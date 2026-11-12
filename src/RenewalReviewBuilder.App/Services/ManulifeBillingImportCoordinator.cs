using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public interface IManulifeBillingStatementParser
{
    BillingExtractionResult Extract(string pdfPath);
}

public interface IManulifeBillingPersistence
{
    Task<ManulifeBillingPersistenceResult> SaveBillingExtractionAsync(
        BillingExtractionResult result,
        ManulifeBillingSaveMode saveMode = ManulifeBillingSaveMode.Cancel,
        long? selectedClientId = null,
        string performedBy = "");
}

public interface IManulifeBillingCsvExporter
{
    BillingCsvExportResult Write(BillingExtractionResult result, string outputFolder);
}

public static class ManulifeBillingLivesReconciliation
{
    public static (int? Difference, string Status, string Message) Evaluate(
        int? pdfLives,
        int? calculatedLives,
        ManulifeBillingValidationOptions? options = null)
    {
        var configured = options ?? new ManulifeBillingValidationOptions();
        if (!pdfLives.HasValue || !calculatedLives.HasValue)
        {
            return (null, BillingReconciliationStatuses.NotAvailable, "PDF lives comparison value was not available.");
        }

        var difference = calculatedLives.Value - pdfLives.Value;
        var variance = Math.Abs(difference);
        if (variance <= configured.ExactLivesTolerance)
        {
            return (difference, BillingReconciliationStatuses.Passed, "");
        }

        if (variance <= configured.LivesWarningTolerance)
        {
            return (difference, BillingReconciliationStatuses.Warning, $"Lives differ by {difference}.");
        }

        return (difference, BillingReconciliationStatuses.Failed, $"Lives differ materially by {difference}.");
    }
}

public sealed class ManulifeBillingCompletionPolicy
{
    private static readonly string[] RequiredReconciliationLevels =
    [
        "Row-level section total",
        "Province total",
        "Tax validation",
        "Grand total"
    ];

    public BillingCompletionValidationResult Evaluate(
        BillingExtractionResult result,
        bool hasExplicitClientSelection = false)
    {
        var validation = new BillingCompletionValidationResult();

        if (!result.ReadablePdfExtraction || result.PageCount <= 0)
        {
            Add(validation, "UnreadablePdf", "The PDF did not produce readable billing text.", blocking: true);
        }

        if (!hasExplicitClientSelection && !result.ClientIdentityConfirmed)
        {
            Add(validation, "ClientIdentityNotConfirmed", "The billing client identity could not be confirmed.", blocking: true);
        }

        if (!result.CarrierIdentityConfirmed ||
            !result.Carrier.Equals("Manulife", StringComparison.OrdinalIgnoreCase))
        {
            Add(validation, "CarrierNotConfirmed", "The document could not be confirmed as a Manulife billing statement.", blocking: true);
        }

        if (string.IsNullOrWhiteSpace(result.BillingPeriod) ||
            !result.PeriodStart.HasValue ||
            !result.PeriodEnd.HasValue ||
            result.PeriodStart > result.PeriodEnd)
        {
            Add(validation, "InvalidBillingPeriod", "The billing period is missing or invalid.", blocking: true);
        }

        if (!result.Details.Any(IsMeaningfulDetail))
        {
            Add(validation, "NoBillingDetails", "No meaningful class or billable detail rows were extracted.", blocking: true);
        }

        if (!result.PdfGrandPremiumBeforeTax.HasValue ||
            !result.PdfGrandTax.HasValue ||
            !result.PdfGrandTotalWithTax.HasValue)
        {
            Add(validation, "MissingStatementTotals", "Required statement-level premium, tax, or total values are unavailable.", blocking: false);
        }

        if (!result.ReconciliationExecuted)
        {
            Add(validation, "ReconciliationNotExecuted", "Billing reconciliation did not execute.", blocking: false);
        }

        foreach (var level in RequiredReconciliationLevels)
        {
            var rows = result.ReconciliationResults
                .Where(row => row.ReconciliationLevel.Equals(level, StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (rows.Count == 0)
            {
                Add(validation, $"Missing{SanitizeCode(level)}", $"Required {level} reconciliation was not created.", blocking: false);
                continue;
            }

            if (rows.Any(row => row.Status.Equals(BillingReconciliationStatuses.Failed, StringComparison.OrdinalIgnoreCase)))
            {
                Add(validation, $"Failed{SanitizeCode(level)}", $"At least one required {level} reconciliation failed.", blocking: false);
            }
            else if (rows.Any(row => row.Status.Equals(BillingReconciliationStatuses.NotAvailable, StringComparison.OrdinalIgnoreCase)))
            {
                Add(validation, $"Unavailable{SanitizeCode(level)}", $"At least one required {level} reconciliation is not available.", blocking: false);
            }

            if (rows.Any(row => row.LivesStatus.Equals(BillingReconciliationStatuses.Failed, StringComparison.OrdinalIgnoreCase)))
            {
                Add(validation, $"LivesMismatch{SanitizeCode(level)}", $"A material lives mismatch exists at {level}.", blocking: false);
            }
            else if (rows.Any(row => row.LivesStatus.Equals(BillingReconciliationStatuses.Warning, StringComparison.OrdinalIgnoreCase)))
            {
                Add(validation, $"LivesWarning{SanitizeCode(level)}", $"A lives variance requires review at {level}.", blocking: false);
            }
        }

        foreach (var warning in result.Warnings.Where(warning => !string.IsNullOrWhiteSpace(warning)))
        {
            Add(validation, "ParserWarning", warning, blocking: false);
        }

        return validation;
    }

    private static bool IsMeaningfulDetail(BillingOptionPremiumDetail row)
    {
        return !string.IsNullOrWhiteSpace(row.Benefit) &&
               !string.IsNullOrWhiteSpace(row.ClassCode) &&
               (row.Lives.HasValue ||
                row.Volume.HasValue ||
                row.MonthlyPremium != 0m ||
                row.PremiumAdjustments != 0m ||
                row.TotalPremiumAdjustments != 0m);
    }

    private static void Add(BillingCompletionValidationResult result, string code, string message, bool blocking)
    {
        if (result.Issues.Any(issue => issue.Code.Equals(code, StringComparison.OrdinalIgnoreCase)))
        {
            return;
        }

        result.Issues.Add(new BillingCompletionIssue
        {
            Code = code,
            Message = message,
            IsBlocking = blocking
        });
    }

    private static string SanitizeCode(string value)
    {
        return new string((value ?? "").Where(char.IsLetterOrDigit).ToArray());
    }
}

public sealed class ManulifeBillingImportCoordinator
{
    private readonly IManulifeBillingStatementParser _parser;
    private readonly IManulifeBillingPersistence _persistence;
    private readonly IManulifeBillingCsvExporter _csvExporter;
    private readonly ManulifeBillingCompletionPolicy _completionPolicy;

    public ManulifeBillingImportCoordinator(
        IManulifeBillingPersistence persistence,
        IManulifeBillingStatementParser? parser = null,
        IManulifeBillingCsvExporter? csvExporter = null,
        ManulifeBillingCompletionPolicy? completionPolicy = null)
    {
        _persistence = persistence;
        _parser = parser ?? new ManulifeBillingStatementParser();
        _csvExporter = csvExporter ?? new BillingCsvWriter();
        _completionPolicy = completionPolicy ?? new ManulifeBillingCompletionPolicy();
    }

    public async Task<ManulifeBillingImportResult> ImportAsync(ManulifeBillingImportRequest request)
    {
        var import = new ManulifeBillingImportResult();
        if (string.IsNullOrWhiteSpace(request.InputPath) ||
            !File.Exists(request.InputPath) ||
            !Path.GetExtension(request.InputPath).Equals(".pdf", StringComparison.OrdinalIgnoreCase))
        {
            import.UserMessage = "Select a readable Manulife billing PDF before importing.";
            import.TechnicalErrors.Add("The input file was missing, unreadable, or was not a PDF.");
            import.ErrorCount = 1;
            return import;
        }

        BillingExtractionResult extraction;
        try
        {
            extraction = await Task.Run(() => _parser.Extract(request.InputPath));
            import.Extraction = extraction;
            import.ParsedRowCount = extraction.Details.Count;
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            import.UserMessage = "The Manulife billing PDF could not be read. No data was saved.";
            import.TechnicalErrors.Add($"{ex.GetType().Name}: {ex.Message}");
            import.ErrorCount = 1;
            return import;
        }

        var explicitClientSelection = request.SelectedClientId.HasValue ||
                                      !string.IsNullOrWhiteSpace(request.SelectedClientName) ||
                                      !string.IsNullOrWhiteSpace(request.SelectedClientNumber);
        if (HasClientMismatch(extraction, request))
        {
            import.UserMessage = "The selected client does not match the client identity on the billing statement. No data was saved.";
            import.TechnicalErrors.Add("Selected client identity did not match the parsed billing identity.");
            import.ErrorCount = 1;
            return import;
        }

        if (explicitClientSelection)
        {
            extraction.ClientIdentityConfirmed = true;
            extraction.ClientName = FirstNonBlank(request.SelectedClientName, extraction.ClientName);
            extraction.ClientNumber = FirstNonBlank(request.SelectedClientNumber, extraction.ClientNumber);
        }

        var completion = _completionPolicy.Evaluate(extraction, explicitClientSelection);
        import.WarningCount = completion.WarningCount;
        import.ErrorCount = completion.ErrorCount;
        import.TechnicalErrors.AddRange(completion.Issues.Select(issue => $"{issue.Code}: {issue.Message}"));
        import.ReconciliationStatus = OverallReconciliationStatus(extraction.ReconciliationResults);
        if (!completion.CanPersist)
        {
            extraction.ImportStatus = BillingImportStatuses.Failed;
            import.FinalStatus = BillingImportStatuses.Failed;
            import.UserMessage = "The document failed required Manulife billing checks. No data was saved.";
            return import;
        }

        extraction.ImportStatus = completion.RecommendedStatus;
        var persistence = await _persistence.SaveBillingExtractionAsync(
            extraction,
            request.SaveMode,
            request.SelectedClientId,
            FirstNonBlank(request.PerformedBy, Environment.UserName));
        import.Persistence = persistence;
        import.ImportBatchId = persistence.ImportBatchId;
        import.DocumentId = persistence.DocumentId;
        import.StatementVersionId = persistence.StatementVersionId;
        import.SavedRowCount = persistence.SavedRowCount;

        if (!persistence.Succeeded)
        {
            import.FinalStatus = persistence.Outcome is ManulifeBillingSaveOutcomes.Duplicate or ManulifeBillingSaveOutcomes.ExistingVersion
                ? BillingImportStatuses.NeedsReview
                : BillingImportStatuses.Failed;
            import.ExportStatus = BillingExportStatuses.Skipped;
            import.UserMessage = persistence.Message;
            import.ErrorCount++;
            import.TechnicalErrors.Add($"Persistence {persistence.Outcome}: {persistence.Message}");
            return import;
        }

        import.Succeeded = true;
        import.FinalStatus = extraction.ImportStatus;
        if (persistence.CrossClientHashMatch)
        {
            import.FinalStatus = BillingImportStatuses.NeedsReview;
            import.WarningCount++;
        }

        if (!request.ExportCsv)
        {
            import.ExportStatus = BillingExportStatuses.NotRequested;
        }
        else if (string.IsNullOrWhiteSpace(request.OutputFolder))
        {
            import.ExportStatus = BillingExportStatuses.Skipped;
            import.WarningCount++;
        }
        else
        {
            try
            {
                import.CsvExport = _csvExporter.Write(extraction, request.OutputFolder);
                import.ExportStatus = BillingExportStatuses.Completed;
            }
            catch (Exception ex)
            {
                AppLogger.Error(ex);
                import.ExportStatus = BillingExportStatuses.Failed;
                import.WarningCount++;
                import.TechnicalErrors.Add($"CSV export failed: {ex.Message}");
            }
        }

        import.UserMessage = import.FinalStatus == BillingImportStatuses.Completed
            ? "Billing extraction, validation, and database save completed."
            : "Billing data was saved and requires review.";
        return import;
    }

    private static bool HasClientMismatch(BillingExtractionResult extraction, ManulifeBillingImportRequest request)
    {
        if (!extraction.ClientIdentityConfirmed)
        {
            return false;
        }

        if (!string.IsNullOrWhiteSpace(request.SelectedClientNumber) &&
            !string.IsNullOrWhiteSpace(extraction.ClientNumber) &&
            !request.SelectedClientNumber.Equals(extraction.ClientNumber, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return !string.IsNullOrWhiteSpace(request.SelectedClientName) &&
               !string.IsNullOrWhiteSpace(extraction.ClientName) &&
               !request.SelectedClientName.Equals(extraction.ClientName, StringComparison.OrdinalIgnoreCase);
    }

    private static string OverallReconciliationStatus(IEnumerable<BillingReconciliationResult> rows)
    {
        var statuses = rows.Select(row => row.Status).ToList();
        if (statuses.Any(status => status.Equals(BillingReconciliationStatuses.Failed, StringComparison.OrdinalIgnoreCase)))
        {
            return BillingReconciliationStatuses.Failed;
        }

        if (statuses.Any(status => status.Equals(BillingReconciliationStatuses.Warning, StringComparison.OrdinalIgnoreCase)))
        {
            return BillingReconciliationStatuses.Warning;
        }

        return statuses.Any(status => status.Equals(BillingReconciliationStatuses.Passed, StringComparison.OrdinalIgnoreCase))
            ? BillingReconciliationStatuses.Passed
            : BillingReconciliationStatuses.NotAvailable;
    }

    private static string FirstNonBlank(params string?[] values)
    {
        return values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)) ?? "";
    }
}

using Microsoft.Data.Sqlite;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed partial class DatabaseService
{
    public async Task<HistoricalRenewalImportPreview> PreviewHistoricalRenewalImportAsync(
        HistoricalRenewalImportRequest request)
    {
        await InitializeCoreAsync();
        var validationMessage = ValidateHistoricalRenewalRequest(request);
        if (!string.IsNullOrWhiteSpace(validationMessage))
        {
            return new HistoricalRenewalImportPreview { Message = validationMessage };
        }

        var sourceHash = await ComputeFileHashAsync(Path.GetFullPath(request.RenewalPdfPath));
        using var connection = await OpenConnectionAsync();
        return await BuildHistoricalRenewalPreviewAsync(connection, null, request, sourceHash);
    }

    public async Task<HistoricalRenewalImportResult> ImportHistoricalRenewalAsync(
        HistoricalRenewalImportRequest request)
    {
        await InitializeCoreAsync();
        var validationMessage = ValidateHistoricalRenewalRequest(request);
        if (!string.IsNullOrWhiteSpace(validationMessage))
        {
            return new HistoricalRenewalImportResult { Message = validationMessage };
        }

        var fullPdfPath = Path.GetFullPath(request.RenewalPdfPath);
        var sourceHash = await ComputeFileHashAsync(fullPdfPath);
        using var connection = await OpenConnectionAsync();
        using var transaction = connection.BeginTransaction();
        try
        {
            var preview = await BuildHistoricalRenewalPreviewAsync(connection, transaction, request, sourceHash);
            if (!preview.CanImport)
            {
                transaction.Rollback();
                return new HistoricalRenewalImportResult
                {
                    WasDuplicate = preview.IsDuplicateSource || preview.IsDuplicateRenewal,
                    Message = preview.Message
                };
            }

            var primaryDocument = await UpsertDocumentAsync(
                connection,
                transaction,
                request.ClientId,
                DocumentTypes.Renewal,
                BenefitsIntelligenceModules.RenewalIntelligence,
                fullPdfPath,
                request.Carrier,
                request.EffectiveDate,
                null,
                null);
            if (primaryDocument.DocumentId is not > 0)
            {
                throw new InvalidOperationException("The historical renewal source document could not be registered.");
            }

            await RegisterHistoricalSupportingWorkbookAsync(connection, transaction, request);

            var now = DateTime.Now;
            await InsertRenewalAsync(
                connection,
                transaction,
                new RenewalRecord
                {
                    ClientId = request.ClientId,
                    RenewalYear = request.RenewalYear,
                    Carrier = request.Carrier,
                    BenefitLine = "Overall",
                    RateChangePercent = decimal.Round(request.FinalAcceptedIncreasePercent * 100m, 4),
                    RenewalStatus = "Completed",
                    SourceDocumentId = primaryDocument.DocumentId,
                    DateLoaded = now
                });

            using var renewalIdCommand = CreateCommand(connection, transaction, "SELECT last_insert_rowid();");
            var renewalId = ToLong(await renewalIdCommand.ExecuteScalarAsync());
            var renewalVersionId = await InsertHistoricalCompletedRenewalVersionAsync(
                connection,
                transaction,
                request,
                primaryDocument.DocumentId.Value,
                renewalId,
                now);

            transaction.Commit();
            InvalidateExecutiveReportingCache();
            return new HistoricalRenewalImportResult
            {
                Succeeded = true,
                Message = $"Imported the completed {request.RenewalYear} renewal without changing the current client or renewal selection.",
                DocumentId = primaryDocument.DocumentId,
                RenewalId = renewalId,
                RenewalVersionId = renewalVersionId
            };
        }
        catch (Exception ex)
        {
            transaction.Rollback();
            AppLogger.Error(ex);
            return new HistoricalRenewalImportResult
            {
                Message = $"The historical renewal was not imported. No changes were committed: {ex.Message}"
            };
        }
    }

    private static async Task<HistoricalRenewalImportPreview> BuildHistoricalRenewalPreviewAsync(
        SqliteConnection connection,
        SqliteTransaction? transaction,
        HistoricalRenewalImportRequest request,
        string sourceHash)
    {
        using var clientCommand = CreateCommand(
            connection,
            transaction,
            "SELECT * FROM Clients WHERE client_id = $client_id LIMIT 1;");
        AddParameter(clientCommand, "$client_id", request.ClientId);
        using var clientReader = await clientCommand.ExecuteReaderAsync();
        if (!await clientReader.ReadAsync())
        {
            return new HistoricalRenewalImportPreview
            {
                Message = "The selected persisted client no longer exists. No renewal was imported."
            };
        }
        var client = ReadClient(clientReader);
        await clientReader.CloseAsync();

        using var currentCommand = CreateCommand(
            connection,
            transaction,
            "SELECT MAX(effective_date) FROM renewal_versions WHERE client_id = $client_id AND TRIM(IFNULL(effective_date, '')) <> ''; ");
        AddParameter(currentCommand, "$client_id", request.ClientId);
        var currentDateText = (await currentCommand.ExecuteScalarAsync())?.ToString();
        var currentDate = DateOnly.TryParse(currentDateText, out var parsedCurrentDate)
            ? parsedCurrentDate
            : client.EffectiveDate;
        var isHistoricalDate = currentDate.HasValue && request.EffectiveDate < currentDate.Value;

        using var sourceDuplicateCommand = CreateCommand(
            connection,
            transaction,
            "SELECT 1 FROM Documents WHERE client_id = $client_id AND document_type = $document_type AND file_hash = $file_hash LIMIT 1;");
        AddParameter(sourceDuplicateCommand, "$client_id", request.ClientId);
        AddParameter(sourceDuplicateCommand, "$document_type", DocumentTypes.Renewal);
        AddParameter(sourceDuplicateCommand, "$file_hash", sourceHash);
        var isDuplicateSource = await sourceDuplicateCommand.ExecuteScalarAsync() is not null;

        using var renewalDuplicateCommand = CreateCommand(
            connection,
            transaction,
            """
            SELECT 1
            FROM renewal_versions
            WHERE client_id = $client_id
              AND IFNULL(effective_date, '') = $effective_date
              AND LOWER(TRIM(IFNULL(carrier, ''))) = LOWER(TRIM($carrier))
              AND LOWER(TRIM(status)) IN ('completed', 'final approved')
            LIMIT 1;
            """);
        AddParameter(renewalDuplicateCommand, "$client_id", request.ClientId);
        AddParameter(renewalDuplicateCommand, "$effective_date", ToDateText(request.EffectiveDate));
        AddParameter(renewalDuplicateCommand, "$carrier", Clean(request.Carrier));
        var isDuplicateRenewal = await renewalDuplicateCommand.ExecuteScalarAsync() is not null;

        var checks = new List<string>
        {
            $"Client locked to {client.ClientName} (database client {client.ClientId}).",
            currentDate.HasValue
                ? $"Historical date check: {request.EffectiveDate:MMM d, yyyy} is {(isHistoricalDate ? "earlier" : "not earlier")} than the current {currentDate.Value:MMM d, yyyy} renewal."
                : "Historical date check failed because the current renewal effective date is unavailable.",
            isDuplicateSource ? "The renewal PDF was already imported for this client." : "No duplicate renewal PDF hash was found.",
            isDuplicateRenewal ? "A completed renewal already exists for this effective date and carrier." : "No completed renewal exists for this effective date and carrier.",
            "The current client carrier, client effective date, and selected renewal versions will not be updated.",
            "The imported renewal version will be saved as non-current."
        };
        var canImport = isHistoricalDate && !isDuplicateSource && !isDuplicateRenewal;
        var message = canImport
            ? "All historical-import safety checks passed."
            : !currentDate.HasValue
                ? "The import is blocked because the current renewal effective date could not be verified."
                : !isHistoricalDate
                    ? $"The import is blocked because the renewal date must be earlier than {currentDate.Value:MMM d, yyyy}."
                    : isDuplicateSource
                        ? "The import is blocked because this renewal PDF was already imported for the selected client."
                        : "The import is blocked because a completed renewal already exists for this date and carrier.";

        return new HistoricalRenewalImportPreview
        {
            CanImport = canImport,
            Message = message,
            Client = client,
            CurrentRenewalEffectiveDate = currentDate,
            IsHistoricalDate = isHistoricalDate,
            IsDuplicateSource = isDuplicateSource,
            IsDuplicateRenewal = isDuplicateRenewal,
            RenewalPdfHash = sourceHash,
            Checks = checks
        };
    }

    private static async Task RegisterHistoricalSupportingWorkbookAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        HistoricalRenewalImportRequest request)
    {
        if (string.IsNullOrWhiteSpace(request.SupportingWorkbookPath))
        {
            return;
        }

        var fullPath = Path.GetFullPath(request.SupportingWorkbookPath);
        var hash = await ComputeFileHashAsync(fullPath);
        var existingId = await FindDocumentIdByHashAsync(
            connection,
            transaction,
            request.ClientId,
            DocumentTypes.HistoricalRenewalSupportingWorkbook,
            hash);
        if (existingId.HasValue)
        {
            return;
        }

        await UpsertDocumentAsync(
            connection,
            transaction,
            request.ClientId,
            DocumentTypes.HistoricalRenewalSupportingWorkbook,
            BenefitsIntelligenceModules.RenewalIntelligence,
            fullPath,
            request.Carrier,
            request.EffectiveDate,
            null,
            null);
    }

    private static async Task<long> InsertHistoricalCompletedRenewalVersionAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        HistoricalRenewalImportRequest request,
        long sourceDocumentId,
        long renewalId,
        DateTime now)
    {
        using var command = CreateCommand(
            connection,
            transaction,
            """
            INSERT INTO renewal_versions (
                client_id, version_name, source_module, status, is_selected,
                effective_date, source_document_id, date_created,
                renewal_id, renewal_year, carrier, updated_at, created_by, notes,
                requested_increase_percent, final_accepted_increase_percent,
                negotiation_status, negotiation_stage, negotiation_completed_date,
                negotiated_savings_annual, negotiated_savings_monthly, negotiated_savings_type,
                negotiated_by, renewal_result, renewal_completed_date,
                renewal_approved_by, renewal_accepted_date, renewal_version_number
            )
            VALUES (
                $client_id, $version_name, $source_module, $status, 0,
                $effective_date, $source_document_id, $date_created,
                $renewal_id, $renewal_year, $carrier, $updated_at, $created_by, $notes,
                $requested_increase_percent, $final_accepted_increase_percent,
                $negotiation_status, $negotiation_stage, $negotiation_completed_date,
                $negotiated_savings_annual, $negotiated_savings_monthly, $negotiated_savings_type,
                $negotiated_by, $renewal_result, $renewal_completed_date,
                $renewal_approved_by, $renewal_accepted_date, 1
            );
            SELECT last_insert_rowid();
            """);
        var importedBy = FirstNonBlank(request.ImportedBy, "Desktop user");
        var notes = string.IsNullOrWhiteSpace(request.SupportingWorkbookPath)
            ? $"Historical renewal imported from {Path.GetFileName(request.RenewalPdfPath)}."
            : $"Historical renewal imported from {Path.GetFileName(request.RenewalPdfPath)} with supporting workbook {Path.GetFileName(request.SupportingWorkbookPath)}.";
        AddParameter(command, "$client_id", request.ClientId);
        AddParameter(command, "$version_name", RenewalVersionNames.FinalApproved);
        AddParameter(command, "$source_module", BenefitsIntelligenceModules.RenewalIntelligence);
        AddParameter(command, "$status", "Completed");
        AddParameter(command, "$effective_date", ToDateText(request.EffectiveDate));
        AddParameter(command, "$source_document_id", sourceDocumentId);
        AddParameter(command, "$date_created", ToDateTimeText(now));
        AddParameter(command, "$renewal_id", renewalId);
        AddParameter(command, "$renewal_year", request.RenewalYear);
        AddParameter(command, "$carrier", Clean(request.Carrier));
        AddParameter(command, "$updated_at", ToDateTimeText(now));
        AddParameter(command, "$created_by", importedBy);
        AddParameter(command, "$notes", notes);
        AddParameter(command, "$requested_increase_percent", request.RequestedIncreasePercent);
        AddParameter(command, "$final_accepted_increase_percent", request.FinalAcceptedIncreasePercent);
        AddParameter(command, "$negotiation_status", NegotiationStatus.Accepted.ToString());
        AddParameter(command, "$negotiation_stage", "Completed");
        AddParameter(command, "$negotiation_completed_date", ToDateText(request.EffectiveDate));
        AddParameter(command, "$negotiated_savings_annual", request.NegotiatedSavingsAnnual);
        AddParameter(command, "$negotiated_savings_monthly", request.NegotiatedSavingsAnnual.HasValue ? decimal.Round(request.NegotiatedSavingsAnnual.Value / 12m, 2) : null);
        AddParameter(command, "$negotiated_savings_type", request.NegotiatedSavingsAnnual.HasValue ? NegotiatedSavingsType.Recurring.ToString() : null);
        AddParameter(command, "$negotiated_by", importedBy);
        AddParameter(command, "$renewal_result", "Completed");
        AddParameter(command, "$renewal_completed_date", ToDateText(request.EffectiveDate));
        AddParameter(command, "$renewal_approved_by", importedBy);
        AddParameter(command, "$renewal_accepted_date", ToDateText(request.EffectiveDate));
        return ToLong(await command.ExecuteScalarAsync());
    }

    private static string ValidateHistoricalRenewalRequest(HistoricalRenewalImportRequest request)
    {
        if (request.ClientId <= 0) return "A persisted client is required.";
        if (string.IsNullOrWhiteSpace(request.RenewalPdfPath) || !File.Exists(request.RenewalPdfPath)) return "Select an existing historical renewal PDF.";
        if (!Path.GetExtension(request.RenewalPdfPath).Equals(".pdf", StringComparison.OrdinalIgnoreCase)) return "The primary historical renewal source must be a PDF.";
        if (!string.IsNullOrWhiteSpace(request.SupportingWorkbookPath))
        {
            if (!File.Exists(request.SupportingWorkbookPath)) return "The selected supporting workbook was not found.";
            var extension = Path.GetExtension(request.SupportingWorkbookPath);
            if (!extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) &&
                !extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase) &&
                !extension.Equals(".xls", StringComparison.OrdinalIgnoreCase))
            {
                return "The supporting source must be an Excel workbook.";
            }
        }
        if (request.EffectiveDate == default) return "Enter the historical renewal effective date.";
        if (request.RenewalYear != request.EffectiveDate.Year) return "Renewal year must match the effective-date year.";
        if (string.IsNullOrWhiteSpace(request.Carrier)) return "Enter the historical carrier.";
        if (request.FinalAcceptedIncreasePercent < -1m || request.FinalAcceptedIncreasePercent > 5m) return "Final accepted increase must be between -100% and 500%.";
        if (request.RequestedIncreasePercent is < -1m or > 5m) return "Requested increase must be between -100% and 500%.";
        if (request.NegotiatedSavingsAnnual < 0m) return "Negotiated annual savings cannot be negative.";
        return "";
    }
}

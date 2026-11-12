using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services.Normalization;

namespace RenewalReviewBuilder.App.Services;

public sealed class CanonicalControlledBookletReprocessingService
{
    private static readonly HashSet<string> RevalidatedPlanDesignIssueCodes = new(StringComparer.OrdinalIgnoreCase)
    {
        ControlledBookletReprocessingIssueCodes.NoStagedRows,
        ControlledBookletReprocessingIssueCodes.PercentageOutOfRange,
        ControlledBookletReprocessingIssueCodes.MultiplePercentagesPreserved,
        ControlledBookletReprocessingIssueCodes.ProvisionIdentityMissing,
        ControlledBookletReprocessingIssueCodes.DuplicateProvisionIdentity,
        ControlledBookletReprocessingIssueCodes.RowLoss,
        ControlledBookletReprocessingIssueCodes.BenefitMissing,
        ControlledBookletReprocessingIssueCodes.NewValidationIssue
    };

    private readonly DatabaseService _databaseService;
    private readonly Func<BookletExtractionRequest, IProgress<BookletExtractionProgress>?, CancellationToken, Task<BookletExtractionResult>> _runExtraction;
    private readonly Func<IReadOnlyList<string>, CancellationToken, Task<IReadOnlyList<BookletClientIdentity>>> _discoverIdentities;

    public CanonicalControlledBookletReprocessingService(
        DatabaseService? databaseService = null,
        Func<BookletExtractionRequest, IProgress<BookletExtractionProgress>?, CancellationToken, Task<BookletExtractionResult>>? runExtraction = null,
        Func<IReadOnlyList<string>, CancellationToken, Task<IReadOnlyList<BookletClientIdentity>>>? discoverIdentities = null)
    {
        _databaseService = databaseService ?? new DatabaseService();
        var bookletRunner = new BookletExtractorRunner();
        _runExtraction = runExtraction ?? bookletRunner.RunAsync;
        _discoverIdentities = discoverIdentities ?? bookletRunner.DiscoverClientIdentitiesAsync;
    }

    public Task<IReadOnlyList<ControlledBookletDocumentOption>> GetDocumentsAsync() =>
        _databaseService.GetControlledBookletDocumentsAsync();

    public Task<ControlledBookletReprocessingRun?> GetRunAsync(long runId) =>
        _databaseService.GetControlledBookletReprocessingRunAsync(runId);

    public Task<ControlledBookletReprocessingRun?> GetActiveRunForDocumentAsync(long documentId) =>
        _databaseService.GetActiveControlledBookletReprocessingRunForDocumentAsync(documentId);

    public async Task<ControlledBookletReprocessingRun> StageAsync(
        ControlledBookletReprocessingRequest request,
        IProgress<BookletExtractionProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var document = await _databaseService.GetControlledBookletDocumentAsync(request.SourceDocumentId)
            ?? throw new InvalidOperationException($"Registered document {request.SourceDocumentId} was not found.");
        var runId = await _databaseService.BeginControlledBookletReprocessingRunAsync(document, request.RequestedBy);
        var run = CreateRun(runId, document, request.RequestedBy);

        try
        {
            var currentRows = await _databaseService.GetPlanDesignForDocumentAsync(document.DocumentId);
            ApplyDocumentIdentity(currentRows, document);
            run.BeforeRows.AddRange(PlanDesignCanonicalMapper.NormalizeProvisionRows(currentRows));
            await AddSourcePreflightIssuesAsync(run, document, cancellationToken);
            if (HasBlockingIssues(run))
            {
                return await FinalizeStageAsync(run);
            }

            progress?.Report(new BookletExtractionProgress(
                "Confirming the registered booklet client identity...",
                0,
                1,
                0,
                0));
            var identities = await _discoverIdentities([document.FilePath], cancellationToken);
            var identity = SelectIdentity(identities, document.FilePath);
            ApplyIdentity(run, identity);
            AddIdentityIssues(run, document, identity);
            if (HasBlockingIssues(run))
            {
                return await FinalizeStageAsync(run);
            }

            run.OutputFolder = BuildOutputFolder(request.OutputRoot, document.DocumentId, run.RunId);
            var extraction = await _runExtraction(
                new BookletExtractionRequest
                {
                    InputPaths = [document.FilePath],
                    OutputFolder = run.OutputFolder
                },
                progress,
                cancellationToken);
            run.MasterWorkbookPath = extraction.MasterWorkbookPath;
            run.ReviewWorkbookPath = extraction.ReviewWorkbookPath;
            run.BenefitAdministrationRecords.AddRange(extraction.BenefitAdministrationRecords);

            AddExtractionIssues(run, extraction);
            var stagedRows = DatabaseService.ReadPlanDesignRowsForBookletExtraction(
                extraction,
                [document.FilePath]);
            AddRowIdentityIssues(run, document, stagedRows);
            ApplyDocumentIdentity(stagedRows, document);
            var canonicalStagedRows = PlanDesignCanonicalMapper.NormalizeProvisionRows(stagedRows);
            ValidateStagedRows(run, document, canonicalStagedRows);
            run.StagedRows.AddRange(canonicalStagedRows);

            await AddComparisonIssuesAsync(run, document);
            run.Changes.AddRange(BuildChanges(run.BeforeRows, run.StagedRows));
            return await FinalizeStageAsync(run);
        }
        catch (OperationCanceledException)
        {
            run.Status = ControlledBookletReprocessingStatuses.Failed;
            run.ErrorMessage = "Controlled booklet reprocessing was cancelled before approval.";
            AddIssue(
                run,
                ControlledBookletReprocessingIssueCodes.ExtractionFailed,
                run.ErrorMessage,
                isBlocking: true,
                rawSourceReference: document.FilePath);
            await FinalizeStageAsync(run);
            throw;
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            run.Status = ControlledBookletReprocessingStatuses.Failed;
            run.ErrorMessage = ex.Message;
            AddIssue(
                run,
                ControlledBookletReprocessingIssueCodes.ExtractionFailed,
                $"Staging failed: {ex.Message}",
                isBlocking: true,
                rawSourceReference: document.FilePath);
            run.StagedAt = DateTime.Now;
            UpdateRunCounts(run);
            await _databaseService.SaveControlledBookletReprocessingStageAsync(run);
            return await ReloadRunAsync(run.RunId);
        }
    }

    public async Task<ControlledBookletReprocessingRun> ApproveAndCommitAsync(
        long runId,
        string approvedBy,
        CancellationToken cancellationToken = default)
    {
        var run = await ReloadRunAsync(runId);
        if (!run.CanApprove)
        {
            throw new InvalidOperationException("This booklet is not ready for approval. Resolve every blocking issue and stage it again.");
        }

        var document = await _databaseService.GetControlledBookletDocumentAsync(run.SourceDocumentId);
        var preCommitError = await ValidateCommitIdentityAsync(run, document, cancellationToken);
        if (!string.IsNullOrWhiteSpace(preCommitError))
        {
            await _databaseService.FailControlledBookletRunAsync(run.RunId, preCommitError);
            return await ReloadRunAsync(run.RunId);
        }

        string backupPath;
        try
        {
            backupPath = await _databaseService.CreateControlledBookletBackupAsync(run.RunId, run.SourceDocumentId);
            await _databaseService.BeginControlledBookletCommitAsync(
                run.RunId,
                FirstNonBlank(approvedBy, Environment.UserName),
                backupPath);

            var client = await _databaseService.GetClientAsync(run.ClientId)
                ?? throw new InvalidOperationException($"Client {run.ClientId} was not found before commit.");
            ApplyDocumentIdentity(run.StagedRows, document!);
            var save = await _databaseService.SavePlanDesignAsync(
                client,
                [run.SourceFilePath],
                run.StagedRows,
                run.ClientId,
                run.BenefitAdministrationRecords,
                expectedSourceDocumentId: run.SourceDocumentId,
                expectedSourceFileHash: run.SourceFileHash);
            if (!save.Succeeded)
            {
                throw new InvalidOperationException(save.Message);
            }

            if (save.DocumentId != run.SourceDocumentId || save.InsertedRows != run.StagedRowCount)
            {
                throw new InvalidOperationException(
                    $"Post-save identity mismatch. Expected document {run.SourceDocumentId} and {run.StagedRowCount} rows, " +
                    $"but the save reported document {save.DocumentId?.ToString(CultureInfo.InvariantCulture) ?? "none"} and {save.InsertedRows} rows.");
            }

            var audit = await _databaseService.GetPlanDesignPostCommitAuditAsync(run.SourceDocumentId);
            if (audit.RowCount != run.StagedRowCount || audit.OutOfRangeCount > 0)
            {
                throw new InvalidOperationException(
                    $"Post-commit audit failed. Expected {run.StagedRowCount} rows and no out-of-range percentages; " +
                    $"found {audit.RowCount} rows and {audit.OutOfRangeCount} out-of-range value(s). The backup path is retained for rollback.");
            }

            await _databaseService.CompleteControlledBookletCommitAsync(
                run.RunId,
                audit.RowCount,
                $"Approved Plan Design replacement completed. Backup: {backupPath}");
            return await ReloadRunAsync(run.RunId);
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            await _databaseService.FailControlledBookletRunAsync(
                run.RunId,
                $"Commit failed: {ex.Message}");
            return await ReloadRunAsync(run.RunId);
        }
    }

    public async Task<ControlledBookletReprocessingRun> RejectAsync(
        long runId,
        string rejectedBy,
        string reason)
    {
        await _databaseService.RejectControlledBookletRunAsync(
            runId,
            FirstNonBlank(rejectedBy, Environment.UserName),
            FirstNonBlank(reason, "Rejected during controlled booklet review."));
        return await ReloadRunAsync(runId);
    }

    public async Task<ControlledBookletReprocessingRun> ApplyStagedEditAsync(
        long runId,
        ControlledBookletStagedEditRequest request)
    {
        ArgumentNullException.ThrowIfNull(request);
        var run = await ReloadRunAsync(runId);
        if (!run.CanEdit)
        {
            throw new InvalidOperationException(
                "Only blocked or ready-for-approval booklet reviews can be corrected.");
        }

        var reason = request.Reason.Trim();
        if (string.IsNullOrWhiteSpace(reason))
        {
            throw new InvalidOperationException("A correction reason is required for the audit history.");
        }

        var rowsByKey = BuildComparableRows(run.StagedRows);
        if (!rowsByKey.TryGetValue(request.ComparisonKey, out var comparable))
        {
            throw new InvalidOperationException(
                "The selected staged provision is no longer available. Refresh the review and select it again.");
        }

        var row = comparable.Row;
        var editedBy = FirstNonBlank(request.EditedBy, Environment.UserName);
        var editedAt = DateTime.Now;
        var newRateText = request.ReimbursementPercent.HasValue
            ? $"{request.ReimbursementPercent.Value.ToString("0.####", CultureInfo.InvariantCulture)}%"
            : "";
        var edits = new List<ControlledBookletReprocessingEdit>();
        AddEdit(
            edits,
            run.RunId,
            request.ComparisonKey,
            "reimbursement_percent",
            FirstNonBlank(row.ReimbursementText, row.ReimbursementPercent?.ToString(CultureInfo.InvariantCulture)),
            newRateText,
            reason,
            editedBy,
            editedAt);
        AddEdit(edits, run.RunId, request.ComparisonKey, "maximum_amount", row.MaximumAmount, request.MaximumAmount, reason, editedBy, editedAt);
        AddEdit(edits, run.RunId, request.ComparisonKey, "frequency", row.Frequency, request.Frequency, reason, editedBy, editedAt);
        AddEdit(edits, run.RunId, request.ComparisonKey, "waiting_period", row.WaitingPeriod, request.WaitingPeriod, reason, editedBy, editedAt);
        AddEdit(edits, run.RunId, request.ComparisonKey, "termination_age", row.TerminationAge, request.TerminationAge, reason, editedBy, editedAt);
        AddEdit(edits, run.RunId, request.ComparisonKey, "coverage_description", row.CoverageDescription, request.CoverageDescription, reason, editedBy, editedAt);
        AddEdit(edits, run.RunId, request.ComparisonKey, "notes", row.Notes, request.Notes, reason, editedBy, editedAt);
        if (edits.Count == 0)
        {
            throw new InvalidOperationException("No staged values were changed.");
        }

        row.ReimbursementPercent = request.ReimbursementPercent;
        row.ReimbursementText = newRateText;
        row.MaximumAmount = request.MaximumAmount.Trim();
        row.Frequency = request.Frequency.Trim();
        row.WaitingPeriod = request.WaitingPeriod.Trim();
        row.TerminationAge = request.TerminationAge.Trim();
        row.CoverageDescription = request.CoverageDescription.Trim();
        row.Notes = request.Notes.Trim();

        var document = await _databaseService.GetControlledBookletDocumentAsync(run.SourceDocumentId)
            ?? throw new InvalidOperationException($"Registered document {run.SourceDocumentId} was not found.");
        run.Issues.RemoveAll(issue => RevalidatedPlanDesignIssueCodes.Contains(issue.IssueCode));
        ValidateStagedRows(run, document, run.StagedRows);
        await AddComparisonIssuesAsync(run, document);
        run.Changes.Clear();
        run.Changes.AddRange(BuildChanges(run.BeforeRows, run.StagedRows));
        run.Status = HasBlockingIssues(run)
            ? ControlledBookletReprocessingStatuses.Blocked
            : ControlledBookletReprocessingStatuses.ReadyForApproval;
        run.StagedAt = editedAt;
        UpdateRunCounts(run);
        await _databaseService.SaveControlledBookletReprocessingStageAsync(run, edits);
        return await ReloadRunAsync(run.RunId);
    }

    private async Task AddSourcePreflightIssuesAsync(
        ControlledBookletReprocessingRun run,
        ControlledBookletDocumentOption document,
        CancellationToken cancellationToken)
    {
        if (DocumentIdentityReconciliationStatuses.IsExcludedFromProcessing(document.ProcessingStatus))
        {
            AddIssue(
                run,
                ControlledBookletReprocessingIssueCodes.SourceDocumentQuarantined,
                $"Document {document.DocumentId} is excluded from processing because its identity status is '{document.ProcessingStatus}'.",
                isBlocking: true,
                fieldName: "processing_status",
                rawSourceReference: document.FilePath);
            return;
        }

        if (!document.DocumentType.Equals(DocumentTypes.BenefitBooklet, StringComparison.OrdinalIgnoreCase))
        {
            AddIssue(
                run,
                ControlledBookletReprocessingIssueCodes.SourceDocumentTypeInvalid,
                $"Document {document.DocumentId} is registered as '{document.DocumentType}', not a benefit booklet.",
                isBlocking: true,
                fieldName: "document_type");
        }

        if (!File.Exists(document.FilePath))
        {
            AddIssue(
                run,
                ControlledBookletReprocessingIssueCodes.SourceFileMissing,
                $"The registered booklet file was not found: {document.FilePath}",
                isBlocking: true,
                fieldName: "file_path",
                rawSourceReference: document.FilePath);
            return;
        }

        if (string.IsNullOrWhiteSpace(document.FileHash))
        {
            AddIssue(
                run,
                ControlledBookletReprocessingIssueCodes.SourceHashMissing,
                "The registered document does not have a SHA-256 source hash.",
                isBlocking: true,
                fieldName: "file_hash",
                rawSourceReference: document.FilePath);
        }
        else
        {
            var currentHash = await ComputeSha256Async(document.FilePath, cancellationToken);
            if (!currentHash.Equals(document.FileHash, StringComparison.OrdinalIgnoreCase))
            {
                AddIssue(
                    run,
                    ControlledBookletReprocessingIssueCodes.SourceHashChanged,
                    "The file on disk no longer matches the SHA-256 hash stored with the selected document.",
                    isBlocking: true,
                    fieldName: "file_hash",
                    rawSourceReference: document.FilePath);
            }
        }

        var assignments = await _databaseService.GetControlledBookletHashAssignmentsAsync(document.FileHash);
        var otherClients = assignments
            .Where(item => item.ClientId != document.ClientId)
            .Select(item => $"{item.ClientName} (document {item.DocumentId})")
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        if (otherClients.Count > 0)
        {
            AddIssue(
                run,
                ControlledBookletReprocessingIssueCodes.CrossClientDuplicateHash,
                $"The same source file hash is assigned to another client: {string.Join(", ", otherClients)}. Reconcile document ownership before reprocessing.",
                isBlocking: true,
                fieldName: "file_hash",
                rawSourceReference: document.FilePath);
        }
    }

    private static void AddIdentityIssues(
        ControlledBookletReprocessingRun run,
        ControlledBookletDocumentOption document,
        BookletClientIdentity? identity)
    {
        if (!BookletClientIdentityGuardrail.IsConfirmed(identity))
        {
            AddIssue(
                run,
                ControlledBookletReprocessingIssueCodes.SourceIdentityUnconfirmed,
                "The booklet did not provide a reliable client identity with at least 80% confidence.",
                isBlocking: true,
                fieldName: "client_name",
                rawSourceReference: document.FilePath);
            return;
        }

        if (!BookletClientIdentityGuardrail.ClientNamesLikelyMatch(document.ClientName, identity!.ClientName))
        {
            AddIssue(
                run,
                ControlledBookletReprocessingIssueCodes.SourceClientMismatch,
                $"The booklet identifies '{identity.ClientName}', but document {document.DocumentId} is assigned to '{document.ClientName}'.",
                isBlocking: true,
                fieldName: "client_name",
                rawSourceReference: document.FilePath);
        }

        if (!string.IsNullOrWhiteSpace(document.Carrier) &&
            !string.IsNullOrWhiteSpace(identity.Carrier) &&
            !CarrierNamesLikelyMatch(document.Carrier, identity.Carrier))
        {
            AddIssue(
                run,
                ControlledBookletReprocessingIssueCodes.SourceCarrierMismatch,
                $"The booklet carrier '{identity.Carrier}' differs from the registered carrier '{document.Carrier}'.",
                isBlocking: false,
                fieldName: "carrier",
                rawSourceReference: document.FilePath);
        }
    }

    private static void AddExtractionIssues(
        ControlledBookletReprocessingRun run,
        BookletExtractionResult extraction)
    {
        if (extraction.TotalPdfsFound != 1 || extraction.SuccessfullyProcessed != 1 || extraction.ErrorCount > 0)
        {
            AddIssue(
                run,
                ControlledBookletReprocessingIssueCodes.ExtractionFailed,
                $"Expected one successful booklet extraction. Found {extraction.TotalPdfsFound}, processed {extraction.SuccessfullyProcessed}, errors {extraction.ErrorCount}.",
                isBlocking: true,
                rawSourceReference: run.SourceFilePath);
        }

        if (extraction.ReviewRequiredCount > 0)
        {
            AddIssue(
                run,
                ControlledBookletReprocessingIssueCodes.ExtractionReviewRequired,
                $"The extraction review workbook contains {extraction.ReviewRequiredCount} row(s) requiring review.",
                isBlocking: true,
                rawSourceReference: extraction.ReviewWorkbookPath);
        }

        if (extraction.ReviewRecommendedCount > 0)
        {
            AddIssue(
                run,
                ControlledBookletReprocessingIssueCodes.ExtractionReviewRecommended,
                $"The extraction review workbook contains {extraction.ReviewRecommendedCount} medium-confidence row(s) recommended for review.",
                isBlocking: false,
                rawSourceReference: extraction.ReviewWorkbookPath);
        }
    }

    private static void AddRowIdentityIssues(
        ControlledBookletReprocessingRun run,
        ControlledBookletDocumentOption document,
        IEnumerable<PlanDesignRecord> rows)
    {
        foreach (var row in rows)
        {
            if (!string.IsNullOrWhiteSpace(row.SourceFileName) &&
                !Path.GetFileName(row.SourceFileName).Equals(document.FileName, StringComparison.OrdinalIgnoreCase))
            {
                AddIssue(
                    run,
                    ControlledBookletReprocessingIssueCodes.SourceLineageConflict,
                    $"An extracted row cites '{Path.GetFileName(row.SourceFileName)}' instead of the selected booklet '{document.FileName}'.",
                    isBlocking: true,
                    row.BenefitCategory,
                    row.BenefitName,
                    "source_file_name",
                    row.SourceFileName,
                    row.SourcePage);
            }

            if (!string.IsNullOrWhiteSpace(row.ClientName) &&
                !BookletClientIdentityGuardrail.ClientNamesLikelyMatch(document.ClientName, row.ClientName))
            {
                AddIssue(
                    run,
                    ControlledBookletReprocessingIssueCodes.SourceClientMismatch,
                    $"An extracted Plan Design row identifies '{row.ClientName}' instead of '{document.ClientName}'.",
                    isBlocking: true,
                    row.BenefitCategory,
                    row.BenefitName,
                    "client_name",
                    document.FilePath,
                    row.SourcePage);
            }
        }
    }

    private static void ValidateStagedRows(
        ControlledBookletReprocessingRun run,
        ControlledBookletDocumentOption document,
        IReadOnlyList<PlanDesignRecord> rows)
    {
        if (rows.Count == 0)
        {
            AddIssue(
                run,
                ControlledBookletReprocessingIssueCodes.NoStagedRows,
                "The extractor produced no Plan Design rows for the selected booklet.",
                isBlocking: true,
                rawSourceReference: run.MasterWorkbookPath);
            return;
        }

        AddProvisionIdentityIssues(run, rows);

        foreach (var row in rows)
        {
            var guardrail = CanonicalSourceLineageValueGuardrails.EvaluatePercentage(
                row.ReimbursementPercent,
                row.ReimbursementText);
            row.ReimbursementPercent = guardrail.NormalizedValue;
            if (guardrail.Status == CanonicalValueGuardrailStatuses.Rejected)
            {
                AddIssue(
                    run,
                    ControlledBookletReprocessingIssueCodes.PercentageOutOfRange,
                    guardrail.Message,
                    isBlocking: true,
                    row.BenefitCategory,
                    row.BenefitName,
                    "reimbursement_percent",
                    FirstNonBlank(row.ReimbursementText, row.CoverageDescription, document.FilePath),
                    row.SourcePage);
            }
            else if (guardrail.Status == CanonicalValueGuardrailStatuses.PreservedForReview)
            {
                AddIssue(
                    run,
                    ControlledBookletReprocessingIssueCodes.MultiplePercentagesPreserved,
                    guardrail.Message,
                    isBlocking: true,
                    row.BenefitCategory,
                    row.BenefitName,
                    "reimbursement_percent",
                    FirstNonBlank(row.ReimbursementText, row.CoverageDescription, document.FilePath),
                    row.SourcePage);
            }
        }

        var source = NormalizedSourceMetadata.ForFile(
            SourceTypes.Pdf,
            document.Carrier,
            document.FilePath,
            PdfExtractionMethods.ExternalBookletExtractor,
            0.80m);
        source.SourceDocumentId = document.DocumentId.ToString(CultureInfo.InvariantCulture);
        var normalized = NormalizePlanDesignData.FromRows(rows, source);
        var validationIssues = new NormalizedDataValidator().Validate(normalized);
        foreach (var issue in validationIssues)
        {
            AddIssue(
                run,
                issue.IssueCode,
                issue.Message,
                isBlocking: true,
                fieldName: issue.FieldName,
                rawSourceReference: issue.RawSourceReference);
        }
    }

    private static void AddProvisionIdentityIssues(
        ControlledBookletReprocessingRun run,
        IReadOnlyList<PlanDesignRecord> rows)
    {
        foreach (var row in rows.Where(row =>
                     string.IsNullOrWhiteSpace(row.ProvisionCode) ||
                     string.IsNullOrWhiteSpace(row.ProvisionLabel)))
        {
            AddIssue(
                run,
                ControlledBookletReprocessingIssueCodes.ProvisionIdentityMissing,
                "A Plan Design row could not be assigned a stable provision identity.",
                isBlocking: true,
                row.BenefitCategory,
                row.BenefitName,
                "provision_code",
                FirstNonBlank(row.CoverageDescription, row.SourceSection),
                row.SourcePage);
        }

        foreach (var duplicate in rows
                     .Where(row => !string.IsNullOrWhiteSpace(row.ProvisionCode))
                     .GroupBy(PlanDesignCanonicalMapper.BuildProvisionComparisonKey, StringComparer.OrdinalIgnoreCase)
                     .Where(group => group.Count() > 1))
        {
            var row = duplicate.First();
            AddIssue(
                run,
                ControlledBookletReprocessingIssueCodes.DuplicateProvisionIdentity,
                $"Provision '{row.ProvisionLabel}' occurs {duplicate.Count()} times for the same benefit and class.",
                isBlocking: true,
                row.BenefitCategory,
                row.ProvisionLabel,
                "provision_code",
                duplicate.Key,
                row.SourcePage);
        }
    }

    private async Task AddComparisonIssuesAsync(
        ControlledBookletReprocessingRun run,
        ControlledBookletDocumentOption document)
    {
        if (run.BeforeRows.Count > 0 && run.StagedRows.Count < run.BeforeRows.Count)
        {
            AddIssue(
                run,
                ControlledBookletReprocessingIssueCodes.RowLoss,
                $"The staged extraction has {run.StagedRows.Count} row(s), fewer than the current {run.BeforeRows.Count}. Row loss requires review before replacement.",
                isBlocking: true,
                fieldName: "row_count",
                rawSourceReference: document.FilePath);
        }

        var beforeBenefits = run.BeforeRows
            .Select(TopLevelBenefitKey)
            .Where(key => !string.IsNullOrWhiteSpace(key))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        var stagedBenefits = run.StagedRows
            .Select(TopLevelBenefitKey)
            .Where(key => !string.IsNullOrWhiteSpace(key))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        foreach (var missingBenefit in beforeBenefits.Except(stagedBenefits, StringComparer.OrdinalIgnoreCase))
        {
            AddIssue(
                run,
                ControlledBookletReprocessingIssueCodes.BenefitMissing,
                $"The current Plan Design contains '{missingBenefit}', but the staged extraction does not.",
                isBlocking: true,
                benefitCategory: missingBenefit,
                fieldName: "benefit_category",
                rawSourceReference: document.FilePath);
        }

        var existingIssueCodes = await _databaseService.GetPlanDesignValidationIssueCodesForDocumentAsync(document.DocumentId);
        foreach (var issue in run.Issues
                     .Where(issue => issue.IsBlocking)
                     .Where(issue => issue.IssueCode is
                         ValidationIssueCodes.PlanDesignPercentageOutOfRange or
                         ValidationIssueCodes.PlanDesignMultiplePercentagesPreserved)
                     .Where(issue => !existingIssueCodes.Contains(issue.IssueCode))
                     .ToList())
        {
            AddIssue(
                run,
                ControlledBookletReprocessingIssueCodes.NewValidationIssue,
                $"Staging introduced validation issue '{issue.IssueCode}' that was not recorded for the current document.",
                isBlocking: true,
                issue.BenefitCategory,
                issue.BenefitName,
                issue.FieldName,
                issue.RawSourceReference,
                issue.SourcePage);
        }
    }

    private async Task<string> ValidateCommitIdentityAsync(
        ControlledBookletReprocessingRun run,
        ControlledBookletDocumentOption? document,
        CancellationToken cancellationToken)
    {
        if (document is null)
        {
            return $"Registered document {run.SourceDocumentId} no longer exists.";
        }

        if (document.ClientId != run.ClientId ||
            !document.FilePath.Equals(run.SourceFilePath, StringComparison.OrdinalIgnoreCase) ||
            !document.FileHash.Equals(run.SourceFileHash, StringComparison.OrdinalIgnoreCase))
        {
            return "The registered client, file path, or source hash changed after staging. Stage the booklet again.";
        }

        if (!File.Exists(document.FilePath))
        {
            return "The staged booklet file is no longer available at its registered path.";
        }

        var currentHash = await ComputeSha256Async(document.FilePath, cancellationToken);
        if (!currentHash.Equals(run.SourceFileHash, StringComparison.OrdinalIgnoreCase))
        {
            return "The booklet file changed after staging. Stage the current file again before approval.";
        }

        var assignments = await _databaseService.GetControlledBookletHashAssignmentsAsync(run.SourceFileHash);
        if (assignments.Any(item => item.ClientId != run.ClientId))
        {
            return "The booklet hash is now assigned to another client. Reconcile document ownership before approval.";
        }

        return "";
    }

    private async Task<ControlledBookletReprocessingRun> FinalizeStageAsync(ControlledBookletReprocessingRun run)
    {
        run.StagedAt = DateTime.Now;
        if (!run.Status.Equals(ControlledBookletReprocessingStatuses.Failed, StringComparison.OrdinalIgnoreCase))
        {
            run.Status = HasBlockingIssues(run)
                ? ControlledBookletReprocessingStatuses.Blocked
                : ControlledBookletReprocessingStatuses.ReadyForApproval;
        }

        UpdateRunCounts(run);
        await _databaseService.SaveControlledBookletReprocessingStageAsync(run);
        return await ReloadRunAsync(run.RunId);
    }

    private async Task<ControlledBookletReprocessingRun> ReloadRunAsync(long runId)
    {
        return await _databaseService.GetControlledBookletReprocessingRunAsync(runId)
            ?? throw new InvalidOperationException($"Controlled booklet reprocessing run {runId} was not found.");
    }

    private static ControlledBookletReprocessingRun CreateRun(
        long runId,
        ControlledBookletDocumentOption document,
        string requestedBy)
    {
        return new ControlledBookletReprocessingRun
        {
            RunId = runId,
            SourceDocumentId = document.DocumentId,
            ClientId = document.ClientId,
            ClientName = document.ClientName,
            Carrier = document.Carrier,
            SourceFileName = document.FileName,
            SourceFilePath = document.FilePath,
            SourceFileHash = document.FileHash,
            RequestedBy = FirstNonBlank(requestedBy, Environment.UserName),
            RequestedAt = DateTime.Now,
            Status = ControlledBookletReprocessingStatuses.Extracting
        };
    }

    private static void ApplyIdentity(ControlledBookletReprocessingRun run, BookletClientIdentity? identity)
    {
        if (identity is null)
        {
            return;
        }

        run.DetectedClientName = identity.ClientName;
        run.DetectedCarrier = identity.Carrier;
        run.IdentityConfidence = identity.ConfidenceScore;
        run.IdentitySource = identity.IdentitySource;
    }

    private static void ApplyDocumentIdentity(
        IEnumerable<PlanDesignRecord> rows,
        ControlledBookletDocumentOption document)
    {
        foreach (var row in rows)
        {
            row.ClientId = document.ClientId;
            row.ClientName = document.ClientName;
            row.Carrier = FirstNonBlank(row.Carrier, document.Carrier);
            row.SourceDocumentId = document.DocumentId;
            row.SourceFileName = document.FileName;
            row.DateLoaded = DateTime.Now;
        }
    }

    private static BookletClientIdentity? SelectIdentity(
        IReadOnlyList<BookletClientIdentity> identities,
        string sourcePath)
    {
        return identities.FirstOrDefault(identity => PathsEqual(identity.SourcePath, sourcePath))
               ?? identities.FirstOrDefault();
    }

    private static IReadOnlyList<ControlledBookletReprocessingChange> BuildChanges(
        IReadOnlyList<PlanDesignRecord> beforeRows,
        IReadOnlyList<PlanDesignRecord> stagedRows)
    {
        var before = BuildComparableRows(beforeRows);
        var after = BuildComparableRows(stagedRows);
        var keys = before.Keys
            .Concat(after.Keys)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(key => key, StringComparer.OrdinalIgnoreCase);
        var changes = new List<ControlledBookletReprocessingChange>();
        foreach (var key in keys)
        {
            before.TryGetValue(key, out var oldRow);
            after.TryGetValue(key, out var newRow);
            var changeType = oldRow is null
                ? ControlledBookletReprocessingChangeTypes.Added
                : newRow is null
                    ? ControlledBookletReprocessingChangeTypes.Removed
                    : oldRow.DisplayValue.Equals(newRow.DisplayValue, StringComparison.OrdinalIgnoreCase)
                        ? ControlledBookletReprocessingChangeTypes.Unchanged
                        : ControlledBookletReprocessingChangeTypes.Changed;
            var source = newRow?.Row ?? oldRow?.Row;
            changes.Add(new ControlledBookletReprocessingChange
            {
                ChangeType = changeType,
                ComparisonKey = key,
                ClassName = source?.ClassName ?? "",
                BenefitCategory = source?.BenefitCategory ?? "",
                BenefitName = FirstNonBlank(source?.ProvisionLabel, source?.BenefitName),
                BeforeValue = oldRow?.DisplayValue ?? "",
                AfterValue = newRow?.DisplayValue ?? "",
                SourcePage = newRow?.Row.SourcePage ?? oldRow?.Row.SourcePage
            });
        }

        return changes
            .OrderBy(change => ChangeOrder(change.ChangeType))
            .ThenBy(change => change.BenefitCategory, StringComparer.OrdinalIgnoreCase)
            .ThenBy(change => change.BenefitName, StringComparer.OrdinalIgnoreCase)
            .ThenBy(change => change.ClassName, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static Dictionary<string, ComparablePlanDesignRow> BuildComparableRows(
        IEnumerable<PlanDesignRecord> rows)
    {
        var result = new Dictionary<string, ComparablePlanDesignRow>(StringComparer.OrdinalIgnoreCase);
        foreach (var row in rows)
        {
            var comparable = new ComparablePlanDesignRow(
                BuildBaseComparisonKey(row),
                BuildDisplayValue(row),
                row);
            var key = comparable.Key;
            if (result.ContainsKey(key))
            {
                key = $"{key}|duplicate|{BuildDuplicateDiscriminator(row)}";
                var collision = 2;
                while (result.ContainsKey(key))
                {
                    key = $"{comparable.Key}|duplicate|{BuildDuplicateDiscriminator(row)}|{collision++}";
                }
            }

            result[key] = comparable with { Key = key };
        }

        return result;
    }

    private static string BuildBaseComparisonKey(PlanDesignRecord row)
    {
        return PlanDesignCanonicalMapper.BuildProvisionComparisonKey(row);
    }

    private static string BuildDuplicateDiscriminator(PlanDesignRecord row)
    {
        var source = string.Join(
            "|",
            row.SourceSection,
            row.SourcePage?.ToString(CultureInfo.InvariantCulture) ?? "",
            PlanDesignCanonicalMapper.NormalizeComparableText(row.CoverageDescription));
        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(source)))[..12].ToLowerInvariant();
    }

    internal static string BuildDisplayValue(PlanDesignRecord row)
    {
        var rate = FirstNonBlank(
            row.ReimbursementText,
            row.ReimbursementPercent.HasValue ? $"{row.ReimbursementPercent.Value:0.####}%" : "");
        var coverageDescription = RemoveRepeatedStructuredLines(row, rate);
        return string.Join(
            " | ",
            new[]
            {
                string.IsNullOrWhiteSpace(rate) ? "" : $"Reimbursement: {rate}",
                string.IsNullOrWhiteSpace(row.MaximumAmount) ? "" : $"Maximum: {row.MaximumAmount}",
                string.IsNullOrWhiteSpace(row.Frequency) ? "" : $"Frequency: {row.Frequency}",
                string.IsNullOrWhiteSpace(row.WaitingPeriod) ? "" : $"Waiting: {row.WaitingPeriod}",
                string.IsNullOrWhiteSpace(row.TerminationAge) ? "" : $"Termination: {row.TerminationAge}",
                coverageDescription
            }.Where(value => !string.IsNullOrWhiteSpace(value)));
    }

    private static string RemoveRepeatedStructuredLines(PlanDesignRecord row, string rate)
    {
        var retained = row.CoverageDescription
            .Replace("\r\n", "\n", StringComparison.Ordinal)
            .Split('\n')
            .Select(line => line.Trim())
            .Where(line => !string.IsNullOrWhiteSpace(line))
            .Where(line => !MatchesStructuredLine(line, "Coinsurance", rate))
            .Where(line => !MatchesStructuredLine(line, "Reimbursement", rate))
            .Where(line => !MatchesStructuredLine(line, "Maximum", row.MaximumAmount))
            .Where(line => !MatchesStructuredLine(line, "Frequency", row.Frequency))
            .Where(line => !MatchesStructuredLine(line, "Trip Duration", row.Frequency));
        return string.Join(Environment.NewLine, retained);
    }

    private static bool MatchesStructuredLine(string line, string label, string expectedValue)
    {
        if (string.IsNullOrWhiteSpace(expectedValue))
        {
            return false;
        }

        var separator = line.IndexOf(':');
        if (separator <= 0 || !line[..separator].Trim().Equals(label, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        var actual = PlanDesignCanonicalMapper.NormalizeComparableText(line[(separator + 1)..]);
        var expected = PlanDesignCanonicalMapper.NormalizeComparableText(expectedValue);
        return actual.Equals(expected, StringComparison.Ordinal);
    }

    internal static string TopLevelBenefitKey(PlanDesignRecord row)
    {
        var code = PlanDesignCanonicalMapper.ResolveProvisionBenefitCode(row);
        return FirstNonBlank(
            code,
            PlanDesignCanonicalMapper.NormalizeBenefitKey(row.BenefitCategory),
            PlanDesignCanonicalMapper.NormalizeBenefitKey(row.BenefitName));
    }

    private static void UpdateRunCounts(ControlledBookletReprocessingRun run)
    {
        run.BeforeRowCount = run.BeforeRows.Count;
        run.StagedRowCount = run.StagedRows.Count;
        run.AddedCount = run.Changes.Count(change => change.ChangeType == ControlledBookletReprocessingChangeTypes.Added);
        run.ChangedCount = run.Changes.Count(change => change.ChangeType == ControlledBookletReprocessingChangeTypes.Changed);
        run.RemovedCount = run.Changes.Count(change => change.ChangeType == ControlledBookletReprocessingChangeTypes.Removed);
        run.BlockingIssueCount = run.Issues.Count(issue => issue.IsBlocking);
        run.WarningCount = run.Issues.Count(issue => !issue.IsBlocking);
    }

    private static void AddIssue(
        ControlledBookletReprocessingRun run,
        string issueCode,
        string message,
        bool isBlocking,
        string benefitCategory = "",
        string benefitName = "",
        string fieldName = "",
        string rawSourceReference = "",
        int? sourcePage = null)
    {
        if (run.Issues.Any(existing =>
                existing.IssueCode.Equals(issueCode, StringComparison.OrdinalIgnoreCase) &&
                existing.FieldName.Equals(fieldName, StringComparison.OrdinalIgnoreCase) &&
                existing.Message.Equals(message, StringComparison.OrdinalIgnoreCase) &&
                (existing.BenefitCategory.Equals(benefitCategory, StringComparison.OrdinalIgnoreCase) ||
                 string.IsNullOrWhiteSpace(existing.BenefitCategory) ||
                 string.IsNullOrWhiteSpace(benefitCategory)) &&
                (existing.BenefitName.Equals(benefitName, StringComparison.OrdinalIgnoreCase) ||
                 string.IsNullOrWhiteSpace(existing.BenefitName) ||
                 string.IsNullOrWhiteSpace(benefitName))))
        {
            return;
        }

        run.Issues.Add(new ControlledBookletReprocessingIssue
        {
            RunId = run.RunId,
            IssueCode = issueCode,
            Severity = isBlocking
                ? ControlledBookletReprocessingSeverities.Error
                : ControlledBookletReprocessingSeverities.Warning,
            IsBlocking = isBlocking,
            BenefitCategory = benefitCategory,
            BenefitName = benefitName,
            FieldName = fieldName,
            Message = message,
            RawSourceReference = rawSourceReference,
            SourcePage = sourcePage
        });
    }

    private static void AddEdit(
        ICollection<ControlledBookletReprocessingEdit> edits,
        long runId,
        string comparisonKey,
        string fieldName,
        string? previousValue,
        string? newValue,
        string reason,
        string editedBy,
        DateTime editedAt)
    {
        var before = previousValue?.Trim() ?? "";
        var after = newValue?.Trim() ?? "";
        if (before.Equals(after, StringComparison.Ordinal))
        {
            return;
        }

        edits.Add(new ControlledBookletReprocessingEdit
        {
            RunId = runId,
            ComparisonKey = comparisonKey,
            FieldName = fieldName,
            PreviousValue = before,
            NewValue = after,
            Reason = reason,
            EditedBy = editedBy,
            EditedAt = editedAt
        });
    }

    private static bool HasBlockingIssues(ControlledBookletReprocessingRun run) =>
        run.Issues.Any(issue => issue.IsBlocking);

    private static int ChangeOrder(string changeType) => changeType switch
    {
        ControlledBookletReprocessingChangeTypes.Changed => 0,
        ControlledBookletReprocessingChangeTypes.Removed => 1,
        ControlledBookletReprocessingChangeTypes.Added => 2,
        _ => 3
    };

    private static bool CarrierNamesLikelyMatch(string left, string right)
    {
        var leftKey = BookletClientIdentityGuardrail.NormalizeGroupKey(left)
            .Replace("FINANCIAL", "", StringComparison.OrdinalIgnoreCase)
            .Replace("INSURANCE", "", StringComparison.OrdinalIgnoreCase);
        var rightKey = BookletClientIdentityGuardrail.NormalizeGroupKey(right)
            .Replace("FINANCIAL", "", StringComparison.OrdinalIgnoreCase)
            .Replace("INSURANCE", "", StringComparison.OrdinalIgnoreCase);
        return leftKey.Equals(rightKey, StringComparison.OrdinalIgnoreCase) ||
               (Math.Min(leftKey.Length, rightKey.Length) >= 4 &&
                (leftKey.Contains(rightKey, StringComparison.OrdinalIgnoreCase) ||
                 rightKey.Contains(leftKey, StringComparison.OrdinalIgnoreCase)));
    }

    private static string BuildOutputFolder(string outputRoot, long documentId, long runId)
    {
        var root = string.IsNullOrWhiteSpace(outputRoot)
            ? Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "Renewal Review Builder",
                "Controlled Booklet Reprocessing")
            : outputRoot.Trim();
        var folder = Path.Combine(root, $"Document_{documentId}_Run_{runId}_{DateTime.Now:yyyyMMdd_HHmmss}");
        Directory.CreateDirectory(folder);
        return folder;
    }

    private static async Task<string> ComputeSha256Async(string filePath, CancellationToken cancellationToken)
    {
        await using var stream = File.OpenRead(filePath);
        var hash = await SHA256.HashDataAsync(stream, cancellationToken);
        return Convert.ToHexString(hash);
    }

    private static bool PathsEqual(string left, string right)
    {
        if (string.IsNullOrWhiteSpace(left) || string.IsNullOrWhiteSpace(right))
        {
            return false;
        }

        try
        {
            return Path.GetFullPath(left).Equals(Path.GetFullPath(right), StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return left.Equals(right, StringComparison.OrdinalIgnoreCase);
        }
    }

    private static string FirstNonBlank(params string?[] values) =>
        values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value))?.Trim() ?? "";

    private sealed record ComparablePlanDesignRow(
        string Key,
        string DisplayValue,
        PlanDesignRecord Row);
}

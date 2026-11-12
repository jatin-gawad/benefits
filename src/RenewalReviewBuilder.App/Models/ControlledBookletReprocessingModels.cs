namespace RenewalReviewBuilder.App.Models;

public static class ControlledBookletReprocessingContract
{
    public const string Version = "CANONICAL_CONTROLLED_BOOKLET_REPROCESSING_V1";
}

public static class ControlledBookletReprocessingStatuses
{
    public const string Extracting = "Extracting";
    public const string Blocked = "Blocked";
    public const string ReadyForApproval = "Ready for Approval";
    public const string Committing = "Committing";
    public const string Committed = "Committed";
    public const string Rejected = "Rejected";
    public const string Failed = "Failed";
}

public static class DocumentIdentityReconciliationStatuses
{
    public const string IdentityQuarantined = "identity_quarantined";
    public const string SupersededDuplicate = "superseded_duplicate";

    public static bool IsExcludedFromProcessing(string? status) => status is
        IdentityQuarantined or
        SupersededDuplicate;
}

public static class ControlledBookletReprocessingSeverities
{
    public const string Error = "Error";
    public const string Warning = "Warning";
    public const string Information = "Information";
}

public static class ControlledBookletReprocessingIssueCodes
{
    public const string SourceDocumentMissing = "source_document_missing";
    public const string SourceDocumentTypeInvalid = "source_document_type_invalid";
    public const string SourceDocumentQuarantined = "source_document_quarantined";
    public const string SourceFileMissing = "source_file_missing";
    public const string SourceHashMissing = "source_hash_missing";
    public const string SourceHashChanged = "source_hash_changed";
    public const string CrossClientDuplicateHash = "cross_client_duplicate_hash";
    public const string SourceIdentityUnconfirmed = "source_identity_unconfirmed";
    public const string SourceClientMismatch = "source_client_mismatch";
    public const string SourceCarrierMismatch = "source_carrier_mismatch";
    public const string ExtractionFailed = "extraction_failed";
    public const string ExtractionReviewRequired = "extraction_review_required";
    public const string ExtractionReviewRecommended = "extraction_review_recommended";
    public const string NoStagedRows = "no_staged_rows";
    public const string SourceLineageConflict = CanonicalSourceLineageValueGuardrailIssueCodes.SourceLineageConflict;
    public const string PercentageOutOfRange = CanonicalSourceLineageValueGuardrailIssueCodes.PercentageOutOfRange;
    public const string MultiplePercentagesPreserved = CanonicalSourceLineageValueGuardrailIssueCodes.MultiplePercentagesPreserved;
    public const string ProvisionIdentityMissing = "plan_design_provision_identity_missing";
    public const string DuplicateProvisionIdentity = "plan_design_duplicate_provision_identity";
    public const string RowLoss = "plan_design_row_loss";
    public const string BenefitMissing = "plan_design_benefit_missing";
    public const string NewValidationIssue = "new_validation_issue";
    public const string SourceChangedBeforeCommit = "source_changed_before_commit";
    public const string PostCommitMismatch = "post_commit_mismatch";
}

public static class ControlledBookletReprocessingChangeTypes
{
    public const string Added = "Added";
    public const string Changed = "Changed";
    public const string Removed = "Removed";
    public const string Unchanged = "Unchanged";
}

public sealed class ControlledBookletDocumentOption
{
    public long DocumentId { get; set; }
    public long ClientId { get; set; }
    public string ClientName { get; set; } = "";
    public string Carrier { get; set; } = "";
    public string DocumentType { get; set; } = "";
    public string FileName { get; set; } = "";
    public string FilePath { get; set; } = "";
    public string FileHash { get; set; } = "";
    public string ProcessingStatus { get; set; } = "";
    public int CurrentRowCount { get; set; }
    public int OutOfRangePercentageCount { get; set; }
    public int CrossClientHashAssignmentCount { get; set; }

    public string DisplayName =>
        $"{ClientName} | {FirstNonBlank(Carrier, "Carrier not recorded")} | {FileName} | Document {DocumentId}";

    public string HashPreview => string.IsNullOrWhiteSpace(FileHash)
        ? "Not recorded"
        : FileHash.Length <= 16 ? FileHash : $"{FileHash[..16]}...";

    public string ReviewSummary => CrossClientHashAssignmentCount > 0
        ? $"Duplicate source assigned to {CrossClientHashAssignmentCount} other client(s)"
        : OutOfRangePercentageCount > 0
            ? $"{OutOfRangePercentageCount} out-of-range percentage row(s)"
            : "No known percentage errors";

    private static string FirstNonBlank(params string[] values) =>
        values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value))?.Trim() ?? "";
}

public sealed class ControlledBookletReprocessingRequest
{
    public long SourceDocumentId { get; init; }
    public string OutputRoot { get; init; } = "";
    public string RequestedBy { get; init; } = Environment.UserName;
}

public sealed class ControlledBookletReprocessingRun
{
    public long RunId { get; set; }
    public string WorkflowVersion { get; set; } = ControlledBookletReprocessingContract.Version;
    public long SourceDocumentId { get; set; }
    public long ClientId { get; set; }
    public string ClientName { get; set; } = "";
    public string Carrier { get; set; } = "";
    public string SourceFileName { get; set; } = "";
    public string SourceFilePath { get; set; } = "";
    public string SourceFileHash { get; set; } = "";
    public string DetectedClientName { get; set; } = "";
    public string DetectedCarrier { get; set; } = "";
    public decimal? IdentityConfidence { get; set; }
    public string IdentitySource { get; set; } = "";
    public string Status { get; set; } = ControlledBookletReprocessingStatuses.Extracting;
    public string RequestedBy { get; set; } = Environment.UserName;
    public DateTime RequestedAt { get; set; } = DateTime.Now;
    public DateTime? StagedAt { get; set; }
    public string ApprovedBy { get; set; } = "";
    public DateTime? ApprovedAt { get; set; }
    public DateTime? CommittedAt { get; set; }
    public string RejectedBy { get; set; } = "";
    public DateTime? RejectedAt { get; set; }
    public string OutputFolder { get; set; } = "";
    public string MasterWorkbookPath { get; set; } = "";
    public string ReviewWorkbookPath { get; set; } = "";
    public string BackupPath { get; set; } = "";
    public int BeforeRowCount { get; set; }
    public int StagedRowCount { get; set; }
    public int PostCommitRowCount { get; set; }
    public int AddedCount { get; set; }
    public int ChangedCount { get; set; }
    public int RemovedCount { get; set; }
    public int BlockingIssueCount { get; set; }
    public int WarningCount { get; set; }
    public string Notes { get; set; } = "";
    public string ErrorMessage { get; set; } = "";
    public List<PlanDesignRecord> BeforeRows { get; } = [];
    public List<PlanDesignRecord> StagedRows { get; } = [];
    public List<BenefitAdministrationExtractionCandidate> BenefitAdministrationRecords { get; } = [];
    public List<ControlledBookletReprocessingIssue> Issues { get; } = [];
    public List<ControlledBookletReprocessingChange> Changes { get; } = [];
    public List<ControlledBookletReprocessingEdit> Edits { get; } = [];

    public bool CanApprove =>
        Status.Equals(ControlledBookletReprocessingStatuses.ReadyForApproval, StringComparison.OrdinalIgnoreCase) &&
        BlockingIssueCount == 0 &&
        StagedRowCount > 0;

    public bool CanReject => Status is
        ControlledBookletReprocessingStatuses.Extracting or
        ControlledBookletReprocessingStatuses.Blocked or
        ControlledBookletReprocessingStatuses.ReadyForApproval;

    public bool CanEdit =>
        StagedRowCount > 0 &&
        Status is ControlledBookletReprocessingStatuses.Blocked or
            ControlledBookletReprocessingStatuses.ReadyForApproval;

    public string DifferenceSummary =>
        $"Added {AddedCount} | Changed {ChangedCount} | Removed {RemovedCount}";

    public string ValidationSummary => BlockingIssueCount > 0
        ? $"{BlockingIssueCount} blocking issue(s), {WarningCount} warning(s)"
        : $"No blockers, {WarningCount} warning(s)";
}

public sealed class ControlledBookletReprocessingIssue
{
    public long IssueId { get; set; }
    public long RunId { get; set; }
    public string IssueCode { get; set; } = "";
    public string Severity { get; set; } = ControlledBookletReprocessingSeverities.Warning;
    public bool IsBlocking { get; set; }
    public string BenefitCategory { get; set; } = "";
    public string BenefitName { get; set; } = "";
    public string FieldName { get; set; } = "";
    public string Message { get; set; } = "";
    public string RawSourceReference { get; set; } = "";
    public int? SourcePage { get; set; }
}

public sealed class ControlledBookletReprocessingChange
{
    public long ChangeId { get; set; }
    public long RunId { get; set; }
    public string ChangeType { get; set; } = "";
    public string ComparisonKey { get; set; } = "";
    public string ClassName { get; set; } = "";
    public string BenefitCategory { get; set; } = "";
    public string BenefitName { get; set; } = "";
    public string BeforeValue { get; set; } = "";
    public string AfterValue { get; set; } = "";
    public int? SourcePage { get; set; }
}

public sealed class ControlledBookletStagedEditRequest
{
    public string ComparisonKey { get; init; } = "";
    public decimal? ReimbursementPercent { get; init; }
    public string MaximumAmount { get; init; } = "";
    public string Frequency { get; init; } = "";
    public string WaitingPeriod { get; init; } = "";
    public string TerminationAge { get; init; } = "";
    public string CoverageDescription { get; init; } = "";
    public string Notes { get; init; } = "";
    public string Reason { get; init; } = "";
    public string EditedBy { get; init; } = Environment.UserName;
}

public sealed class ControlledBookletReprocessingEdit
{
    public long EditId { get; set; }
    public long RunId { get; set; }
    public string ComparisonKey { get; set; } = "";
    public string FieldName { get; set; } = "";
    public string PreviousValue { get; set; } = "";
    public string NewValue { get; set; } = "";
    public string Reason { get; set; } = "";
    public string EditedBy { get; set; } = "";
    public DateTime EditedAt { get; set; }
}

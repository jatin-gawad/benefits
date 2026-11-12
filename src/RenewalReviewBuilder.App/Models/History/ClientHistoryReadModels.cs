namespace RenewalReviewBuilder.App.Models.History;

public sealed class ClientHistoryEventReadModel
{
    public string EventId { get; init; } = "";
    public string StableSourceKey { get; init; } = "";
    public long ClientId { get; init; }
    public string SourceModule { get; init; } = "";
    public string SourceEntityType { get; init; } = "";
    public string SourceEntityId { get; init; } = "";
    public DateOnly? EffectiveDate { get; init; }
    public HistoryDateReference EffectiveDateReference { get; init; } = new();
    public DateTime? RecordedDate { get; init; }
    public string EventType { get; init; } = "";
    public string Category { get; init; } = "";
    public string Title { get; init; } = "";
    public string Summary { get; init; } = "";
    public string BusinessRationale { get; init; } = "";
    public string BusinessImpact { get; init; } = "";
    public HistoryChangeSet ChangeSet { get; init; } = new();
    public string FinancialImpact { get; init; } = "";
    public string RecruitmentImpact { get; init; } = "";
    public string RetentionImpact { get; init; } = "";
    public string AdministrationImpact { get; init; } = "";
    public string EmployeeExperienceImpact { get; init; } = "";
    public string RequestedBy { get; init; } = "";
    public string ApprovedBy { get; init; } = "";
    public string Notes { get; init; } = "";
    public IReadOnlyList<string> RelatedRecords { get; init; } = [];
    public IReadOnlyList<HistoryEvidenceReference> Evidence { get; init; } = [];
    public decimal Confidence { get; init; }
    public HistoryValidationStatus ValidationStatus { get; init; }
    public HistoryOrigin Origin { get; init; }
    public bool IsExcluded { get; init; }
    public bool IsManuallyEdited { get; init; }
    public DateTime CreatedAt { get; init; }
    public DateTime UpdatedAt { get; init; }
}

public sealed class HistoryChangeSet { public IReadOnlyList<HistoryChangeItem> Changes { get; init; } = []; }
public sealed class HistoryChangeItem { public string FieldKey { get; init; } = ""; public string DisplayLabel { get; init; } = ""; public HistoryValueType ValueType { get; init; } = HistoryValueType.Text; public string PreviousValue { get; init; } = ""; public string NewValue { get; init; } = ""; public string Unit { get; init; } = ""; public string PreviousReferenceId { get; init; } = ""; public string NewReferenceId { get; init; } = ""; }
public sealed class HistoryEvidenceReference { public string EventId { get; init; } = ""; public long? DocumentId { get; init; } public long? DocumentVersionId { get; init; } public string SourceFactId { get; init; } = ""; public string SourceEntityId { get; init; } = ""; public int? PageNumber { get; init; } public string SheetName { get; init; } = ""; public string CellRange { get; init; } = ""; public string SectionName { get; init; } = ""; public string ExtractedTextReference { get; init; } = ""; public decimal Confidence { get; init; } public HistoryValidationStatus ValidationStatus { get; init; } }
public sealed class ManualClientHistoryEvent { public long Id { get; init; } public long ClientId { get; init; } public DateOnly? EffectiveDate { get; init; } public string EventType { get; init; } = ""; public string Category { get; init; } = ""; public string Title { get; init; } = ""; public string Summary { get; init; } = ""; public string BusinessRationale { get; init; } = ""; public string BusinessImpact { get; init; } = ""; public string Notes { get; init; } = ""; public string CreatedBy { get; init; } = ""; public DateTime CreatedAt { get; init; } public DateTime UpdatedAt { get; init; } public bool IsExcluded { get; init; } }

namespace RenewalReviewBuilder.App.Models;

public static class ClientHistoryPreviewStatuses
{
    public const string Parsed = "Parsed";
    public const string NeedsReview = "Needs Review";
    public const string Failed = "Failed";
    public const string Approved = "Approved";
    public const string Rejected = "Rejected";
    public const string Imported = "Imported";
}

public sealed class ClientHistoryPreviewEvent
{
    public DateOnly? EffectiveDate { get; set; }
    public string EventType { get; set; } = "";
    public string Title { get; set; } = "";
    public string Description { get; set; } = "";
    public string PreviousValue { get; set; } = "";
    public string NewValue { get; set; } = "";
    public string BenefitCategory { get; set; } = "";
    public string SourceSheet { get; set; } = "";
    public int SourceRow { get; set; }
    public string SourceFile { get; set; } = "";
    public string OriginalText { get; set; } = "";
    public string ExtractionMethod { get; set; } = "";
    public decimal Confidence { get; set; }
    public string ValidationMessages { get; set; } = "";
    public string PreviewStatus { get; set; } = ClientHistoryPreviewStatuses.Parsed;
    public string DuplicateStatus { get; set; } = "";
}

public sealed class ClientHistoryEventRecord
{
    public long HistoryEventId { get; set; }
    public long ClientId { get; set; }
    public DateOnly? EffectiveDate { get; set; }
    public string EventType { get; set; } = "";
    public string BenefitCategory { get; set; } = "";
    public string ChangeType { get; set; } = "";
    public string Title { get; set; } = "";
    public string Description { get; set; } = "";
    public string PreviousValue { get; set; } = "";
    public string NewValue { get; set; } = "";
    public long? SourceDocumentId { get; set; }
    public string SourceSheet { get; set; } = "";
    public int SourceRow { get; set; }
    public string OriginalText { get; set; } = "";
    public string ExtractionMethod { get; set; } = "";
    public decimal Confidence { get; set; }
    public string ReviewStatus { get; set; } = "";
    public DateTime CreatedAt { get; set; }
    public string CreatedBy { get; set; } = "";
    public DateTime UpdatedAt { get; set; }
    public string UpdatedBy { get; set; } = "";
    public long? ImportBatchId { get; set; }
    public string SourceFingerprint { get; set; } = "";
}

public sealed class ClientHistoryImportResult
{
    public int Approved { get; set; }
    public int Imported { get; set; }
    public int SkippedDuplicates { get; set; }
    public int Rejected { get; set; }
    public int Failed { get; set; }
    public List<string> Warnings { get; } = [];
}

public sealed class ClientHistoryWorkbookPreview
{
    public string SourceFile { get; init; } = "";
    public List<ClientHistoryPreviewEvent> Records { get; } = [];
    public IReadOnlyDictionary<string, int> CountsByEventType => Records
        .Where(record => record.PreviewStatus != ClientHistoryPreviewStatuses.Failed)
        .GroupBy(record => record.EventType, StringComparer.OrdinalIgnoreCase)
        .ToDictionary(group => group.Key, group => group.Count(), StringComparer.OrdinalIgnoreCase);
}

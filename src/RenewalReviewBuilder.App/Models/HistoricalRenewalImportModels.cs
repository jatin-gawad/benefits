namespace RenewalReviewBuilder.App.Models;

public sealed class HistoricalRenewalImportRequest
{
    public long ClientId { get; init; }
    public string RenewalPdfPath { get; init; } = "";
    public string SupportingWorkbookPath { get; init; } = "";
    public string Carrier { get; init; } = "";
    public DateOnly EffectiveDate { get; init; }
    public int RenewalYear { get; init; }
    public decimal? RequestedIncreasePercent { get; init; }
    public decimal FinalAcceptedIncreasePercent { get; init; }
    public decimal? NegotiatedSavingsAnnual { get; init; }
    public string ImportedBy { get; init; } = "Desktop user";
}

public sealed class HistoricalRenewalImportPreview
{
    public bool CanImport { get; init; }
    public string Message { get; init; } = "";
    public ClientRecord? Client { get; init; }
    public DateOnly? CurrentRenewalEffectiveDate { get; init; }
    public bool IsHistoricalDate { get; init; }
    public bool IsDuplicateSource { get; init; }
    public bool IsDuplicateRenewal { get; init; }
    public string RenewalPdfHash { get; init; } = "";
    public IReadOnlyList<string> Checks { get; init; } = [];
}

public sealed class HistoricalRenewalImportResult
{
    public bool Succeeded { get; init; }
    public bool WasDuplicate { get; init; }
    public string Message { get; init; } = "";
    public long? DocumentId { get; init; }
    public long? RenewalId { get; init; }
    public long? RenewalVersionId { get; init; }
}

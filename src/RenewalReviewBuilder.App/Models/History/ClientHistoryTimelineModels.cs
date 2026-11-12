namespace RenewalReviewBuilder.App.Models.History;

public sealed class ClientHistoryFilter
{
    public string Search { get; init; } = "";
    public string Category { get; init; } = "All Events";
    public DateOnly? From { get; init; }
    public DateOnly? To { get; init; }
    public bool OldestFirst { get; init; }
    public bool GroupByYear { get; init; }
}

public sealed class ClientHistoryTimelineRow
{
    public string EventId { get; init; } = "";
    public DateOnly? EffectiveDate { get; init; }
    public string EffectiveDateDisplay { get; init; } = "";
    public string IconGlyph { get; init; } = "&#xE81C;";
    public string AccentColor { get; init; } = "#18A55B";
    public string AccentBackground { get; init; } = "#EAF8F0";
    public string EventType { get; init; } = "";
    public string Title { get; init; } = "";
    public string Summary { get; init; } = "";
    public string PrimaryImpact { get; init; } = "";
    public int EvidenceCount { get; init; }
    public decimal Confidence { get; init; }
    public bool IsSelected { get; init; }
}

public sealed class ClientHistoryKpis
{
    public int? YearsAsClient { get; init; }
    public decimal? TotalValueCreated { get; init; }
    public int BenefitImprovementCount { get; init; }
    public int SuccessfulNegotiationCount { get; init; }
    public decimal? AverageRenewalIncrease { get; init; }
    public int CompletedRenewalCount { get; init; }
    public decimal? IndustryBenchmark { get; init; }
    public string CurrentCarrier { get; init; } = "Not available";
    public DateOnly? CarrierEffectiveDate { get; init; }
}

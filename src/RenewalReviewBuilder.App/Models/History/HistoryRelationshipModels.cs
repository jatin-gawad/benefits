using System.Globalization;

namespace RenewalReviewBuilder.App.Models.History;

public sealed class HistoryRelationshipMilestone
{
    public string EventId { get; init; } = "";
    public DateOnly? EffectiveDate { get; init; }
    public HistoryDateReference DateReference { get; init; } = new();
    public string Label { get; init; } = "";
    public string IconGlyph { get; init; } = "\uE81C";
    public string AccentColor { get; init; } = "#18A55B";
    public string AccentBackground { get; init; } = "#EAF8F0";
    public string ScreenReaderDescription { get; init; } = "";
    public bool IsSelected { get; init; }
}

public sealed class HistoryEvolutionActivityItem
{
    public string EventId { get; init; } = "";
    public HistoryDateReference DateReference { get; init; } = new();
    public string Category { get; init; } = "";
    public string Title { get; init; } = "";
    public string Summary { get; init; } = "";
    public string AccentColor { get; init; } = "#0B5DE8";
    public string AccentBackground { get; init; } = "#EEF4FF";
    public bool IsSelected { get; init; }
}
public sealed class HistoryRenewalEvolutionPoint
{
    public string EventId { get; init; } = "";
    public long? RenewalId { get; init; }
    public long? CarrierId { get; init; }
    public HistoryDateReference DateReference { get; init; } = new();
    public int? Year => DateReference.Year;
    public decimal? RequestedIncrease { get; init; }
    public decimal? NegotiatedIncrease { get; init; }
    public decimal? FinalIncrease { get; init; }
    public decimal? IndustryComparison { get; init; }
    public decimal? NegotiatedSavingsAnnual { get; init; }
    public bool IsCarrierChange { get; init; }
    public bool IsCurrent { get; init; }
    public string RenewalStatus { get; init; } = "";
    public string DetailDescription { get; init; } = "";
    public bool IsSelected { get; init; }

    public string EffectiveDateDisplay => DateReference.DisplayText;
    public string RequestedIncreaseDisplay => PercentOrUnavailable(RequestedIncrease);
    public string NegotiatedIncreaseDisplay => PercentOrUnavailable(NegotiatedIncrease);
    public string FinalIncreaseDisplay => PercentOrUnavailable(FinalIncrease);
    public string IndustryComparisonDisplay => PercentOrUnavailable(IndustryComparison);
    public string NegotiatedSavingsAnnualDisplay => NegotiatedSavingsAnnual.HasValue
        ? NegotiatedSavingsAnnual.Value.ToString("C0", CultureInfo.GetCultureInfo("en-CA"))
        : "Unavailable";

    private static string PercentOrUnavailable(decimal? value) => value.HasValue
        ? value.Value.ToString("P1", CultureInfo.InvariantCulture)
        : "Unavailable";
}
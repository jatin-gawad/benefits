using System.Globalization;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public enum AsoFundingStatusKind
{
    Unavailable,
    Surplus,
    Balanced,
    WellFunded,
    AttentionNeeded
}

public sealed class AsoFundingStatusResult
{
    public AsoFundingStatusKind Status { get; init; }
    public decimal? SourceBalance { get; init; }
    public decimal? AvailableSurplus { get; init; }
    public decimal? RequiredFloat { get; init; }
    public decimal? FundingMargin { get; init; }
    public decimal? FundingRatio { get; init; }
    public bool IsPositive { get; init; }
    public bool NeedsFundingAttention { get; init; }
    public string StatusTitle { get; init; } = "";
    public string Detail { get; init; } = "";
    public string Insight { get; init; } = "";
    public string Recommendation { get; init; } = "";
}

/// <summary>
/// Interprets the carrier ASO balance convention without changing stored source values.
/// Negative current balances are source surpluses; positive current balances are deficits.
/// </summary>
public static class AsoFundingStatusService
{
    private static readonly CultureInfo CanadianCulture = CultureInfo.GetCultureInfo("en-CA");

    public static AsoFundingStatusResult Evaluate(AsoFinancialActivityPeriod? period)
    {
        var balance = period?.CurrentBalance;
        if (!balance.HasValue)
        {
            return new AsoFundingStatusResult
            {
                Status = AsoFundingStatusKind.Unavailable,
                RequiredFloat = period?.RequiredFloat,
                StatusTitle = "Funding review required",
                Detail = "Current balance is unavailable in the latest saved statement.",
                Insight = "\u2022 Funding ratio is waiting for a saved current balance and required float.",
                Recommendation = "Import or confirm ASO activity"
            };
        }

        var requiredFloat = period!.RequiredFloat;
        var sourceStatus = period.BalanceStatus.Trim();
        var sourceSaysDeficit = sourceStatus.Equals("Deficit", StringComparison.OrdinalIgnoreCase);
        var sourceSaysSurplus = sourceStatus.Equals("Surplus", StringComparison.OrdinalIgnoreCase) ||
                                sourceStatus.Equals("Funds Available", StringComparison.OrdinalIgnoreCase);
        var availableSurplusFromSource = sourceSaysSurplus
            ? Math.Abs(balance.Value)
            : balance.Value < 0m ? -balance.Value : 0m;
        if (requiredFloat is < 0m)
        {
            return new AsoFundingStatusResult
            {
                Status = AsoFundingStatusKind.Unavailable,
                SourceBalance = balance,
                AvailableSurplus = availableSurplusFromSource,
                RequiredFloat = requiredFloat,
                StatusTitle = "Funding review required",
                Detail = $"Required float cannot be negative ({Money(requiredFloat.Value)}). Confirm the source value before evaluating funding sufficiency.",
                Insight = "\u2022 Funding ratio is unavailable because the required float is invalid.",
                Recommendation = "Confirm required float"
            };
        }

        if (sourceSaysDeficit)
        {
            return Deficit(balance.Value, Math.Abs(balance.Value), requiredFloat);
        }

        if (!sourceSaysSurplus && balance.Value > 0m)
        {
            return Deficit(balance.Value, balance.Value, requiredFloat);
        }

        var availableSurplus = sourceSaysSurplus ? Math.Abs(balance.Value) : -balance.Value;
        if (!requiredFloat.HasValue)
        {
            return availableSurplus > 0m
                ? new AsoFundingStatusResult
                {
                    Status = AsoFundingStatusKind.Surplus,
                    SourceBalance = balance,
                    AvailableSurplus = availableSurplus,
                    StatusTitle = "Surplus Recorded",
                    Detail = $"Surplus: {Money(availableSurplus)}. Required float is unavailable, so funding sufficiency cannot be calculated.",
                    Insight = "\u2022 Funding ratio is unavailable until a required float is recorded.",
                    Recommendation = "Confirm required float",
                    IsPositive = true
                }
                : Balanced(requiredFloat, null, null);
        }

        if (requiredFloat.Value <= 0m)
        {
            return availableSurplus > 0m
                ? new AsoFundingStatusResult
                {
                    Status = AsoFundingStatusKind.Surplus,
                    SourceBalance = balance,
                    AvailableSurplus = availableSurplus,
                    RequiredFloat = requiredFloat,
                    FundingMargin = availableSurplus,
                    StatusTitle = "Surplus Recorded",
                    Detail = $"Surplus: {Money(availableSurplus)}; no positive required float is recorded for the latest period.",
                    Insight = "\u2022 Funding ratio is not applicable because the latest required float is $0.00.",
                    Recommendation = "Maintain and monitor funding",
                    IsPositive = true
                }
                : Balanced(requiredFloat, 0m, null);
        }

        var margin = availableSurplus - requiredFloat.Value;
        var ratio = availableSurplus / requiredFloat.Value;
        if (margin >= 0m)
        {
            return new AsoFundingStatusResult
            {
                Status = AsoFundingStatusKind.WellFunded,
                SourceBalance = balance,
                AvailableSurplus = availableSurplus,
                RequiredFloat = requiredFloat,
                FundingMargin = margin,
                FundingRatio = ratio,
                StatusTitle = "Well Funded",
                Detail = $"Surplus: {Money(margin)} above the latest required float.",
                Insight = $"\u2022 Funding ratio: {Percent(ratio)} ({Money(availableSurplus)} available surplus / {Money(requiredFloat.Value)} required float).",
                Recommendation = "Maintain and monitor funding",
                IsPositive = true
            };
        }

        return new AsoFundingStatusResult
        {
            Status = AsoFundingStatusKind.AttentionNeeded,
            SourceBalance = balance,
            AvailableSurplus = availableSurplus,
            RequiredFloat = requiredFloat,
            FundingMargin = margin,
            FundingRatio = ratio,
            StatusTitle = "Funding Attention Needed",
            Detail = $"Shortfall: {Money(Math.Abs(margin))} against the latest required float.",
            Insight = $"\u2022 Funding ratio: {Percent(ratio)} ({Money(availableSurplus)} available surplus / {Money(requiredFloat.Value)} required float).",
            Recommendation = "Review funding requirements",
            NeedsFundingAttention = true
        };
    }

    private static AsoFundingStatusResult Deficit(decimal sourceBalance, decimal deficitAmount, decimal? requiredFloat)
    {
        var comparableFloat = requiredFloat is > 0m ? requiredFloat.Value : 0m;
        var fundingNeed = deficitAmount + comparableFloat;
        var floatDetail = !requiredFloat.HasValue
            ? " Required float is unavailable."
            : requiredFloat.Value > 0m
                ? $" Required float: {Money(requiredFloat.Value)}."
                : " No positive required float is recorded.";

        return new AsoFundingStatusResult
        {
            Status = AsoFundingStatusKind.AttentionNeeded,
            SourceBalance = sourceBalance,
            AvailableSurplus = 0m,
            RequiredFloat = requiredFloat,
            FundingMargin = -fundingNeed,
            FundingRatio = requiredFloat is > 0m ? 0m : null,
            StatusTitle = "Funding Attention Needed",
            Detail = $"Deficit: {Money(deficitAmount)} in the latest saved statement.{floatDetail}",
            Insight = requiredFloat is > 0m
                ? $"\u2022 Funding need: {Money(fundingNeed)} including the source deficit and required float."
                : "\u2022 Funding ratio is unavailable because no positive required float is recorded.",
            Recommendation = "Review funding requirements",
            NeedsFundingAttention = true
        };
    }

    private static AsoFundingStatusResult Balanced(decimal? requiredFloat, decimal? margin, decimal? ratio)
    {
        var hasPositiveFloat = requiredFloat is > 0m;
        return new AsoFundingStatusResult
        {
            Status = hasPositiveFloat ? AsoFundingStatusKind.AttentionNeeded : AsoFundingStatusKind.Balanced,
            SourceBalance = 0m,
            AvailableSurplus = 0m,
            RequiredFloat = requiredFloat,
            FundingMargin = margin,
            FundingRatio = ratio,
            StatusTitle = hasPositiveFloat ? "Funding Attention Needed" : "Balanced Position",
            Detail = hasPositiveFloat
                ? $"Shortfall: {Money(requiredFloat!.Value)} against the latest required float."
                : requiredFloat.HasValue
                    ? "Latest confirmed balance is balanced; no positive required float is recorded."
                    : "Latest confirmed balance is balanced. Required float is unavailable.",
            Insight = hasPositiveFloat
                ? $"\u2022 Funding ratio: 0.0% ($0.00 available surplus / {Money(requiredFloat!.Value)} required float)."
                : "\u2022 Funding ratio is unavailable because no positive required float is recorded.",
            Recommendation = hasPositiveFloat ? "Review funding requirements" : "Maintain and monitor funding",
            IsPositive = !hasPositiveFloat,
            NeedsFundingAttention = hasPositiveFloat
        };
    }

    private static string Money(decimal value) =>
        value.ToString("$#,0.00;($#,0.00);$0.00", CanadianCulture);

    private static string Percent(decimal value) =>
        value.ToString("P1", CanadianCulture);
}

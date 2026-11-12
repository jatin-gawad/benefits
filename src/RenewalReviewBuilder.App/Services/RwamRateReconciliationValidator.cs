using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed class RwamRateReconciliationValidator
{
    public RwamRateReconciliationComparison Validate(
        RwamExpectedRateValue expected,
        RateLine parsed,
        RateLine canonical,
        RateLine displayed)
    {
        var parsedRate = PeriodRate(parsed, expected.Period);
        var displayedRate = PeriodRate(displayed, expected.Period);
        var parsedPremium = PeriodPremium(parsed, expected.Period);
        var displayedPremium = PeriodPremium(displayed, expected.Period);
        var canonicalRate = PeriodRate(canonical, expected.Period);
        var recalculated = expected.ExpectedStatus.Equals(
            RateLineValidationStatuses.NeedsReview,
            StringComparison.OrdinalIgnoreCase)
            ? null
            : IndependentlyRecalculatePremium(expected, canonicalRate);

        var rateVariance = parsedRate.HasValue ? parsedRate.Value - expected.SourceRate : (decimal?)null;
        var premiumVariance = parsedPremium.HasValue ? parsedPremium.Value - expected.SourcePremium : (decimal?)null;
        var rateMatches = parsedRate.HasValue && displayedRate.HasValue &&
                          decimal.Round(parsedRate.Value, expected.PrintedRatePrecision) == expected.SourceRate &&
                          decimal.Round(displayedRate.Value, expected.PrintedRatePrecision) == expected.SourceRate;
        var premiumMatches = parsedPremium == expected.SourcePremium &&
                             displayedPremium == expected.SourcePremium;
        var basisMatches = canonical.RateBasis.Equals(expected.RateBasis, StringComparison.OrdinalIgnoreCase);
        var calculationMatches = expected.ExpectedStatus.Equals(
                                     RateLineValidationStatuses.NeedsReview,
                                     StringComparison.OrdinalIgnoreCase) ||
                                 recalculated == expected.SourcePremium;

        var status = expected.ExpectedStatus.Equals(
            RateLineValidationStatuses.NeedsReview,
            StringComparison.OrdinalIgnoreCase)
            ? "NEEDS_REVIEW_EXPECTED"
            : rateMatches && premiumMatches && basisMatches && calculationMatches
                ? "MATCH"
                : "FAIL";

        var explanation = status switch
        {
            "MATCH" => "Source, parser, canonical mapping, displayed values, basis, and independent premium recalculation agree.",
            "NEEDS_REVIEW_EXPECTED" => parsed.ValidationMessage,
            _ => BuildFailureExplanation(rateMatches, premiumMatches, basisMatches, calculationMatches)
        };

        return new RwamRateReconciliationComparison
        {
            SourceFile = expected.SourceFile,
            Page = expected.Page,
            ClassDivision = expected.ClassDivision,
            Benefit = expected.Benefit,
            Tier = expected.Tier,
            Period = expected.Period,
            RateBasis = canonical.RateBasis,
            SourceRate = expected.SourceRate,
            ParsedRate = parsedRate,
            RateReconciliationRate = displayedRate,
            SourcePremium = expected.SourcePremium,
            RecalculatedPremium = recalculated,
            RateReconciliationPremium = displayedPremium,
            RateVariance = rateVariance,
            PremiumVariance = premiumVariance,
            Status = status,
            Explanation = explanation
        };
    }

    private static decimal? PeriodRate(RateLine line, string period)
    {
        return period.Equals("Current", StringComparison.OrdinalIgnoreCase)
            ? line.CurrentRate
            : line.RenewalRate;
    }

    private static decimal? PeriodPremium(RateLine line, string period)
    {
        return period.Equals("Current", StringComparison.OrdinalIgnoreCase)
            ? line.CurrentPremium
            : line.RenewalPremium;
    }

    private static decimal? IndependentlyRecalculatePremium(RwamExpectedRateValue expected, decimal? rate)
    {
        if (!rate.HasValue)
        {
            return null;
        }

        var raw = expected.RateBasis switch
        {
            "per_1000_volume" => expected.Quantity * rate.Value / 1000m,
            "per_100_volume" => expected.Quantity * rate.Value / 100m,
            "per_10_volume" => expected.Quantity * rate.Value / 10m,
            "per_life" => expected.Quantity * rate.Value,
            _ => (decimal?)null
        };
        return raw.HasValue
            ? decimal.Round(raw.Value, 2, MidpointRounding.AwayFromZero)
            : null;
    }

    private static string BuildFailureExplanation(
        bool rateMatches,
        bool premiumMatches,
        bool basisMatches,
        bool calculationMatches)
    {
        var issues = new List<string>();
        if (!rateMatches) issues.Add("rate mismatch");
        if (!premiumMatches) issues.Add("premium mismatch");
        if (!basisMatches) issues.Add("rate-basis mismatch");
        if (!calculationMatches) issues.Add("premium recalculation mismatch");
        return string.Join("; ", issues);
    }
}

public sealed class RwamExpectedRateValue
{
    public string SourceFile { get; set; } = "";
    public int Page { get; set; }
    public string ClassDivision { get; set; } = "";
    public string Benefit { get; set; } = "";
    public string Tier { get; set; } = "";
    public string Period { get; set; } = "";
    public string RateBasis { get; set; } = "";
    public decimal Quantity { get; set; }
    public decimal SourceRate { get; set; }
    public int PrintedRatePrecision { get; set; }
    public decimal SourcePremium { get; set; }
    public string ExpectedStatus { get; set; } = RateLineValidationStatuses.Supported;
}

public sealed class RwamRateReconciliationComparison
{
    public string SourceFile { get; set; } = "";
    public int Page { get; set; }
    public string ClassDivision { get; set; } = "";
    public string Benefit { get; set; } = "";
    public string Tier { get; set; } = "";
    public string Period { get; set; } = "";
    public string RateBasis { get; set; } = "";
    public decimal SourceRate { get; set; }
    public decimal? ParsedRate { get; set; }
    public decimal? RateReconciliationRate { get; set; }
    public decimal SourcePremium { get; set; }
    public decimal? RecalculatedPremium { get; set; }
    public decimal? RateReconciliationPremium { get; set; }
    public decimal? RateVariance { get; set; }
    public decimal? PremiumVariance { get; set; }
    public string Status { get; set; } = "";
    public string Explanation { get; set; } = "";
}

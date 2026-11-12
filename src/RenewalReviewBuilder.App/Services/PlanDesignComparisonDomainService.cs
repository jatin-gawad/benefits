using System.Globalization;
using System.Text.RegularExpressions;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed class PlanDesignComparisonDomainService
{
    public PlanDesignComparisonClassification Classify(
        string provisionName,
        string currentValue,
        string selectedValue,
        string canonicalStatus,
        string currentAvailability,
        string selectedAvailability)
    {
        if (RequiresReview(canonicalStatus, currentAvailability, selectedAvailability))
        {
            return new(PlanDesignComparisonStatuses.ReviewRequired, "Source validation or extraction status requires review.", "Review Required");
        }

        if (IsMissing(currentValue) && IsMissing(selectedValue))
        {
            return new(PlanDesignComparisonStatuses.NotComparable, "Neither plan has a comparable normalized value.", "Not available");
        }

        if (IsMissing(currentValue))
        {
            return new(PlanDesignComparisonStatuses.Added, "The selected plan has a provision not found in the current plan.", "Informational");
        }

        if (IsMissing(selectedValue))
        {
            return new(PlanDesignComparisonStatuses.Removed, "The current-plan provision was not found in the selected plan.", "Review Required");
        }

        if (canonicalStatus.Equals(CurrentVsQuotePlanDesignStatuses.Same, StringComparison.OrdinalIgnoreCase) ||
            Normalize(currentValue).Equals(Normalize(selectedValue), StringComparison.OrdinalIgnoreCase))
        {
            return new(PlanDesignComparisonStatuses.Same, "Normalized values match.", "Informational");
        }

        if (TryClassifyNumericProvision(provisionName, currentValue, selectedValue, out var classification))
        {
            return classification;
        }

        return new(PlanDesignComparisonStatuses.Changed, "Values differ, but the saved comparison data does not support a directional coverage conclusion.", "Informational");
    }

    public static string FormatCost(PlanDesignComparisonCostInput? input)
    {
        if (input is null || input.Kind == PlanDesignComparisonCostKinds.NotAvailable)
        {
            return "Not available";
        }

        if (input.Kind == PlanDesignComparisonCostKinds.Preview)
        {
            return "Preview";
        }

        if (!input.Amount.HasValue)
        {
            return "Not available";
        }

        var sign = input.Amount.Value > 0m ? "+" : "";
        return input.Kind switch
        {
            PlanDesignComparisonCostKinds.Pepm => $"{sign}{input.Amount.Value.ToString("C2", CultureInfo.CurrentCulture)} PEPM",
            PlanDesignComparisonCostKinds.Annual => $"{sign}{input.Amount.Value.ToString("C0", CultureInfo.CurrentCulture)} annual",
            PlanDesignComparisonCostKinds.Percentage => $"{sign}{input.Amount.Value.ToString("0.##", CultureInfo.InvariantCulture)}%",
            _ => "Not available"
        };
    }

    private static bool RequiresReview(string canonicalStatus, string currentAvailability, string selectedAvailability) =>
        canonicalStatus.Equals(CurrentVsQuotePlanDesignStatuses.NeedsReview, StringComparison.OrdinalIgnoreCase) ||
        currentAvailability is PlanDesignCanonicalAvailability.NeedsReview or PlanDesignCanonicalAvailability.Partial ||
        selectedAvailability is PlanDesignCanonicalAvailability.NeedsReview or PlanDesignCanonicalAvailability.Partial;

    private static bool TryClassifyNumericProvision(string provision, string current, string selected, out PlanDesignComparisonClassification result)
    {
        var normalizedProvision = provision.ToLowerInvariant();
        if (normalizedProvision.Contains("coinsurance") || normalizedProvision.Contains("reimbursement"))
        {
            return TryCompare(ParsePercent(current), ParsePercent(selected), higherIsBetter: true, "Higher reimbursement", out result);
        }

        if (normalizedProvision.Contains("deductible"))
        {
            return TryCompare(ParseCurrency(current), ParseCurrency(selected), higherIsBetter: false, "Lower deductible", out result);
        }

        if (normalizedProvision.Contains("waiting") || normalizedProvision.Contains("elimination"))
        {
            return TryCompare(ParseDurationDays(current), ParseDurationDays(selected), higherIsBetter: false, "Shorter waiting period", out result);
        }

        if (normalizedProvision.Contains("termination"))
        {
            return TryCompare(ParseAge(current), ParseAge(selected), higherIsBetter: true, "Later termination age", out result);
        }

        if (normalizedProvision.Contains("maximum") || normalizedProvision.Contains("benefit amount") || normalizedProvision.Contains("benefit maximum"))
        {
            var currency = TryCompare(ParseCurrency(current), ParseCurrency(selected), higherIsBetter: true, "Higher coverage limit", out result);
            return currency || TryCompare(ParseMultiplier(current), ParseMultiplier(selected), higherIsBetter: true, "Higher earnings multiple", out result);
        }

        result = default;
        return false;
    }

    private static bool TryCompare(decimal? current, decimal? selected, bool higherIsBetter, string rule, out PlanDesignComparisonClassification result)
    {
        if (!current.HasValue || !selected.HasValue || current.Value == selected.Value)
        {
            result = default;
            return false;
        }

        var improved = higherIsBetter ? selected.Value > current.Value : selected.Value < current.Value;
        result = improved
            ? new(PlanDesignComparisonStatuses.Improved, $"{rule} is supported by the normalized numeric values.", "Better")
            : new(PlanDesignComparisonStatuses.Reduced, $"{rule} is supported by the normalized numeric values.", "Review Required");
        return true;
    }

    private static decimal? ParsePercent(string value) => TryParseMatch(value, @"\b(100|[1-9]?\d(?:\.\d+)?)\s*(?:%|\bpercent\b)");
    private static decimal? ParseCurrency(string value) => TryParseMatch(value, @"\$\s*([\d,]+(?:\.\d{1,2})?)");
    private static decimal? ParseMultiplier(string value) => TryParseMatch(value, @"\b(\d+(?:\.\d+)?)\s*(?:x|times)\b");
    private static decimal? ParseDurationDays(string value) => TryParseMatch(value, @"\b(\d+(?:\.\d+)?)\s*days?\b");
    private static decimal? ParseAge(string value) => TryParseMatch(value, @"\bage\s*(\d+(?:\.\d+)?)\b");

    private static decimal? TryParseMatch(string value, string pattern)
    {
        var match = Regex.Match(value ?? "", pattern, RegexOptions.IgnoreCase);
        return match.Success && decimal.TryParse(match.Groups[1].Value.Replace(",", ""), NumberStyles.Number, CultureInfo.InvariantCulture, out var parsed)
            ? parsed
            : null;
    }

    private static bool IsMissing(string value) => string.IsNullOrWhiteSpace(value) ||
        value.Equals("Not found", StringComparison.OrdinalIgnoreCase) ||
        value.Equals("Not available", StringComparison.OrdinalIgnoreCase) ||
        value.Equals("N/A", StringComparison.OrdinalIgnoreCase);

    private static string Normalize(string value) => PlanDesignCanonicalMapper.NormalizeComparableText(value)
        .Replace("per cent", "%", StringComparison.Ordinal)
        .Replace("percent", "%", StringComparison.Ordinal)
        .Replace(" %", "%", StringComparison.Ordinal);
}

public readonly record struct PlanDesignComparisonClassification(string Status, string Rationale, string Recommendation);

public static class PlanDesignComparisonCostKinds
{
    public const string Pepm = "PEPM";
    public const string Annual = "Annual";
    public const string Percentage = "Percentage";
    public const string Preview = "Preview";
    public const string NotAvailable = "Not available";
}

public sealed class PlanDesignComparisonCostInput
{
    public string Kind { get; init; } = PlanDesignComparisonCostKinds.NotAvailable;
    public decimal? Amount { get; init; }
}

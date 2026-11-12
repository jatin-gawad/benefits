using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public static class UnderwritingRiskCalculator
{
    public static decimal? CalculateChange(
        IReadOnlyCollection<DemographicLine> demographics,
        IReadOnlyCollection<UnderwritingReferenceRate> referenceRates,
        string rateType)
    {
        var candidates = referenceRates
            .Where(rate => rate.RateType.Equals(rateType, StringComparison.OrdinalIgnoreCase))
            .ToList();
        if (demographics.Count == 0 || candidates.Count == 0)
        {
            return null;
        }

        var versionName = candidates[0].VersionName;
        var rates = candidates
            .Where(rate => rate.VersionName.Equals(versionName, StringComparison.OrdinalIgnoreCase))
            .ToList();
        var divisor = IsPerThousandRateType(rateType)
            ? 1_000m
            : 100m;

        var currentVolume = demographics.Sum(row => (row.CurrentFemaleVolume ?? 0m) + (row.CurrentMaleVolume ?? 0m));
        var previousVolume = demographics.Sum(row => (row.PriorFemaleVolume ?? 0m) + (row.PriorMaleVolume ?? 0m));
        if (currentVolume <= 0m || previousVolume <= 0m)
        {
            return null;
        }

        decimal currentWeighted = 0m;
        decimal previousWeighted = 0m;
        foreach (var rate in rates)
        {
            foreach (var row in demographics.Where(row => AgeBandMatches(row.AgeBand, rate.AgeBand, rateType)))
            {
                var current = rate.Gender.Equals(UnderwritingRateGenders.Female, StringComparison.OrdinalIgnoreCase)
                    ? row.CurrentFemaleVolume ?? 0m
                    : row.CurrentMaleVolume ?? 0m;
                var previous = rate.Gender.Equals(UnderwritingRateGenders.Female, StringComparison.OrdinalIgnoreCase)
                    ? row.PriorFemaleVolume ?? 0m
                    : row.PriorMaleVolume ?? 0m;
                currentWeighted += current * rate.Rate / divisor;
                previousWeighted += previous * rate.Rate / divisor;
            }
        }

        if (previousWeighted <= 0m)
        {
            return null;
        }

        return ((currentWeighted / currentVolume) / (previousWeighted / previousVolume)) - 1m;
    }

    public static string? RateTypeForBenefit(string? benefit)
    {
        var normalized = NormalizeBenefit(benefit);
        if ((normalized.Equals("LIFE", StringComparison.Ordinal) ||
             normalized.Contains("BASIC LIFE", StringComparison.Ordinal) ||
             normalized.Contains("EMPLOYEE LIFE", StringComparison.Ordinal)) &&
            !normalized.Contains("DEPENDENT", StringComparison.Ordinal) &&
            !normalized.Contains("OPTIONAL", StringComparison.Ordinal) &&
            !normalized.Contains("SPOUS", StringComparison.Ordinal) &&
            !normalized.Contains("CHILD", StringComparison.Ordinal))
        {
            return UnderwritingRateTypes.Mortality;
        }

        return normalized.Equals("LTD", StringComparison.Ordinal) ||
               normalized.Contains("LONG TERM DISABILITY", StringComparison.Ordinal)
            ? UnderwritingRateTypes.Morbidity
            : null;
    }

    public static decimal? CalculateRate(decimal? currentRate, decimal? changeInRisk)
    {
        return currentRate.HasValue && changeInRisk.HasValue
            ? decimal.Round(currentRate.Value * (1m + changeInRisk.Value), 3, MidpointRounding.AwayFromZero)
            : null;
    }

    public static decimal? CalculatePremium(string? rateType, decimal? volume, decimal? calculatedRate)
    {
        if (!volume.HasValue || !calculatedRate.HasValue || string.IsNullOrWhiteSpace(rateType))
        {
            return null;
        }

        var divisor = rateType.Equals(UnderwritingRateTypes.Mortality, StringComparison.OrdinalIgnoreCase)
            ? 1_000m
            : 100m;
        return decimal.Round(volume.Value / divisor * calculatedRate.Value, 2, MidpointRounding.AwayFromZero);
    }

    private static bool AgeBandMatches(string? demographicAgeBand, string referenceAgeBand, string rateType)
    {
        var demographic = IsLifeAgeBandRateType(rateType)
            ? DemographicAgeBandNormalizer.NormalizeLifeAgeBand(demographicAgeBand)
            : DemographicAgeBandNormalizer.NormalizeLtdAgeBand(demographicAgeBand);
        var reference = IsLifeAgeBandRateType(rateType)
            ? DemographicAgeBandNormalizer.NormalizeLifeAgeBand(referenceAgeBand)
            : DemographicAgeBandNormalizer.NormalizeLtdAgeBand(referenceAgeBand);

        if (reference.Equals("65-69", StringComparison.OrdinalIgnoreCase) &&
            demographic.Equals("65+", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        if (reference.Equals("60-64", StringComparison.OrdinalIgnoreCase) &&
            demographic.Equals("60 & over", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return demographic.Equals(reference, StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsPerThousandRateType(string rateType)
    {
        return rateType.Equals(UnderwritingRateTypes.Mortality, StringComparison.OrdinalIgnoreCase) ||
               rateType.Equals(UnderwritingRateTypes.SoaBasicLifeMortality, StringComparison.OrdinalIgnoreCase) ||
               rateType.Equals(UnderwritingRateTypes.SoaGroupLifeWaiver, StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsLifeAgeBandRateType(string rateType)
    {
        return rateType.Equals(UnderwritingRateTypes.Mortality, StringComparison.OrdinalIgnoreCase) ||
               rateType.Equals(UnderwritingRateTypes.SoaBasicLifeMortality, StringComparison.OrdinalIgnoreCase);
    }

    private static string NormalizeBenefit(string? benefit)
    {
        var cleaned = (benefit ?? "")
            .Replace("&", " AND ", StringComparison.Ordinal)
            .Replace("-", " ", StringComparison.Ordinal)
            .Replace("/", " ", StringComparison.Ordinal)
            .ToUpperInvariant();
        return string.Join(" ", cleaned.Split(' ', StringSplitOptions.RemoveEmptyEntries));
    }
}

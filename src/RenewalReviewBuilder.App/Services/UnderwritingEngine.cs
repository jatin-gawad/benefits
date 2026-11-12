using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

// Shared calculation layer for renewals, quotes, negotiation workbench results, and future-facing projections.
// Claims Experience explains what happened historically; this engine calculates what rates and premiums should be.
public sealed class UnderwritingEngine
{
    public UnderwritingCalculationResult Calculate(UnderwritingCalculationInput input)
    {
        var currentPremium = FirstValue(
            input.CurrentPremium,
            CalculatePremium(input.BenefitLine, input.Volume, input.Lives, input.CurrentRate, input.RateBasis));
        var proposedPremium = FirstValue(
            input.ProposedPremium,
            CalculatePremium(input.BenefitLine, input.Volume, input.Lives, input.ProposedRate, input.RateBasis));
        var negotiatedPremium = FirstValue(
            input.NegotiatedPremium,
            CalculatePremium(input.BenefitLine, input.Volume, input.Lives, input.NegotiatedRate, input.RateBasis));

        var trend = NormalizeRate(input.Trend) ?? 0m;
        var targetLossRatio = NormalizeRate(input.TargetLossRatio);
        var poolingRate = NormalizeRate(input.PoolingAdjustment) ?? 0m;
        var ibnrRate = NormalizeRate(input.IbnrAdjustment) ?? 0m;
        var retentionAdminFeeRate = NormalizeRate(input.RetentionAdminFee) ?? 0m;
        var manualAdjustment = input.ManualAdjustment ?? 0m;

        var trendedClaims = CalculateTrendedClaims(input.CurrentClaims, trend);
        var poolingAdjustment = Multiply(trendedClaims, poolingRate);
        var ibnrAdjustment = Multiply(trendedClaims, ibnrRate);
        var projectedClaims = SumNullable(trendedClaims, poolingAdjustment, ibnrAdjustment, manualAdjustment == 0m ? (decimal?)null : manualAdjustment);
        var baseRequiredPremium = targetLossRatio is > 0m && projectedClaims.HasValue
            ? decimal.Round(projectedClaims.Value / targetLossRatio.Value, 2)
            : (decimal?)null;
        var retentionAdminFeeAdjustment = Multiply(baseRequiredPremium, retentionAdminFeeRate);
        var requiredPremium = SumNullable(baseRequiredPremium, retentionAdminFeeAdjustment);
        var requiredRate = CalculateRate(input.BenefitLine, input.Volume, input.Lives, requiredPremium, input.RateBasis);
        var projectedLossRatio = CalculateLossRatioPercent(projectedClaims, FirstValue(proposedPremium, requiredPremium));

        return new UnderwritingCalculationResult
        {
            SourceModule = input.SourceModule,
            VersionName = input.VersionName,
            BenefitLine = input.BenefitLine,
            CurrentClaims = input.CurrentClaims,
            CurrentPremium = currentPremium,
            TrendedClaims = trendedClaims,
            PoolingAdjustment = poolingAdjustment,
            IbnrAdjustment = ibnrAdjustment,
            RetentionAdminFeeAdjustment = retentionAdminFeeAdjustment,
            ManualAdjustment = manualAdjustment == 0m ? null : manualAdjustment,
            ProjectedClaims = projectedClaims,
            RequiredPremium = requiredPremium,
            ProposedPremium = proposedPremium,
            NegotiatedPremium = negotiatedPremium,
            RequiredRate = requiredRate,
            TargetLossRatio = targetLossRatio,
            ProjectedLossRatio = projectedLossRatio,
            RenewalChangePercent = CalculatePercentChange(currentPremium, FirstValue(proposedPremium, requiredPremium)),
            NegotiatedChangePercent = CalculatePercentChange(currentPremium, negotiatedPremium),
            Savings = CalculateSavings(proposedPremium, negotiatedPremium),
            Trend = trend
        };
    }

    public static decimal? CalculatePremium(string benefit, decimal? volume, int? lives, decimal? rate)
    {
        return CalculatePremium(benefit, volume, lives, rate, "");
    }

    public static decimal? CalculatePremium(string benefit, decimal? volume, int? lives, decimal? rate, string? rateBasis)
    {
        if (!rate.HasValue)
        {
            return null;
        }

        if (IsPerEmployeeRateBasis(rateBasis))
        {
            return lives.HasValue
                ? decimal.Round(lives.Value * rate.Value, 2, MidpointRounding.AwayFromZero)
                : null;
        }

        var divisor = RateVolumeDivisor(benefit, rateBasis);
        if (divisor.HasValue)
        {
            return volume.HasValue
                ? decimal.Round(volume.Value * rate.Value / divisor.Value, 2, MidpointRounding.AwayFromZero)
                : null;
        }

        return lives.HasValue
            ? decimal.Round(lives.Value * rate.Value, 2, MidpointRounding.AwayFromZero)
            : null;
    }

    public static decimal? CalculateRate(string benefit, decimal? volume, int? lives, decimal? premium)
    {
        return CalculateRate(benefit, volume, lives, premium, "");
    }

    public static decimal? CalculateRate(string benefit, decimal? volume, int? lives, decimal? premium, string? rateBasis)
    {
        if (!premium.HasValue)
        {
            return null;
        }

        if (IsPerEmployeeRateBasis(rateBasis))
        {
            return lives is > 0 ? decimal.Round(premium.Value / lives.Value, 4) : null;
        }

        var divisor = RateVolumeDivisor(benefit, rateBasis);
        if (divisor.HasValue)
        {
            return volume is > 0m ? decimal.Round(premium.Value * divisor.Value / volume.Value, 4) : null;
        }

        return lives is > 0 ? decimal.Round(premium.Value / lives.Value, 4) : null;
    }

    public static bool IsPerEmployeeRateBasis(string? rateBasis)
    {
        var normalized = new string((rateBasis ?? "")
            .Where(char.IsLetterOrDigit)
            .Select(char.ToLowerInvariant)
            .ToArray());
        return normalized.Contains("peremployee", StringComparison.Ordinal) ||
               normalized is "employee" or "lives" or "life";
    }

    public static int? RateVolumeDivisor(string benefit)
    {
        return RateVolumeDivisor(benefit, null);
    }

    public static int? RateVolumeDivisor(string benefit, string? rateBasis)
    {
        var explicitDivisor = ExplicitRateVolumeDivisor(rateBasis);
        if (explicitDivisor.HasValue)
        {
            return explicitDivisor;
        }

        var normalized = (benefit ?? "").Trim().ToUpperInvariant();

        if (normalized.Contains("DEPENDENT LIFE", StringComparison.Ordinal))
        {
            return null;
        }

        if (normalized.Equals("LIFE", StringComparison.Ordinal) ||
            normalized.Contains("BASIC LIFE", StringComparison.Ordinal) ||
            normalized.Contains("EMPLOYEE LIFE", StringComparison.Ordinal) ||
            normalized.Contains("OPTIONAL LIFE", StringComparison.Ordinal) ||
            normalized.Contains("ACCIDENTAL DEATH", StringComparison.Ordinal) ||
            normalized.Contains("DISMEMBERMENT", StringComparison.Ordinal) ||
            normalized.Contains("AD&D", StringComparison.Ordinal) ||
            normalized.Contains("AD AND D", StringComparison.Ordinal) ||
            normalized.Equals("ASI", StringComparison.Ordinal) ||
            normalized.EndsWith(" ASI", StringComparison.Ordinal) ||
            normalized.Contains("CRITICAL ILLNESS", StringComparison.Ordinal))
        {
            return 1000;
        }

        if (normalized.Contains("LONG TERM DISABILITY", StringComparison.Ordinal) ||
            normalized.Contains("LONG-TERM DISABILITY", StringComparison.Ordinal) ||
            normalized.Equals("LTD", StringComparison.Ordinal))
        {
            return 100;
        }

        if (normalized.Contains("SHORT TERM DISABILITY", StringComparison.Ordinal) ||
            normalized.Contains("SHORT-TERM DISABILITY", StringComparison.Ordinal) ||
            normalized.Equals("STD", StringComparison.Ordinal))
        {
            return 10;
        }

        return null;
    }

    private static int? ExplicitRateVolumeDivisor(string? rateBasis)
    {
        if (string.IsNullOrWhiteSpace(rateBasis))
        {
            return null;
        }

        var normalized = rateBasis.Replace('_', ' ');
        var perIndex = normalized.IndexOf("per", StringComparison.OrdinalIgnoreCase);
        var slashIndex = normalized.IndexOf('/');
        var markerIndex = perIndex >= 0 && (slashIndex < 0 || perIndex < slashIndex)
            ? perIndex
            : slashIndex;
        if (markerIndex < 0)
        {
            return null;
        }

        var markerLength = markerIndex == perIndex ? 3 : 1;
        var numericText = new string(normalized[(markerIndex + markerLength)..]
            .TakeWhile(character => char.IsDigit(character) || char.IsWhiteSpace(character) || character is '$' or ',')
            .Where(char.IsDigit)
            .ToArray());

        return int.TryParse(numericText, out var divisor) && divisor > 0 ? divisor : null;
    }

    public static decimal? CalculateTrendedClaims(decimal? claims, decimal? trend)
    {
        var trendRate = NormalizeRate(trend) ?? 0m;
        return claims.HasValue ? decimal.Round(claims.Value * (1m + trendRate), 2) : null;
    }

    public static decimal? CalculateRequiredPremium(decimal? projectedClaims, decimal? targetLossRatio)
    {
        var target = NormalizeRate(targetLossRatio);
        return projectedClaims.HasValue && target is > 0m
            ? decimal.Round(projectedClaims.Value / target.Value, 2)
            : null;
    }

    public static decimal? CalculateLossRatioPercent(decimal? claims, decimal? premium)
    {
        return claims.HasValue && premium is > 0m
            ? decimal.Round(claims.Value / premium.Value * 100m, 1)
            : null;
    }

    public static decimal CalculateLossRatioFactor(decimal claims, decimal premium)
    {
        return premium == 0m ? 0m : decimal.Round(claims / premium, 6);
    }

    public static decimal CalculateTrendedLossRatioFactor(decimal lossRatio, decimal trend)
    {
        return decimal.Round(lossRatio * (1m + (NormalizeRate(trend) ?? 0m)), 6);
    }

    public static decimal CalculateRequiredChangeFactor(decimal trendedLossRatio, decimal targetLossRatio)
    {
        var target = NormalizeRate(targetLossRatio) ?? 0m;
        return target == 0m ? 0m : decimal.Round(trendedLossRatio / target - 1m, 6);
    }

    public static decimal? CalculatePercentChange(decimal? current, decimal? renewal)
    {
        return current is > 0m && renewal.HasValue
            ? decimal.Round((renewal.Value / current.Value - 1m) * 100m, 1)
            : null;
    }

    public static decimal? CalculateSavings(decimal? proposedPremium, decimal? negotiatedPremium)
    {
        return proposedPremium.HasValue && negotiatedPremium.HasValue
            ? decimal.Round(proposedPremium.Value - negotiatedPremium.Value, 2)
            : null;
    }

    public static decimal? NormalizeRate(decimal? value)
    {
        if (!value.HasValue)
        {
            return null;
        }

        return Math.Abs(value.Value) > 1m ? value.Value / 100m : value.Value;
    }

    private static decimal? Multiply(decimal? value, decimal rate)
    {
        return value.HasValue && rate != 0m ? decimal.Round(value.Value * rate, 2) : null;
    }

    private static decimal? SumNullable(params decimal?[] values)
    {
        return values.Any(value => value.HasValue)
            ? decimal.Round(values.Sum(value => value ?? 0m), 2)
            : null;
    }

    private static decimal? FirstValue(params decimal?[] values)
    {
        return values.FirstOrDefault(value => value.HasValue);
    }
}

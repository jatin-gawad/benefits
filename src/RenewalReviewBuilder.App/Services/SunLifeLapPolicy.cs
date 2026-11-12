using System.Globalization;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed record SunLifeLapAssessment(
    int LatestEhcLives,
    DateOnly LatestMonth,
    decimal ProjectedLapLevel,
    bool CrossedBelowFifty,
    int? HighestEarlierEhcLives)
{
    public string ParticipationBasisText
    {
        get
        {
            var basis = $"{LatestEhcLives.ToString("#,0", CultureInfo.InvariantCulture)} EHC lives ({LatestMonth.ToString("MMM yyyy", CultureInfo.InvariantCulture)})";
            return CrossedBelowFifty && HighestEarlierEhcLives.HasValue
                ? $"{basis}; crossed below 50 from {HighestEarlierEhcLives.Value.ToString("#,0", CultureInfo.InvariantCulture)}"
                : basis;
        }
    }
}

public static class SunLifeLapPolicy
{
    public const int HigherLapMinimumLives = 50;
    public const decimal LowerLapLevel = 10_000m;
    public const decimal HigherLapLevel = 15_000m;

    public static decimal ResolveLapLevel(int ehcLives)
    {
        if (ehcLives < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(ehcLives));
        }

        return ehcLives < HigherLapMinimumLives ? LowerLapLevel : HigherLapLevel;
    }

    public static SunLifeLapAssessment? Evaluate(
        string? carrier,
        IEnumerable<MonthlyClaimsRecord> monthlyClaims)
    {
        if (!IsSunLife(carrier))
        {
            return null;
        }

        var history = monthlyClaims
            .Where(row => IsEhc(row.BenefitLine) && row.Month.HasValue && row.Lives is > 0)
            .GroupBy(row => row.Month!.Value)
            .Select(group => new
            {
                Month = group.Key,
                Lives = group.Max(row => row.Lives!.Value)
            })
            .OrderBy(row => row.Month)
            .ToList();

        if (history.Count == 0)
        {
            return null;
        }

        var latest = history[^1];
        var highestEarlierLives = history
            .Take(history.Count - 1)
            .Select(row => (int?)row.Lives)
            .Max();
        var crossedBelowFifty = latest.Lives < HigherLapMinimumLives &&
                                highestEarlierLives is >= HigherLapMinimumLives;

        return new SunLifeLapAssessment(
            latest.Lives,
            latest.Month,
            ResolveLapLevel(latest.Lives),
            crossedBelowFifty,
            highestEarlierLives);
    }

    private static bool IsSunLife(string? carrier)
    {
        if (string.IsNullOrWhiteSpace(carrier))
        {
            return false;
        }

        var normalized = new string(carrier.Where(char.IsLetterOrDigit).ToArray());
        return normalized.StartsWith("SunLife", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsEhc(string? benefitLine)
    {
        if (string.IsNullOrWhiteSpace(benefitLine))
        {
            return false;
        }

        var normalized = new string(benefitLine.Where(char.IsLetterOrDigit).ToArray());
        return normalized.Equals("EHC", StringComparison.OrdinalIgnoreCase) ||
               normalized.Equals("Health", StringComparison.OrdinalIgnoreCase) ||
               normalized.StartsWith("ExtendedHealth", StringComparison.OrdinalIgnoreCase);
    }
}

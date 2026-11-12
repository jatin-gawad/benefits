using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed class ClaimsRangeCommentaryBuilder
{
    public const string Footnote = "Counts are by certificate and are not normalized for changes in eligible lives.";

    public string Build(ClaimsRangeComparison comparison)
    {
        ArgumentNullException.ThrowIfNull(comparison);
        if (comparison.PriorReport is null || comparison.CurrentReport is null || comparison.Rows.Count == 0)
        {
            return "Claims Range commentary is unavailable until both prior and current reports are complete.";
        }

        var health = Metrics(comparison, ClaimsRangeBenefitKeys.Health);
        var dental = Metrics(comparison, ClaimsRangeBenefitKeys.Dental);
        var totalSentence = BuildTotalSentence(health, dental);
        var detailSentence = BuildPriorityDetail(health, dental);
        return string.IsNullOrWhiteSpace(detailSentence)
            ? totalSentence
            : $"{totalSentence} {detailSentence}";
    }

    private static string BuildTotalSentence(BenefitMetrics health, BenefitMetrics dental)
    {
        if (health.CurrentTotal == dental.CurrentTotal && health.PriorTotal == dental.PriorTotal)
        {
            if (health.CurrentTotal == 0)
            {
                return health.PriorTotal == 0
                    ? "No claiming certificates were reported for EHC or dental in either period."
                    : $"No claiming certificates were reported for EHC or dental in the current period, compared with {health.PriorTotal} each in the prior period.";
            }

            return $"EHC and dental each show {CountPhrase(health.CurrentTotal)}, {TrendPhrase(health.CurrentTotal, health.PriorTotal)} in the prior period.";
        }

        return $"EHC shows {CountPhrase(health.CurrentTotal)} versus {health.PriorTotal} prior; dental shows {CountPhrase(dental.CurrentTotal)} versus {dental.PriorTotal} prior.";
    }

    private static string BuildPriorityDetail(BenefitMetrics health, BenefitMetrics dental)
    {
        var aboveTenThousand = health.CurrentOverTenThousand + dental.CurrentOverTenThousand;
        if (aboveTenThousand > 0)
        {
            return BuildThresholdSentence(health, dental, aboveTenThousand, overTenThousand: true);
        }

        var atLeastFiveThousand = health.CurrentAtLeastFiveThousand + dental.CurrentAtLeastFiveThousand;
        if (atLeastFiveThousand > 0)
        {
            return BuildThresholdSentence(health, dental, atLeastFiveThousand, overTenThousand: false);
        }

        var healthShift = ShiftClause("EHC's", health);
        var dentalShift = ShiftClause("dental's", dental);
        if (!string.IsNullOrWhiteSpace(healthShift) && !string.IsNullOrWhiteSpace(dentalShift))
        {
            return $"{healthShift}, while {LowercaseFirst(dentalShift)}.";
        }

        var singleShift = FirstNonBlank(healthShift, dentalShift);
        return string.IsNullOrWhiteSpace(singleShift) ? "" : $"{singleShift}.";
    }

    private static string BuildThresholdSentence(
        BenefitMetrics health,
        BenefitMetrics dental,
        int combined,
        bool overTenThousand)
    {
        var healthCount = overTenThousand ? health.CurrentOverTenThousand : health.CurrentAtLeastFiveThousand;
        var dentalCount = overTenThousand ? dental.CurrentOverTenThousand : dental.CurrentAtLeastFiveThousand;
        var threshold = overTenThousand ? "above $10,000" : "at $5,000 or above";
        if (healthCount > 0 && dentalCount > 0)
        {
            return $"Current counts include {healthCount} EHC and {dentalCount} dental {CertificateWord(combined)} {threshold}.";
        }

        var benefit = healthCount > 0 ? "EHC" : "Dental";
        var count = healthCount > 0 ? healthCount : dentalCount;
        return $"{benefit} includes {count} current {CertificateWord(count)} {threshold}.";
    }

    private static string ShiftClause(string subject, BenefitMetrics metrics)
    {
        if (metrics.PriorHighest is null || metrics.CurrentHighest is null ||
            metrics.PriorHighest.RangeCode.Equals(metrics.CurrentHighest.RangeCode, StringComparison.OrdinalIgnoreCase))
        {
            return "";
        }

        return $"{subject} highest active band moved from {metrics.PriorHighest.RangeLabel} to {metrics.CurrentHighest.RangeLabel}";
    }

    private static BenefitMetrics Metrics(ClaimsRangeComparison comparison, string benefit)
    {
        var rows = comparison.Rows
            .Where(row => row.Benefit.Equals(benefit, StringComparison.OrdinalIgnoreCase))
            .ToList();
        return new BenefitMetrics(
            rows.Sum(row => row.PriorCount),
            rows.Sum(row => row.CurrentCount),
            rows.Where(row => row.PriorCount > 0).OrderByDescending(row => row.SortOrder).FirstOrDefault(),
            rows.Where(row => row.CurrentCount > 0).OrderByDescending(row => row.SortOrder).FirstOrDefault(),
            rows.Where(row => row.LowerBoundInclusive >= 5000).Sum(row => row.CurrentCount),
            rows.Where(row =>
                    row.RangeCode.Equals(ClaimsRangeCodes.Over10000, StringComparison.OrdinalIgnoreCase) ||
                    (!row.UpperBoundInclusive.HasValue && row.LowerBoundInclusive >= 10000))
                .Sum(row => row.CurrentCount));
    }

    private static string CountPhrase(int count) => $"{count} {CertificateWord(count)}";

    private static string CertificateWord(int count) => count == 1
        ? "claiming certificate"
        : "claiming certificates";

    private static string TrendPhrase(int current, int prior)
    {
        return current.CompareTo(prior) switch
        {
            < 0 => $"down from {prior}",
            > 0 => $"up from {prior}",
            _ => $"unchanged from {prior}"
        };
    }

    private static string LowercaseFirst(string value)
    {
        return string.IsNullOrEmpty(value)
            ? value
            : char.ToLowerInvariant(value[0]) + value[1..];
    }

    private static string FirstNonBlank(params string[] values) =>
        values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)) ?? "";

    private sealed record BenefitMetrics(
        int PriorTotal,
        int CurrentTotal,
        ClaimsRangeComparisonRow? PriorHighest,
        ClaimsRangeComparisonRow? CurrentHighest,
        int CurrentAtLeastFiveThousand,
        int CurrentOverTenThousand);
}

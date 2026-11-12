using System.Globalization;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed class AssurePlusExperienceSummaryBuilder
{
    public AssurePlusExperienceSummary? Build(
        IReadOnlyCollection<AssurePlusPricingReport> reports,
        DateOnly? requestedEffectiveDate = null)
    {
        var usableReports = reports
            .Where(report => report.EffectiveDate.HasValue && report.Benefits.Count > 0)
            .OrderBy(report => report.EffectiveDate)
            .ToList();
        if (usableReports.Count == 0)
        {
            return null;
        }

        var currentReport = SelectCurrentReport(usableReports, requestedEffectiveDate);
        var currentYear = (requestedEffectiveDate ?? currentReport.EffectiveDate)!.Value.Year;
        var firstYear = currentYear - 2;
        var summary = new AssurePlusExperienceSummary
        {
            EffectiveDate = currentReport.EffectiveDate,
            HealthOutcome = CloneBenefit(currentReport.Benefits.FirstOrDefault(IsHealth)),
            DentalOutcome = CloneBenefit(currentReport.Benefits.FirstOrDefault(IsDental))
        };

        foreach (var warning in usableReports.SelectMany(report => report.Warnings).Distinct(StringComparer.OrdinalIgnoreCase))
        {
            summary.Warnings.Add(warning);
        }

        var activeBenefits = new List<string>();
        if (summary.HealthOutcome is not null)
        {
            activeBenefits.Add("Health");
        }

        if (summary.DentalOutcome is not null)
        {
            activeBenefits.Add("Dental");
        }

        foreach (var benefit in activeBenefits)
        {
            for (var year = firstYear; year <= currentYear; year++)
            {
                var observations = GetObservations(usableReports, benefit, year).ToList();
                if (observations.Count == 0)
                {
                    summary.Warnings.Add($"AssurePLUS {benefit} experience status is missing for {year}.");
                    continue;
                }

                var distinctRatings = observations
                    .Select(observation => NormalizeRating(observation.Rating))
                    .Where(value => value.Length > 0)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();
                var selected = observations
                    .OrderByDescending(observation => observation.IsRecentYear)
                    .ThenByDescending(observation => observation.ReportDate)
                    .First();
                var conflict = distinctRatings.Count > 1;
                if (conflict)
                {
                    summary.Warnings.Add(
                        $"AssurePLUS {benefit} experience status conflicts for {year}: {string.Join(" versus ", distinctRatings)}. " +
                        $"The recent-year value '{selected.Rating}' was used.");
                }

                summary.AnnualRatings.Add(new AssurePlusAnnualRating
                {
                    Benefit = benefit,
                    Year = year,
                    Rating = selected.Rating,
                    SourceReference = string.Join("; ", observations
                        .Select(observation => observation.SourceReference)
                        .Where(value => value.Length > 0)
                        .Distinct(StringComparer.OrdinalIgnoreCase)),
                    Conflict = conflict
                });
            }
        }

        summary.HealthCommentary = BuildCommentary(summary, "Health", firstYear, currentYear);
        summary.DentalCommentary = BuildCommentary(summary, "Dental", firstYear, currentYear);
        summary.Takeaway = BuildTakeaway(summary.DentalOutcome ?? summary.HealthOutcome);
        return summary;
    }

    private static AssurePlusPricingReport SelectCurrentReport(
        IReadOnlyList<AssurePlusPricingReport> reports,
        DateOnly? requestedEffectiveDate)
    {
        if (requestedEffectiveDate.HasValue)
        {
            var exact = reports.LastOrDefault(report => report.EffectiveDate == requestedEffectiveDate);
            if (exact is not null)
            {
                return exact;
            }

            var sameYear = reports.LastOrDefault(report => report.EffectiveDate!.Value.Year == requestedEffectiveDate.Value.Year);
            if (sameYear is not null)
            {
                return sameYear;
            }
        }

        return reports[^1];
    }

    private static IEnumerable<RatingObservation> GetObservations(
        IEnumerable<AssurePlusPricingReport> reports,
        string benefit,
        int year)
    {
        foreach (var report in reports)
        {
            var effectiveYear = report.EffectiveDate!.Value.Year;
            var row = report.Benefits.FirstOrDefault(item =>
                item.Benefit.Equals(benefit, StringComparison.OrdinalIgnoreCase));
            if (row is null)
            {
                continue;
            }

            if (effectiveYear - 1 == year && !string.IsNullOrWhiteSpace(row.PreviousYearRating))
            {
                yield return new RatingObservation(
                    row.PreviousYearRating,
                    false,
                    report.EffectiveDate.Value,
                    row.SourceReference);
            }

            if (effectiveYear == year && !string.IsNullOrWhiteSpace(row.RecentYearRating))
            {
                yield return new RatingObservation(
                    row.RecentYearRating,
                    true,
                    report.EffectiveDate.Value,
                    row.SourceReference);
            }
        }
    }

    private static string BuildCommentary(
        AssurePlusExperienceSummary summary,
        string benefit,
        int firstYear,
        int currentYear)
    {
        var rows = summary.AnnualRatings
            .Where(row => row.Benefit.Equals(benefit, StringComparison.OrdinalIgnoreCase))
            .OrderBy(row => row.Year)
            .ToList();
        if (rows.Count == 0)
        {
            return $"{benefit} experience history was not available.";
        }

        if (rows.Count == currentYear - firstYear + 1 &&
            rows.Select(row => NormalizeRating(row.Rating)).Distinct(StringComparer.OrdinalIgnoreCase).Count() == 1)
        {
            return $"{benefit} remained {rows[0].Rating} from {firstYear} to {currentYear}.";
        }

        var changes = rows
            .Zip(rows.Skip(1), (previous, current) => (previous, current))
            .Where(pair => !NormalizeRating(pair.previous.Rating).Equals(
                NormalizeRating(pair.current.Rating),
                StringComparison.OrdinalIgnoreCase))
            .ToList();
        if (changes.Count == 1)
        {
            var change = changes[0];
            var remained = rows.Last().Year > change.current.Year &&
                           rows.SkipWhile(row => row.Year < change.current.Year)
                               .All(row => NormalizeRating(row.Rating).Equals(
                                   NormalizeRating(change.current.Rating),
                                   StringComparison.OrdinalIgnoreCase));
            return remained
                ? $"{benefit} moved from {change.previous.Rating} to {change.current.Rating} in {change.current.Year} and remained there in {currentYear}."
                : $"{benefit} moved from {change.previous.Rating} to {change.current.Rating} in {change.current.Year}.";
        }

        var first = rows.First();
        var last = rows.Last();
        return $"{benefit} moved from {first.Rating} in {first.Year} to {last.Rating} in {last.Year}.";
    }

    private static string BuildTakeaway(AssurePlusBenefitPricing? outcome)
    {
        if (outcome is null)
        {
            return "";
        }

        var benefitName = outcome.Benefit.Equals("Dental", StringComparison.OrdinalIgnoreCase)
            ? "Dental"
            : "Health";
        if (NormalizeRating(outcome.RecentYearRating).Equals("Top Performing", StringComparison.OrdinalIgnoreCase) &&
            outcome.StandardChangePercent is > 0m &&
            outcome.TotalChangePercent is < 0m)
        {
            return $"Top Performing {benefitName.ToLowerInvariant()} experience produced a rate decrease despite the {outcome.StandardChangePercent.Value.ToString("0.00", CultureInfo.InvariantCulture)}% standard increase.";
        }

        if (outcome.ExperienceAdjustmentPercent is not null and not 0m)
        {
            return $"{benefitName} experience adjusted the standard rate change by {FormatSignedPercent(outcome.ExperienceAdjustmentPercent.Value)}.";
        }

        return $"{benefitName} pricing reflects the standard change with no experience adjustment.";
    }

    private static string FormatSignedPercent(decimal value) =>
        $"{(value > 0m ? "+" : value < 0m ? "−" : "")}{Math.Abs(value).ToString("0.00", CultureInfo.InvariantCulture)}%";

    private static AssurePlusBenefitPricing? CloneBenefit(AssurePlusBenefitPricing? source) => source is null
        ? null
        : new AssurePlusBenefitPricing
        {
            Benefit = source.Benefit,
            PreviousYearRating = source.PreviousYearRating,
            RecentYearRating = source.RecentYearRating,
            StandardChangePercent = source.StandardChangePercent,
            ExperienceAdjustmentPercent = source.ExperienceAdjustmentPercent,
            TotalChangePercent = source.TotalChangePercent,
            SourceReference = source.SourceReference
        };

    private static string NormalizeRating(string value) => string.Join(' ',
        value.Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries));

    private static bool IsHealth(AssurePlusBenefitPricing item) =>
        item.Benefit.Equals("Health", StringComparison.OrdinalIgnoreCase);

    private static bool IsDental(AssurePlusBenefitPricing item) =>
        item.Benefit.Equals("Dental", StringComparison.OrdinalIgnoreCase);

    private sealed record RatingObservation(
        string Rating,
        bool IsRecentYear,
        DateOnly ReportDate,
        string SourceReference);
}

using RenewalReviewBuilder.App.Models;
using System.Globalization;

namespace RenewalReviewBuilder.App.Services.CanonicalClaims;

public sealed class CanonicalClaimsParityService
{
    /// <summary>Checks only additive measures. Unique claimants are deliberately excluded from rollups.</summary>
    public List<CanonicalClaimsParityResult> ReconcileUtilization(CanonicalClaimsImportEnvelope envelope, decimal tolerance = 0.01m)
    {
        var results = new List<CanonicalClaimsParityResult>();
        foreach (var levelPair in new[] { (CanonicalClaimsSourceAggregationLevel.Class, CanonicalClaimsSourceAggregationLevel.Contract, "DetailToContract"), (CanonicalClaimsSourceAggregationLevel.Contract, CanonicalClaimsSourceAggregationLevel.Sponsor, "ContractToSponsor") })
        foreach (var measure in new[] { "Occurrences", "Submitted", "Eligible", "Paid" })
        {
            var lower = Sum(envelope.UtilizationFacts.Where(f => f.SourceAggregationLevel == levelPair.Item1 && f.RelationshipType == CanonicalClaimsRelationshipType.Total), measure);
            var upper = Sum(envelope.UtilizationFacts.Where(f => f.SourceAggregationLevel == levelPair.Item2 && f.RelationshipType == CanonicalClaimsRelationshipType.Total), measure);
            if (lower.HasValue && upper.HasValue) Compare(results, $"Utilization{levelPair.Item3}{measure}", "All", lower, upper, null, tolerance);
        }
        foreach (var measure in new[] { "Occurrences", "Submitted", "Eligible", "Paid" })
        {
            var parts = Sum(envelope.UtilizationFacts.Where(f => f.SourceAggregationLevel == CanonicalClaimsSourceAggregationLevel.Sponsor && f.RelationshipType is CanonicalClaimsRelationshipType.Employee or CanonicalClaimsRelationshipType.Dependant or CanonicalClaimsRelationshipType.Partner), measure);
            var total = Sum(envelope.UtilizationFacts.Where(f => f.SourceAggregationLevel == CanonicalClaimsSourceAggregationLevel.Sponsor && f.RelationshipType == CanonicalClaimsRelationshipType.Total), measure);
            if (parts.HasValue && total.HasValue) Compare(results, $"UtilizationRelationshipTotal{measure}", "All", parts, total, null, tolerance);
        }
        return results;
    }

    private static decimal? Sum(IEnumerable<CanonicalClaimsUtilizationFact> facts, string measure)
    {
        var values = facts.Select(f => measure switch { "Occurrences" => (decimal?)f.Occurrences, "Submitted" => f.SubmittedAmount, "Eligible" => f.EligibleAmount, _ => f.PaidAmount }).Where(v => v.HasValue).ToList();
        return values.Count == 0 ? null : values.Sum(v => v!.Value);
    }
    public List<CanonicalClaimsParityResult> Compare(ClaimsExperienceDataModelResult legacy, CanonicalClaimsImportEnvelope canonical, decimal tolerance = 0.01m)
    {
        var results = new List<CanonicalClaimsParityResult>();
        foreach (var legacyFact in legacy.Summaries)
        {
            var canonicalFact = canonical.FinancialFacts.FirstOrDefault(fact => fact.Benefit.Equals(NormalizeBenefit(legacyFact.BenefitLine), StringComparison.OrdinalIgnoreCase));
            Compare(results, "Premium", legacyFact.BenefitLine, legacyFact.Premium, canonicalFact?.Premium, canonicalFact?.FactId, tolerance);
            Compare(results, "PaidClaims", legacyFact.BenefitLine, legacyFact.PaidClaims, canonicalFact?.PaidClaims, canonicalFact?.FactId, tolerance);
            Compare(results, "IncurredClaims", legacyFact.BenefitLine, legacyFact.IncurredClaims, canonicalFact?.IncurredClaims, canonicalFact?.FactId, tolerance);
            Compare(results, "Lives", legacyFact.BenefitLine, legacyFact.Lives, canonicalFact?.Lives, canonicalFact?.FactId, 0m);
            Compare(results, "LossRatioInput", legacyFact.BenefitLine, legacyFact.LossRatio, canonicalFact?.LossRatio, canonicalFact?.FactId, tolerance);
        }

        foreach (var legacyMonth in legacy.MonthlyClaims)
        {
            var canonicalMonth = canonical.MonthlyFacts.FirstOrDefault(fact => fact.Month == legacyMonth.Month && fact.Benefit.Equals(NormalizeBenefit(legacyMonth.BenefitLine), StringComparison.OrdinalIgnoreCase));
            Compare(results, "MonthlyPaidClaims", legacyMonth.BenefitLine, legacyMonth.PaidClaims, canonicalMonth?.PaidClaims, canonicalMonth?.FactId, tolerance, legacyMonth.Month?.ToString("yyyy-MM") ?? "");
        }
        return results;
    }

    public List<CanonicalClaimsParityResult> Compare(CarrierPdfParseResult legacy, CanonicalClaimsImportEnvelope canonical, decimal tolerance = 0.01m)
    {
        ArgumentNullException.ThrowIfNull(legacy);
        var results = new List<CanonicalClaimsParityResult>();

        Compare(results, "ReportingPeriodStart", "All", PackDate(ParsedPeriodStart(legacy)), PackDate(canonical.Report.Period.Start), null, 0m);
        Compare(results, "ReportingPeriodEnd", "All", PackDate(ParsedPeriodEnd(legacy)), PackDate(canonical.Report.Period.End), null, 0m);

        foreach (var row in legacy.Pdf.ExperienceLines)
        {
            var fact = canonical.FinancialFacts.FirstOrDefault(candidate =>
                candidate.Evidence.RawCarrierLabel.Equals(row.Benefit, StringComparison.OrdinalIgnoreCase)
                && candidate.DimensionValue.Equals(row.Period, StringComparison.OrdinalIgnoreCase));
            Compare(results, "Premium", row.Benefit, row.BilledPremium, fact?.Premium, fact?.FactId, tolerance, row.Period);
            Compare(results, "PaidClaims", row.Benefit, row.PaidClaims, fact?.PaidClaims, fact?.FactId, tolerance, row.Period);
            Compare(results, "IncurredClaims", row.Benefit, row.IncurredClaims, fact?.IncurredClaims, fact?.FactId, tolerance, row.Period);
            Compare(results, "Lives", row.Benefit, row.AverageLives, fact?.Lives, fact?.FactId, 0m, row.Period);
            Compare(results, "LossRatioInput", row.Benefit, row.IncurredLossRatio, fact?.LossRatio, fact?.FactId, tolerance, row.Period);
        }

        foreach (var row in legacy.WorkbookFromPdf?.ClaimsDistributionRows ?? [])
        {
            var rawLabel = $"{row.Benefit}|{row.ServiceCategory}";
            var fact = canonical.DistributionFacts.FirstOrDefault(candidate =>
                candidate.Evidence.RawCarrierLabel.Equals(rawLabel, StringComparison.OrdinalIgnoreCase)
                && candidate.DimensionValue.Equals(row.ReportPeriod, StringComparison.OrdinalIgnoreCase)
                && candidate.Division.Equals(NormalizeDimension(row.Division), StringComparison.OrdinalIgnoreCase)
                && candidate.Class.Equals(NormalizeDimension(row.Class), StringComparison.OrdinalIgnoreCase));
            Compare(results, "DistributionAmount", row.Benefit, row.AmountPaid, fact?.MonetaryValue, fact?.FactId, tolerance, row.ReportPeriod);
            Compare(results, "DistributionPercent", row.Benefit, row.PercentPaid, fact?.PercentageValue, fact?.FactId, tolerance, row.ReportPeriod);
            Compare(results, "DistributionCount", row.Benefit, row.Occurrences, fact?.CountValue, fact?.FactId, 0m, row.ReportPeriod);
            Compare(results, "DistributionAverage", row.Benefit, row.AverageClaim, fact?.AverageAmount, fact?.FactId, tolerance, row.ReportPeriod);
        }

        foreach (var row in legacy.WorkbookFromPdf?.ClaimTrends ?? [])
        {
            var rawLabel = $"{row.Benefit}|{row.Category}";
            var fact = canonical.CategoryFacts.FirstOrDefault(candidate =>
                candidate.Evidence.RawCarrierLabel.Equals(rawLabel, StringComparison.OrdinalIgnoreCase)
                && candidate.DimensionValue.Equals(row.Period, StringComparison.OrdinalIgnoreCase));
            var benchmark = canonical.BenchmarkFacts.FirstOrDefault(candidate =>
                candidate.MeasurementType == "CarrierBenchmarkPercent"
                && candidate.Evidence.RawCarrierLabel.Equals(rawLabel, StringComparison.OrdinalIgnoreCase)
                && candidate.DimensionValue.Equals(row.Period, StringComparison.OrdinalIgnoreCase));
            Compare(results, "CategoryAmount", row.Benefit, row.Amount, fact?.MonetaryValue, fact?.FactId, tolerance, row.Period);
            Compare(results, "CategoryPercent", row.Benefit, row.PercentOfTotal, fact?.PercentageValue, fact?.FactId, tolerance, row.Period);
            Compare(results, "BenchmarkPercent", row.Benefit, row.BenchmarkPercent, benchmark?.PercentageValue, benchmark?.FactId, tolerance, row.Period);
        }

        var legacyMeaningful = legacy.Pdf.ExperienceLines.Count + (legacy.WorkbookFromPdf?.ClaimTrends.Count ?? 0) + (legacy.WorkbookFromPdf?.ClaimsDistributionRows.Count ?? 0);
        var canonicalMeaningful = canonical.FinancialFacts.Count + canonical.CategoryFacts.Count(fact => fact.MeasurementType == "ClaimsCategory") + canonical.DistributionFacts.Count;
        Compare(results, "MeaningfulDetailRowCount", "All", legacyMeaningful, canonicalMeaningful, null, 0m);
        return results;
    }

    private static void Compare(List<CanonicalClaimsParityResult> results, string metric, string benefit, decimal? legacy, decimal? canonical, string? factId, decimal tolerance, string period = "")
    {
        decimal? variance = legacy.HasValue && canonical.HasValue ? canonical.Value - legacy.Value : null;
        var classification = !legacy.HasValue ? CanonicalClaimsParityClassification.MissingLegacyField :
            !canonical.HasValue ? CanonicalClaimsParityClassification.MissingCanonicalField :
            variance.GetValueOrDefault() == 0m ? CanonicalClaimsParityClassification.ExactMatch :
            Math.Abs(variance.GetValueOrDefault()) <= tolerance ? CanonicalClaimsParityClassification.WithinRoundingTolerance :
            CanonicalClaimsParityClassification.UnexpectedMismatch;
        results.Add(new CanonicalClaimsParityResult
        {
            Metric = metric,
            Benefit = NormalizeBenefit(benefit),
            Period = period,
            LegacyValue = legacy,
            CanonicalValue = canonical,
            Variance = variance,
            Tolerance = tolerance,
            Classification = classification,
            Explanation = classification switch
            {
                CanonicalClaimsParityClassification.ExactMatch => "Legacy and canonical values match.",
                CanonicalClaimsParityClassification.WithinRoundingTolerance => "Values differ only within the approved rounding tolerance.",
                CanonicalClaimsParityClassification.MissingLegacyField => "The legacy result did not report this value.",
                CanonicalClaimsParityClassification.MissingCanonicalField => "The required canonical value is missing.",
                _ => "Values differ beyond the approved tolerance."
            },
            RelatedCanonicalFactIds = factId ?? "",
            ReviewStatus = classification is CanonicalClaimsParityClassification.ExactMatch or CanonicalClaimsParityClassification.WithinRoundingTolerance ? "Not Required" : "Needs Review"
        });
    }

    private static string NormalizeBenefit(string value)
    {
        return value.Trim() switch
        {
            var item when item.Equals("Extended Health Care", StringComparison.OrdinalIgnoreCase) => "EHC",
            var item when item.Equals("Health", StringComparison.OrdinalIgnoreCase) => "EHC",
            var item when item.Equals("Dental Care", StringComparison.OrdinalIgnoreCase) => "Dental",
            var item => item
        };
    }

    private static decimal? PackDate(DateOnly? value) => value.HasValue ? value.Value.Year * 10000m + value.Value.Month * 100m + value.Value.Day : null;

    private static DateOnly? ParsedPeriodStart(CarrierPdfParseResult parsed)
    {
        var values = new List<DateOnly>();
        foreach (var row in parsed.Pdf.ExperienceLines)
            if (TryParsePeriodLabel(row.Period, out var start, out _)) values.Add(start);
        foreach (var row in parsed.WorkbookFromPdf?.ClaimsDistributionRows ?? [])
            if (DateOnly.TryParse(row.PeriodStart, out var start)) values.Add(start);
        return values.Count == 0 ? null : values.Min();
    }

    private static DateOnly? ParsedPeriodEnd(CarrierPdfParseResult parsed)
    {
        var values = new List<DateOnly>();
        foreach (var row in parsed.Pdf.ExperienceLines)
            if (TryParsePeriodLabel(row.Period, out _, out var end)) values.Add(end);
        foreach (var row in parsed.WorkbookFromPdf?.ClaimsDistributionRows ?? [])
            if (DateOnly.TryParse(row.PeriodEnd, out var end)) values.Add(end);
        return values.Count == 0 ? null : values.Max();
    }

    private static string NormalizeDimension(string value) => value.Equals("ALL", StringComparison.OrdinalIgnoreCase) ? "All" : value.Trim();

    private static bool TryParsePeriodLabel(string label, out DateOnly start, out DateOnly end)
    {
        start = default;
        end = default;
        var parts = label.Split(" - ", 2, StringSplitOptions.TrimEntries);
        if (parts.Length != 2) return false;
        return DateOnly.TryParse(parts[0], CultureInfo.GetCultureInfo("en-CA"), DateTimeStyles.AllowWhiteSpaces, out start)
            && DateOnly.TryParse(parts[1], CultureInfo.GetCultureInfo("en-CA"), DateTimeStyles.AllowWhiteSpaces, out end);
    }
}

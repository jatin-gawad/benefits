using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

/// <summary>Applies verified enrollment counts to parsed Allstate per-employee CI rates.</summary>
public sealed class AllstateCriticalIllnessRateLineService
{
    public void Apply(AnalysisResult analysis, AllstateCriticalIllnessEnrollment? enrollment)
    {
        if (enrollment?.HasEntries != true)
        {
            return;
        }

        var lines = analysis.Pdf.RateLines
            .Where(line =>
                line.SourceCarrier.Equals(AllstateCriticalIllnessPdfParser.CarrierName, StringComparison.OrdinalIgnoreCase) &&
                line.Benefit.Equals("Critical Illness", StringComparison.OrdinalIgnoreCase) &&
                line.RateBasis.Equals(AllstateCriticalIllnessPdfParser.PerEmployeeRateBasis, StringComparison.OrdinalIgnoreCase))
            .ToList();
        if (lines.Count == 0)
        {
            analysis.SourceChecks.Add(new SourceCheck
            {
                Area = "Allstate CI enrollment",
                Status = "Needs review",
                Detail = "Enrollment counts were entered, but no matching Allstate Critical Illness application was parsed from the selected PDFs."
            });
            return;
        }

        decimal currentAdded = 0m;
        decimal renewalAdded = 0m;
        foreach (var line in lines)
        {
            var lives = line.TierOrClass.Contains("Family", StringComparison.OrdinalIgnoreCase)
                ? enrollment.FamilyCount
                : enrollment.EmployeeOnlyCount;
            line.Lives = lives;
            line.CurrentPremium = UnderwritingEngine.CalculatePremium(
                line.Benefit, line.Volume, lives, line.CurrentRate, line.RateBasis);
            line.RenewalPremium = UnderwritingEngine.CalculatePremium(
                line.Benefit, line.Volume, lives, line.RenewalRate, line.RateBasis);
            line.CarrierProposedPremium = line.RenewalPremium;
            line.NegotiatedPremium = line.RenewalPremium;
            line.PercentChange = UnderwritingEngine.CalculatePercentChange(line.CurrentPremium, line.RenewalPremium);
            line.IncludeInPremiumTotal = true;
            line.SourceRowRole = "Detail";
            line.ValidationStatus = RateLineValidationStatuses.Supported;
            line.ValidationMessage = $"Calculated as {lives} enrolled employee(s) x the Allstate monthly {line.TierOrClass} rate. Rate is per employee, not per $1,000 of benefit. The same sold plan-year rate is carried across the consolidated review; no separate Allstate renewal increase was extracted.";
            currentAdded += line.CurrentPremium ?? 0m;
            renewalAdded += line.RenewalPremium ?? 0m;
        }

        analysis.Pdf.CurrentMonthlyPremium = AddToCarrierTotal(analysis.Pdf.CurrentMonthlyPremium, currentAdded);
        analysis.Pdf.RenewalMonthlyPremium = AddToCarrierTotal(analysis.Pdf.RenewalMonthlyPremium, renewalAdded);
        if (analysis.Pdf.CurrentMonthlyPremium.HasValue && analysis.Pdf.RenewalMonthlyPremium.HasValue)
        {
            analysis.Pdf.MonthlyPremiumChange = analysis.Pdf.RenewalMonthlyPremium.Value - analysis.Pdf.CurrentMonthlyPremium.Value;
            analysis.Pdf.MonthlyPremiumChangePercent = UnderwritingEngine.CalculatePercentChange(
                analysis.Pdf.CurrentMonthlyPremium,
                analysis.Pdf.RenewalMonthlyPremium);
        }

        analysis.SourceChecks.Add(new SourceCheck
        {
            Area = "Allstate CI enrollment",
            Status = "Calculated",
            Detail = $"EE {enrollment.EmployeeOnlyCount}, Family {enrollment.FamilyCount}; monthly current {ParsingHelpers.Money(currentAdded)}, renewal {ParsingHelpers.Money(renewalAdded)}. Per-employee tier rates were used."
        });
    }

    private static decimal? AddToCarrierTotal(decimal? carrierTotal, decimal standalonePremium) =>
        carrierTotal.HasValue ? carrierTotal.Value + standalonePremium : standalonePremium;
}

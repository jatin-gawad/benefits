using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

/// <summary>Turns manual HCSA funding assumptions into ordinary monthly rate lines.</summary>
public sealed class HcsaRateLineService
{
    public const string ManualSourceReference = "Manual HCSA calculator";

    public void Apply(AnalysisResult analysis, HcsaCalculation? calculation)
    {
        if (calculation?.HasEntries != true)
        {
            return;
        }

        // A manual calculator entry is intentional: it replaces any carrier HCSA
        // premium line so the account is never counted twice.
        analysis.Pdf.RateLines.RemoveAll(line => IsHcsa(line.Benefit));

        var includedClasses = calculation.IncludedClasses().ToList();
        foreach (var item in includedClasses)
        {
            var currentMonthlyRate = item.EligibleEmployees > 0
                ? decimal.Round(item.CurrentMonthlyCost / item.EligibleEmployees, 6, MidpointRounding.AwayFromZero)
                : 0m;
            var renewalMonthlyRate = item.EligibleEmployees > 0
                ? decimal.Round(item.RenewalMonthlyCost / item.EligibleEmployees, 6, MidpointRounding.AwayFromZero)
                : 0m;

            analysis.Pdf.RateLines.Add(new RateLine
            {
                Benefit = "HCSA",
                TierOrClass = string.IsNullOrWhiteSpace(item.ClassName) ? "All eligible employees" : item.ClassName.Trim(),
                Lives = item.EligibleEmployees,
                CurrentRate = currentMonthlyRate,
                CurrentPremium = item.CurrentMonthlyCost,
                RenewalRate = renewalMonthlyRate,
                RenewalPremium = item.RenewalMonthlyCost,
                CarrierProposedRate = renewalMonthlyRate,
                CarrierProposedPremium = item.RenewalMonthlyCost,
                NegotiatedRate = renewalMonthlyRate,
                NegotiatedPremium = item.RenewalMonthlyCost,
                PercentChange = UnderwritingEngine.CalculatePercentChange(item.CurrentMonthlyCost, item.RenewalMonthlyCost),
                FundingArrangement = "Employer-funded HCSA",
                SourceReference = $"{ManualSourceReference}: {item.EligibleEmployees} eligible employee(s), {item.UsagePercent:0.##}% usage, {item.AdministrationFeePercent:0.##}% administration fee."
            });
        }

        analysis.SourceChecks.Add(new SourceCheck
        {
            Area = "HCSA funding",
            Status = "Entered",
            Detail = $"Manual HCSA calculator supplied {includedClasses.Count} class line(s): current annual {ParsingHelpers.Money(calculation.CurrentAnnualCost)}, renewal annual {ParsingHelpers.Money(calculation.RenewalAnnualCost)}."
        });
    }

    private static bool IsHcsa(string benefit) =>
        benefit.Equals("HCSA", StringComparison.OrdinalIgnoreCase) ||
        benefit.Equals("HSA", StringComparison.OrdinalIgnoreCase) ||
        benefit.Contains("health spending", StringComparison.OrdinalIgnoreCase);
}

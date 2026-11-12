using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

/// <summary>Creates repeatable findings from normalized comparison data only.</summary>
public sealed class FinancialArrangementInsightService
{
    public IReadOnlyList<FinancialArrangementInsight> Generate(RenewalFinancialArrangementComparison comparison, FinancialImpactComparison? impact = null)
    {
        var insights = new List<FinancialArrangementInsight>();
        foreach (var load in comparison.AdministrativeLoadComparisons)
        {
            AddRateChange(insights, "CarrierAdministration", load.Benefit, "Carrier administrative charge", load.CurrentCarrierAdministrationRate, load.ProposedCarrierAdministrationRate, "FA-ADMIN");
            if (load.CalculationWarnings.Count > 0) insights.Add(new FinancialArrangementInsight { InsightType = "CalculationBasisChanged", Severity = FinancialArrangementInsightSeverity.Warning, Benefit = load.Benefit, Title = "Administrative load requires review", Description = string.Join(" ", load.CalculationWarnings), RuleCode = "FA-BASIS", RequiresAction = true });
        }
        if (!string.Equals(comparison.CurrentCommissionBasis, comparison.ProposedCommissionBasis, StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(comparison.CurrentCommissionBasis) && !string.IsNullOrWhiteSpace(comparison.ProposedCommissionBasis)) insights.Add(new FinancialArrangementInsight { InsightType = "CommissionBasisChanged", Severity = FinancialArrangementInsightSeverity.Warning, Title = "Commission calculation basis changed", Description = $"Commission changed from {comparison.CurrentCommissionBasis} to {comparison.ProposedCommissionBasis}; rates were not treated as directly comparable.", CurrentValue = comparison.CurrentCommissionRate, ProposedValue = comparison.ProposedCommissionRate, RuleCode = "FA-COMMISSION-BASIS", RequiresAction = true });
        else AddRateChange(insights, "Commission", "All", "Commission", comparison.CurrentCommissionRate, comparison.ProposedCommissionRate, "FA-COMMISSION");
        AddRateChange(insights, "PoolingRate", "EHC", "Pooling charge", comparison.CurrentPoolingRate, comparison.ProposedPoolingRate, "FA-POOL-RATE");
        if (comparison.CurrentPoolingThreshold != comparison.ProposedPoolingThreshold && comparison.CurrentPoolingThreshold.HasValue && comparison.ProposedPoolingThreshold.HasValue) insights.Add(new FinancialArrangementInsight { InsightType = "PoolingThresholdChanged", Severity = comparison.ProposedPoolingThreshold > comparison.CurrentPoolingThreshold ? FinancialArrangementInsightSeverity.Attention : FinancialArrangementInsightSeverity.Positive, Benefit = "EHC", Title = "Pooling threshold changed", Description = $"Pooling threshold changed from {comparison.CurrentPoolingThreshold:C0} to {comparison.ProposedPoolingThreshold:C0}.", CurrentValue = comparison.CurrentPoolingThreshold, ProposedValue = comparison.ProposedPoolingThreshold, RuleCode = "FA-POOL-THRESHOLD", RequiresAction = true });
        if (comparison.IsStale) insights.Add(new FinancialArrangementInsight { InsightType = "SnapshotStale", Severity = FinancialArrangementInsightSeverity.Warning, Title = "Financial arrangement snapshot is stale", Description = "Canonical financial terms changed after this open comparison snapshot was created.", RuleCode = "FA-STALE", RequiresAction = true });
        if (impact is not null && impact.CalculationCompleteness != FinancialCalculationCompleteness.Complete) insights.Add(new FinancialArrangementInsight { InsightType = "FinancialArrangementIncomplete", Severity = FinancialArrangementInsightSeverity.Warning, Title = "Financial arrangement is incomplete", Description = string.Join(" ", impact.MissingInputs), RuleCode = "FA-INCOMPLETE", RequiresAction = true });
        return insights;
    }

    private static void AddRateChange(List<FinancialArrangementInsight> insights, string type, string benefit, string label, decimal? current, decimal? proposed, string ruleCode)
    {
        if (!current.HasValue || !proposed.HasValue) return;
        if (current == proposed) return;
        var increased = proposed > current;
        insights.Add(new FinancialArrangementInsight { InsightType = type, Severity = increased ? FinancialArrangementInsightSeverity.Attention : FinancialArrangementInsightSeverity.Positive, Benefit = benefit, Title = $"{label} {(increased ? "increased" : "decreased")}", Description = $"{label} {(increased ? "increased" : "decreased")} from {current:P2} to {proposed:P2}.", CurrentValue = current, ProposedValue = proposed, RuleCode = ruleCode, RequiresAction = increased });
    }
}

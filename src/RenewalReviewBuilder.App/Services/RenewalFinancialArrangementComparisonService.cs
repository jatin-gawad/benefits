using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

/// <summary>Read-model builder. It does not save or edit financial terms.</summary>
public sealed class RenewalFinancialArrangementComparisonService
{
    private readonly FinancialArrangementCalculationService _calculations = new();

    public RenewalFinancialArrangementComparison Build(
        long renewalId,
        FinancialArrangement current,
        FinancialArrangement proposed,
        CommissionArrangement? currentCommission,
        CommissionArrangement? proposedCommission,
        FinancialArrangementScenario proposedScenario,
        bool isStale = false)
    {
        var comparison = new RenewalFinancialArrangementComparison
        {
            RenewalId = renewalId,
            ClientId = current.ClientId,
            Carrier = proposed.Carrier,
            ProposedScenario = proposedScenario,
            IsStale = isStale,
            SnapshotStatus = isStale ? "NeedsReview" : "Draft"
        };
        comparison.CurrentCommissionRate = currentCommission?.Rate;
        comparison.ProposedCommissionRate = proposedCommission?.Rate;
        comparison.CurrentCommissionBasis = currentCommission?.Basis ?? "";
        comparison.ProposedCommissionBasis = proposedCommission?.Basis ?? "";
        foreach (var benefit in new[] { "EHC", "Dental" })
        {
            comparison.AdministrativeLoadComparisons.Add(_calculations.Compare(benefit, current.Charges, proposed.Charges, currentCommission, proposedCommission));
        }
        return comparison;
    }
}

using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed class FinancialArrangementCalculationService
{
    private static readonly FinancialChargeType[] IncludedAdministrationTypes =
    [FinancialChargeType.GeneralAdministration, FinancialChargeType.BenefitPaymentsAdministration, FinancialChargeType.ProfitCharge];

    public decimal? CalculateCarrierAdministrationRate(IEnumerable<FinancialArrangementCharge> charges, string benefit)
    {
        var included = charges.Where(charge => IncludedAdministrationTypes.Contains(charge.ChargeType))
            .Where(charge => charge.Benefit.Equals("All", StringComparison.OrdinalIgnoreCase) || charge.Benefit.Equals(benefit, StringComparison.OrdinalIgnoreCase))
            .ToList();
        return included.Count == 0 || included.Any(charge => !charge.Rate.HasValue) ? null : included.Sum(charge => charge.Rate!.Value);
    }

    public decimal? CalculateTotalAdministrativeLoad(decimal? administrationRate, string administrationBasis, CommissionArrangement? commission)
    {
        if (!administrationRate.HasValue || commission?.Rate is null || !SameBasis(administrationBasis, commission.Basis)) return null;
        return administrationRate.Value + commission.Rate.Value;
    }

    public AdministrativeLoadComparison Compare(string benefit, IEnumerable<FinancialArrangementCharge> currentCharges, IEnumerable<FinancialArrangementCharge> proposedCharges, CommissionArrangement? currentCommission, CommissionArrangement? proposedCommission)
    {
        const string administrationBasis = "Total benefit payments";
        var currentRate = CalculateCarrierAdministrationRate(currentCharges, benefit);
        var proposedRate = CalculateCarrierAdministrationRate(proposedCharges, benefit);
        var comparison = new AdministrativeLoadComparison
        {
            Benefit = benefit,
            CurrentCarrierAdministrationRate = currentRate,
            ProposedCarrierAdministrationRate = proposedRate,
            CurrentCommissionRate = currentCommission?.Rate,
            ProposedCommissionRate = proposedCommission?.Rate,
            HasCompatibleCurrentBasis = currentCommission is not null && SameBasis(administrationBasis, currentCommission.Basis),
            HasCompatibleProposedBasis = proposedCommission is not null && SameBasis(administrationBasis, proposedCommission.Basis)
        };
        comparison.CurrentTotalAdministrativeLoadRate = CalculateTotalAdministrativeLoad(currentRate, administrationBasis, currentCommission);
        comparison.ProposedTotalAdministrativeLoadRate = CalculateTotalAdministrativeLoad(proposedRate, administrationBasis, proposedCommission);
        if (!comparison.HasCompatibleCurrentBasis || !comparison.HasCompatibleProposedBasis) comparison.CalculationWarnings.Add("Carrier administration and commission have different calculation bases; their percentages were not combined.");
        return comparison;
    }

    private static bool SameBasis(string left, string right) => left.Trim().Equals(right.Trim(), StringComparison.OrdinalIgnoreCase);
}

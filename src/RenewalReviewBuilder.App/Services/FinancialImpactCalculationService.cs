using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed class FinancialImpactCalculationService
{
    private readonly FinancialCalculationBasisResolver _resolver;

    public FinancialImpactCalculationService(FinancialCalculationBasisResolver? resolver = null) => _resolver = resolver ?? new FinancialCalculationBasisResolver();

    public FinancialImpactComparison Calculate(FinancialArrangement current, FinancialArrangement proposed, CommissionArrangement? currentCommission, CommissionArrangement? proposedCommission, PoolingArrangement? currentPooling, PoolingArrangement? proposedPooling, FinancialCalculationInputSet currentInputs, FinancialCalculationInputSet proposedInputs)
    {
        var result = new FinancialImpactComparison
        {
            CurrentCarrierAdministration = CalculateCharges(current.Charges, currentInputs, "Carrier administration"),
            ProposedCarrierAdministration = CalculateCharges(proposed.Charges, proposedInputs, "Carrier administration"),
            CurrentCommission = CalculateCommission(currentCommission, currentInputs),
            ProposedCommission = CalculateCommission(proposedCommission, proposedInputs),
            CurrentPooling = CalculatePooling(currentPooling, currentInputs),
            ProposedPooling = CalculatePooling(proposedPooling, proposedInputs),
            CurrentSpecialCharges = CalculateCharges(current.Charges, currentInputs, "Special charges", FinancialChargeType.PrintingExpense, FinancialChargeType.ReportPreparationExpense, FinancialChargeType.TripCancellation, FinancialChargeType.OtherCarrierCharge),
            ProposedSpecialCharges = CalculateCharges(proposed.Charges, proposedInputs, "Special charges", FinancialChargeType.PrintingExpense, FinancialChargeType.ReportPreparationExpense, FinancialChargeType.TripCancellation, FinancialChargeType.OtherCarrierCharge),
            CurrentTaxesAndInterest = CalculateCharges(current.Charges, currentInputs, "Taxes and interest", FinancialChargeType.PremiumTax, FinancialChargeType.InterestTerm),
            ProposedTaxesAndInterest = CalculateCharges(proposed.Charges, proposedInputs, "Taxes and interest", FinancialChargeType.PremiumTax, FinancialChargeType.InterestTerm)
        };
        result.CurrentTotalAdministrativeCost = SumIfAllAvailable(result.CurrentCarrierAdministration.Amount, result.CurrentCommission.Amount);
        result.ProposedTotalAdministrativeCost = SumIfAllAvailable(result.ProposedCarrierAdministration.Amount, result.ProposedCommission.Amount);
        result.CurrentTotalAsoExpense = SumIfAllAvailable(result.CurrentCarrierAdministration.Amount, result.CurrentCommission.Amount, result.CurrentPooling.Amount, result.CurrentSpecialCharges.Amount, result.CurrentTaxesAndInterest.Amount);
        result.ProposedTotalAsoExpense = SumIfAllAvailable(result.ProposedCarrierAdministration.Amount, result.ProposedCommission.Amount, result.ProposedPooling.Amount, result.ProposedSpecialCharges.Amount, result.ProposedTaxesAndInterest.Amount);
        CollectIssues(result, result.CurrentCarrierAdministration, result.ProposedCarrierAdministration, result.CurrentCommission, result.ProposedCommission, result.CurrentPooling, result.ProposedPooling, result.CurrentSpecialCharges, result.ProposedSpecialCharges, result.CurrentTaxesAndInterest, result.ProposedTaxesAndInterest);
        result.CalculationCompleteness = result.MissingInputs.Count == 0 ? FinancialCalculationCompleteness.Complete : result.CurrentTotalAsoExpense.HasValue || result.ProposedTotalAsoExpense.HasValue ? FinancialCalculationCompleteness.PartiallyCalculated : FinancialCalculationCompleteness.NotCalculable;
        return result;
    }

    private CalculatedFinancialAmount CalculateCharges(IEnumerable<FinancialArrangementCharge> charges, FinancialCalculationInputSet inputs, string label, params FinancialChargeType[] types)
    {
        var calculated = new List<CalculatedFinancialAmount>();
        var includedTypes = types.Length == 0 ? new[] { FinancialChargeType.GeneralAdministration, FinancialChargeType.BenefitPaymentsAdministration, FinancialChargeType.ProfitCharge } : types;
        foreach (var charge in charges.Where(charge => includedTypes.Contains(charge.ChargeType)).Where(charge => charge.Rate.HasValue || charge.Amount.HasValue))
        {
            if (charge.Amount.HasValue) calculated.Add(new CalculatedFinancialAmount { Amount = charge.Amount, Formula = charge.Formula, CalculatedAtUtc = DateTime.UtcNow });
            else calculated.Add(CalculateRate(charge.Rate!.Value, charge.Basis, inputs, $"{charge.ChargeType} rate × {charge.Basis}"));
        }
        if (calculated.Count == 0) return new CalculatedFinancialAmount { Warning = $"No {label} terms were available.", CalculatedAtUtc = DateTime.UtcNow };
        if (calculated.Any(item => !item.Amount.HasValue)) return new CalculatedFinancialAmount { Warning = $"At least one {label} component could not be calculated.", CalculatedAtUtc = DateTime.UtcNow };
        return new CalculatedFinancialAmount { Amount = calculated.Sum(item => item.Amount!.Value), Formula = $"Sum of {label} components", CalculatedAtUtc = DateTime.UtcNow };
    }

    private CalculatedFinancialAmount CalculateCommission(CommissionArrangement? commission, FinancialCalculationInputSet inputs) => commission?.Amount is not null ? new CalculatedFinancialAmount { Amount = commission.Amount, Formula = "Confirmed commission amount", CalculatedAtUtc = DateTime.UtcNow } : commission?.Rate is not null ? CalculateRate(commission.Rate.Value, commission.Basis, inputs, $"Commission rate × {commission.Basis}") : new CalculatedFinancialAmount { Warning = "Commission is unavailable.", CalculatedAtUtc = DateTime.UtcNow };
    private CalculatedFinancialAmount CalculatePooling(PoolingArrangement? pooling, FinancialCalculationInputSet inputs) => pooling?.Rate is not null ? CalculateRate(pooling.Rate.Value, pooling.Basis, inputs, $"Pooling rate × {pooling.Basis}") : new CalculatedFinancialAmount { Warning = "Pooling is unavailable.", CalculatedAtUtc = DateTime.UtcNow };
    private CalculatedFinancialAmount CalculateRate(decimal rate, string basis, FinancialCalculationInputSet inputs, string formula)
    {
        var resolved = _resolver.Resolve(basis, inputs);
        return new CalculatedFinancialAmount { Amount = resolved.Value.HasValue ? resolved.Value.Value * rate : null, BasisValue = resolved.Value, Basis = basis, BasisSource = resolved.Source, BasisPeriod = resolved.Period, Formula = formula, CalculatedAtUtc = DateTime.UtcNow, Warning = resolved.Warning };
    }
    private static decimal? SumIfAllAvailable(params decimal?[] values) => values.All(value => value.HasValue) ? values.Sum(value => value!.Value) : null;
    private static void CollectIssues(FinancialImpactComparison result, params CalculatedFinancialAmount[] amounts)
    {
        foreach (var warning in amounts
            .Select(amount => amount.Warning)
            .Where(warning => !string.IsNullOrWhiteSpace(warning))
            .Distinct(StringComparer.Ordinal))
        {
            result.MissingInputs.Add(warning);
        }
    }
}

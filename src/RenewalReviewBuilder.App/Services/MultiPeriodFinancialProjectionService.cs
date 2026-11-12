using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

/// <summary>Recalculates each annual period from that period's inputs; it never multiplies year one by the number of years.</summary>
public sealed class MultiPeriodFinancialProjectionService
{
    private readonly FinancialImpactCalculationService _impactCalculationService;
    public MultiPeriodFinancialProjectionService(FinancialImpactCalculationService? impactCalculationService = null) => _impactCalculationService = impactCalculationService ?? new FinancialImpactCalculationService();

    public MultiPeriodFinancialProjection Project(IEnumerable<FinancialProjectionPeriodInput> periods, FinancialArrangement current, FinancialArrangement proposed, CommissionArrangement? currentCommission, CommissionArrangement? proposedCommission, PoolingArrangement? currentPooling, PoolingArrangement? proposedPooling)
    {
        var projection = new MultiPeriodFinancialProjection();
        foreach (var period in periods)
        {
            projection.Periods.Add(new FinancialProjectionPeriodResult { Period = period.Period, Impact = _impactCalculationService.Calculate(current, proposed, currentCommission, proposedCommission, currentPooling, proposedPooling, period.CurrentInputs, period.ProposedInputs) });
        }
        projection.TotalProposedAdministrativeCost = SumIfAllAvailable(projection.Periods.Select(period => period.Impact.ProposedTotalAdministrativeCost));
        projection.TotalProposedAsoExpense = SumIfAllAvailable(projection.Periods.Select(period => period.Impact.ProposedTotalAsoExpense));
        return projection;
    }

    private static decimal? SumIfAllAvailable(IEnumerable<decimal?> values)
    {
        var list = values.ToList();
        return list.Count > 0 && list.All(value => value.HasValue) ? list.Sum(value => value!.Value) : null;
    }
}

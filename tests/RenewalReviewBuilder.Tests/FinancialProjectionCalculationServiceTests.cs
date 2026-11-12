using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class FinancialProjectionCalculationServiceTests
{
    [Fact]
    public void Nine_completed_aso_months_annualize_actual_recurring_components_and_credit()
    {
        var request = RequestWithMonths(9, totalBenefitPayment: 100m, admin: 10m, pooling: 5m, taxes: 2m, interest: -1m);

        var historical = new FinancialProjectionCalculationService()
            .CalculateHistoricalAnnualizedRunRate(request);

        Assert.Equal(1200m, Component(historical, FinancialCostComponentType.Claims).CalculatedAmount);
        Assert.Equal(120m, Component(historical, FinancialCostComponentType.Administration).CalculatedAmount);
        Assert.Equal(60m, Component(historical, FinancialCostComponentType.Pooling).CalculatedAmount);
        Assert.Equal(24m, Component(historical, FinancialCostComponentType.Taxes).CalculatedAmount);
        Assert.Equal(-12m, Component(historical, FinancialCostComponentType.Credits).CalculatedAmount);
        Assert.Equal(1392m, historical.TotalClientCost);
        Assert.Equal("Not Available", Component(historical, FinancialCostComponentType.OtherCharges).ValidationStatus);
    }

    [Fact]
    public void Twelve_completed_months_are_not_reannualized()
    {
        var historical = new FinancialProjectionCalculationService()
            .CalculateHistoricalAnnualizedRunRate(RequestWithMonths(12, totalBenefitPayment: 100m, admin: 0m, pooling: 0m, taxes: 0m, interest: 0m));

        Assert.Equal(1200m, Component(historical, FinancialCostComponentType.Claims).CalculatedAmount);
    }

    [Fact]
    public void Current_and_renewal_terms_apply_independently_to_the_same_projected_claims_basis()
    {
        var request = RequestWithMonths(9, 100m, 0m, 0m, 0m, 0m);
        request.ProjectedInputs.ValidatedAsoActivity.Add(new FinancialCalculationInput
        {
            Basis = "Annualized total benefit payments (9-month rolling window)",
            Value = 1200m,
            Source = "Saved ASO activity",
            IsValidated = true
        });
        request.CurrentArrangement.Charges.Add(new FinancialArrangementCharge { FinancialArrangementChargeId = 1, ChargeType = FinancialChargeType.GeneralAdministration, Rate = .10m, Basis = "Total benefit payments" });
        request.RenewalArrangement.Charges.Add(new FinancialArrangementCharge { FinancialArrangementChargeId = 2, ChargeType = FinancialChargeType.GeneralAdministration, Rate = .15m, Basis = "Total benefit payments" });
        request.CurrentCommission = new CommissionArrangement { CommissionArrangementId = 3, Rate = .05m, Basis = "Total benefit payments" };
        request.RenewalCommission = new CommissionArrangement { CommissionArrangementId = 4, Rate = .04m, Basis = "Total benefit payments" };

        var projections = new FinancialProjectionCalculationService().Calculate(request);
        var current = projections.Single(result => result.ProjectionType == FinancialProjectionType.CurrentTermsProjectedCost);
        var renewal = projections.Single(result => result.ProjectionType == FinancialProjectionType.RenewalTermsProjectedCost);

        Assert.Equal(1380m, current.TotalClientCost);
        Assert.Equal(1428m, renewal.TotalClientCost);
    }

    [Fact]
    public void Non_pooled_ehc_is_not_substituted_from_total_benefit_payments_and_is_reported_once()
    {
        var request = RequestWithMonths(9, 100m, 0m, 0m, 0m, 0m);
        request.ProjectedInputs.ValidatedAsoActivity.Add(new FinancialCalculationInput { Basis = "Annualized total benefit payments (9-month rolling window)", Value = 1200m, Source = "Saved ASO activity", IsValidated = true });
        request.CurrentPooling = new PoolingArrangement { PoolingArrangementId = 1, Rate = .077m, Basis = "Non-pooled EHC benefit payments" };

        var current = new FinancialProjectionCalculationService().Calculate(request)
            .Single(result => result.ProjectionType == FinancialProjectionType.CurrentTermsProjectedCost);

        Assert.DoesNotContain(current.Components, component => component.ComponentType == FinancialCostComponentType.Pooling && component.IncludedInTotal);
        Assert.Single(current.Validations, validation => validation.Code == "MissingBasis");
    }

    [Fact]
    public void Composite_administrative_load_excludes_separate_administration_and_commission_terms()
    {
        var request = RequestWithMonths(9, 100m, 0m, 0m, 0m, 0m);
        request.ProjectedInputs.ValidatedAsoActivity.Add(new FinancialCalculationInput { Basis = "Total benefit payments", Value = 1000m, Source = "Claims forecast", IsValidated = true });
        request.CurrentArrangement.Charges.Add(new FinancialArrangementCharge { FinancialArrangementChargeId = 1, ChargeType = FinancialChargeType.GeneralAdministration, Rate = .12m, Basis = "Total benefit payments", Formula = "Total administrative load" });
        request.CurrentArrangement.Charges.Add(new FinancialArrangementCharge { FinancialArrangementChargeId = 2, ChargeType = FinancialChargeType.GeneralAdministration, Rate = .07m, Basis = "Total benefit payments" });
        request.CurrentCommission = new CommissionArrangement { CommissionArrangementId = 3, Rate = .04m, Basis = "Total benefit payments" };

        var current = new FinancialProjectionCalculationService().Calculate(request)
            .Single(result => result.ProjectionType == FinancialProjectionType.CurrentTermsProjectedCost);

        Assert.Equal(1120m, current.TotalClientCost);
        Assert.Single(current.Validations, validation => validation.Code == "CompositeAdministrativeLoad");
        Assert.Single(current.Components, component => component.ComponentType == FinancialCostComponentType.Administration && component.IncludedInTotal);
        Assert.DoesNotContain(current.Components, component => component.ComponentType == FinancialCostComponentType.Commission && component.IncludedInTotal);
    }

    [Fact]
    public void Historical_components_retain_saved_aso_lineage()
    {
        var historical = new FinancialProjectionCalculationService()
            .CalculateHistoricalAnnualizedRunRate(RequestWithMonths(9, 100m, 10m, 5m, 2m, 0m));

        var claims = Component(historical, FinancialCostComponentType.Claims);

        Assert.Equal("1,2,3,4,5,6,7,8,9", claims.SourceFactIds);
        Assert.Equal("10", claims.SourceDocumentVersionIds);
        Assert.Equal("aso.xlsx", claims.EvidenceIds);
    }

    private static FinancialCostComponentResult Component(FinancialProjectionResult result, FinancialCostComponentType type) =>
        Assert.Single(result.Components, component => component.ComponentType == type);

    private static FinancialProjectionRequest RequestWithMonths(int monthCount, decimal totalBenefitPayment, decimal admin, decimal pooling, decimal taxes, decimal interest)
    {
        var start = new DateOnly(2025, 10, 1);
        var periods = Enumerable.Range(0, monthCount)
            .Select(index => new FinancialHistoricalPeriodSource
            {
                StatementId = 10,
                SourceDocumentId = 20,
                SourceVersionId = "v1",
                EvidenceId = "aso.xlsx",
                Period = new AsoFinancialActivityPeriod
                {
                    PeriodId = index + 1,
                    ActivityMonth = start.AddMonths(index),
                    TotalBenefitPayment = totalBenefitPayment,
                    AdminFees = admin,
                    PoolingCharges = pooling,
                    TotalTaxes = taxes,
                    InterestCreditCharge = interest
                }
            })
            .ToList();
        return new FinancialProjectionRequest
        {
            CompletedHistoricalPeriods = periods,
            CurrentArrangement = new FinancialArrangement(),
            RenewalArrangement = new FinancialArrangement()
        };
    }
}

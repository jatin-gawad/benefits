using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

/// <summary>
/// Keeps historical ASO activity and pricing-term projections separate. It is a derived read model;
/// source records and their evidence remain owned by their existing import workflows.
/// </summary>
public sealed class FinancialProjectionCalculationService
{
    private readonly FinancialCalculationBasisResolver _basisResolver;

    public FinancialProjectionCalculationService(FinancialCalculationBasisResolver? basisResolver = null) =>
        _basisResolver = basisResolver ?? new FinancialCalculationBasisResolver();

    public IReadOnlyList<FinancialProjectionResult> Calculate(FinancialProjectionRequest request) =>
    [
        CalculateHistoricalAnnualizedRunRate(request),
        CalculateTermsProjection(request, FinancialProjectionType.CurrentTermsProjectedCost, "Current terms", request.CurrentArrangement, request.CurrentCommission, request.CurrentPooling),
        CalculateTermsProjection(request, FinancialProjectionType.RenewalTermsProjectedCost, "Renewal terms", request.RenewalArrangement, request.RenewalCommission, request.RenewalPooling)
    ];

    public FinancialProjectionResult CalculateHistoricalAnnualizedRunRate(FinancialProjectionRequest request)
    {
        var periods = request.CompletedHistoricalPeriods
            .Where(source => source.Period.ActivityMonth.HasValue && !source.Period.IsTotal)
            .OrderBy(source => source.Period.ActivityMonth)
            .ToList();
        var result = new FinancialProjectionResult
        {
            ProjectionType = FinancialProjectionType.HistoricalAnnualizedRunRate,
            Scenario = "Historical ASO activity",
            CompletedMonthCount = periods.Count,
            IsInsufficientHistory = periods.Count < request.MinimumCompleteMonths
        };
        if (result.IsInsufficientHistory)
        {
            result.Validations.Add(new FinancialProjectionValidationResult
            {
                Code = "InsufficientHistory",
                RequiresReview = true,
                Message = $"At least {request.MinimumCompleteMonths} completed ASO months are required; {periods.Count} are available."
            });
            return result;
        }

        AddHistorical(result, periods, FinancialCostComponentType.Claims, period => period.TotalBenefitPayment, "Total benefit payments", true);
        AddHistorical(result, periods, FinancialCostComponentType.Administration, period => period.AdminFees, "Administration fees", true);
        AddHistorical(result, periods, FinancialCostComponentType.Pooling, period => period.PoolingCharges, "Pooling charges", true);
        AddHistorical(result, periods, FinancialCostComponentType.Taxes, period => period.TotalTaxes, "Taxes", true);
        AddHistoricalInterestAndCredits(result, periods);
        AddUnavailable(result, FinancialCostComponentType.Commission, "ASO activity does not contain a commission field.");
        AddUnavailable(result, FinancialCostComponentType.OtherCharges, "ASO Other charges are not classified as recurring or one-time in the saved source.");
        ValidateTotal(result, "HistoricalTotal");
        return result;
    }

    private FinancialProjectionResult CalculateTermsProjection(
        FinancialProjectionRequest request,
        FinancialProjectionType projectionType,
        string scenario,
        FinancialArrangement arrangement,
        CommissionArrangement? commission,
        PoolingArrangement? pooling)
    {
        var result = new FinancialProjectionResult { ProjectionType = projectionType, Scenario = scenario };
        var projectedClaims = _basisResolver.Resolve(FinancialCalculationBasisCatalog.Display(FinancialCalculationBasisId.TotalBenefitPayments), request.ProjectedInputs);
        if (projectedClaims.Value.HasValue)
        {
            result.Components.Add(new FinancialCostComponentResult
            {
                ComponentType = FinancialCostComponentType.Claims,
                ProjectedBasisAmount = projectedClaims.Value,
                CalculatedAmount = projectedClaims.Value,
                IncludedInTotal = true,
                SourceType = projectedClaims.Source,
                CalculationMethod = "Resolved projected claims basis",
                SourceDocumentVersionIds = projectedClaims.Period,
                Scenario = scenario
            });
        }
        else
        {
            AddUnavailable(result, FinancialCostComponentType.Claims, projectedClaims.Warning, scenario);
            result.Validations.Add(new FinancialProjectionValidationResult { Code = "MissingProjectedClaims", RequiresReview = true, Message = projectedClaims.Warning });
        }

        var terms = BuildTerms(arrangement, commission, pooling);
        var composites = terms.Where(term => term.IsCompositeAdministrativeLoad).ToList();
        if (composites.Count > 0)
        {
            var conflicting = string.Join(", ", terms.Where(term => term.ComponentType is FinancialCostComponentType.Administration or FinancialCostComponentType.Commission).Select(term => term.Name));
            result.Validations.Add(new FinancialProjectionValidationResult { Code = "CompositeAdministrativeLoad", RequiresReview = true, ConflictingTerms = conflicting, Message = "A total administrative load term overlaps administration and commission terms; duplicate components were excluded." });
            terms = terms.Where(term => term.IsCompositeAdministrativeLoad || term.ComponentType is not (FinancialCostComponentType.Administration or FinancialCostComponentType.Commission)).ToList();
        }

        foreach (var duplicate in terms.GroupBy(term => term.DuplicateKey).Where(group => group.Count() > 1))
        {
            var retained = duplicate.OrderBy(term => term.TermId).First();
            var excluded = duplicate.Where(term => !ReferenceEquals(term, retained)).ToList();
            terms.RemoveAll(term => excluded.Contains(term));
            result.Validations.Add(new FinancialProjectionValidationResult { Code = "DuplicateTerm", RequiresReview = true, ConflictingTerms = string.Join(", ", duplicate.Select(term => term.Name)), Message = "Duplicate pricing terms were detected; only the earliest saved term was applied." });
        }

        foreach (var term in terms)
        {
            if (IsInsuredBenefit(term.Benefit) && !term.IsExplicitlyAsoFunded)
            {
                result.Components.Add(ExcludedTerm(term, scenario, "Insured benefits are excluded unless the saved term explicitly identifies ASO funding."));
                continue;
            }
            if (term.FixedAmount.HasValue)
            {
                AddCalculatedTerm(result, term, term.FixedAmount.Value, null, "Saved flat amount", scenario);
                continue;
            }
            if (!term.Rate.HasValue)
            {
                result.Components.Add(ExcludedTerm(term, scenario, "The saved term has neither a rate nor a fixed amount."));
                continue;
            }
            var basisId = FinancialCalculationBasisCatalog.Canonicalize(term.Basis);
            if (basisId == FinancialCalculationBasisId.Unknown || basisId == FinancialCalculationBasisId.FlatAmount)
            {
                result.Components.Add(ExcludedTerm(term, scenario, "The saved percentage term does not identify a supported canonical calculation basis."));
                result.Validations.Add(new FinancialProjectionValidationResult { Code = "UnsupportedBasis", RequiresReview = true, Message = $"{term.Name} requires a supported canonical basis; saved value: {term.Basis}." });
                continue;
            }
            var resolved = _basisResolver.Resolve(FinancialCalculationBasisCatalog.Display(basisId), request.ProjectedInputs);
            if (!resolved.Value.HasValue)
            {
                result.Components.Add(ExcludedTerm(term, scenario, resolved.Warning));
                result.Validations.Add(new FinancialProjectionValidationResult { Code = "MissingBasis", RequiresReview = true, Message = resolved.Warning });
                continue;
            }
            AddCalculatedTerm(result, term, resolved.Value.Value * term.Rate.Value, resolved.Value, $"{FinancialCalculationBasisCatalog.Display(basisId)} × saved rate", scenario);
        }

        ValidateTotal(result, projectionType.ToString());
        return result;
    }

    private static void AddHistorical(
        FinancialProjectionResult result,
        IReadOnlyList<FinancialHistoricalPeriodSource> periods,
        FinancialCostComponentType component,
        Func<AsoFinancialActivityPeriod, decimal?> value,
        string label,
        bool include)
    {
        var values = periods.Select(source => value(source.Period)).ToList();
        if (values.Any(item => !item.HasValue))
        {
            AddUnavailable(result, component, $"{label} is missing for one or more completed ASO months.");
            result.Validations.Add(new FinancialProjectionValidationResult { Code = "IncompleteHistoricalComponent", RequiresReview = true, Message = $"{label} cannot be annualized because its completed-month window is incomplete." });
            return;
        }
        var actual = values.Sum(item => item!.Value);
        var annualized = actual * 12m / periods.Count;
        result.Components.Add(new FinancialCostComponentResult
        {
            ComponentType = component,
            ActualPeriodAmount = actual,
            CompletedMonthCount = periods.Count,
            AnnualizedAmount = annualized,
            CalculatedAmount = annualized,
            IncludedInTotal = include,
            SourceType = "Saved ASO activity",
            CalculationMethod = "Period amount × 12 ÷ completed month count",
            SourceFactIds = string.Join(",", periods.Select(source => source.Period.PeriodId)),
            SourceDocumentVersionIds = string.Join(",", periods.Select(source => source.StatementId).Distinct()),
            EvidenceIds = string.Join(",", periods.Select(source => source.EvidenceId).Where(id => !string.IsNullOrWhiteSpace(id)).Distinct()),
            Scenario = result.Scenario
        });
    }

    private static void AddHistoricalInterestAndCredits(FinancialProjectionResult result, IReadOnlyList<FinancialHistoricalPeriodSource> periods)
    {
        var values = periods.Select(source => source.Period.InterestCreditCharge).ToList();
        if (values.Any(value => !value.HasValue))
        {
            AddUnavailable(result, FinancialCostComponentType.Interest, "Interest is missing for one or more completed ASO months.");
            AddUnavailable(result, FinancialCostComponentType.Credits, "Interest credits are unavailable because interest is incomplete.");
            return;
        }
        var actual = values.Sum(value => value!.Value);
        var annualized = actual * 12m / periods.Count;
        var type = annualized < 0m ? FinancialCostComponentType.Credits : FinancialCostComponentType.Interest;
        result.Components.Add(new FinancialCostComponentResult
        {
            ComponentType = type,
            ActualPeriodAmount = actual,
            CompletedMonthCount = periods.Count,
            AnnualizedAmount = annualized,
            CalculatedAmount = annualized,
            IncludedInTotal = true,
            SourceType = "Saved ASO activity",
            CalculationMethod = "Period interest/credit × 12 ÷ completed month count",
            SourceFactIds = string.Join(",", periods.Select(source => source.Period.PeriodId)),
            SourceDocumentVersionIds = string.Join(",", periods.Select(source => source.StatementId).Distinct()),
            Scenario = result.Scenario
        });
    }

    private static void AddUnavailable(FinancialProjectionResult result, FinancialCostComponentType component, string reason, string? scenario = null) =>
        result.Components.Add(new FinancialCostComponentResult { ComponentType = component, SourceType = "Unavailable", Scenario = scenario ?? result.Scenario, ExclusionReason = reason, ValidationStatus = "Not Available" });

    private static void AddCalculatedTerm(FinancialProjectionResult result, PricingTerm term, decimal amount, decimal? basis, string method, string scenario)
    {
        var component = amount < 0m ? FinancialCostComponentType.Credits : term.ComponentType;
        result.Components.Add(new FinancialCostComponentResult
        {
            ComponentType = component,
            ProjectedBasisAmount = basis,
            RateOrFixedTerm = term.FixedAmount ?? term.Rate,
            CalculatedAmount = amount,
            IncludedInTotal = true,
            SourceType = "Saved pricing term",
            CalculationMethod = method,
            SourceFactIds = term.TermId.ToString(),
            SourceDocumentVersionIds = term.SourceDocumentId?.ToString() ?? "",
            EvidenceIds = term.SourceEvidence,
            Scenario = scenario
        });
    }

    private static FinancialCostComponentResult ExcludedTerm(PricingTerm term, string scenario, string reason) => new()
    {
        ComponentType = term.ComponentType,
        RateOrFixedTerm = term.FixedAmount ?? term.Rate,
        SourceType = "Saved pricing term",
        SourceFactIds = term.TermId.ToString(),
        SourceDocumentVersionIds = term.SourceDocumentId?.ToString() ?? "",
        EvidenceIds = term.SourceEvidence,
        Scenario = scenario,
        ExclusionReason = reason,
        ValidationStatus = "Input Review Required"
    };

    private static void ValidateTotal(FinancialProjectionResult result, string code)
    {
        if (result.Components.Where(component => component.IncludedInTotal).Any(component => !component.CalculatedAmount.HasValue))
            result.Validations.Add(new FinancialProjectionValidationResult { Code = code, RequiresReview = true, Message = "The displayed total excludes unavailable components; review component availability." });
    }

    private static List<PricingTerm> BuildTerms(FinancialArrangement arrangement, CommissionArrangement? commission, PoolingArrangement? pooling)
    {
        var terms = arrangement.Charges.Select(charge => new PricingTerm(
            charge.FinancialArrangementChargeId,
            charge.ChargeType.ToString(),
            ChargeComponent(charge.ChargeType),
            charge.Basis,
            charge.Rate,
            charge.Amount,
            charge.Benefit,
            arrangement.SourceDocumentId,
            charge.SourceText,
            charge.Formula.Contains("total administrative load", StringComparison.OrdinalIgnoreCase) || charge.SourceText.Contains("total administrative load", StringComparison.OrdinalIgnoreCase),
            charge.Formula.Contains("ASO", StringComparison.OrdinalIgnoreCase) || charge.SourceText.Contains("ASO", StringComparison.OrdinalIgnoreCase))).ToList();
        if (commission is not null) terms.Add(new PricingTerm(commission.CommissionArrangementId, "Commission", FinancialCostComponentType.Commission, commission.Basis, commission.Rate, commission.Amount, "All", commission.SourceDocumentId, commission.Source, false, commission.Source.Contains("ASO", StringComparison.OrdinalIgnoreCase)));
        if (pooling is not null) terms.Add(new PricingTerm(pooling.PoolingArrangementId, "Pooling", FinancialCostComponentType.Pooling, pooling.Basis, pooling.Rate, null, pooling.Benefit, pooling.SourceDocumentId, pooling.SourceText, false, pooling.SourceText.Contains("ASO", StringComparison.OrdinalIgnoreCase)));
        return terms;
    }

    private static FinancialCostComponentType ChargeComponent(FinancialChargeType type) => type switch
    {
        FinancialChargeType.GeneralAdministration or FinancialChargeType.BenefitPaymentsAdministration or FinancialChargeType.ProfitCharge => FinancialCostComponentType.Administration,
        FinancialChargeType.PremiumTax => FinancialCostComponentType.Taxes,
        FinancialChargeType.InterestTerm => FinancialCostComponentType.Interest,
        _ => FinancialCostComponentType.OtherCharges
    };

    private static bool IsInsuredBenefit(string value) => value.Equals("Life", StringComparison.OrdinalIgnoreCase) || value.Equals("LTD", StringComparison.OrdinalIgnoreCase) || value.Equals("Long Term Disability", StringComparison.OrdinalIgnoreCase);

    private sealed record PricingTerm(long TermId, string Name, FinancialCostComponentType ComponentType, string Basis, decimal? Rate, decimal? FixedAmount, string Benefit, long? SourceDocumentId, string SourceEvidence, bool IsCompositeAdministrativeLoad, bool IsExplicitlyAsoFunded)
    {
        public string DuplicateKey => $"{ComponentType}|{FinancialCalculationBasisCatalog.Canonicalize(Basis)}|{Rate}|{FixedAmount}";
    }
}

namespace RenewalReviewBuilder.App.Models;

public enum FinancialCalculationCompleteness
{
    Complete,
    PartiallyCalculated,
    NotCalculable,
    RequiresReview
}

public enum FinancialProjectionType
{
    HistoricalAnnualizedRunRate,
    CurrentTermsProjectedCost,
    RenewalTermsProjectedCost
}

public enum FinancialCostComponentType
{
    Claims,
    Administration,
    Commission,
    Pooling,
    Taxes,
    Interest,
    OtherCharges,
    Credits
}

public enum FinancialCalculationBasisId
{
    Unknown,
    TotalBenefitPayments,
    NonPooledEhcBenefitPayments,
    DentalBenefitPayments,
    BilledPremium,
    Lives,
    FlatAmount
}

public sealed class FinancialCostComponentResult
{
    public FinancialCostComponentType ComponentType { get; init; }
    public decimal? ActualPeriodAmount { get; init; }
    public int CompletedMonthCount { get; init; }
    public decimal? AnnualizedAmount { get; init; }
    public decimal? ProjectedBasisAmount { get; init; }
    public decimal? RateOrFixedTerm { get; init; }
    public decimal? CalculatedAmount { get; init; }
    public string SourceType { get; init; } = "";
    public string CalculationMethod { get; init; } = "";
    public bool IncludedInTotal { get; init; }
    public string ExclusionReason { get; init; } = "";
    public string SourceFactIds { get; init; } = "";
    public string SourceDocumentVersionIds { get; init; } = "";
    public string EvidenceIds { get; init; } = "";
    public string Scenario { get; init; } = "";
    public string ValidationStatus { get; init; } = "Available";
}

public sealed class FinancialProjectionValidationResult
{
    public string Code { get; init; } = "";
    public string Message { get; init; } = "";
    public bool RequiresReview { get; init; }
    public string ConflictingTerms { get; init; } = "";
}

public sealed class FinancialProjectionResult
{
    public FinancialProjectionType ProjectionType { get; init; }
    public string Scenario { get; init; } = "";
    public int CompletedMonthCount { get; init; }
    public bool IsInsufficientHistory { get; init; }
    public List<FinancialCostComponentResult> Components { get; } = [];
    public List<FinancialProjectionValidationResult> Validations { get; } = [];
    public decimal? TotalClientCost =>
        Components.Any(component => component.IncludedInTotal) &&
        Components.Where(component => component.IncludedInTotal).All(component => component.CalculatedAmount.HasValue)
            ? Components.Where(component => component.IncludedInTotal).Sum(component => component.CalculatedAmount!.Value)
            : null;
    public bool RequiresInputReview => Validations.Any(validation => validation.RequiresReview);
}

public sealed class FinancialHistoricalPeriodSource
{
    public AsoFinancialActivityPeriod Period { get; init; } = new();
    public long StatementId { get; init; }
    public long? SourceDocumentId { get; init; }
    public string SourceVersionId { get; init; } = "";
    public string EvidenceId { get; init; } = "";
}

public sealed class FinancialProjectionRequest
{
    public IReadOnlyList<FinancialHistoricalPeriodSource> CompletedHistoricalPeriods { get; init; } = [];
    public int MinimumCompleteMonths { get; init; } = 6;
    public FinancialArrangement CurrentArrangement { get; init; } = new();
    public FinancialArrangement RenewalArrangement { get; init; } = new();
    public CommissionArrangement? CurrentCommission { get; set; }
    public CommissionArrangement? RenewalCommission { get; set; }
    public PoolingArrangement? CurrentPooling { get; set; }
    public PoolingArrangement? RenewalPooling { get; set; }
    public FinancialCalculationInputSet ProjectedInputs { get; init; } = new();
}

public sealed class FinancialCalculationInput
{
    public string Basis { get; set; } = "";
    public decimal? Value { get; set; }
    public string Source { get; set; } = "";
    public string Period { get; set; } = "";
    public bool IsValidated { get; set; }
}

public sealed class FinancialCalculationInputSet
{
    public List<FinancialCalculationInput> FinalizedRenewalProjection { get; } = [];
    public List<FinancialCalculationInput> ConfirmedRenewalFinancialAnalysis { get; } = [];
    public List<FinancialCalculationInput> ValidatedClaimsProjection { get; } = [];
    public List<FinancialCalculationInput> ValidatedAsoActivity { get; } = [];
    public List<FinancialCalculationInput> ActiveClientFinancialData { get; } = [];
}

public sealed class ResolvedFinancialCalculationBasis
{
    public string Basis { get; set; } = "";
    public decimal? Value { get; set; }
    public string Source { get; set; } = "Unavailable";
    public string Period { get; set; } = "";
    public DateTime ResolvedAtUtc { get; set; }
    public string Warning { get; set; } = "";
}

public sealed class CalculatedFinancialAmount
{
    public decimal? Amount { get; set; }
    public decimal? BasisValue { get; set; }
    public string Basis { get; set; } = "";
    public string BasisSource { get; set; } = "";
    public string BasisPeriod { get; set; } = "";
    public string Formula { get; set; } = "";
    public DateTime CalculatedAtUtc { get; set; }
    public string Warning { get; set; } = "";
}

public sealed class FinancialImpactComparison
{
    public CalculatedFinancialAmount CurrentCarrierAdministration { get; set; } = new();
    public CalculatedFinancialAmount ProposedCarrierAdministration { get; set; } = new();
    public CalculatedFinancialAmount CurrentCommission { get; set; } = new();
    public CalculatedFinancialAmount ProposedCommission { get; set; } = new();
    public CalculatedFinancialAmount CurrentPooling { get; set; } = new();
    public CalculatedFinancialAmount ProposedPooling { get; set; } = new();
    public CalculatedFinancialAmount CurrentSpecialCharges { get; set; } = new();
    public CalculatedFinancialAmount ProposedSpecialCharges { get; set; } = new();
    public CalculatedFinancialAmount CurrentTaxesAndInterest { get; set; } = new();
    public CalculatedFinancialAmount ProposedTaxesAndInterest { get; set; } = new();
    public decimal? CurrentTotalAdministrativeCost { get; set; }
    public decimal? ProposedTotalAdministrativeCost { get; set; }
    public decimal? CurrentTotalAsoExpense { get; set; }
    public decimal? ProposedTotalAsoExpense { get; set; }
    public FinancialCalculationCompleteness CalculationCompleteness { get; set; }
    public List<string> MissingInputs { get; } = [];
    public List<string> CalculationWarnings { get; } = [];
}

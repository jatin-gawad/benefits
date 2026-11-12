namespace RenewalReviewBuilder.App.Models;

public static class AsoFinancialImportStatuses
{
    public const string Complete = "Complete";
    public const string NeedsReview = "Needs Review";
}

public sealed class AsoFinancialActivityStatement
{
    public long StatementId { get; set; }
    public long ClientId { get; set; }
    public long? SourceDocumentId { get; set; }
    public string ClientName { get; set; } = "";
    public string ContractNumber { get; set; } = "";
    public string Carrier { get; set; } = "";
    public string BillingArrangement { get; set; } = "";
    public string Divisions { get; set; } = "";
    public string Benefits { get; set; } = "";
    public DateOnly? PeriodStart { get; set; }
    public DateOnly? PeriodEnd { get; set; }
    public DateOnly? StatementDate { get; set; }
    public string SourceFundingStatus { get; set; } = "";
    public int? SourcePageStart { get; set; }
    public int? SourcePageEnd { get; set; }
    public string SourceFilePath { get; set; } = "";
    public string SourceFileName { get; set; } = "";
    public string SourceSheet { get; set; } = "";
    public string SourceHash { get; set; } = "";
    public int VersionNumber { get; set; } = 1;
    public bool IsCurrent { get; set; } = true;
    public string ImportStatus { get; set; } = AsoFinancialImportStatuses.Complete;
    public string ValidationSummary { get; set; } = "";
    public DateTime ImportedAt { get; set; }
    public string ImportedBy { get; set; } = "";
    public List<AsoFinancialActivityPeriod> Periods { get; } = [];
    public AsoFinancialActivityPeriod? Total { get; set; }
    public List<string> ValidationIssues { get; } = [];
}

public sealed class AsoFinancialActivityPeriod
{
    public long PeriodId { get; set; }
    public long StatementId { get; set; }
    public DateOnly? ActivityMonth { get; set; }
    public string PeriodLabel { get; set; } = "";
    public bool IsTotal { get; set; }
    public int DisplayOrder { get; set; }
    public decimal? OpeningBalance { get; set; }
    public decimal? PaymentReceivedNotAllocated { get; set; }
    public decimal? DentalClaimsPaid { get; set; }
    public decimal? ClaimsHandlingCharges { get; set; }
    public decimal? GeneralAdminCharge { get; set; }
    public decimal? ProfitCharge { get; set; }
    public decimal? TotalServiceFees { get; set; }
    public decimal? ConsultingFees { get; set; }
    public decimal? Transfers { get; set; }
    public decimal? DeficitPaymentRefund { get; set; }
    public decimal? TotalDeposit { get; set; }
    public decimal? OrstDeposit { get; set; }
    public decimal? QstDeposit { get; set; }
    public decimal? HstDeposit { get; set; }
    public decimal? TvqDeposit { get; set; }
    public decimal? SkpstDeposit { get; set; }
    public decimal? NetDeposit { get; set; }
    public decimal? UnpooledBenefitPayments { get; set; }
    public decimal? PooledBenefitPayments { get; set; }
    public decimal? TotalBenefitPayment { get; set; }
    public decimal? LossRatio { get; set; }
    public decimal? PoolingCharges { get; set; }
    public decimal? AdminFees { get; set; }
    public decimal? OtherCharges { get; set; }
    public decimal? TotalTaxes { get; set; }
    public decimal? Gst { get; set; }
    public decimal? Hst { get; set; }
    public decimal? Tvq { get; set; }
    public decimal? InsuredPremiumTax { get; set; }
    public decimal? PremiumTaxQuebec { get; set; }
    public decimal? PremiumTaxOntario { get; set; }
    public decimal? PremiumTaxNewfoundland { get; set; }
    public decimal? RetroactiveTaxQst { get; set; }
    public decimal? RetroactiveTaxOrst { get; set; }
    public decimal? RetroactiveTaxSkpst { get; set; }
    public decimal? InterestCreditCharge { get; set; }
    public decimal? TotalCharges { get; set; }
    public decimal? SubTotal { get; set; }
    public decimal? RequiredFloat { get; set; }
    public decimal? CurrentBalance { get; set; }
    public string BalanceStatus { get; set; } = "";
}

public sealed class AsoFinancialActivitySaveResult
{
    public bool Succeeded { get; init; }
    public string Message { get; init; } = "";
    public string DatabasePath { get; init; } = "";
    public long? ClientId { get; init; }
    public long? DocumentId { get; init; }
    public long? StatementId { get; init; }
    public int InsertedRows { get; init; }
    public bool WasDuplicate { get; init; }
    public bool RequiresVerifiedAsoPolicyLink { get; init; }
    public string WorkbookContractNumber { get; init; } = "";
    public string SelectedClientNumber { get; init; } = "";
}

/// <summary>
/// A saved monthly ASO observation used for a read-only annual forecast.
/// The source statement and import timestamp preserve the audit trail when
/// duplicate source files contain the same activity month.
/// </summary>
public sealed record AsoFinancialActivityMonthObservation(
    DateOnly ActivityMonth,
    decimal? TotalBenefitPayment,
    string SourceFileName,
    DateTime ImportedAt);

/// <summary>
/// Validated rolling ASO activity used to annualize projected costs. A result
/// is complete only when every requested consecutive month has a benefit total.
/// </summary>
public sealed class AsoAnnualForecastWindowResult
{
    public int RequestedMonths { get; init; }
    public DateOnly? StartMonth { get; init; }
    public DateOnly? EndMonth { get; init; }
    public IReadOnlyList<AsoFinancialActivityMonthObservation> Observations { get; init; } = [];
    public IReadOnlyList<DateOnly> MissingMonths { get; init; } = [];
    public IReadOnlyList<DateOnly> MissingBenefitPaymentMonths { get; init; } = [];

    public bool HasSourceMonths => EndMonth.HasValue;
    public bool IsComplete =>
        HasSourceMonths &&
        Observations.Count == RequestedMonths &&
        MissingMonths.Count == 0 &&
        MissingBenefitPaymentMonths.Count == 0;

    public decimal? ObservedBenefitPayments =>
        IsComplete ? Observations.Sum(observation => observation.TotalBenefitPayment!.Value) : null;

    public decimal? AnnualizedBenefitPayments =>
        ObservedBenefitPayments is decimal observed
            ? observed * 12m / RequestedMonths
            : null;
}

/// <summary>
/// The experience period explicitly recorded by a processed renewal source.
/// It is deliberately separate from an ASO statement period: a missing value
/// means the application must not infer that the two periods agree.
/// </summary>
public sealed class AsoExperiencePeriodReference
{
    public DateOnly? PeriodStart { get; init; }
    public DateOnly? PeriodEnd { get; init; }
    public string SourceDescription { get; init; } = "";
}

namespace RenewalReviewBuilder.App.Models;

public static class RateLineValidationStatuses
{
    public const string Supported = "Supported";
    public const string NeedsReview = "Needs Review";
}

public sealed class RateLine
{
    public string Benefit { get; set; } = "";
    public string TierOrClass { get; set; } = "";
    public decimal? Volume { get; set; }
    public int? Lives { get; set; }
    public decimal? CurrentRate { get; set; }
    public decimal? CurrentPremium { get; set; }
    public decimal? RenewalRate { get; set; }
    public decimal? RenewalPremium { get; set; }
    public decimal? CarrierProposedRate { get; set; }
    public decimal? CarrierProposedPremium { get; set; }
    public decimal? NegotiatedRate { get; set; }
    public decimal? NegotiatedPremium { get; set; }
    public decimal? PercentChange { get; set; }
    public decimal? CalculatedCurrentPremium { get; set; }
    public decimal? CalculatedRenewalPremium { get; set; }
    public decimal? CurrentVariance { get; set; }
    public decimal? RenewalVariance { get; set; }
    public string SourceReference { get; set; } = "";
    public string AccountClass { get; set; } = "";
    /// <summary>Carrier-stated funding basis, for example Insured or ASO - Budgeted.</summary>
    public string FundingArrangement { get; set; } = "";

    /// <summary>Authoritative calculation basis used by the application, when it can be established without guessing.</summary>
    public string RateBasis { get; set; } = "";

    /// <summary>Whether this printed row contributes to the document's premium total. Source component/subtotal rows can remain visible without double counting.</summary>
    public bool IncludeInPremiumTotal { get; set; } = true;

    /// <summary>Source role such as Detail, Summary, or Conditional.</summary>
    public string SourceRowRole { get; set; } = "Detail";

    /// <summary>Supported or Needs Review. Blank retains the behavior of existing carrier parsers.</summary>
    public string ValidationStatus { get; set; } = "";
    public string ValidationMessage { get; set; } = "";
    public string SourceCarrier { get; set; } = "";
    public string SourceFileName { get; set; } = "";
    public int? SourcePage { get; set; }
    public string SourceText { get; set; } = "";
    public int? CurrentRatePrecision { get; set; }
    public int? RenewalRatePrecision { get; set; }
    public string CurrentPeriodLabel { get; set; } = "";
    public string RenewalPeriodLabel { get; set; } = "";
    public DateOnly? RenewalEffectiveDate { get; set; }
}

public sealed class ExperienceLine
{
    public string Benefit { get; set; } = "";
    public string Period { get; set; } = "";
    public int? AverageLives { get; set; }
    public decimal? BilledPremium { get; set; }
    public decimal? PaidClaims { get; set; }
    public decimal? IncurredClaims { get; set; }
    public decimal? PaidLossRatio { get; set; }
    public decimal? IncurredLossRatio { get; set; }
    public decimal? Weighting { get; set; }
    public decimal? ProposedAdjustment { get; set; }
    public decimal? TrendFactor { get; set; }
    public decimal? Credibility { get; set; }
    public decimal? TargetLossRatio { get; set; }
    public string SourceReference { get; set; } = "";
}

public sealed class ClaimTrend
{
    public string Benefit { get; set; } = "";
    public string Category { get; set; } = "";
    public string Period { get; set; } = "";
    public decimal? Amount { get; set; }
    public decimal? PercentOfTotal { get; set; }
    public decimal? PriorPeriodPercent { get; set; }
    public decimal? BenchmarkPercent { get; set; }
    public decimal? RegionalIndustryBenchmarkPercent { get; set; }
    public string SourceReference { get; set; } = "";
}

public sealed class ServiceCategoryTrend
{
    public string Benefit { get; set; } = "";
    public string Category { get; set; } = "";
    public string PriorPeriodLabel { get; set; } = "";
    public decimal PriorAmount { get; set; }
    public string CurrentPeriodLabel { get; set; } = "";
    public decimal CurrentAmount { get; set; }
    public string SourceReference { get; set; } = "";
    public List<ServiceCategoryPeriodValue> Periods { get; } = [];
}

public sealed class ServiceCategoryPeriodValue
{
    public string PeriodLabel { get; set; } = "";
    public decimal Amount { get; set; }
    public string SourceReference { get; set; } = "";
}

public sealed class DrugTrend
{
    public string Din { get; set; } = "";
    public string DrugName { get; set; } = "";
    public string Treats { get; set; } = "";
    public string Period { get; set; } = "";
    public int? PrescriptionCount { get; set; }
    public decimal? SubmittedTotal { get; set; }
    public decimal? NetAmountPaid { get; set; }
    public decimal? AllowableTotal { get; set; }
    public decimal? PriorPaid { get; set; }
    public string SourceReference { get; set; } = "";
}

public sealed class IndicationTrend
{
    public string Indication { get; set; } = "";
    public int? PrescriptionCount { get; set; }
    public decimal? SubmittedTotal { get; set; }
    public decimal? AllowableTotal { get; set; }
    public decimal? PriorPaid { get; set; }
    public string SourceReference { get; set; } = "";
}

public sealed class QuartileMetric
{
    public string Benefit { get; set; } = "";
    public string Quartile { get; set; } = "";
    public string Category { get; set; } = "";
    public decimal? PercentOfTotal { get; set; }
    public decimal? AverageClaimCost { get; set; }
    public string SourceReference { get; set; } = "";
}

public sealed class CertificateClaim
{
    public string Certificate { get; set; } = "";
    public decimal? ExtendedHealthCare { get; set; }
    public decimal? Dental { get; set; }
    public decimal? Total { get; set; }
    public string Quartile { get; set; } = "";
    public string SourceReference { get; set; } = "";
}

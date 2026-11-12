using RenewalReviewBuilder.App.Services;

namespace RenewalReviewBuilder.App.Models;

public static class RenewalVersionNames
{
    public const string CarrierProposed = "Carrier Proposed";
    public const string ConsultantModel = "Consultant Model";
    public const string Negotiated = "Negotiated";
    public const string FinalApproved = "Final Approved";
    public const string AlternateOption = "Alternate Option";
}

public static class RateValueBasisNames
{
    public const string CarrierStated = "Carrier-stated";
    public const string Net = "Net";
    public const string Gross = "Gross";
}

public sealed class RenewalVersion
{
    public string VersionName { get; set; } = RenewalVersionNames.CarrierProposed;
    public string SourceModule { get; set; } = BenefitsIntelligenceModules.RenewalIntelligence;
    public string Status { get; set; } = "";
    public bool IsSelected { get; set; }
    public DateTime CreatedAt { get; set; } = DateTime.Now;
    public decimal? RequestedIncreasePercent { get; set; }
    public decimal? NegotiatedIncreasePercent { get; set; }
    public decimal? FinalAcceptedIncreasePercent { get; set; }
    public NegotiationStatus? NegotiationStatus { get; set; }
    public string? NegotiationStage { get; set; }
    public DateOnly? NegotiationCompletedDate { get; set; }
    public decimal? NegotiatedSavingsAnnual { get; set; }
    public decimal? NegotiatedSavingsMonthly { get; set; }
    public NegotiatedSavingsType? NegotiatedSavingsType { get; set; }
    public string? NegotiatedBy { get; set; }
    public decimal? IndustryBenchmarkPercent { get; set; }
    public string? IndustryBenchmarkSource { get; set; }
    public string? IndustryBenchmarkPopulation { get; set; }
    public int? IndustryBenchmarkYear { get; set; }
    public decimal? IndustryBenchmarkConfidence { get; set; }
    public RenewalRecommendationType? RenewalRecommendationType { get; set; }
    public string? RenewalDecisionReason { get; set; }
    public string? RenewalStrategySummary { get; set; }
    public RenewalRiskLevel? RenewalRiskLevel { get; set; }
    public string? RenewalComplexity { get; set; }
    public string? RenewalResult { get; set; }
    public DateOnly? RenewalCompletedDate { get; set; }
    public string? RenewalApprovedBy { get; set; }
    public DateOnly? RenewalAcceptedDate { get; set; }
    public int? RenewalVersionNumber { get; set; }
    public long? RenewalScenarioId { get; set; }
    public long? PreviousRenewalScenarioId { get; set; }
    public ScenarioType? ScenarioType { get; set; }
    public long? HistoryEventId { get; set; }
    public string? RelationshipMilestoneId { get; set; }
    public string? EvolutionGroupId { get; set; }
    public List<RenewalVersionRateLine> RateLines { get; } = [];
    public List<UnderwritingCalculationResult> Calculations { get; } = [];

    public bool HasRateLines => RateLines.Count > 0;

    public decimal? CurrentMonthlyPremium => HasRateLines
        ? RateLines.Sum(line => line.CurrentPremium ?? 0m)
        : null;

    public decimal? RenewalMonthlyPremium => HasRateLines
        ? RateLines.Sum(line => line.VersionPremium ?? line.ProposedPremium ?? 0m)
        : null;

    public decimal? MonthlyPremiumChange => CurrentMonthlyPremium.HasValue && RenewalMonthlyPremium.HasValue
        ? RenewalMonthlyPremium.Value - CurrentMonthlyPremium.Value
        : null;

    public decimal? MonthlyPremiumChangePercent => UnderwritingEngine.CalculatePercentChange(CurrentMonthlyPremium, RenewalMonthlyPremium);
}

public sealed class RenewalVersionRateLine
{
    public string Benefit { get; set; } = "";
    public string TierOrClass { get; set; } = "";
    public string AccountClass { get; set; } = "";
    public decimal? Volume { get; set; }
    public int? Lives { get; set; }
    public decimal? CurrentRate { get; set; }
    public decimal? CurrentPremium { get; set; }
    public decimal? ProposedRate { get; set; }
    public decimal? ProposedPremium { get; set; }
    public decimal? VersionRate { get; set; }
    public decimal? VersionPremium { get; set; }
    public decimal? ChangePercent { get; set; }
    public string RateBasis { get; set; } = "";
    public string RateValueBasis { get; set; } = RateValueBasisNames.CarrierStated;
    public string ClientPresentationRateBasis { get; set; } = RateValueBasisNames.CarrierStated;
    public decimal? CurrentNetRate { get; set; }
    public decimal? CurrentGrossRate { get; set; }
    public decimal? ProposedNetRate { get; set; }
    public decimal? ProposedGrossRate { get; set; }
    public decimal? VersionNetRate { get; set; }
    public decimal? VersionGrossRate { get; set; }
    public string RateScheduleSourceFile { get; set; } = "";
    public string RateScheduleSourceSheet { get; set; } = "";
    public int? RateScheduleSourceRow { get; set; }
    public string FundingArrangement { get; set; } = "";
    public string SourceReference { get; set; } = "";
    public bool IncludeInPremiumTotal { get; set; } = true;
    public string SourceRowRole { get; set; } = "Detail";
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

    public bool HasNetAndGrossRates =>
        CurrentNetRate.HasValue || CurrentGrossRate.HasValue ||
        ProposedNetRate.HasValue || ProposedGrossRate.HasValue ||
        VersionNetRate.HasValue || VersionGrossRate.HasValue;
}

public sealed class UnderwritingCalculationInput
{
    public string SourceModule { get; set; } = BenefitsIntelligenceModules.UnderwritingIntelligence;
    public string VersionName { get; set; } = RenewalVersionNames.ConsultantModel;
    public string BenefitLine { get; set; } = "";
    public decimal? CurrentClaims { get; set; }
    public decimal? CurrentPremium { get; set; }
    public decimal? ProposedPremium { get; set; }
    public decimal? NegotiatedPremium { get; set; }
    public decimal? CurrentRate { get; set; }
    public decimal? ProposedRate { get; set; }
    public decimal? NegotiatedRate { get; set; }
    public decimal? Volume { get; set; }
    public int? Lives { get; set; }
    public string RateBasis { get; set; } = "";
    public decimal? Trend { get; set; }
    public decimal? TargetLossRatio { get; set; }
    public decimal? PoolingAdjustment { get; set; }
    public decimal? IbnrAdjustment { get; set; }
    public decimal? RetentionAdminFee { get; set; }
    public decimal? ManualAdjustment { get; set; }
}

public sealed class UnderwritingCalculationResult
{
    public string SourceModule { get; set; } = BenefitsIntelligenceModules.UnderwritingIntelligence;
    public string VersionName { get; set; } = RenewalVersionNames.ConsultantModel;
    public string BenefitLine { get; set; } = "";
    public decimal? CurrentClaims { get; set; }
    public decimal? CurrentPremium { get; set; }
    public decimal? TrendedClaims { get; set; }
    public decimal? PoolingAdjustment { get; set; }
    public decimal? IbnrAdjustment { get; set; }
    public decimal? RetentionAdminFeeAdjustment { get; set; }
    public decimal? ManualAdjustment { get; set; }
    public decimal? ProjectedClaims { get; set; }
    public decimal? RequiredPremium { get; set; }
    public decimal? ProposedPremium { get; set; }
    public decimal? NegotiatedPremium { get; set; }
    public decimal? RequiredRate { get; set; }
    public decimal? TargetLossRatio { get; set; }
    public decimal? ProjectedLossRatio { get; set; }
    public decimal? RenewalChangePercent { get; set; }
    public decimal? NegotiatedChangePercent { get; set; }
    public decimal? Savings { get; set; }
    public decimal? Trend { get; set; }
}

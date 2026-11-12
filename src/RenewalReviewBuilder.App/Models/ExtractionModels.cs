namespace RenewalReviewBuilder.App.Models;

public sealed class PdfExtractionResult
{
    public string SourcePath { get; set; } = "";
    public string Carrier { get; set; } = "";
    public string ClientName { get; set; } = "";
    public string PolicyNumber { get; set; } = "";
    public DateOnly? RenewalDate { get; set; }
    public List<RateLine> RateLines { get; } = [];
    public List<AsoBudgetedBenefit> AsoBudgetedBenefits { get; } = [];
    public List<ExperienceLine> ExperienceLines { get; } = [];
    public List<SunLifeExperienceRow> SunLifeExperienceRows { get; } = [];
    public List<EmpireLifePooledBenefit> EmpireLifePooledBenefits { get; } = [];
    public List<EmpireLifeExperienceRow> EmpireLifeExperienceRows { get; } = [];
    public List<AssurePlusPricingReport> AssurePlusPricingReports { get; } = [];
    public List<DemographicLine> LifeDemographics { get; } = [];
    public List<DemographicLine> LtdDemographics { get; } = [];
    public decimal? CurrentMonthlyPremium { get; set; }
    public decimal? RenewalMonthlyPremium { get; set; }
    public decimal? MonthlyPremiumChange { get; set; }
    public decimal? MonthlyPremiumChangePercent { get; set; }
    public decimal? CurrentEhcPremium { get; set; }
    public decimal? RenewalEhcPremium { get; set; }
    public decimal? CurrentDentalPremium { get; set; }
    public decimal? RenewalDentalPremium { get; set; }
    public decimal? PoolingCharge { get; set; }
    public decimal? PoolingThreshold { get; set; }
    public decimal? PooledPremium { get; set; }
    public decimal? DrugStopLossClaims { get; set; }
    public decimal? OtherPooledClaims { get; set; }
    public List<string> Warnings { get; } = [];
}

/// <summary>
/// A carrier-stated ASO budget arrangement. These amounts are budgeted plan
/// contributions, not insured premium, and the expense percentage excludes premium tax when so stated by the source.
/// </summary>
public sealed class AsoBudgetedBenefit
{
    public string Benefit { get; set; } = "";
    public string UnderwritingArrangement { get; set; } = "ASO - Budgeted";
    public decimal? CurrentExpensePercent { get; set; }
    public decimal? RenewalExpensePercent { get; set; }
    public string SourceReference { get; set; } = "";
}

public sealed class SunLifeExperienceRow
{
    public string Benefit { get; set; } = "";
    public string Period { get; set; } = "";
    public int? AverageCoveredEmployees { get; set; }
    public decimal? BilledPremium { get; set; }
    public decimal? AdjustedPremium { get; set; }
    public decimal? PaidClaims { get; set; }
    public decimal? IncurredClaims { get; set; }
    public decimal? AdjustedLossRatio { get; set; }
    public decimal? Weighting { get; set; }
    public decimal? LargeAmountPoolingPremium { get; set; }
    public decimal? QuebecIndustryPoolingPremium { get; set; }
    public decimal? PooledClaims { get; set; }
    public decimal? QuebecIndustryPooledClaims { get; set; }
    public decimal? WeightedAdjustedLossRatio { get; set; }
    public decimal? Credibility { get; set; }
    public decimal? Ibnr { get; set; }
    public decimal? TargetLossRatio { get; set; }
    public decimal? AnnualTrendFactor { get; set; }
    public decimal? RateAdjustment { get; set; }
    public string SourceReference { get; set; } = "";
}

public sealed class EmpireLifePooledBenefit
{
    public string Benefit { get; set; } = "";
    public decimal? BilledPremium { get; set; }
    public decimal? PaidClaims { get; set; }
    public string SourceReference { get; set; } = "";
}

public sealed class EmpireLifeExperienceRow
{
    public string Benefit { get; set; } = "";
    public string Period { get; set; } = "";
    public decimal? NonPooledPremium { get; set; }
    public decimal? AdjustedPremium { get; set; }
    public decimal? PaidClaims { get; set; }
    public decimal? TrendedClaims { get; set; }
    public decimal? IncurredClaims { get; set; }
    public decimal? IncurredLossRatio { get; set; }
    public decimal? WeightedLossRatio { get; set; }
    public decimal? TargetLossRatio { get; set; }
    public decimal? Credibility { get; set; }
    public string SourceReference { get; set; } = "";
}

public sealed class WorkbookExtractionResult
{
    public string SourcePath { get; set; } = "";
    public List<ClaimTrend> ClaimTrends { get; } = [];
    public List<ClaimsDistributionRow> ClaimsDistributionRows { get; } = [];
    public ClaimsDistributionValidation ClaimsDistributionValidation { get; set; } = new();
    public List<ServiceCategoryTrend> EhcSpendingTrends { get; } = [];
    public List<ServiceCategoryTrend> DentalSpendingTrends { get; } = [];
    public List<ServiceCategoryTrend> ParamedicalServiceTrends { get; } = [];
    public List<ClaimsTrendValidation> ClaimsTrendValidations { get; } = [];
    public List<DrugTrend> DrugTrends { get; } = [];
    public List<DrugTrend> PriorDrugTrends { get; } = [];
    public List<IndicationTrend> IndicationTrends { get; } = [];
    public List<QuartileMetric> QuartileMetrics { get; } = [];
    public List<CertificateClaim> CertificateClaims { get; } = [];
    public List<string> MissingSheets { get; } = [];
    public List<string> Warnings { get; } = [];
}

public sealed class ClaimsTrendValidation
{
    public string Area { get; set; } = "";
    public string ClientName { get; set; } = "";
    public string PolicyNumber { get; set; } = "";
    public string PriorPeriod { get; set; } = "";
    public string CurrentPeriod { get; set; } = "";
    public string SourceSheetUsed { get; set; } = "";
    public string Status { get; set; } = "";
    public string Warning { get; set; } = "";
}

public sealed class ClaimsDistributionRow
{
    public string CarrierName { get; set; } = "";
    public string PolicyNumber { get; set; } = "";
    public string ClientName { get; set; } = "";
    public string ReportPeriod { get; set; } = "";
    public string PeriodStart { get; set; } = "";
    public string PeriodEnd { get; set; } = "";
    public string PriorReportPeriod { get; set; } = "";
    public string PriorPeriodStart { get; set; } = "";
    public string PriorPeriodEnd { get; set; } = "";
    public string Benefit { get; set; } = "";
    public string Division { get; set; } = "";
    public string Class { get; set; } = "";
    public string ReportCreatedDate { get; set; } = "";
    public string SourceFile { get; set; } = "";
    public string ServiceCategory { get; set; } = "";
    public string ServiceDescriptionOriginal { get; set; } = "";
    public decimal? AmountPaid { get; set; }
    public decimal? PercentPaid { get; set; }
    public decimal? PriorAmountPaid { get; set; }
    public decimal? PriorPercentPaid { get; set; }
    public decimal? NormPercentPaid { get; set; }
    public int? Occurrences { get; set; }
    public int? PriorOccurrences { get; set; }
    public decimal? PercentOccurrences { get; set; }
    public decimal? NormPercentOccurrences { get; set; }
    public decimal? AverageClaim { get; set; }
    public string SourceReference { get; set; } = "";
}

public sealed class ClaimsDistributionValidation
{
    public bool TableFound { get; set; }
    public bool ParamedicalFound { get; set; }
    public bool ParamedicalDetailFound { get; set; }
    public bool UnclassifiedParamedicalDetailFound { get; set; }
    public bool TotalRowFound { get; set; }
    public decimal? TotalAmountPaid { get; set; }
    public decimal? SumOfServiceAmounts { get; set; }
    public bool AmountsBalanceToTotal { get; set; }
    public decimal? ParamedicalSubtotalAmountPaid { get; set; }
    public decimal? SumOfParamedicalDetailAmounts { get; set; }
    public decimal? ParamedicalDetailVariance { get; set; }
    public bool? ParamedicalDetailsBalanceToTotal { get; set; }
    public decimal? PriorParamedicalSubtotalAmountPaid { get; set; }
    public decimal? SumOfPriorParamedicalDetailAmounts { get; set; }
    public decimal? PriorParamedicalDetailVariance { get; set; }
    public bool? PriorParamedicalDetailsBalanceToTotal { get; set; }
    public List<string> Warnings { get; } = [];
}

public sealed class PdfPageText
{
    public int PageNumber { get; set; }
    public string Text { get; set; } = "";
    public List<string> Lines { get; set; } = [];
}

public sealed class DemographicLine
{
    public string Benefit { get; set; } = "";
    public string AgeBand { get; set; } = "";
    public int? CurrentFemaleLives { get; set; }
    public decimal? CurrentFemaleVolume { get; set; }
    public int? CurrentMaleLives { get; set; }
    public decimal? CurrentMaleVolume { get; set; }
    public int? PriorFemaleLives { get; set; }
    public decimal? PriorFemaleVolume { get; set; }
    public int? PriorMaleLives { get; set; }
    public decimal? PriorMaleVolume { get; set; }
    public string SourceReference { get; set; } = "";
}

public sealed class NemMaximumRow
{
    public string Benefit { get; set; } = "";
    public string CurrentMaximum { get; set; } = "";
    public string CurrentNem { get; set; } = "";
    public string AvailableNem { get; set; } = "";
    public string SourceFileName { get; set; } = "";
    public int? SourcePage { get; set; }
    public string SourceReference { get; set; } = "";
    public decimal ConfidenceScore { get; set; }
}

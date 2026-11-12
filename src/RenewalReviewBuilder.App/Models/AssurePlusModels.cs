namespace RenewalReviewBuilder.App.Models;

public sealed class AssurePlusPricingReport
{
    public string SourcePath { get; set; } = "";
    public int SourcePage { get; set; }
    public DateOnly? EffectiveDate { get; set; }
    public List<AssurePlusBenefitPricing> Benefits { get; } = [];
    public List<AssurePlusDistributionRow> DistributionRows { get; } = [];
    public List<string> Warnings { get; } = [];
    public string SourceReference => SourcePage > 0
        ? $"{Path.GetFileName(SourcePath)} page {SourcePage}"
        : Path.GetFileName(SourcePath);
}

public sealed class AssurePlusBenefitPricing
{
    public string Benefit { get; set; } = "";
    public string PreviousYearRating { get; set; } = "";
    public string RecentYearRating { get; set; } = "";
    public decimal? StandardChangePercent { get; set; }
    public decimal? ExperienceAdjustmentPercent { get; set; }
    public decimal? TotalChangePercent { get; set; }
    public string SourceReference { get; set; } = "";
}

public sealed class AssurePlusDistributionRow
{
    public string Benefit { get; set; } = "";
    public string Category { get; set; } = "";
    public decimal? YourClaimsPercent { get; set; }
    public decimal? BenchmarkClaimsPercent { get; set; }
    public string SourceReference { get; set; } = "";
}

public sealed class AssurePlusExperienceSummary
{
    public DateOnly? EffectiveDate { get; set; }
    public List<AssurePlusAnnualRating> AnnualRatings { get; } = [];
    public AssurePlusBenefitPricing? HealthOutcome { get; set; }
    public AssurePlusBenefitPricing? DentalOutcome { get; set; }
    public string HealthCommentary { get; set; } = "";
    public string DentalCommentary { get; set; } = "";
    public string Takeaway { get; set; } = "";
    public List<string> Warnings { get; } = [];
}

public sealed class AssurePlusAnnualRating
{
    public int Year { get; set; }
    public string Benefit { get; set; } = "";
    public string Rating { get; set; } = "";
    public string SourceReference { get; set; } = "";
    public bool Conflict { get; set; }
}

namespace RenewalReviewBuilder.App.Models;

/// <summary>Actual dental ASO activity from a carrier monthly billing statement.</summary>
public sealed class AsoBillingActivityRecord
{
    public DateOnly ServiceMonth { get; set; }
    public decimal PaidDentalClaims { get; set; }
    public decimal? AdministrationFee { get; set; }
    public decimal? Gst { get; set; }
    public decimal? InterestCreditCharge { get; set; }
    public decimal? TotalAsoCost { get; set; }
    public string SourcePath { get; set; } = "";
}

public sealed class AsoBillingFolderExtractionResult
{
    public List<AsoBillingActivityRecord> Records { get; } = [];
    public List<AsoFinancialActivityStatement> FinancialStatements { get; } = [];
    public List<string> Warnings { get; } = [];
}

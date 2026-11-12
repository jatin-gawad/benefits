namespace RenewalReviewBuilder.App.Models;

public sealed class FinancialProjectionPeriodInput
{
    public string Period { get; set; } = "";
    public FinancialCalculationInputSet CurrentInputs { get; set; } = new();
    public FinancialCalculationInputSet ProposedInputs { get; set; } = new();
}

public sealed class FinancialProjectionPeriodResult
{
    public string Period { get; set; } = "";
    public FinancialImpactComparison Impact { get; set; } = new();
}

public sealed class MultiPeriodFinancialProjection
{
    public List<FinancialProjectionPeriodResult> Periods { get; } = [];
    public decimal? TotalProposedAdministrativeCost { get; set; }
    public decimal? TotalProposedAsoExpense { get; set; }
}

namespace RenewalReviewBuilder.App.Models;

public enum FinancialArrangementInsightSeverity
{
    Information,
    Positive,
    Attention,
    Warning,
    Critical
}

public sealed class FinancialArrangementInsight
{
    public string InsightType { get; set; } = "";
    public FinancialArrangementInsightSeverity Severity { get; set; }
    public string Benefit { get; set; } = "";
    public string Title { get; set; } = "";
    public string Description { get; set; } = "";
    public decimal? CurrentValue { get; set; }
    public decimal? ProposedValue { get; set; }
    public string RuleCode { get; set; } = "";
    public bool RequiresAction { get; set; }
}

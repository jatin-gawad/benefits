namespace RenewalReviewBuilder.App.Models;

public enum FindingSeverity
{
    Info,
    Warning,
    Critical
}

public sealed class Finding
{
    public FindingSeverity Severity { get; set; }
    public string Title { get; set; } = "";
    public string Evidence { get; set; } = "";
    public string RecommendedAction { get; set; } = "";
    public string ClientExplanation { get; set; } = "";
    public string SourceReference { get; set; } = "";
}

public sealed class SourceCheck
{
    public string Area { get; set; } = "";
    public string Status { get; set; } = "";
    public string Detail { get; set; } = "";
}


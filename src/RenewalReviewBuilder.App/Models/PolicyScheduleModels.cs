namespace RenewalReviewBuilder.App.Models;

/// <summary>Source-shaped facts from a group policy schedule. These are not persistence entities.</summary>
public sealed class PolicyScheduleExtractionResult
{
    public string Carrier { get; set; } = "";
    public string DocumentType { get; set; } = DocumentTypes.GroupPolicySchedule;
    public string ExtractorVersion { get; set; } = "MANULIFE_POLICY_SCHEDULE_V1";
    public string SourcePath { get; set; } = "";
    public string PolicyNumber { get; set; } = "";
    public DateOnly? EffectiveDate { get; set; }
    public List<PolicyScheduleSection> Sections { get; } = [];
    public List<PolicyScheduleWarning> Warnings { get; } = [];
}

public sealed class PolicyScheduleSection
{
    public string Benefit { get; set; } = "";
    public string Heading { get; set; } = "";
    public int StartPage { get; set; }
    public List<string> RawClassifications { get; } = [];
    public List<string> ClassificationCodes { get; } = [];
    public List<PolicyScheduleField> Fields { get; } = [];
}

public sealed class PolicyScheduleField
{
    public string Label { get; set; } = "";
    public string Value { get; set; } = "";
    public string CanonicalHint { get; set; } = "";
    public int PageNumber { get; set; }
    public string SectionHeading { get; set; } = "";
    public string SourceExcerpt { get; set; } = "";
    public decimal Confidence { get; set; }
}

public sealed class PolicyScheduleWarning
{
    public string Code { get; set; } = "";
    public string Message { get; set; } = "";
    public int? PageNumber { get; set; }
    public string SourceExcerpt { get; set; } = "";
}

public sealed class PolicyScheduleMappingResult
{
    public List<PlanDesignRecord> Rows { get; } = [];
    public List<ValidationIssueRecord> ValidationIssues { get; } = [];
}

namespace RenewalReviewBuilder.App.Models;

public static class PlanDesignVerificationStatuses
{
    public const string Confirmed = "Confirmed";
    public const string EquivalentWording = "Equivalent Wording";
    public const string MissingFromSource = "Missing From Source";
    public const string SourceOnlyDetail = "Source Only Detail";
    public const string ProposedPlanChange = "Proposed Plan Change";
    public const string MaterialConflict = "Material Conflict";
    public const string EffectiveDateDifference = "Effective Date Difference";
    public const string ClassMappingRequired = "Class Mapping Required";
    public const string ManualReviewRequired = "Manual Review Required";
    public const string UnsupportedComparison = "Unsupported Comparison";
}

public static class PlanDesignStates
{
    public const string CurrentContract = "Current Contract";
    public const string FutureRenewal = "Future Renewal";
    public const string NegotiatedFuture = "Negotiated Future";
    public const string Historical = "Historical";
    public const string Approved = "Approved";
    public const string Superseded = "Superseded";
}

public sealed class PlanDesignEvidenceFact
{
    public PlanDesignCanonicalFact Fact { get; init; } = new();
    public string PlanState { get; init; } = PlanDesignStates.CurrentContract;
    public DateOnly? EffectiveDate { get; init; }
    public bool IsManualApproved { get; init; }
    public int SourceAuthority { get; init; }
}

public sealed class PlanDesignVerificationResult
{
    public PlanDesignEvidenceFact? Approved { get; init; }
    public PlanDesignEvidenceFact? Evidence { get; init; }
    public string Status { get; init; } = PlanDesignVerificationStatuses.UnsupportedComparison;
    public decimal Confidence { get; init; }
    public string Message { get; init; } = "";
}

public sealed class PlanDesignVerificationWorkspaceRow
{
    public string Benefit { get; init; } = "";
    public string ClassName { get; init; } = "";
    public string Field { get; init; } = "";
    public string ApprovedValue { get; init; } = "";
    public string EvidenceValue { get; init; } = "";
    public string Status { get; init; } = "";
    public string SourceDocument { get; init; } = "";
    public int? SourcePage { get; init; }
    public string SourceExcerpt { get; init; } = "";
}

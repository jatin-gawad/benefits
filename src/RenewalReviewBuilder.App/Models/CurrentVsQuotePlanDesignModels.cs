namespace RenewalReviewBuilder.App.Models;

public static class CurrentVsQuotePlanDesignStatuses
{
    public const string Same = "Same";
    public const string Different = "Different";
    public const string NeedsReview = "Needs Review";
    public const string CurrentOnly = "Current Only";
    public const string QuoteOnly = "Quote Only";
    public const string Unavailable = "Unavailable";
}

public sealed class CurrentVsQuotePlanDesignComparison
{
    public long ClientId { get; init; }
    public string ClientName { get; init; } = "";
    public string ClientNumber { get; init; } = "";
    public string CurrentCarrier { get; init; } = "";
    public string ProfileVersion { get; init; } = PlanDesignCanonicalProfileVersions.Current;
    public string ReadinessStatus { get; init; } = PlanDesignCanonicalReadinessStatuses.NotReady;
    public DateTime BuiltAtUtc { get; init; }
    public bool HasCurrentPlan { get; init; }
    public bool HasQuotePlan { get; init; }
    public bool CanCompare { get; init; }
    public List<CurrentVsQuotePlanDesignQuote> Quotes { get; } = [];
    public List<CurrentVsQuotePlanDesignRow> Rows { get; } = [];
    public List<PlanDesignCanonicalIssue> Issues { get; } = [];
}

public sealed class CurrentVsQuotePlanDesignQuote
{
    public string QuoteKey { get; init; } = "";
    public long? ProposalId { get; init; }
    public string Carrier { get; init; } = "";
    public string SourceFileName { get; init; } = "";
    public string DisplayName { get; init; } = "";
    public string Reference { get; init; } = "";
}

public sealed class CurrentVsQuotePlanDesignRow
{
    public string ComparisonKey { get; init; } = "";
    public string CanonicalBenefitCode { get; init; } = "";
    public string BenefitCategory { get; init; } = "";
    public string BenefitName { get; init; } = "";
    public string ClassName { get; init; } = "";
    public string ProvisionName { get; init; } = "";
    public string CurrentValue { get; init; } = "";
    public string CurrentAvailability { get; init; } = PlanDesignCanonicalAvailability.Missing;
    public string Status { get; init; } = CurrentVsQuotePlanDesignStatuses.Unavailable;
    public PlanDesignCanonicalEvidence? CurrentEvidence { get; init; }
    public List<CurrentVsQuotePlanDesignValue> QuoteValues { get; } = [];
}

public sealed class CurrentVsQuotePlanDesignValue
{
    public string QuoteKey { get; init; } = "";
    public long? ProposalId { get; init; }
    public string Carrier { get; init; } = "";
    public string QuoteDisplayName { get; init; } = "";
    public string Value { get; init; } = "";
    public string Status { get; init; } = CurrentVsQuotePlanDesignStatuses.Unavailable;
    public string Availability { get; init; } = PlanDesignCanonicalAvailability.Missing;
    public string ValidationStatus { get; init; } = "";
    public PlanDesignCanonicalEvidence? Evidence { get; init; }
}

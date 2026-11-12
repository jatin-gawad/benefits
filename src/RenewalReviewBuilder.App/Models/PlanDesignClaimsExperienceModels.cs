namespace RenewalReviewBuilder.App.Models;

public enum PlanDesignClaimsReadMode { Legacy, CanonicalWithLegacyFallback, CanonicalOnly, Compare }
public static class PlanDesignClaimsMatchStatuses
{
    public const string ExactMatch = "ExactMatch";
    public const string BenefitLevelFallback = "BenefitLevelFallback";
    public const string CategoryLevelFallback = "CategoryLevelFallback";
    public const string ServiceMappingRequired = "ServiceMappingRequired";
    public const string ClassMismatch = "ClassMismatch";
    public const string OutsidePlanEffectivePeriod = "OutsidePlanEffectivePeriod";
    public const string MultiplePlanVersionsMatch = "MultiplePlanVersionsMatch";
    public const string PlanDesignUnavailable = "PlanDesignUnavailable";
}

/// <summary>Plan Design projection assembled from canonical claims; it is not a second claims fact model.</summary>
public sealed class PlanDesignClaimsExperienceReadModel
{
    public long ClientId { get; init; }
    public long CanonicalClaimsReportId { get; init; }
    public long? PlanDesignVersionId { get; init; }
    public string BenefitCode { get; init; } = ""; public string BenefitName { get; init; } = "";
    public string CategoryCode { get; init; } = ""; public string CategoryName { get; init; } = "";
    public string ServiceCode { get; init; } = ""; public string ServiceName { get; init; } = "";
    public string ProvinceCode { get; init; } = ""; public CanonicalClaimsRelationshipType RelationshipType { get; init; }
    public string? DivisionCode { get; init; } public string? ClassCode { get; init; }
    public DateOnly? PeriodStart { get; init; } public DateOnly? PeriodEnd { get; init; }
    public long? UniqueClaimants { get; init; } public long? Occurrences { get; init; }
    public decimal? SubmittedAmount { get; init; } public decimal? EligibleAmount { get; init; } public decimal? PaidAmount { get; init; }
    public IReadOnlyDictionary<DateOnly, decimal> MonthlyPaidAmounts { get; init; } = new Dictionary<DateOnly, decimal>();
    public string ValidationStatus { get; init; } = ""; public string EvidenceStatus { get; init; } = "";
    public IReadOnlyCollection<long> SourceDocumentIds { get; init; } = []; public decimal? MappingConfidence { get; init; }
    public string MatchStatus { get; init; } = PlanDesignClaimsMatchStatuses.PlanDesignUnavailable;
    public string ConfidenceLevel { get; init; } = "Unavailable"; public string FallbackReason { get; init; } = "";
    public CanonicalClaimsUtilizationMetrics Metrics => CanonicalClaimsUtilizationMetricCalculator.Calculate(new CanonicalClaimsUtilizationFact { UniqueClaimants=UniqueClaimants, Occurrences=Occurrences, SubmittedAmount=SubmittedAmount, EligibleAmount=EligibleAmount, PaidAmount=PaidAmount });
}

public sealed class PlanDesignClaimsExperienceResult
{
    public PlanDesignClaimsReadMode ReadMode { get; init; }
    public bool UsedCanonical { get; init; }
    public string FallbackReason { get; init; } = "";
    public List<PlanDesignClaimsExperienceReadModel> Rows { get; } = [];
    public List<CanonicalClaimsParityResult> ParityResults { get; } = [];
}

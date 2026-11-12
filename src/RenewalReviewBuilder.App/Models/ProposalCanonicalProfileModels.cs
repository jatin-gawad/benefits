namespace RenewalReviewBuilder.App.Models;

public static class ProposalCanonicalProfileVersions
{
    public const string Current = "PROPOSAL_CANONICAL_PROFILE_V1";
}

public static class ProposalCanonicalReadinessStatuses
{
    public const string Ready = "Ready";
    public const string Partial = "Partial";
    public const string NeedsReview = "Needs Review";
    public const string NotReady = "Not Ready";
}

public static class ProposalCanonicalIssueSeverities
{
    public const string Information = "Information";
    public const string Warning = "Warning";
    public const string Error = "Error";
}

public sealed class ProposalCanonicalProfile
{
    public string ProfileVersion { get; set; } = ProposalCanonicalProfileVersions.Current;
    public long ClientId { get; set; }
    public string ClientName { get; set; } = "";
    public string ClientNumber { get; set; } = "";
    public DateTime BuiltAtUtc { get; set; } = DateTime.UtcNow;
    public List<ProposalCanonicalQuote> Proposals { get; } = [];
    public List<ProposalCanonicalSourceUsage> Sources { get; } = [];
    public List<ProposalCanonicalIssue> Issues { get; } = [];

    public bool IsReadOnly => true;
    public bool HasProposals => Proposals.Count > 0;
    public bool CanCompareCarriers => Proposals.Count(proposal => proposal.HasCoreFacts) >= 2;
    public bool CanCompareFinancials => Proposals.Count(proposal => proposal.HasFinancialFacts) >= 2;
    public bool CanComparePlanDesign => Proposals.Count(proposal => proposal.HasPlanDesignFacts) >= 2;
    public bool CanSupportProposalIntelligence => Proposals.Any(proposal => proposal.HasCoreFacts);

    public string ReadinessStatus
    {
        get
        {
            if (!HasProposals || Proposals.All(proposal => !proposal.HasCoreFacts))
            {
                return ProposalCanonicalReadinessStatuses.NotReady;
            }

            if (Issues.Any(issue => issue.Severity == ProposalCanonicalIssueSeverities.Error) ||
                Proposals.Any(proposal => proposal.ReadinessStatus == ProposalCanonicalReadinessStatuses.NeedsReview))
            {
                return ProposalCanonicalReadinessStatuses.NeedsReview;
            }

            if (Issues.Any(issue => issue.Severity == ProposalCanonicalIssueSeverities.Warning) ||
                Proposals.Any(proposal => proposal.ReadinessStatus != ProposalCanonicalReadinessStatuses.Ready))
            {
                return ProposalCanonicalReadinessStatuses.Partial;
            }

            return ProposalCanonicalReadinessStatuses.Ready;
        }
    }

    public string ReadinessMessage => ReadinessStatus switch
    {
        ProposalCanonicalReadinessStatuses.Ready => "Saved proposal facts, validation, and source evidence are ready for Proposal Intelligence.",
        ProposalCanonicalReadinessStatuses.NeedsReview => "Saved proposal data is available, but one or more validation or extraction issues require review.",
        ProposalCanonicalReadinessStatuses.Partial => "Saved proposal data is usable, with missing facts or validation noted in the profile issues.",
        _ => "No usable saved proposal facts were found for this client."
    };
}

public sealed class ProposalCanonicalQuote
{
    public long ProposalId { get; set; }
    public string StoredClientId { get; set; } = "";
    public long? CarrierId { get; set; }
    public string ClientName { get; set; } = "";
    public string CarrierKey { get; set; } = "";
    public string CarrierName { get; set; } = "";
    public string DetectedCarrier { get; set; } = "";
    public DateOnly? ProposalDate { get; set; }
    public string ProposalDateSourceText { get; set; } = "";
    public DateOnly? EffectiveDate { get; set; }
    public string EffectiveDateSourceText { get; set; } = "";
    public string SourceFilePath { get; set; } = "";
    public string SourceFileName { get; set; } = "";
    public string SourceFileHash { get; set; } = "";
    public List<string> SourceFiles { get; } = [];
    public string DetectionConfidence { get; set; } = "";
    public decimal? ExtractionConfidence { get; set; }
    public string ExtractionMethod { get; set; } = "Proposal Engine";
    public string Status { get; set; } = "";
    public int ReprocessCount { get; set; }
    public DateTime? CreatedAt { get; set; }
    public DateTime? UpdatedAt { get; set; }
    public ProposalCanonicalEvidence Evidence { get; set; } = new();

    public List<ProposalCanonicalBenefitDefinition> Benefits { get; } = [];
    public List<ProposalCanonicalBenefitLine> BenefitLines { get; } = [];
    public List<ProposalCanonicalRate> Rates { get; } = [];
    public List<ProposalCanonicalPremium> Premiums { get; } = [];
    public List<ProposalCanonicalPlanDesignFact> PlanDesignFacts { get; } = [];
    public List<ProposalCanonicalPlanDesignFact> ValidationFacts { get; } = [];
    public List<ProposalCanonicalDemographic> Demographics { get; } = [];
    public List<ProposalCanonicalUnderwritingAssumption> UnderwritingAssumptions { get; } = [];
    public List<ProposalCanonicalPoolingRetention> PoolingRetention { get; } = [];
    public List<ProposalCanonicalException> Exceptions { get; } = [];
    public List<ProposalCanonicalAuditRecord> AuditTrail { get; } = [];

    public bool HasCoreFacts => BenefitLines.Count > 0 || Rates.Count > 0 || Premiums.Count > 0 || PlanDesignFacts.Count > 0;
    public bool HasFinancialFacts => Rates.Count > 0 || Premiums.Count > 0 || BenefitLines.Any(row => row.MonthlyPremium.HasValue || row.AnnualPremium.HasValue || row.Rate.HasValue);
    public bool HasPlanDesignFacts => PlanDesignFacts.Count > 0;
    public bool HasBookletValidation => ValidationFacts.Count > 0;
    public bool HasSourceEvidence => !string.IsNullOrWhiteSpace(SourceFileHash) ||
                                     !string.IsNullOrWhiteSpace(SourceFileName) ||
                                     AllEvidence().Any(evidence => !string.IsNullOrWhiteSpace(evidence.SourceFileName));
    public int ValidationMismatchCount => ValidationFacts.Count(fact => fact.IsMismatch);
    public int ReviewIssueCount => Exceptions.Count(row => row.NeedsReview) + ValidationMismatchCount;

    public decimal? TotalMonthlyPremium => SumWithFallback(
        Premiums.Select(row => row.MonthlyPremium),
        BenefitLines.Select(row => row.MonthlyPremium),
        Rates.Select(row => row.MonthlyPremium));

    public decimal? TotalAnnualPremium => SumWithFallback(
        Premiums.Select(row => row.AnnualPremium),
        BenefitLines.Select(row => row.AnnualPremium),
        Rates.Select(row => row.AnnualPremium));

    public decimal? TotalEmployerCost => Sum(Premiums.Select(row => row.EmployerCost));
    public decimal? TotalEmployeeCost => Sum(Premiums.Select(row => row.EmployeeCost));
    public decimal? TotalCost => Sum(Premiums.Select(row => row.TotalCost));

    public string ReadinessStatus
    {
        get
        {
            if (!HasCoreFacts)
            {
                return ProposalCanonicalReadinessStatuses.NotReady;
            }

            if (ValidationMismatchCount > 0 || Exceptions.Any(row => row.NeedsReview))
            {
                return ProposalCanonicalReadinessStatuses.NeedsReview;
            }

            if (!HasFinancialFacts || !HasPlanDesignFacts || !HasBookletValidation || !HasSourceEvidence)
            {
                return ProposalCanonicalReadinessStatuses.Partial;
            }

            return ProposalCanonicalReadinessStatuses.Ready;
        }
    }

    private IEnumerable<ProposalCanonicalEvidence> AllEvidence()
    {
        yield return Evidence;
        foreach (var evidence in Benefits.Select(row => row.Evidence)
                     .Concat(BenefitLines.Select(row => row.Evidence))
                     .Concat(Rates.Select(row => row.Evidence))
                     .Concat(Premiums.Select(row => row.Evidence))
                     .Concat(PlanDesignFacts.Select(row => row.Evidence))
                     .Concat(ValidationFacts.Select(row => row.Evidence))
                     .Concat(Demographics.Select(row => row.Evidence))
                     .Concat(UnderwritingAssumptions.Select(row => row.Evidence))
                     .Concat(PoolingRetention.Select(row => row.Evidence))
                     .Concat(Exceptions.Select(row => row.Evidence)))
        {
            yield return evidence;
        }
    }

    private static decimal? SumWithFallback(params IEnumerable<decimal?>[] candidates)
    {
        foreach (var candidate in candidates)
        {
            var result = Sum(candidate);
            if (result.HasValue)
            {
                return result;
            }
        }

        return null;
    }

    private static decimal? Sum(IEnumerable<decimal?> values)
    {
        var available = values.Where(value => value.HasValue).Select(value => value!.Value).ToList();
        return available.Count == 0 ? null : available.Sum();
    }
}

public sealed class ProposalCanonicalBenefitDefinition
{
    public string BenefitKey { get; set; } = "";
    public string BenefitName { get; set; } = "";
    public string NormalizedBenefitName { get; set; } = "";
    public string DisplayName { get; set; } = "";
    public string BenefitCategory { get; set; } = "";
    public int? SortOrder { get; set; }
    public bool IsStandard { get; set; }
    public bool IsKnown { get; set; }
    public ProposalCanonicalEvidence Evidence { get; set; } = new();
}

public sealed class ProposalCanonicalBenefitLine
{
    public string BenefitLine { get; set; } = "";
    public string ClassName { get; set; } = "";
    public string Tier { get; set; } = "";
    public decimal? Lives { get; set; }
    public decimal? Volume { get; set; }
    public decimal? Rate { get; set; }
    public decimal? MonthlyPremium { get; set; }
    public decimal? AnnualPremium { get; set; }
    public ProposalCanonicalEvidence Evidence { get; set; } = new();
}

public sealed class ProposalCanonicalRate
{
    public string BenefitLine { get; set; } = "";
    public string ClassName { get; set; } = "";
    public string Tier { get; set; } = "";
    public string AgeBand { get; set; } = "";
    public string Gender { get; set; } = "";
    public string Province { get; set; } = "";
    public decimal? Lives { get; set; }
    public decimal? Volume { get; set; }
    public decimal? Rate { get; set; }
    public string RateBasis { get; set; } = "";
    public decimal? MonthlyPremium { get; set; }
    public decimal? AnnualPremium { get; set; }
    public ProposalCanonicalEvidence Evidence { get; set; } = new();
}

public sealed class ProposalCanonicalPremium
{
    public string BenefitLine { get; set; } = "";
    public string ClassName { get; set; } = "";
    public string Tier { get; set; } = "";
    public decimal? MonthlyPremium { get; set; }
    public decimal? AnnualPremium { get; set; }
    public decimal? EmployerCost { get; set; }
    public decimal? EmployeeCost { get; set; }
    public decimal? TotalCost { get; set; }
    public ProposalCanonicalEvidence Evidence { get; set; } = new();
}

public sealed class ProposalCanonicalPlanDesignFact
{
    public string BenefitLine { get; set; } = "";
    public string ProvisionName { get; set; } = "";
    public string ProposalValue { get; set; } = "";
    public string BookletValue { get; set; } = "";
    public string Variance { get; set; } = "";
    public string ValidationStatus { get; set; } = "";
    public string BookletCarrier { get; set; } = "";
    public List<string> BookletClassNames { get; } = [];
    public string ComparisonMethod { get; set; } = "";
    public ProposalCanonicalEvidence Evidence { get; set; } = new();
    public ProposalCanonicalEvidence BookletEvidence { get; set; } = new();

    public bool IsMismatch
    {
        get
        {
            if (!string.IsNullOrWhiteSpace(ValidationStatus) &&
                !ValidationStatus.Equals("valid", StringComparison.OrdinalIgnoreCase) &&
                !ValidationStatus.Equals("match", StringComparison.OrdinalIgnoreCase) &&
                !ValidationStatus.Equals("matched", StringComparison.OrdinalIgnoreCase) &&
                !ValidationStatus.Equals("pass", StringComparison.OrdinalIgnoreCase) &&
                !ValidationStatus.Equals("passes", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (!string.IsNullOrWhiteSpace(Variance) &&
                !Variance.Equals("none", StringComparison.OrdinalIgnoreCase) &&
                !Variance.Equals("match", StringComparison.OrdinalIgnoreCase) &&
                !Variance.Contains("no variance", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return !string.IsNullOrWhiteSpace(ProposalValue) &&
                   !string.IsNullOrWhiteSpace(BookletValue) &&
                   !Normalize(ProposalValue).Equals(Normalize(BookletValue), StringComparison.OrdinalIgnoreCase);
        }
    }

    private static string Normalize(string value) =>
        new(value.Where(char.IsLetterOrDigit).Select(char.ToLowerInvariant).ToArray());
}

public sealed class ProposalCanonicalDemographic
{
    public string BenefitLine { get; set; } = "";
    public string ClassName { get; set; } = "";
    public string Tier { get; set; } = "";
    public decimal? Lives { get; set; }
    public decimal? Volume { get; set; }
    public ProposalCanonicalEvidence Evidence { get; set; } = new();
}

public sealed class ProposalCanonicalUnderwritingAssumption
{
    public string Name { get; set; } = "";
    public string Value { get; set; } = "";
    public ProposalCanonicalEvidence Evidence { get; set; } = new();
}

public sealed class ProposalCanonicalPoolingRetention
{
    public string BenefitLine { get; set; } = "";
    public string PoolingLevel { get; set; } = "";
    public decimal? PoolingCharge { get; set; }
    public decimal? Commission { get; set; }
    public decimal? Retention { get; set; }
    public decimal? TargetLossRatio { get; set; }
    public string RateGuarantee { get; set; } = "";
    public string PooledBenefits { get; set; } = "";
    public string ExperienceRatedBenefits { get; set; } = "";
    public string Ibnr { get; set; } = "";
    public string RateCaps { get; set; } = "";
    public string FirstRenewalCap { get; set; } = "";
    public string SecondRenewalCap { get; set; } = "";
    public ProposalCanonicalEvidence Evidence { get; set; } = new();
}

public sealed class ProposalCanonicalException
{
    public string ExceptionType { get; set; } = "";
    public string FieldName { get; set; } = "";
    public string Message { get; set; } = "";
    public string BenefitLine { get; set; } = "";
    public string ProvisionName { get; set; } = "";
    public ProposalCanonicalEvidence Evidence { get; set; } = new();

    public bool NeedsReview => ExceptionType.Contains("review", StringComparison.OrdinalIgnoreCase) ||
                               ExceptionType.Contains("missing", StringComparison.OrdinalIgnoreCase) ||
                               Evidence.Status.Contains("review", StringComparison.OrdinalIgnoreCase) ||
                               Evidence.Status.Contains("missing", StringComparison.OrdinalIgnoreCase);
}

public sealed class ProposalCanonicalAuditRecord
{
    public string Action { get; set; } = "";
    public string TableName { get; set; } = "";
    public int? RowsLoaded { get; set; }
    public string Message { get; set; } = "";
    public ProposalCanonicalEvidence Evidence { get; set; } = new();
}

public sealed class ProposalCanonicalEvidence
{
    public string SourceTable { get; set; } = "";
    public long? SourceId { get; set; }
    public long? ProposalId { get; set; }
    public string SourceFileName { get; set; } = "";
    public string SourceFileHash { get; set; } = "";
    public int? SourcePage { get; set; }
    public string SourceSheet { get; set; } = "";
    public decimal? ConfidenceScore { get; set; }
    public string Status { get; set; } = "";
    public string SourceType { get; set; } = "";
    public string SourceLocation { get; set; } = "";
    public string ExtractionMethod { get; set; } = "Proposal Engine";
    public string RawPayloadJson { get; set; } = "";
    public DateTime? CreatedAt { get; set; }
    public DateTime? UpdatedAt { get; set; }
}

public sealed class ProposalCanonicalSourceUsage
{
    public string TableName { get; set; } = "";
    public bool Exists { get; set; }
    public int RowsFound { get; set; }
    public bool Used { get; set; }
    public string Notes { get; set; } = "";
}

public sealed class ProposalCanonicalIssue
{
    public string Code { get; set; } = "";
    public string Severity { get; set; } = ProposalCanonicalIssueSeverities.Warning;
    public string Message { get; set; } = "";
    public long? ProposalId { get; set; }
    public string Carrier { get; set; } = "";
    public string SourceTable { get; set; } = "";
    public long? SourceId { get; set; }
}

namespace RenewalReviewBuilder.App.Models;

public static class PlanDesignCanonicalProfileVersions
{
    public const string Current = "PLAN_DESIGN_CANONICAL_PROFILE_V1";
}

public static class PlanDesignCanonicalScopes
{
    public const string CurrentPlan = "Current Plan";
    public const string Quote = "Quote";
    public const string ProposalValidation = "Proposal Validation";
}

public static class PlanDesignCanonicalAvailability
{
    public const string Available = "Available";
    public const string Partial = "Partial";
    public const string Missing = "Missing";
    public const string NeedsReview = "Needs Review";
}

public static class PlanDesignCanonicalReadinessStatuses
{
    public const string Ready = "Ready";
    public const string Partial = "Partial";
    public const string NotReady = "Not Ready";
    public const string NeedsReview = "Needs Review";
}

public static class PlanDesignCanonicalIssueSeverities
{
    public const string Information = "Information";
    public const string Warning = "Warning";
    public const string Error = "Error";
}

public static class CanonicalSourceLineageValueGuardrailVersions
{
    public const string Current = "CANONICAL_SOURCE_LINEAGE_VALUE_GUARDRAILS_V1";
}

public static class CanonicalSourceLineageStatuses
{
    public const string Verified = "Verified";
    public const string Unlinked = "Unlinked";
    public const string Missing = "Missing";
    public const string NeedsReview = "Needs Review";
}

public static class CanonicalValueGuardrailStatuses
{
    public const string Valid = "Valid";
    public const string NotApplicable = "Not Applicable";
    public const string PreservedForReview = "Preserved for Review";
    public const string Rejected = "Rejected";
}

public static class CanonicalSourceLineageValueGuardrailIssueCodes
{
    public const string SourceLineageConflict = "source_lineage_conflict";
    public const string SourceLineageUnlinked = "source_lineage_unlinked";
    public const string PercentageOutOfRange = "plan_design_percentage_out_of_range";
    public const string MultiplePercentagesPreserved = "plan_design_multiple_percentages_preserved";
}

public sealed class PlanDesignCanonicalProfile
{
    public string ProfileVersion { get; set; } = PlanDesignCanonicalProfileVersions.Current;
    public long ClientId { get; set; }
    public string ClientName { get; set; } = "";
    public string ClientNumber { get; set; } = "";
    public string Carrier { get; set; } = "";
    public DateTime BuiltAtUtc { get; set; } = DateTime.UtcNow;
    public List<PlanDesignCanonicalFact> Facts { get; } = [];
    public List<PlanDesignCanonicalSourceUsage> Sources { get; } = [];
    public List<PlanDesignCanonicalIssue> Issues { get; } = [];
    public List<PlanDesignFactorRecord> FactorReferences { get; } = [];
    public List<string> ContextBenefitLines { get; } = [];

    public bool IsReadOnly => true;

    public IReadOnlyList<PlanDesignCanonicalFact> CurrentPlanFacts =>
        Facts.Where(fact => fact.Scope == PlanDesignCanonicalScopes.CurrentPlan).ToList();

    public IReadOnlyList<PlanDesignCanonicalFact> QuoteFacts =>
        Facts.Where(fact => fact.Scope == PlanDesignCanonicalScopes.Quote).ToList();

    public IReadOnlyList<PlanDesignCanonicalFact> ProposalValidationFacts =>
        Facts.Where(fact => fact.Scope == PlanDesignCanonicalScopes.ProposalValidation).ToList();

    public bool HasCurrentPlan => CurrentPlanFacts.Count > 0;
    public bool HasQuotePlan => QuoteFacts.Count > 0;
    public bool HasProposalValidation => ProposalValidationFacts.Count > 0;
    public bool CanCompareCurrentVsQuote => HasCurrentPlan && HasQuotePlan;
    public bool CanSupportRenewalIntelligence => HasCurrentPlan;
    public bool CanSupportProposalIntelligence => CanCompareCurrentVsQuote;
    public bool CanSupportPlanDesignIntelligence => HasCurrentPlan;

    public string ReadinessStatus
    {
        get
        {
            if (!HasCurrentPlan)
            {
                return PlanDesignCanonicalReadinessStatuses.NotReady;
            }

            if (Issues.Any(issue => issue.Code is
                    "proposal_plan_design_mismatch" or
                    CanonicalSourceLineageValueGuardrailIssueCodes.SourceLineageConflict or
                    CanonicalSourceLineageValueGuardrailIssueCodes.SourceLineageUnlinked or
                    CanonicalSourceLineageValueGuardrailIssueCodes.PercentageOutOfRange or
                    CanonicalSourceLineageValueGuardrailIssueCodes.MultiplePercentagesPreserved) ||
                Facts.Any(fact => fact.Availability == PlanDesignCanonicalAvailability.NeedsReview))
            {
                return PlanDesignCanonicalReadinessStatuses.NeedsReview;
            }

            if (!HasQuotePlan ||
                !HasProposalValidation ||
                Issues.Any(issue => issue.Severity is PlanDesignCanonicalIssueSeverities.Warning or PlanDesignCanonicalIssueSeverities.Error) ||
                Facts.Any(fact => fact.Availability == PlanDesignCanonicalAvailability.Partial))
            {
                return PlanDesignCanonicalReadinessStatuses.Partial;
            }

            return PlanDesignCanonicalReadinessStatuses.Ready;
        }
    }
}

public sealed class PlanDesignCanonicalFact
{
    public string GuardrailVersion { get; set; } = CanonicalSourceLineageValueGuardrailVersions.Current;
    public string Scope { get; set; } = "";
    public long ClientId { get; set; }
    public long? ProposalId { get; set; }
    public string Carrier { get; set; } = "";
    public string CanonicalBenefitCode { get; set; } = "";
    public string NormalizedBenefitKey { get; set; } = "";
    public string BenefitCategory { get; set; } = "";
    public string BenefitName { get; set; } = "";
    public string ClassName { get; set; } = "";
    public string FeatureName { get; set; } = "";
    public string FeatureValue { get; set; } = "";
    public string BookletValue { get; set; } = "";
    public string Variance { get; set; } = "";
    public string ValidationStatus { get; set; } = "";
    public string Status { get; set; } = "";
    public decimal? ReimbursementPercent { get; set; }
    public string Coinsurance { get; set; } = "";
    public string Deductible { get; set; } = "";
    public string Maximum { get; set; } = "";
    public string Frequency { get; set; } = "";
    public string WaitingPeriod { get; set; } = "";
    public string TerminationAge { get; set; } = "";
    public string ComparisonKey { get; set; } = "";
    public string Availability { get; set; } = PlanDesignCanonicalAvailability.Missing;

    public string SourceTable { get; set; } = "";
    public string SourceId { get; set; } = "";
    public long? SourceDocumentId { get; set; }
    public string SourceDocument { get; set; } = "";
    public string SourceDocumentType { get; set; } = "";
    public string SourceType { get; set; } = "";
    public string SourceCarrier { get; set; } = "";
    public string SourceFileName { get; set; } = "";
    public string TrackedSourceFileName { get; set; } = "";
    public string SourceLineageStatus { get; set; } = CanonicalSourceLineageStatuses.Unlinked;
    public string SourceLineageMessage { get; set; } = "";
    public string ValueGuardrailStatus { get; set; } = CanonicalValueGuardrailStatuses.NotApplicable;
    public string ValueGuardrailMessage { get; set; } = "";
    public string SourceFileHash { get; set; } = "";
    public string SourceSheet { get; set; } = "";
    public int? SourcePage { get; set; }
    public string SourceText { get; set; } = "";
    public string SourceApiEndpoint { get; set; } = "";
    public long? ImportBatchId { get; set; }
    public DateTime? ImportedAt { get; set; }
    public string ImportedBy { get; set; } = "";
    public decimal? ConfidenceScore { get; set; }
    public string ExtractionMethod { get; set; } = "";
    public string RawSourceReference { get; set; } = "";
    public DateTime? DateLoaded { get; set; }

    public PlanDesignCanonicalEvidence Evidence => new()
    {
        SourceTable = SourceTable,
        SourceId = SourceId,
        SourceDocumentId = SourceDocumentId,
        SourceDocument = SourceDocument,
        SourceDocumentType = SourceDocumentType,
        SourceType = SourceType,
        SourceCarrier = SourceCarrier,
        SourceFileName = SourceFileName,
        TrackedSourceFileName = TrackedSourceFileName,
        SourceLineageStatus = SourceLineageStatus,
        SourceLineageMessage = SourceLineageMessage,
        ValueGuardrailStatus = ValueGuardrailStatus,
        ValueGuardrailMessage = ValueGuardrailMessage,
        GuardrailVersion = GuardrailVersion,
        SourceFileHash = SourceFileHash,
        SourceSheet = SourceSheet,
        SourcePage = SourcePage,
        SourceText = SourceText,
        SourceApiEndpoint = SourceApiEndpoint,
        ImportBatchId = ImportBatchId,
        ImportedAt = ImportedAt,
        ImportedBy = ImportedBy,
        ConfidenceScore = ConfidenceScore,
        ExtractionMethod = ExtractionMethod,
        RawSourceReference = RawSourceReference
    };
}

public sealed class PlanDesignCanonicalEvidence
{
    public string GuardrailVersion { get; set; } = CanonicalSourceLineageValueGuardrailVersions.Current;
    public string SourceTable { get; set; } = "";
    public string SourceId { get; set; } = "";
    public long? SourceDocumentId { get; set; }
    public string SourceDocument { get; set; } = "";
    public string SourceDocumentType { get; set; } = "";
    public string SourceType { get; set; } = "";
    public string SourceCarrier { get; set; } = "";
    public string SourceFileName { get; set; } = "";
    public string TrackedSourceFileName { get; set; } = "";
    public string SourceLineageStatus { get; set; } = CanonicalSourceLineageStatuses.Unlinked;
    public string SourceLineageMessage { get; set; } = "";
    public string ValueGuardrailStatus { get; set; } = CanonicalValueGuardrailStatuses.NotApplicable;
    public string ValueGuardrailMessage { get; set; } = "";
    public string SourceFileHash { get; set; } = "";
    public string SourceSheet { get; set; } = "";
    public int? SourcePage { get; set; }
    public string SourceText { get; set; } = "";
    public string SourceApiEndpoint { get; set; } = "";
    public long? ImportBatchId { get; set; }
    public DateTime? ImportedAt { get; set; }
    public string ImportedBy { get; set; } = "";
    public decimal? ConfidenceScore { get; set; }
    public string ExtractionMethod { get; set; } = "";
    public string RawSourceReference { get; set; } = "";
}

public sealed class PlanDesignCanonicalSourceUsage
{
    public string TableName { get; set; } = "";
    public bool Exists { get; set; }
    public int RowsFound { get; set; }
    public bool Used { get; set; }
    public string Notes { get; set; } = "";
}

public sealed class PlanDesignCanonicalIssue
{
    public string Code { get; set; } = "";
    public string Severity { get; set; } = PlanDesignCanonicalIssueSeverities.Warning;
    public string Message { get; set; } = "";
    public string Benefit { get; set; } = "";
    public string Provision { get; set; } = "";
    public string SourceTable { get; set; } = "";
    public string SourceId { get; set; } = "";
}

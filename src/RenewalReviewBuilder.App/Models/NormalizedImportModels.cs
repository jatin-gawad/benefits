namespace RenewalReviewBuilder.App.Models;

public static class SourceTypes
{
    public const string Pdf = "pdf";
    public const string Excel = "excel";
    public const string Api = "api";
    public const string Manual = "manual";
}

public static class ImportStatuses
{
    public const string Created = "Created";
    public const string Extracted = "Extracted";
    public const string Normalized = "Normalized";
    public const string Validated = "Validated";
    public const string Completed = "Completed";
    public const string NeedsReview = "Needs Review";
    public const string Failed = "Failed";
}

public static class ValidationIssueCodes
{
    public const string MissingClientName = "missing_client_name";
    public const string MissingCarrier = "missing_carrier";
    public const string MissingRenewalEffectiveDate = "missing_renewal_effective_date";
    public const string MissingPremium = "missing_premium";
    public const string MissingClaims = "missing_claims";
    public const string MissingRates = "missing_rates";
    public const string MissingClass = "missing_class";
    public const string ClaimsPeriodMismatch = "claims_period_mismatch";
    public const string RenewalPeriodMismatch = "renewal_period_mismatch";
    public const string BenefitNotMapped = "benefit_not_mapped";
    public const string ProvinceMissing = "province_missing";
    public const string OptionClassMissing = "option_class_missing";
    public const string PlanDesignPercentageOutOfRange = CanonicalSourceLineageValueGuardrailIssueCodes.PercentageOutOfRange;
    public const string PlanDesignMultiplePercentagesPreserved = CanonicalSourceLineageValueGuardrailIssueCodes.MultiplePercentagesPreserved;
    public const string PlanDesignProvisionIdentityMissing = "plan_design_provision_identity_missing";
    public const string PlanDesignDuplicateProvisionIdentity = "plan_design_duplicate_provision_identity";
    public const string PolicyScheduleClassificationReviewRequired = "policy_schedule_classification_review_required";
    public const string PolicyScheduleProcessorNotRegistered = "policy_schedule_processor_not_registered";
    public const string PolicyScheduleExtractionNotImplemented = "policy_schedule_extraction_not_implemented";
    public const string PolicyScheduleNoBenefitSections = "policy_schedule_no_benefit_sections";
    public const string PolicyScheduleClassMappingRequired = "policy_schedule_class_mapping_required";
    public const string PolicyScheduleUnsupportedTerm = "policy_schedule_unsupported_term";
    public const string PolicySchedulePartialExtraction = "policy_schedule_partial_extraction";
}

public sealed class NormalizedSourceMetadata
{
    public string SourceType { get; set; } = "";
    public string SourceCarrier { get; set; } = "";
    public string SourceFileName { get; set; } = "";
    public string SourceDocumentId { get; set; } = "";
    public string SourceApiEndpoint { get; set; } = "";
    public long? ImportBatchId { get; set; }
    public DateTime ImportedAt { get; set; } = DateTime.Now;
    public string ImportedBy { get; set; } = Environment.UserName;
    public decimal? ConfidenceScore { get; set; }
    public string ExtractionMethod { get; set; } = "";
    public string RawSourceReference { get; set; } = "";

    public static NormalizedSourceMetadata ForFile(
        string sourceType,
        string carrier,
        string filePath,
        string extractionMethod,
        decimal? confidenceScore = null)
    {
        return new NormalizedSourceMetadata
        {
            SourceType = sourceType,
            SourceCarrier = carrier,
            SourceFileName = string.IsNullOrWhiteSpace(filePath) ? "" : Path.GetFileName(filePath),
            SourceDocumentId = "",
            SourceApiEndpoint = "",
            ImportedAt = DateTime.Now,
            ImportedBy = Environment.UserName,
            ConfidenceScore = confidenceScore,
            ExtractionMethod = extractionMethod,
            RawSourceReference = filePath
        };
    }

    public static NormalizedSourceMetadata ForApi(string carrier, string endpoint)
    {
        return new NormalizedSourceMetadata
        {
            SourceType = SourceTypes.Api,
            SourceCarrier = carrier,
            SourceApiEndpoint = endpoint,
            ImportedAt = DateTime.Now,
            ImportedBy = Environment.UserName,
            ExtractionMethod = "Carrier API",
            RawSourceReference = endpoint
        };
    }

    public NormalizedSourceMetadata Copy()
    {
        return new NormalizedSourceMetadata
        {
            SourceType = SourceType,
            SourceCarrier = SourceCarrier,
            SourceFileName = SourceFileName,
            SourceDocumentId = SourceDocumentId,
            SourceApiEndpoint = SourceApiEndpoint,
            ImportBatchId = ImportBatchId,
            ImportedAt = ImportedAt,
            ImportedBy = ImportedBy,
            ConfidenceScore = ConfidenceScore,
            ExtractionMethod = ExtractionMethod,
            RawSourceReference = RawSourceReference
        };
    }
}

public abstract class NormalizedRecord
{
    public NormalizedSourceMetadata Source { get; set; } = new();
}

public sealed class NormalizedImportPayload
{
    public string DocumentType { get; set; } = DocumentTypes.Other;
    public NormalizedSourceMetadata Source { get; set; } = new();
    public NormalizedClient? Client { get; set; }
    public NormalizedCarrier? Carrier { get; set; }
    public NormalizedPolicy? Policy { get; set; }
    public List<NormalizedRenewal> Renewals { get; } = [];
    public List<NormalizedRenewalVersion> RenewalVersions { get; } = [];
    public List<NormalizedBenefit> Benefits { get; } = [];
    public List<NormalizedBenefitClass> BenefitClasses { get; } = [];
    public List<NormalizedBenefitOption> BenefitOptions { get; } = [];
    public List<NormalizedRate> Rates { get; } = [];
    public List<NormalizedPremium> Premiums { get; } = [];
    public List<NormalizedClaim> Claims { get; } = [];
    public List<NormalizedClaimsExperience> ClaimsExperience { get; } = [];
    public List<NormalizedBillingStatement> BillingStatements { get; } = [];
    public List<NormalizedPooling> Pooling { get; } = [];
    public List<NormalizedCommission> Commissions { get; } = [];
    public List<NormalizedPlanDesign> PlanDesign { get; } = [];
    public List<NormalizedQuote> Quotes { get; } = [];
    public List<NormalizedProposal> Proposals { get; } = [];
    public List<NormalizedFlexibleBenefitPlan> FlexibleBenefitPlans { get; } = [];
    public List<ValidationIssueRecord> ValidationIssues { get; } = [];

    public void ApplySourceToRecords()
    {
        Apply(Client);
        Apply(Carrier);
        Apply(Policy);
        foreach (var record in AllRecords())
        {
            Apply(record);
        }
    }

    private IEnumerable<NormalizedRecord> AllRecords()
    {
        return Renewals
            .Cast<NormalizedRecord>()
            .Concat(RenewalVersions)
            .Concat(Benefits)
            .Concat(BenefitClasses)
            .Concat(BenefitOptions)
            .Concat(Rates)
            .Concat(Premiums)
            .Concat(Claims)
            .Concat(ClaimsExperience)
            .Concat(BillingStatements)
            .Concat(Pooling)
            .Concat(Commissions)
            .Concat(PlanDesign)
            .Concat(Quotes)
            .Concat(Proposals)
            .Concat(FlexibleBenefitPlans);
    }

    private void Apply(NormalizedRecord? record)
    {
        if (record is not null)
        {
            record.Source = Source.Copy();
        }
    }
}

public sealed class NormalizedClient : NormalizedRecord
{
    public string ClientName { get; set; } = "";
    public string ClientNumber { get; set; } = "";
    public string Province { get; set; } = "";
    public string Advisor { get; set; } = "";
}

public sealed class NormalizedCarrier : NormalizedRecord
{
    public string CarrierName { get; set; } = "";
}

public sealed class NormalizedPolicy : NormalizedRecord
{
    public string PolicyNumber { get; set; } = "";
    public string ClientName { get; set; } = "";
    public string CarrierName { get; set; } = "";
    public DateOnly? EffectiveDate { get; set; }
    public DateOnly? RenewalDate { get; set; }
}

public sealed class NormalizedRenewal : NormalizedRecord
{
    public string CarrierName { get; set; } = "";
    public string BenefitLine { get; set; } = "";
    public DateOnly? EffectiveDate { get; set; }
    public int? RenewalYear { get; set; }
    public decimal? CurrentPremium { get; set; }
    public decimal? RenewalPremium { get; set; }
    public decimal? Claims { get; set; }
    public decimal? LossRatio { get; set; }
    public int? Lives { get; set; }
}

public sealed class NormalizedRenewalVersion : NormalizedRecord
{
    public string VersionName { get; set; } = "";
    public string Status { get; set; } = "";
    public DateOnly? EffectiveDate { get; set; }
    public decimal? RequestedIncreasePercent { get; set; }
    public decimal? NegotiatedIncreasePercent { get; set; }
    public decimal? FinalAcceptedIncreasePercent { get; set; }
    public NegotiationStatus? NegotiationStatus { get; set; }
    public string? NegotiationStage { get; set; }
    public DateOnly? NegotiationCompletedDate { get; set; }
    public decimal? NegotiatedSavingsAnnual { get; set; }
    public decimal? NegotiatedSavingsMonthly { get; set; }
    public NegotiatedSavingsType? NegotiatedSavingsType { get; set; }
    public string? NegotiatedBy { get; set; }
    public decimal? IndustryBenchmarkPercent { get; set; }
    public string? IndustryBenchmarkSource { get; set; }
    public string? IndustryBenchmarkPopulation { get; set; }
    public int? IndustryBenchmarkYear { get; set; }
    public decimal? IndustryBenchmarkConfidence { get; set; }
    public RenewalRecommendationType? RenewalRecommendationType { get; set; }
    public string? RenewalDecisionReason { get; set; }
    public string? RenewalStrategySummary { get; set; }
    public RenewalRiskLevel? RenewalRiskLevel { get; set; }
    public string? RenewalComplexity { get; set; }
    public string? RenewalResult { get; set; }
    public DateOnly? RenewalCompletedDate { get; set; }
    public string? RenewalApprovedBy { get; set; }
    public DateOnly? RenewalAcceptedDate { get; set; }
    public int? RenewalVersionNumber { get; set; }
    public long? RenewalScenarioId { get; set; }
    public long? PreviousRenewalScenarioId { get; set; }
    public ScenarioType? ScenarioType { get; set; }
    public long? HistoryEventId { get; set; }
    public string? RelationshipMilestoneId { get; set; }
    public string? EvolutionGroupId { get; set; }
}

public sealed class NormalizedBenefit : NormalizedRecord
{
    public string BenefitLine { get; set; } = "";
    public string BenefitCategory { get; set; } = "";
}

public sealed class NormalizedBenefitClass : NormalizedRecord
{
    public string ClassName { get; set; } = "";
    public string ClassCode { get; set; } = "";
}

public sealed class NormalizedBenefitOption : NormalizedRecord
{
    public string OptionName { get; set; } = "";
    public string ClassCode { get; set; } = "";
    public string BenefitLine { get; set; } = "";
}

public sealed class NormalizedRate : NormalizedRecord
{
    public string BenefitLine { get; set; } = "";
    public string ClassName { get; set; } = "";
    public decimal? CurrentRate { get; set; }
    public decimal? ProposedRate { get; set; }
    public decimal? NegotiatedRate { get; set; }
    public string RateBasis { get; set; } = "";
    public DateOnly? EffectiveDate { get; set; }
}

public sealed class NormalizedPremium : NormalizedRecord
{
    public string BenefitLine { get; set; } = "";
    public string ClassName { get; set; } = "";
    public string PeriodLabel { get; set; } = "";
    public decimal? Amount { get; set; }
}

public sealed class NormalizedClaim : NormalizedRecord
{
    public string BenefitLine { get; set; } = "";
    public string Category { get; set; } = "";
    public string PeriodLabel { get; set; } = "";
    public decimal? PaidClaims { get; set; }
    public decimal? IncurredClaims { get; set; }
}

public sealed class NormalizedClaimsExperience : NormalizedRecord
{
    public string BenefitLine { get; set; } = "";
    public DateOnly? PeriodStart { get; set; }
    public DateOnly? PeriodEnd { get; set; }
    public decimal? Premium { get; set; }
    public decimal? PaidClaims { get; set; }
    public decimal? IncurredClaims { get; set; }
    public decimal? LossRatio { get; set; }
    public int? Lives { get; set; }
}

public sealed class NormalizedBillingStatement : NormalizedRecord
{
    public string BillingPeriod { get; set; } = "";
    public string Province { get; set; } = "";
    public string BenefitLine { get; set; } = "";
    public string ClassCode { get; set; } = "";
    public string OptionLabel { get; set; } = "";
    public int? Lives { get; set; }
    // Kept for compatibility: this is the current-month premium before adjustments.
    public decimal? MonthlyPremium { get; set; }
    public decimal? PremiumAdjustments { get; set; }
    public decimal? PremiumBeforeTax { get; set; }
    public decimal? SalesTax { get; set; }
    public decimal? TotalPremiumAndTax { get; set; }
}

public sealed class NormalizedPooling : NormalizedRecord
{
    public string BenefitLine { get; set; } = "";
    public decimal? PoolingThreshold { get; set; }
    public decimal? PoolingCharge { get; set; }
}

public sealed class NormalizedCommission : NormalizedRecord
{
    public string BenefitLine { get; set; } = "";
    public decimal? CommissionRate { get; set; }
    public decimal? CommissionAmount { get; set; }
}

public sealed class NormalizedPlanDesign : NormalizedRecord
{
    public string BenefitCategory { get; set; } = "";
    public string BenefitName { get; set; } = "";
    public string ProvisionCode { get; set; } = "";
    public string ProvisionLabel { get; set; } = "";
    public string CoverageDescription { get; set; } = "";
    public decimal? ReimbursementPercent { get; set; }
    public string ReimbursementText { get; set; } = "";
    public string MaximumAmount { get; set; } = "";
    public string Frequency { get; set; } = "";
    public string WaitingPeriod { get; set; } = "";
    public string TerminationAge { get; set; } = "";
    public string ClassName { get; set; } = "";
}

public sealed class NormalizedQuote : NormalizedRecord
{
    public string CarrierName { get; set; } = "";
    public string QuoteName { get; set; } = "";
    public decimal? MonthlyPremium { get; set; }
    public string Recommendation { get; set; } = "";
}

public sealed class NormalizedProposal : NormalizedRecord
{
    public string CarrierName { get; set; } = "";
    public string ProposalName { get; set; } = "";
    public decimal? ProposedPremium { get; set; }
    public string Status { get; set; } = "";
}

public sealed class NormalizedFlexibleBenefitPlan : NormalizedRecord
{
    public string AdministratorName { get; set; } = "";
    public string PlanName { get; set; } = "";
    public string PlanType { get; set; } = "";
    public DateOnly? EffectiveDate { get; set; }
    public int ClassCount { get; set; }
    public int EnabledCategoryCount { get; set; }
    public string ReviewStatus { get; set; } = FlexibleBenefitReviewStatuses.Pending;
    public FlexibleBenefitExtractionResult? Extraction { get; set; }
}

public sealed class ImportBatchRecord
{
    public long ImportBatchId { get; set; }
    public string SourceType { get; set; } = "";
    public string SourceCarrier { get; set; } = "";
    public string SourceFileName { get; set; } = "";
    public string SourceDocumentId { get; set; } = "";
    public string SourceApiEndpoint { get; set; } = "";
    public string DocumentType { get; set; } = "";
    public string Status { get; set; } = "";
    public DateTime ImportedAt { get; set; }
    public string ImportedBy { get; set; } = "";
    public decimal? ConfidenceScore { get; set; }
    public string ExtractionMethod { get; set; } = "";
    public string RawSourceReference { get; set; } = "";
    public string Notes { get; set; } = "";
}

public sealed class ImportLogRecord
{
    public long ImportLogId { get; set; }
    public long? ImportBatchId { get; set; }
    public string StepName { get; set; } = "";
    public string Status { get; set; } = "";
    public string Message { get; set; } = "";
    public DateTime CreatedAt { get; set; } = DateTime.Now;
}

public sealed class ValidationIssueRecord
{
    public long ValidationIssueId { get; set; }
    public long? ImportBatchId { get; set; }
    public long? SourceDocumentId { get; set; }
    public string IssueCode { get; set; } = "";
    public string Severity { get; set; } = "Warning";
    public string TableName { get; set; } = "";
    public string FieldName { get; set; } = "";
    public string Message { get; set; } = "";
    public string RawSourceReference { get; set; } = "";
    public DateTime CreatedAt { get; set; } = DateTime.Now;
}

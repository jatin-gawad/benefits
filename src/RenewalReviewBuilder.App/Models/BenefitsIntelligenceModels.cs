namespace RenewalReviewBuilder.App.Models;

public static class BenefitsIntelligenceModules
{
    public const string RenewalIntelligence = "Renewal Intelligence";
    public const string ClaimsExperienceIntelligence = "Claims Experience Intelligence";
    public const string BenefitBookletIntelligence = "Benefit Booklet Intelligence";
    public const string ProposalComparisonIntelligence = "Proposal Comparison Intelligence";
    public const string UnderwritingIntelligence = "Underwriting Intelligence";
    public const string CarrierIntelligence = "Carrier Intelligence";
    public const string BillingPremiumIntelligence = "Billing Premium Intelligence";
    public const string FinancialIntelligence = "Financial Intelligence";
    public const string FlexibleBenefitsIntelligence = "Flexible Benefits Intelligence";
}

public static class DocumentTypes
{
    public const string Renewal = "renewal";
    public const string HistoricalRenewalSupportingWorkbook = "historical_renewal_supporting_workbook";
    public const string ClaimsExperience = "claims_experience";
    public const string BenefitBooklet = "benefit_booklet";
    public const string GroupPolicySchedule = "group_policy_schedule";
    public const string Proposal = "proposal";
    public const string Billing = "billing";
    public const string AsoFinancialActivity = "aso_financial_activity";
    public const string FlexibleBenefitsPlan = "flexible_benefits_plan";
    public const string Other = "other";
}

public static class PdfExtractionMethods
{
    public const string PyMuPdf = "PYMUPDF";
    public const string PdfPlumber = "PDFPLUMBER";
    public const string Docling = "DOCLING";
    public const string PaddleOcr = "PADDLEOCR";
    public const string AzureDocumentIntelligence = "AZURE_DOCUMENT_INTELLIGENCE";
    public const string PdfPig = "PDFPIG";
    public const string ExternalBookletExtractor = "EXTERNAL_BOOKLET_EXTRACTOR";
    public const string WindowsLocalOcr = "WINDOWS_LOCAL_OCR";
}

public sealed class DatabaseOperationResult
{
    public bool Succeeded { get; init; }
    public string Message { get; init; } = "";
    public string DatabasePath { get; init; } = "";

    public static DatabaseOperationResult Success(string message, string databasePath)
    {
        return new DatabaseOperationResult
        {
            Succeeded = true,
            Message = message,
            DatabasePath = databasePath
        };
    }

    public static DatabaseOperationResult Failure(string message, string databasePath)
    {
        return new DatabaseOperationResult
        {
            Succeeded = false,
            Message = message,
            DatabasePath = databasePath
        };
    }
}

public sealed class DatabaseSaveResult
{
    public bool Succeeded { get; init; }
    public string Message { get; init; } = "";
    public string DatabasePath { get; init; } = "";
    public long? ClientId { get; init; }
    public long? DocumentId { get; init; }
    public int InsertedRows { get; init; }
    public int UpdatedRows { get; init; }

    public static DatabaseSaveResult Success(
        string message,
        string databasePath,
        long? clientId = null,
        long? documentId = null,
        int insertedRows = 0,
        int updatedRows = 0)
    {
        return new DatabaseSaveResult
        {
            Succeeded = true,
            Message = message,
            DatabasePath = databasePath,
            ClientId = clientId,
            DocumentId = documentId,
            InsertedRows = insertedRows,
            UpdatedRows = updatedRows
        };
    }

    public static DatabaseSaveResult Failure(string message, string databasePath)
    {
        return new DatabaseSaveResult
        {
            Succeeded = false,
            Message = message,
            DatabasePath = databasePath
        };
    }
}

public sealed class ClientRecord
{
    public long ClientId { get; set; }
    public string ClientName { get; set; } = "";
    public string ClientNumber { get; set; } = "";
    public string Carrier { get; set; } = "";
    public string Industry { get; set; } = "";
    public string Province { get; set; } = "";
    public string Advisor { get; set; } = "";
    public DateOnly? EffectiveDate { get; set; }
    public int? RenewalMonth { get; set; }
    public string Status { get; set; } = "Active";
    public DateTime DateCreated { get; set; }
    public DateTime DateUpdated { get; set; }
}

public sealed class DocumentRecord
{
    public long DocumentId { get; set; }
    public long ClientId { get; set; }
    public string DocumentType { get; set; } = "";
    public string ModuleName { get; set; } = "";
    public string FileName { get; set; } = "";
    public string FileType { get; set; } = "";
    public string FilePath { get; set; } = "";
    public string FileHash { get; set; } = "";
    public string Carrier { get; set; } = "";
    public DateOnly? EffectiveDate { get; set; }
    public DateOnly? PeriodStart { get; set; }
    public DateOnly? PeriodEnd { get; set; }
    public DateTime UploadDate { get; set; }
    public DateTime ProcessedDate { get; set; }
    public string ProcessingStatus { get; set; } = "";
}

public sealed class PdfExtractionLogRecord
{
    public long LogId { get; set; }
    public long DocumentId { get; set; }
    public string ExtractionMethod { get; set; } = "";
    public decimal? ConfidenceScore { get; set; }
    public string FallbackReason { get; set; } = "";
    public int? PageCount { get; set; }
    public int? TableCount { get; set; }
    public decimal? ProcessingTime { get; set; }
    public string ErrorMessage { get; set; } = "";
    public bool ManualReviewRequired { get; set; }
    public DateTime DateProcessed { get; set; }
}

public sealed class RenewalRecord
{
    public long RenewalId { get; set; }
    public long ClientId { get; set; }
    public int RenewalYear { get; set; }
    public string Carrier { get; set; } = "";
    public string BenefitLine { get; set; } = "";
    public decimal? CurrentRate { get; set; }
    public decimal? RenewalRate { get; set; }
    public decimal? RateChangePercent { get; set; }
    public decimal? Premium { get; set; }
    public decimal? Claims { get; set; }
    public decimal? LossRatio { get; set; }
    public int? Lives { get; set; }
    public string RenewalStatus { get; set; } = "";
    public long? SourceDocumentId { get; set; }
    public DateTime DateLoaded { get; set; }
}

public sealed class RenewalQueueRecord
{
    public long ClientId { get; set; }
    public long RenewalId { get; set; }
    public long? SourceDocumentId { get; set; }
    public string SourceFilePath { get; set; } = "";
    public string ClientName { get; set; } = "";
    public string PolicyNumber { get; set; } = "";
    public string Carrier { get; set; } = "";
    public DateOnly? RenewalDate { get; set; }
    public int RenewalYear { get; set; }
    public decimal? CurrentPremium { get; set; }
    public decimal? RenewalPremium { get; set; }
    public decimal? PremiumChangePercent { get; set; }
    public decimal? EhcLossRatio { get; set; }
    public decimal? PoolingThreshold { get; set; }
    public decimal? PoolingCharge { get; set; }
    public decimal? DentalLossRatio { get; set; }
    public int? Lives { get; set; }
    public string Status { get; set; } = "";
    public string SelectedRenewalVersion { get; set; } = "";
    public DateTime DateLoaded { get; set; }
}

public sealed class RateRecord
{
    public long RateId { get; set; }
    public long ClientId { get; set; }
    public string Carrier { get; set; } = "";
    public string BenefitLine { get; set; } = "";
    public string ClassName { get; set; } = "";
    public decimal? CurrentRate { get; set; }
    public decimal? ProposedRate { get; set; }
    public decimal? NegotiatedRate { get; set; }
    public string RateBasis { get; set; } = "";
    public DateOnly? EffectiveDate { get; set; }
    public long? SourceDocumentId { get; set; }
    public DateTime DateLoaded { get; set; }
}

public sealed class RenewalAssumptionRecord
{
    public long AssumptionId { get; set; }
    public long ClientId { get; set; }
    public string Carrier { get; set; } = "";
    public string BenefitLine { get; set; } = "";
    public decimal? Trend { get; set; }
    public decimal? TargetLossRatio { get; set; }
    public decimal? PoolingCharge { get; set; }
    public decimal? PoolingThreshold { get; set; }
    public decimal? Ibnr { get; set; }
    public decimal? AdminFee { get; set; }
    public decimal? RiskCharge { get; set; }
    public DateOnly? EffectiveDate { get; set; }
}

public static class TlrBasisOptions
{
    public const string NetOfPremiumTaxOnly = "Net of premium tax only";
    public const string NetOfTaxAndCommission = "Net of tax and commission";
    public const string GrossBilledPremium = "Gross billed premium";
    public const string IncludesAdminMargin = "Includes admin margin";
    public const string ExcludesAdminMargin = "Excludes admin margin";

    public static IReadOnlyList<string> All { get; } =
    [
        NetOfPremiumTaxOnly,
        NetOfTaxAndCommission,
        GrossBilledPremium,
        IncludesAdminMargin,
        ExcludesAdminMargin
    ];

    public static string Normalize(string? value)
    {
        var cleaned = value?.Trim() ?? "";
        if (string.IsNullOrWhiteSpace(cleaned))
        {
            return "";
        }

        var exact = All.FirstOrDefault(option => option.Equals(cleaned, StringComparison.OrdinalIgnoreCase));
        if (exact is not null)
        {
            return exact;
        }

        var normalized = cleaned.ToLowerInvariant();
        if (normalized.Contains("net tlr before premium tax", StringComparison.Ordinal))
        {
            return NetOfPremiumTaxOnly;
        }

        if (normalized.Contains("premium tax", StringComparison.Ordinal) && !normalized.Contains("commission", StringComparison.Ordinal))
        {
            return NetOfPremiumTaxOnly;
        }

        if (normalized.Contains("tax", StringComparison.Ordinal) && normalized.Contains("commission", StringComparison.Ordinal))
        {
            return NetOfTaxAndCommission;
        }

        if (normalized.Contains("gross", StringComparison.Ordinal) || normalized.Contains("billed", StringComparison.Ordinal))
        {
            return GrossBilledPremium;
        }

        if (normalized.Contains("pooled target", StringComparison.Ordinal) ||
            normalized.Contains("starter manual", StringComparison.Ordinal))
        {
            return GrossBilledPremium;
        }

        if (normalized.Contains("includes", StringComparison.Ordinal) && normalized.Contains("admin", StringComparison.Ordinal))
        {
            return IncludesAdminMargin;
        }

        if (normalized.Contains("excludes", StringComparison.Ordinal) && normalized.Contains("admin", StringComparison.Ordinal))
        {
            return ExcludesAdminMargin;
        }

        return cleaned;
    }
}

public sealed class CarrierAssumptionRecord
{
    public long CarrierAssumptionId { get; set; }
    public string Carrier { get; set; } = "";
    public string BlockName { get; set; } = "";
    public string BenefitLine { get; set; } = "";
    public string FundingType { get; set; } = "";
    public decimal? TargetLossRatio { get; set; }
    public string TlrBasis { get; set; } = "";
    public decimal? PremiumTax { get; set; }
    public decimal? Commission { get; set; }
    public decimal? AdminMargin { get; set; }
    public decimal? RiskMargin { get; set; }
    public decimal? Trend { get; set; }
    public decimal? CredibilityFactor { get; set; }
    public decimal? PoolingThreshold { get; set; }
    public decimal? PoolingCharge { get; set; }
    public decimal? ReserveIbnr { get; set; }
    public DateOnly? EffectiveDate { get; set; }
    public string Status { get; set; } = "Active";
    public string Notes { get; set; } = "";
    public DateTime DateLoaded { get; set; }
}

public sealed class RenewalProjectionRecord
{
    public long ProjectionId { get; set; }
    public long ClientId { get; set; }
    public string BenefitLine { get; set; } = "";
    public decimal? CurrentPremium { get; set; }
    public decimal? ProjectedClaims { get; set; }
    public decimal? TrendAdjustment { get; set; }
    public decimal? RequiredPremium { get; set; }
    public decimal? ProposedPremium { get; set; }
    public decimal? NegotiatedPremium { get; set; }
    public decimal? RenewalChangePercent { get; set; }
    public decimal? Savings { get; set; }
}

public sealed class RenewalVersionRecord
{
    public long VersionId { get; set; }
    public long ClientId { get; set; }
    public long? RenewalId { get; set; }
    public int? RenewalYear { get; set; }
    public string Carrier { get; set; } = "";
    public string VersionName { get; set; } = "";
    public string SourceModule { get; set; } = "";
    public string Status { get; set; } = "";
    public bool IsSelected { get; set; }
    public DateOnly? EffectiveDate { get; set; }
    public long? SourceDocumentId { get; set; }
    public DateTime DateCreated { get; set; }
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
    public List<RenewalRateVersionRecord> Rates { get; } = [];
    public List<RenewalCalculationRecord> Calculations { get; } = [];
}

public sealed class RenewalRateVersionRecord
{
    public long RateVersionId { get; set; }
    public long VersionId { get; set; }
    public string BenefitLine { get; set; } = "";
    public string ClassName { get; set; } = "";
    public string AccountClass { get; set; } = "";
    public decimal? Volume { get; set; }
    public int? Lives { get; set; }
    public decimal? CurrentRate { get; set; }
    public decimal? CurrentPremium { get; set; }
    public decimal? ProposedRate { get; set; }
    public decimal? ProposedPremium { get; set; }
    public decimal? VersionRate { get; set; }
    public decimal? VersionPremium { get; set; }
    public decimal? ChangePercent { get; set; }
    public string RateBasis { get; set; } = "";
    public string RateValueBasis { get; set; } = RateValueBasisNames.CarrierStated;
    public string ClientPresentationRateBasis { get; set; } = RateValueBasisNames.CarrierStated;
    public decimal? CurrentNetRate { get; set; }
    public decimal? CurrentGrossRate { get; set; }
    public decimal? ProposedNetRate { get; set; }
    public decimal? ProposedGrossRate { get; set; }
    public decimal? VersionNetRate { get; set; }
    public decimal? VersionGrossRate { get; set; }
    public string RateScheduleSourceFile { get; set; } = "";
    public string RateScheduleSourceSheet { get; set; } = "";
    public int? RateScheduleSourceRow { get; set; }
    public string SourceReference { get; set; } = "";
}

public sealed class RenewalCalculationRecord
{
    public long CalculationId { get; set; }
    public long VersionId { get; set; }
    public string BenefitLine { get; set; } = "";
    public decimal? TrendedClaims { get; set; }
    public decimal? ProjectedClaims { get; set; }
    public decimal? RequiredPremium { get; set; }
    public decimal? ProjectedLossRatio { get; set; }
    public decimal? RenewalChangePercent { get; set; }
    public decimal? NegotiatedChangePercent { get; set; }
    public decimal? Savings { get; set; }
}

public sealed class NegotiationWorkbenchSaveRequest
{
    public long ClientId { get; set; }
    public long RenewalId { get; set; }
    public string ClientName { get; set; } = "";
    public string Carrier { get; set; } = "";
    public DateOnly? EffectiveDate { get; set; }
    public int RenewalYear { get; set; }
    public string VersionName { get; set; } = RenewalVersionNames.Negotiated;
    public string Status { get; set; } = "Negotiated";
    public string CreatedBy { get; set; } = "";
    public string Notes { get; set; } = "";
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
    public List<NegotiationRateSaveRecord> Rates { get; } = [];
    public List<UnderwritingCalculationResult> Calculations { get; } = [];
    public NegotiationAssumptionSaveRecord Assumptions { get; set; } = new();
    public NegotiationNoteRecord? Note { get; set; }
}

public sealed class NegotiationRateSaveRecord
{
    public string BenefitLine { get; set; } = "";
    public string ClassName { get; set; } = "";
    public string AccountClass { get; set; } = "";
    public string RateBasis { get; set; } = "";
    public decimal? Volume { get; set; }
    public int? Lives { get; set; }
    public decimal? CurrentRate { get; set; }
    public decimal? ProposedRate { get; set; }
    public decimal? ModelledRate { get; set; }
    public decimal? NegotiatedRate { get; set; }
    public decimal? CurrentPremium { get; set; }
    public decimal? ProposedPremium { get; set; }
    public decimal? ModelledPremium { get; set; }
    public decimal? NegotiatedPremium { get; set; }
    public decimal? ProposedChangePct { get; set; }
    public decimal? ModelledChangePct { get; set; }
    public decimal? NegotiatedChangePct { get; set; }
    public decimal? Savings { get; set; }
    public string Status { get; set; } = "";
    public string Notes { get; set; } = "";
    public string Source { get; set; } = "";
    public string RateValueBasis { get; set; } = RateValueBasisNames.CarrierStated;
    public string ClientPresentationRateBasis { get; set; } = RateValueBasisNames.CarrierStated;
    public decimal? CurrentNetRate { get; set; }
    public decimal? CurrentGrossRate { get; set; }
    public decimal? ProposedNetRate { get; set; }
    public decimal? ProposedGrossRate { get; set; }
    public decimal? VersionNetRate { get; set; }
    public decimal? VersionGrossRate { get; set; }
    public string RateScheduleSourceFile { get; set; } = "";
    public string RateScheduleSourceSheet { get; set; } = "";
    public int? RateScheduleSourceRow { get; set; }
}

public sealed class NegotiationAssumptionSaveRecord
{
    public string BenefitLine { get; set; } = "All";
    public decimal? TrendPct { get; set; }
    public decimal? TargetLossRatioPct { get; set; }
    public decimal? PoolingChargePct { get; set; }
    public decimal? PoolingThreshold { get; set; }
    public decimal? IbnrPct { get; set; }
    public decimal? AdminFeePct { get; set; }
    public decimal? RiskChargePct { get; set; }
    public int? Lives { get; set; }
    public decimal? Volume { get; set; }
    public decimal? CredibilityFactor { get; set; }
    public decimal? ManualAdjustmentPct { get; set; }
}

public sealed class NegotiationNoteRecord
{
    public long NoteId { get; set; }
    public long RenewalId { get; set; }
    public long ClientId { get; set; }
    public long RenewalVersionId { get; set; }
    public int RoundNumber { get; set; }
    public DateOnly? NoteDate { get; set; }
    public string CarrierResponse { get; set; } = "";
    public string ConsultantNotes { get; set; } = "";
    public string NextAction { get; set; } = "";
    public string Status { get; set; } = "";
    public DateTime CreatedAt { get; set; } = DateTime.Now;
    public string CreatedBy { get; set; } = "";
}

public sealed class ClaimsExperienceRecord
{
    public long ClaimsId { get; set; }
    public long ClientId { get; set; }
    public DateOnly? ExperiencePeriodStart { get; set; }
    public DateOnly? ExperiencePeriodEnd { get; set; }
    public string BenefitLine { get; set; } = "";
    public decimal? Premium { get; set; }
    public decimal? PaidClaims { get; set; }
    public decimal? IncurredClaims { get; set; }
    public decimal? LossRatio { get; set; }
    public int? Lives { get; set; }
    public long? SourceDocumentId { get; set; }
    public DateTime DateLoaded { get; set; }
}

public sealed class MonthlyClaimsRecord
{
    public long ClaimMonthId { get; set; }
    public long ClientId { get; set; }
    public string Carrier { get; set; } = "";
    public string BenefitLine { get; set; } = "";
    public DateOnly? Month { get; set; }
    public int? Lives { get; set; }
    public decimal? Premium { get; set; }
    public decimal? PaidClaims { get; set; }
    public decimal? IncurredClaims { get; set; }
    public decimal? ClaimsPlusFees { get; set; }
    public decimal? LossRatio { get; set; }
    public long? SourceDocumentId { get; set; }
    public DateTime DateLoaded { get; set; }
}

public sealed class PlanDesignRecord
{
    public long PlanDesignId { get; set; }
    public long ClientId { get; set; }
    public string ClientName { get; set; } = "";
    public string ClientNumber { get; set; } = "";
    public string Carrier { get; set; } = "";
    public string ClassName { get; set; } = "";
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
    public long? SourceDocumentId { get; set; }
    public string SourceFileName { get; set; } = "";
    public string SourceSection { get; set; } = "";
    public string Notes { get; set; } = "";
    public int? SourcePage { get; set; }
    public DateTime DateLoaded { get; set; }
}

public sealed class ClientHistory
{
    public long ClientId { get; init; }
    public List<DocumentRecord> Documents { get; } = [];
    public List<RenewalRecord> Renewals { get; } = [];
    public List<RateRecord> Rates { get; } = [];
    public List<RenewalAssumptionRecord> RenewalAssumptions { get; } = [];
    public List<RenewalProjectionRecord> RenewalProjections { get; } = [];
    public List<RenewalVersionRecord> RenewalVersions { get; } = [];
    public List<ClaimsExperienceRecord> ClaimsHistory { get; } = [];
    public List<MonthlyClaimsRecord> MonthlyClaims { get; } = [];
    public List<PlanDesignRecord> PlanDesign { get; } = [];
    public List<ClientHistoryEventRecord> Events { get; } = [];
}

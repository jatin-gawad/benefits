namespace RenewalReviewBuilder.App.Models;

public sealed class RenewalProcessingResult
{
    public string ExtractedJson { get; set; } = "";
    public List<RenewalValidationItem> ValidationSummary { get; } = [];
    public string? WorkbookOutputPath { get; set; }
    public DatabaseSaveResult? DatabaseStatus { get; set; }
    public List<string> Warnings { get; } = [];
    public List<string> Errors { get; } = [];
    public string OutputFolder { get; set; } = "";
    public string? ExtractionJsonPath { get; set; }
    public string? ExtractionAuditLogPath { get; set; }
    public string? ValidationSummaryPath { get; set; }
    public ExportResult? Export { get; set; }
    public string Status { get; set; } = RenewalProcessingStatuses.NeedsReview;
    public string PlanDesignStatus { get; set; } = PlanDesignStatuses.NotAvailable;
    public string PlanDesignSource { get; set; } = PlanDesignSources.None;
    public RenewalDocumentDiscoveryResult? DocumentDiscovery { get; set; }
}

public static class RenewalProcessingStatuses
{
    public const string Discovered = "Discovered";
    public const string ReadyToProcess = "Ready to Process";
    public const string Complete = "Complete";
    public const string CompletedWithWarning = "Completed with Warning";
    public const string NeedsReview = "Needs Review";
}

public static class PlanDesignStatuses
{
    public const string Available = "Available";
    public const string NotAvailable = "Not Available";
}

public static class PlanDesignSources
{
    public const string SeparateGir = "Separate GIR";
    public const string EmbeddedRenewal = "Embedded Renewal";
    public const string Booklet = "Booklet";
    public const string PolicySchedule = "Policy Schedule";
    public const string Proposal = "Proposal";
    public const string None = "None";
}

public static class DiscoveryDocumentTypes
{
    public const string RwamRateSheet = "RWAM_RATE_SHEET";
    public const string SunLifeRenewalReport = "SUN_LIFE_RENEWAL_REPORT";
    public const string SunLifeExperienceAnalysis = "SUN_LIFE_EXPERIENCE_ANALYSIS";
    public const string GroupHealthRenewalReport = "GROUPHEALTH_RENEWAL_REPORT";
    public const string GroupHealthClaimsExperience = "GROUPHEALTH_CLAIMS_EXPERIENCE";
    public const string GroupHealthTopDrugs = "GROUPHEALTH_TOP_DRUGS";
    public const string GroupHealthPoolingStatement = "GROUPHEALTH_POOLING_STATEMENT";
    public const string GroupHealthClaimsRange = "GROUPHEALTH_CLAIMS_RANGE";
    public const string GroupHealthReserveReport = "GROUPHEALTH_RESERVE_REPORT";
    public const string GroupSourceRenewalReport = "GROUPSOURCE_RENEWAL_REPORT";
    public const string GroupSourceClaimsExperience = "GROUPSOURCE_CLAIMS_EXPERIENCE";
    public const string GroupSourceClaimsDetail = "GROUPSOURCE_CLAIMS_DETAIL";
    public const string GroupSourceTopDrugs = "GROUPSOURCE_TOP_DRUGS";
    public const string Renewal = "Renewal";
    public const string SeparateGroupInformationReport = "Separate Group Information Report";
    public const string ClaimsExperience = "Claims Experience";
    public const string BillingStatement = "Billing Statement";
    public const string AsoFinancialActivity = "ASO Financial Activity";
    public const string BenefitBooklet = "Benefit Booklet";
    public const string GroupPolicySchedule = "Group Policy Schedule";
    public const string Proposal = "Proposal";
    public const string Workbook = "Workbook";
    public const string WordDocument = "Word Document";
    public const string Other = "Other";
}

public sealed class RenewalDocumentDiscoveryResult
{
    public string ClientFolder { get; set; } = "";
    public List<DiscoveredDocument> Documents { get; } = [];
    public DiscoveredDocument? RenewalDocument { get; set; }
    public PlanDesignSourceDiscovery PlanDesign { get; set; } = PlanDesignSourceDiscovery.None();

    public string Status => RenewalDocument is null
        ? RenewalProcessingStatuses.NeedsReview
        : RenewalProcessingStatuses.Discovered;
}

public sealed class DiscoveredDocument
{
    public string FilePath { get; set; } = "";
    public string FileName { get; set; } = "";
    public string FileExtension { get; set; } = "";
    public string DocumentType { get; set; } = DiscoveryDocumentTypes.Other;
    public string Carrier { get; set; } = "";
    public decimal CarrierConfidenceScore { get; set; }
    public decimal DocumentConfidenceScore { get; set; }
    public string ContractNumber { get; set; } = "";
    public string ContractHolder { get; set; } = "";
    public DateOnly? RateEffectiveDate { get; set; }
    public string Status { get; set; } = "ready_for_extraction";
    public List<string> MatchedSignals { get; } = [];
    public List<int> PagesDetected { get; } = [];
    public DateOnly? RenewalDate { get; set; }
    public bool IsRenewal { get; set; }
    public bool IsSeparateGroupInformationReport { get; set; }
    public bool IsClaimsExperience { get; set; }
    public bool IsBillingStatement { get; set; }
    public bool IsAsoFinancialActivity { get; set; }
    public bool IsBenefitBooklet { get; set; }
    public bool IsGroupPolicySchedule { get; set; }
    public bool IsProposal { get; set; }
    public string EmbeddedPlanDesignSection { get; set; } = "";
    public List<int> EmbeddedPlanDesignPages { get; } = [];
}

public sealed class PlanDesignSourceDiscovery
{
    public string Source { get; set; } = PlanDesignSources.None;
    public string Status { get; set; } = PlanDesignStatuses.NotAvailable;
    public string FilePath { get; set; } = "";
    public string FileName { get; set; } = "";
    public string DocumentType { get; set; } = "";
    public string SectionType { get; set; } = "";
    public string Detail { get; set; } = "";
    public int Priority { get; set; } = int.MaxValue;
    public List<int> PageNumbers { get; } = [];

    public bool IsAvailable => !Source.Equals(PlanDesignSources.None, StringComparison.OrdinalIgnoreCase);

    public static PlanDesignSourceDiscovery None()
    {
        return new PlanDesignSourceDiscovery
        {
            Source = PlanDesignSources.None,
            Status = PlanDesignStatuses.NotAvailable,
            Detail = "No plan design source was found."
        };
    }
}

public sealed class CarrierDetectionResult
{
    public string CarrierName { get; set; } = "";
    public decimal ConfidenceScore { get; set; }
}

public sealed class DetectedSection
{
    public string SectionType { get; set; } = "";
    public string SectionName { get; set; } = "";
    public int PageNumber { get; set; }
    public decimal ConfidenceScore { get; set; }
}

public sealed class RenewalDocumentSnapshot
{
    public string SourceFileName { get; set; } = "";
    public string SourceFilePath { get; set; } = "";
    public List<RenewalDocumentPage> Pages { get; } = [];
}

public sealed class RenewalDocumentPage
{
    public int PageNumber { get; set; }
    public string Text { get; set; } = "";
    public List<string> Lines { get; } = [];
}

public sealed class ExtractionAuditEntry
{
    public string SourceFile { get; set; } = "";
    public int? SourcePage { get; set; }
    public string SectionType { get; set; } = "";
    public string FieldName { get; set; } = "";
    public string ExtractedValue { get; set; } = "";
    public string SourceText { get; set; } = "";
    public decimal ConfidenceScore { get; set; }
    public string ValidationStatus { get; set; } = "";
    public string ValidationMessage { get; set; } = "";
}

public sealed class PlanDesignFeatureExtraction
{
    public string Benefit { get; set; } = "";
    public string Category { get; set; } = "";
    public string FeatureName { get; set; } = "";
    public string FeatureValue { get; set; } = "";
    public string Coinsurance { get; set; } = "";
    public string Deductible { get; set; } = "";
    public string Maximum { get; set; } = "";
    public string Frequency { get; set; } = "";
    public string TerminationAge { get; set; } = "";
    public string ClassNumber { get; set; } = "";
    public string ClassName { get; set; } = "";
    public int? SourcePage { get; set; }
    public string SourceText { get; set; } = "";
    public string SourceFileName { get; set; } = "";
    public string PlanDesignSource { get; set; } = PlanDesignSources.None;
    public decimal ConfidenceScore { get; set; }
}

public sealed class RenewalValidationItem
{
    public string Rule { get; set; } = "";
    public string Status { get; set; } = "";
    public string Message { get; set; } = "";
}

public sealed class RenewalProcessingExtractedJson
{
    public CarrierDetectionResult CarrierDetection { get; set; } = new();
    public RenewalDocumentDiscoveryResult? DocumentDiscovery { get; set; }
    public string Status { get; set; } = RenewalProcessingStatuses.NeedsReview;
    public string PlanDesignStatus { get; set; } = PlanDesignStatuses.NotAvailable;
    public string PlanDesignSource { get; set; } = PlanDesignSources.None;
    public List<DetectedSection> Sections { get; } = [];
    public RenewalCase RenewalCase { get; set; } = new();
    public PdfExtractionResult Pdf { get; set; } = new();
    public WorkbookExtractionResult Workbook { get; set; } = new();
    public List<Finding> Findings { get; } = [];
    public List<SourceCheck> SourceChecks { get; } = [];
    public List<PlanDesignFeatureExtraction> PlanDesignFeatures { get; } = [];
    public List<NemMaximumRow> NemMaximums { get; } = [];
    public List<ExtractionAuditEntry> AuditLog { get; } = [];
}

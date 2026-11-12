namespace RenewalReviewBuilder.App.Models;

public static class BillingImportStatuses
{
    public const string Completed = "Completed";
    public const string NeedsReview = "Needs Review";
    public const string Failed = "Failed";
}

public static class BillingReconciliationStatuses
{
    public const string Passed = "Passed";
    public const string Warning = "Warning";
    public const string Failed = "Failed";
    public const string NotAvailable = "Not Available";
}

public enum ManulifeBillingSaveMode
{
    Cancel = 0,
    ImportAsNewVersion = 1,
    ReplaceExistingVersion = 2
}

public static class ManulifeBillingSaveOutcomes
{
    public const string Saved = "Saved";
    public const string Duplicate = "Duplicate";
    public const string ExistingVersion = "Existing Version";
    public const string Failed = "Failed";
}

public static class BillingExportStatuses
{
    public const string NotRequested = "Not Requested";
    public const string Completed = "Completed";
    public const string Failed = "Failed";
    public const string Skipped = "Skipped";
}

public sealed class ManulifeBillingValidationOptions
{
    public decimal MoneyTolerance { get; init; } = 0.02m;
    public int ExactLivesTolerance { get; init; }
    public int LivesWarningTolerance { get; init; } = 1;
}

public sealed class BillingExtractionResult
{
    public string ClientName { get; set; } = "";
    public string ClientNumber { get; set; } = "";
    public string Carrier { get; set; } = "Manulife";
    public string BillingPeriod { get; set; } = "";
    public DateOnly? PeriodStart { get; set; }
    public DateOnly? PeriodEnd { get; set; }
    public string BillingId { get; set; } = "";
    public string SourceFile { get; set; } = "";
    public int PageCount { get; set; }
    public bool ReadablePdfExtraction { get; set; }
    public bool ClientIdentityConfirmed { get; set; }
    public bool CarrierIdentityConfirmed { get; set; }
    public bool ReconciliationExecuted { get; set; }
    public List<BillingOptionPremiumDetail> Details { get; } = [];
    public List<BillingProvinceSummary> ProvinceSummaries { get; } = [];
    public List<BillingOptionSummary> OptionSummaries { get; } = [];
    public List<BillingClassMapping> ClassMappings { get; } = [];
    public List<BillingReconciliationResult> ReconciliationResults { get; } = [];
    public List<string> Warnings { get; } = [];
    public BillingCsvExportResult? CsvExport { get; set; }

    public decimal? PdfGrandPremiumBeforeTax { get; set; }
    public decimal? PdfGrandTax { get; set; }
    public decimal? PdfGrandTotalWithTax { get; set; }

    // Parsing alone cannot complete an import. The coordinator advances this only after validation and persistence.
    public string ImportStatus { get; set; } = BillingImportStatuses.NeedsReview;
}

public sealed class BillingOptionPremiumDetail
{
    public long? StatementVersionId { get; set; }
    public long ClientId { get; set; }
    public long? SourceDocumentId { get; set; }
    public string BillingPeriod { get; set; } = "";
    public string Province { get; set; } = "";
    public string Benefit { get; set; } = "";
    public string ClassCode { get; set; } = "";
    public string OptionLabel { get; set; } = "";
    public string DependentCode { get; set; } = "";
    public int? Lives { get; set; }
    public decimal? Volume { get; set; }
    public decimal MonthlyPremium { get; set; }
    public decimal PremiumAdjustments { get; set; }
    public decimal TotalPremiumAdjustments { get; set; }
    public int SourcePage { get; set; }
    public string SourceFile { get; set; } = "";
    public decimal Confidence { get; set; } = 0.90m;

    public decimal CurrentMonthPremium => MonthlyPremium;
    public decimal PriorPeriodAdjustment => PremiumAdjustments;
    public decimal PremiumBeforeTax => TotalPremiumAdjustments;
}

public sealed class BillingProvinceSummary
{
    public long? StatementVersionId { get; set; }
    public long ClientId { get; set; }
    public long? SourceDocumentId { get; set; }
    public string BillingPeriod { get; set; } = "";
    public string Province { get; set; } = "";
    public int TotalLives { get; set; }
    public decimal TotalMonthlyPremium { get; set; }
    public decimal EstimatedAnnualPremium { get; set; }
    public decimal TotalTax { get; set; }
    public decimal TotalPremiumWithTax { get; set; }
    public string SourceFile { get; set; } = "";

    public decimal PremiumBeforeTax => TotalMonthlyPremium;
    public decimal SalesTax => TotalTax;
    public decimal TotalPremiumAndTax => TotalPremiumWithTax;
}

public sealed class BillingOptionSummary
{
    public long? StatementVersionId { get; set; }
    public long ClientId { get; set; }
    public long? SourceDocumentId { get; set; }
    public string BillingPeriod { get; set; } = "";
    public string Province { get; set; } = "";
    public string Benefit { get; set; } = "";
    public string ClassCode { get; set; } = "";
    public string OptionLabel { get; set; } = "";
    public int Lives { get; set; }
    public decimal MonthlyPremium { get; set; }
    public decimal EstimatedAnnualPremium { get; set; }
    public string SourceFile { get; set; } = "";
}

public sealed class BillingClassMapping
{
    public long ClientId { get; set; }
    public string Carrier { get; set; } = "Manulife";
    public string ClassCode { get; set; } = "";
    public string OptionLabel { get; set; } = "";
    public string FirstSeenBillingPeriod { get; set; } = "";
    public string LastSeenBillingPeriod { get; set; } = "";
    public string SourceFile { get; set; } = "";
    public DateTime CreatedAt { get; set; } = DateTime.Now;
    public DateTime UpdatedAt { get; set; } = DateTime.Now;
}

public sealed class BillingReconciliationResult
{
    public long? StatementVersionId { get; set; }
    public long ClientId { get; set; }
    public long? SourceDocumentId { get; set; }
    public string BillingPeriod { get; set; } = "";
    public string SourceFile { get; set; } = "";
    public string ReconciliationLevel { get; set; } = "";
    public string Province { get; set; } = "";
    public string Benefit { get; set; } = "";
    public string ClassCode { get; set; } = "";
    public int? PdfLives { get; set; }
    public int? CalculatedLives { get; set; }
    public int? LivesDifference { get; set; }
    public string LivesStatus { get; set; } = BillingReconciliationStatuses.NotAvailable;
    public string LivesWarningMessage { get; set; } = "";
    public decimal? PdfPremium { get; set; }
    public decimal? CalculatedPremium { get; set; }
    public decimal? PremiumDifference { get; set; }
    public decimal? PdfTax { get; set; }
    public decimal? CalculatedTax { get; set; }
    public decimal? TaxDifference { get; set; }
    public decimal? PdfTotalWithTax { get; set; }
    public decimal? CalculatedTotalWithTax { get; set; }
    public decimal? TotalDifference { get; set; }
    public string Status { get; set; } = BillingReconciliationStatuses.NotAvailable;
    public string WarningMessage { get; set; } = "";
    public int? SourcePage { get; set; }
    public DateTime CreatedAt { get; set; } = DateTime.Now;
}

public sealed class BillingCsvExportResult
{
    public string OutputFolder { get; set; } = "";
    public string ClassesDetectedPath { get; set; } = "";
    public string ReconciliationResultsPath { get; set; } = "";
    public string ValidationErrorsPath { get; set; } = "";
}

public sealed class BillingCompletionIssue
{
    public string Code { get; init; } = "";
    public string Message { get; init; } = "";
    public bool IsBlocking { get; init; }
}

public sealed class BillingCompletionValidationResult
{
    public bool CanPersist => Issues.All(issue => !issue.IsBlocking);
    public bool CanComplete => CanPersist && Issues.Count == 0;
    public string RecommendedStatus => !CanPersist
        ? BillingImportStatuses.Failed
        : CanComplete
            ? BillingImportStatuses.Completed
            : BillingImportStatuses.NeedsReview;
    public List<BillingCompletionIssue> Issues { get; } = [];
    public int WarningCount => Issues.Count(issue => !issue.IsBlocking);
    public int ErrorCount => Issues.Count(issue => issue.IsBlocking);
}

public sealed class ManulifeBillingImportRequest
{
    public string InputPath { get; init; } = "";
    public string OutputFolder { get; init; } = "";
    public bool ExportCsv { get; init; } = true;
    public ManulifeBillingSaveMode SaveMode { get; init; } = ManulifeBillingSaveMode.Cancel;
    public long? SelectedClientId { get; init; }
    public string SelectedClientName { get; init; } = "";
    public string SelectedClientNumber { get; init; } = "";
    public string PerformedBy { get; init; } = "";
}

public sealed class ManulifeBillingPersistenceResult
{
    public bool Succeeded { get; init; }
    public string Outcome { get; init; } = ManulifeBillingSaveOutcomes.Failed;
    public string Message { get; init; } = "";
    public string DatabasePath { get; init; } = "";
    public long? ClientId { get; init; }
    public long? DocumentId { get; init; }
    public long? ImportBatchId { get; init; }
    public long? StatementVersionId { get; init; }
    public long? ExistingStatementVersionId { get; init; }
    public int VersionNumber { get; init; }
    public int SavedRowCount { get; init; }
    public bool CrossClientHashMatch { get; init; }
}

public sealed class ManulifeBillingImportResult
{
    public bool Succeeded { get; set; }
    public string FinalStatus { get; set; } = BillingImportStatuses.Failed;
    public long? ImportBatchId { get; set; }
    public long? DocumentId { get; set; }
    public long? StatementVersionId { get; set; }
    public int ParsedRowCount { get; set; }
    public int SavedRowCount { get; set; }
    public int WarningCount { get; set; }
    public int ErrorCount { get; set; }
    public string ReconciliationStatus { get; set; } = BillingReconciliationStatuses.NotAvailable;
    public string ExportStatus { get; set; } = BillingExportStatuses.NotRequested;
    public string UserMessage { get; set; } = "";
    public List<string> TechnicalErrors { get; } = [];
    public BillingExtractionResult? Extraction { get; set; }
    public BillingCsvExportResult? CsvExport { get; set; }
    public ManulifeBillingPersistenceResult? Persistence { get; set; }
}

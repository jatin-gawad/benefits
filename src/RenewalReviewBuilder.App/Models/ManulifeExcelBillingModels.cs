namespace RenewalReviewBuilder.App.Models;

public static class CarrierBillingHeaderStatuses
{
    public const string ExactMatch = "exact_match";
    public const string ApprovedAlias = "approved_alias";
    public const string DynamicallyDetectedBenefit = "dynamically_detected_benefit";
    public const string Unmapped = "unmapped";
    public const string Ambiguous = "ambiguous";
}

public static class CarrierBillingExceptionSeverities
{
    public const string Warning = "Warning";
    public const string Error = "Error";
}

public sealed class ManulifeExcelBillingImportOptions
{
    public decimal ReconciliationTolerance { get; init; } = 0.02m;
    public decimal MinorVarianceThreshold { get; init; } = 0.02m;
    public long MaximumFileSizeBytes { get; init; } = 50L * 1024 * 1024;
}

public sealed class ManulifeExcelBillingImportRequest
{
    public string InputPath { get; init; } = "";
    public long? SelectedClientId { get; init; }
    public string SelectedClientName { get; init; } = "";
    public string BillingPeriod { get; init; } = "";
    public DateOnly? StatementDate { get; init; }
    public string ImportedBy { get; init; } = Environment.UserName;
}

public sealed class ManulifeExcelBillingImportPreview
{
    public string CarrierName { get; set; } = "Manulife";
    public long? SelectedClientId { get; set; }
    public string ClientName { get; set; } = "";
    public string DetectedClientName { get; set; } = "";
    public string BillingPeriod { get; set; } = "";
    public DateOnly? StatementDate { get; set; }
    public string SourceFilePath { get; set; } = "";
    public string SourceFileName { get; set; } = "";
    public string SourceFileHash { get; set; } = "";
    public bool DuplicateFileDetected { get; set; }
    public bool RequiresClientConfirmation { get; set; }
    public int SourceRowCount { get; set; }
    public int EmployeeRecordCount { get; set; }
    public int BenefitRecordCount { get; set; }
    public decimal TotalCurrentMonthPremium { get; set; }
    public decimal TotalPriorPeriodAdjustment { get; set; }
    public decimal TotalSalesTax { get; set; }
    public decimal TotalPremiumAndTax { get; set; }
    public decimal CalculatedPremiumAndTax { get; set; }
    public decimal ReconciliationDifference { get; set; }
    public int ErrorCount => Exceptions.Count(item => item.Severity == CarrierBillingExceptionSeverities.Error);
    public int WarningCount => Exceptions.Count(item => item.Severity == CarrierBillingExceptionSeverities.Warning);
    public bool CanConfirm => !DuplicateFileDetected && !RequiresClientConfirmation && ErrorCount == 0;
    public List<CarrierBillingHeaderMapping> HeaderMappings { get; } = [];
    public List<CarrierBillingEmployeeRecord> EmployeeRecords { get; } = [];
    public List<CarrierBillingBenefitRecord> BenefitRecords { get; } = [];
    public List<CarrierBillingException> Exceptions { get; } = [];
    public List<CarrierBillingAggregate> TotalsByBenefit { get; } = [];
    public List<CarrierBillingAggregate> TotalsByClass { get; } = [];
    public List<CarrierBillingAggregate> TotalsByPlan { get; } = [];
    public List<CarrierBillingAggregate> TotalsByDivision { get; } = [];
    public List<string> NewBenefitCodes { get; } = [];
    public List<string> UnmappedHeaders { get; } = [];
    public List<CarrierBillingPreviewEmployee> SampleEmployeeRecords { get; } = [];
}

public sealed class CarrierBillingHeaderMapping
{
    public int ColumnNumber { get; set; }
    public string SourceHeader { get; set; } = "";
    public string NormalizedHeader { get; set; } = "";
    public string TargetField { get; set; } = "";
    public string BenefitCode { get; set; } = "";
    public string Status { get; set; } = CarrierBillingHeaderStatuses.Unmapped;
}

public sealed class CarrierBillingEmployeeRecord
{
    public long CarrierBillingEmployeeRecordId { get; set; }
    public long CarrierBillingImportId { get; set; }
    public long ClientId { get; set; }
    public long? CarrierId { get; set; }
    public string CertificateNumber { get; set; } = "";
    public string LastName { get; set; } = "";
    public string FirstName { get; set; } = "";
    public string ClassName { get; set; } = "";
    public string ClassCode { get; set; } = "";
    public string PlanName { get; set; } = "";
    public string PlanCode { get; set; } = "";
    public string ProvinceOfWork { get; set; } = "";
    public string ProvinceOfWorkRaw { get; set; } = "";
    public string ProvinceOfResidence { get; set; } = "";
    public string ProvinceOfResidenceRaw { get; set; } = "";
    public decimal? CurrentMonthPremium { get; set; }
    public decimal? PriorPeriodAdjustment { get; set; }
    public decimal? TotalSalesTax { get; set; }
    public decimal? TotalPremiumAndTax { get; set; }
    public decimal CalculatedCurrentMonthPremium { get; set; }
    public decimal CalculatedPriorPeriodAdjustment { get; set; }
    public decimal CalculatedSalesTax { get; set; }
    public decimal CalculatedPremiumAndTax { get; set; }
    public decimal? ReconciliationDifference { get; set; }
    public string DivisionalBillingCode { get; set; } = "";
    public string BillTotal6 { get; set; } = "";
    public string BillTotal7 { get; set; } = "";
    public int SourceRowNumber { get; set; }
    public string SourceRowHash { get; set; } = "";
    public string ValidationStatus { get; set; } = "Passed";
}

public sealed class CarrierBillingBenefitRecord
{
    public long CarrierBillingBenefitRecordId { get; set; }
    public long CarrierBillingEmployeeRecordId { get; set; }
    public long CarrierBillingImportId { get; set; }
    public long ClientId { get; set; }
    public string CertificateNumber { get; set; } = "";
    public string BenefitCode { get; set; } = "";
    public string BenefitName { get; set; } = "";
    public string BenefitPlanCode { get; set; } = "";
    public string DependentCodeOrVolumeRaw { get; set; } = "";
    public string DependentCode { get; set; } = "";
    public decimal? InsuredVolume { get; set; }
    public string OptionCode { get; set; } = "";
    public decimal? CurrentMonthPremium { get; set; }
    public decimal? PriorPeriodAdjustment { get; set; }
    public decimal? TotalSalesTax { get; set; }
    public decimal? BenefitTotal { get; set; }
    public decimal CalculatedBenefitTotal { get; set; }
    public decimal? ReconciliationDifference { get; set; }
    public int SourceRowNumber { get; set; }
    public string ValidationStatus { get; set; } = "Passed";
}

public sealed class CarrierBillingException
{
    public long CarrierBillingExceptionId { get; set; }
    public long CarrierBillingImportId { get; set; }
    public int? SourceRowNumber { get; set; }
    public string CertificateNumber { get; set; } = "";
    public string BenefitCode { get; set; } = "";
    public string FieldName { get; set; } = "";
    public string ExceptionType { get; set; } = "";
    public string Severity { get; set; } = CarrierBillingExceptionSeverities.Warning;
    public string SourceValue { get; set; } = "";
    public string ExpectedValue { get; set; } = "";
    public string Message { get; set; } = "";
    public DateTime CreatedAt { get; set; } = DateTime.Now;
    public string ResolutionStatus { get; set; } = "Open";
}

public sealed class CarrierBillingAggregate
{
    public string Group { get; set; } = "";
    public int EmployeeCount { get; set; }
    public int BenefitRecordCount { get; set; }
    public decimal CurrentMonthPremium { get; set; }
    public decimal PriorPeriodAdjustment { get; set; }
    public decimal SalesTax { get; set; }
    public decimal TotalPremiumAndTax { get; set; }
}

public sealed class CarrierBillingRateSummary
{
    public string BenefitCode { get; set; } = "";
    public string ClassCode { get; set; } = "";
    public string PlanCode { get; set; } = "";
    public string Coverage { get; set; } = "";
    public string RateBasis { get; set; } = "";
    public decimal Rate { get; set; }
    public int RecordCount { get; set; }
    public string Source { get; set; } = "";
}

public sealed class CarrierBillingPreviewEmployee
{
    public int SourceRowNumber { get; set; }
    public string MaskedCertificateNumber { get; set; } = "";
    public string MaskedEmployeeName { get; set; } = "";
    public string ClassCode { get; set; } = "";
    public string PlanCode { get; set; } = "";
    public decimal? TotalPremiumAndTax { get; set; }
    public string ValidationStatus { get; set; } = "";
}

public sealed class CarrierBillingConfirmationResult
{
    public bool Succeeded { get; set; }
    public string Message { get; set; } = "";
    public long? CarrierBillingImportId { get; set; }
    public long? ClientId { get; set; }
    public long? ImportBatchId { get; set; }
    public int SavedEmployeeRecordCount { get; set; }
    public int SavedBenefitRecordCount { get; set; }
    public string DatabasePath { get; set; } = "";
}

public sealed class CarrierBillingImportHistoryRecord
{
    public long CarrierBillingImportId { get; set; }
    public long ClientId { get; set; }
    public string CarrierName { get; set; } = "";
    public string BillingPeriod { get; set; } = "";
    public string SourceFileName { get; set; } = "";
    public int VersionNumber { get; set; }
    public string ImportStatus { get; set; } = "";
    public DateTime ImportedAt { get; set; }
    public int EmployeeRecordCount { get; set; }
    public int BenefitRecordCount { get; set; }
    public decimal TotalPremiumAndTax { get; set; }
    public int ErrorCount { get; set; }
    public int WarningCount { get; set; }
}

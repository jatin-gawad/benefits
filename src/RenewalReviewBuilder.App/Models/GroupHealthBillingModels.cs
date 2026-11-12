namespace RenewalReviewBuilder.App.Models;

public static class GroupHealthBillingImportStatuses
{
    public const string Draft = "Draft";
    public const string Review = "Review";
    public const string Passed = "Passed";
    public const string Warning = "Warning";
    public const string Failed = "Failed";
}

public static class GroupHealthBillingValidationStatuses
{
    public const string Passed = "Passed";
    public const string Warning = "Warning";
    public const string Failed = "Failed";
    public const string Information = "Information";
    public const string NotAvailable = "Not Available";
}

public enum GroupHealthBillingSaveMode
{
    Cancel = 0,
    Replace = 1,
    NewVersion = 2
}

public sealed class GroupHealthBillingValidationOptions
{
    public decimal MoneyTolerance { get; init; } = 0.02m;
    public decimal PremiumChangeThresholdPercent { get; init; } = 10m;
    public decimal MissingKeyFieldDeduction { get; init; } = 1m;
    public decimal WarningDeduction { get; init; } = 1.5m;
    public decimal FailedDeduction { get; init; } = 8m;
}

public sealed class GroupHealthBillingExtractionResult
{
    public long ClientId { get; set; }
    public long? BillingStatementId { get; set; }
    public string ClientName { get; set; } = "";
    public string ClientNumber { get; set; } = "";
    public string Carrier { get; set; } = "GroupHEALTH";
    public int StatementMonth { get; set; }
    public int StatementYear { get; set; }
    public DateOnly? StatementDate { get; set; }
    public DateOnly? StatementPeriodStart { get; set; }
    public DateOnly? StatementPeriodEnd { get; set; }
    public DateTime ImportDate { get; set; } = DateTime.Now;
    public string SourceFilePath { get; set; } = "";
    public string SourceFileName { get; set; } = "";
    public string SourceFileHash { get; set; } = "";
    public int ImportVersion { get; set; } = 1;
    public string ImportStatus { get; set; } = GroupHealthBillingImportStatuses.Draft;
    public decimal BillingIntegrityScore { get; set; } = 100m;
    public int PageCount { get; set; }
    public decimal? PremiumGrandTotal { get; set; }
    public decimal? GstHstGrandTotal { get; set; }
    public decimal? StatementGrandTotal { get; set; }
    public List<GroupHealthBillingRow> Rows { get; } = [];
    public List<GroupHealthBillingStatementTotal> Totals { get; } = [];
    public List<GroupHealthBillingValidationResult> ValidationResults { get; } = [];
    public List<GroupHealthBillingEmployeeChange> EmployeeChangeLog { get; } = [];
    public List<GroupHealthClientRateTableLine> RateLines { get; } = [];
    public GroupHealthBillingIntegrityReport IntegrityReport { get; set; } = new();
}

public sealed class GroupHealthBillingRow
{
    public long BillingStatementRowId { get; set; }
    public long BillingStatementId { get; set; }
    public long ClientId { get; set; }
    public string ClassCode { get; set; } = "";
    public string EmployeeId { get; set; } = "";
    public string CertificateId { get; set; } = "";
    public string EmployeeName { get; set; } = "";
    public decimal? AnnualEarnings { get; set; }
    public string CostDivision { get; set; } = "";
    public string EmploymentClass { get; set; } = "";
    public string RetroMonth { get; set; } = "";
    public decimal? LifeVolume { get; set; }
    public decimal? LifePremium { get; set; }
    public decimal? AdAndDVolume { get; set; }
    public decimal? AdAndDPremium { get; set; }
    public decimal? LtdVolume { get; set; }
    public decimal? LtdPremium { get; set; }
    public string HealthTier { get; set; } = "";
    public decimal? HealthPremium { get; set; }
    public string DentalTier { get; set; } = "";
    public decimal? DentalPremium { get; set; }
    public decimal? EfapPremium { get; set; }
    public decimal? VirtualHealthcarePremium { get; set; }
    public decimal? FlexAdminFee { get; set; }
    public decimal? MemberAdAndDVolume { get; set; }
    public decimal? MemberAdAndDPremium { get; set; }
    public decimal? MemberAdAndDSpVolume { get; set; }
    public decimal? MemberAdAndDSpPremium { get; set; }
    public decimal? MemberAdAndDChVolume { get; set; }
    public decimal? MemberAdAndDChPremium { get; set; }
    public decimal? EmployeePremium { get; set; }
    public decimal? EmployerPremium { get; set; }
    public decimal? TotalPremium { get; set; }
    public decimal? GstHst { get; set; }
    public decimal? RetroPremium { get; set; }
    public decimal? RetroGstHst { get; set; }
    public decimal? RowTotal { get; set; }
    public int SourcePage { get; set; }
    public int SourceLine { get; set; }
    public string ValidationStatus { get; set; } = GroupHealthBillingValidationStatuses.Passed;
    public string ValidationFlags { get; set; } = "";
    public string SourceText { get; set; } = "";
}

public sealed class GroupHealthBillingStatementTotal
{
    public long BillingStatementTotalId { get; set; }
    public long BillingStatementId { get; set; }
    public long ClientId { get; set; }
    public string TotalLevel { get; set; } = "";
    public string ClassCode { get; set; } = "";
    public string Benefit { get; set; } = "";
    public string TotalType { get; set; } = "";
    public int? Lives { get; set; }
    public decimal? Volume { get; set; }
    public decimal? Premium { get; set; }
    public decimal? GstHst { get; set; }
    public decimal? TotalPremium { get; set; }
    public int SourcePage { get; set; }
}

public sealed class GroupHealthBillingValidationResult
{
    public long BillingValidationResultId { get; set; }
    public long BillingStatementId { get; set; }
    public long? BillingStatementRowId { get; set; }
    public long ClientId { get; set; }
    public string ValidationLevel { get; set; } = "";
    public string CheckName { get; set; } = "";
    public string Status { get; set; } = GroupHealthBillingValidationStatuses.Passed;
    public string Message { get; set; } = "";
    public decimal? ExpectedAmount { get; set; }
    public decimal? ActualAmount { get; set; }
    public decimal? DifferenceAmount { get; set; }
    public string Benefit { get; set; } = "";
    public string ClassCode { get; set; } = "";
    public string EmployeeId { get; set; } = "";
    public string CertificateId { get; set; } = "";
    public int? SourcePage { get; set; }
    public DateTime CreatedAt { get; set; } = DateTime.Now;
}

public sealed class GroupHealthBillingEmployeeChange
{
    public long BillingEmployeeChangeLogId { get; set; }
    public long BillingStatementId { get; set; }
    public long ClientId { get; set; }
    public string Carrier { get; set; } = "GroupHEALTH";
    public int StatementMonth { get; set; }
    public int StatementYear { get; set; }
    public string EmployeeId { get; set; } = "";
    public string CertificateId { get; set; } = "";
    public string EmployeeName { get; set; } = "";
    public string ChangeType { get; set; } = "";
    public string FieldName { get; set; } = "";
    public string PreviousValue { get; set; } = "";
    public string CurrentValue { get; set; } = "";
    public decimal? DifferenceAmount { get; set; }
    public decimal? ChangePercent { get; set; }
    public string Severity { get; set; } = GroupHealthBillingValidationStatuses.Information;
    public string Source { get; set; } = "";
    public DateTime CreatedAt { get; set; } = DateTime.Now;
}

public sealed class GroupHealthClientRateTableLine
{
    public long ClientRateTableLineId { get; set; }
    public long ClientRateTableId { get; set; }
    public long BillingStatementId { get; set; }
    public long ClientId { get; set; }
    public string Carrier { get; set; } = "GroupHEALTH";
    public int EffectiveMonth { get; set; }
    public int EffectiveYear { get; set; }
    public string Benefit { get; set; } = "";
    public string ClassCode { get; set; } = "";
    public string Tier { get; set; } = "";
    public string VolumeBasis { get; set; } = "";
    public decimal? EmployeeRate { get; set; }
    public decimal? EmployerRate { get; set; }
    public decimal? TotalRate { get; set; }
    public decimal? CalculatedRate { get; set; }
    public decimal? ExpectedRate { get; set; }
    public string RateStatus { get; set; } = GroupHealthBillingValidationStatuses.NotAvailable;
    public bool IsCalculated { get; set; } = true;
    public string Source { get; set; } = "Calculated from billing import";
}

public sealed class GroupHealthBillingIntegrityReport
{
    public int TotalEmployeesImported { get; set; }
    public decimal TotalPremiumsImported { get; set; }
    public decimal? StatementTotal { get; set; }
    public decimal? DifferenceAmount { get; set; }
    public int PassedChecks { get; set; }
    public int Warnings { get; set; }
    public int FailedChecks { get; set; }
    public int NewHires { get; set; }
    public int Terminations { get; set; }
    public int SalaryChanges { get; set; }
    public int CoverageChanges { get; set; }
    public int TierChanges { get; set; }
    public int DuplicateRows { get; set; }
    public int MissingValues { get; set; }
    public decimal BillingIntegrityScore { get; set; }
    public string FinalStatus { get; set; } = GroupHealthBillingImportStatuses.Draft;
}

public sealed class GroupHealthBillingHistoryRecord
{
    public long BillingStatementId { get; set; }
    public long ClientId { get; set; }
    public string Carrier { get; set; } = "";
    public int StatementMonth { get; set; }
    public int StatementYear { get; set; }
    public DateOnly? StatementDate { get; set; }
    public DateTime ImportDate { get; set; }
    public string SourceFileName { get; set; } = "";
    public int ImportVersion { get; set; }
    public string ImportStatus { get; set; } = "";
    public decimal BillingIntegrityScore { get; set; }
    public decimal? TotalEmployeePremium { get; set; }
    public decimal? TotalEmployerPremium { get; set; }
    public decimal? TotalPremium { get; set; }
    public int EmployeeCount { get; set; }
    public bool IsReplaced { get; set; }
}

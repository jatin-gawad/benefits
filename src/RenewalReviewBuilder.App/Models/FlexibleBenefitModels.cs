using System.Collections.ObjectModel;
using System.Windows;

namespace RenewalReviewBuilder.App.Models;

public enum FlexibleBenefitPlanType
{
    Hcsa,
    Lsa,
    CombinedHcsaAndLsa,
    WellnessSpendingAccount,
    PersonalSpendingAccount,
    Other
}

public enum FlexibleBenefitYearType
{
    CalendarYear,
    FiscalYear,
    PolicyYear,
    Custom,
    Unknown
}

public enum FlexibleBenefitAllocationType
{
    AllHcsa,
    AllLsa,
    Split,
    EmployeeElection,
    NotApplicable,
    Unknown
}

public enum FlexibleBenefitCarryForwardType
{
    Credit,
    Expense,
    None,
    Both,
    Unknown
}

public enum FlexibleBenefitFundingMethod
{
    PreFunded,
    PayAsYouGo,
    Pad,
    OnlinePayment,
    Cheque,
    Mixed,
    Unknown
}

public enum FlexibleBenefitAccountType
{
    Hcsa,
    Lsa,
    Both,
    Other
}

public enum FlexibleBenefitDocumentSubtype
{
    EnrollmentForm,
    AdministrativeServicesAgreement,
    IndemnityAgreement,
    FeeSchedule,
    PlanAmendment,
    FundingAgreement,
    EmployeeEnrollmentSchedule,
    SpendingAccountPlanDesign,
    CombinedEnrollmentAndAgreement,
    Unknown
}

public enum FlexibleBenefitImportMode
{
    Cancel,
    ImportAsNewVersion,
    ReplaceExistingVersion
}

public enum FlexibleBenefitImportOutcome
{
    Completed,
    NeedsReview,
    Failed,
    Duplicate,
    ExistingVersion,
    Cancelled
}

public static class FlexibleBenefitReviewStatuses
{
    public const string Pending = "Pending";
    public const string NeedsReview = "Needs Review";
    public const string Confirmed = "Confirmed";
    public const string Reviewed = "Reviewed";
    public const string Replaced = "Replaced";
}

public static class FlexibleBenefitExtractionStatuses
{
    public const string Extracted = "Extracted";
    public const string Normalized = "Normalized";
    public const string Validated = "Validated";
    public const string Completed = "Completed";
    public const string NeedsReview = "Needs Review";
    public const string Failed = "Failed";
}

public static class FlexibleBenefitDocumentTypes
{
    public const string FlexibleBenefitsPlanDocument = "FlexibleBenefitsPlanDocument";
}

public sealed class FlexibleBenefitPlan
{
    public long? FlexibleBenefitPlanId { get; set; }
    public long ClientId { get; set; }
    public long? CarrierId { get; set; }
    public long? AdministratorId { get; set; }
    public long? SourceDocumentId { get; set; }
    public long? SourceDocumentVersionId { get; set; }
    public string PolicyNumber { get; set; } = "";
    public string AccountNumber { get; set; } = "";
    public string PlanName { get; set; } = "";
    public FlexibleBenefitPlanType PlanType { get; set; } = FlexibleBenefitPlanType.Other;
    public bool? HcsaEnabled { get; set; }
    public bool? LsaEnabled { get; set; }
    public bool? WellnessAccountEnabled { get; set; }
    public bool? CombinedAccountEnabled { get; set; }
    public DateOnly? EffectiveDate { get; set; }
    public DateOnly? TerminationDate { get; set; }
    public FlexibleBenefitYearType BenefitYearType { get; set; } = FlexibleBenefitYearType.Unknown;
    public int? BenefitYearStartMonth { get; set; }
    public int? BenefitYearStartDay { get; set; }
    public FlexibleBenefitAllocationType DefaultAllocationType { get; set; } = FlexibleBenefitAllocationType.Unknown;
    public decimal? DefaultHcsaPercentage { get; set; }
    public decimal? DefaultLsaPercentage { get; set; }
    public bool? PayAsSubmitted { get; set; }
    public int? RunoffDays { get; set; }
    public FlexibleBenefitCarryForwardType CarryForwardType { get; set; } = FlexibleBenefitCarryForwardType.Unknown;
    public bool? CreditCarryForwardEnabled { get; set; }
    public bool? ExpenseCarryForwardEnabled { get; set; }
    public decimal? CarryForwardMaximum { get; set; }
    public int? CarryForwardExpiryMonths { get; set; }
    public int? DefaultWaitingPeriodDays { get; set; }
    public bool? NewHireProrationDefault { get; set; }
    public int? ChildDependentCutoffAge { get; set; }
    public int? StudentDependentCutoffAge { get; set; }
    public FlexibleBenefitFundingMethod FundingMethod { get; set; } = FlexibleBenefitFundingMethod.Unknown;
    public bool? PreFunded { get; set; }
    public bool? PayAsYouGo { get; set; }
    public bool? PadRequested { get; set; }
    public decimal? SetupFee { get; set; }
    public decimal? AdministrationFeePercentage { get; set; }
    public decimal? AdministrationFeeFlatAmount { get; set; }
    public string AdministrationFeeBasis { get; set; } = "";
    public string TaxTreatment { get; set; } = "";
    public string CurrencyCode { get; set; } = "CAD";
    public bool IsCurrent { get; set; } = true;
    public int VersionNumber { get; set; } = 1;
    public long? SupersedesPlanId { get; set; }
    public string ReviewStatus { get; set; } = FlexibleBenefitReviewStatuses.Pending;
    public string ExtractionStatus { get; set; } = FlexibleBenefitExtractionStatuses.Extracted;
    public DateTime CreatedAt { get; set; } = DateTime.Now;
    public string CreatedBy { get; set; } = Environment.UserName;
    public DateTime UpdatedAt { get; set; } = DateTime.Now;
    public string UpdatedBy { get; set; } = Environment.UserName;
}

public sealed class FlexibleBenefitClass
{
    public long? FlexibleBenefitClassId { get; set; }
    public long? FlexibleBenefitPlanId { get; set; }
    public string SourceKey { get; set; } = Guid.NewGuid().ToString("N");
    public string ExternalClassCode { get; set; } = "";
    public string ClassName { get; set; } = "";
    public string JobDescription { get; set; } = "";
    public string EligibilityDescription { get; set; } = "";
    public decimal? AnnualSingleAmount { get; set; }
    public decimal? AnnualCoupleAmount { get; set; }
    public decimal? AnnualFamilyAmount { get; set; }
    public decimal? AnnualEmployeeAmount { get; set; }
    public decimal? AnnualDependentAmount { get; set; }
    public int? WaitingPeriodDays { get; set; }
    public bool? NewHireProrationEnabled { get; set; }
    public decimal? HcsaAmount { get; set; }
    public decimal? LsaAmount { get; set; }
    public decimal? HcsaPercentage { get; set; }
    public decimal? LsaPercentage { get; set; }
    public string CurrencyCode { get; set; } = "CAD";
    public DateOnly? EffectiveDate { get; set; }
    public DateOnly? TerminationDate { get; set; }
    public bool IsCurrent { get; set; } = true;
    public int? SourcePageNumber { get; set; }
    public decimal? ExtractionConfidence { get; set; }
    public string ReviewStatus { get; set; } = FlexibleBenefitReviewStatuses.Pending;
}

public sealed class FlexibleBenefitCategory
{
    public long? FlexibleBenefitCategoryId { get; set; }
    public long? FlexibleBenefitPlanId { get; set; }
    public string SourceKey { get; set; } = Guid.NewGuid().ToString("N");
    public string CategoryCode { get; set; } = "";
    public string CategoryGroup { get; set; } = "";
    public string CategoryName { get; set; } = "";
    public string Description { get; set; } = "";
    public bool? Enabled { get; set; }
    public bool? Taxable { get; set; }
    public FlexibleBenefitAccountType AccountType { get; set; } = FlexibleBenefitAccountType.Other;
    public bool? RequiresReceipt { get; set; }
    public bool? RequiresProviderDetails { get; set; }
    public string Exclusions { get; set; } = "";
    public int? SourcePageNumber { get; set; }
    public decimal? ExtractionConfidence { get; set; }
    public string ReviewStatus { get; set; } = FlexibleBenefitReviewStatuses.Pending;
}

public sealed class FlexibleBenefitEligibilityRule
{
    public long? FlexibleBenefitEligibilityRuleId { get; set; }
    public long? FlexibleBenefitPlanId { get; set; }
    public long? FlexibleBenefitClassId { get; set; }
    public string SourceKey { get; set; } = Guid.NewGuid().ToString("N");
    public string ClassSourceKey { get; set; } = "";
    public string RuleType { get; set; } = "";
    public string Operator { get; set; } = "";
    public string RuleValue { get; set; } = "";
    public string RuleUnit { get; set; } = "";
    public string Description { get; set; } = "";
    public DateOnly? EffectiveDate { get; set; }
    public int? SourcePageNumber { get; set; }
    public decimal? ExtractionConfidence { get; set; }
    public string ReviewStatus { get; set; } = FlexibleBenefitReviewStatuses.Pending;
}

public sealed class FlexibleBenefitFundingConfiguration
{
    public long? FlexibleBenefitFundingConfigurationId { get; set; }
    public long? FlexibleBenefitPlanId { get; set; }
    public string SourceKey { get; set; } = Guid.NewGuid().ToString("N");
    public FlexibleBenefitFundingMethod FundingMethod { get; set; } = FlexibleBenefitFundingMethod.Unknown;
    public bool? PreFundingRequired { get; set; }
    public bool? PadSelected { get; set; }
    public bool? ContinuousFloatRequired { get; set; }
    public bool? OnlineBankPaymentAllowed { get; set; }
    public bool? ChequeAllowed { get; set; }
    public string PaymentTimingDescription { get; set; } = "";
    public bool? ClaimsPaidOnlyAfterFundsClear { get; set; }
    public bool? InterestPaidOnFunds { get; set; }
    public string CurrencyCode { get; set; } = "CAD";
    public DateOnly? EffectiveDate { get; set; }
    public int? SourcePageNumber { get; set; }
    public decimal? ExtractionConfidence { get; set; }
    public string ReviewStatus { get; set; } = FlexibleBenefitReviewStatuses.Pending;
}

public sealed class FlexibleBenefitFee
{
    public long? FlexibleBenefitFeeId { get; set; }
    public long? FlexibleBenefitPlanId { get; set; }
    public string SourceKey { get; set; } = Guid.NewGuid().ToString("N");
    public string FeeType { get; set; } = "";
    public string FeeCalculationType { get; set; } = "";
    public decimal? Percentage { get; set; }
    public decimal? FlatAmount { get; set; }
    public decimal? MinimumAmount { get; set; }
    public decimal? MaximumAmount { get; set; }
    public bool? Taxable { get; set; }
    public string TaxType { get; set; } = "";
    public string Description { get; set; } = "";
    public DateOnly? EffectiveDate { get; set; }
    public int? NoticePeriodDays { get; set; }
    public int? SourcePageNumber { get; set; }
    public decimal? ExtractionConfidence { get; set; }
    public string ReviewStatus { get; set; } = FlexibleBenefitReviewStatuses.Pending;
}

public sealed class FlexibleBenefitAgreement
{
    public long? FlexibleBenefitAgreementId { get; set; }
    public long? FlexibleBenefitPlanId { get; set; }
    public string SourceKey { get; set; } = Guid.NewGuid().ToString("N");
    public string AgreementType { get; set; } = "";
    public string AgreementVersion { get; set; } = "";
    public DateOnly? AgreementDate { get; set; }
    public DateTime? ExecutionDate { get; set; }
    public string GoverningProvince { get; set; } = "";
    public string SubscriberLegalName { get; set; } = "";
    public string AdministratorLegalName { get; set; } = "";
    public string SubscriberSignerName { get; set; } = "";
    public string SubscriberSignerTitle { get; set; } = "";
    public string AdministratorSignerName { get; set; } = "";
    public string AdministratorSignerTitle { get; set; } = "";
    public bool? SubscriberSignaturePresent { get; set; }
    public bool? AdministratorSignaturePresent { get; set; }
    public bool? ElectronicSignature { get; set; }
    public string SignedCity { get; set; } = "";
    public string SignedProvince { get; set; } = "";
    public bool? IndemnityAgreementSigned { get; set; }
    public bool? FundingAgreementSigned { get; set; }
    public string AgreementStatus { get; set; } = "";
    public long? SourceDocumentId { get; set; }
    public int? SourcePageNumber { get; set; }
    public string ReviewStatus { get; set; } = FlexibleBenefitReviewStatuses.Pending;
}

public sealed class FlexibleBenefitContact
{
    public long? FlexibleBenefitContactId { get; set; }
    public long? FlexibleBenefitPlanId { get; set; }
    public long? ClientContactId { get; set; }
    public string SourceKey { get; set; } = Guid.NewGuid().ToString("N");
    public string ContactRole { get; set; } = "";
    public string ContactName { get; set; } = "";
    public string Email { get; set; } = "";
    public string Phone { get; set; } = "";
    public bool IsPrimary { get; set; }
    public int? SourcePageNumber { get; set; }
    public decimal? ExtractionConfidence { get; set; }
    public string ReviewStatus { get; set; } = FlexibleBenefitReviewStatuses.Pending;
}

public sealed class FlexibleBenefitPlanSourceField
{
    public long? FlexibleBenefitPlanSourceFieldId { get; set; }
    public long? SourceDocumentId { get; set; }
    public long? SourceDocumentVersionId { get; set; }
    public string EntityType { get; set; } = "";
    public long? EntityId { get; set; }
    public string EntityKey { get; set; } = "";
    public string FieldName { get; set; } = "";
    public string RawValue { get; set; } = "";
    public string NormalizedValue { get; set; } = "";
    public int? SourcePageNumber { get; set; }
    public string BoundingBoxJson { get; set; } = "";
    public string ExtractionMethod { get; set; } = "";
    public decimal? ExtractionConfidence { get; set; }
    public bool WasManuallyCorrected { get; set; }
    public string CorrectedBy { get; set; } = "";
    public DateTime? CorrectedAt { get; set; }
    public string ReviewStatus { get; set; } = FlexibleBenefitReviewStatuses.Pending;
}

public sealed class FlexibleBenefitDocumentVersion
{
    public long? FlexibleBenefitDocumentVersionId { get; set; }
    public long DocumentId { get; set; }
    public long ClientId { get; set; }
    public string DocumentHash { get; set; } = "";
    public FlexibleBenefitDocumentSubtype DocumentSubtype { get; set; } = FlexibleBenefitDocumentSubtype.Unknown;
    public DateOnly? EffectiveDate { get; set; }
    public DateTime? ExecutionDate { get; set; }
    public int VersionNumber { get; set; } = 1;
    public string VersionAction { get; set; } = FlexibleBenefitImportMode.ImportAsNewVersion.ToString();
    public bool IsCurrent { get; set; } = true;
    public long? SupersedesDocumentVersionId { get; set; }
    public DateTime CreatedAt { get; set; } = DateTime.Now;
    public string CreatedBy { get; set; } = Environment.UserName;
}

public readonly record struct FlexibleBenefitBounds(double X, double Y, double Width, double Height)
{
    public double Right => X + Width;
    public double Bottom => Y + Height;
}

public sealed class FlexibleBenefitOcrWord
{
    public string Text { get; init; } = "";
    public FlexibleBenefitBounds Bounds { get; init; }
}

public sealed class FlexibleBenefitOcrLine
{
    public string Text { get; init; } = "";
    public FlexibleBenefitBounds Bounds { get; init; }
    public List<FlexibleBenefitOcrWord> Words { get; } = [];
}

public sealed class FlexibleBenefitDetectedMark
{
    public FlexibleBenefitBounds Bounds { get; init; }
    public bool? IsSelected { get; init; }
    public string Shape { get; init; } = "Checkbox";
    public decimal Confidence { get; init; }
}

public sealed class FlexibleBenefitOcrPage
{
    public int PageNumber { get; init; }
    public int PixelWidth { get; init; }
    public int PixelHeight { get; init; }
    public string Text { get; init; } = "";
    public List<FlexibleBenefitOcrLine> Lines { get; } = [];
    public List<FlexibleBenefitDetectedMark> Marks { get; } = [];
}

public sealed class FlexibleBenefitOcrDocument
{
    public string FilePath { get; init; } = "";
    public string FileHash { get; init; } = "";
    public string ExtractionMethod { get; init; } = "Windows Media OCR";
    public decimal Confidence { get; init; } = 0.80m;
    public List<FlexibleBenefitOcrPage> Pages { get; } = [];
    public string CombinedText => string.Join(Environment.NewLine, Pages.Select(page => page.Text));
}

public sealed class FlexibleBenefitDocumentClassification
{
    public string Carrier { get; init; } = "";
    public string Administrator { get; init; } = "";
    public string DocumentType { get; init; } = FlexibleBenefitDocumentTypes.FlexibleBenefitsPlanDocument;
    public FlexibleBenefitDocumentSubtype DocumentSubtype { get; init; } = FlexibleBenefitDocumentSubtype.Unknown;
    public bool? IsSigned { get; init; }
    public bool? IsAmendment { get; init; }
    public decimal Confidence { get; init; }
}

public sealed class FlexibleBenefitClientIdentity
{
    public string LegalName { get; set; } = "";
    public string DbaName { get; set; } = "";
    public string Address { get; set; } = "";
    public string City { get; set; } = "";
    public string Province { get; set; } = "";
    public string PostalCode { get; set; } = "";
    public string Phone { get; set; } = "";
}

public sealed class FlexibleBenefitValidationIssue
{
    public string Code { get; init; } = "";
    public string Severity { get; init; } = "Warning";
    public string FieldName { get; init; } = "";
    public string Message { get; init; } = "";
    public int? SourcePageNumber { get; init; }
    public bool IsFailure => Severity.Equals("Error", StringComparison.OrdinalIgnoreCase);
}

public sealed class FlexibleBenefitValidationResult
{
    public List<FlexibleBenefitValidationIssue> Issues { get; } = [];
    public string Status { get; set; } = FlexibleBenefitExtractionStatuses.Failed;
    public bool CanCommit => !Issues.Any(issue => issue.IsFailure);
    public bool HasWarnings => Issues.Any(issue => issue.Severity.Equals("Warning", StringComparison.OrdinalIgnoreCase));
}

public sealed class FlexibleBenefitExtractionResult
{
    public FlexibleBenefitDocumentClassification Classification { get; set; } = new();
    public FlexibleBenefitClientIdentity Client { get; set; } = new();
    public FlexibleBenefitPlan Plan { get; set; } = new();
    public List<FlexibleBenefitClass> Classes { get; } = [];
    public List<FlexibleBenefitCategory> Categories { get; } = [];
    public List<FlexibleBenefitEligibilityRule> EligibilityRules { get; } = [];
    public List<FlexibleBenefitFundingConfiguration> FundingConfigurations { get; } = [];
    public List<FlexibleBenefitFee> Fees { get; } = [];
    public List<FlexibleBenefitAgreement> Agreements { get; } = [];
    public List<FlexibleBenefitContact> Contacts { get; } = [];
    public List<FlexibleBenefitPlanSourceField> SourceFields { get; } = [];
    public List<FlexibleBenefitValidationIssue> ValidationIssues { get; } = [];
    public FlexibleBenefitOcrDocument? OcrDocument { get; set; }
    public string SourcePath { get; set; } = "";
    public string SourceHash { get; set; } = "";
    public decimal Confidence { get; set; }
    public string Status { get; set; } = FlexibleBenefitExtractionStatuses.Extracted;
}

public sealed class FlexibleBenefitImportRequest
{
    public long? ClientId { get; init; }
    public string ClientName { get; init; } = "";
    public string FilePath { get; init; } = "";
    public FlexibleBenefitImportMode ImportMode { get; init; } = FlexibleBenefitImportMode.Cancel;
    public string ImportedBy { get; init; } = Environment.UserName;
}

public sealed class FlexibleBenefitImportResult
{
    public FlexibleBenefitImportOutcome Outcome { get; init; }
    public string Message { get; init; } = "";
    public long? ClientId { get; init; }
    public long? DocumentId { get; init; }
    public long? DocumentVersionId { get; init; }
    public long? FlexibleBenefitPlanId { get; init; }
    public int VersionNumber { get; init; }
    public string Status { get; init; } = ImportStatuses.Failed;
    public FlexibleBenefitExtractionResult? Extraction { get; init; }
    public List<FlexibleBenefitValidationIssue> ValidationIssues { get; } = [];
    public bool Succeeded => Outcome is FlexibleBenefitImportOutcome.Completed or FlexibleBenefitImportOutcome.NeedsReview;
}

public sealed class FlexibleBenefitCategoryDefinition
{
    public string CategoryCode { get; init; } = "";
    public string CategoryGroup { get; init; } = "";
    public string CategoryName { get; init; } = "";
    public string DocumentWording { get; init; } = "";
    public List<string> Aliases { get; init; } = [];
    public bool? Taxable { get; init; }
    public FlexibleBenefitAccountType AccountType { get; init; } = FlexibleBenefitAccountType.Lsa;
    public bool? RequiresReceipt { get; init; }
    public bool? RequiresProviderDetails { get; init; }
    public string Exclusions { get; init; } = "";
}

public sealed class FlexibleBenefitWorkspaceViewModel
{
    public bool HasPlan { get; init; }
    public long? PlanId { get; init; }
    public long? SourceDocumentId { get; init; }
    public long? SourceDocumentVersionId { get; init; }
    public string Administrator { get; init; } = "Not confirmed";
    public string PlanType { get; init; } = "Not confirmed";
    public string EffectiveDate { get; init; } = "Not confirmed";
    public string HcsaSummary { get; init; } = "Not confirmed";
    public string LsaSummary { get; init; } = "Not confirmed";
    public string DefaultAllocation { get; init; } = "Not confirmed";
    public string BenefitYear { get; init; } = "Not confirmed";
    public string Runoff { get; init; } = "Not confirmed";
    public string CarryForward { get; init; } = "Not confirmed";
    public string DependentAges { get; init; } = "Not confirmed";
    public string FundingSummary { get; init; } = "Not confirmed";
    public string FeeSummary { get; init; } = "Not confirmed";
    public string AgreementSummary { get; set; } = "Not confirmed";
    public string ReviewStatus { get; set; } = "Not confirmed";
    public string SourceDocumentPath { get; init; } = "";
    public int VersionNumber { get; init; }
    public Visibility ContentVisibility => HasPlan ? Visibility.Visible : Visibility.Collapsed;
    public Visibility EmptyVisibility => HasPlan ? Visibility.Collapsed : Visibility.Visible;
    public Visibility ReviewItemsVisibility => ReviewItems.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
    public bool CanComparePriorVersion => Versions.Count > 1;
    public ObservableCollection<FlexibleBenefitClassDisplayRow> Classes { get; } = [];
    public ObservableCollection<FlexibleBenefitCategoryDisplayRow> Categories { get; } = [];
    public ObservableCollection<FlexibleBenefitDetailDisplayRow> Funding { get; } = [];
    public ObservableCollection<FlexibleBenefitDetailDisplayRow> Fees { get; } = [];
    public ObservableCollection<FlexibleBenefitDetailDisplayRow> Agreements { get; } = [];
    public ObservableCollection<FlexibleBenefitDetailDisplayRow> Contacts { get; } = [];
    public ObservableCollection<FlexibleBenefitDetailDisplayRow> ReviewItems { get; } = [];
    public ObservableCollection<FlexibleBenefitSourceFieldDisplayRow> SourceFields { get; } = [];
    public ObservableCollection<FlexibleBenefitVersionDisplayRow> Versions { get; } = [];
}

public sealed class FlexibleBenefitClassDisplayRow
{
    public string ClassName { get; init; } = "";
    public string SingleAmount { get; init; } = "Not confirmed";
    public string CoupleAmount { get; init; } = "Not confirmed";
    public string FamilyAmount { get; init; } = "Not confirmed";
    public string WaitingPeriod { get; init; } = "Not confirmed";
    public string Proration { get; init; } = "Not confirmed";
    public string SourcePage { get; init; } = "";
}

public sealed class FlexibleBenefitCategoryDisplayRow
{
    public string CategoryGroup { get; init; } = "";
    public string CategoryName { get; init; } = "";
    public string Enabled { get; init; } = "Not confirmed";
    public string SourcePage { get; init; } = "";
}

public sealed class FlexibleBenefitDetailDisplayRow
{
    public string Label { get; init; } = "";
    public string Value { get; init; } = "Not confirmed";
    public string SourcePage { get; init; } = "";
}

public sealed class FlexibleBenefitSourceFieldDisplayRow
{
    public long SourceFieldId { get; init; }
    public string EntityType { get; init; } = "";
    public string FieldName { get; init; } = "";
    public string RawValue { get; init; } = "";
    public string NormalizedValue { get; init; } = "";
    public string SourcePage { get; init; } = "";
    public string Confidence { get; init; } = "";
    public string ReviewStatus { get; init; } = "";
    public bool WasManuallyCorrected { get; init; }
}

public sealed class FlexibleBenefitVersionDisplayRow
{
    public long PlanId { get; init; }
    public int VersionNumber { get; init; }
    public string EffectiveDate { get; init; } = "";
    public string Status { get; init; } = "";
    public string SourceFile { get; init; } = "";
    public bool IsCurrent { get; init; }
}

public sealed class FlexibleBenefitVersionDifferenceRow
{
    public string Area { get; init; } = "";
    public string Field { get; init; } = "";
    public string PriorValue { get; init; } = "Not confirmed";
    public string CurrentValue { get; init; } = "Not confirmed";
}

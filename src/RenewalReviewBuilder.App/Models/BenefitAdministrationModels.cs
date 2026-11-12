using System.Collections.ObjectModel;
using System.Windows;

namespace RenewalReviewBuilder.App.Models;

public static class BenefitAdministrationReviewStatuses
{
    public const string NeedsReview = "NEEDS_REVIEW";
    public const string Confirmed = "CONFIRMED";
    public const string Rejected = "REJECTED";
}

public static class BenefitAdministrationOrganizationTypes
{
    public const string Insurer = "INSURER";
    public const string Carrier = "CARRIER";
    public const string ThirdPartyAdministrator = "THIRD_PARTY_ADMINISTRATOR";
    public const string DisabilityManager = "DISABILITY_MANAGER";
    public const string AssistanceProvider = "ASSISTANCE_PROVIDER";
    public const string VirtualCareProvider = "VIRTUAL_CARE_PROVIDER";
    public const string SecondOpinionProvider = "SECOND_OPINION_PROVIDER";
    public const string OtherServiceProvider = "OTHER_SERVICE_PROVIDER";
}

public static class BenefitAdministrationProviderRoles
{
    public const string BenefitAdministrator = "BENEFIT_ADMINISTRATOR";
    public const string ClaimsAdministrator = "CLAIMS_ADMINISTRATOR";
    public const string DisabilityManagement = "DISABILITY_MANAGEMENT";
    public const string TravelAssistance = "TRAVEL_ASSISTANCE";
    public const string MedicalSecondOpinion = "MEDICAL_SECOND_OPINION";
    public const string VirtualCare = "VIRTUAL_CARE";
    public const string OptionalBenefitAdministrator = "OPTIONAL_BENEFIT_ADMINISTRATOR";
    public const string Other = "OTHER";
}

public static class CanonicalBenefitTypeCodes
{
    public const string Life = "LIFE";
    public const string AdAndD = "AD_AND_D";
    public const string Ltd = "LTD";
    public const string Ehc = "EHC";
    public const string Travel = "TRAVEL";
    public const string MedicalSecondOpinion = "MEDICAL_SECOND_OPINION";
    public const string VirtualHealthcare = "VIRTUAL_HEALTHCARE";
    public const string Dental = "DENTAL";
    public const string OptionalLife = "OPTIONAL_LIFE";
    public const string OptionalAccidentSeriousIllness = "OPTIONAL_ACCIDENT_SERIOUS_ILLNESS";
    public const string OptionalCriticalIllness = "OPTIONAL_CRITICAL_ILLNESS";
}

public sealed class BenefitAdministrationExtractionCandidate
{
    public string SourceFile { get; init; } = "";
    public string ClientName { get; init; } = "";
    public string BenefitDisplayName { get; init; } = "";
    public string InsurerDisplayName { get; init; } = "";
    public List<string> ServiceProviderDisplayNames { get; init; } = [];
    public string PolicyNumber { get; init; } = "";
    public string ProductName { get; init; } = "";
    public bool IsOptional { get; init; }
    public int? SourcePageNumber { get; init; }
    public string SourceText { get; init; } = "";
    public decimal? ExtractionConfidence { get; init; }
    public long? SourceDocumentId { get; set; }
}

public sealed class OrganizationRecord
{
    public long OrganizationId { get; set; }
    public string CanonicalName { get; set; } = "";
    public string DisplayName { get; set; } = "";
    public string OrganizationTypeCode { get; set; } = "";
    public string NormalizedName { get; set; } = "";
    public bool IsActive { get; set; } = true;
}

public sealed class ClientBenefitAdministrationRecord
{
    public long ClientBenefitAdministrationId { get; set; }
    public long ClientId { get; set; }
    public long? PlanDesignVersionId { get; set; }
    public long BenefitTypeId { get; set; }
    public string BenefitTypeCode { get; set; } = "";
    public string BenefitDisplayName { get; set; } = "";
    public string ProductName { get; set; } = "";
    public string PolicyNumber { get; set; } = "";
    public long? InsurerOrganizationId { get; set; }
    public string InsurerName { get; set; } = "";
    public string InsurerOrganizationTypeCode { get; set; } = "";
    public bool IsOptional { get; set; }
    public DateTime? EffectiveDate { get; set; }
    public DateTime? TerminationDate { get; set; }
    public long? SourceDocumentId { get; set; }
    public string SourceDocumentName { get; set; } = "";
    public string SourceDocumentPath { get; set; } = "";
    public int? SourcePageNumber { get; set; }
    public string SourceText { get; set; } = "";
    public decimal? ExtractionConfidence { get; set; }
    public string ReviewStatusCode { get; set; } = BenefitAdministrationReviewStatuses.NeedsReview;
    public bool IsActive { get; set; } = true;
    public long? CopiedFromClientBenefitAdministrationId { get; set; }
    public DateTime CreatedAtUtc { get; set; }
    public DateTime UpdatedAtUtc { get; set; }
    public List<BenefitServiceProviderAssignmentRecord> ServiceProviders { get; } = [];
}

public sealed class BenefitServiceProviderAssignmentRecord
{
    public long BenefitServiceProviderAssignmentId { get; set; }
    public long ClientBenefitAdministrationId { get; set; }
    public long ProviderOrganizationId { get; set; }
    public string ProviderName { get; set; } = "";
    public string ProviderRoleCode { get; set; } = BenefitAdministrationProviderRoles.Other;
    public string SourceDisplayName { get; set; } = "";
    public bool IsPrimary { get; set; } = true;
    public DateTime? EffectiveDate { get; set; }
    public DateTime? TerminationDate { get; set; }
    public long? SourceDocumentId { get; set; }
    public int? SourcePageNumber { get; set; }
    public string ReviewStatusCode { get; set; } = BenefitAdministrationReviewStatuses.NeedsReview;
    public bool IsActive { get; set; } = true;
    public DateTime CreatedAtUtc { get; set; }
    public DateTime UpdatedAtUtc { get; set; }
}

public sealed class BenefitAdministrationImportRequest
{
    public long ClientId { get; init; }
    public long? PlanDesignVersionId { get; init; }
    public DateTime? EffectiveDate { get; init; }
    public IReadOnlyList<BenefitAdministrationExtractionCandidate> Records { get; init; } = [];
}

public sealed class BenefitAdministrationImportResult
{
    public bool Succeeded { get; init; }
    public int RecordsInserted { get; init; }
    public int ProviderAssignmentsInserted { get; init; }
    public int DuplicatesSkipped { get; init; }
    public List<string> Warnings { get; init; } = [];
    public string Message { get; init; } = "";
}

public sealed class BenefitAdministrationClientExportRow
{
    public string Benefit { get; init; } = "";
    public string InsuranceCompany { get; init; } = "";
    public string ServiceProvider { get; init; } = "";
    public string PolicyNumber { get; init; } = "";
}

public sealed class BenefitAdministrationWorkspaceViewModel
{
    public ObservableCollection<BenefitAdministrationWorkspaceRow> Rows { get; } = [];
    public ObservableCollection<string> InsurerFilters { get; } = ["All insurers"];
    public ObservableCollection<string> ProviderFilters { get; } = ["All providers"];
    public ObservableCollection<string> StatusFilters { get; } = ["All statuses"];
    public ObservableCollection<string> OptionalFilters { get; } = ["All benefits", "Mandatory", "Optional"];
    public bool HasRecords => Rows.Count > 0;
    public Visibility ContentVisibility => HasRecords ? Visibility.Visible : Visibility.Collapsed;
    public Visibility EmptyVisibility => HasRecords ? Visibility.Collapsed : Visibility.Visible;
}

public sealed class BenefitAdministrationWorkspaceRow
{
    public long ClientBenefitAdministrationId { get; init; }
    public string Benefit { get; init; } = "";
    public string Product { get; init; } = "";
    public string InsuranceCompany { get; init; } = "";
    public string ServiceProvider { get; init; } = "";
    public string ProviderRole { get; init; } = "";
    public string PolicyNumber { get; init; } = "";
    public string EffectiveDate { get; init; } = "";
    public string Status { get; init; } = "";
    public string Source { get; init; } = "";
    public string SourcePath { get; init; } = "";
    public long? SourceDocumentId { get; init; }
    public int? SourcePageNumber { get; init; }
    public bool IsOptional { get; init; }
    public IReadOnlyList<BenefitServiceProviderAssignmentRecord> ProviderAssignments { get; init; } = [];
    public string OptionalText => IsOptional ? "Optional" : "";
    public Visibility OptionalVisibility => IsOptional ? Visibility.Visible : Visibility.Collapsed;
    public bool NeedsReview => Status.Equals(BenefitAdministrationReviewStatuses.NeedsReview, StringComparison.OrdinalIgnoreCase);
}

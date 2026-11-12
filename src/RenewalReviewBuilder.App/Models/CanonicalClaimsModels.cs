namespace RenewalReviewBuilder.App.Models;

public enum CanonicalDataAvailability
{
    ReportedZero,
    Present,
    Missing,
    NotAvailable,
    NotApplicable,
    ParseFailure,
    Suppressed
}

public enum CanonicalClaimsValidationStatus { Valid, NeedsReview, Invalid }
public enum CanonicalClaimsRelationshipType { Unknown, Employee, Dependant, Partner, Child, Spouse, Retiree, Other, OtherDependant, Unallocated, Total }
public enum CanonicalClaimsSourceAggregationLevel { Sponsor, Contract, Division, Class, Province, Service, Relationship, Certificate }
public enum CanonicalClaimsEvidencePrecision { Document, Page, Worksheet, Section, Row, Cell }
public enum CanonicalClaimsMappingClassification { Direct, Normalized, Derived, Unavailable, Ambiguous }
public enum CanonicalClaimsParityClassification
{
    ExactMatch,
    WithinRoundingTolerance,
    ExpectedModellingDifference,
    MissingLegacyField,
    MissingCanonicalField,
    SourceScopeDifference,
    ValidationFailure,
    UnexpectedMismatch
}

// Import DTO: the only payload accepted by the canonical write boundary.
public sealed class CanonicalClaimsImportEnvelope
{
    public string MappingVersion { get; init; } = "1.0";
    public CanonicalClaimsReport Report { get; init; } = new();
    public List<CanonicalClaimFinancialFact> FinancialFacts { get; } = [];
    public List<CanonicalClaimMonthlyFact> MonthlyFacts { get; } = [];
    public List<CanonicalClaimCategoryFact> CategoryFacts { get; } = [];
    public List<CanonicalClaimDistributionFact> DistributionFacts { get; } = [];
    public List<CanonicalClaimsUtilizationFact> UtilizationFacts { get; } = [];
    public List<CanonicalBenchmarkFact> BenchmarkFacts { get; } = [];
    public List<CanonicalDrugClaim> DrugClaims { get; } = [];
    public List<CanonicalDrugIndicationFact> DrugIndications { get; } = [];
    public List<CanonicalProviderUtilizationFact> ProviderUtilization { get; } = [];
    public List<CanonicalLargeClaimFact> LargeClaims { get; } = [];
    public List<CanonicalDisabilityClaimFact> DisabilityClaims { get; } = [];
    public List<CanonicalLifeClaimFact> LifeClaims { get; } = [];
    public List<CanonicalClaimsFactRevision> FactRevisions { get; } = [];
    public List<CanonicalClaimsValidationResult> MappingWarnings { get; } = [];
}

// Domain aggregate root: report ownership and identity shared by all child facts.
public sealed class CanonicalClaimsReport
{
    public long ClientId { get; init; }
    public long? CarrierId { get; init; }
    public string Carrier { get; init; } = "";
    public string PolicyNumber { get; init; } = "";
    public string Division { get; init; } = "";
    public string Class { get; init; } = "";
    public long? ImportBatchId { get; init; }
    public long? SourceDocumentId { get; init; }
    public long? SourceDocumentVersionId { get; init; }
    public CanonicalClaimsPeriod Period { get; init; } = new();
    public DateTime CreatedAt { get; init; } = DateTime.UtcNow;
    public DateTime UpdatedAt { get; init; } = DateTime.UtcNow;
}

// Value object: normalized reporting period.
public sealed class CanonicalClaimsPeriod
{
    public DateOnly? Start { get; init; }
    public DateOnly? End { get; init; }
    public string PeriodType { get; init; } = "Experience";
}

// Domain model shared by stable and specialized canonical facts.
public abstract class CanonicalClaimFact
{
    public string FactId { get; init; } = Guid.NewGuid().ToString("N");
    public string Benefit { get; init; } = "";
    public string Coverage { get; init; } = "";
    public string Category { get; init; } = "";
    public string Division { get; init; } = "";
    public string Class { get; init; } = "";
    public string MeasurementType { get; init; } = "";
    public string DimensionType { get; init; } = "";
    public string DimensionValue { get; init; } = "";
    public decimal? MonetaryValue { get; init; }
    public long? CountValue { get; init; }
    public decimal? PercentageValue { get; init; }
    public string Currency { get; init; } = "CAD";
    public string UnitOfMeasure { get; init; } = "";
    public CanonicalDataAvailability Availability { get; init; } = CanonicalDataAvailability.Missing;
    public CanonicalClaimsValidationStatus ValidationStatus { get; init; } = CanonicalClaimsValidationStatus.Valid;
    public long? LegacySourceRecordId { get; init; }
    public CanonicalClaimsSourceEvidence Evidence { get; init; } = new();
    public DateTime CreatedAt { get; init; } = DateTime.UtcNow;
    public DateTime UpdatedAt { get; init; } = DateTime.UtcNow;
}

public sealed class CanonicalClaimFinancialFact : CanonicalClaimFact
{
    public decimal? Premium { get; init; }
    public decimal? PaidClaims { get; init; }
    public decimal? IncurredClaims { get; init; }
    public decimal? LossRatio { get; init; }
    public int? Lives { get; init; }
}

public sealed class CanonicalClaimMonthlyFact : CanonicalClaimFact
{
    public DateOnly? Month { get; init; }
    public decimal? Premium { get; init; }
    public decimal? PaidClaims { get; init; }
    public decimal? IncurredClaims { get; init; }
    public decimal? ClaimsPlusFees { get; init; }
    public decimal? LossRatio { get; init; }
    public int? Lives { get; init; }
}

public sealed class CanonicalClaimCategoryFact : CanonicalClaimFact { }
/// <summary>
/// Service-utilization observation. Unique claimants are semi-additive and must never be summed across
/// service, category, province, relationship, division, class, or aggregation levels unless a source total
/// explicitly supplies the distinct count at that exact grain.
/// </summary>
public sealed class CanonicalClaimsUtilizationFact
{
    public string CanonicalClaimsUtilizationFactId { get; init; } = Guid.NewGuid().ToString("N");
    public long? CanonicalClaimsReportId { get; init; }
    public string BenefitCode { get; init; } = "";
    public string BenefitName { get; init; } = "";
    public string CategoryCode { get; init; } = "";
    public string CategoryName { get; init; } = "";
    public string ServiceCode { get; init; } = "";
    public string ServiceName { get; init; } = "";
    public string ProvinceCode { get; init; } = "";
    public CanonicalClaimsRelationshipType RelationshipType { get; init; } = CanonicalClaimsRelationshipType.Unknown;
    public string? DivisionCode { get; init; }
    public string? ClassCode { get; init; }
    public long? UniqueClaimants { get; init; }
    public long? Occurrences { get; init; }
    public decimal? SubmittedAmount { get; init; }
    public decimal? EligibleAmount { get; init; }
    public decimal? PaidAmount { get; init; }
    public decimal? ReportedPaidPercentage { get; init; }
    public string Currency { get; init; } = "CAD";
    public CanonicalClaimsValidationStatus ValidationStatus { get; init; } = CanonicalClaimsValidationStatus.Valid;
    public DateOnly? MeasurementStartDate { get; init; }
    public DateOnly? MeasurementEndDate { get; init; }
    public CanonicalClaimsSourceAggregationLevel SourceAggregationLevel { get; init; } = CanonicalClaimsSourceAggregationLevel.Service;
    public string MappingVersion { get; init; } = "1.0";
    public string RecordStatus { get; init; } = "Active";
    public DateTime CreatedAtUtc { get; init; } = DateTime.UtcNow;
    public CanonicalClaimsSourceEvidence Evidence { get; init; } = new();
}

public sealed class CanonicalClaimsFactRevision
{
    public string RevisionId { get; init; } = Guid.NewGuid().ToString("N");
    public string FactId { get; init; } = "";
    public string NaturalKey { get; init; } = "";
    public string PreviousSourceReference { get; init; } = "";
    public string RevisedSourceReference { get; init; } = "";
    public decimal? PreviousSubmittedAmount { get; init; }
    public decimal? RevisedSubmittedAmount { get; init; }
    public decimal? PreviousEligibleAmount { get; init; }
    public decimal? RevisedEligibleAmount { get; init; }
    public decimal? PreviousPaidAmount { get; init; }
    public decimal? RevisedPaidAmount { get; init; }
    public string ReviewStatus { get; init; } = "NeedsReview";
    public DateTime CreatedAtUtc { get; init; } = DateTime.UtcNow;
}
public sealed record CanonicalClaimsUtilizationMetrics(decimal? OccurrencesPerClaimant, decimal? PaidPerClaimant, decimal? PaidPerOccurrence, decimal? EligiblePerOccurrence, decimal? SubmittedToEligibleRatio, decimal? EligibleToPaidRatio);
public sealed record CanonicalClaimsUtilizationDiagnostic(int FactCount, int BenefitCount, int ServiceCount, int EvidenceCount, long TotalOccurrences, decimal TotalSubmitted, decimal TotalEligible, decimal TotalPaid, int ValidationIssueCount);
public static class CanonicalClaimsUtilizationMetricCalculator
{
    public static CanonicalClaimsUtilizationMetrics Calculate(CanonicalClaimsUtilizationFact fact) => new(
        Divide(fact.Occurrences, fact.UniqueClaimants), Divide(fact.PaidAmount, fact.UniqueClaimants), Divide(fact.PaidAmount, fact.Occurrences), Divide(fact.EligibleAmount, fact.Occurrences), Divide(fact.EligibleAmount, fact.SubmittedAmount), Divide(fact.PaidAmount, fact.EligibleAmount));
    private static decimal? Divide(long? numerator, long? denominator) => numerator.HasValue && denominator is > 0 ? decimal.Round((decimal)numerator.Value / denominator.Value, 6) : null;
    private static decimal? Divide(decimal? numerator, long? denominator) => numerator.HasValue && denominator is > 0 ? decimal.Round(numerator.Value / denominator.Value, 6) : null;
    private static decimal? Divide(decimal? numerator, decimal? denominator) => numerator.HasValue && denominator is > 0m ? decimal.Round(numerator.Value / denominator.Value, 6) : null;
    public static CanonicalClaimsUtilizationDiagnostic Summarize(IEnumerable<CanonicalClaimsUtilizationFact> facts, int validationIssueCount)
    {
        var items=facts.ToList(); return new(items.Count, items.Select(f=>f.BenefitCode).Where(x=>x.Length>0).Distinct().Count(), items.Select(f=>f.ServiceCode).Where(x=>x.Length>0).Distinct().Count(), items.Count(f=>f.Evidence.SourceDocumentId.HasValue), items.Sum(f=>f.Occurrences??0), items.Sum(f=>f.SubmittedAmount??0m), items.Sum(f=>f.EligibleAmount??0m), items.Sum(f=>f.PaidAmount??0m), validationIssueCount);
    }
}
public sealed class CanonicalClaimDistributionFact : CanonicalClaimFact { public decimal? AverageAmount { get; init; } }
public sealed class CanonicalBenchmarkFact : CanonicalClaimFact { public string BenchmarkDefinition { get; init; } = ""; }
public sealed class CanonicalDrugClaim : CanonicalClaimFact { public string DrugName { get; init; } = ""; }
public sealed class CanonicalDrugIndicationFact : CanonicalClaimFact { public string Indication { get; init; } = ""; }
public sealed class CanonicalProviderUtilizationFact : CanonicalClaimFact { public string ProviderType { get; init; } = ""; }
public sealed class CanonicalLargeClaimFact : CanonicalClaimFact { public string ClaimReference { get; init; } = ""; }
public sealed class CanonicalDisabilityClaimFact : CanonicalClaimFact { public string DisabilityStatus { get; init; } = ""; }
public sealed class CanonicalLifeClaimFact : CanonicalClaimFact { public string ClaimType { get; init; } = ""; }

// Value object: source and calculation lineage attached to each fact.
public sealed class CanonicalClaimsSourceEvidence
{
    public long? SourceDocumentId { get; init; }
    public long? SourceDocumentVersionId { get; init; }
    public long? ImportBatchId { get; init; }
    public int? SourcePageNumber { get; init; }
    public string WorksheetName { get; init; } = "";
    public int? SourceRow { get; init; }
    public string SourceColumnOrCell { get; init; } = "";
    public string SourceTableOrSection { get; init; } = "";
    public string RawCarrierLabel { get; init; } = "";
    public string NormalizedLabel { get; init; } = "";
    public string RawExtractedValue { get; init; } = "";
    public string ExtractionMethod { get; init; } = "";
    public string ParserName { get; init; } = "ClaimsExperienceDataModelBuilder";
    public string ParserVersion { get; init; } = "1.0";
    public string MappingVersion { get; init; } = "1.0";
    public decimal? ConfidenceScore { get; init; }
    public CanonicalClaimsEvidencePrecision Precision { get; init; } = CanonicalClaimsEvidencePrecision.Document;
    public CanonicalClaimsMappingClassification MappingClassification { get; init; } = CanonicalClaimsMappingClassification.Direct;
    public string NormalizationRule { get; init; } = "";
    public string DerivationFormula { get; init; } = "";
    public bool WasNormalized { get; init; }
    public string WarningCode { get; init; } = "";
    public string WarningMessage { get; init; } = "";
}

// Read models returned by validation, parity, persistence, and query boundaries.
public sealed class CanonicalClaimsValidationResult
{
    public string Code { get; init; } = "";
    public string Severity { get; init; } = "Warning";
    public string FieldName { get; init; } = "";
    public string Message { get; init; } = "";
    public string? FactId { get; init; }
    public CanonicalClaimsMappingClassification? MappingClassification { get; init; }
    public string SourceReference { get; init; } = "";
    public string NormalizationRule { get; init; } = "";
}

public sealed class CanonicalClaimsParityResult
{
    public string Metric { get; init; } = "";
    public string Benefit { get; init; } = "";
    public string Period { get; init; } = "";
    public decimal? LegacyValue { get; init; }
    public decimal? CanonicalValue { get; init; }
    public decimal? Variance { get; init; }
    public decimal Tolerance { get; init; } = 0.01m;
    public CanonicalClaimsParityClassification Classification { get; init; }
    public string Explanation { get; init; } = "";
    public string RelatedLegacyRecordIds { get; init; } = "";
    public string RelatedCanonicalFactIds { get; init; } = "";
    public string ReviewStatus { get; init; } = "Not Required";
}

public sealed class CanonicalClaimsImportResult
{
    public bool Succeeded { get; init; }
    public bool WasIdempotentRetry { get; init; }
    public long? CanonicalReportId { get; init; }
    public int PersistedFactCount { get; init; }
    public string Message { get; init; } = "";
    public string CompletionStatus { get; init; } = ImportStatuses.Failed;
    public TimeSpan Duration { get; set; }
    public List<CanonicalClaimsValidationResult> ValidationResults { get; } = [];
    public List<CanonicalClaimsParityResult> ParityResults { get; } = [];
}

public sealed class CanonicalClaimsReportReadModel
{
    public long CanonicalReportId { get; init; }
    public long ClientId { get; init; }
    public string Carrier { get; init; } = "";
    public string PolicyNumber { get; init; } = "";
    public long? SourceDocumentId { get; init; }
    public long? SourceDocumentVersionId { get; init; }
    public CanonicalClaimsPeriod Period { get; init; } = new();
    public string ValidationStatus { get; init; } = "";
    public DateTime CreatedAt { get; init; }
    public DateTime UpdatedAt { get; init; }
    public List<CanonicalClaimFinancialFact> FinancialFacts { get; } = [];
    public List<CanonicalClaimMonthlyFact> MonthlyFacts { get; } = [];
    public List<CanonicalClaimCategoryFact> CategoryFacts { get; } = [];
    public List<CanonicalClaimDistributionFact> DistributionFacts { get; } = [];
    public List<CanonicalClaimsUtilizationFact> UtilizationFacts { get; } = [];
    public List<CanonicalBenchmarkFact> BenchmarkFacts { get; } = [];
    public List<CanonicalClaimsParityResult> ParityResults { get; } = [];
    public bool RequiredParityComplete => ParityResults.Count > 0 && ParityResults.All(item => item.Classification is CanonicalClaimsParityClassification.ExactMatch or CanonicalClaimsParityClassification.WithinRoundingTolerance);
}


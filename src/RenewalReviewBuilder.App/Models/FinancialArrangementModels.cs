namespace RenewalReviewBuilder.App.Models;

public enum FinancialChargeType
{
    GeneralAdministration,
    BenefitPaymentsAdministration,
    ProfitCharge,
    PrintingExpense,
    ReportPreparationExpense,
    TripCancellation,
    PremiumTax,
    InterestTerm,
    OtherCarrierCharge
}

public enum FinancialArrangementScenario
{
    Current,
    InitialRenewal,
    Negotiated,
    FinalAccepted,
    Alternative
}

public enum ChangeClassification
{
    Added,
    Removed,
    Increased,
    Decreased,
    Unchanged,
    BasisChanged,
    StructureChanged,
    NotComparable,
    MissingCurrent,
    MissingProposed,
    RequiresReview
}

public sealed class FinancialArrangement
{
    public long FinancialArrangementId { get; set; }
    public long ClientId { get; set; }
    public string Carrier { get; set; } = "";
    public DateOnly EffectiveDate { get; set; }
    public DateOnly? ExpiryDate { get; set; }
    public string Status { get; set; } = "Extracted";
    public string Source { get; set; } = "";
    public long? SourceDocumentId { get; set; }
    // Contractual terms are stored with their document evidence; activity workbooks stay performance-only.
    public decimal? TargetLossRatio { get; set; }
    public string TlrBasis { get; set; } = "";
    public string TargetLossRatioSourceText { get; set; } = "";
    public int? TargetLossRatioSourcePage { get; set; }
    public decimal? TargetLossRatioExtractionConfidence { get; set; }
    public List<FinancialArrangementCharge> Charges { get; } = [];
}

public sealed class FinancialArrangementCharge
{
    public long FinancialArrangementChargeId { get; set; }
    public FinancialChargeType ChargeType { get; set; }
    public string Benefit { get; set; } = "All";
    public decimal? Rate { get; set; }
    public decimal? Amount { get; set; }
    public string Formula { get; set; } = "";
    public string Basis { get; set; } = "Total benefit payments";
    public string SourceText { get; set; } = "";
    public int? SourcePage { get; set; }
    public decimal? ExtractionConfidence { get; set; }
}

public sealed class FinancialArrangementAuditEvent
{
    public long FinancialArrangementAuditEventId { get; set; }
    public string EntityType { get; set; } = "";
    public long EntityId { get; set; }
    public long ClientId { get; set; }
    public long? RenewalId { get; set; }
    public string EventType { get; set; } = "";
    public string PreviousValue { get; set; } = "";
    public string NewValue { get; set; } = "";
    public DateOnly EffectiveDate { get; set; }
    public string User { get; set; } = "";
    public string Reason { get; set; } = "";
    public string Source { get; set; } = "";
    public DateTime OccurredAtUtc { get; set; }
}

public sealed class PoolingArrangement
{
    public long PoolingArrangementId { get; set; }
    public long ClientId { get; set; }
    public string Carrier { get; set; } = "";
    public string Benefit { get; set; } = "EHC";
    public decimal? Rate { get; set; }
    public decimal? Threshold { get; set; }
    public string Basis { get; set; } = "Billed premium";
    public string PoolingType { get; set; } = "";
    public DateOnly EffectiveDate { get; set; }
    public DateOnly? ExpiryDate { get; set; }
    public long? SourceDocumentId { get; set; }
    public string SourceText { get; set; } = "";
    public int? SourcePage { get; set; }
    public decimal? ExtractionConfidence { get; set; }
}

public sealed class CommissionArrangement
{
    public long CommissionArrangementId { get; set; }
    public long ClientId { get; set; }
    public string Carrier { get; set; } = "";
    public decimal? Rate { get; set; }
    public decimal? Amount { get; set; }
    public string Basis { get; set; } = "Total benefit payments";
    public string Source { get; set; } = "AOR";
    public string ConfirmationStatus { get; set; } = "RequiresReview";
    public DateOnly EffectiveDate { get; set; }
    public DateOnly? ExpiryDate { get; set; }
    public long? SoldProposalId { get; set; }
    public long? SourceDocumentId { get; set; }
}

public sealed class FinancialArrangementVersion
{
    public long FinancialArrangementVersionId { get; set; }
    public long ClientId { get; set; }
    public string Carrier { get; set; } = "";
    public FinancialArrangementScenario Scenario { get; set; }
    public DateOnly EffectiveDate { get; set; }
    public long FinancialArrangementId { get; set; }
    public long? CommissionArrangementId { get; set; }
    public long? PoolingArrangementId { get; set; }
}

/// <summary>
/// Read-only, evidence-bearing view of a saved financial-arrangement version for canonical history.
/// </summary>
public sealed class FinancialArrangementHistoryVersion
{
    public long FinancialArrangementVersionId { get; set; }
    public long ClientId { get; set; }
    public string Carrier { get; set; } = "";
    public FinancialArrangementScenario Scenario { get; set; }
    public DateOnly EffectiveDate { get; set; }
    public long FinancialArrangementId { get; set; }
    public string ArrangementStatus { get; set; } = "";
    public string Source { get; set; } = "";
    public long? SourceDocumentId { get; set; }
    public decimal? TargetLossRatio { get; set; }
    public string TargetLossRatioBasis { get; set; } = "";
    public string TargetLossRatioSourceText { get; set; } = "";
    public int? TargetLossRatioSourcePage { get; set; }
    public decimal? TargetLossRatioConfidence { get; set; }
    public long? CommissionArrangementId { get; set; }
    public string CommissionConfirmationStatus { get; set; } = "";
    public decimal? CommissionRate { get; set; }
    public decimal? CommissionAmount { get; set; }
    public string CommissionBasis { get; set; } = "";
    public string CommissionSource { get; set; } = "";
    public long? CommissionSourceDocumentId { get; set; }
    public long? PoolingArrangementId { get; set; }
    public string PoolingBenefit { get; set; } = "";
    public decimal? PoolingRate { get; set; }
    public decimal? PoolingThreshold { get; set; }
    public string PoolingBasis { get; set; } = "";
    public string PoolingType { get; set; } = "";
    public long? PoolingSourceDocumentId { get; set; }
    public string PoolingSourceText { get; set; } = "";
    public int? PoolingSourcePage { get; set; }
    public decimal? PoolingConfidence { get; set; }
    public List<FinancialArrangementCharge> Charges { get; } = [];
    public DateTime CreatedAtUtc { get; set; }
}

public sealed class RenewalFinancialArrangementSnapshot
{
    public long RenewalFinancialArrangementSnapshotId { get; set; }
    public long RenewalId { get; set; }
    public long ClientId { get; set; }
    public string Carrier { get; set; } = "";
    public FinancialArrangementScenario Scenario { get; set; }
    public string SnapshotType { get; set; } = "Initial";
    public string SnapshotStatus { get; set; } = "Draft";
    public DateTime AsOfUtc { get; set; }
    public long CurrentArrangementVersionId { get; set; }
    public long ProposedArrangementVersionId { get; set; }
    public string SourceVersion { get; set; } = "";
    public bool IsStale { get; set; }
    public string FinalizedBy { get; set; } = "";
    public DateTime? FinalizedAtUtc { get; set; }
}

public sealed class FinancialArrangementExtraction
{
    public string Carrier { get; set; } = "Manulife";
    public DateOnly EffectiveDate { get; set; }
    public decimal? TargetLossRatio { get; set; }
    public string TlrBasis { get; set; } = "";
    public string TargetLossRatioSourceText { get; set; } = "";
    public int? TargetLossRatioSourcePage { get; set; }
    public decimal? TargetLossRatioExtractionConfidence { get; set; }
    public List<FinancialArrangementCharge> Charges { get; } = [];
    public List<PoolingArrangement> Pooling { get; } = [];
    public List<string> Warnings { get; } = [];

    public bool HasIncompleteTargetLossRatioEvidence =>
        !string.IsNullOrWhiteSpace(TargetLossRatioSourceText) &&
        (!TargetLossRatio.HasValue || string.IsNullOrWhiteSpace(TlrBasis));
}

/// <summary>
/// Verified, current financial terms for one client and carrier. This is a read model;
/// the source-of-truth remains the versioned FinancialArrangement and PoolingArrangement rows.
/// </summary>
public sealed class CurrentClientFinancialTerms
{
    public long FinancialArrangementVersionId { get; set; }
    public FinancialArrangementScenario Scenario { get; set; }
    public long ClientId { get; set; }
    public string Carrier { get; set; } = "";
    public DateOnly EffectiveDate { get; set; }
    public string Source { get; set; } = "";
    public long? SourceDocumentId { get; set; }
    public decimal? TargetLossRatio { get; set; }
    public string TlrBasis { get; set; } = "";
    public string TargetLossRatioSourceText { get; set; } = "";
    public int? TargetLossRatioSourcePage { get; set; }
    public decimal? TargetLossRatioExtractionConfidence { get; set; }
    public decimal? PoolingRate { get; set; }
    public decimal? PoolingThreshold { get; set; }
    public string PoolingBasis { get; set; } = "";
    public string PoolingSourceText { get; set; } = "";
    public int? PoolingSourcePage { get; set; }
    public decimal? PoolingExtractionConfidence { get; set; }
}

public sealed class AdministrativeLoadComparison
{
    public string Benefit { get; set; } = "All";
    public decimal? CurrentCarrierAdministrationRate { get; set; }
    public decimal? ProposedCarrierAdministrationRate { get; set; }
    public decimal? CurrentCommissionRate { get; set; }
    public decimal? ProposedCommissionRate { get; set; }
    public string CurrentCommissionBasis { get; set; } = "";
    public string ProposedCommissionBasis { get; set; } = "";
    public decimal? CurrentTotalAdministrativeLoadRate { get; set; }
    public decimal? ProposedTotalAdministrativeLoadRate { get; set; }
    public bool HasCompatibleCurrentBasis { get; set; }
    public bool HasCompatibleProposedBasis { get; set; }
    public List<string> CalculationWarnings { get; } = [];
}

public sealed class RenewalFinancialArrangementComparison
{
    public long RenewalId { get; set; }
    public long ClientId { get; set; }
    public string Carrier { get; set; } = "";
    public FinancialArrangementScenario ProposedScenario { get; set; }
    public List<AdministrativeLoadComparison> AdministrativeLoadComparisons { get; } = [];
    public decimal? CurrentCommissionRate { get; set; }
    public decimal? ProposedCommissionRate { get; set; }
    public string CurrentCommissionBasis { get; set; } = "";
    public string ProposedCommissionBasis { get; set; } = "";
    public decimal? CurrentPoolingRate { get; set; }
    public decimal? ProposedPoolingRate { get; set; }
    public decimal? CurrentPoolingThreshold { get; set; }
    public decimal? ProposedPoolingThreshold { get; set; }
    public bool IsStale { get; set; }
    public string SnapshotStatus { get; set; } = "Draft";
}

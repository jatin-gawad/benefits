namespace RenewalReviewBuilder.App.Models.History;

public enum HistoryValueType { Text, Number, Percentage, Money, Age, Duration, CoverageLimit, Reimbursement, CarrierReference, VendorReference, Boolean, Date, CompositeCoverage }
public enum HistoryValidationStatus { Valid, NeedsReview, Warning, Failed }
public enum HistoryOrigin { CanonicalProjection, ManualEvent, MigratedLegacy }
public enum HistoryMode { Timeline, Evolution, Audit, Relationship, Renewal }

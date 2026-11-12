namespace RenewalReviewBuilder.App.Models;

public static class ProposalIntelligenceAnalysisVersions
{
    public const string Current = "PROPOSAL_INTELLIGENCE_ANALYSIS_V1";
}

public static class ProposalIntelligenceComparisonStatuses
{
    public const string Same = "Same";
    public const string Different = "Different";
    public const string Incomplete = "Incomplete";
    public const string NotComparable = "Not Comparable";
    public const string NeedsReview = "Needs Review";
}

public static class ProposalIntelligenceFindingCategories
{
    public const string Readiness = "Readiness";
    public const string Financial = "Financial";
    public const string PlanDesign = "Plan Design";
    public const string Underwriting = "Underwriting";
    public const string Validation = "Validation";
    public const string Evidence = "Evidence";
}

public sealed class ProposalIntelligenceAnalysis
{
    public string AnalysisVersion { get; init; } = ProposalIntelligenceAnalysisVersions.Current;
    public string ProfileVersion { get; init; } = ProposalCanonicalProfileVersions.Current;
    public long ClientId { get; init; }
    public string ClientName { get; init; } = "";
    public string ClientNumber { get; init; } = "";
    public DateTime BuiltAtUtc { get; init; } = DateTime.UtcNow;
    public string ReadinessStatus { get; set; } = ProposalCanonicalReadinessStatuses.NotReady;
    public string ReadinessMessage { get; set; } = "No usable saved proposal facts were found for this client.";
    public bool CanCompareCarriers { get; init; }
    public bool CanCompareFinancials { get; init; }
    public bool CanComparePlanDesign { get; init; }
    public List<ProposalIntelligenceQuoteAnalysis> Quotes { get; } = [];
    public List<ProposalIntelligenceRateComparisonRow> RateComparisons { get; } = [];
    public List<ProposalIntelligencePremiumComparisonRow> PremiumComparisons { get; } = [];
    public List<ProposalIntelligencePlanDesignComparisonRow> PlanDesignComparisons { get; } = [];
    public List<ProposalIntelligenceTermComparisonRow> UnderwritingComparisons { get; } = [];
    public List<ProposalIntelligenceFinding> Findings { get; } = [];

    public bool IsReadOnly => true;
    public bool HasProposals => Quotes.Count > 0;
    public bool HasComparativeAnalysis => CanCompareCarriers &&
                                          (RateComparisons.Any(row => row.IsComparable) ||
                                           PremiumComparisons.Any(row => row.IsComparable) ||
                                           PlanDesignComparisons.Any(row => row.IsComparable) ||
                                           UnderwritingComparisons.Any(row => row.IsComparable));
    public int ReviewFindingCount => Findings.Count(finding =>
        finding.Severity == ProposalCanonicalIssueSeverities.Error ||
        finding.Severity == ProposalCanonicalIssueSeverities.Warning);
    public decimal? LowestAnnualPremium => Quotes
        .Where(quote => quote.AnnualPremium.HasValue)
        .Select(quote => quote.AnnualPremium)
        .Min();
    public decimal? HighestAnnualPremium => Quotes
        .Where(quote => quote.AnnualPremium.HasValue)
        .Select(quote => quote.AnnualPremium)
        .Max();
    public decimal? AnnualPremiumSpread => LowestAnnualPremium.HasValue && HighestAnnualPremium.HasValue
        ? HighestAnnualPremium.Value - LowestAnnualPremium.Value
        : null;
    public decimal? AnnualPremiumSpreadPercent => LowestAnnualPremium is > 0m && AnnualPremiumSpread.HasValue
        ? AnnualPremiumSpread.Value / LowestAnnualPremium.Value * 100m
        : null;
}

public sealed class ProposalIntelligenceQuoteAnalysis
{
    public string QuoteKey { get; init; } = "";
    public long ProposalId { get; init; }
    public long? CarrierId { get; init; }
    public string CarrierKey { get; init; } = "";
    public string CarrierName { get; init; } = "";
    public string DisplayName { get; init; } = "";
    public DateOnly? ProposalDate { get; init; }
    public DateOnly? EffectiveDate { get; init; }
    public string Status { get; init; } = "";
    public string ReadinessStatus { get; init; } = ProposalCanonicalReadinessStatuses.NotReady;
    public decimal? MonthlyPremium { get; init; }
    public decimal? AnnualPremium { get; init; }
    public decimal? EmployerCost { get; init; }
    public decimal? EmployeeCost { get; init; }
    public decimal? TotalCost { get; init; }
    public decimal? AnnualPremiumDifferenceFromLowest { get; set; }
    public decimal? AnnualPremiumDifferencePercentFromLowest { get; set; }
    public bool IsLowestAnnualPremium { get; set; }
    public bool HasCoreFacts { get; init; }
    public bool HasFinancialFacts { get; init; }
    public bool HasPlanDesignFacts { get; init; }
    public bool HasBookletValidation { get; init; }
    public bool HasSourceEvidence { get; init; }
    public int RateCount { get; init; }
    public int PremiumCount { get; init; }
    public int PlanDesignFactCount { get; init; }
    public int ValidationMismatchCount { get; init; }
    public int UnderwritingTermCount { get; init; }
    public int ReviewIssueCount { get; init; }
    public string SourceFileName { get; init; } = "";
    public string SourceFileHash { get; init; } = "";
    public decimal? ExtractionConfidence { get; init; }
    public string ExtractionMethod { get; init; } = "";
    public ProposalCanonicalEvidence Evidence { get; init; } = new();
}

public sealed class ProposalIntelligenceRateComparisonRow
{
    public string ComparisonKey { get; init; } = "";
    public string BenefitLine { get; init; } = "";
    public string ClassName { get; init; } = "";
    public string Tier { get; init; } = "";
    public string AgeBand { get; init; } = "";
    public string Gender { get; init; } = "";
    public string Province { get; init; } = "";
    public string RateBasis { get; init; } = "";
    public string ComparisonStatus { get; set; } = ProposalIntelligenceComparisonStatuses.NotComparable;
    public int ExpectedProposalCount { get; init; }
    public int MissingProposalCount { get; set; }
    public bool HasRateDifference { get; set; }
    public bool HasPremiumDifference { get; set; }
    public List<ProposalIntelligenceRateValue> Values { get; } = [];

    public bool IsComparable => Values.Select(value => value.ProposalId).Distinct().Count() >= 2;
}

public sealed class ProposalIntelligenceRateValue
{
    public string QuoteKey { get; init; } = "";
    public long ProposalId { get; init; }
    public string CarrierName { get; init; } = "";
    public string QuoteDisplayName { get; init; } = "";
    public decimal? Lives { get; init; }
    public decimal? Volume { get; init; }
    public decimal? Rate { get; init; }
    public decimal? MonthlyPremium { get; init; }
    public decimal? AnnualPremium { get; init; }
    public ProposalCanonicalEvidence Evidence { get; init; } = new();
}

public sealed class ProposalIntelligencePremiumComparisonRow
{
    public string ComparisonKey { get; init; } = "";
    public string BenefitLine { get; init; } = "";
    public string ClassName { get; init; } = "";
    public string Tier { get; init; } = "";
    public string ComparisonStatus { get; set; } = ProposalIntelligenceComparisonStatuses.NotComparable;
    public int ExpectedProposalCount { get; init; }
    public int MissingProposalCount { get; set; }
    public bool HasPremiumDifference { get; set; }
    public bool HasCostShareDifference { get; set; }
    public List<ProposalIntelligencePremiumValue> Values { get; } = [];

    public bool IsComparable => Values.Select(value => value.ProposalId).Distinct().Count() >= 2;
}

public sealed class ProposalIntelligencePremiumValue
{
    public string QuoteKey { get; init; } = "";
    public long ProposalId { get; init; }
    public string CarrierName { get; init; } = "";
    public string QuoteDisplayName { get; init; } = "";
    public decimal? MonthlyPremium { get; init; }
    public decimal? AnnualPremium { get; init; }
    public decimal? EmployerCost { get; init; }
    public decimal? EmployeeCost { get; init; }
    public decimal? TotalCost { get; init; }
    public ProposalCanonicalEvidence Evidence { get; init; } = new();
}

public sealed class ProposalIntelligencePlanDesignComparisonRow
{
    public string ComparisonKey { get; init; } = "";
    public string BenefitLine { get; init; } = "";
    public string ProvisionName { get; init; } = "";
    public string ComparisonStatus { get; set; } = ProposalIntelligenceComparisonStatuses.NotComparable;
    public int ExpectedProposalCount { get; init; }
    public int MissingProposalCount { get; set; }
    public bool HasProposalDifference { get; set; }
    public bool HasBookletMismatch { get; set; }
    public List<ProposalIntelligencePlanDesignValue> Values { get; } = [];

    public bool IsComparable => Values.Select(value => value.ProposalId).Distinct().Count() >= 2;
}

public sealed class ProposalIntelligencePlanDesignValue
{
    public string QuoteKey { get; init; } = "";
    public long ProposalId { get; init; }
    public string CarrierName { get; init; } = "";
    public string QuoteDisplayName { get; init; } = "";
    public string ProposalValue { get; init; } = "";
    public string BookletValue { get; init; } = "";
    public string ValidationStatus { get; init; } = "";
    public string Variance { get; init; } = "";
    public bool IsMismatch { get; init; }
    public ProposalCanonicalEvidence Evidence { get; init; } = new();
    public ProposalCanonicalEvidence BookletEvidence { get; init; } = new();
}

public sealed class ProposalIntelligenceTermComparisonRow
{
    public string ComparisonKey { get; init; } = "";
    public string Category { get; init; } = "";
    public string BenefitLine { get; init; } = "";
    public string TermName { get; init; } = "";
    public string ComparisonStatus { get; set; } = ProposalIntelligenceComparisonStatuses.NotComparable;
    public int ExpectedProposalCount { get; init; }
    public int MissingProposalCount { get; set; }
    public bool HasDifference { get; set; }
    public List<ProposalIntelligenceTermValue> Values { get; } = [];

    public bool IsComparable => Values.Select(value => value.ProposalId).Distinct().Count() >= 2;
}

public sealed class ProposalIntelligenceTermValue
{
    public string QuoteKey { get; init; } = "";
    public long ProposalId { get; init; }
    public string CarrierName { get; init; } = "";
    public string QuoteDisplayName { get; init; } = "";
    public string Value { get; init; } = "";
    public ProposalCanonicalEvidence Evidence { get; init; } = new();
}

public sealed class ProposalIntelligenceFinding
{
    public string Code { get; init; } = "";
    public string Severity { get; init; } = ProposalCanonicalIssueSeverities.Warning;
    public string Category { get; init; } = ProposalIntelligenceFindingCategories.Readiness;
    public string Message { get; init; } = "";
    public long? ProposalId { get; init; }
    public string Carrier { get; init; } = "";
    public string BenefitLine { get; init; } = "";
    public string FieldName { get; init; } = "";
    public string SourceTable { get; init; } = "";
    public long? SourceId { get; init; }
    public ProposalCanonicalEvidence? Evidence { get; init; }
}

using System.Globalization;

namespace RenewalReviewBuilder.App.Models;

public static class PlanDesignReadinessStatuses
{
    public const string Ready = "Ready";
    public const string Partial = "Partial";
    public const string NotReady = "Not Ready";
    public const string NeedsReview = "Needs Review";
}

public static class PlanDesignFindingTypes
{
    public const string MissingData = "Missing Data";
    public const string InconsistentData = "Inconsistent Data";
    public const string ComparisonOpportunity = "Comparison Opportunity";
    public const string PricingImpactOpportunity = "Pricing Impact Opportunity";
    public const string NeedsReview = "Needs Review";
    public const string NormalizationSuggestion = "Normalization Suggestion";
}

public static class PlanDesignRiskLevels
{
    public const string Low = "Low";
    public const string Medium = "Medium";
    public const string High = "High";
}

public sealed class PlanDesignIntelligenceClientOption
{
    public long ClientId { get; set; }
    public string ClientName { get; set; } = "";
    public string ClientNumber { get; set; } = "";
    public string Carrier { get; set; } = "";

    public string DisplayName =>
        string.IsNullOrWhiteSpace(Carrier)
            ? ClientName
            : $"{ClientName} ({Carrier})";
}

public sealed class PlanDesignIntelligenceResult
{
    public long ClientId { get; set; }
    public string ClientName { get; set; } = "";
    public string ClientNumber { get; set; } = "";
    public string Carrier { get; set; } = "";
    public DateTime AnalyzedAt { get; set; } = DateTime.Now;
    public List<PlanDesignSummaryRow> SummaryRows { get; } = [];
    public List<PlanDesignReadinessRow> ReadinessRows { get; } = [];
    public List<PlanDesignIntelligenceFinding> Findings { get; } = [];
    public List<PlanDesignSourceTableUsage> SourceTables { get; } = [];

    public IReadOnlyList<PlanDesignIntelligenceFinding> MissingDataFindings =>
        Findings
            .Where(finding => finding.FindingType is PlanDesignFindingTypes.MissingData or PlanDesignFindingTypes.InconsistentData or PlanDesignFindingTypes.NeedsReview)
            .ToList();

    public IReadOnlyList<PlanDesignIntelligenceFinding> NormalizationSuggestions =>
        Findings
            .Where(finding => finding.FindingType == PlanDesignFindingTypes.NormalizationSuggestion)
            .ToList();

    public IReadOnlyList<PlanDesignIntelligenceFinding> PricingImpactOpportunities =>
        Findings
            .Where(finding => finding.FindingType == PlanDesignFindingTypes.PricingImpactOpportunity)
            .ToList();

    public string OverallReadinessStatus
    {
        get
        {
            if (ReadinessRows.Any(row => row.Status == PlanDesignReadinessStatuses.NeedsReview))
            {
                return PlanDesignReadinessStatuses.NeedsReview;
            }

            if (ReadinessRows.Any(row => row.Status == PlanDesignReadinessStatuses.NotReady))
            {
                return PlanDesignReadinessStatuses.NotReady;
            }

            if (ReadinessRows.Any(row => row.Status == PlanDesignReadinessStatuses.Partial))
            {
                return PlanDesignReadinessStatuses.Partial;
            }

            return ReadinessRows.Count == 0 ? PlanDesignReadinessStatuses.NotReady : PlanDesignReadinessStatuses.Ready;
        }
    }
}

public sealed class PlanDesignSummaryRow
{
    public string SourceTable { get; set; } = "";
    public string SourceDocument { get; set; } = "";
    public string SourceDocumentType { get; set; } = "";
    public string Carrier { get; set; } = "";
    public string BenefitCategory { get; set; } = "";
    public string BenefitName { get; set; } = "";
    public int FeatureCount { get; set; }
    public DateTime? LastLoadedDate { get; set; }

    public string LastLoadedText =>
        LastLoadedDate.HasValue
            ? LastLoadedDate.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
            : "";
}

public sealed class PlanDesignReadinessRow
{
    public string Comparison { get; set; } = "";
    public string Status { get; set; } = PlanDesignReadinessStatuses.NotReady;
    public string Reason { get; set; } = "";
}

public sealed class PlanDesignSourceTableUsage
{
    public string TableName { get; set; } = "";
    public bool Exists { get; set; }
    public int RowsFound { get; set; }
    public bool Used { get; set; }
    public string Notes { get; set; } = "";

    public string ExistsText => Exists ? "Yes" : "No";
    public string UsedText => Used ? "Yes" : "No";
}

public sealed class PlanDesignIntelligenceFinding
{
    public long ClientId { get; set; }
    public string Benefit { get; set; } = "";
    public string Category { get; set; } = "";
    public string FindingType { get; set; } = "";
    public string FindingMessage { get; set; } = "";
    public string RiskLevel { get; set; } = PlanDesignRiskLevels.Low;
    public decimal Confidence { get; set; }
    public string SourceTable { get; set; } = "";
    public string SourceId { get; set; } = "";
    public string RecommendedAction { get; set; } = "";

    public string ConfidenceText => Confidence.ToString("0%", CultureInfo.InvariantCulture);
    public string SourceReference => string.IsNullOrWhiteSpace(SourceId) ? SourceTable : $"{SourceTable} #{SourceId}";
}

using System.Globalization;

namespace RenewalReviewBuilder.App.Models;

public sealed class PlanDesignFactorRecord
{
    public const string PlaceholderConfidence = "Estimated Placeholder";
    public const string GeneralPlaceholderCarrier = "General Placeholder";

    public long Id { get; set; }
    public string CarrierName { get; set; } = "";
    public string BenefitCategory { get; set; } = "";
    public string BenefitLine { get; set; } = "";
    public string CurrentDesign { get; set; } = "";
    public string ProposedDesign { get; set; } = "";
    public string ChangeType { get; set; } = "";
    public decimal Factor { get; set; } = 1.0000m;
    public string FactorDirection { get; set; } = "";
    public DateOnly? EffectiveDate { get; set; }
    public DateOnly? ExpiryDate { get; set; }
    public string SourceDocument { get; set; } = "";
    public string Notes { get; set; } = "";
    public string ConfidenceLevel { get; set; } = "";
    public DateTime CreatedAt { get; set; }
    public DateTime UpdatedAt { get; set; }

    public bool IsEstimatedPlaceholder =>
        ConfidenceLevel.Equals(PlaceholderConfidence, StringComparison.OrdinalIgnoreCase);

    public bool IsActiveOn(DateOnly date)
    {
        return (!EffectiveDate.HasValue || EffectiveDate.Value <= date) &&
               (!ExpiryDate.HasValue || ExpiryDate.Value >= date);
    }

    public bool IsExpiringSoon(DateOnly date, int days = 90)
    {
        return ExpiryDate.HasValue &&
               ExpiryDate.Value >= date &&
               ExpiryDate.Value <= date.AddDays(days);
    }

    public string FactorText => Factor.ToString("0.0000", CultureInfo.InvariantCulture);
}

public sealed class PlanDesignFactorFilter
{
    public string SearchTerm { get; set; } = "";
    public string CarrierName { get; set; } = "";
    public string BenefitCategory { get; set; } = "";
    public string BenefitLine { get; set; } = "";
    public string ChangeType { get; set; } = "";
    public string ConfidenceLevel { get; set; } = "";
    public DateOnly? EffectiveDate { get; set; }
}

public sealed class PlanDesignFactorSummary
{
    public int TotalFactors { get; set; }
    public int CarriersCovered { get; set; }
    public int ActiveFactors { get; set; }
    public int EstimatedPlaceholderFactors { get; set; }
    public int ExpiringSoon { get; set; }
}

public sealed class PlanDesignFactorEstimate
{
    public required PlanDesignFactorRecord FactorRecord { get; init; }
    public decimal CurrentPremium { get; init; }
    public decimal EstimatedRevisedPremium { get; init; }
    public decimal DollarChange { get; init; }
    public decimal PercentageChange { get; init; }
    public string FactorSource { get; init; } = "";
    public string ConfidenceLevel { get; init; } = "";
}

public sealed class PlanDesignFactorImportResult
{
    public int InsertedCount { get; set; }
    public int UpdatedCount { get; set; }
    public int SkippedCount { get; set; }
    public List<string> Warnings { get; } = [];
}

public static class PlanDesignFactorReferenceData
{
    public static IReadOnlyList<string> BenefitCategories { get; } =
    [
        "Extended Health Care",
        "Drugs",
        "Dental",
        "Basic Life",
        "AD&D",
        "Short Term Disability",
        "Long Term Disability",
        "Other"
    ];

    public static IReadOnlyList<string> ChangeTypes { get; } =
    [
        "Reimbursement Change",
        "Generic Drug Substitution",
        "Mandatory Generic",
        "Prior Authorization",
        "Biosimilar Switching",
        "Annual Maximum Change",
        "Paramedical Maximum Change",
        "Vision Maximum Change",
        "Dental Recall Change",
        "Add Major Dental",
        "Add Orthodontics",
        "Life Volume Change",
        "LTD Maximum Change",
        "LTD Elimination Period Change",
        "STD Waiting Period Change",
        "Other"
    ];

    public static IReadOnlyList<string> ConfidenceLevels { get; } =
    [
        "Carrier Provided",
        "Underwriter Reviewed",
        "Consultant Estimate",
        PlanDesignFactorRecord.PlaceholderConfidence
    ];

    public static IReadOnlyList<string> Directions { get; } =
    [
        "Increase",
        "Decrease",
        "Neutral"
    ];
}

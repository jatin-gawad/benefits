namespace RenewalReviewBuilder.App.Models;

public enum BenefitRuleType
{
    TieredReimbursement
}

public enum BenefitRuleThresholdBasis
{
    Person,
    Certificate
}

public enum BenefitRuleResetPeriod
{
    CalendarYear,
    PolicyYear
}

public enum BenefitRuleNormalizationStatus
{
    Normalized,
    NeedsReview
}

public sealed class PlanBenefitRule
{
    public long? PlanBenefitRuleId { get; set; }
    public long? PlanDesignFeatureId { get; set; }
    public long? ClientId { get; set; }
    public long? SourceDocumentId { get; set; }
    public BenefitRuleType RuleType { get; set; } = BenefitRuleType.TieredReimbursement;
    public decimal? Tier1ReimbursementPercent { get; set; }
    public decimal? Tier1ThresholdAmount { get; set; }
    public decimal? Tier2ReimbursementPercent { get; set; }
    public BenefitRuleThresholdBasis? ThresholdBasis { get; set; }
    public BenefitRuleResetPeriod? ResetPeriod { get; set; }
    public string CurrencyCode { get; set; } = "";
    public string SourceText { get; set; } = "";
    public string SourceLocation { get; set; } = "";
    public BenefitRuleNormalizationStatus NormalizationStatus { get; set; } = BenefitRuleNormalizationStatus.NeedsReview;
    public decimal? ConfidenceScore { get; set; }
    public int DisplayOrder { get; set; } = 10;
    public bool IsActive { get; set; } = true;
    public DateOnly? EffectiveDate { get; set; }
    public DateOnly? TerminationDate { get; set; }
    public DateTime CreatedAtUtc { get; set; } = DateTime.UtcNow;
    public DateTime UpdatedAtUtc { get; set; } = DateTime.UtcNow;

    public IReadOnlyList<string> Validate()
    {
        var errors = new List<string>();
        if (RuleType != BenefitRuleType.TieredReimbursement)
        {
            errors.Add("The rule type is not supported.");
            return errors;
        }

        ValidatePercentage(Tier1ReimbursementPercent, "Tier 1 reimbursement percentage", errors);
        ValidatePercentage(Tier2ReimbursementPercent, "Tier 2 reimbursement percentage", errors);

        if (!Tier1ThresholdAmount.HasValue || Tier1ThresholdAmount.Value <= 0m)
        {
            errors.Add("The Tier 1 threshold amount must be greater than zero.");
        }

        if (!ThresholdBasis.HasValue)
        {
            errors.Add("The threshold basis is required.");
        }

        if (!ResetPeriod.HasValue)
        {
            errors.Add("The reset period is required.");
        }

        return errors;
    }

    public bool IsComplete => Validate().Count == 0;

    public bool IsDisplayable =>
        IsActive &&
        NormalizationStatus == BenefitRuleNormalizationStatus.Normalized &&
        IsComplete;

    private static void ValidatePercentage(decimal? value, string label, ICollection<string> errors)
    {
        if (!value.HasValue)
        {
            errors.Add($"{label} is required.");
        }
        else if (value.Value is < 0m or > 100m)
        {
            errors.Add($"{label} must be between 0 and 100.");
        }
    }
}

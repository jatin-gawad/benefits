namespace RenewalReviewBuilder.App.Models;

// Nullable renewal intelligence fields deliberately use these explicit values only when a future
// workflow supplies them. Existing imports leave the corresponding properties null.
public enum NegotiationStatus
{
    Unknown,
    Requested,
    InNegotiation,
    Accepted,
    Rejected,
    Withdrawn
}

public enum NegotiatedSavingsType
{
    Unknown,
    Recurring,
    OneTime,
    Mixed
}

public enum RenewalRecommendationType
{
    Unknown,
    Maintain,
    IncreaseBenefits,
    ReduceBenefits,
    CarrierChange,
    MarketSurvey,
    FundingAdjustment
}

public enum RenewalRiskLevel
{
    Unknown,
    Low,
    Moderate,
    High
}

public enum ScenarioType
{
    Unknown,
    Requested,
    Negotiated,
    Accepted,
    Forecast
}

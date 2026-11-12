namespace RenewalReviewBuilder.App.Services.Intelligence;

public static class GenerateRenewalRecommendation
{
    public static string Run(decimal? renewalIncreasePercent, decimal? lossRatioPercent)
    {
        if (lossRatioPercent is >= 95m)
        {
            return "High claims pressure. Review the carrier increase and prepare claim-driver support.";
        }

        if (renewalIncreasePercent is > 15m && lossRatioPercent is < 75m)
        {
            return "Increase appears high relative to claims. Consider negotiation or alternate options.";
        }

        if (renewalIncreasePercent is <= 5m)
        {
            return "Renewal change is modest. Validate assumptions and consider accepting if service is strong.";
        }

        return "Review claims, trend, pooling, and rate assumptions before final recommendation.";
    }
}

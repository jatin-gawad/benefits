namespace RenewalReviewBuilder.App.Services.Intelligence;

public static class CalculateClaimsUtilization
{
    public static decimal? Run(decimal? claims, decimal? premium)
    {
        return UnderwritingEngine.CalculateLossRatioPercent(claims, premium);
    }
}

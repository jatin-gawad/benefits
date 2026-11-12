namespace RenewalReviewBuilder.App.Services.Intelligence;

public static class CalculateTrend
{
    public static decimal? Run(decimal? claims, decimal? annualTrend)
    {
        return UnderwritingEngine.CalculateTrendedClaims(claims, annualTrend);
    }
}

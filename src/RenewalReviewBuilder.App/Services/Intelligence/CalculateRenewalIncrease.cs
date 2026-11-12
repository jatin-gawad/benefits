namespace RenewalReviewBuilder.App.Services.Intelligence;

public static class CalculateRenewalIncrease
{
    public static decimal? Run(decimal? currentPremium, decimal? renewalPremium)
    {
        return UnderwritingEngine.CalculatePercentChange(currentPremium, renewalPremium);
    }
}

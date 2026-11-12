namespace RenewalReviewBuilder.App.Services.Intelligence;

public static class CalculatePremiumChange
{
    public static decimal? Run(decimal? currentPremium, decimal? renewalPremium)
    {
        return currentPremium.HasValue && renewalPremium.HasValue
            ? decimal.Round(renewalPremium.Value - currentPremium.Value, 2)
            : null;
    }
}

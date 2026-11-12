namespace RenewalReviewBuilder.App.Services.Intelligence;

public static class CalculatePoolingImpact
{
    public static decimal? Run(decimal? claims, decimal? poolingCharge)
    {
        var normalizedCharge = UnderwritingEngine.NormalizeRate(poolingCharge);
        return claims.HasValue && normalizedCharge.HasValue
            ? decimal.Round(claims.Value * normalizedCharge.Value, 2)
            : null;
    }
}

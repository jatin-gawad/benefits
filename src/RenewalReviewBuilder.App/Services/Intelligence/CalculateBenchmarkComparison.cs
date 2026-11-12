namespace RenewalReviewBuilder.App.Services.Intelligence;

public static class CalculateBenchmarkComparison
{
    public static decimal? Run(decimal? clientValue, decimal? benchmarkValue)
    {
        return clientValue.HasValue && benchmarkValue is not null
            ? decimal.Round(clientValue.Value - benchmarkValue.Value, 2)
            : null;
    }
}

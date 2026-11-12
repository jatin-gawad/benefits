using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

/// <summary>
/// Builds an auditable rolling forecast window from saved ASO monthly activity.
/// It never fills gaps or annualizes a partial sequence.
/// </summary>
public sealed class AsoFinancialActivityForecastService
{
    public AsoAnnualForecastWindowResult BuildRollingWindow(
        IEnumerable<AsoFinancialActivityMonthObservation> observations,
        int requestedMonths)
    {
        if (requestedMonths is not (6 or 9 or 12))
            throw new ArgumentOutOfRangeException(nameof(requestedMonths), "Forecast periods must be 6, 9, or 12 months.");

        var latestByMonth = observations
            .GroupBy(observation => observation.ActivityMonth)
            .ToDictionary(
                group => group.Key,
                group => group
                    .OrderByDescending(observation => observation.ImportedAt)
                    .ThenByDescending(observation => observation.SourceFileName, StringComparer.OrdinalIgnoreCase)
                    .First());

        if (latestByMonth.Count == 0)
            return new AsoAnnualForecastWindowResult { RequestedMonths = requestedMonths };

        var endMonth = latestByMonth.Keys.Max();
        var startMonth = endMonth.AddMonths(-(requestedMonths - 1));
        var expectedMonths = Enumerable.Range(0, requestedMonths)
            .Select(index => startMonth.AddMonths(index))
            .ToList();
        var selectedObservations = expectedMonths
            .Where(latestByMonth.ContainsKey)
            .Select(month => latestByMonth[month])
            .ToList();

        return new AsoAnnualForecastWindowResult
        {
            RequestedMonths = requestedMonths,
            StartMonth = startMonth,
            EndMonth = endMonth,
            Observations = selectedObservations,
            MissingMonths = expectedMonths.Where(month => !latestByMonth.ContainsKey(month)).ToList(),
            MissingBenefitPaymentMonths = selectedObservations
                .Where(observation => !observation.TotalBenefitPayment.HasValue)
                .Select(observation => observation.ActivityMonth)
                .ToList()
        };
    }
}

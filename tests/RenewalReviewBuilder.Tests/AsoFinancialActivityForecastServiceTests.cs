using System.Collections.Generic;
using System.Linq;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class AsoFinancialActivityForecastServiceTests
{
    private readonly AsoFinancialActivityForecastService _service = new();

    [Fact]
    public void BuildRollingWindow_Annualizes_ACompleteSixMonthSequence()
    {
        var result = _service.BuildRollingWindow(CreateMonths(new DateOnly(2026, 1, 1), 6, 100m), 6);

        Assert.True(result.IsComplete);
        Assert.Equal(600m, result.ObservedBenefitPayments);
        Assert.Equal(1_200m, result.AnnualizedBenefitPayments);
        Assert.Equal(new DateOnly(2026, 1, 1), result.StartMonth);
        Assert.Equal(new DateOnly(2026, 6, 1), result.EndMonth);
    }

    [Fact]
    public void BuildRollingWindow_DoesNotAnnualize_WhenARequiredMonthIsMissing()
    {
        var observations = CreateMonths(new DateOnly(2025, 10, 1), 9, 100m)
            .Where(observation => observation.ActivityMonth != new DateOnly(2026, 2, 1));

        var result = _service.BuildRollingWindow(observations, 9);

        Assert.False(result.IsComplete);
        Assert.Equal([new DateOnly(2026, 2, 1)], result.MissingMonths);
        Assert.Null(result.AnnualizedBenefitPayments);
    }

    [Fact]
    public void BuildRollingWindow_DoesNotAnnualize_WhenBenefitPaymentsAreMissing()
    {
        var observations = CreateMonths(new DateOnly(2026, 1, 1), 6, 100m).ToList();
        observations[2] = observations[2] with { TotalBenefitPayment = null };

        var result = _service.BuildRollingWindow(observations, 6);

        Assert.False(result.IsComplete);
        Assert.Equal([new DateOnly(2026, 3, 1)], result.MissingBenefitPaymentMonths);
        Assert.Null(result.AnnualizedBenefitPayments);
    }

    [Fact]
    public void BuildRollingWindow_UsesTheLatestImport_WhenAMonthHasMultipleStatements()
    {
        var observations = CreateMonths(new DateOnly(2026, 1, 1), 6, 100m).ToList();
        observations.Add(new AsoFinancialActivityMonthObservation(
            new DateOnly(2026, 6, 1),
            200m,
            "Corrected-ASOActivity.xlsx",
            new DateTime(2026, 7, 2)));

        var result = _service.BuildRollingWindow(observations, 6);

        Assert.True(result.IsComplete);
        Assert.Equal(700m, result.ObservedBenefitPayments);
        Assert.Equal(1_400m, result.AnnualizedBenefitPayments);
        Assert.Contains(result.Observations, observation => observation.SourceFileName == "Corrected-ASOActivity.xlsx");
    }

    [Theory]
    [InlineData(3)]
    [InlineData(10)]
    public void BuildRollingWindow_RejectsUnsupportedWindows(int requestedMonths)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            _service.BuildRollingWindow([], requestedMonths));
    }

    private static IEnumerable<AsoFinancialActivityMonthObservation> CreateMonths(DateOnly start, int count, decimal amount) =>
        Enumerable.Range(0, count).Select(index => new AsoFinancialActivityMonthObservation(
            start.AddMonths(index),
            amount,
            $"ASO-{index + 1}.xlsx",
            new DateTime(2026, 7, 1)));
}

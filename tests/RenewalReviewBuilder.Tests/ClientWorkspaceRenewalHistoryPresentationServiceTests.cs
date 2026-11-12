using RenewalReviewBuilder.App.Services;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class ClientWorkspaceRenewalHistoryPresentationServiceTests
{
    [Fact]
    public void Build_UsesLatestFiveSavedCyclesAndScalesClientPremiums()
    {
        var service = new ClientWorkspaceRenewalHistoryPresentationService();
        var entries = Enumerable.Range(2022, 6)
            .Select(year => Entry(
                year,
                currentPremium: 100_000m + ((year - 2022) * 5_000m),
                renewalPremium: 105_000m + ((year - 2022) * 6_000m),
                ehcLossRatio: 70m + (year - 2022),
                dentalLossRatio: 60m + (year - 2022)))
            .ToList();

        var result = service.Build(entries, "Active");

        Assert.Equal(["2023", "2024", "2025", "2026", "2027"], result.PremiumPoints.Select(point => point.Label));
        Assert.Equal(5, result.LossRatioPoints.Count);
        Assert.All(result.PremiumPoints, point => Assert.True(point.CurrentBarHeight > 0));
        Assert.All(result.PremiumPoints, point => Assert.True(point.RenewalBarHeight > 0));
        Assert.Equal("5 saved renewals: 2023 to 2027", result.PeriodsText);
    }

    [Fact]
    public void Build_UsesActualPriorMetricsAndExistingHealthScoreRules()
    {
        var service = new ClientWorkspaceRenewalHistoryPresentationService();

        var result = service.Build(
        [
            Entry(2026, 120_000m, 127_200m, 86.2m, 28.1m, premiumChange: 6m),
            Entry(2027, 128_291m, 145_657m, 86.5m, 62.7m, premiumChange: 13.5m, highRisk: true)
        ], "Active");

        Assert.Equal(58, result.CurrentHealthScore);
        Assert.Equal("-22 pts vs last renewal", result.HealthScoreComparison.Text);
        Assert.Equal(ClientWorkspaceComparisonTone.Negative, result.HealthScoreComparison.Tone);
        Assert.Equal("vs prior: 86.2%", result.EhcComparison.Text);
        Assert.Equal(ClientWorkspaceComparisonTone.Negative, result.EhcComparison.Tone);
        Assert.Equal("vs prior: 28.1%", result.DentalComparison.Text);
        Assert.Equal(ClientWorkspaceComparisonTone.Negative, result.DentalComparison.Tone);
    }

    [Fact]
    public void Build_DeduplicatesSameCycleUsingNewestLoadedRecord()
    {
        var service = new ClientWorkspaceRenewalHistoryPresentationService();
        var older = Entry(2027, 100_000m, 110_000m, 75m, 65m, dateLoaded: new DateTime(2026, 1, 1));
        var newer = Entry(2027, 120_000m, 135_000m, 82m, 71m, dateLoaded: new DateTime(2026, 2, 1));

        var result = service.Build([older, newer], "Active");

        var point = Assert.Single(result.PremiumPoints);
        Assert.Equal(120_000m, point.CurrentPremium);
        Assert.Equal(135_000m, point.RenewalPremium);
        Assert.Equal(82m, Assert.Single(result.LossRatioPoints).EhcLossRatio);
    }

    [Fact]
    public void Build_ShowsUnavailableStatesWhenNoSavedHistoryExists()
    {
        var result = new ClientWorkspaceRenewalHistoryPresentationService().Build([], "Active");

        Assert.Null(result.Current);
        Assert.Null(result.CurrentHealthScore);
        Assert.Empty(result.PremiumPoints);
        Assert.Empty(result.LossRatioPoints);
        Assert.Equal("Score unavailable", result.HealthScoreComparison.Text);
        Assert.Equal("No saved renewal history", result.PeriodsText);
    }

    private static ClientWorkspaceRenewalHistoryEntry Entry(
        int year,
        decimal currentPremium,
        decimal renewalPremium,
        decimal ehcLossRatio,
        decimal dentalLossRatio,
        decimal premiumChange = 4m,
        bool highRisk = false,
        DateTime? dateLoaded = null)
    {
        return new ClientWorkspaceRenewalHistoryEntry
        {
            RenewalDate = new DateOnly(year, 2, 1),
            RenewalYear = year,
            CurrentPremium = currentPremium,
            RenewalPremium = renewalPremium,
            PremiumChangePercent = premiumChange,
            EhcLossRatio = ehcLossRatio,
            DentalLossRatio = dentalLossRatio,
            IsHighRisk = highRisk,
            DateLoaded = dateLoaded ?? new DateTime(year - 1, 8, 1)
        };
    }
}

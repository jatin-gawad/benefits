using RenewalReviewBuilder.App;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class SunLifeLapPolicyTests
{
    [Theory]
    [InlineData(49, 10_000)]
    [InlineData(50, 15_000)]
    [InlineData(51, 15_000)]
    public void ResolveLapLevel_UsesFiftyAsTheInclusiveHigherBoundary(int lives, decimal expected)
    {
        Assert.Equal(expected, SunLifeLapPolicy.ResolveLapLevel(lives));
    }

    [Fact]
    public void Evaluate_UsesLatestEhcLivesAndFlagsA54To41ThresholdCrossing()
    {
        var assessment = SunLifeLapPolicy.Evaluate(
            "Sun Life",
            [
                new MonthlyClaimsRecord { BenefitLine = "EHC", Month = new DateOnly(2024, 10, 1), Lives = 54 },
                new MonthlyClaimsRecord { BenefitLine = "Extended Health Care", Month = new DateOnly(2025, 10, 1), Lives = 49 },
                new MonthlyClaimsRecord { BenefitLine = "EHC", Month = new DateOnly(2026, 6, 1), Lives = 41 },
                new MonthlyClaimsRecord { BenefitLine = "Dental", Month = new DateOnly(2026, 6, 1), Lives = 75 }
            ]);

        Assert.NotNull(assessment);
        Assert.Equal(41, assessment.LatestEhcLives);
        Assert.Equal(new DateOnly(2026, 6, 1), assessment.LatestMonth);
        Assert.Equal(10_000m, assessment.ProjectedLapLevel);
        Assert.True(assessment.CrossedBelowFifty);
        Assert.Equal(54, assessment.HighestEarlierEhcLives);
        Assert.Contains("crossed below 50 from 54", assessment.ParticipationBasisText);
    }

    [Fact]
    public void Evaluate_DoesNotApplyTheSunLifeRuleToOtherCarriers()
    {
        var assessment = SunLifeLapPolicy.Evaluate(
            "Canada Life",
            [new MonthlyClaimsRecord { BenefitLine = "EHC", Month = new DateOnly(2026, 6, 1), Lives = 41 }]);

        Assert.Null(assessment);
    }

    [Fact]
    public void Workspace_UpdatesRenewalSummaryAndRetainsTheCurrentContractThresholdForAudit()
    {
        var workspace = new ClientWorkspaceViewModel { PoolingThreshold = 15_000m };
        var assessment = new SunLifeLapAssessment(41, new DateOnly(2026, 6, 1), 10_000m, true, 54);

        workspace.ApplySunLifeLapAssessment(assessment);

        Assert.Equal(10_000m, workspace.PoolingThreshold);
        Assert.Equal(15_000m, workspace.CurrentContractPoolingThreshold);
        Assert.Equal(10_000m, workspace.ProjectedPoolingThreshold);
        Assert.Equal(41, workspace.LatestEhcInsuredLives);
        Assert.Contains("Projected Sun Life LAP at renewal", workspace.PoolingThresholdBasis);
        Assert.Contains("current contract $15,000", workspace.PoolingThresholdBasis);
    }

    [Fact]
    public void Workspace_UsesProjectedLapWhenNoVerifiedThresholdExists()
    {
        var workspace = new ClientWorkspaceViewModel();
        var assessment = new SunLifeLapAssessment(41, new DateOnly(2026, 6, 1), 10_000m, true, 54);

        workspace.ApplySunLifeLapAssessment(assessment);

        Assert.Equal(10_000m, workspace.PoolingThreshold);
        Assert.Equal("$10,000", workspace.PoolingThresholdText);
        Assert.StartsWith("Projected Sun Life LAP", workspace.PoolingThresholdBasis);
    }
}

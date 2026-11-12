using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class RenewalEngineTests
{
    [Fact]
    public void BuildRenewalVersions_Preserves_Proposed_Premium_When_Claims_Model_Has_No_Positive_Requirement()
    {
        var analysis = new AnalysisResult();
        analysis.Pdf.RateLines.Add(new RateLine
        {
            Benefit = "Life",
            CurrentPremium = 100m,
            RenewalPremium = 105m,
            CurrentRate = 0.20m,
            RenewalRate = 0.21m,
            Volume = 500000m
        });
        analysis.Pdf.ExperienceLines.Add(new ExperienceLine
        {
            Benefit = "Life",
            Period = "01-Jun-2025 - 31-May-2026",
            BilledPremium = 1200m,
            PaidClaims = 0m,
            PaidLossRatio = 0m
        });

        new RenewalEngine().BuildRenewalVersions(analysis);

        var model = Assert.Single(
            analysis.RenewalVersions,
            version => version.VersionName == RenewalVersionNames.ConsultantModel);
        var life = Assert.Single(model.RateLines);

        Assert.Equal(105m, life.VersionPremium);
        Assert.Equal(0.21m, life.VersionRate);
    }

    [Fact]
    public void BuildRenewalVersions_Uses_Loss_Ratio_For_Rate_Row_Current_Claims()
    {
        var analysis = new AnalysisResult();
        analysis.Pdf.RateLines.Add(new RateLine
        {
            Benefit = "Dental",
            TierOrClass = "Single",
            CurrentPremium = 200m,
            RenewalPremium = 220m,
            Lives = 4
        });
        analysis.Pdf.RateLines.Add(new RateLine
        {
            Benefit = "Dental",
            TierOrClass = "Family",
            CurrentPremium = 300m,
            RenewalPremium = 330m,
            Lives = 3
        });
        analysis.Pdf.ExperienceLines.Add(new ExperienceLine
        {
            Benefit = "Dental",
            Period = "01-Apr-2025 - 31-Mar-2026",
            IncurredClaims = 8307m,
            IncurredLossRatio = 65m
        });

        new RenewalEngine().BuildRenewalVersions(analysis);

        var model = Assert.Single(
            analysis.RenewalVersions,
            version => version.VersionName == RenewalVersionNames.ConsultantModel);
        Assert.Collection(
            model.Calculations.Where(calculation => calculation.BenefitLine == "Dental"),
            single => Assert.Equal(130m, single.CurrentClaims),
            family => Assert.Equal(195m, family.CurrentClaims));
    }
}

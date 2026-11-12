using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class UnderwritingEngineTests
{
    [Fact]
    public void Calculates_Shared_Underwriting_Projection()
    {
        var engine = new UnderwritingEngine();

        var result = engine.Calculate(new UnderwritingCalculationInput
        {
            BenefitLine = "Extended Health Care",
            CurrentClaims = 100000m,
            CurrentPremium = 120000m,
            ProposedPremium = 150000m,
            NegotiatedPremium = 145000m,
            Trend = 0.10m,
            TargetLossRatio = 0.80m,
            PoolingAdjustment = 0.05m,
            IbnrAdjustment = 0.02m,
            RetentionAdminFee = 0.08m,
            ManualAdjustment = 1000m
        });

        Assert.Equal(110000m, result.TrendedClaims);
        Assert.Equal(5500m, result.PoolingAdjustment);
        Assert.Equal(2200m, result.IbnrAdjustment);
        Assert.Equal(118700m, result.ProjectedClaims);
        Assert.Equal(160245m, result.RequiredPremium);
        Assert.Equal(25.0m, result.RenewalChangePercent);
        Assert.Equal(20.8m, result.NegotiatedChangePercent);
        Assert.Equal(5000m, result.Savings);
        Assert.Equal(79.1m, result.ProjectedLossRatio);
    }

    [Fact]
    public void Calculates_Volume_And_Lives_Based_Premiums()
    {
        Assert.Equal(50m, UnderwritingEngine.CalculatePremium("Basic Life", 100000m, null, 0.50m));
        Assert.Equal(420m, UnderwritingEngine.CalculatePremium("Dental", null, 12, 35m));
    }

    [Fact]
    public void Calculates_Allstate_Critical_Illness_Per_Employee_Not_Per_Thousand()
    {
        var premium = UnderwritingEngine.CalculatePremium(
            "Critical Illness",
            25000m,
            10,
            16.96m,
            AllstateCriticalIllnessPdfParser.PerEmployeeRateBasis);

        Assert.Equal(169.60m, premium);
        Assert.Equal(
            16.96m,
            UnderwritingEngine.CalculateRate(
                "Critical Illness",
                25000m,
                10,
                premium,
                AllstateCriticalIllnessPdfParser.PerEmployeeRateBasis));
    }

    [Fact]
    public void Calculates_Critical_Illness_Using_Explicit_Printed_Volume_Basis()
    {
        const decimal volume = 1_620_000m;
        const decimal currentRate = 6.74m;
        const decimal renewalRate = 7.84m;

        Assert.Equal(10_000, UnderwritingEngine.RateVolumeDivisor("Critical Illness", "Per $10,000"));
        Assert.Equal(1_091.88m, UnderwritingEngine.CalculatePremium("Critical Illness", volume, null, currentRate, "Per $10,000"));
        Assert.Equal(1_270.08m, UnderwritingEngine.CalculatePremium("Critical Illness", volume, null, renewalRate, "rate/$10,000"));
        Assert.Equal(currentRate, UnderwritingEngine.CalculateRate("Critical Illness", volume, null, 1_091.88m, "Per $10,000 insured volume"));
    }

    [Fact]
    public void Critical_Illness_Without_Printed_Basis_Defaults_To_Per_Thousand()
    {
        Assert.Equal(1_000, UnderwritingEngine.RateVolumeDivisor("Critical Illness", ""));
        Assert.Equal(10_918.80m, UnderwritingEngine.CalculatePremium("Critical Illness", 1_620_000m, null, 6.74m, ""));
    }
}

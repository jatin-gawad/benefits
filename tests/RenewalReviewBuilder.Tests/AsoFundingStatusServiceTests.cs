using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class AsoFundingStatusServiceTests
{
    [Fact]
    public void Source_surplus_with_zero_float_never_becomes_a_shortfall()
    {
        var result = AsoFundingStatusService.Evaluate(Period(-47828.42m, 0m, "Surplus"));

        Assert.Equal(AsoFundingStatusKind.Surplus, result.Status);
        Assert.Equal(47828.42m, result.AvailableSurplus);
        Assert.Equal(47828.42m, result.FundingMargin);
        Assert.Null(result.FundingRatio);
        Assert.True(result.IsPositive);
        Assert.False(result.NeedsFundingAttention);
        Assert.Equal("Surplus Recorded", result.StatusTitle);
        Assert.Contains("Surplus: $47,828.42", result.Detail);
        Assert.DoesNotContain("Shortfall", result.Detail);
        Assert.Equal("Maintain and monitor funding", result.Recommendation);
    }

    [Fact]
    public void Source_surplus_with_missing_float_preserves_surplus_and_requests_float_confirmation()
    {
        var result = AsoFundingStatusService.Evaluate(Period(-47828.42m, null, "Surplus"));

        Assert.Equal(AsoFundingStatusKind.Surplus, result.Status);
        Assert.Equal(47828.42m, result.AvailableSurplus);
        Assert.Null(result.FundingMargin);
        Assert.Null(result.FundingRatio);
        Assert.True(result.IsPositive);
        Assert.False(result.NeedsFundingAttention);
        Assert.Equal("Confirm required float", result.Recommendation);
    }

    [Fact]
    public void Source_surplus_above_positive_float_is_well_funded()
    {
        var result = AsoFundingStatusService.Evaluate(Period(-47828.42m, 20000m, "Surplus"));

        Assert.Equal(AsoFundingStatusKind.WellFunded, result.Status);
        Assert.Equal(27828.42m, result.FundingMargin);
        Assert.Equal(2.391421m, result.FundingRatio);
        Assert.True(result.IsPositive);
        Assert.False(result.NeedsFundingAttention);
        Assert.Equal("Well Funded", result.StatusTitle);
    }

    [Fact]
    public void Source_surplus_below_positive_float_reports_only_the_true_shortfall()
    {
        var result = AsoFundingStatusService.Evaluate(Period(-47828.42m, 50000m, "Surplus"));

        Assert.Equal(AsoFundingStatusKind.AttentionNeeded, result.Status);
        Assert.Equal(-2171.58m, result.FundingMargin);
        Assert.Equal(0.9565684m, result.FundingRatio);
        Assert.False(result.IsPositive);
        Assert.True(result.NeedsFundingAttention);
        Assert.Contains("Shortfall: $2,171.58", result.Detail);
    }

    [Fact]
    public void Source_deficit_cannot_become_well_funded()
    {
        var result = AsoFundingStatusService.Evaluate(Period(10000m, 20000m, "Deficit"));

        Assert.Equal(AsoFundingStatusKind.AttentionNeeded, result.Status);
        Assert.Equal(-30000m, result.FundingMargin);
        Assert.Equal(0m, result.FundingRatio);
        Assert.False(result.IsPositive);
        Assert.True(result.NeedsFundingAttention);
        Assert.Contains("Deficit: $10,000.00", result.Detail);
        Assert.DoesNotContain("Well Funded", result.StatusTitle);
    }

    [Fact]
    public void Explicit_source_deficit_controls_a_bracketed_negative_balance()
    {
        var result = AsoFundingStatusService.Evaluate(Period(-165.28m, null, "Deficit"));

        Assert.Equal(AsoFundingStatusKind.AttentionNeeded, result.Status);
        Assert.Equal(-165.28m, result.SourceBalance);
        Assert.Equal(0m, result.AvailableSurplus);
        Assert.Equal(-165.28m, result.FundingMargin);
        Assert.True(result.NeedsFundingAttention);
        Assert.Contains("Deficit: $165.28", result.Detail);
    }

    [Fact]
    public void Explicit_funds_available_controls_a_positive_sun_life_balance()
    {
        var result = AsoFundingStatusService.Evaluate(Period(457.86m, null, "Funds Available"));

        Assert.Equal(AsoFundingStatusKind.Surplus, result.Status);
        Assert.Equal(457.86m, result.SourceBalance);
        Assert.Equal(457.86m, result.AvailableSurplus);
        Assert.True(result.IsPositive);
        Assert.False(result.NeedsFundingAttention);
        Assert.Contains("Surplus: $457.86", result.Detail);
    }

    [Fact]
    public void Negative_required_float_requires_source_review()
    {
        var result = AsoFundingStatusService.Evaluate(Period(-47828.42m, -100m, "Surplus"));

        Assert.Equal(AsoFundingStatusKind.Unavailable, result.Status);
        Assert.Equal(47828.42m, result.AvailableSurplus);
        Assert.Null(result.FundingRatio);
        Assert.False(result.IsPositive);
        Assert.False(result.NeedsFundingAttention);
        Assert.Equal("Funding review required", result.StatusTitle);
        Assert.Equal("Confirm required float", result.Recommendation);
    }

    [Fact]
    public void Missing_balance_remains_unavailable()
    {
        var result = AsoFundingStatusService.Evaluate(Period(null, 10000m, ""));

        Assert.Equal(AsoFundingStatusKind.Unavailable, result.Status);
        Assert.Null(result.SourceBalance);
        Assert.False(result.IsPositive);
        Assert.False(result.NeedsFundingAttention);
        Assert.Equal("Funding review required", result.StatusTitle);
    }

    private static AsoFinancialActivityPeriod Period(decimal? balance, decimal? requiredFloat, string status) =>
        new()
        {
            CurrentBalance = balance,
            RequiredFloat = requiredFloat,
            BalanceStatus = status
        };
}

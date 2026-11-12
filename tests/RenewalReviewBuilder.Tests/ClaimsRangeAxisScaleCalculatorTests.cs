using RenewalReviewBuilder.App.Services;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class ClaimsRangeAxisScaleCalculatorTests
{
    private readonly ClaimsRangeAxisScaleCalculator _calculator = new();

    [Theory]
    [InlineData(0, 1, 1)]
    [InlineData(1, 2, 1)]
    [InlineData(4, 5, 1)]
    [InlineData(200, 250, 50)]
    [InlineData(201, 250, 50)]
    [InlineData(1250, 1500, 250)]
    public void Calculate_Uses_Integer_Ticks_And_Has_No_Fixed_Count_Cap(
        int largestCount,
        double expectedMaximum,
        double expectedMajorUnit)
    {
        var scale = _calculator.Calculate([0, largestCount]);

        Assert.Equal(0d, scale.Minimum);
        Assert.Equal(expectedMaximum, scale.Maximum);
        Assert.Equal(expectedMajorUnit, scale.MajorUnit);
        Assert.Equal(0d, scale.MajorUnit % 1d);
    }

    [Fact]
    public void Calculate_Rejects_Negative_Certificate_Counts()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => _calculator.Calculate([2, -1]));
    }
}

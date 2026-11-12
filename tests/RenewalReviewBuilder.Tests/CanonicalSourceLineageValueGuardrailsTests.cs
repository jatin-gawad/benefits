using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class CanonicalSourceLineageValueGuardrailsTests
{
    [Fact]
    public void ResolveSourceFileName_LinkedDocumentWinsAndConflictRequiresReview()
    {
        var result = CanonicalSourceLineageValueGuardrails.ResolveSourceFileName(
            87,
            "Allstate Critical Illness Policy.pdf",
            "Synthetic Client Booklet.pdf");

        Assert.Equal("Synthetic Client Booklet.pdf", result.ResolvedSourceFileName);
        Assert.Equal(CanonicalSourceLineageStatuses.NeedsReview, result.Status);
        Assert.Contains("linked document filename is canonical", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("80%", 80)]
    [InlineData("0.8", 80)]
    [InlineData("100", 100)]
    public void ParseSinglePercentage_NormalizesOnlyOneValidValue(string source, decimal expected)
    {
        Assert.Equal(expected, CanonicalSourceLineageValueGuardrails.ParseSinglePercentage(source));
    }

    [Theory]
    [InlineData("80% / 90%")]
    [InlineData("8090%")]
    [InlineData("8050")]
    public void ParseSinglePercentage_DoesNotInventAnUnsafeValue(string source)
    {
        Assert.Null(CanonicalSourceLineageValueGuardrails.ParseSinglePercentage(source));
    }

    [Fact]
    public void EvaluatePercentage_RejectsOutOfRangeNumericValue()
    {
        var result = CanonicalSourceLineageValueGuardrails.EvaluatePercentage(8050m, "8050%");

        Assert.Null(result.NormalizedValue);
        Assert.Equal(CanonicalValueGuardrailStatuses.Rejected, result.Status);
    }

    [Fact]
    public void EvaluatePercentage_PreservesLegitimateTieredTextForReview()
    {
        var result = CanonicalSourceLineageValueGuardrails.EvaluatePercentage(
            null,
            "80% for the first tier; 100% thereafter");

        Assert.Null(result.NormalizedValue);
        Assert.Equal(CanonicalValueGuardrailStatuses.PreservedForReview, result.Status);
    }
}

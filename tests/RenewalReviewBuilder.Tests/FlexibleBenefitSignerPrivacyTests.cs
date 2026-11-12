using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services.FlexibleBenefits;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class FlexibleBenefitSignerPrivacyTests
{
    [Fact]
    public async Task ExplicitSignerLabels_UseTheSuppliedNamesAndSourceEvidence()
    {
        var result = await Parse("Subscriber signer name: Example Subscriber", "Administrator signer name: Example Administrator");
        var agreement = Assert.Single(result.Agreements);
        Assert.Equal("Example Subscriber", agreement.SubscriberSignerName);
        Assert.Equal("Example Administrator", agreement.AdministratorSignerName);
        Assert.Contains(result.SourceFields, field => field.NormalizedValue == "Example Subscriber");
    }

    [Fact]
    public async Task UnlabelledNames_AreNotAssignedToSignerRoles()
    {
        var result = await Parse("Example Subscriber", "Example Administrator");
        var agreement = Assert.Single(result.Agreements);
        Assert.Equal("", agreement.SubscriberSignerName);
        Assert.Equal("", agreement.AdministratorSignerName);
        Assert.Equal(FlexibleBenefitReviewStatuses.NeedsReview, agreement.ReviewStatus);
    }

    [Fact]
    public async Task AmbiguousSignerLabels_RequireReview()
    {
        var result = await Parse("Subscriber signer name: Example One", "Subscriber signer name: Example Two");
        var agreement = Assert.Single(result.Agreements);
        Assert.Equal("", agreement.SubscriberSignerName);
        Assert.Equal(FlexibleBenefitReviewStatuses.NeedsReview, agreement.ReviewStatus);
    }

    private static Task<FlexibleBenefitExtractionResult> Parse(params string[] signerLines)
    {
        var lines = new[] { "National HealthClaim LifeHealth+", "Administrative Services Indemnity Agreement", "Signed in (city): Example City" }.Concat(signerLines).ToArray();
        var page = new FlexibleBenefitOcrPage { PageNumber = 1, Text = string.Join(Environment.NewLine, lines) };
        foreach (var line in lines) page.Lines.Add(new FlexibleBenefitOcrLine { Text = line });
        var document = new FlexibleBenefitOcrDocument { FilePath = "synthetic-agreement.pdf" };
        document.Pages.Add(page);
        return new NationalHealthClaimFlexibleBenefitParser().ParseAsync(document);
    }
}

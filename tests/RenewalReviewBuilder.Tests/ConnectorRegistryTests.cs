using RenewalReviewBuilder.App.Connectors;
using RenewalReviewBuilder.App.Connectors.Api;
using RenewalReviewBuilder.App.Connectors.Pdf;
using RenewalReviewBuilder.App.Models;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class ConnectorRegistryTests
{
    [Fact]
    public void Resolve_ManulifePdfRenewal_UsesCarrierSpecificConnector()
    {
        var registry = new ConnectorRegistry();

        var connector = registry.Resolve(SourceTypes.Pdf, "Manulife", DocumentTypes.Renewal);

        Assert.IsType<ManulifePdfRenewalConnector>(connector);
        Assert.True(connector.IsEnabled);
        Assert.Equal("Connected", connector.Status);
    }

    [Fact]
    public void Resolve_NationalHealthClaimFlexibleBenefits_UsesCarrierSpecificConnector()
    {
        var registry = new ConnectorRegistry();

        var connector = registry.Resolve(
            SourceTypes.Pdf,
            "National HealthClaim",
            DocumentTypes.FlexibleBenefitsPlan);

        Assert.IsType<NationalHealthClaimFlexibleBenefitsConnector>(connector);
        Assert.True(connector.IsEnabled);
        Assert.Equal("Connected", connector.Status);
    }

    [Fact]
    public async Task CanadaLifeApiConnector_IsDisabledPlaceholder()
    {
        var registry = new ConnectorRegistry();

        var connector = registry.Resolve(SourceTypes.Api, "Canada Life", DocumentTypes.Renewal);

        Assert.IsType<CanadaLifeApiConnector>(connector);
        Assert.False(connector.IsEnabled);
        Assert.Equal("Not Connected", connector.Status);
        var error = await Assert.ThrowsAsync<NotImplementedException>(() => connector.ExtractAsync(new BenefitsConnectorRequest()));
        Assert.Equal("Future API connector placeholder", error.Message);
    }

    [Fact]
    public void PolicySchedule_Resolution_Uses_Source_Carrier_And_DocumentType()
    {
        var registry = new ConnectorRegistry();

        Assert.True(registry.TryResolve(SourceTypes.Pdf, "Manulife", DocumentTypes.GroupPolicySchedule, out var manulife));
        Assert.IsType<ManulifePolicyScheduleConnector>(manulife);
        Assert.False(registry.TryResolve(SourceTypes.Pdf, "Sun Life", DocumentTypes.GroupPolicySchedule, out _));
        Assert.False(registry.TryResolve(SourceTypes.Pdf, "Canada Life", DocumentTypes.GroupPolicySchedule, out _));
    }
}

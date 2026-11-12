using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class NegotiationWorkbenchServiceTests
{
    [Fact]
    public void ApplySavedVersion_ReplacesStaleNegotiatedRows_WithoutChangingCarrierProposal()
    {
        var analysis = new AnalysisResult
        {
            SelectedRenewalVersionName = RenewalVersionNames.CarrierProposed
        };
        var carrier = new RenewalVersion { VersionName = RenewalVersionNames.CarrierProposed };
        carrier.RateLines.Add(new RenewalVersionRateLine
        {
            Benefit = "Extended Health Care",
            TierOrClass = "Single",
            Lives = 14,
            CurrentRate = 206.54m,
            CurrentPremium = 2_891.56m,
            ProposedRate = 190.84m,
            ProposedPremium = 2_671.76m,
            VersionRate = 190.84m,
            VersionPremium = 2_671.76m
        });
        analysis.RenewalVersions.Add(carrier);

        var staleNegotiated = new RenewalVersion { VersionName = RenewalVersionNames.Negotiated };
        staleNegotiated.RateLines.Add(new RenewalVersionRateLine
        {
            Benefit = "Extended Health Care",
            TierOrClass = "Single",
            VersionRate = 190.84m,
            VersionPremium = 2_671.76m
        });
        analysis.RenewalVersions.Add(staleNegotiated);

        var request = new NegotiationWorkbenchSaveRequest
        {
            VersionName = RenewalVersionNames.Negotiated,
            Status = "Negotiated"
        };
        request.Rates.Add(new NegotiationRateSaveRecord
        {
            BenefitLine = "Extended Health Care",
            ClassName = "Single",
            RateBasis = "single",
            Lives = 14,
            CurrentRate = 206.54m,
            CurrentPremium = 2_891.56m,
            ProposedRate = 190.84m,
            ProposedPremium = 2_671.76m,
            ModelledRate = 144.7136m,
            ModelledPremium = 2_026m,
            NegotiatedRate = 164.82m,
            NegotiatedPremium = 2_307m,
            Source = "Western Belting AB revised.pdf"
        });
        request.Rates.Add(new NegotiationRateSaveRecord
        {
            BenefitLine = "Critical Illness",
            ClassName = "A",
            RateBasis = "composite",
            CurrentRate = 6.43m,
            CurrentPremium = 321.50m,
            ProposedRate = 6.25m,
            ProposedPremium = 312.50m,
            NegotiatedPremium = 313m,
            Source = "Western Belting AB revised.pdf"
        });

        var version = new NegotiationWorkbenchService().ApplySavedVersion(
            analysis,
            request,
            RenewalVersionNames.Negotiated);

        Assert.Equal(RenewalVersionNames.Negotiated, analysis.SelectedRenewalVersionName);
        Assert.Same(version, staleNegotiated);
        var negotiatedHealth = Assert.Single(version.RateLines, line => line.Benefit == "Extended Health Care");
        Assert.Equal(164.82m, negotiatedHealth.VersionRate);
        Assert.Equal(2_307m, negotiatedHealth.VersionPremium);
        Assert.Equal(190.84m, negotiatedHealth.ProposedRate);
        Assert.Equal(2_671.76m, negotiatedHealth.ProposedPremium);

        var carrierHealth = Assert.Single(carrier.RateLines);
        Assert.Equal(190.84m, carrierHealth.VersionRate);
        Assert.Equal(2_671.76m, carrierHealth.VersionPremium);

        var negotiatedCriticalIllness = Assert.Single(version.RateLines, line => line.Benefit == "Critical Illness");
        Assert.Null(negotiatedCriticalIllness.VersionRate);
        Assert.Equal(313m, negotiatedCriticalIllness.VersionPremium);

        var powerPointTable = PowerPointRateTableBuilder
            .BuildPlaceholderTables(analysis)["<<RateTable_LineRates>>"];
        var powerPointHealth = Assert.Single(powerPointTable.Rows, row =>
            row.Length > 2 &&
            row[1].StartsWith("Extended Health Care", StringComparison.OrdinalIgnoreCase) &&
            row[2].Equals("Single", StringComparison.OrdinalIgnoreCase));

        Assert.Equal("190.84", powerPointHealth[7]);
        Assert.Equal("$2,671.76", powerPointHealth[8]);
        Assert.Equal("164.82", powerPointHealth[9]);
        Assert.Equal("$2,307.00", powerPointHealth[10]);
    }
}

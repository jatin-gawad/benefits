using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Models.History;
using RenewalReviewBuilder.App.Services.ClientHistory;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class HistoryRelationshipTimelineServiceTests
{
    [Fact]
    public void Qualifies_Only_Strategic_Milestones()
    {
        var service = new HistoryRelationshipTimelineService();
        var milestones = service.BuildMilestones(
        [
            Event("carrier", "Carrier / Vendor", new DateOnly(2024, 1, 1)),
            Event("document", "Document", new DateOnly(2023, 1, 1))
        ]);

        var milestone = Assert.Single(milestones);
        Assert.Equal("carrier", milestone.EventId);
    }

    [Fact]
    public void Orders_Renewals_Handles_YearOnly_Dates_And_Current_State()
    {
        var points = new HistoryRenewalEvolutionService().Build(
        [
            Renewal(24, 2024, true),
            Renewal(22, 2022, false)
        ]);

        Assert.Equal(new[] { 2022, 2024 }, points.Select(point => point.Year!.Value));
        Assert.True(points[1].IsCurrent);
        Assert.Null(points[0].IndustryComparison);
        Assert.Equal(HistoryDatePrecision.YearOnly, points[0].DateReference.Precision);
        Assert.Null(points[0].DateReference.ExactDate);
    }

    [Fact]
    public void Uses_Only_Canonical_Nullable_Renewal_Fields()
    {
        var version = Renewal(44, 2024, true);
        version.RequestedIncreasePercent = 0.08m;
        version.NegotiatedIncreasePercent = 0.06m;
        version.FinalAcceptedIncreasePercent = 0.05m;
        version.IndustryBenchmarkPercent = 0.07m;
        version.NegotiatedSavingsAnnual = 12000m;

        var point = Assert.Single(new HistoryRenewalEvolutionService().Build([version]));

        Assert.Equal(0.08m, point.RequestedIncrease);
        Assert.Equal(0.06m, point.NegotiatedIncrease);
        Assert.Equal(0.05m, point.FinalIncrease);
        Assert.Equal(0.07m, point.IndustryComparison);
        Assert.Equal(12000m, point.NegotiatedSavingsAnnual);
    }

    private static ClientHistoryEventReadModel Event(string id, string category, DateOnly date) => new()
    {
        EventId = id,
        StableSourceKey = id,
        Category = category,
        Title = id,
        EffectiveDate = date,
        EffectiveDateReference = HistoryDateReference.FromExact(date)
    };

    private static RenewalVersionRecord Renewal(long id, int year, bool selected) => new()
    {
        VersionId = id,
        RenewalYear = year,
        IsSelected = selected,
        Status = "Pending"
    };
}
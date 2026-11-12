using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using RenewalReviewBuilder.App.Models.History;
using RenewalReviewBuilder.App.Models.Querying;
using RenewalReviewBuilder.App.Services.ClientHistory;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class ClientHistoryAsOfQueryTests
{
    [Fact]
    public async Task AsOf_Query_Filters_Canonical_Dates_Without_Manufacturing_A_Day_For_YearOnly_Events()
    {
        var builder = new ClientHistoryProjectionBuilder([new StubContributor("Renewal", HistoricalQuerySupport.Supported,
        [Event("year", HistoryDateReference.FromYear(2025)), Event("future", HistoryDateReference.FromExact(new DateOnly(2026, 1, 2)))] )]);

        var result = await builder.BuildAsOfAsync(new AsOfQueryContext { ClientId = 8, AsOfDate = new DateOnly(2025, 12, 31) });

        Assert.Single(result.Events);
        Assert.Equal("year", result.Events[0].EventId);
        Assert.Equal(HistoryDatePrecision.YearOnly, result.Events[0].EffectiveDateReference.Precision);
        Assert.Null(result.Events[0].EffectiveDateReference.ExactDate);
        Assert.Equal(HistoricalQuerySupport.Supported, Assert.Single(result.ContributorSupport).Support);
    }

    [Fact]
    public async Task AsOf_Query_Reports_Limitations_And_Does_Not_Fabricate_Empty_Source_Data()
    {
        var builder = new ClientHistoryProjectionBuilder(
        [
            new StubContributor("PlanDesign", HistoricalQuerySupport.Limited, [], "Load-time dates only."),
            new StubContributor("Financial", HistoricalQuerySupport.Unavailable, [], "No canonical history.")
        ]);

        var result = await builder.BuildAsOfAsync(new AsOfQueryContext { ClientId = 8, AsOfDate = new DateOnly(2025, 12, 31) });

        Assert.Empty(result.Events);
        Assert.Collection(result.ContributorSupport,
            plan => { Assert.Equal(HistoricalQuerySupport.Limited, plan.Support); Assert.Equal("Load-time dates only.", plan.Detail); },
            financial => Assert.Equal(HistoricalQuerySupport.Unavailable, financial.Support));
    }

    [Fact]
    public async Task Legacy_Build_Remains_Unfiltered_While_AsOf_Options_Control_Manual_And_Unvalidated_Events()
    {
        var manual = Event("manual", HistoryDateReference.FromExact(new DateOnly(2027, 1, 1)), "Manual", HistoryValidationStatus.NeedsReview);
        var builder = new ClientHistoryProjectionBuilder([new StubContributor("Manual", HistoricalQuerySupport.Supported, [manual])]);

        Assert.Single(await builder.BuildAsync(8));
        Assert.Empty((await builder.BuildAsOfAsync(new AsOfQueryContext { ClientId = 8, AsOfDate = new DateOnly(2028, 1, 1) })).Events);
        Assert.Single((await builder.BuildAsOfAsync(new AsOfQueryContext { ClientId = 8, AsOfDate = new DateOnly(2028, 1, 1), IncludeManualEvents = true, IncludeUnvalidatedEvents = true })).Events);
    }

    private static ClientHistoryEventReadModel Event(string id, HistoryDateReference date, string module = "Renewal", HistoryValidationStatus status = HistoryValidationStatus.Valid) => new()
    {
        EventId = id, StableSourceKey = id, ClientId = 8, SourceModule = module, SourceEntityType = module, SourceEntityId = id,
        EffectiveDate = date.ExactDate, EffectiveDateReference = date, RecordedDate = new DateTime(2025, 1, 1),
        Category = "Renewals", EventType = "Renewal", Title = id, ValidationStatus = status
    };

    private sealed class StubContributor(string sourceModule, HistoricalQuerySupport support, IReadOnlyList<ClientHistoryEventReadModel> events, string detail = "") : IClientHistoryContributor
    {
        public string SourceModule => sourceModule;
        public Task<IReadOnlyList<ClientHistoryEventReadModel>> ContributeAsync(long clientId, CancellationToken cancellationToken = default) => Task.FromResult(events);
        public Task<ClientHistoryContribution> ContributeAsync(AsOfQueryContext context) => Task.FromResult(new ClientHistoryContribution
        {
            Events = events,
            HistoricalSupport = new HistoricalQuerySupportStatus { SourceModule = SourceModule, Support = support, Detail = detail }
        });
    }
}
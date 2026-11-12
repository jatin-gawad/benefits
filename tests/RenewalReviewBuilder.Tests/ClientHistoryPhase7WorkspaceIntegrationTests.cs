using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Models.History;
using RenewalReviewBuilder.App.Services.ClientHistory;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class ClientHistoryPhase7WorkspaceIntegrationTests
{
    [Fact]
    public async Task Load_Populates_Milestones_And_RenewalPoints_From_Canonical_Sources()
    {
        var state = new HistorySelectionState();
        using var controller = CreateController(state, Events(), Versions());

        await controller.LoadAsync(17);

        Assert.Contains(controller.RelationshipMilestones, milestone => milestone.EventId == "carrier-change");
        var point = Assert.Single(controller.RenewalPoints);
        Assert.Equal(0.08m, point.RequestedIncrease);
        Assert.Equal(0.06m, point.NegotiatedIncrease);
        Assert.Equal(0.05m, point.FinalIncrease);
        Assert.Equal(0.07m, point.IndustryComparison);
        Assert.Equal(12000m, point.NegotiatedSavingsAnnual);
        Assert.False(controller.IsRelationshipTimelineLoading);
        Assert.False(controller.IsRenewalEvolutionLoading);
        Assert.Equal("", controller.RelationshipTimelineError);
        Assert.Equal("", controller.RenewalEvolutionError);
    }

    [Fact]
    public async Task Empty_And_Partial_Canonical_History_Use_States_Without_Fabricated_Values()
    {
        using var empty = CreateController(new HistorySelectionState(), [], []);
        await empty.LoadAsync(17);

        Assert.Empty(empty.RelationshipMilestones);
        Assert.Empty(empty.RenewalPoints);
        Assert.Contains("No canonical history", empty.RelationshipTimelineEmptyState);
        Assert.Contains("No canonical renewal versions", empty.RenewalEvolutionEmptyState);

        var partialVersion = new RenewalVersionRecord
        {
            VersionId = 5,
            RenewalId = 50,
            RenewalYear = 2025,
            Status = "Pending"
        };
        using var partial = CreateController(
            new HistorySelectionState(),
            [RenewalEvent(partialVersion)],
            [partialVersion]);

        await partial.LoadAsync(17);

        Assert.Empty(partial.RelationshipMilestones);

        var point = Assert.Single(partial.RenewalPoints);
        Assert.Null(point.RequestedIncrease);
        Assert.Null(point.NegotiatedIncrease);
        Assert.Null(point.FinalIncrease);
        Assert.Null(point.IndustryComparison);
        Assert.Null(point.NegotiatedSavingsAnnual);
        Assert.Equal(HistoryDatePrecision.YearOnly, point.DateReference.Precision);
        Assert.Null(point.DateReference.ExactDate);
        Assert.Equal("2025", point.EffectiveDateDisplay);
    }

    [Fact]
    public async Task Milestone_And_Renewal_Selection_Update_The_One_Shared_SelectionState()
    {
        var state = new HistorySelectionState();
        using var controller = CreateController(state, Events(), Versions());
        await controller.LoadAsync(17);

        var carrierMilestone = controller.RelationshipMilestones.Single(milestone => milestone.EventId == "carrier-change");
        Assert.True(controller.SelectRelationshipMilestone(carrierMilestone));
        Assert.Equal("carrier-change", state.SelectedHistoryEventId);
        Assert.Equal(42L, state.SelectedRenewalId);
        Assert.Equal(9L, state.SelectedCarrierId);
        Assert.Equal(2026, state.SelectedYear);
        Assert.True(state.IsDetailsDrawerOpen);

        state.IsDetailsDrawerOpen = false;
        var point = Assert.Single(controller.RenewalPoints);
        Assert.True(controller.SelectRenewalPoint(point));
        Assert.Equal(42L, state.SelectedRenewalId);
        Assert.Equal(RenewalEventId(point), state.SelectedHistoryEventId);
        Assert.Equal(2026, state.SelectedYear);
        Assert.True(state.IsDetailsDrawerOpen);
        Assert.Same(state, controller.Selection);
        Assert.Same(state, controller.Presentation.Selection);
    }

    [Fact]
    public async Task Filters_Refresh_Compact_Components_And_Clear_Restores_Them()
    {
        using var controller = CreateController(new HistorySelectionState(), Events(), Versions());
        await controller.LoadAsync(17);

        controller.SetFilters("carrier change", "All benefit areas", null, null);
        Assert.Single(controller.RelationshipMilestones);
        Assert.Empty(controller.RenewalPoints);

        controller.ClearFilters();
        Assert.NotEmpty(controller.RelationshipMilestones);
        Assert.Single(controller.RenewalPoints);

        // Replay selection no longer filters the presentation; as-of projection queries own that responsibility.
        controller.Selection.ReplayAsOfDate = new DateOnly(2025, 12, 31);
        Assert.NotEmpty(controller.RelationshipMilestones);
        Assert.Single(controller.RenewalPoints);
        controller.Selection.ReplayAsOfDate = null;
        Assert.NotEmpty(controller.RelationshipMilestones);
        Assert.Single(controller.RenewalPoints);
    }

    [Fact]
    public async Task Refresh_And_Client_Change_Reload_Both_Canonical_Collections()
    {
        var clientsLoaded = new List<long>();
        var projection = new Projection(clientId =>
        {
            clientsLoaded.Add(clientId);
            return Task.FromResult<IReadOnlyList<ClientHistoryEventReadModel>>(Events());
        });
        using var controller = new ClientHistoryWorkspaceController(
            new HistorySelectionState(),
            projection,
            (clientId, _) => Task.FromResult<IReadOnlyList<RenewalVersionRecord>>(Versions()));

        await controller.LoadAsync(17);
        await controller.RefreshAsync();
        await controller.LoadAsync(18);

        Assert.Equal(new[] { 17L, 17L, 18L }, clientsLoaded);
        Assert.Equal(18L, controller.DatabaseClientId);
        Assert.NotEmpty(controller.RelationshipMilestones);
        Assert.NotEmpty(controller.RenewalPoints);
    }

    [Fact]
    public async Task Errors_Are_Exposed_Without_Leaving_Stale_Compact_Collections()
    {
        using var controller = new ClientHistoryWorkspaceController(
            new HistorySelectionState(),
            new Projection(_ => throw new InvalidOperationException("projection failure")),
            (_, _) => throw new InvalidOperationException("renewal failure"));

        await controller.LoadAsync(17);

        Assert.Empty(controller.RelationshipMilestones);
        Assert.Empty(controller.RenewalPoints);
        Assert.Contains("projection failure", controller.RelationshipTimelineError);
        Assert.Contains("renewal failure", controller.RenewalEvolutionError);
        Assert.False(controller.IsRelationshipTimelineLoading);
        Assert.False(controller.IsRenewalEvolutionLoading);
    }

    [Fact]
    public void Resolver_Uses_Related_Canonical_References_Only()
    {
        Assert.Equal(42L, HistorySelectionReferenceResolver.ResolveRenewalId(["Document:3", "Renewal:42"]));
        Assert.Equal(9L, HistorySelectionReferenceResolver.ResolveCarrierId(["Carrier:9"]));
        Assert.Null(HistorySelectionReferenceResolver.ResolveRenewalId(["Renewal:not-a-number"]));
        Assert.Null(HistorySelectionReferenceResolver.ResolveCarrierId([]));
    }

    private static ClientHistoryWorkspaceController CreateController(
        HistorySelectionState state,
        IReadOnlyList<ClientHistoryEventReadModel> events,
        IReadOnlyList<RenewalVersionRecord> versions) =>
        new(
            state,
            new Projection(_ => Task.FromResult(events)),
            (_, _) => Task.FromResult(versions));

    private static IReadOnlyList<ClientHistoryEventReadModel> Events() =>
    [
        new ClientHistoryEventReadModel
        {
            EventId = "carrier-change",
            StableSourceKey = "CarrierVendor:Carrier:9:change",
            SourceModule = "CarrierVendor",
            SourceEntityType = "Carrier",
            SourceEntityId = "9",
            Category = "Carrier / Vendor",
            EventType = "Carrier change",
            Title = "Carrier change",
            EffectiveDate = new DateOnly(2026, 8, 1),
            EffectiveDateReference = HistoryDateReference.FromExact(new DateOnly(2026, 8, 1)),
            RelatedRecords = ["Renewal:42", "Carrier:9"]
        },
        RenewalEvent(Versions().Single())
    ];

    private static IReadOnlyList<RenewalVersionRecord> Versions() =>
    [
        new RenewalVersionRecord
        {
            VersionId = 5,
            RenewalId = 42,
            EffectiveDate = new DateOnly(2026, 8, 1),
            Status = "Accepted",
            IsSelected = true,
            RequestedIncreasePercent = 0.08m,
            NegotiatedIncreasePercent = 0.06m,
            FinalAcceptedIncreasePercent = 0.05m,
            IndustryBenchmarkPercent = 0.07m,
            NegotiatedSavingsAnnual = 12000m
        }
    ];

    private static ClientHistoryEventReadModel RenewalEvent(RenewalVersionRecord version) => new()
    {
        EventId = RenewalEventId(version),
        StableSourceKey = RenewalEventId(version),
        SourceModule = "Renewal",
        SourceEntityType = "RenewalVersion",
        SourceEntityId = version.VersionId.ToString(),
        Category = "Renewals",
        EventType = version.IsSelected ? "Current renewal" : "Renewal",
        Title = "Renewal",
        EffectiveDate = version.EffectiveDate,
        EffectiveDateReference = version.EffectiveDate.HasValue
            ? HistoryDateReference.FromExact(version.EffectiveDate.Value)
            : HistoryDateReference.FromYear(version.RenewalYear!.Value),
        RelatedRecords = version.RenewalId.HasValue ? [$"Renewal:{version.RenewalId.Value}"] : []
    };

    private static string RenewalEventId(HistoryRenewalEvolutionPoint point) => point.EventId;
    private static string RenewalEventId(RenewalVersionRecord version) =>
        ClientHistoryProjectionBuilder.StableKey("Renewal", "RenewalVersion", version.VersionId, "version");

    private sealed class Projection(Func<long, Task<IReadOnlyList<ClientHistoryEventReadModel>>> load) : IClientHistoryProjectionBuilder
    {
        public Task<IReadOnlyList<ClientHistoryEventReadModel>> BuildAsync(long clientId, CancellationToken cancellationToken = default) => load(clientId);
    }
}

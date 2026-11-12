using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Models.History;
using RenewalReviewBuilder.App.Models.Querying;
using RenewalReviewBuilder.App.Services.ClientHistory;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class ClientHistoryReplaySynchronizationTests
{
    [Fact]
    public void Replay_Is_Inactive_By_Default_And_Uses_One_Selection_Store()
    {
        var state = new HistorySelectionState();
        using var controller = Controller(state, new RecordingProjection(CurrentEvents()), Versions());

        Assert.False(controller.IsReplayActive);
        Assert.Null(controller.ReplayAsOfDate);
        Assert.True(controller.IsCurrentProjection);
        Assert.Same(state, controller.Selection);
        Assert.Same(state, controller.Presentation.Selection);
    }

    [Fact]
    public async Task Activate_Replay_Uses_AsOf_Context_And_GoToCurrent_Restores_Current_Kpis()
    {
        var projection = new RecordingProjection(CurrentEvents());
        using var controller = Controller(new HistorySelectionState(), projection, Versions());
        await controller.LoadAsync(73);

        await controller.ActivateReplayAsync(new DateOnly(2024, 12, 31));

        Assert.True(controller.IsReplayActive);
        Assert.Equal(new DateOnly(2024, 12, 31), Assert.Single(projection.AsOfContexts).AsOfDate);
        Assert.DoesNotContain(controller.Events, item => item.EventId == "future");
        Assert.Equal(8, controller.Kpis.YearsAsClient);
        Assert.Equal(100m, controller.Kpis.TotalValueCreated);
        Assert.Equal(0.04m, controller.Kpis.AverageRenewalIncrease);
        Assert.Equal(1, controller.Kpis.CompletedRenewalCount);
        Assert.Equal("Unavailable for Replay", controller.CurrentCarrierDisplay);
        Assert.Contains("Carrier history is unavailable", controller.ReplaySupportSummary);

        await controller.GoToCurrentAsync();

        Assert.False(controller.IsReplayActive);
        Assert.True(controller.IsCurrentProjection);
        Assert.Contains(controller.Events, item => item.EventId == "future");
        Assert.Equal(0.06m, controller.Kpis.AverageRenewalIncrease);
        Assert.Equal(2, controller.Kpis.CompletedRenewalCount);
        Assert.Equal("Across 2 completed renewals", controller.AverageRenewalIncreaseSubtitle);
        Assert.Equal("Manulife", controller.CurrentCarrierDisplay);
    }

    [Fact]
    public async Task Replay_Limited_Carrier_Support_Does_Not_Expose_Legacy_Current_State_Carrier()
    {
        using var controller = Controller(
            new HistorySelectionState(),
            new LimitedCarrierProjection(CurrentEvents()),
            Versions());
        await controller.LoadAsync(73);
        Assert.Equal("Manulife", controller.CurrentCarrierDisplay);

        await controller.ActivateReplayAsync(new DateOnly(2024, 12, 31));

        Assert.Contains(controller.Events, item => item.EventId == "carrier-current");
        Assert.Equal("Manulife", controller.Kpis.CurrentCarrier);
        Assert.Equal("Limited for Replay", controller.CurrentCarrierDisplay);
        Assert.Contains("limited", controller.ReplaySupportSummary, StringComparison.OrdinalIgnoreCase);

        await controller.GoToCurrentAsync();
        Assert.Equal("Manulife", controller.CurrentCarrierDisplay);
    }

    [Fact]
    public async Task Replay_Preserves_Mode_Filters_And_Visible_Selection()
    {
        var state = new HistorySelectionState();
        var projection = new RecordingProjection(CurrentEvents());
        using var controller = Controller(state, projection, Versions());
        await controller.LoadAsync(73);
        controller.Presentation.SelectMode(HistoryMode.Evolution);
        controller.SetFilters("value", "All benefit areas", null, null);
        Assert.True(controller.SelectHistoryEvent("value"));

        await controller.ActivateReplayAsync(new DateOnly(2024, 12, 31));

        Assert.Equal(HistoryMode.Evolution, state.ActiveHistoryMode);
        Assert.Equal("value", controller.Presentation.SearchText);
        Assert.Equal("value", state.SelectedHistoryEventId);
        controller.Presentation.CloseDrawer();
        Assert.Equal("value", state.SelectedHistoryEventId);
        Assert.False(state.IsDetailsDrawerOpen);
    }

    [Fact]
    public async Task Replay_Clears_Selection_When_Selected_Event_Is_No_Longer_Canonical()
    {
        var state = new HistorySelectionState();
        using var controller = Controller(state, new RecordingProjection(CurrentEvents()), Versions());
        await controller.LoadAsync(73);
        Assert.True(controller.SelectHistoryEvent("future"));

        await controller.ActivateReplayAsync(new DateOnly(2024, 12, 31));

        Assert.Equal("", state.SelectedHistoryEventId);
        Assert.False(state.IsDetailsDrawerOpen);
    }

    [Fact]
    public async Task Timeline_Relationship_Renewal_And_Evolution_Selections_Synchronize_Shared_Identifiers()
    {
        var state = new HistorySelectionState();
        using var controller = Controller(state, new RecordingProjection(CurrentEvents()), Versions());
        await controller.LoadAsync(73);

        Assert.True(controller.SelectHistoryEvent("relationship"));
        Assert.Equal(42L, state.SelectedRenewalId);
        Assert.Equal(9L, state.SelectedCarrierId);
        Assert.Equal(2024, state.SelectedYear);
        Assert.Equal(new long[] { 700 }, state.SelectedDocumentIds);

        var milestone = controller.RelationshipMilestones.Single(item => item.EventId == "relationship");
        Assert.True(controller.SelectRelationshipMilestone(milestone));
        Assert.True(milestone.EventId == state.SelectedHistoryEventId);

        var point = controller.RenewalPoints.Single(item => item.RenewalId == 42);
        Assert.True(controller.SelectRenewalPoint(point));
        Assert.Equal(42L, state.SelectedRenewalId);
        Assert.Equal(point.EventId, state.SelectedHistoryEventId);

        controller.Presentation.SelectEvolutionStep("value");
        Assert.Equal("value", state.SelectedEvolutionStepId);
        Assert.Equal("value", state.SelectedHistoryEventId);
        Assert.True(state.IsDetailsDrawerOpen);
    }

    [Fact]
    public async Task Renewal_Point_With_No_Canonical_Renewal_Id_Is_Not_Selected_By_Default()
    {
        var versions = new[]
        {
            new RenewalVersionRecord { VersionId = 99, RenewalYear = 2026, Status = "Pending" }
        };
        using var controller = Controller(new HistorySelectionState(), new RecordingProjection([]), versions);

        await controller.LoadAsync(73);

        var point = Assert.Single(controller.RenewalPoints);
        Assert.Null(point.RenewalId);
        Assert.False(point.IsSelected);
        Assert.Equal("", controller.Selection.SelectedHistoryEventId);
    }

    [Fact]
    public async Task Newer_Replay_Supersedes_Stale_Result_Even_When_Older_Query_Ignores_Cancellation()
    {
        var projection = new ControlledProjection(CurrentEvents());
        using var controller = Controller(new HistorySelectionState(), projection, Versions());
        await controller.LoadAsync(73);

        var older = controller.ActivateReplayAsync(new DateOnly(2023, 12, 31));
        await projection.OlderStarted.Task;
        var newer = controller.ActivateReplayAsync(new DateOnly(2024, 12, 31));
        await projection.NewerStarted.Task;
        projection.Newer.SetResult(Result(new DateOnly(2024, 12, 31), [Event("newer", new DateOnly(2024, 1, 1))]));
        await newer;
        projection.Older.SetResult(Result(new DateOnly(2023, 12, 31), [Event("stale", new DateOnly(2023, 1, 1))]));
        await older;

        Assert.Equal("newer", Assert.Single(controller.Events).EventId);
        Assert.Equal(new DateOnly(2024, 12, 31), controller.ReplayAsOfDate);
    }

    [Fact]
    public async Task YearOnly_Event_Remains_YearOnly_During_Replay()
    {
        var yearOnly = new ClientHistoryEventReadModel
        {
            EventId = "year-only", StableSourceKey = "year-only", ClientId = 73, SourceModule = "Renewal",
            SourceEntityType = "Renewal", SourceEntityId = "year-only", EffectiveDateReference = HistoryDateReference.FromYear(2024),
            EventType = "Renewal", Category = "Renewals", Title = "Year-only renewal", ValidationStatus = HistoryValidationStatus.Valid
        };
        var builder = new ClientHistoryProjectionBuilder([new StaticContributor([yearOnly])]);

        var result = await builder.BuildAsOfAsync(new AsOfQueryContext { ClientId = 73, AsOfDate = new DateOnly(2024, 6, 30), IncludeManualEvents = true, IncludeUnvalidatedEvents = true });

        var item = Assert.Single(result.Events);
        Assert.Equal(HistoryDatePrecision.YearOnly, item.EffectiveDateReference.Precision);
        Assert.Null(item.EffectiveDateReference.ExactDate);
    }

    private static ClientHistoryWorkspaceController Controller(HistorySelectionState state, IClientHistoryProjectionBuilder projection, IReadOnlyList<RenewalVersionRecord> versions) =>
        new(state, projection, (_, _) => Task.FromResult(versions));

    private static IReadOnlyList<ClientHistoryEventReadModel> CurrentEvents() =>
    [
        Event("start", new DateOnly(2016, 3, 1)),
        ValueEvent(),
        RelationshipEvent(),
        CarrierCurrentEvent(),
        RenewalEvent(),
        Event("future", new DateOnly(2026, 10, 1))
    ];

    private static IReadOnlyList<RenewalVersionRecord> Versions() =>
    [
        new RenewalVersionRecord { VersionId = 4, RenewalId = 41, RenewalYear = 2022, Status = "Completed", FinalAcceptedIncreasePercent = 0.04m },
        new RenewalVersionRecord { VersionId = 44, RenewalId = 41, RenewalYear = 2022, Status = "Final Approved", FinalAcceptedIncreasePercent = 0.05m },
        new RenewalVersionRecord { VersionId = 5, RenewalId = 42, RenewalYear = 2025, Status = "Completed", FinalAcceptedIncreasePercent = 0.08m },
        new RenewalVersionRecord { VersionId = 6, RenewalId = 43, RenewalYear = 2025, Status = "Pending", FinalAcceptedIncreasePercent = 0.20m }
    ];

    private static ClientHistoryEventReadModel Event(string id, DateOnly? date) => new()
    {
        EventId = id,
        StableSourceKey = id,
        ClientId = 73,
        SourceModule = id == "carrier-current" ? "CarrierVendor" : "Manual",
        SourceEntityType = "History",
        SourceEntityId = id,
        EffectiveDate = date,
        EffectiveDateReference = date.HasValue ? HistoryDateReference.FromExact(date.Value) : new HistoryDateReference(),
        EventType = "History",
        Category = "Other",
        Title = id,
        Confidence = 1m,
        ValidationStatus = HistoryValidationStatus.Valid
    };

    private static ClientHistoryEventReadModel ValueEvent() => new()
    {
        EventId = "value", StableSourceKey = "value", ClientId = 73, SourceModule = "Manual", SourceEntityType = "History", SourceEntityId = "value",
        EffectiveDate = new DateOnly(2024, 5, 31), EffectiveDateReference = HistoryDateReference.FromExact(new DateOnly(2024, 5, 31)),
        EventType = "Benefit Plan Amendments", Category = "Plan Design", Title = "Value", FinancialImpact = "$100", Confidence = 1m,
        ValidationStatus = HistoryValidationStatus.Valid,
        ChangeSet = new HistoryChangeSet { Changes = [new HistoryChangeItem { DisplayLabel = "Vision", PreviousValue = "$200", NewValue = "$350" }] }
    };

    private static ClientHistoryEventReadModel RelationshipEvent() => new()
    {
        EventId = "relationship", StableSourceKey = "relationship", ClientId = 73, SourceModule = "Manual", SourceEntityType = "History", SourceEntityId = "relationship",
        EffectiveDate = new DateOnly(2024, 10, 1), EffectiveDateReference = HistoryDateReference.FromExact(new DateOnly(2024, 10, 1)),
        EventType = "Carrier change", Category = "Carrier / Vendor", Title = "Moved to Manulife", RelatedRecords = ["Renewal:42", "Carrier:9"],
        Evidence = [new HistoryEvidenceReference { DocumentId = 700 }], Confidence = 1m, ValidationStatus = HistoryValidationStatus.Valid
    };

    private static ClientHistoryEventReadModel RenewalEvent() => new()
    {
        EventId = ClientHistoryProjectionBuilder.StableKey("Renewal", "RenewalVersion", 5, "version"),
        StableSourceKey = ClientHistoryProjectionBuilder.StableKey("Renewal", "RenewalVersion", 5, "version"),
        ClientId = 73, SourceModule = "Renewal", SourceEntityType = "RenewalVersion", SourceEntityId = "5",
        EffectiveDateReference = HistoryDateReference.FromYear(2025), EventType = "Renewal", Category = "Renewals", Title = "2025 renewal",
        RelatedRecords = ["Renewal:42"], Confidence = 1m, ValidationStatus = HistoryValidationStatus.Valid
    };
    private static ClientHistoryEventReadModel CarrierCurrentEvent() => new()
    {
        EventId = "carrier-current", StableSourceKey = "carrier-current", ClientId = 73, SourceModule = "CarrierVendor", SourceEntityType = "Client", SourceEntityId = "73",
        EventType = "Current carrier", Category = "Carrier / Vendor", Title = "Manulife", Confidence = 1m, ValidationStatus = HistoryValidationStatus.Valid
    };
    private static ClientHistoryAsOfQueryResult Result(
        DateOnly date,
        IReadOnlyList<ClientHistoryEventReadModel> events,
        HistoricalQuerySupport carrierSupport = HistoricalQuerySupport.Unavailable) => new()
    {
        Context = new AsOfQueryContext { ClientId = 73, AsOfDate = date },
        Events = events,
        ContributorSupport =
        [
            new HistoricalQuerySupportStatus
            {
                SourceModule = "CarrierVendor",
                Support = carrierSupport,
                Detail = carrierSupport == HistoricalQuerySupport.Limited
                    ? "Carrier history has limited replay support."
                    : "Carrier history is unavailable for this replay date."
            }
        ]
    };

    private sealed class RecordingProjection(IReadOnlyList<ClientHistoryEventReadModel> current) : IClientHistoryProjectionBuilder
    {
        public List<AsOfQueryContext> AsOfContexts { get; } = [];
        public Task<IReadOnlyList<ClientHistoryEventReadModel>> BuildAsync(long clientId, CancellationToken cancellationToken = default) => Task.FromResult(current);
        public Task<ClientHistoryAsOfQueryResult> BuildAsOfAsync(AsOfQueryContext context)
        {
            AsOfContexts.Add(context);
            var visible = current.Where(item => item.SourceModule != "CarrierVendor" && (!item.EffectiveDate.HasValue || item.EffectiveDate.Value <= context.AsOfDate)).ToList();
            return Task.FromResult(Result(context.AsOfDate!.Value, visible));
        }
    }

    private sealed class ControlledProjection(IReadOnlyList<ClientHistoryEventReadModel> current) : IClientHistoryProjectionBuilder
    {
        public TaskCompletionSource OlderStarted { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public TaskCompletionSource NewerStarted { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public TaskCompletionSource<ClientHistoryAsOfQueryResult> Older { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public TaskCompletionSource<ClientHistoryAsOfQueryResult> Newer { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public Task<IReadOnlyList<ClientHistoryEventReadModel>> BuildAsync(long clientId, CancellationToken cancellationToken = default) => Task.FromResult(current);
        public Task<ClientHistoryAsOfQueryResult> BuildAsOfAsync(AsOfQueryContext context)
        {
            if (context.AsOfDate!.Value.Year == 2023) { OlderStarted.SetResult(); return Older.Task; }
            NewerStarted.SetResult(); return Newer.Task;
        }
    }

    private sealed class LimitedCarrierProjection(IReadOnlyList<ClientHistoryEventReadModel> current) : IClientHistoryProjectionBuilder
    {
        public Task<IReadOnlyList<ClientHistoryEventReadModel>> BuildAsync(
            long clientId,
            CancellationToken cancellationToken = default) => Task.FromResult(current);

        public Task<ClientHistoryAsOfQueryResult> BuildAsOfAsync(AsOfQueryContext context)
        {
            var visible = current
                .Where(item => !item.EffectiveDate.HasValue || item.EffectiveDate.Value <= context.AsOfDate)
                .ToList();
            return Task.FromResult(Result(
                context.AsOfDate!.Value,
                visible,
                HistoricalQuerySupport.Limited));
        }
    }

    private sealed class StaticContributor(IReadOnlyList<ClientHistoryEventReadModel> events) : IClientHistoryContributor
    {
        public string SourceModule => "Renewal";
        public Task<IReadOnlyList<ClientHistoryEventReadModel>> ContributeAsync(long clientId, CancellationToken cancellationToken = default) => Task.FromResult(events);
    }
}

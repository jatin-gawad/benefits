using System.Collections.Generic;
using RenewalReviewBuilder.App.Models.History;
using RenewalReviewBuilder.App.Services.ClientHistory;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class ClientHistoryPhase5Tests
{
    private readonly ClientHistoryTimelineService _timeline = new();

    [Fact]
    public void Selection_State_Opens_Closes_And_Keeps_Selected_Event()
    {
        var state = new HistorySelectionState();
        state.SelectedHistoryEventId = "Renewal:Renewal:12:record";
        state.IsDetailsDrawerOpen = true;
        Assert.True(state.IsDetailsDrawerOpen);
        Assert.Equal("Renewal:Renewal:12:record", state.SelectedHistoryEventId);
        state.IsDetailsDrawerOpen = false;
        Assert.False(state.IsDetailsDrawerOpen);
        Assert.Equal("Renewal:Renewal:12:record", state.SelectedHistoryEventId);
    }

    [Fact]
    public void Filtered_Timeline_Excludes_Excluded_Events_And_Navigates_Within_Filtered_Order()
    {
        var rows = _timeline.BuildRows(Events(), new ClientHistoryFilter { Category = "Plan Design" });
        Assert.Equal(2, rows.Count);
        Assert.Equal("plan-2", rows[0].EventId);
        Assert.Equal("plan-1", rows[1].EventId);
    }

    [Fact]
    public void Search_And_Filter_State_Is_Not_Mutated_By_Selection_State()
    {
        var filter = new ClientHistoryFilter { Search = "dental", Category = "Plan Design", From = new DateOnly(2023, 1, 1), To = new DateOnly(2025, 12, 31), OldestFirst = true };
        var state = new HistorySelectionState { SelectedYear = 2024, SelectedHistoryEventId = "plan-1", IsDetailsDrawerOpen = true };
        _ = _timeline.BuildRows(Events(), filter);
        state.IsDetailsDrawerOpen = false;
        Assert.Equal("dental", filter.Search); Assert.Equal("Plan Design", filter.Category); Assert.True(filter.OldestFirst); Assert.Equal(2024, state.SelectedYear);
    }

    [Fact]
    public void Typed_ChangeSet_And_Evidence_Are_Available_Without_Unsafe_Assumptions()
    {
        var change = new HistoryChangeSet { Changes = [new HistoryChangeItem { FieldKey = "maximum", DisplayLabel = "Annual maximum", ValueType = HistoryValueType.Money, PreviousValue = "$500", NewValue = "$1,000", Unit = "per year" }] };
        var evidence = new HistoryEvidenceReference { EventId = "plan-1", DocumentId = 42, SheetName = "Group History", CellRange = "B13", Confidence = .95m, ValidationStatus = HistoryValidationStatus.Valid };
        Assert.Single(change.Changes); Assert.Equal(HistoryValueType.Money, change.Changes[0].ValueType); Assert.Equal(42, evidence.DocumentId); Assert.Equal("B13", evidence.CellRange);
        var missingEvidence = new ClientHistoryEventReadModel { EventId = "none", StableSourceKey = "Manual:Manual:1:event" };
        Assert.Empty(missingEvidence.Evidence);
        Assert.Empty(missingEvidence.ChangeSet.Changes);
    }

    [Fact]
    public void Temporary_Review_Rule_Is_Only_NeedsReview()
    {
        Assert.True(ShowsReviewActions(HistoryValidationStatus.NeedsReview));
        Assert.False(ShowsReviewActions(HistoryValidationStatus.Valid));
        Assert.False(ShowsReviewActions(HistoryValidationStatus.Warning));
    }

    private static bool ShowsReviewActions(HistoryValidationStatus status) => status == HistoryValidationStatus.NeedsReview;
    private static IReadOnlyList<ClientHistoryEventReadModel> Events() =>
    [
        Event("plan-1", "Plan Design", new DateOnly(2024, 5, 1), false, "Dental coverage"),
        Event("plan-2", "Plan Design", new DateOnly(2025, 1, 1), false, "Dental maximum"),
        Event("carrier-1", "Carrier / Vendor", new DateOnly(2025, 2, 1), false, "Carrier change"),
        Event("excluded", "Plan Design", new DateOnly(2026, 1, 1), true, "Excluded dental item")
    ];
    private static ClientHistoryEventReadModel Event(string id, string category, DateOnly date, bool excluded, string title) => new() { EventId = id, StableSourceKey = id, Category = category, EventType = category, Title = title, Summary = title, EffectiveDate = date, IsExcluded = excluded, Confidence = .9m, ValidationStatus = HistoryValidationStatus.Valid };
}

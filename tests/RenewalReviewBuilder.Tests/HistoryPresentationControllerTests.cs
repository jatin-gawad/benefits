using System.Collections.Generic;
using RenewalReviewBuilder.App.Models.History;
using RenewalReviewBuilder.App.Services.ClientHistory;
using Xunit;

namespace RenewalReviewBuilder.Tests;
public sealed class HistoryPresentationControllerTests
{
    [Fact]
    public void Switches_Modes_And_Drawer_Preserves_Evolution_And_Selection()
    {
        var state=new HistorySelectionState(); var controller=new HistoryPresentationController(state); controller.SetEvents(Events());
        Assert.True(controller.IsTimelineVisible); controller.SelectMode(HistoryMode.Evolution); Assert.True(controller.IsEvolutionVisible); Assert.False(controller.IsTimelineVisible);
        controller.SelectEvolutionStep("vision"); Assert.True(controller.IsDetailsDrawerOpen); Assert.Equal("vision",state.SelectedHistoryEventId); Assert.True(Assert.Single(controller.FilteredEvolutionTracks).Steps[0].IsSelected);
        controller.CloseDrawer(); Assert.True(controller.IsEvolutionVisible); Assert.False(controller.IsDetailsDrawerOpen); controller.SelectMode(HistoryMode.Timeline); Assert.True(controller.IsTimelineVisible); Assert.False(controller.IsAuditVisible);
    }
    [Fact]
    public void Filters_Clear_And_Empty_State_Use_Only_Canonical_Typed_Changes()
    {
        var controller=new HistoryPresentationController(new HistorySelectionState()); controller.SetEvents(Events()); controller.SetFilters("vision","Plan Design",new DateOnly(2024,1,1),new DateOnly(2024,12,31));
        Assert.Single(controller.FilteredEvolutionTracks); Assert.Equal("Vision",controller.FilteredEvolutionTracks[0].Subject); controller.SetFilters("missing","All benefit areas",null,null); Assert.Empty(controller.FilteredEvolutionTracks); controller.ClearFilters(); Assert.Single(controller.FilteredEvolutionTracks); Assert.Equal("",controller.SearchText);
    }
    [Fact]
    public void Selected_Step_Moves_Persists_Across_Drawer_And_Mode_And_Clears()
    {
        var state=new HistorySelectionState(); var controller=new HistoryPresentationController(state); controller.SetEvents([Event("vision"),Event("dental")]); controller.SelectMode(HistoryMode.Evolution);
        controller.SelectEvolutionStep("vision"); Assert.True(Step(controller,"vision").IsSelected); controller.CloseDrawer(); Assert.True(Step(controller,"vision").IsSelected);
        controller.SelectEvolutionStep("dental"); Assert.False(Step(controller,"vision").IsSelected); Assert.True(Step(controller,"dental").IsSelected); controller.SelectMode(HistoryMode.Timeline); controller.SelectMode(HistoryMode.Evolution); Assert.True(Step(controller,"dental").IsSelected);
        controller.ClearSelection(); Assert.All(controller.FilteredEvolutionTracks.SelectMany(track=>track.Steps), step=>Assert.False(step.IsSelected));
    }
    private static HistoryEvolutionStep Step(HistoryPresentationController controller,string id)=>Assert.Single(controller.FilteredEvolutionTracks.SelectMany(track=>track.Steps), step=>step.EventId==id);
    private static ClientHistoryEventReadModel Event(string id)=>new(){EventId=id,StableSourceKey=id,Category="Plan Design",Title=id,EffectiveDate=new DateOnly(2024,5,1),ChangeSet=new HistoryChangeSet{Changes=[new HistoryChangeItem{DisplayLabel=id,PreviousValue="$1",NewValue="$2",ValueType=HistoryValueType.Money}]}};    private static IReadOnlyList<ClientHistoryEventReadModel> Events()=>[new ClientHistoryEventReadModel{EventId="vision",StableSourceKey="vision",Category="Plan Design",Title="Vision update",EffectiveDate=new DateOnly(2024,5,1),ChangeSet=new HistoryChangeSet{Changes=[new HistoryChangeItem{DisplayLabel="Vision",PreviousValue="$200",NewValue="$350",ValueType=HistoryValueType.Money}]}}];
}

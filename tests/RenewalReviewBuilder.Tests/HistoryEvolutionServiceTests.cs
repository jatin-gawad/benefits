using RenewalReviewBuilder.App.Models.History;
using RenewalReviewBuilder.App.Services.ClientHistory;
using Xunit;

namespace RenewalReviewBuilder.Tests;
public sealed class HistoryEvolutionServiceTests
{
    [Fact]
    public void Groups_Typed_Changes_Orders_Steps_And_Preserves_Selection_Identity()
    {
        var events = new[] { Event("carrier", new DateOnly(2024,10,1), "Carrier", "GroupHealth", "Manulife", HistoryValueType.CarrierReference), Event("vision", new DateOnly(2021,8,1), "Vision", "$200", "$350", HistoryValueType.Money), Event("vision-later", new DateOnly(2024,5,1), "Vision", "$350", "$500", HistoryValueType.Money) };
        var tracks = new HistoryEvolutionService().BuildTracks(events);
        var vision = Assert.Single(tracks, track => track.Subject == "Vision");
        Assert.Equal(new[] { "vision", "vision-later" }, vision.Steps.Select(step => step.EventId));
        Assert.Equal("Changed", vision.Steps[0].Direction);
        Assert.Equal("carrier", Assert.Single(tracks, t => t.Subject == "Carrier").Steps[0].EventId);
    }
    [Fact]
    public void Supports_Multiple_Changes_Missing_Previous_And_Financial_Values()
    {
        var e = new ClientHistoryEventReadModel { EventId="financial", StableSourceKey="financial", Category="Financial", EffectiveDate=new DateOnly(2024,1,1), ChangeSet=new HistoryChangeSet { Changes=[new HistoryChangeItem { DisplayLabel="ASO administration fee", NewValue="12%", ValueType=HistoryValueType.Percentage }, new HistoryChangeItem { DisplayLabel="Pooling threshold", PreviousValue="$5,000", NewValue="$7,500", ValueType=HistoryValueType.Money }] } };
        var tracks = new HistoryEvolutionService().BuildTracks([e]);
        Assert.Equal(2, tracks.Count); Assert.Equal("Established", Assert.Single(tracks, t=>t.Subject=="ASO administration fee").Steps[0].Direction); Assert.Equal("$7,500", Assert.Single(tracks,t=>t.Subject=="Pooling threshold").Steps[0].NewValue);
    }
    private static ClientHistoryEventReadModel Event(string id, DateOnly date, string label, string previous, string next, HistoryValueType type) => new() { EventId=id, StableSourceKey=id, EffectiveDate=date, ChangeSet=new HistoryChangeSet { Changes=[new HistoryChangeItem { DisplayLabel=label, PreviousValue=previous, NewValue=next, ValueType=type }] }, Evidence=[new HistoryEvidenceReference { DocumentId=1 }] };
}

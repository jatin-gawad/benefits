using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using RenewalReviewBuilder.App.Models.History;
using RenewalReviewBuilder.App.Models.Querying;
using RenewalReviewBuilder.App.Services.ClientHistory;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class ClientHistoryProjectionResilienceTests
{
    [Fact]
    public void Canonical_Factory_Owns_Evidence_And_Assigns_Final_Event_Id()
    {
        var source = new HistoryEvidenceReference
        {
            EventId = "shared-owner",
            DocumentId = 17,
            DocumentVersionId = 18,
            SourceFactId = "fact-19",
            SourceEntityId = "source-20",
            PageNumber = 21,
            SheetName = "History",
            CellRange = "B22",
            SectionName = "Carrier",
            ExtractedTextReference = "Manulife",
            Confidence = .97m,
            ValidationStatus = HistoryValidationStatus.Valid
        };

        var historyEvent = CanonicalHistoryEventFactory.Create(
            "CarrierVendor",
            "CarrierAssignment",
            "23",
            "assignment",
            24,
            HistoryDateReference.FromExact(new DateOnly(2025, 1, 1)),
            new DateTime(2025, 1, 2),
            "Carrier assignment",
            "Carrier / Vendor",
            "Carrier assigned",
            "Canonical carrier assignment.",
            HistoryValidationStatus.Valid,
            .97m,
            [source],
            stableKey: "CarrierVendor:CarrierAssignment:24:23:assignment");

        var owned = Assert.Single(historyEvent.Evidence);
        Assert.NotSame(source, owned);
        Assert.Equal(historyEvent.EventId, owned.EventId);
        Assert.Equal("shared-owner", source.EventId);
        Assert.Equal(source.DocumentId, owned.DocumentId);
        Assert.Equal(source.DocumentVersionId, owned.DocumentVersionId);
        Assert.Equal(source.SourceFactId, owned.SourceFactId);
        Assert.Equal(source.SourceEntityId, owned.SourceEntityId);
        Assert.Equal(source.PageNumber, owned.PageNumber);
        Assert.Equal(source.SheetName, owned.SheetName);
        Assert.Equal(source.CellRange, owned.CellRange);
        Assert.Equal(source.SectionName, owned.SectionName);
        Assert.Equal(source.ExtractedTextReference, owned.ExtractedTextReference);
        Assert.Equal(source.Confidence, owned.Confidence);
        Assert.Equal(source.ValidationStatus, owned.ValidationStatus);
    }

    [Fact]
    public async Task Projection_Continues_After_Contributor_Failure_And_Reports_Unavailable()
    {
        var expected = Event("surviving-event");
        var builder = new ClientHistoryProjectionBuilder(
        [
            new ThrowingContributor("Claims", new InvalidOperationException("database read failed")),
            new StaticContributor("Renewal", [expected])
        ]);

        var result = await builder.BuildAsOfAsync(new AsOfQueryContext
        {
            ClientId = 31,
            IncludeManualEvents = true,
            IncludeUnvalidatedEvents = true,
            IncludeFutureEvents = true
        });

        Assert.Same(expected, Assert.Single(result.Events));
        var failed = Assert.Single(result.ContributorSupport, item => item.SourceModule == "Claims");
        Assert.Equal(HistoricalQuerySupport.Unavailable, failed.Support);
        Assert.Contains("could not be loaded", failed.Detail, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(
            HistoricalQuerySupport.Supported,
            Assert.Single(result.ContributorSupport, item => item.SourceModule == "Renewal").Support);
    }

    [Fact]
    public async Task Projection_Does_Not_Swallow_Caller_Cancellation()
    {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        var builder = new ClientHistoryProjectionBuilder(
        [
            new CancelingContributor("Claims"),
            new StaticContributor("Renewal", [Event("must-not-load")])
        ]);

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => builder.BuildAsOfAsync(new AsOfQueryContext
        {
            ClientId = 31,
            CancellationToken = cancellation.Token
        }));
    }

    private static ClientHistoryEventReadModel Event(string id) => new()
    {
        EventId = id,
        StableSourceKey = id,
        ClientId = 31,
        SourceModule = "Renewal",
        SourceEntityType = "Renewal",
        SourceEntityId = id,
        EffectiveDateReference = HistoryDateReference.FromYear(2025),
        EventType = "Renewal",
        Category = "Renewals",
        Title = id,
        Confidence = 1m,
        ValidationStatus = HistoryValidationStatus.Valid
    };

    private sealed class StaticContributor(
        string sourceModule,
        IReadOnlyList<ClientHistoryEventReadModel> events) : IClientHistoryContributor
    {
        public string SourceModule => sourceModule;

        public Task<IReadOnlyList<ClientHistoryEventReadModel>> ContributeAsync(
            long clientId,
            CancellationToken cancellationToken = default) => Task.FromResult(events);
    }

    private sealed class ThrowingContributor(
        string sourceModule,
        Exception exception) : IClientHistoryContributor
    {
        public string SourceModule => sourceModule;

        public Task<IReadOnlyList<ClientHistoryEventReadModel>> ContributeAsync(
            long clientId,
            CancellationToken cancellationToken = default) => Task.FromException<IReadOnlyList<ClientHistoryEventReadModel>>(exception);
    }

    private sealed class CancelingContributor(string sourceModule) : IClientHistoryContributor
    {
        public string SourceModule => sourceModule;

        public Task<IReadOnlyList<ClientHistoryEventReadModel>> ContributeAsync(
            long clientId,
            CancellationToken cancellationToken = default) => Task.FromCanceled<IReadOnlyList<ClientHistoryEventReadModel>>(cancellationToken);
    }
}

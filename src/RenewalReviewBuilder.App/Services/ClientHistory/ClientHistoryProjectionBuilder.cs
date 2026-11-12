using System.Threading;
using System.Threading.Tasks;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Models.History;
using RenewalReviewBuilder.App.Models.Querying;

namespace RenewalReviewBuilder.App.Services.ClientHistory;

public sealed class ClientHistoryProjectionBuilder : IClientHistoryProjectionBuilder, IClientHistoryQueryService
{
    private readonly IReadOnlyList<IClientHistoryContributor> _contributors;
    public ClientHistoryProjectionBuilder(IEnumerable<IClientHistoryContributor> contributors) => _contributors = contributors.ToList();

    public static ClientHistoryProjectionBuilder CreateDefault(DatabaseService database) => new(
    [
        new RenewalHistoryContributor(database),
        new PlanDesignHistoryContributor(database),
        new DocumentHistoryContributor(database),
        new ManualHistoryContributor(database),
        new CarrierVendorHistoryContributor(database),
        new FinancialHistoryContributor(database),
        new ProposalMarketHistoryContributor(database),
        new ClaimsHistoryContributor(database),
        new ImplementationHistoryContributor(),
        new MeetingHistoryContributor()
    ]);

    public async Task<IReadOnlyList<ClientHistoryEventReadModel>> BuildAsync(long clientId, CancellationToken cancellationToken = default) =>
        (await BuildAsOfAsync(new AsOfQueryContext { ClientId = clientId, IncludeManualEvents = true, IncludeUnvalidatedEvents = true, IncludeFutureEvents = true, CancellationToken = cancellationToken })).Events;

    public async Task<ClientHistoryAsOfQueryResult> BuildAsOfAsync(AsOfQueryContext context)
    {
        var source = new List<ClientHistoryEventReadModel>();
        var support = new List<HistoricalQuerySupportStatus>();
        foreach (var contributor in _contributors)
        {
            context.CancellationToken.ThrowIfCancellationRequested();
            try
            {
                var contribution = await contributor.ContributeAsync(context);
                context.CancellationToken.ThrowIfCancellationRequested();
                source.AddRange(contribution.Events);
                support.Add(contribution.HistoricalSupport);
            }
            catch (OperationCanceledException) when (context.CancellationToken.IsCancellationRequested)
            {
                throw;
            }
            catch (Exception)
            {
                support.Add(new HistoricalQuerySupportStatus
                {
                    SourceModule = contributor.SourceModule,
                    Support = HistoricalQuerySupport.Unavailable,
                    Detail = $"{contributor.SourceModule} history is unavailable because its canonical contributor could not be loaded."
                });
            }
        }
        var events = source.Where(item => (context.IncludeManualEvents || !item.SourceModule.Equals("Manual", StringComparison.OrdinalIgnoreCase)) && (context.IncludeUnvalidatedEvents || item.ValidationStatus == HistoryValidationStatus.Valid) && (context.IncludeFutureEvents || !context.AsOfDate.HasValue || DateReference(item).IsOnOrBefore(context.AsOfDate.Value)))
            .GroupBy(item => item.StableSourceKey, StringComparer.Ordinal).Select(group => group.OrderByDescending(item => item.UpdatedAt).First()).OrderByDescending(item => DateReference(item).SortYear).ThenByDescending(item => DateReference(item).SortMonth).ThenByDescending(item => DateReference(item).SortDay).ThenByDescending(item => item.RecordedDate).ThenBy(item => item.StableSourceKey, StringComparer.Ordinal).ToList();
        return new ClientHistoryAsOfQueryResult { Context = context, Events = events, ContributorSupport = support };
    }
    public Task<IReadOnlyList<ClientHistoryEventReadModel>> GetAsync(long clientId, CancellationToken cancellationToken = default) => BuildAsync(clientId, cancellationToken);
    public Task<ClientHistoryAsOfQueryResult> GetAsOfAsync(AsOfQueryContext context) => BuildAsOfAsync(context);
    private static HistoryDateReference DateReference(ClientHistoryEventReadModel historyEvent) => historyEvent.EffectiveDateReference.Precision != HistoryDatePrecision.Unknown
        ? historyEvent.EffectiveDateReference
        : historyEvent.EffectiveDate.HasValue ? HistoryDateReference.FromExact(historyEvent.EffectiveDate.Value) : new HistoryDateReference();
    public static string StableKey(string module, string entityType, object entityId, string subtype) => $"{module}:{entityType}:{entityId}:{subtype}";
}

public abstract class DatabaseHistoryContributor : IClientHistoryContributor
{
    protected readonly DatabaseService Database; protected DatabaseHistoryContributor(DatabaseService database) => Database = database;
    public abstract string SourceModule { get; }
    public abstract Task<IReadOnlyList<ClientHistoryEventReadModel>> ContributeAsync(long clientId, CancellationToken cancellationToken = default);
    public virtual async Task<ClientHistoryContribution> ContributeAsync(AsOfQueryContext context) => new()
    {
        Events = await ContributeAsync(context.ClientId, context.CancellationToken),
        HistoricalSupport = new HistoricalQuerySupportStatus { SourceModule = SourceModule, Support = HistoricalQuerySupport.Supported }
    };
    protected static ClientHistoryEventReadModel Event(string module, string type, long id, string subtype, long clientId, DateOnly? effective, DateTime recorded, string eventType, string category, string title, string summary, long? documentId = null, decimal confidence = .90m) => new() { EventId=ClientHistoryProjectionBuilder.StableKey(module,type,id,subtype), StableSourceKey=ClientHistoryProjectionBuilder.StableKey(module,type,id,subtype), ClientId=clientId, SourceModule=module, SourceEntityType=type, SourceEntityId=id.ToString(), EffectiveDate=effective, RecordedDate=recorded, EventType=eventType, Category=category, Title=title, Summary=summary, Evidence=documentId.HasValue ? [new HistoryEvidenceReference { DocumentId=documentId, SourceEntityId=id.ToString(), Confidence=confidence, ValidationStatus=HistoryValidationStatus.Valid }] : [], Confidence=confidence, ValidationStatus=HistoryValidationStatus.Valid, Origin=HistoryOrigin.CanonicalProjection, EffectiveDateReference=effective.HasValue ? HistoryDateReference.FromExact(effective.Value) : new HistoryDateReference(), CreatedAt=recorded, UpdatedAt=recorded };
}

public sealed class RenewalHistoryContributor : DatabaseHistoryContributor
{
    public RenewalHistoryContributor(DatabaseService database) : base(database) { } public override string SourceModule => "Renewal";
    public override async Task<IReadOnlyList<ClientHistoryEventReadModel>> ContributeAsync(long clientId, CancellationToken cancellationToken = default)
    {
        var history = await Database.GetClientHistoryAsync(clientId);
        var legacyRecords = history.Renewals.Select(record =>
        {
            var date = record.RenewalYear > 0 ? HistoryDateReference.FromYear(record.RenewalYear) : new HistoryDateReference();
            return new ClientHistoryEventReadModel
            {
                EventId = ClientHistoryProjectionBuilder.StableKey(SourceModule, "Renewal", record.RenewalId, "record"),
                StableSourceKey = ClientHistoryProjectionBuilder.StableKey(SourceModule, "Renewal", record.RenewalId, "record"),
                ClientId = clientId,
                SourceModule = SourceModule,
                SourceEntityType = "Renewal",
                SourceEntityId = record.RenewalId.ToString(),
                EffectiveDateReference = date,
                RecordedDate = record.DateLoaded,
                EventType = "Renewal",
                Category = "Renewals",
                Title = record.RenewalYear > 0 ? $"{record.RenewalYear} renewal" : "Renewal",
                Summary = $"{record.Carrier} {record.RenewalStatus}".Trim(),
                Evidence = record.SourceDocumentId.HasValue ? [new HistoryEvidenceReference { DocumentId = record.SourceDocumentId, SourceEntityId = record.RenewalId.ToString(), Confidence = .90m, ValidationStatus = HistoryValidationStatus.Valid }] : [],
                Confidence = .90m,
                ValidationStatus = HistoryValidationStatus.Valid,
                Origin = HistoryOrigin.CanonicalProjection,
                CreatedAt = record.DateLoaded,
                UpdatedAt = record.DateLoaded
            };
        });

        var versionRecords = history.RenewalVersions.Select(version =>
        {
            var date = version.EffectiveDate.HasValue
                ? HistoryDateReference.FromExact(version.EffectiveDate.Value)
                : version.RenewalYear.HasValue ? HistoryDateReference.FromYear(version.RenewalYear.Value) : new HistoryDateReference();
            var eventId = ClientHistoryProjectionBuilder.StableKey(SourceModule, "RenewalVersion", version.VersionId, "version");
            return new ClientHistoryEventReadModel
            {
                EventId = eventId,
                StableSourceKey = eventId,
                ClientId = clientId,
                SourceModule = SourceModule,
                SourceEntityType = "RenewalVersion",
                SourceEntityId = version.VersionId.ToString(),
                EffectiveDate = version.EffectiveDate,
                EffectiveDateReference = date,
                RecordedDate = version.DateCreated,
                EventType = version.IsSelected ? "Current renewal" : "Renewal",
                Category = "Renewals",
                Title = version.VersionName,
                Summary = version.Status,
                RelatedRecords = version.RenewalId.HasValue ? [$"Renewal:{version.RenewalId.Value}"] : [],
                Evidence = version.SourceDocumentId.HasValue ? [new HistoryEvidenceReference { DocumentId = version.SourceDocumentId, SourceEntityId = version.VersionId.ToString(), Confidence = .90m, ValidationStatus = HistoryValidationStatus.Valid }] : [],
                Confidence = .90m,
                ValidationStatus = HistoryValidationStatus.Valid,
                Origin = HistoryOrigin.CanonicalProjection,
                CreatedAt = version.DateCreated,
                UpdatedAt = version.DateCreated
            };
        });

        return legacyRecords.Concat(versionRecords).ToList();
    }
}
public sealed class DocumentHistoryContributor : DatabaseHistoryContributor
{
    public DocumentHistoryContributor(DatabaseService database) : base(database) { } public override string SourceModule => "Document";
    public override async Task<IReadOnlyList<ClientHistoryEventReadModel>> ContributeAsync(long clientId, CancellationToken cancellationToken = default) { var h=await Database.GetClientHistoryAsync(clientId); return h.Documents.Select(d => Event(SourceModule,"Document",d.DocumentId,"received",clientId,d.EffectiveDate,d.ProcessedDate,"Document received",d.ModuleName,d.FileName,d.DocumentType,d.DocumentId,.80m)).ToList(); }
}
public sealed class ManualHistoryContributor : DatabaseHistoryContributor
{
    public ManualHistoryContributor(DatabaseService database) : base(database) { }
    public override string SourceModule => "Manual";

    public override async Task<IReadOnlyList<ClientHistoryEventReadModel>> ContributeAsync(long clientId, CancellationToken cancellationToken = default)
    {
        var history = await Database.GetClientHistoryAsync(clientId);
        return history.Events.Select(historyEvent =>
        {
            var eventId = ClientHistoryProjectionBuilder.StableKey(SourceModule, "ImportedHistory", historyEvent.HistoryEventId, "event");
            var hasCanonicalChange = !string.IsNullOrWhiteSpace(historyEvent.PreviousValue) ||
                                     !string.IsNullOrWhiteSpace(historyEvent.NewValue);
            var changes = hasCanonicalChange
                ? new HistoryChangeSet
                {
                    Changes =
                    [
                        new HistoryChangeItem
                        {
                            FieldKey = historyEvent.ChangeType,
                            DisplayLabel = string.IsNullOrWhiteSpace(historyEvent.ChangeType) ? historyEvent.Title : historyEvent.ChangeType,
                            PreviousValue = historyEvent.PreviousValue,
                            NewValue = historyEvent.NewValue
                        }
                    ]
                }
                : new HistoryChangeSet();
            var evidence = !string.IsNullOrWhiteSpace(historyEvent.SourceSheet) || historyEvent.SourceRow > 0 || historyEvent.SourceDocumentId.HasValue
                ? new[]
                {
                    new HistoryEvidenceReference
                    {
                        EventId = eventId,
                        DocumentId = historyEvent.SourceDocumentId,
                        SourceEntityId = historyEvent.HistoryEventId.ToString(),
                        SheetName = historyEvent.SourceSheet,
                        CellRange = historyEvent.SourceRow > 0 ? $"Row {historyEvent.SourceRow}" : "",
                        ExtractedTextReference = historyEvent.OriginalText,
                        Confidence = historyEvent.Confidence,
                        ValidationStatus = IsApproved(historyEvent) ? HistoryValidationStatus.Valid : HistoryValidationStatus.NeedsReview
                    }
                }
                : [];

            return new ClientHistoryEventReadModel
            {
                EventId = eventId,
                StableSourceKey = eventId,
                ClientId = clientId,
                SourceModule = SourceModule,
                SourceEntityType = "ImportedHistory",
                SourceEntityId = historyEvent.HistoryEventId.ToString(),
                EffectiveDate = historyEvent.EffectiveDate,
                EffectiveDateReference = historyEvent.EffectiveDate.HasValue
                    ? HistoryDateReference.FromExact(historyEvent.EffectiveDate.Value)
                    : new HistoryDateReference(),
                RecordedDate = historyEvent.CreatedAt,
                EventType = historyEvent.EventType,
                Category = Category(historyEvent),
                Title = historyEvent.Title,
                Summary = historyEvent.Description,
                ChangeSet = changes,
                Evidence = evidence,
                Confidence = historyEvent.Confidence,
                ValidationStatus = IsApproved(historyEvent) ? HistoryValidationStatus.Valid : HistoryValidationStatus.NeedsReview,
                Origin = HistoryOrigin.ManualEvent,
                CreatedAt = historyEvent.CreatedAt,
                UpdatedAt = historyEvent.UpdatedAt
            };
        }).ToList();
    }

    private static bool IsApproved(ClientHistoryEventRecord historyEvent) =>
        historyEvent.ReviewStatus.Equals(ClientHistoryPreviewStatuses.Imported, StringComparison.OrdinalIgnoreCase) ||
        historyEvent.ReviewStatus.Equals(ClientHistoryPreviewStatuses.Approved, StringComparison.OrdinalIgnoreCase);

    private static string Category(ClientHistoryEventRecord historyEvent)
    {
        if (historyEvent.EventType.Contains("Carrier", StringComparison.OrdinalIgnoreCase)) return "Carrier / Vendor";
        if (historyEvent.EventType.Contains("Benefit Plan", StringComparison.OrdinalIgnoreCase)) return "Plan Design";
        if (historyEvent.EventType.Contains("Market", StringComparison.OrdinalIgnoreCase)) return "Market";
        if (historyEvent.EventType.Contains("Meeting", StringComparison.OrdinalIgnoreCase)) return "Administration";
        return string.IsNullOrWhiteSpace(historyEvent.BenefitCategory) ? "Other" : historyEvent.BenefitCategory;
    }
}

public sealed class ImplementationHistoryContributor : EmptyHistoryContributor { public ImplementationHistoryContributor() : base("Implementation", HistoricalQuerySupport.Unavailable, "Implementation history is not available from a canonical contributor.") {} }
public sealed class MeetingHistoryContributor : EmptyHistoryContributor { public MeetingHistoryContributor() : base("Meeting", HistoricalQuerySupport.Unavailable, "Meeting history is not available from a canonical contributor.") {} }
public abstract class EmptyHistoryContributor : IClientHistoryContributor
{
    private readonly HistoricalQuerySupport _support; private readonly string _detail;
    protected EmptyHistoryContributor(string source, HistoricalQuerySupport support = HistoricalQuerySupport.Unavailable, string detail = "No canonical historical contributor is available.") { SourceModule=source; _support=support; _detail=detail; }
    public string SourceModule { get; }
    public Task<IReadOnlyList<ClientHistoryEventReadModel>> ContributeAsync(long clientId, CancellationToken cancellationToken = default) => Task.FromResult<IReadOnlyList<ClientHistoryEventReadModel>>([]);
    public Task<ClientHistoryContribution> ContributeAsync(AsOfQueryContext context) => Task.FromResult(new ClientHistoryContribution { HistoricalSupport = new HistoricalQuerySupportStatus { SourceModule = SourceModule, Support = _support, Detail = _detail } });
}

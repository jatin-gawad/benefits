using System.Threading;
using System.Threading.Tasks;
using RenewalReviewBuilder.App.Models.History;
using RenewalReviewBuilder.App.Models.Querying;

namespace RenewalReviewBuilder.App.Services.ClientHistory;
public sealed class ClientHistoryContribution
{
    public IReadOnlyList<ClientHistoryEventReadModel> Events { get; init; } = [];
    public HistoricalQuerySupportStatus HistoricalSupport { get; init; } = new();
}
public sealed class ClientHistoryAsOfQueryResult
{
    public AsOfQueryContext Context { get; init; } = new();
    public IReadOnlyList<ClientHistoryEventReadModel> Events { get; init; } = [];
    public IReadOnlyList<HistoricalQuerySupportStatus> ContributorSupport { get; init; } = [];
}
public interface IClientHistoryContributor
{
    string SourceModule { get; }
    Task<IReadOnlyList<ClientHistoryEventReadModel>> ContributeAsync(long clientId, CancellationToken cancellationToken = default);
    async Task<ClientHistoryContribution> ContributeAsync(AsOfQueryContext context) => new()
    {
        Events = await ContributeAsync(context.ClientId, context.CancellationToken),
        HistoricalSupport = new HistoricalQuerySupportStatus { SourceModule = SourceModule, Support = HistoricalQuerySupport.Supported }
    };
}
public interface IClientHistoryProjectionBuilder
{
    Task<IReadOnlyList<ClientHistoryEventReadModel>> BuildAsync(long clientId, CancellationToken cancellationToken = default);
    async Task<ClientHistoryAsOfQueryResult> BuildAsOfAsync(AsOfQueryContext context) => new()
    {
        Context = context,
        Events = await BuildAsync(context.ClientId, context.CancellationToken)
    };
}
public interface IClientHistoryQueryService
{
    Task<IReadOnlyList<ClientHistoryEventReadModel>> GetAsync(long clientId, CancellationToken cancellationToken = default);
    Task<ClientHistoryAsOfQueryResult> GetAsOfAsync(AsOfQueryContext context);
}
public interface IClientHistoryFilterService { IReadOnlyList<ClientHistoryEventReadModel> Filter(IEnumerable<ClientHistoryEventReadModel> events, HistorySelectionState selection); }
public interface IClientHistorySelectionService { HistorySelectionState Selection { get; } }
public interface IHistoryEvidenceService { IReadOnlyList<HistoryEvidenceReference> GetEvidence(ClientHistoryEventReadModel historyEvent); }
public interface IHistoryValueCalculationService { string CalculateFinancialImpact(ClientHistoryEventReadModel historyEvent); }
public interface IHistoryDeduplicationService { IReadOnlyList<ClientHistoryEventReadModel> Deduplicate(IEnumerable<ClientHistoryEventReadModel> events); }
public interface IHistoryEvolutionService { IReadOnlyList<ClientHistoryEventReadModel> BuildEvolution(IEnumerable<ClientHistoryEventReadModel> events); }
public interface IHistoryRenewalEvolutionService { IReadOnlyList<ClientHistoryEventReadModel> BuildRenewalEvolution(IEnumerable<ClientHistoryEventReadModel> events); }
public interface IManualHistoryEventService { Task<IReadOnlyList<ManualClientHistoryEvent>> GetAsync(long clientId, CancellationToken cancellationToken = default); }
public interface IHistoryReviewService { Task<IReadOnlyList<string>> GetMigrationReviewMessagesAsync(long clientId, CancellationToken cancellationToken = default); }
public interface IHistoryPermissionService { bool CanReview(ClientHistoryEventReadModel historyEvent); }

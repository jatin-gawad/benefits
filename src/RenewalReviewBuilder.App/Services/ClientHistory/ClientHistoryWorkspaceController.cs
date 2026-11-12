using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Threading;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Models.History;
using RenewalReviewBuilder.App.Models.Querying;

namespace RenewalReviewBuilder.App.Services.ClientHistory;

/// <summary>
/// Keeps compact History workspace state testable while sharing the one HistorySelectionState.
/// </summary>
public sealed class ClientHistoryWorkspaceController : INotifyPropertyChanged, IDisposable
{
    private readonly IClientHistoryProjectionBuilder _projectionBuilder;
    private readonly Func<long, CancellationToken, Task<IReadOnlyList<RenewalVersionRecord>>> _renewalVersionsLoader;
    private readonly HistoryRelationshipTimelineService _relationshipTimeline;
    private readonly HistoryRenewalEvolutionService _renewalEvolution;
    private CancellationTokenSource? _loadCancellation;
    private IReadOnlyList<ClientHistoryEventReadModel> _events = [];
    private IReadOnlyList<RenewalVersionRecord> _renewalVersions = [];
    private long? _clientId;
    private bool _isRelationshipTimelineLoading;
    private bool _isRenewalEvolutionLoading;
    private string _relationshipTimelineEmptyState = "";
    private string _renewalEvolutionEmptyState = "";
    private string _relationshipTimelineError = "";
    private string _renewalEvolutionError = "";
    private IReadOnlyList<HistoricalQuerySupportStatus> _replaySupport = [];
    private ClientHistoryKpis _kpis = new();
    private bool _isReplayLoading;
    private string _replayError = "";

    public ClientHistoryWorkspaceController(
        HistorySelectionState selection,
        IClientHistoryProjectionBuilder projectionBuilder,
        Func<long, CancellationToken, Task<IReadOnlyList<RenewalVersionRecord>>> renewalVersionsLoader,
        HistoryRelationshipTimelineService? relationshipTimeline = null,
        HistoryRenewalEvolutionService? renewalEvolution = null)
    {
        Selection = selection;
        _projectionBuilder = projectionBuilder;
        _renewalVersionsLoader = renewalVersionsLoader;
        _relationshipTimeline = relationshipTimeline ?? new HistoryRelationshipTimelineService();
        _renewalEvolution = renewalEvolution ?? new HistoryRenewalEvolutionService();
        Presentation = new HistoryPresentationController(selection);
        Presentation.PropertyChanged += Presentation_PropertyChanged;
        Selection.PropertyChanged += (_, _) => RefreshCompactComponents();
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    public HistorySelectionState Selection { get; }
    public HistoryPresentationController Presentation { get; }
    public long? DatabaseClientId => _clientId;
    public IReadOnlyList<ClientHistoryEventReadModel> Events => _events;
    public ObservableCollection<HistoryRelationshipMilestone> RelationshipMilestones { get; } = [];
    public ObservableCollection<HistoryRenewalEvolutionPoint> RenewalPoints { get; } = [];
    public bool IsRelationshipTimelineLoading { get => _isRelationshipTimelineLoading; private set => Set(ref _isRelationshipTimelineLoading, value); }
    public bool IsRenewalEvolutionLoading { get => _isRenewalEvolutionLoading; private set => Set(ref _isRenewalEvolutionLoading, value); }
    public string RelationshipTimelineEmptyState { get => _relationshipTimelineEmptyState; private set => Set(ref _relationshipTimelineEmptyState, value); }
    public string RenewalEvolutionEmptyState { get => _renewalEvolutionEmptyState; private set => Set(ref _renewalEvolutionEmptyState, value); }
    public string RelationshipTimelineError { get => _relationshipTimelineError; private set => Set(ref _relationshipTimelineError, value); }
    public string RenewalEvolutionError { get => _renewalEvolutionError; private set => Set(ref _renewalEvolutionError, value); }
    public string RelationshipTimelineStatusText => IsRelationshipTimelineLoading
        ? "Loading relationship milestones…"
        : !string.IsNullOrWhiteSpace(RelationshipTimelineError)
            ? RelationshipTimelineError
            : RelationshipTimelineEmptyState;
    public IReadOnlyList<HistoricalQuerySupportStatus> ReplaySupport => _replaySupport;
    public ClientHistoryKpis Kpis { get => _kpis; private set => Set(ref _kpis, value); }
    public bool IsReplayLoading { get => _isReplayLoading; private set => Set(ref _isReplayLoading, value); }
    public string ReplayError { get => _replayError; private set => Set(ref _replayError, value); }
    public bool IsReplayActive => Selection.IsReplayActive;
    public DateOnly? ReplayAsOfDate => Selection.ReplayAsOfDate;
    public bool IsCurrentProjection => !IsReplayActive;
    public string ReplayDisplayLabel => IsReplayActive && ReplayAsOfDate.HasValue ? $"As of {ReplayAsOfDate.Value:MMM d, yyyy}" : "Current";
    public string CurrentCarrierDisplay
    {
        get
        {
            if (!IsReplayActive) return Kpis.CurrentCarrier;

            var carrierSupport = ReplaySupport
                .Where(item => item.SourceModule.Equals("CarrierVendor", StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (carrierSupport.Any(item => item.Support == HistoricalQuerySupport.Unavailable))
            {
                return "Unavailable for Replay";
            }

            if (carrierSupport.Any(item => item.Support == HistoricalQuerySupport.Limited))
            {
                return "Limited for Replay";
            }

            return Kpis.CurrentCarrier;
        }
    }
    public string ReplaySupportSummary => string.Join(" ", ReplaySupport.Where(item => item.Support != HistoricalQuerySupport.Supported).Select(item => item.Detail).Where(item => !string.IsNullOrWhiteSpace(item)));
    public string AverageRenewalIncreaseSubtitle => Kpis.CompletedRenewalCount == 0
        ? "No completed renewals"
        : $"Across {Kpis.CompletedRenewalCount} completed renewal{(Kpis.CompletedRenewalCount == 1 ? "" : "s")}";
    public string RenewalEvolutionStatusText => IsRenewalEvolutionLoading
        ? "Loading renewal evolution…"
        : !string.IsNullOrWhiteSpace(RenewalEvolutionError)
            ? RenewalEvolutionError
            : RenewalEvolutionEmptyState;

    public async Task LoadAsync(long clientId, CancellationToken cancellationToken = default)
    {
        _loadCancellation?.Cancel();
        _loadCancellation?.Dispose();
        _loadCancellation = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        var token = _loadCancellation.Token;
        _clientId = clientId;
        OnPropertyChanged(nameof(DatabaseClientId));
        IsRelationshipTimelineLoading = true;
        IsRenewalEvolutionLoading = true;
        RelationshipTimelineError = "";
        RenewalEvolutionError = "";
        RelationshipTimelineEmptyState = "";
        RenewalEvolutionEmptyState = "";

        try
        {
            IReadOnlyList<ClientHistoryEventReadModel> loadedEvents;
            IReadOnlyList<HistoricalQuerySupportStatus> loadedSupport;
            if (Selection.IsReplayActive && Selection.ReplayAsOfDate.HasValue)
            {
                var result = await _projectionBuilder.BuildAsOfAsync(new AsOfQueryContext { ClientId = clientId, AsOfDate = Selection.ReplayAsOfDate, IncludeManualEvents = true, IncludeUnvalidatedEvents = true, CancellationToken = token });
                loadedEvents = result.Events;
                loadedSupport = result.ContributorSupport;
            }
            else
            {
                loadedEvents = await _projectionBuilder.BuildAsync(clientId, token);
                loadedSupport = [];
            }
            if (token.IsCancellationRequested) return;
            _events = loadedEvents;
            _replaySupport = loadedSupport;
            if (!string.IsNullOrWhiteSpace(Selection.SelectedHistoryEventId) && !_events.Any(item => item.EventId == Selection.SelectedHistoryEventId)) Selection.ClearSelection();
            Presentation.SetEvents(_events);
            Kpis = new ClientHistoryTimelineService().CalculateKpis(_events, Selection.ReplayAsOfDate ?? DateOnly.FromDateTime(DateTime.Today)); OnPropertyChanged(nameof(CurrentCarrierDisplay)); OnPropertyChanged(nameof(ReplaySupport)); OnPropertyChanged(nameof(IsReplayActive)); OnPropertyChanged(nameof(ReplayAsOfDate)); OnPropertyChanged(nameof(ReplayDisplayLabel)); OnPropertyChanged(nameof(ReplaySupportSummary)); OnPropertyChanged(nameof(IsCurrentProjection));
            OnPropertyChanged(nameof(Events));
            RefreshRelationshipMilestones();
        }
        catch (OperationCanceledException) when (token.IsCancellationRequested)
        {
            return;
        }
        catch (Exception ex)
        {
            _events = [];
            Presentation.SetEvents(_events);
            RelationshipMilestones.Clear();
            RelationshipTimelineError = $"Relationship timeline could not be loaded: {ex.Message}";
            OnPropertyChanged(nameof(Events));
        }
        finally
        {
            IsRelationshipTimelineLoading = false;
        }

        try
        {
            var loadedRenewalVersions = await _renewalVersionsLoader(clientId, token);
            if (token.IsCancellationRequested) return;
            _renewalVersions = loadedRenewalVersions;
            RefreshRenewalPoints();
            RefreshKpis();
        }
        catch (OperationCanceledException) when (token.IsCancellationRequested)
        {
            return;
        }
        catch (Exception ex)
        {
            _renewalVersions = [];
            RenewalPoints.Clear();
            RenewalEvolutionError = $"Renewal evolution could not be loaded: {ex.Message}";
        }
        finally
        {
            IsRenewalEvolutionLoading = false;
        }
    }

    public async Task ActivateReplayAsync(DateOnly asOfDate, CancellationToken cancellationToken = default) { Selection.ReplayAsOfDate = asOfDate; Selection.IsReplayActive = true; await RefreshAsync(cancellationToken); }
    public async Task ExitReplayAsync(CancellationToken cancellationToken = default) { Selection.ReplayAsOfDate = null; Selection.IsReplayActive = false; await RefreshAsync(cancellationToken); }
    public Task GoToCurrentAsync(CancellationToken cancellationToken = default) => ExitReplayAsync(cancellationToken);
    public Task RefreshAsync(CancellationToken cancellationToken = default) => _clientId.HasValue
        ? LoadAsync(_clientId.Value, cancellationToken)
        : Task.CompletedTask;

    public void SetFilters(string? searchText, string? category, DateOnly? dateFrom, DateOnly? dateTo)
    {
        Presentation.SetFilters(searchText, category, dateFrom, dateTo);
        RefreshCompactComponents();
    }

    public void ClearFilters()
    {
        Presentation.ClearFilters();
        RefreshCompactComponents();
    }

    public bool SelectRelationshipMilestone(HistoryRelationshipMilestone milestone) => SelectHistoryEvent(milestone.EventId);

    public bool SelectRenewalPoint(HistoryRenewalEvolutionPoint point)
    {
        Selection.SelectedRenewalId = point.RenewalId;
        Selection.SelectedYear = point.DateReference.Year;
        Selection.SelectedCarrierId = point.CarrierId;
        return string.IsNullOrWhiteSpace(point.EventId)
            ? false
            : Presentation.SelectHistoryEvent(point.EventId, point.RenewalId, point.CarrierId, point.DateReference.Year);
    }

    public bool SelectHistoryEvent(string eventId)
    {
        var historyEvent = _events.FirstOrDefault(item => item.EventId == eventId);
        if (historyEvent is null)
        {
            return false;
        }

        return Presentation.SelectHistoryEvent(
            eventId,
            HistorySelectionReferenceResolver.ResolveRenewalId(historyEvent.RelatedRecords),
            HistorySelectionReferenceResolver.ResolveCarrierId(historyEvent.RelatedRecords),
            historyEvent.EffectiveDateReference.Year ?? historyEvent.EffectiveDate?.Year);
    }

    /// <summary>
    /// Selects a canonical event through the shared selection state and ensures the existing drawer is open.
    /// This remains idempotent when the same visible event is chosen again.
    /// </summary>
    public bool OpenHistoryEventDetails(string eventId)
    {
        if (!SelectHistoryEvent(eventId))
        {
            return false;
        }

        Presentation.OpenDrawer();
        return true;
    }

    public void Dispose()
    {
        _loadCancellation?.Cancel();
        _loadCancellation?.Dispose();
    }

    private void Presentation_PropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName is nameof(HistoryPresentationController.FilteredEvents) or nameof(HistoryPresentationController.SearchText) or nameof(HistoryPresentationController.SelectedCategory) or nameof(HistoryPresentationController.DateFrom) or nameof(HistoryPresentationController.DateTo))
        {
            RefreshCompactComponents();
        }
    }

    private void RefreshCompactComponents()
    {
        RefreshRelationshipMilestones();
        RefreshRenewalPoints();
    }

    private void RefreshRelationshipMilestones()
    {
        RelationshipMilestones.Clear();
        if (_events.Count == 0)
        {
            RelationshipTimelineEmptyState = "No canonical history is available for this client.";
            return;
        }

        foreach (var milestone in _relationshipTimeline.BuildMilestones(Presentation.FilteredEvents))
        {
            RelationshipMilestones.Add(new HistoryRelationshipMilestone { EventId = milestone.EventId, EffectiveDate = milestone.EffectiveDate, DateReference = milestone.DateReference, Label = milestone.Label, IconGlyph = milestone.IconGlyph, AccentColor = milestone.AccentColor, AccentBackground = milestone.AccentBackground, ScreenReaderDescription = milestone.ScreenReaderDescription, IsSelected = milestone.EventId == Selection.SelectedHistoryEventId });
        }

        RelationshipTimelineEmptyState = RelationshipMilestones.Count == 0
            ? "No strategic relationship milestones match the current history context."
            : "";
    }

    private void RefreshRenewalPoints()
    {
        RenewalPoints.Clear();
        var versions = FilteredRenewalVersions();
        foreach (var point in _renewalEvolution.Build(versions))
        {
            RenewalPoints.Add(new HistoryRenewalEvolutionPoint { EventId = point.EventId, RenewalId = point.RenewalId, CarrierId = point.CarrierId, DateReference = point.DateReference, RequestedIncrease = point.RequestedIncrease, NegotiatedIncrease = point.NegotiatedIncrease, FinalIncrease = point.FinalIncrease, IndustryComparison = point.IndustryComparison, NegotiatedSavingsAnnual = point.NegotiatedSavingsAnnual, IsCarrierChange = point.IsCarrierChange, IsCurrent = point.IsCurrent, RenewalStatus = point.RenewalStatus, DetailDescription = point.DetailDescription, IsSelected = (point.RenewalId.HasValue && point.RenewalId == Selection.SelectedRenewalId) || (!string.IsNullOrWhiteSpace(point.EventId) && point.EventId == Selection.SelectedHistoryEventId) });
        }

        RenewalEvolutionEmptyState = _renewalVersions.Count == 0
            ? "No canonical renewal versions are available for this client."
            : RenewalPoints.Count == 0
                ? "No renewal versions match the current history context."
                : RenewalPoints.Count == 1
                    ? "Only one canonical renewal version is available."
                    : "";
    }

    private void RefreshKpis()
    {
        var baseKpis = new ClientHistoryTimelineService().CalculateKpis(_events, Selection.ReplayAsOfDate ?? DateOnly.FromDateTime(DateTime.Today));
        var completedRenewals = ReplayVisibleRenewalVersions()
            .Where(IsCompletedRenewalForKpi)
            .Select(item => new { Version = item, Identity = CompletedRenewalIdentity(item) })
            .Where(item => !string.IsNullOrWhiteSpace(item.Identity))
            .GroupBy(item => item.Identity, StringComparer.Ordinal)
            .Select(group => group
                .Select(item => item.Version)
                .OrderByDescending(item => item.Status.Trim().Equals("Completed", StringComparison.OrdinalIgnoreCase))
                .ThenByDescending(item => item.RenewalCompletedDate ?? item.RenewalAcceptedDate ?? item.EffectiveDate ?? DateOnly.MinValue)
                .ThenByDescending(item => item.DateCreated)
                .ThenByDescending(item => item.VersionId)
                .First())
            .ToList();
        Kpis = new ClientHistoryKpis
        {
            YearsAsClient = baseKpis.YearsAsClient,
            TotalValueCreated = baseKpis.TotalValueCreated,
            BenefitImprovementCount = baseKpis.BenefitImprovementCount,
            SuccessfulNegotiationCount = baseKpis.SuccessfulNegotiationCount,
            CurrentCarrier = baseKpis.CurrentCarrier,
            CarrierEffectiveDate = baseKpis.CarrierEffectiveDate,
            CompletedRenewalCount = completedRenewals.Count,
            AverageRenewalIncrease = completedRenewals.Count == 0
                ? null
                : completedRenewals.Average(item => item.FinalAcceptedIncreasePercent!.Value)
        };
        OnPropertyChanged(nameof(CurrentCarrierDisplay));
        OnPropertyChanged(nameof(AverageRenewalIncreaseSubtitle));
    }

    private static bool IsCompletedRenewalForKpi(RenewalVersionRecord version) =>
        version.FinalAcceptedIncreasePercent.HasValue &&
        (version.Status.Trim().Equals("Completed", StringComparison.OrdinalIgnoreCase) ||
         version.Status.Trim().Equals("Final Approved", StringComparison.OrdinalIgnoreCase));

    private static string CompletedRenewalIdentity(RenewalVersionRecord version)
    {
        if (version.RenewalId.HasValue) return $"renewal:{version.RenewalId.Value}";
        if (version.EffectiveDate.HasValue) return $"date:{version.EffectiveDate.Value:yyyy-MM-dd}";
        return version.RenewalYear.HasValue ? $"year:{version.RenewalYear.Value}" : "";
    }

    private IReadOnlyList<RenewalVersionRecord> ReplayVisibleRenewalVersions()
    {
        IEnumerable<RenewalVersionRecord> visible = _renewalVersions;
        if (Selection.IsReplayActive && Selection.ReplayAsOfDate.HasValue)
        {
            var asOfDate = Selection.ReplayAsOfDate.Value;
            visible = visible.Where(version => RenewalDateReference(version).IsOnOrBefore(asOfDate));
        }
        return visible.ToList();
    }

    private IReadOnlyList<RenewalVersionRecord> FilteredRenewalVersions()
    {
        var replayVisible = ReplayVisibleRenewalVersions();
        if (!Presentation.HasActiveFilters)
        {
            return replayVisible;
        }

        var visibleVersionIds = Presentation.FilteredEvents
            .Where(item => item.SourceModule.Equals("Renewal", StringComparison.OrdinalIgnoreCase) && item.SourceEntityType.Equals("RenewalVersion", StringComparison.OrdinalIgnoreCase))
            .Select(item => long.TryParse(item.SourceEntityId, out var versionId) ? versionId : (long?)null)
            .Where(versionId => versionId.HasValue)
            .Select(versionId => versionId!.Value)
            .ToHashSet();

        return replayVisible.Where(version => visibleVersionIds.Contains(version.VersionId)).ToList();
    }

    private static HistoryDateReference RenewalDateReference(RenewalVersionRecord version) =>
        version.EffectiveDate.HasValue
            ? HistoryDateReference.FromExact(version.EffectiveDate.Value)
            : version.RenewalYear.HasValue
                ? HistoryDateReference.FromYear(version.RenewalYear.Value)
                : new HistoryDateReference();
    private void Set<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value))
        {
            return;
        }

        field = value;
        OnPropertyChanged(propertyName);
        if (propertyName is nameof(IsRelationshipTimelineLoading) or nameof(RelationshipTimelineError) or nameof(RelationshipTimelineEmptyState))
        {
            OnPropertyChanged(nameof(RelationshipTimelineStatusText));
        }
        if (propertyName is nameof(IsRenewalEvolutionLoading) or nameof(RenewalEvolutionError) or nameof(RenewalEvolutionEmptyState))
        {
            OnPropertyChanged(nameof(RenewalEvolutionStatusText));
        }
    }

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
}

public static class HistorySelectionReferenceResolver
{
    public static long? ResolveRenewalId(IEnumerable<string> relatedRecords) => ResolveId(relatedRecords, "Renewal:");
    public static long? ResolveCarrierId(IEnumerable<string> relatedRecords) => ResolveId(relatedRecords, "Carrier:");

    private static long? ResolveId(IEnumerable<string> relatedRecords, string prefix)
    {
        var value = relatedRecords.FirstOrDefault(item => item.StartsWith(prefix, StringComparison.OrdinalIgnoreCase));
        return value is not null && long.TryParse(value[prefix.Length..], out var id) ? id : null;
    }
}

using System.ComponentModel;
using System.Runtime.CompilerServices;
using RenewalReviewBuilder.App.Models.History;

namespace RenewalReviewBuilder.App.Services.ClientHistory;

public sealed class HistoryPresentationController : INotifyPropertyChanged
{
    private readonly HistorySelectionState _selection;
    private readonly HistoryEvolutionService _evolution = new();
    private IReadOnlyList<ClientHistoryEventReadModel> _events = [];
    private string _searchText = "";
    private string _category = "All benefit areas";
    private DateOnly? _from;
    private DateOnly? _to;

    public event PropertyChangedEventHandler? PropertyChanged;
    public IReadOnlyList<ClientHistoryEventReadModel> FilteredEvents { get; private set; } = [];
    public IReadOnlyList<HistoryEvolutionTrack> FilteredEvolutionTracks { get; private set; } = [];
    public IReadOnlyList<HistoryEvolutionActivityItem> EvolutionActivityEvents { get; private set; } = [];
    public HistorySelectionState Selection => _selection;
    public HistoryMode ActiveHistoryMode => _selection.ActiveHistoryMode;
    public bool IsTimelineVisible => ActiveHistoryMode == HistoryMode.Timeline;
    public bool IsEvolutionVisible => ActiveHistoryMode == HistoryMode.Evolution;
    public bool IsAuditVisible => ActiveHistoryMode == HistoryMode.Audit;
    public bool IsDetailsDrawerOpen => _selection.IsDetailsDrawerOpen;
    public string SelectedHistoryEventId => _selection.SelectedHistoryEventId;
    public string SelectedEvolutionStepId => _selection.SelectedEvolutionStepId;
    public string SearchText => _searchText;
    public string SelectedCategory => _category;
    public DateOnly? DateFrom => _from;
    public DateOnly? DateTo => _to;
    public bool HasActiveFilters =>
        !string.IsNullOrWhiteSpace(_searchText) ||
        !_category.Equals("All benefit areas", StringComparison.OrdinalIgnoreCase) ||
        _from.HasValue ||
        _to.HasValue;

    public HistoryPresentationController(HistorySelectionState selection)
    {
        _selection = selection;
        _selection.PropertyChanged += (_, _) =>
        {
            Rebuild();
            NotifyState();
        };
    }

    public void SetEvents(IReadOnlyList<ClientHistoryEventReadModel> events)
    {
        _events = events;
        Rebuild();
    }

    public void SelectMode(HistoryMode mode)
    {
        _selection.ActiveHistoryMode = mode;
        NotifyState();
    }

    public void SetFilters(string? search, string? category, DateOnly? from, DateOnly? to)
    {
        _searchText = search?.Trim() ?? "";
        _category = string.IsNullOrWhiteSpace(category) ? "All benefit areas" : category;
        _from = from;
        _to = to;
        Rebuild();
        OnChanged(nameof(SearchText));
        OnChanged(nameof(SelectedCategory));
        OnChanged(nameof(DateFrom));
        OnChanged(nameof(DateTo));
        OnChanged(nameof(HasActiveFilters));
    }

    public void ClearFilters() => SetFilters("", "All benefit areas", null, null);

    public void SelectEvolutionStep(string id) => SelectHistoryEvent(id);

    public bool SelectHistoryEvent(string id, long? renewalId = null, long? carrierId = null, int? year = null)
    {
        var historyEvent = _events.FirstOrDefault(item => item.EventId == id);
        if (historyEvent is null)
        {
            return false;
        }

        _selection.SelectedHistoryEventId = id;
        _selection.SelectedEvolutionStepId = id;
        _selection.SelectedDocumentIds = historyEvent.Evidence.Where(item => item.DocumentId.HasValue).Select(item => item.DocumentId!.Value).ToList();
        _selection.SelectedRenewalId = renewalId;
        _selection.SelectedCarrierId = carrierId;
        _selection.SelectedYear = year ?? DateReference(historyEvent).Year;
        _selection.IsDetailsDrawerOpen = true;
        Rebuild();
        return true;
    }

    public void OpenDrawer() => _selection.IsDetailsDrawerOpen = true;
    public void CloseDrawer() => _selection.IsDetailsDrawerOpen = false;

    public void ClearSelection()
    {
        _selection.ClearSelection();
        Rebuild();
        NotifyState();
    }

    private void Rebuild()
    {
        FilteredEvents = _events.Where(Matches).ToList();
        OnChanged(nameof(FilteredEvents));

        EvolutionActivityEvents = FilteredEvents
            .Where(item => !item.IsExcluded && DateReference(item).Precision != HistoryDatePrecision.Unknown)
            .Where(item => item.Category.Contains("Plan", StringComparison.OrdinalIgnoreCase) ||
                           item.Category.Contains("Carrier", StringComparison.OrdinalIgnoreCase) ||
                           item.Category.Contains("Market", StringComparison.OrdinalIgnoreCase) ||
                           item.Category.Contains("Administration", StringComparison.OrdinalIgnoreCase))
            .OrderByDescending(item => DateReference(item).SortYear)
            .ThenByDescending(item => DateReference(item).SortMonth)
            .ThenByDescending(item => DateReference(item).SortDay)
            .Take(6)
            .Select(item => new HistoryEvolutionActivityItem
            {
                EventId = item.EventId,
                DateReference = DateReference(item),
                Category = item.Category,
                Title = item.Title,
                Summary = item.Summary,
                AccentColor = ActivityAccentColor(item.Category),
                AccentBackground = ActivityAccentBackground(item.Category),
                IsSelected = item.EventId == SelectedEvolutionStepId
            })
            .ToList();
        OnChanged(nameof(EvolutionActivityEvents));
        var tracks = _evolution.BuildTracks(FilteredEvents, "All benefit areas", _from, _to);
        FilteredEvolutionTracks = tracks.Select(track => new HistoryEvolutionTrack
        {
            Subject = track.Subject,
            Steps = track.Steps.Select(step => new HistoryEvolutionStep
            {
                EventId = step.EventId,
                EffectiveDate = step.EffectiveDate,
                PreviousValue = step.PreviousValue,
                NewValue = step.NewValue,
                Unit = step.Unit,
                Direction = step.Direction,
                RelatedRenewal = step.RelatedRenewal,
                EvidenceCount = step.EvidenceCount,
                IsSelected = step.EventId == SelectedEvolutionStepId
            }).ToList()
        }).ToList();
        OnChanged(nameof(FilteredEvolutionTracks));
        OnChanged(nameof(HasActiveFilters));
    }

    private bool Matches(ClientHistoryEventReadModel historyEvent) =>
        (_category.Equals("All benefit areas", StringComparison.OrdinalIgnoreCase) ||
         historyEvent.Category.Equals(_category, StringComparison.OrdinalIgnoreCase)) &&
        (string.IsNullOrWhiteSpace(_searchText) ||
         $"{historyEvent.Category} {historyEvent.Title} {historyEvent.Summary} {historyEvent.SourceModule} {string.Join(' ', historyEvent.ChangeSet.Changes.Select(change => $"{change.DisplayLabel} {change.PreviousValue} {change.NewValue}"))}"
             .Contains(_searchText, StringComparison.OrdinalIgnoreCase)) &&
        MatchesDateContext(historyEvent);

    private bool MatchesDateContext(ClientHistoryEventReadModel historyEvent)
    {
        var date = DateReference(historyEvent);
        if (_from.HasValue && !date.IsOnOrAfter(_from.Value))
        {
            return false;
        }
        if (_to.HasValue && !date.IsOnOrBefore(_to.Value))
        {
            return false;
        }
        return true;
    }

    private static string ActivityAccentColor(string category) =>
        category.Contains("Carrier", StringComparison.OrdinalIgnoreCase) ? "#0B5DE8" :
        category.Contains("Market", StringComparison.OrdinalIgnoreCase) || category.Contains("Financial", StringComparison.OrdinalIgnoreCase) ? "#F47A13" :
        category.Contains("Administration", StringComparison.OrdinalIgnoreCase) ? "#6D28D9" : "#18A55B";

    private static string ActivityAccentBackground(string category) =>
        category.Contains("Carrier", StringComparison.OrdinalIgnoreCase) ? "#EEF4FF" :
        category.Contains("Market", StringComparison.OrdinalIgnoreCase) || category.Contains("Financial", StringComparison.OrdinalIgnoreCase) ? "#FFF1E6" :
        category.Contains("Administration", StringComparison.OrdinalIgnoreCase) ? "#F3EDFF" : "#EAF8F0";
    private static HistoryDateReference DateReference(ClientHistoryEventReadModel historyEvent) =>
        historyEvent.EffectiveDateReference.Precision != HistoryDatePrecision.Unknown
            ? historyEvent.EffectiveDateReference
            : historyEvent.EffectiveDate.HasValue
                ? HistoryDateReference.FromExact(historyEvent.EffectiveDate.Value)
                : new HistoryDateReference();

    private void NotifyState()
    {
        OnChanged(nameof(ActiveHistoryMode));
        OnChanged(nameof(IsTimelineVisible));
        OnChanged(nameof(IsEvolutionVisible));
        OnChanged(nameof(IsAuditVisible));
        OnChanged(nameof(IsDetailsDrawerOpen));
        OnChanged(nameof(SelectedHistoryEventId));
        OnChanged(nameof(SelectedEvolutionStepId));
        OnChanged(nameof(HasActiveFilters));
    }

    private void OnChanged(string? propertyName) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
}
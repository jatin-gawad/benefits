using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Models.History;
using RenewalReviewBuilder.App.Services;
using RenewalReviewBuilder.App.Services.ClientHistory;

namespace RenewalReviewBuilder.App.Views;

public partial class ClientHistoryWorkspaceView : UserControl
{
    private readonly DatabaseService _database;
    private readonly ClientHistoryImportAction _importAction = new();
    private IReadOnlyList<ClientHistoryEventReadModel> _events = [];
    private bool _isSynchronizingTimelineSelection;

    public HistorySelectionState Selection { get; } = new();
    public ClientHistoryWorkspaceController Workspace { get; }
    public HistoryPresentationController Presentation => Workspace.Presentation;
    public ObservableCollection<ClientHistoryTimelineRow> TimelineRows { get; } = [];
    public ObservableCollection<HistoryEvolutionTrack> EvolutionTracks { get; } = [];
    public ObservableCollection<HistoryRelationshipMilestone> RelationshipMilestones => Workspace.RelationshipMilestones;
    public ObservableCollection<HistoryRenewalEvolutionPoint> RenewalPoints => Workspace.RenewalPoints;
    public bool IsRelationshipTimelineLoading => Workspace.IsRelationshipTimelineLoading;
    public bool IsRenewalEvolutionLoading => Workspace.IsRenewalEvolutionLoading;
    public string RelationshipTimelineEmptyState => Workspace.RelationshipTimelineEmptyState;
    public string RenewalEvolutionEmptyState => Workspace.RenewalEvolutionEmptyState;
    public string RelationshipTimelineError => Workspace.RelationshipTimelineError;
    public string RenewalEvolutionError => Workspace.RenewalEvolutionError;

    public ClientHistoryWorkspaceView()
    {
        _database = new DatabaseService();
        Workspace = new ClientHistoryWorkspaceController(
            Selection,
            ClientHistoryProjectionBuilder.CreateDefault(_database),
            async (clientId, _) => (IReadOnlyList<RenewalVersionRecord>)await _database.GetRenewalVersionsAsync(clientId));

        InitializeComponent();
        DataContext = this;
        PreviewKeyDown += (_, eventArgs) =>
        {
            if (eventArgs.Key == System.Windows.Input.Key.Escape)
            {
                CloseDrawer();
            }
        };
        Selection.PropertyChanged += (_, eventArgs) =>
        {
            if (eventArgs.PropertyName == nameof(HistorySelectionState.IsDetailsDrawerOpen))
            {
                SetDrawer();
            }
            if (eventArgs.PropertyName == nameof(HistorySelectionState.ActiveHistoryMode))
            {
                UpdateMode();
            }
            if (eventArgs.PropertyName is nameof(HistorySelectionState.SelectedHistoryEventId) or nameof(HistorySelectionState.SelectedRenewalId) or nameof(HistorySelectionState.SelectedCarrierId) or nameof(HistorySelectionState.SelectedYear))
            {
                SyncVisualSelection();
                if (Selection.IsDetailsDrawerOpen)
                {
                    RenderDrawer();
                }
            }
        };
        Presentation.PropertyChanged += (_, eventArgs) =>
        {
            if (eventArgs.PropertyName == nameof(HistoryPresentationController.FilteredEvents))
            {
                RefreshTimelineRows();
            }
            if (eventArgs.PropertyName is nameof(HistoryPresentationController.FilteredEvolutionTracks) or nameof(HistoryPresentationController.EvolutionActivityEvents))
            {
                RefreshEvolution();
            }
        };
        UpdateCategoryButtons();
    }

    public async Task LoadAsync(long clientId)
    {
        await Workspace.LoadAsync(clientId);
        ApplyLoadedProjection();
    }

    private async void AddHistory_Click(object sender, RoutedEventArgs eventArgs)
    {
        if (Workspace.DatabaseClientId is not > 0)
        {
            MessageBox.Show(
                "Save or select a persisted client before importing Client History.",
                "Add Client History",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        var owner = Window.GetWindow(this);
        if (owner is null)
        {
            MessageBox.Show(
                "The Client History import window could not be opened from this workspace.",
                "Add Client History",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        AddHistoryButton.IsEnabled = false;
        try
        {
            await _importAction.RunAsync(
                Workspace.DatabaseClientId.Value,
                databaseClientId =>
                {
                    var importWindow = new ClientHistoryImportPreviewWindow(databaseClientId)
                    {
                        Owner = owner
                    };
                    importWindow.ShowDialog();
                },
                async cancellationToken =>
                {
                    await Workspace.RefreshAsync(cancellationToken);
                    ApplyLoadedProjection();
                });
        }
        catch (OperationCanceledException)
        {
            // A newer client/projection load superseded this refresh.
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"Client History could not be refreshed after import: {ex.Message}",
                "Add Client History",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        finally
        {
            AddHistoryButton.IsEnabled = Workspace.DatabaseClientId is > 0;
        }
    }

    private async void AddHistoricalRenewal_Click(object sender, RoutedEventArgs eventArgs)
    {
        if (Workspace.DatabaseClientId is not > 0)
        {
            MessageBox.Show(
                "Save or select a persisted client before importing a historical renewal.",
                "Add Historical Renewal",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        var owner = Window.GetWindow(this);
        if (owner is null)
        {
            MessageBox.Show(
                "The historical renewal import window could not be opened from this workspace.",
                "Add Historical Renewal",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        AddHistoricalRenewalButton.IsEnabled = false;
        try
        {
            var importWindow = new HistoricalRenewalImportWindow(Workspace.DatabaseClientId.Value)
            {
                Owner = owner
            };
            importWindow.ShowDialog();
            if (importWindow.ImportSucceeded)
            {
                await Workspace.RefreshAsync();
                ApplyLoadedProjection();
            }
        }
        catch (OperationCanceledException)
        {
            // A newer client/projection load superseded this refresh.
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"Client History could not be refreshed after the historical renewal import: {ex.Message}",
                "Add Historical Renewal",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        finally
        {
            AddHistoricalRenewalButton.IsEnabled = Workspace.DatabaseClientId is > 0;
        }
    }

    private void ApplyLoadedProjection()
    {
        _events = Workspace.Events;
        RefreshTimelineRows();
        RefreshEvolution();
        if (Selection.IsDetailsDrawerOpen)
        {
            RenderDrawer();
        }
    }

    private async void ApplyReplay_Click(object sender, RoutedEventArgs eventArgs)
    {
        if (ToDateOnly(ReplayDatePicker.SelectedDate) is not DateOnly date) return;
        await Workspace.ActivateReplayAsync(date);
        _events = Workspace.Events; RefreshTimelineRows(); RefreshEvolution();
    }

    private async void GoToCurrent_Click(object sender, RoutedEventArgs eventArgs)
    {
        await Workspace.GoToCurrentAsync();
        _events = Workspace.Events; RefreshTimelineRows(); RefreshEvolution();
    }
    private void RelationshipMilestone_Click(object sender, RoutedEventArgs eventArgs)
    {
        if ((sender as FrameworkElement)?.DataContext is HistoryRelationshipMilestone milestone &&
            Workspace.SelectRelationshipMilestone(milestone))
        {
            RenderDrawer();
        }
    }

    private void RenewalPoint_Click(object sender, RoutedEventArgs eventArgs)
    {
        if ((sender as FrameworkElement)?.DataContext is HistoryRenewalEvolutionPoint point &&
            Workspace.SelectRenewalPoint(point))
        {
            RenderDrawer();
        }
    }

    private void HistorySearchTextBox_TextChanged(object sender, TextChangedEventArgs eventArgs)
    {
        if (HistorySearchWatermark is not null)
        {
            HistorySearchWatermark.Visibility = string.IsNullOrEmpty(HistorySearchTextBox.Text) ? Visibility.Visible : Visibility.Collapsed;
        }
        ApplyFilters();
    }

    private void HistoryDateFilter_Changed(object? sender, SelectionChangedEventArgs eventArgs) => ApplyFilters();

    private void Category_Click(object sender, RoutedEventArgs eventArgs)
    {
        if ((sender as Button)?.Tag is string category)
        {
            Workspace.SetFilters(Presentation.SearchText, category, Presentation.DateFrom, Presentation.DateTo);
            UpdateCategoryButtons();
        }
    }

    private void ClearFilters_Click(object sender, RoutedEventArgs eventArgs)
    {
        HistorySearchTextBox.Text = "";
        HistoryDateFromPicker.SelectedDate = null;
        HistoryDateToPicker.SelectedDate = null;
        Workspace.ClearFilters();
        UpdateCategoryButtons();
    }

    private void UpdateCategoryButtons()
    {
        if (CategoryPanel is null)
        {
            return;
        }

        foreach (var button in CategoryPanel.Children.OfType<Button>())
        {
            var selected = button.Tag is string category && category.Equals(Presentation.SelectedCategory, StringComparison.OrdinalIgnoreCase);
            button.Background = new SolidColorBrush(selected ? Color.FromRgb(238, 244, 255) : Colors.White);
            button.BorderBrush = new SolidColorBrush(selected ? Color.FromRgb(11, 93, 232) : Color.FromRgb(217, 226, 239));
            button.Foreground = new SolidColorBrush(selected ? Color.FromRgb(11, 93, 232) : Color.FromRgb(7, 26, 69));
        }
    }
    private void ApplyFilters()
    {
        if (HistorySearchTextBox is null)
        {
            return;
        }

        Workspace.SetFilters(
            HistorySearchTextBox.Text,
            Presentation.SelectedCategory,
            ToDateOnly(HistoryDateFromPicker?.SelectedDate),
            ToDateOnly(HistoryDateToPicker?.SelectedDate));
    }

    private void RefreshTimelineRows()
    {
        var rows = new ClientHistoryTimelineService().BuildRows(
            Presentation.FilteredEvents,
            new ClientHistoryFilter { Category = "All Events" });
        var wasSynchronizing = _isSynchronizingTimelineSelection;
        _isSynchronizingTimelineSelection = true;
        try
        {
            TimelineRows.Clear();
            foreach (var row in rows)
            {
                TimelineRows.Add(row);
            }
            SyncVisualSelection();
        }
        finally
        {
            _isSynchronizingTimelineSelection = wasSynchronizing;
        }
    }

    private void SyncVisualSelection()
    {
        var selectedRow = TimelineRows.FirstOrDefault(row => row.EventId == Selection.SelectedHistoryEventId);
        if (ReferenceEquals(TimelineGrid.SelectedItem, selectedRow))
        {
            return;
        }

        var wasSynchronizing = _isSynchronizingTimelineSelection;
        _isSynchronizingTimelineSelection = true;
        try
        {
            TimelineGrid.SelectedItem = selectedRow;
            if (selectedRow is not null)
            {
                TimelineGrid.ScrollIntoView(selectedRow);
            }
        }
        finally
        {
            _isSynchronizingTimelineSelection = wasSynchronizing;
        }
    }
    private static DateOnly? ToDateOnly(DateTime? value) => value.HasValue ? DateOnly.FromDateTime(value.Value) : null;

    private void Timeline_Click(object sender, RoutedEventArgs eventArgs)
    {
        Presentation.SelectMode(HistoryMode.Timeline);
        UpdateMode();
    }

    private void Evolution_Click(object sender, RoutedEventArgs eventArgs)
    {
        Presentation.SelectMode(HistoryMode.Evolution);
        UpdateMode();
    }

    private void UpdateMode()
    {
        TimelineGrid.Visibility = Visibility.Visible;
        EvolutionPanel.Visibility = Visibility.Visible;
        RefreshEvolution();
        SyncVisualSelection();
    }

    private void RefreshEvolution()
    {
        EvolutionTracks.Clear();
        foreach (var track in Presentation.FilteredEvolutionTracks)
        {
            EvolutionTracks.Add(track);
        }

        var hasStructuredTracks = EvolutionTracks.Count > 0;
        var hasCanonicalActivity = Presentation.EvolutionActivityEvents.Count > 0;
        StructuredEvolutionList.Visibility = hasStructuredTracks ? Visibility.Visible : Visibility.Collapsed;
        EvolutionActivityList.Visibility = !hasStructuredTracks && hasCanonicalActivity ? Visibility.Visible : Visibility.Collapsed;

        EvolutionStatusText.Text = hasStructuredTracks
            ? "Changes are grouped by normalized canonical subject."
            : hasCanonicalActivity
                ? "Structured before/after values are unavailable; showing canonical history activity."
                : _events.Count == 0
                    ? "No canonical history is available for this client."
                    : "No structured canonical changes are available for Evolution View.";
    }
    private void EvolutionStep_Click(object sender, RoutedEventArgs eventArgs)
    {
        if ((sender as Button)?.Tag is string eventId && Workspace.SelectHistoryEvent(eventId))
        {
            RenderDrawer();
        }
    }

    private void TimelineGrid_SelectionChanged(object sender, SelectionChangedEventArgs eventArgs)
    {
        if (_isSynchronizingTimelineSelection)
        {
            return;
        }

        if ((sender as ListBox)?.SelectedItem is ClientHistoryTimelineRow row)
        {
            OpenTimelineDetails(row, focusDrawer: false);
        }
    }

    private void TimelineDetails_Click(object sender, RoutedEventArgs eventArgs)
    {
        if ((sender as FrameworkElement)?.DataContext is not ClientHistoryTimelineRow row)
        {
            return;
        }

        eventArgs.Handled = true;
        OpenTimelineDetails(row, focusDrawer: true);
    }

    private bool OpenTimelineDetails(ClientHistoryTimelineRow row, bool focusDrawer)
    {
        if (string.IsNullOrWhiteSpace(row.EventId) || !Workspace.OpenHistoryEventDetails(row.EventId))
        {
            return false;
        }

        // Apply the visual state directly as well as through notifications so re-clicking an
        // already-selected event reliably reopens and repaints the one shared drawer.
        SyncVisualSelection();
        SetDrawer();
        RenderDrawer();
        if (focusDrawer)
        {
            DetailsDrawer.BringIntoView();
            DrawerTitleText.BringIntoView();
            DrawerCloseButton.Focus();
        }

        return true;
    }
    private void Previous_Click(object sender, RoutedEventArgs eventArgs) => Move(-1);
    private void Next_Click(object sender, RoutedEventArgs eventArgs) => Move(1);

    private void Move(int delta)
    {
        var index = TimelineRows.ToList().FindIndex(row => row.EventId == Selection.SelectedHistoryEventId);
        var next = index + delta;
        if (next < 0 || next >= TimelineRows.Count)
        {
            return;
        }

        if (Workspace.SelectHistoryEvent(TimelineRows[next].EventId))
        {
            RenderDrawer();
        }
    }

    private void RenderDrawer()
    {
        var item = _events.FirstOrDefault(historyEvent => historyEvent.EventId == Selection.SelectedHistoryEventId);
        if (item is null)
        {
            return;
        }

        var position = TimelineRows.ToList().FindIndex(row => row.EventId == item.EventId) + 1;
        PositionText.Text = $"{position} of {TimelineRows.Count}";
        DrawerTypeText.Text = item.EventType;
        DrawerTitleText.Text = item.Title;
        DrawerTypeText.Text = Workspace.IsReplayActive ? $"{item.EventType} | {Workspace.ReplayDisplayLabel}" : item.EventType;
        DrawerDateText.Text = item.EffectiveDateReference.Precision != HistoryDatePrecision.Unknown
            ? item.EffectiveDateReference.DisplayText
            : item.EffectiveDate?.ToString("MMM d, yyyy") ?? "Date unavailable";
        var timelineRow = TimelineRows.FirstOrDefault(row => row.EventId == item.EventId);
        if (timelineRow is not null)
        {
            DrawerIconText.Text = timelineRow.IconGlyph;
            DrawerIconText.Foreground = (Brush?)new BrushConverter().ConvertFromString(timelineRow.AccentColor) ?? Brushes.RoyalBlue;
            DrawerIconBackground.Background = (Brush?)new BrushConverter().ConvertFromString(timelineRow.AccentBackground) ?? Brushes.AliceBlue;
        }

        var impact = string.Join(" ", new[] { item.BusinessImpact, item.FinancialImpact }
            .Where(value => !string.IsNullOrWhiteSpace(value)));
        ImpactText.Text = impact;
        BusinessImpactSection.Visibility = string.IsNullOrWhiteSpace(impact) ? Visibility.Collapsed : Visibility.Visible;

        var details = ClientHistoryDrawerContentFormatter.Details(item);
        DetailsText.Text = string.Join(Environment.NewLine, details);
        BeforeAfterSection.Visibility = item.ChangeSet.Changes.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
        ChangesList.ItemsSource = item.ChangeSet.Changes.Select(change => $"{change.DisplayLabel}: {change.PreviousValue} → {change.NewValue}");
        EvidenceSection.Visibility = item.Evidence.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
        EvidenceList.ItemsSource = ClientHistoryDrawerContentFormatter.Evidence(item);
        RelatedList.ItemsSource = item.RelatedRecords;
        RelatedRecordsSection.Visibility = item.RelatedRecords.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
        ReviewActions.Visibility = new HistoryPermissionService().CanReview(item) ? Visibility.Visible : Visibility.Collapsed;
    }

    private void OpenSource_Click(object sender, RoutedEventArgs eventArgs)
    {
        var item = _events.FirstOrDefault(historyEvent => historyEvent.EventId == Selection.SelectedHistoryEventId);
        if (item?.Evidence.FirstOrDefault()?.DocumentId is null)
        {
            return;
        }

        MessageBox.Show(
            $"Source document {item.Evidence.First().DocumentId} is available through the existing Documents workspace.",
            "History source");
    }

    private void CloseDrawer_Click(object sender, RoutedEventArgs eventArgs) => CloseDrawer();

    private void CloseDrawer() => Presentation.CloseDrawer();

    private void SetDrawer()
    {
        var open = Selection.IsDetailsDrawerOpen;
        DetailsDrawer.Visibility = open ? Visibility.Visible : Visibility.Collapsed;
        DrawerColumn.Width = open ? new GridLength(255) : new GridLength(0);
    }
}

using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;

namespace RenewalReviewBuilder.App.Views;

/// <summary>Reviews one client-history workbook and imports explicitly approved rows.</summary>
public sealed class ClientHistoryImportPreviewWindow : Window
{
    private readonly ClientHistoryWorkbookPreviewParser _parser = new();
    private readonly DatabaseService _databaseService = new();
    private readonly long _clientId;
    private readonly TextBlock _summary = new() { Text = "Choose one Excel workbook to preview, review, and import approved history rows.", Margin = new Thickness(0, 0, 0, 10) };
    private readonly TabControl _tabs = new();
    private ClientHistoryWorkbookPreview? _preview;

    public ClientHistoryImportPreviewWindow(long clientId)
    {
        _clientId = clientId;
        Title = "Import Client History";
        Width = 1250;
        Height = 720;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        var browse = new Button { Content = "Preview Workbook...", Padding = new Thickness(12, 6, 12, 6), Margin = new Thickness(0, 0, 10, 10) };
        browse.Click += PreviewWorkbook_Click;
        var approve = new Button { Content = "Approve Selected", Padding = new Thickness(12, 6, 12, 6), Margin = new Thickness(0, 0, 10, 10) };
        approve.Click += (_, _) => SetSelectedStatus(ClientHistoryPreviewStatuses.Approved);
        var reject = new Button { Content = "Reject Selected", Padding = new Thickness(12, 6, 12, 6), Margin = new Thickness(0, 0, 10, 10) };
        reject.Click += (_, _) => SetSelectedStatus(ClientHistoryPreviewStatuses.Rejected);
        var selectAll = new Button { Content = "Approve Parsed", Padding = new Thickness(12, 6, 12, 6), Margin = new Thickness(0, 0, 10, 10) };
        selectAll.Click += (_, _) => { if (_preview is not null) foreach (var row in _preview.Records.Where(row => row.PreviewStatus == ClientHistoryPreviewStatuses.Parsed)) row.PreviewStatus = ClientHistoryPreviewStatuses.Approved; RefreshTabs(); };
        var import = new Button { Content = "Import Approved", Padding = new Thickness(12, 6, 12, 6), Margin = new Thickness(0, 0, 10, 10) };
        import.Click += ImportApproved_Click;
        var close = new Button { Content = "Close", Padding = new Thickness(12, 6, 12, 6), Margin = new Thickness(0, 0, 0, 10) };
        close.Click += (_, _) => Close();
        var toolbar = new StackPanel { Orientation = Orientation.Horizontal };
        toolbar.Children.Add(browse); toolbar.Children.Add(approve); toolbar.Children.Add(reject); toolbar.Children.Add(selectAll); toolbar.Children.Add(import); toolbar.Children.Add(close);
        var panel = new DockPanel { Margin = new Thickness(18) };
        DockPanel.SetDock(toolbar, Dock.Top); DockPanel.SetDock(_summary, Dock.Top);
        panel.Children.Add(toolbar); panel.Children.Add(_summary); panel.Children.Add(_tabs);
        Content = panel;
    }

    private void PreviewWorkbook_Click(object? sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog { Filter = "Excel workbooks (*.xlsx)|*.xlsx", CheckFileExists = true };
        if (dialog.ShowDialog(this) != true) return;
        _preview = _parser.Parse(dialog.FileName);
        RefreshTabs();
    }

    private async void ImportApproved_Click(object? sender, RoutedEventArgs e)
    {
        if (_preview is null) return;
        var result = await _databaseService.ImportApprovedClientHistoryEventsAsync(_clientId, _preview.Records);
        RefreshTabs();
        _summary.Text = $"Approved: {result.Approved}; Imported: {result.Imported}; Skipped duplicates: {result.SkippedDuplicates}; Rejected: {result.Rejected}; Failed: {result.Failed}. {string.Join(" ", result.Warnings)}";
    }

    private void SetSelectedStatus(string status)
    {
        if (_tabs.SelectedContent is not DataGrid grid || grid.SelectedItems.Count == 0) return;
        foreach (var row in grid.SelectedItems.OfType<ClientHistoryPreviewEvent>()) row.PreviewStatus = status;
        RefreshTabs();
    }

    private void RefreshTabs()
    {
        if (_preview is null) return;
        var preview = _preview;
        _summary.Text = $"Preview: {preview.Records.Count} record(s). Approve records before importing. {string.Join("; ", preview.CountsByEventType.Select(pair => $"{pair.Key}: {pair.Value}"))}";
        _tabs.Items.Clear();
        AddTab("Parsed", preview.Records.Where(row => row.PreviewStatus == ClientHistoryPreviewStatuses.Parsed));
        AddTab("Needs Review", preview.Records.Where(row => row.PreviewStatus == ClientHistoryPreviewStatuses.NeedsReview));
        AddTab("Approved", preview.Records.Where(row => row.PreviewStatus == ClientHistoryPreviewStatuses.Approved));
        AddTab("Rejected", preview.Records.Where(row => row.PreviewStatus == ClientHistoryPreviewStatuses.Rejected));
        AddTab("Imported", preview.Records.Where(row => row.PreviewStatus == ClientHistoryPreviewStatuses.Imported));
        AddTab("Failed", preview.Records.Where(row => row.PreviewStatus == ClientHistoryPreviewStatuses.Failed));
    }

    private void AddTab(string label, IEnumerable<ClientHistoryPreviewEvent> rows)
    {
        var items = rows.ToList();
        var grid = new DataGrid { ItemsSource = items, AutoGenerateColumns = false, IsReadOnly = false, CanUserAddRows = false, SelectionMode = DataGridSelectionMode.Extended };
        grid.Columns.Add(new DataGridTextColumn { Header = "Date", Binding = new System.Windows.Data.Binding(nameof(ClientHistoryPreviewEvent.EffectiveDate)), Width = 100 });
        grid.Columns.Add(new DataGridTextColumn { Header = "Type", Binding = new System.Windows.Data.Binding(nameof(ClientHistoryPreviewEvent.EventType)), Width = 150 });
        grid.Columns.Add(new DataGridTextColumn { Header = "Title", Binding = new System.Windows.Data.Binding(nameof(ClientHistoryPreviewEvent.Title)), Width = 230 });
        grid.Columns.Add(new DataGridTextColumn { Header = "Description", Binding = new System.Windows.Data.Binding(nameof(ClientHistoryPreviewEvent.Description)), Width = new DataGridLength(1, DataGridLengthUnitType.Star) });
        grid.Columns.Add(new DataGridTextColumn { Header = "Sheet", Binding = new System.Windows.Data.Binding(nameof(ClientHistoryPreviewEvent.SourceSheet)), Width = 120 });
        grid.Columns.Add(new DataGridTextColumn { Header = "Row", Binding = new System.Windows.Data.Binding(nameof(ClientHistoryPreviewEvent.SourceRow)), Width = 55 });
        grid.Columns.Add(new DataGridTextColumn { Header = "Review notes", Binding = new System.Windows.Data.Binding(nameof(ClientHistoryPreviewEvent.ValidationMessages)), Width = 220 });
        grid.Columns.Add(new DataGridTextColumn { Header = "Duplicate", Binding = new System.Windows.Data.Binding(nameof(ClientHistoryPreviewEvent.DuplicateStatus)), Width = 130, IsReadOnly = true });
        _tabs.Items.Add(new TabItem { Header = $"{label} ({items.Count})", Content = grid });
    }
}

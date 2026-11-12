using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using System.Windows.Threading;
using Microsoft.Win32;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;

namespace RenewalReviewBuilder.App.Views;

public partial class ClassComparisonWorkspaceView : UserControl
{
    private readonly ClassComparisonExportService _exportService = new();
    private readonly PlanDesignWorkspaceDataService _dataService = new();
    private readonly SemaphoreSlim _preferenceSaveGate = new(1, 1);
    private readonly DispatcherTimer _searchDebounceTimer;
    private bool _isUpdatingControls;

    public event EventHandler<ClassComparisonNavigationEventArgs>? NavigateTabRequested;
    public event EventHandler? RetryRequested;

    public ClassComparisonWorkspaceView()
    {
        InitializeComponent();
        _searchDebounceTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(250) };
        _searchDebounceTimer.Tick += SearchDebounceTimer_Tick;
        DataContextChanged += ClassComparisonWorkspaceView_DataContextChanged;
    }

    private ClassComparisonWorkspaceViewModel? Model => DataContext as ClassComparisonWorkspaceViewModel;

    private void ClassComparisonWorkspaceView_DataContextChanged(object sender, DependencyPropertyChangedEventArgs e)
    {
        RefreshAfterMutation(Model);
        UpdateResponsiveLayout(ActualWidth);
    }

    private void BaselineClassComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_isUpdatingControls || Model is not { } model || BaselineClassComboBox.SelectedValue is not string baseline)
        {
            return;
        }

        var comparison = model.ComparisonClassName;
        if (baseline.Equals(comparison, StringComparison.OrdinalIgnoreCase))
        {
            comparison = model.Classes.FirstOrDefault(item => !item.ClassName.Equals(baseline, StringComparison.OrdinalIgnoreCase))?.ClassName ?? "";
        }
        ApplyPairSelection(model, baseline, comparison);
    }

    private void ComparisonClassComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_isUpdatingControls || Model is not { } model || ComparisonClassComboBox.SelectedItem is not ClassComparisonClassCard selected)
        {
            return;
        }
        ApplyPairSelection(model, model.BaselineClassName, selected.ClassName);
    }

    private void ApplyPairSelection(ClassComparisonWorkspaceViewModel model, string baseline, string comparison)
    {
        if (!ClassComparisonEngine.TrySetPairClasses(model, baseline, comparison, out var message))
        {
            model.StatusText = message;
            RefreshAfterMutation(model);
            return;
        }
        model.StatusText = $"Comparing {model.SelectedComparisonClass?.DisplayName ?? comparison} against {model.BaselineClass?.DisplayName ?? baseline}.";
        RefreshAfterMutation(model);
        QueuePreferenceSave(model);
    }

    private void SwapClasses_Click(object sender, RoutedEventArgs e)
    {
        if (Model is not { } model || !ClassComparisonEngine.SwapPairClasses(model)) return;
        RefreshAfterMutation(model);
        QueuePreferenceSave(model);
    }

    private void AllBenefits_Click(object sender, RoutedEventArgs e)
    {
        if (Model is not { } model || !model.ShowDifferencesOnly) return;
        ClassComparisonEngine.SetShowDifferencesOnly(model, false);
        RefreshAfterMutation(model);
        QueuePreferenceSave(model);
    }

    private void KeyDifferences_Click(object sender, RoutedEventArgs e)
    {
        if (Model is not { } model || model.ShowDifferencesOnly) return;
        ClassComparisonEngine.SetShowDifferencesOnly(model, true);
        RefreshAfterMutation(model);
        QueuePreferenceSave(model);
    }

    private void ShowDifferencesOnly_Click(object sender, RoutedEventArgs e)
    {
        if (_isUpdatingControls || Model is not { } model || sender is not CheckBox checkBox) return;
        ClassComparisonEngine.SetShowDifferencesOnly(model, checkBox.IsChecked == true);
        RefreshAfterMutation(model);
        QueuePreferenceSave(model);
    }

    private void HighlightDifferences_Click(object sender, RoutedEventArgs e)
    {
        if (_isUpdatingControls || Model is not { } model || sender is not CheckBox checkBox) return;
        ClassComparisonEngine.SetHighlightDifferences(model, checkBox.IsChecked == true);
        RefreshAfterMutation(model);
        QueuePreferenceSave(model);
    }

    private void SearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_isUpdatingControls) return;
        _searchDebounceTimer.Stop();
        _searchDebounceTimer.Start();
    }

    private void SearchDebounceTimer_Tick(object? sender, EventArgs e)
    {
        _searchDebounceTimer.Stop();
        if (Model is not { } model) return;
        ClassComparisonEngine.SetSearchText(model, SearchTextBox.Text);
        ComparisonDataGrid.Items.Refresh();
    }

    private void RefreshAfterMutation(ClassComparisonWorkspaceViewModel? model)
    {
        if (model is null) return;
        _isUpdatingControls = true;
        try
        {
            BaselineClassComboBox.Items.Refresh();
            ComparisonClassComboBox.Items.Refresh();
            BaselineClassComboBox.SelectedValue = model.BaselineClassName;
            ComparisonClassComboBox.SelectedItem = model.SelectedComparisonClass;
            ComparisonDataGrid.Items.Refresh();
            BaselineValueColumn.Header = BuildClassHeader(model.BaselineClass?.DisplayName ?? model.BaselineClassName, true);
            ComparisonValueColumn.Header = BuildClassHeader(model.SelectedComparisonClass?.DisplayName ?? model.ComparisonClassName, false);
        }
        finally
        {
            _isUpdatingControls = false;
        }
    }

    private static FrameworkElement BuildClassHeader(string className, bool baseline)
    {
        var panel = new StackPanel { Orientation = Orientation.Horizontal };
        panel.Children.Add(new TextBlock
        {
            Text = string.IsNullOrWhiteSpace(className) ? "Class" : className,
            FontSize = 13,
            FontWeight = FontWeights.SemiBold,
            Foreground = new SolidColorBrush(Color.FromRgb(11, 35, 65)),
            TextTrimming = TextTrimming.CharacterEllipsis,
            MaxWidth = 260,
            VerticalAlignment = VerticalAlignment.Center
        });
        if (baseline)
        {
            panel.Children.Add(new Border
            {
                Background = new SolidColorBrush(Color.FromRgb(240, 253, 244)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(134, 239, 172)),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(4),
                Padding = new Thickness(7, 2, 7, 2),
                Margin = new Thickness(9, 0, 0, 0),
                Child = new TextBlock { Text = "Baseline", Foreground = new SolidColorBrush(Color.FromRgb(21, 128, 61)), FontSize = 11 }
            });
        }
        return panel;
    }

    private async void ExportComparison_Click(object sender, RoutedEventArgs e)
    {
        if (Model is not { HasEnoughClasses: true, HasComparableFeatures: true } model)
        {
            MessageBox.Show("No two-class comparison is loaded to export.", "Class Comparison", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }
        AppPaths.EnsureCreated();
        var reportsFolder = Path.Combine(AppPaths.AppDataFolder, "Reports");
        Directory.CreateDirectory(reportsFolder);
        var dialog = new SaveFileDialog
        {
            Title = "Export Class Comparison",
            Filter = "Excel Workbook (*.xlsx)|*.xlsx",
            DefaultExt = ".xlsx",
            AddExtension = true,
            InitialDirectory = reportsFolder,
            FileName = $"ClassComparison_{SafeFileNamePart(model.ClientName)}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
        };
        if (dialog.ShowDialog() != true) return;
        try
        {
            var result = _exportService.ExportToExcel(model, dialog.FileName);
            model.StatusText = $"Exported {result.ComparisonRowCount} visible comparison rows.";
            MessageBox.Show($"The class comparison was exported to:\n\n{result.OutputPath}", "Class Comparison Export", MessageBoxButton.OK, MessageBoxImage.Information);
            Process.Start(new ProcessStartInfo(Path.GetDirectoryName(result.OutputPath)!) { UseShellExecute = true });
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            model.StatusText = "The class comparison export failed.";
            MessageBox.Show(ex.Message, "Class Comparison Export", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        await Task.CompletedTask;
    }

    private void MoreActions_Click(object sender, RoutedEventArgs e)
    {
        if (MoreActionsButton.ContextMenu is not { } menu) return;
        menu.PlacementTarget = MoreActionsButton;
        menu.Placement = PlacementMode.Bottom;
        menu.IsOpen = true;
    }

    private async void SaveComparison_Click(object sender, RoutedEventArgs e)
    {
        if (Model is not { } model) return;
        var result = await SavePreferencesAsync(model);
        MessageBox.Show(result.Message, "Save Class Comparison", MessageBoxButton.OK, result.Succeeded ? MessageBoxImage.Information : MessageBoxImage.Warning);
    }

    private void ResetFilters_Click(object sender, RoutedEventArgs e)
    {
        if (Model is not { } model) return;
        model.SearchText = "";
        model.ShowDifferencesOnly = false;
        model.HighlightDifferences = true;
        SearchTextBox.Text = "";
        ClassComparisonEngine.Refresh(model);
        RefreshAfterMutation(model);
        QueuePreferenceSave(model);
    }

    private void Retry_Click(object sender, RoutedEventArgs e) => RetryRequested?.Invoke(this, EventArgs.Empty);
    private void ViewAuditHistory_Click(object sender, RoutedEventArgs e) => RequestNavigation("Notes & History");
    private void OpenAdvancedTab_Click(object sender, RoutedEventArgs e)
    {
        if (sender is FrameworkElement { Tag: string tabName }) RequestNavigation(tabName);
    }

    private void RequestNavigation(string tabName)
    {
        if (Model is { } model)
        {
            model.ActiveTab = tabName;
            QueuePreferenceSave(model);
        }
        NavigateTabRequested?.Invoke(this, new ClassComparisonNavigationEventArgs(tabName));
    }

    private void ClassComparisonWorkspaceView_SizeChanged(object sender, SizeChangedEventArgs e) => UpdateResponsiveLayout(e.NewSize.Width);

    private void UpdateResponsiveLayout(double width)
    {
        if (width <= 0) return;
        var narrow = width < 1050;
        Grid.SetColumn(BaselineSelectorHost, 0); Grid.SetColumnSpan(BaselineSelectorHost, narrow ? 3 : 1); Grid.SetRow(BaselineSelectorHost, 0);
        Grid.SetColumn(SwapClassesButton, narrow ? 0 : 1); Grid.SetColumnSpan(SwapClassesButton, narrow ? 3 : 1); Grid.SetRow(SwapClassesButton, narrow ? 1 : 0);
        Grid.SetColumn(ComparisonSelectorHost, narrow ? 0 : 2); Grid.SetColumnSpan(ComparisonSelectorHost, narrow ? 3 : 1); Grid.SetRow(ComparisonSelectorHost, narrow ? 2 : 0);
        SwapClassesButton.Margin = narrow ? new Thickness(0, 10, 0, 10) : new Thickness(15, 29, 15, 0);

        Grid.SetRow(ComparisonActionPanel, narrow ? 1 : 0); Grid.SetColumn(ComparisonActionPanel, narrow ? 0 : 1); Grid.SetColumnSpan(ComparisonActionPanel, narrow ? 2 : 1);
        ComparisonActionPanel.HorizontalAlignment = narrow ? HorizontalAlignment.Left : HorizontalAlignment.Right;
        ComparisonActionPanel.Margin = narrow ? new Thickness(0, 14, 0, 0) : new Thickness(0);

        var narrowHeader = width < 900;
        Grid.SetRow(TableHeaderControls, narrowHeader ? 1 : 0); Grid.SetColumn(TableHeaderControls, narrowHeader ? 0 : 1); Grid.SetColumnSpan(TableHeaderControls, narrowHeader ? 2 : 1);
        TableHeaderControls.HorizontalAlignment = narrowHeader ? HorizontalAlignment.Left : HorizontalAlignment.Right;
        TableHeaderControls.Margin = narrowHeader ? new Thickness(0, 12, 0, 0) : new Thickness(0);
        SearchTextBox.Width = narrowHeader ? 240 : 300;
    }

    private void QueuePreferenceSave(ClassComparisonWorkspaceViewModel model) => _ = SavePreferencesAsync(model);

    private async Task<DatabaseOperationResult> SavePreferencesAsync(ClassComparisonWorkspaceViewModel model)
    {
        await _preferenceSaveGate.WaitAsync();
        try
        {
            var result = await _dataService.SaveClassComparisonPreferencesAsync(model, ResolveUserName(model));
            if (!result.Succeeded) model.StatusText = result.Message;
            return result;
        }
        finally
        {
            _preferenceSaveGate.Release();
        }
    }

    private static string ResolveUserName(ClassComparisonWorkspaceViewModel model) =>
        string.IsNullOrWhiteSpace(model.Consultant) || model.Consultant.Equals("Not available", StringComparison.OrdinalIgnoreCase)
            ? Environment.UserName
            : model.Consultant;

    private static string SafeFileNamePart(string value)
    {
        var safe = new string((value ?? "Client").Select(ch => Path.GetInvalidFileNameChars().Contains(ch) ? '_' : ch).ToArray()).Trim();
        return string.IsNullOrWhiteSpace(safe) ? "Client" : safe;
    }
}

public sealed class ClassComparisonNavigationEventArgs(string tabName) : EventArgs
{
    public string TabName { get; } = tabName;
}

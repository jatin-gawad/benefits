using System.Globalization;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;
using ClosedXML.Excel;
using Microsoft.Win32;
using RenewalReviewBuilder.App.Repositories;
using RenewalReviewBuilder.App.Services;
using RenewalReviewBuilder.App.Services.FlexibleBenefits;

namespace RenewalReviewBuilder.App.Views;

public partial class PlanDesignWorkspaceView : UserControl
{
    private readonly PlanDesignWorkspaceDataService _dataService = new();
    private readonly FlexibleBenefitImportService _flexibleBenefitImportService = new();
    private readonly IFlexibleBenefitRepository _flexibleBenefitRepository = new FlexibleBenefitRepository();
    private readonly BenefitAdministrationService _benefitAdministrationService = new();
    private readonly ClassComparisonExportService _classComparisonExportService = new();
    private readonly PlanDesignComparisonExportService _planDesignComparisonExportService = new();
    private CancellationTokenSource? _flexibleBenefitImportCancellation;
    private long? _clientId;
    private string _clientName = "";
    private string _carrier = "";
    private Models.PlanDesignWorkspaceClientContext? _clientContext;
    private bool _isLoadingClientSelection;

    public PlanDesignWorkspaceView()
    {
        InitializeComponent();
        DataContext = CreateInitialModel();
        UpdateWorkspaceChrome("Class Comparison");
        UpdateClassWorkspaceLayout("Class Comparison");
        UpdateSecondaryNavigation("Class Comparison");
        UpdateHeaderResponsiveLayout(ActualWidth);
    }

    public async Task LoadClientAsync(
        long? clientId,
        string? clientName,
        string? carrier,
        Models.PlanDesignWorkspaceClientContext? clientContext = null,
        string? initialTab = null)
    {
        _isLoadingClientSelection = true;
        try
        {
            _clientId = clientId;
            _clientName = clientName ?? "";
            _carrier = carrier ?? "";
            _clientContext = clientContext;
            var model = await _dataService.LoadAsync(clientId, _clientName, _carrier, clientContext);
            DataContext = model;
            ResetBenefitAdministrationFilters();
            UpdateClassColumnHeaders(model);
            SelectWorkspaceTab(string.IsNullOrWhiteSpace(initialTab) ? model.ClassComparison.ActiveTab : initialTab);
            UpdateWorkspaceChrome();
            _ = Dispatcher.BeginInvoke(
                new Action(() => PlanWorkspaceTitleTextBlock.BringIntoView()),
                System.Windows.Threading.DispatcherPriority.Loaded);
        }
        finally
        {
            _isLoadingClientSelection = false;
        }
    }

    private async void WorkspaceTabs_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (!ReferenceEquals(sender, WorkspaceTabs) || WorkspaceTabs.SelectedItem is not TabItem selectedTab)
        {
            return;
        }

        var header = selectedTab.Header?.ToString() ?? "Class Comparison";
        if (!string.Equals(selectedTab.Tag?.ToString(), "Advanced", StringComparison.OrdinalIgnoreCase))
        {
            HideAdvancedTabs();
        }

        UpdateWorkspaceChrome(header);
        UpdateClassWorkspaceLayout(header);
        UpdateSecondaryNavigation(header);

        if (!_isLoadingClientSelection && DataContext is Models.PlanDesignWorkspaceViewModel model)
        {
            model.ClassComparison.ActiveTab = header;
            var result = await _dataService.SaveClassComparisonPreferencesAsync(
                model.ClassComparison,
                ResolveComparisonUserName(model.ClassComparison));
            if (!result.Succeeded)
            {
                model.ClassComparison.StatusText = result.Message;
            }
        }
    }

    private void UpdateWorkspaceChrome(string? selectedHeader = null)
    {
        var header = selectedHeader;
        if (string.IsNullOrWhiteSpace(header) && WorkspaceTabs?.SelectedItem is TabItem selectedTab)
        {
            header = selectedTab.Header?.ToString();
        }

        header = string.IsNullOrWhiteSpace(header) ? "Class Comparison" : header;

        var clientName = FirstNonBlank(
            _clientName,
            (DataContext as Models.PlanDesignWorkspaceViewModel)?.ClientName);

        PlanWorkspaceTitleTextBlock.Text = "Plan Design";
        PlanWorkspaceSubtitleTextBlock.Text = "Review the current program, compare employee classes, model alternatives, and assess overall value.";
        PlanWorkspaceHeaderActionsPanel.Visibility = header.Equals("Plan Design Comparison", StringComparison.OrdinalIgnoreCase)
            ? Visibility.Visible
            : Visibility.Collapsed;
    }

    private void UpdateClassWorkspaceLayout(string selectedHeader)
    {
        if (PlanAskBenefitsCard is null ||
            PlanNextActionsCard is null)
        {
            return;
        }

        var usesClassWorkspaceLayout = selectedHeader is
            "Class Comparison" or
            "Plan Design Comparison" or
            "Benefit Summary" or
            "Rate Comparison" or
            "Cost Analysis" or
            "Value Score" or
            "Summary" or
            "Notes & History";
        var contentVisibility = DataContext is Models.PlanDesignWorkspaceViewModel model
            ? model.PlanDesignContentVisibility
            : Visibility.Collapsed;

        PlanAskBenefitsCard.Visibility = usesClassWorkspaceLayout ? Visibility.Collapsed : contentVisibility;
        PlanNextActionsCard.Visibility = usesClassWorkspaceLayout ? Visibility.Collapsed : contentVisibility;
    }

    private void UpdateSecondaryNavigation(string selectedHeader)
    {
        if (PlanSecondaryNavigationPanel is null)
        {
            return;
        }

        var activeBrush = TryFindResource("PlanPurpleBrush") as Brush ?? Brushes.Purple;
        var inactiveBrush = TryFindResource("PlanTextSecondaryBrush") as Brush ?? Brushes.SlateGray;
        foreach (var button in PlanSecondaryNavigationPanel.Children.OfType<Button>())
        {
            var isActive = string.Equals(button.Tag?.ToString(), selectedHeader, StringComparison.OrdinalIgnoreCase);
            button.Foreground = isActive ? activeBrush : inactiveBrush;
            button.BorderBrush = isActive ? activeBrush : Brushes.Transparent;
            button.FontWeight = isActive ? FontWeights.Bold : FontWeights.SemiBold;
        }
    }

    private void PlanSecondaryNavigation_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button { Tag: string tabName })
        {
            SelectWorkspaceTab(tabName);
        }
    }

    public void OpenSubpage(string subpage)
    {
        SelectWorkspaceTab(subpage);
    }

    private void PlanDesignWorkspaceView_SizeChanged(object sender, SizeChangedEventArgs e)
    {
        UpdateHeaderResponsiveLayout(e.NewSize.Width);
    }

    private void UpdateHeaderResponsiveLayout(double width)
    {
        if (PlanWorkspaceHeaderActionsPanel is null ||
            PlanWorkspaceHeaderActionsColumn is null ||
            PlanWorkspaceHeaderPrimaryActions is null ||
            PlanWorkspaceHeaderSelectorActions is null)
        {
            return;
        }

        var stackActions = width > 0 && width < 1100;
        if (stackActions)
        {
            Grid.SetRow(PlanWorkspaceHeaderActionsPanel, 1);
            Grid.SetColumn(PlanWorkspaceHeaderActionsPanel, 0);
            Grid.SetColumnSpan(PlanWorkspaceHeaderActionsPanel, 2);
            PlanWorkspaceHeaderActionsPanel.HorizontalAlignment = HorizontalAlignment.Stretch;
            PlanWorkspaceHeaderActionsPanel.Margin = new Thickness(0, 14, 0, 0);
            PlanWorkspaceHeaderActionsColumn.Width = new GridLength(0);
            PlanWorkspaceHeaderPrimaryActions.HorizontalAlignment = HorizontalAlignment.Left;
            PlanWorkspaceHeaderSelectorActions.HorizontalAlignment = HorizontalAlignment.Left;
        }
        else
        {
            Grid.SetRow(PlanWorkspaceHeaderActionsPanel, 0);
            Grid.SetColumn(PlanWorkspaceHeaderActionsPanel, 1);
            Grid.SetColumnSpan(PlanWorkspaceHeaderActionsPanel, 1);
            PlanWorkspaceHeaderActionsPanel.HorizontalAlignment = HorizontalAlignment.Right;
            PlanWorkspaceHeaderActionsPanel.Margin = new Thickness(18, 0, 0, 0);
            PlanWorkspaceHeaderActionsColumn.Width = GridLength.Auto;
            PlanWorkspaceHeaderPrimaryActions.HorizontalAlignment = HorizontalAlignment.Right;
            PlanWorkspaceHeaderSelectorActions.HorizontalAlignment = HorizontalAlignment.Right;
        }
    }

    private void SelectWorkspaceTab(string? tabName, bool revealAdvanced = false)
    {
        var targetName = string.IsNullOrWhiteSpace(tabName) ? "Class Comparison" : tabName.Trim();
        var target = WorkspaceTabs.Items
            .OfType<TabItem>()
            .FirstOrDefault(tab => string.Equals(tab.Header?.ToString(), targetName, StringComparison.OrdinalIgnoreCase));
        target ??= WorkspaceTabs.Items.OfType<TabItem>().FirstOrDefault();
        if (target is null)
        {
            return;
        }

        if (revealAdvanced && string.Equals(target.Tag?.ToString(), "Advanced", StringComparison.OrdinalIgnoreCase))
        {
            target.Visibility = Visibility.Visible;
        }

        WorkspaceTabs.SelectedItem = target;
        target.BringIntoView();
        UpdateWorkspaceChrome(target.Header?.ToString());
        UpdateClassWorkspaceLayout(target.Header?.ToString() ?? "Class Comparison");
        UpdateSecondaryNavigation(target.Header?.ToString() ?? "Class Comparison");
    }

    private void HideAdvancedTabs()
    {
        foreach (var tab in WorkspaceTabs.Items
                     .OfType<TabItem>()
                     .Where(item => string.Equals(item.Tag?.ToString(), "Advanced", StringComparison.OrdinalIgnoreCase)))
        {
            tab.Visibility = Visibility.Collapsed;
        }
    }

    private void ClassComparison_NavigateTabRequested(object? sender, ClassComparisonNavigationEventArgs e)
    {
        var isAdvanced = e.TabName is "Benefit Details" or "Paramedical" or "Exclusions & Limitations" or "Readiness Checklist" or "Raw Extraction";
        SelectWorkspaceTab(e.TabName, isAdvanced);
    }

    private async void ClassComparison_RetryRequested(object? sender, EventArgs e)
    {
        await LoadClientAsync(_clientId, _clientName, _carrier, _clientContext, "Class Comparison");
    }

    private void PlanComparisonAllBenefits_Click(object sender, RoutedEventArgs e) =>
        SetPlanComparisonViewMode(Models.PlanDesignComparisonViewModes.AllBenefits);

    private void PlanComparisonChangedOnly_Click(object sender, RoutedEventArgs e) =>
        SetPlanComparisonViewMode(Models.PlanDesignComparisonViewModes.ChangedOnly);

    private void PlanComparisonKeyBenefits_Click(object sender, RoutedEventArgs e) =>
        SetPlanComparisonViewMode(Models.PlanDesignComparisonViewModes.KeyBenefitsOnly);

    private void SetPlanComparisonViewMode(string viewMode)
    {
        if (DataContext is Models.PlanDesignWorkspaceViewModel model)
        {
            model.PlanComparison.ViewMode = viewMode;
        }
    }

    private void PlanComparisonExpandAll_Click(object sender, RoutedEventArgs e)
    {
        if (DataContext is Models.PlanDesignWorkspaceViewModel model)
        {
            model.PlanComparison.SetAllGroupsExpanded(true);
        }
    }

    private void PlanComparisonCollapseAll_Click(object sender, RoutedEventArgs e)
    {
        if (DataContext is Models.PlanDesignWorkspaceViewModel model)
        {
            model.PlanComparison.SetAllGroupsExpanded(false);
        }
    }

    private void PlanComparisonGroupToggle_Click(object sender, RoutedEventArgs e)
    {
        if (DataContext is Models.PlanDesignWorkspaceViewModel model &&
            (sender as FrameworkElement)?.DataContext is Models.PlanDesignComparisonDisplayRow row)
        {
            model.PlanComparison.ToggleGroup(row.Group);
        }
    }

    private void PlanComparisonDetails_Click(object sender, RoutedEventArgs e)
    {
        if ((sender as FrameworkElement)?.DataContext is Models.PlanDesignComparisonDisplayRow row &&
            row.Provision is not null)
        {
            row.Provision.IsDetailsExpanded = !row.Provision.IsDetailsExpanded;
        }
    }

    private void PlanComparisonResetFilters_Click(object sender, RoutedEventArgs e)
    {
        if (DataContext is not Models.PlanDesignWorkspaceViewModel model)
        {
            return;
        }

        model.PlanComparison.SearchText = "";
        model.PlanComparison.HighlightDifferences = true;
        model.PlanComparison.ViewMode = model.PlanComparison.HasDifferences
            ? Models.PlanDesignComparisonViewModes.ChangedOnly
            : Models.PlanDesignComparisonViewModes.AllBenefits;
        model.PlanComparison.SetAllGroupsExpanded(true);
    }


    private void DownloadPlanDesignSummary_Click(object sender, RoutedEventArgs e)
    {
        if (DataContext is not Models.PlanDesignWorkspaceViewModel model || !model.PlanComparison.HasComparisonScenario)
        {
            MessageBox.Show(
                "No plan design comparison is loaded to download.",
                "Plan Design Comparison",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        try
        {
            AppPaths.EnsureCreated();
            var outputFolder = Path.Combine(AppPaths.AppDataFolder, "Reports");
            Directory.CreateDirectory(outputFolder);
            var path = Path.Combine(
                outputFolder,
                $"PlanDesignComparison_{SafeFileNamePart(model.ClientName)}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");

            var result = _planDesignComparisonExportService.ExportToExcel(model.PlanComparison, path);
            MessageBox.Show(
                $"The plan design comparison was downloaded to:\n\n{result.OutputPath}",
                "Plan Design Comparison",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            Process.Start(new ProcessStartInfo(outputFolder) { UseShellExecute = true });
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            MessageBox.Show(
                ex.Message,
                "Plan Design Comparison",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }
    }

    private async void SaveComparisonNote_Click(object sender, RoutedEventArgs e)
    {
        if (DataContext is not Models.PlanDesignWorkspaceViewModel workspaceModel ||
            workspaceModel.ClassComparison is not { ClientId: not null } comparisonModel)
        {
            MessageBox.Show(
                "Select a saved client before adding a class comparison note.",
                "Class Comparison Note",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        try
        {
            var className = ComparisonNoteClassComboBox.SelectedValue?.ToString() ?? comparisonModel.BaselineClassName;
            var note = await _dataService.SaveClassComparisonNoteAsync(
                comparisonModel.ClientId.Value,
                comparisonModel.AnalystNote,
                className,
                ResolveComparisonUserName(comparisonModel));
            comparisonModel.Notes.Insert(0, note);
            comparisonModel.History.Insert(0, new Models.ClassComparisonHistoryRow
            {
                HistoryId = note.NoteId,
                DateText = note.CreatedAtText,
                Author = note.Author,
                ClassName = note.ClassName,
                FieldName = "Analyst note",
                PreviousValue = "Not available",
                NewValue = note.NoteText,
                Reason = "Analyst note saved",
                SourceDocument = "Class Comparison workspace"
            });
            comparisonModel.AnalystNote = "";
            comparisonModel.NotifyCollectionsChanged();
            comparisonModel.StatusText = "Analyst note saved to SQLite.";
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            MessageBox.Show(ex.Message, "Class Comparison Note", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private static string ResolveComparisonUserName(Models.ClassComparisonWorkspaceViewModel model)
    {
        return string.IsNullOrWhiteSpace(model.Consultant) || model.Consultant.Equals("Not available", StringComparison.OrdinalIgnoreCase)
            ? Environment.UserName
            : model.Consultant;
    }

    private void UpdateClassColumnHeaders(Models.PlanDesignWorkspaceViewModel model)
    {
        var class1Header = CompactClassHeader(FirstNonBlank(model.Class1Name, "All Classes"));
        var class2Header = CompactClassHeader(FirstNonBlank(model.Class2Name));
        var class3Header = CompactClassHeader(FirstNonBlank(model.Class3Name));
        var hasSecondClass = !string.IsNullOrWhiteSpace(class2Header);
        var hasThirdClass = !string.IsNullOrWhiteSpace(class3Header);

        ComparisonClass1Column.Header = class1Header;
        EhcClass1Column.Header = class1Header;
        ParamedicalClass1Column.Header = class1Header;

        ComparisonClass2Column.Header = hasSecondClass ? class2Header : "Class 2";
        EhcClass2Column.Header = hasSecondClass ? class2Header : "Class 2";
        ParamedicalClass2Column.Header = hasSecondClass ? class2Header : "Class 2";

        ComparisonClass3Column.Header = hasThirdClass ? class3Header : "Class 3";
        EhcClass3Column.Header = hasThirdClass ? class3Header : "Class 3";
        ParamedicalClass3Column.Header = hasThirdClass ? class3Header : "Class 3";

        var secondColumnVisibility = hasSecondClass ? Visibility.Visible : Visibility.Collapsed;
        ComparisonClass2Column.Visibility = secondColumnVisibility;
        EhcClass2Column.Visibility = secondColumnVisibility;
        ParamedicalClass2Column.Visibility = secondColumnVisibility;

        var thirdColumnVisibility = hasThirdClass ? Visibility.Visible : Visibility.Collapsed;
        ComparisonClass3Column.Visibility = thirdColumnVisibility;
        EhcClass3Column.Visibility = thirdColumnVisibility;
        ParamedicalClass3Column.Visibility = thirdColumnVisibility;

    }

    private static string CompactClassHeader(string value)
    {
        var text = FirstNonBlank(value);
        if (string.IsNullOrWhiteSpace(text))
        {
            return "";
        }

        var splitIndex = text.IndexOf(" - ", StringComparison.Ordinal);
        var prefix = splitIndex > 0 ? text[..splitIndex].Trim() : text;
        if (prefix.StartsWith("Class ", StringComparison.OrdinalIgnoreCase) && prefix.Length <= 18)
        {
            return prefix;
        }

        return text.Length <= 34
            ? text
            : $"{text[..31]}...";
    }

    private static string FirstNonBlank(params string?[] values)
    {
        foreach (var value in values)
        {
            if (!string.IsNullOrWhiteSpace(value))
            {
                return value.Trim();
            }
        }

        return string.Empty;
    }

    private async void RefreshData_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            await LoadClientAsync(_clientId, _clientName, _carrier, _clientContext);
            if (DataContext is Models.PlanDesignWorkspaceViewModel model)
            {
                model.LastRefreshedText = $"{model.LastRefreshedText} Refreshed {DateTime.Now:MMM d, yyyy h:mm tt}.";
            }
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            MessageBox.Show(
                "Plan Design data could not be refreshed. The workspace is still open.",
                "Plan Design Workspace",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }
    }

    private void ViewSourceDocument_Click(object sender, RoutedEventArgs e)
    {
        if (DataContext is Models.PlanDesignWorkspaceViewModel model &&
            !string.IsNullOrWhiteSpace(model.SourceDocumentPath) &&
            File.Exists(model.SourceDocumentPath))
        {
            Process.Start(new ProcessStartInfo(model.SourceDocumentPath) { UseShellExecute = true });
            return;
        }

        MessageBox.Show(
            "No linked source document file path is saved for this client yet. The extracted plan design can still be reviewed from SQLite.",
            "Plan Design Workspace",
            MessageBoxButton.OK,
            MessageBoxImage.Information);
    }

    private async void ImportFlexibleBenefits_Click(object sender, RoutedEventArgs e)
    {
        if (!_clientId.HasValue || _clientId <= 0)
        {
            MessageBox.Show(
                "Select a saved client before importing an HCSA or LSA document.",
                "Flexible Benefits Import",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        var dialog = new OpenFileDialog
        {
            Title = "Select HCSA, LSA, or Flexible Benefits PDF",
            Filter = "PDF documents (*.pdf)|*.pdf|All files (*.*)|*.*",
            CheckFileExists = true,
            Multiselect = false
        };
        if (dialog.ShowDialog(Window.GetWindow(this)) != true)
        {
            return;
        }

        var importMode = Models.FlexibleBenefitImportMode.Cancel;
        if (DataContext is Models.PlanDesignWorkspaceViewModel currentModel && currentModel.FlexibleBenefits.HasPlan)
        {
            var versionChoice = MessageBox.Show(
                "A flexible-benefits plan already exists for this client.\n\n" +
                "Yes: preserve both and import as a new version\n" +
                "No: preserve the prior version and mark it replaced\n" +
                "Cancel: do not import",
                "Choose Version Action",
                MessageBoxButton.YesNoCancel,
                MessageBoxImage.Question,
                MessageBoxResult.Cancel);
            if (versionChoice == MessageBoxResult.Cancel)
            {
                return;
            }

            importMode = versionChoice == MessageBoxResult.Yes
                ? Models.FlexibleBenefitImportMode.ImportAsNewVersion
                : Models.FlexibleBenefitImportMode.ReplaceExistingVersion;
        }

        if (sender is Button button)
        {
            button.IsEnabled = false;
        }

        _flexibleBenefitImportCancellation?.Cancel();
        _flexibleBenefitImportCancellation?.Dispose();
        _flexibleBenefitImportCancellation = new CancellationTokenSource();
        try
        {
            if (DataContext is Models.PlanDesignWorkspaceViewModel loadingModel)
            {
                loadingModel.LastRefreshedText = "Extracting the flexible-benefits document locally. The signed PDF remains the legal source.";
            }

            var result = await _flexibleBenefitImportService.ImportAsync(
                new Models.FlexibleBenefitImportRequest
                {
                    ClientId = _clientId,
                    ClientName = _clientName,
                    FilePath = dialog.FileName,
                    ImportMode = importMode,
                    ImportedBy = Environment.UserName
                },
                _flexibleBenefitImportCancellation.Token);

            if (result.Succeeded)
            {
                await LoadClientAsync(_clientId, _clientName, _carrier, _clientContext);
                SelectWorkspaceTab("Flexible Benefits");
            }

            var icon = result.Outcome switch
            {
                Models.FlexibleBenefitImportOutcome.Completed => MessageBoxImage.Information,
                Models.FlexibleBenefitImportOutcome.NeedsReview => MessageBoxImage.Warning,
                Models.FlexibleBenefitImportOutcome.Duplicate => MessageBoxImage.Information,
                Models.FlexibleBenefitImportOutcome.ExistingVersion => MessageBoxImage.Warning,
                _ => MessageBoxImage.Error
            };
            MessageBox.Show(result.Message, "Flexible Benefits Import", MessageBoxButton.OK, icon);
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            MessageBox.Show(
                "The flexible-benefits document could not be imported. No partial plan version was saved.",
                "Flexible Benefits Import",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        finally
        {
            if (sender is Button finishedButton)
            {
                finishedButton.IsEnabled = true;
            }
        }
    }

    private void ViewFlexibleBenefitDocument_Click(object sender, RoutedEventArgs e)
    {
        if (DataContext is not Models.PlanDesignWorkspaceViewModel model ||
            !model.FlexibleBenefits.HasPlan ||
            string.IsNullOrWhiteSpace(model.FlexibleBenefits.SourceDocumentPath) ||
            !File.Exists(model.FlexibleBenefits.SourceDocumentPath))
        {
            MessageBox.Show(
                "The linked flexible-benefits source PDF is not available at its saved path.",
                "Flexible Benefits Source",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        Process.Start(new ProcessStartInfo(model.FlexibleBenefits.SourceDocumentPath) { UseShellExecute = true });
    }

    private void ViewFlexibleBenefitSourcePage_Click(object sender, RoutedEventArgs e)
    {
        if (DataContext is not Models.PlanDesignWorkspaceViewModel model ||
            FlexibleSourceFieldsGrid.SelectedItem is not Models.FlexibleBenefitSourceFieldDisplayRow sourceField)
        {
            MessageBox.Show(
                "Select an extraction-detail row first.",
                "View Source Page",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        var pageMatch = Regex.Match(sourceField.SourcePage ?? "", "\\d+");
        if (!pageMatch.Success ||
            string.IsNullOrWhiteSpace(model.FlexibleBenefits.SourceDocumentPath) ||
            !File.Exists(model.FlexibleBenefits.SourceDocumentPath))
        {
            MessageBox.Show(
                "This field does not have an available source page link.",
                "View Source Page",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        var sourceUri = new Uri(model.FlexibleBenefits.SourceDocumentPath).AbsoluteUri + $"#page={pageMatch.Value}";
        Process.Start(new ProcessStartInfo(sourceUri) { UseShellExecute = true });
    }

    private async void EditFlexibleBenefitSourceField_Click(object sender, RoutedEventArgs e)
    {
        if (FlexibleSourceFieldsGrid.SelectedItem is not Models.FlexibleBenefitSourceFieldDisplayRow sourceField)
        {
            MessageBox.Show(
                "Select an extraction-detail row to edit and confirm.",
                "Flexible Benefits Review",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        var dialog = new FlexibleBenefitCorrectionDialog(sourceField)
        {
            Owner = Window.GetWindow(this)
        };
        if (dialog.ShowDialog() != true)
        {
            return;
        }

        var result = await _flexibleBenefitRepository.CorrectSourceFieldAsync(
            sourceField.SourceFieldId,
            dialog.ConfirmedValue,
            Environment.UserName,
            dialog.Reason);
        if (result.Succeeded)
        {
            await LoadClientAsync(_clientId, _clientName, _carrier, _clientContext);
            SelectWorkspaceTab("Flexible Benefits");
        }

        MessageBox.Show(
            result.Message,
            "Flexible Benefits Review",
            MessageBoxButton.OK,
            result.Succeeded ? MessageBoxImage.Information : MessageBoxImage.Warning);
    }

    private async void MarkFlexibleBenefitsReviewed_Click(object sender, RoutedEventArgs e)
    {
        if (DataContext is not Models.PlanDesignWorkspaceViewModel model || !model.FlexibleBenefits.PlanId.HasValue)
        {
            return;
        }

        if (MessageBox.Show(
                "Mark the current flexible-benefits plan version and all of its extracted fields as reviewed?",
                "Mark Flexible Benefits Reviewed",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question,
                MessageBoxResult.No) != MessageBoxResult.Yes)
        {
            return;
        }

        var result = await _flexibleBenefitRepository.MarkReviewedAsync(
            model.FlexibleBenefits.PlanId.Value,
            Environment.UserName);
        if (result.Succeeded)
        {
            await LoadClientAsync(_clientId, _clientName, _carrier, _clientContext);
            SelectWorkspaceTab("Flexible Benefits");
        }

        MessageBox.Show(
            result.Message,
            "Flexible Benefits Review",
            MessageBoxButton.OK,
            result.Succeeded ? MessageBoxImage.Information : MessageBoxImage.Warning);
    }

    private async void CompareFlexibleBenefitVersions_Click(object sender, RoutedEventArgs e)
    {
        if (DataContext is not Models.PlanDesignWorkspaceViewModel model ||
            !model.FlexibleBenefits.PlanId.HasValue)
        {
            return;
        }

        var prior = FlexibleVersionsGrid.SelectedItem as Models.FlexibleBenefitVersionDisplayRow;
        if (prior is null || prior.IsCurrent)
        {
            prior = model.FlexibleBenefits.Versions.FirstOrDefault(version => !version.IsCurrent);
        }

        if (prior is null)
        {
            MessageBox.Show(
                "No prior flexible-benefits plan version is available for comparison.",
                "Compare Flexible Benefits Versions",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        var differences = await _flexibleBenefitRepository.CompareVersionsAsync(
            model.FlexibleBenefits.PlanId.Value,
            prior.PlanId);
        var dialog = new FlexibleBenefitVersionComparisonDialog(
            model.FlexibleBenefits.VersionNumber,
            prior.VersionNumber,
            differences)
        {
            Owner = Window.GetWindow(this)
        };
        dialog.ShowDialog();
    }

    private void ResetBenefitAdministrationFilters()
    {
        if (BenefitAdministrationInsurerFilter is null ||
            BenefitAdministrationProviderFilter is null ||
            BenefitAdministrationStatusFilter is null ||
            BenefitAdministrationOptionalFilter is null)
        {
            return;
        }

        BenefitAdministrationInsurerFilter.SelectedIndex = 0;
        BenefitAdministrationProviderFilter.SelectedIndex = 0;
        BenefitAdministrationStatusFilter.SelectedIndex = 0;
        BenefitAdministrationOptionalFilter.SelectedIndex = 0;
        ApplyBenefitAdministrationFilter();
    }

    private void BenefitAdministrationFilter_Changed(object sender, SelectionChangedEventArgs e)
    {
        ApplyBenefitAdministrationFilter();
    }

    private void ApplyBenefitAdministrationFilter()
    {
        if (BenefitAdministrationDataGrid?.ItemsSource is null)
        {
            return;
        }

        var insurer = BenefitAdministrationInsurerFilter?.SelectedItem?.ToString() ?? "All insurers";
        var provider = BenefitAdministrationProviderFilter?.SelectedItem?.ToString() ?? "All providers";
        var status = BenefitAdministrationStatusFilter?.SelectedItem?.ToString() ?? "All statuses";
        var optional = BenefitAdministrationOptionalFilter?.SelectedItem?.ToString() ?? "All benefits";
        var view = CollectionViewSource.GetDefaultView(BenefitAdministrationDataGrid.ItemsSource);
        view.Filter = item =>
        {
            if (item is not Models.BenefitAdministrationWorkspaceRow row)
            {
                return false;
            }

            return (insurer.Equals("All insurers", StringComparison.OrdinalIgnoreCase) || row.InsuranceCompany.Equals(insurer, StringComparison.OrdinalIgnoreCase)) &&
                   (provider.Equals("All providers", StringComparison.OrdinalIgnoreCase) || row.ProviderAssignments.Any(assignment => assignment.ProviderName.Equals(provider, StringComparison.OrdinalIgnoreCase))) &&
                   (status.Equals("All statuses", StringComparison.OrdinalIgnoreCase) || row.Status.Equals(status, StringComparison.OrdinalIgnoreCase)) &&
                   (optional.Equals("All benefits", StringComparison.OrdinalIgnoreCase) ||
                    optional.Equals("Optional", StringComparison.OrdinalIgnoreCase) == row.IsOptional);
        };
        view.Refresh();
    }

    private void ViewBenefitAdministrationSource_Click(object sender, RoutedEventArgs e)
    {
        var row = (sender as FrameworkElement)?.DataContext as Models.BenefitAdministrationWorkspaceRow;
        if (row is null || string.IsNullOrWhiteSpace(row.SourcePath) || !File.Exists(row.SourcePath))
        {
            MessageBox.Show("The saved source document could not be found.", "Benefit Administration", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        Process.Start(new ProcessStartInfo(row.SourcePath) { UseShellExecute = true });
    }

    private async void ConfirmBenefitAdministration_Click(object sender, RoutedEventArgs e)
    {
        var row = (sender as FrameworkElement)?.DataContext as Models.BenefitAdministrationWorkspaceRow;
        if (row is null)
        {
            return;
        }

        try
        {
            await _benefitAdministrationService.ConfirmAsync(row.ClientBenefitAdministrationId);
            await LoadClientAsync(_clientId, _clientName, _carrier, _clientContext);
            SelectWorkspaceTab("Benefit Administration");
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            MessageBox.Show(ex.Message, "Benefit Administration", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private async void EditBenefitAdministration_Click(object sender, RoutedEventArgs e)
    {
        var row = (sender as FrameworkElement)?.DataContext as Models.BenefitAdministrationWorkspaceRow;
        if (row is null)
        {
            return;
        }

        try
        {
            var organizations = await _benefitAdministrationService.GetOrganizationsAsync();
            var dialog = new BenefitAdministrationEditDialog(row, organizations) { Owner = Window.GetWindow(this) };
            if (dialog.ShowDialog() != true)
            {
                return;
            }

            await _benefitAdministrationService.CorrectInsurerAsync(row.ClientBenefitAdministrationId, dialog.SelectedInsurerOrganizationId);
            await _benefitAdministrationService.CorrectPolicyNumberAsync(row.ClientBenefitAdministrationId, dialog.PolicyNumber);
            await LoadClientAsync(_clientId, _clientName, _carrier, _clientContext);
            SelectWorkspaceTab("Benefit Administration");
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            MessageBox.Show(ex.Message, "Benefit Administration", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private async void AddBenefitProvider_Click(object sender, RoutedEventArgs e)
    {
        var row = (sender as FrameworkElement)?.DataContext as Models.BenefitAdministrationWorkspaceRow;
        if (row is null)
        {
            return;
        }

        try
        {
            var organizations = await _benefitAdministrationService.GetOrganizationsAsync();
            var dialog = new BenefitProviderAssignmentDialog(row.Benefit, organizations) { Owner = Window.GetWindow(this) };
            if (dialog.ShowDialog() != true || dialog.RequestedAction != BenefitProviderDialogAction.Save || dialog.SelectedProvider is null)
            {
                return;
            }

            await _benefitAdministrationService.AddProviderAssignmentAsync(new Models.BenefitServiceProviderAssignmentRecord
            {
                ClientBenefitAdministrationId = row.ClientBenefitAdministrationId,
                ProviderOrganizationId = dialog.SelectedProvider.OrganizationId,
                ProviderRoleCode = dialog.SelectedRole,
                SourceDisplayName = dialog.SelectedProvider.DisplayName,
                EffectiveDate = dialog.EffectiveDate,
                TerminationDate = dialog.TerminationDate,
                SourceDocumentId = row.SourceDocumentId,
                SourcePageNumber = row.SourcePageNumber,
                IsPrimary = true
            });
            await LoadClientAsync(_clientId, _clientName, _carrier, _clientContext);
            SelectWorkspaceTab("Benefit Administration");
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            MessageBox.Show(ex.Message, "Benefit Administration", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private async void ManageBenefitProviders_Click(object sender, RoutedEventArgs e)
    {
        var row = (sender as FrameworkElement)?.DataContext as Models.BenefitAdministrationWorkspaceRow;
        var activeAssignments = row?.ProviderAssignments.Where(assignment => assignment.IsActive).ToList() ?? [];
        if (row is null || activeAssignments.Count == 0)
        {
            MessageBox.Show("No active service-provider assignment is available to manage.", "Benefit Administration", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        try
        {
            Models.BenefitServiceProviderAssignmentRecord? assignment = activeAssignments[0];
            if (activeAssignments.Count > 1)
            {
                var choice = new BenefitProviderChoiceDialog(activeAssignments) { Owner = Window.GetWindow(this) };
                if (choice.ShowDialog() != true)
                {
                    return;
                }
                assignment = choice.SelectedAssignment;
            }

            if (assignment is null)
            {
                return;
            }

            var organizations = await _benefitAdministrationService.GetOrganizationsAsync();
            var dialog = new BenefitProviderAssignmentDialog(row.Benefit, organizations, assignment) { Owner = Window.GetWindow(this) };
            if (dialog.ShowDialog() != true)
            {
                return;
            }

            if (dialog.RequestedAction == BenefitProviderDialogAction.Remove)
            {
                await _benefitAdministrationService.RemoveProviderAssignmentAsync(assignment.BenefitServiceProviderAssignmentId);
            }
            else if (dialog.RequestedAction == BenefitProviderDialogAction.EndDate && dialog.TerminationDate.HasValue)
            {
                await _benefitAdministrationService.EndDateProviderAssignmentAsync(
                    assignment.BenefitServiceProviderAssignmentId,
                    dialog.TerminationDate.Value);
            }
            else if (dialog.RequestedAction == BenefitProviderDialogAction.Save && dialog.SelectedProvider is not null)
            {
                assignment.ProviderOrganizationId = dialog.SelectedProvider.OrganizationId;
                assignment.ProviderRoleCode = dialog.SelectedRole;
                assignment.SourceDisplayName = dialog.SelectedProvider.DisplayName;
                assignment.EffectiveDate = dialog.EffectiveDate;
                assignment.TerminationDate = dialog.TerminationDate;
                await _benefitAdministrationService.UpdateProviderAssignmentAsync(assignment);
            }

            await LoadClientAsync(_clientId, _clientName, _carrier, _clientContext);
            SelectWorkspaceTab("Benefit Administration");
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            MessageBox.Show(ex.Message, "Benefit Administration", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private async void ExportBenefitAdministration_Click(object sender, RoutedEventArgs e)
    {
        if (!_clientId.HasValue)
        {
            MessageBox.Show("Select a client before exporting.", "Benefit Administration", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var rows = await _benefitAdministrationService.GetClientFacingExportRowsAsync(_clientId.Value);
        if (rows.Count == 0)
        {
            MessageBox.Show("No active Benefit Administration rows are available to export.", "Benefit Administration", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        AppPaths.EnsureCreated();
        var outputFolder = Path.Combine(AppPaths.AppDataFolder, "Reports");
        Directory.CreateDirectory(outputFolder);
        var path = Path.Combine(outputFolder, $"BenefitAdministration_{SafeFileNamePart(_clientName)}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.Worksheets.Add("Benefit Administration");
            sheet.Cell(1, 1).Value = "Benefit";
            sheet.Cell(1, 2).Value = "Insurance Company";
            sheet.Cell(1, 3).Value = "Service Provider";
            sheet.Cell(1, 4).Value = "Policy Number";
            for (var index = 0; index < rows.Count; index++)
            {
                sheet.Cell(index + 2, 1).Value = rows[index].Benefit;
                sheet.Cell(index + 2, 2).Value = rows[index].InsuranceCompany;
                sheet.Cell(index + 2, 3).Value = rows[index].ServiceProvider;
                sheet.Cell(index + 2, 4).Value = rows[index].PolicyNumber;
            }
            var table = sheet.Range(1, 1, rows.Count + 1, 4).CreateTable("BenefitAdministrationSummary");
            table.Theme = XLTableTheme.TableStyleMedium2;
            sheet.SheetView.FreezeRows(1);
            sheet.Columns().AdjustToContents(12, 50);
            workbook.SaveAs(path);
        }

        Process.Start(new ProcessStartInfo(outputFolder) { UseShellExecute = true });
    }

    private void EmptyUploadBooklet_Click(object sender, RoutedEventArgs e)
    {
        MessageBox.Show(
            "Upload Booklet will use the existing Benefit Booklet upload workflow from the main platform actions.",
            "Plan Design Workspace",
            MessageBoxButton.OK,
            MessageBoxImage.Information);
    }

    private void EmptyExtractPlanDesign_Click(object sender, RoutedEventArgs e)
    {
        MessageBox.Show(
            "Extract Plan Design will use the existing Document Intelligence extraction workflow for this selected client.",
            "Plan Design Workspace",
            MessageBoxButton.OK,
            MessageBoxImage.Information);
    }

    private void EmptyLinkSourceDocument_Click(object sender, RoutedEventArgs e)
    {
        MessageBox.Show(
            "Link Source Document will connect a saved document record to this client's plan design once document linking is enabled.",
            "Plan Design Workspace",
            MessageBoxButton.OK,
            MessageBoxImage.Information);
    }

    private void ExportRawRows_Click(object sender, RoutedEventArgs e)
    {
        if (DataContext is not Models.PlanDesignWorkspaceViewModel model || model.RawExtractionRows.Count == 0)
        {
            MessageBox.Show(
                "No raw plan design rows are loaded for this client yet.",
                "Raw Extraction",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        AppPaths.EnsureCreated();
        var outputFolder = Path.Combine(AppPaths.AppDataFolder, "Reports");
        Directory.CreateDirectory(outputFolder);

        var safeClientName = SafeFileNamePart(model.ClientName);
        var path = Path.Combine(outputFolder, $"PlanDesignRawExtraction_{safeClientName}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");

        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.Worksheets.Add("Raw Extraction");
            var headers = new[]
            {
                "Client",
                "Carrier",
                "Source Document",
                "Benefit",
                "Category",
                "Feature",
                "Value",
                "Coinsurance",
                "Deductible",
                "Maximum",
                "Frequency",
                "Termination Age",
                "Class",
                "Source Page",
                "Confidence",
                "Source Text",
                "Loaded"
            };

            for (var column = 0; column < headers.Length; column++)
            {
                sheet.Cell(1, column + 1).Value = headers[column];
            }

            var rowNumber = 2;
            foreach (var row in model.RawExtractionRows)
            {
                sheet.Cell(rowNumber, 1).Value = model.ClientName;
                sheet.Cell(rowNumber, 2).Value = model.Carrier;
                sheet.Cell(rowNumber, 3).Value = model.SourceDocument;
                sheet.Cell(rowNumber, 4).Value = row.Benefit;
                sheet.Cell(rowNumber, 5).Value = row.Category;
                sheet.Cell(rowNumber, 6).Value = row.FeatureName;
                sheet.Cell(rowNumber, 7).Value = row.FeatureValue;
                sheet.Cell(rowNumber, 8).Value = row.Coinsurance;
                sheet.Cell(rowNumber, 9).Value = row.Deductible;
                sheet.Cell(rowNumber, 10).Value = row.Maximum;
                sheet.Cell(rowNumber, 11).Value = row.Frequency;
                sheet.Cell(rowNumber, 12).Value = row.TerminationAge;
                sheet.Cell(rowNumber, 13).Value = row.ClassText;
                sheet.Cell(rowNumber, 14).Value = row.SourcePage;
                sheet.Cell(rowNumber, 15).Value = row.ConfidenceText;
                sheet.Cell(rowNumber, 16).Value = row.SourceText;
                sheet.Cell(rowNumber, 17).Value = row.LoadedText;
                rowNumber++;
            }

            var usedRange = sheet.Range(1, 1, Math.Max(rowNumber - 1, 1), headers.Length);
            usedRange.CreateTable("RawPlanDesignExtraction");
            sheet.Row(1).Style.Font.Bold = true;
            sheet.SheetView.FreezeRows(1);
            sheet.Columns().AdjustToContents(8, 48);
            sheet.Column(16).Width = 70;
            sheet.Column(16).Style.Alignment.WrapText = true;
            workbook.SaveAs(path);
        }

        MessageBox.Show(
            $"Raw extraction rows were exported to:\n\n{path}",
            "Raw Extraction",
            MessageBoxButton.OK,
            MessageBoxImage.Information);

        Process.Start(new ProcessStartInfo(outputFolder) { UseShellExecute = true });
    }

    private static string SafeFileNamePart(string value)
    {
        var safe = new string((value ?? "Client")
            .Select(ch => Path.GetInvalidFileNameChars().Contains(ch) ? '_' : ch)
            .ToArray());
        safe = safe.Trim();
        return string.IsNullOrWhiteSpace(safe) ? "Client" : safe;
    }

    private static Models.PlanDesignWorkspaceViewModel CreateInitialModel()
    {
        var model = new Models.PlanDesignWorkspaceViewModel
        {
            ClientName = "Select a client",
            Carrier = "",
            SourceDocument = "Not loaded",
            Quote = "No client selected",
            PolicyEffectiveDate = "N/A",
            CalcEffectiveDate = "N/A",
            RatesVersion = "N/A",
            ClassesFound = "0",
            Class1Name = "All Classes",
            Readiness = "Select client",
            HasPlanDesign = false,
            EmptyStateMessage = "Select a client from the Client Workspace to load client-specific plan design data.",
            LastRefreshedText = "No client selected."
        };

        model.ClientSelectorOptions.Add(new Models.PlanDesignClientOption { ClientName = "Select a client" });
        model.SummaryCards.Add(new Models.PlanDesignSummaryCard { Label = "Client", Value = "Select a client", Detail = "Open from Client Workspace" });
        model.SummaryCards.Add(new Models.PlanDesignSummaryCard { Label = "Carrier", Value = "N/A", Detail = "No client selected" });
        model.SummaryCards.Add(new Models.PlanDesignSummaryCard { Label = "Source Document", Value = "N/A", Detail = "No source selected" });
        model.SummaryCards.Add(new Models.PlanDesignSummaryCard { Label = "Effective Date", Value = "N/A", Detail = "No client selected" });
        model.SummaryCards.Add(new Models.PlanDesignSummaryCard { Label = "Classes Found", Value = "0", Detail = "No plan design loaded" });
        model.SummaryCards.Add(new Models.PlanDesignSummaryCard { Label = "Readiness", Value = "Select client", Detail = "No client selected" });
        return model;
    }
}

public sealed class StringNotEmptyToVisibilityConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        return string.IsNullOrWhiteSpace(value?.ToString())
            ? Visibility.Collapsed
            : Visibility.Visible;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        throw new NotSupportedException();
    }
}

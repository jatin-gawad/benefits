using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Threading;
using ClosedXML.Excel;
using Microsoft.Win32;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Repositories;
using RenewalReviewBuilder.App.Services;
using RenewalReviewBuilder.App.ViewModels;
using RenewalReviewBuilder.App.Views;
using Forms = System.Windows.Forms;

namespace RenewalReviewBuilder.App;

public partial class MainWindow : Window
{
    private enum ResponsiveDisplayPreference
    {
        Auto,
        Comfortable,
        Compact
    }

    private enum HomeDateRangeBasis
    {
        EffectiveDate,
        WorkingDate
    }

    private sealed class ProspectPresentationEvidence
    {
        public string ClientName { get; init; } = "";
        public bool HasProposalPresentation { get; set; }
        public bool HasRenewalPresentation { get; set; }
        public DateTime LastUpdated { get; set; } = DateTime.MinValue;
    }

    private sealed record HomeDateRangeSelection(
        DateOnly? Start,
        DateOnly? End,
        string Label,
        HomeDateRangeBasis Basis);

    private sealed record ClientWorkspaceChartMarker(
        double Left,
        double Top,
        string ToolTip);

    private sealed record AssurePlusHistoryReviewRow(
        int Year,
        string Health,
        string Dental,
        string Validation);

    private sealed record AssurePlusPricingReviewRow(
        string EffectiveDate,
        string Benefit,
        string PreviousYearRating,
        string RecentYearRating,
        string StandardChange,
        string ExperienceAdjustment,
        string TotalChange,
        string Source);

    private sealed record AssurePlusDistributionReviewRow(
        int RenewalYear,
        string Benefit,
        string Category,
        string YourClaims,
        string BenchmarkClaims,
        string Source);

    private static readonly string WestlandOneDriveRoot = ResolveWestlandOneDriveRoot();
    private static readonly string DefaultPdfPath = "";
    private static readonly string DefaultWorkbookPath = "";
    private static readonly string DefaultTemplatePath = ResolveWestlandPath("RenewalAutomation", "templates", "renewal_template.pptx");
    private static readonly string DefaultClaimsPowerPointTemplatePath = ResolveWestlandPath("RenewalAutomation", "templates", "Experience ReportV1.pptx");
    private static readonly string WestlandSharePointDocumentsUrl = Environment.GetEnvironmentVariable("WESTLAND_SHAREPOINT_DOCUMENTS_URL") ?? "";
    private const double DefaultMinWindowWidth = 720;
    private const double DefaultMinWindowHeight = 480;
    private const double NarrowLayoutBreakpoint = 1050;
    private const double CompactLayoutBreakpoint = 1450;
    private const string HeaderSearchPlaceholder = "Search clients by name, ID, or carrier...";
    private static readonly TimeSpan ClaimsRefreshTimeout = TimeSpan.FromMinutes(2);
    private static readonly string[] ClientIndustryOptions =
    [
        "Construction & Trades",
        "Manufacturing",
        "Retail / Wholesale",
        "Transportation / Warehousing",
        "Professional / Technology",
        "Financial / Insurance",
        "Health Care / Social Assistance",
        "Education / Non-profit",
        "Hospitality / Food Services",
        "Other Services",
        "Government / Public Sector",
        "Other / Review Required"
    ];
    private readonly ReviewPackBuilder _builder;
    private readonly ClaimsRangeReviewViewModel _claimsRangeReviewViewModel = new();
    private readonly ClaimsExperienceDataModelBuilder _claimsDataModelBuilder = new();
    private readonly ClaimsExperiencePowerPointExporter _claimsPowerPointExporter = new();
    private readonly BookletExtractorRunner _bookletExtractorRunner = new();
    private readonly DatabaseService _databaseService = new();
    private readonly ClientClaimsDashboardDataService _clientClaimsDashboardDataService;
    private readonly BenefitsMasterDatabaseExportService _benefitsMasterDatabaseExportService = new();
    private readonly ClientWorkspaceRenewalHistoryPresentationService _clientWorkspaceRenewalHistoryPresentationService = new();
    private readonly IFlexibleBenefitRepository _flexibleBenefitRepository = new FlexibleBenefitRepository();
    private readonly FinancialArrangementService _financialArrangementService = new(AppPaths.BenefitsIntelligenceDatabasePath);
    private readonly BatchDocumentProcessorService _batchDocumentProcessorService;
    private readonly PlanDesignIntelligenceService _planDesignIntelligenceService;
    private readonly ProposalPipelineOverviewService _proposalPipelineOverviewService;
    private readonly ProposalEngineRunner _proposalEngineRunner = new();
    private readonly ManulifeBillingImportCoordinator _manulifeBillingImportCoordinator;
    private readonly NetGrossRateScheduleImporter _netGrossRateScheduleImporter = new();
    private readonly RevisedRateScheduleImporter _revisedRateScheduleImporter = new();
    private readonly UserSettingsStore _settingsStore = new();
    private readonly SharePointStagedExportPublisher _sharePointExportPublisher = new();
    private readonly ObservableCollection<string> _additionalPdfPaths = [];
    private readonly ObservableCollection<string> _additionalWorkbookPaths = [];
    private readonly ObservableCollection<string> _bookletInputPaths = [];
    private HcsaCalculation _hcsaCalculation = new();
    private AllstateCriticalIllnessEnrollment _allstateCriticalIllnessEnrollment = new();
    private readonly ObservableCollection<BatchClientDiscoveryResult> _batchDiscoveryRows = [];
    private readonly ObservableCollection<RenewalQueueItem> _renewalQueue = [];
    private readonly ObservableCollection<ClientListItem> _clientList = [];
    private bool _headerSearchShowsPlaceholder = true;
    private readonly ObservableCollection<UnderwritingBenefitRow> _underwritingBenefitRows = [];
    private readonly ObservableCollection<CarrierAssumptionLibraryRow> _carrierAssumptionRows = [];
    private readonly ObservableCollection<CarrierBlockProfileRow> _carrierBlockProfiles = [];
    private readonly ObservableCollection<CarrierAssumptionSummaryRow> _carrierAssumptionSummaries = [];
    private readonly ObservableCollection<PlanDesignIntelligenceClientOption> _planDesignIntelligenceClients = [];
    private readonly ObservableCollection<PlanDesignSummaryRow> _planDesignIntelligenceSummaryRows = [];
    private readonly ObservableCollection<PlanDesignReadinessRow> _planDesignIntelligenceReadinessRows = [];
    private readonly ObservableCollection<PlanDesignIntelligenceFinding> _planDesignIntelligenceMissingRows = [];
    private readonly ObservableCollection<PlanDesignIntelligenceFinding> _planDesignIntelligenceNormalizationRows = [];
    private readonly ObservableCollection<PlanDesignIntelligenceFinding> _planDesignIntelligencePricingRows = [];
    private readonly ObservableCollection<PlanDesignIntelligenceFinding> _planDesignIntelligenceFindingRows = [];
    private readonly ObservableCollection<PlanDesignSourceTableUsage> _planDesignIntelligenceSourceRows = [];
    private readonly ObservableCollection<PlanDesignFactorRow> _planDesignFactorRows = [];
    private readonly ObservableCollection<PlanDesignFactorRow> _planDesignFactorOptions = [];
    private readonly ObservableCollection<ExecutiveMetricCard> _reportsKpiCards = [];
    private readonly ObservableCollection<ExecutivePipelineRow> _reportsRenewalPipelineRows = [];
    private readonly ObservableCollection<ExecutivePipelineRow> _reportsProposalPipelineRows = [];
    private readonly ObservableCollection<ProposalPipelineStageRow> _proposalPipelineOverviewRows = [];
    private readonly ObservableCollection<ExecutivePipelineRow> _reportsQuotePipelineRows = [];
    private readonly ObservableCollection<ExecutiveChartRow> _reportsRenewalsByMonthRows = [];
    private readonly ObservableCollection<ExecutiveChartRow> _reportsRenewalsByCarrierRows = [];
    private readonly ObservableCollection<ExecutiveChartRow> _reportsRenewalsByConsultantRows = [];
    private readonly ObservableCollection<ExecutiveSummaryRow> _reportsRenewalAnalyticsRows = [];
    private readonly ObservableCollection<ExecutiveSummaryRow> _reportsClaimsAnalyticsRows = [];
    private readonly ObservableCollection<ExecutiveSummaryRow> _reportsUnderwritingAnalyticsRows = [];
    private readonly ObservableCollection<ExecutiveSummaryRow> _reportsCarrierAnalyticsRows = [];
    private readonly ObservableCollection<ExecutiveSummaryRow> _reportsProposalAnalyticsRows = [];
    private readonly ObservableCollection<ExecutiveSummaryRow> _reportsPlanDesignAnalyticsRows = [];
    private readonly ObservableCollection<ExecutiveSummaryRow> _reportsBenchmarkRows = [];
    private readonly ObservableCollection<ExecutiveSummaryRow> _reportsAlertRows = [];
    private readonly ObservableCollection<ExecutiveSummaryRow> _reportsExecutiveDecisionRows = [];
    private readonly ObservableCollection<ExecutiveSummaryRow> _reportsDrillDownRows = [];
    private readonly ObservableCollection<ExecutiveDataCategoryRecord> _reportsDataCategoryRows = [];
    private readonly ObservableCollection<ExecutiveSummaryRow> _reportsClientSummaryRows = [];
    private readonly ObservableCollection<ExecutiveSummaryRow> _reportsClientHealthComponentRows = [];
    private readonly ObservableCollection<ExecutiveSummaryRow> _reportsClientAnalyticsRows = [];
    private readonly ObservableCollection<ExecutiveSummaryRow> _reportsClientHistoryRows = [];
    private readonly ObservableCollection<string> _reportsInsightRows = [];
    private readonly ObservableCollection<string> _reportsClientActionRows = [];
    private readonly ObservableCollection<ExecutiveSummaryRow> _bookOfBusinessActionRows = [];
    private readonly ObservableCollection<ExecutiveSummaryRow> _bookOfBusinessRenewalCalendarRows = [];
    private readonly ObservableCollection<ExecutiveSummaryRow> _bookOfBusinessPipelineRows = [];
    private readonly ObservableCollection<ExecutiveSummaryRow> _bookOfBusinessCarrierRows = [];
    private readonly ObservableCollection<ExecutiveSummaryRow> _bookOfBusinessIndustryRows = [];
    private readonly ObservableCollection<BookOfBusinessClientRow> _bookOfBusinessClientRows = [];
    private readonly ObservableCollection<BillingReconciliationResult> _billingReconciliationRows = [];
    private readonly ObservableCollection<NegotiationRateRow> _negotiationRates = [];
    private readonly ObservableCollection<NegotiationRateRow> _netGrossRateScheduleReviewRows = [];
    private readonly ObservableCollection<NegotiationClaimsRow> _negotiationClaimsRows = [];
    private readonly ObservableCollection<NegotiationModelReferenceRow> _modelReferenceRows = [];
    private readonly UnderwritingEngine _underwritingEngine = new();
    private readonly NegotiationWorkbenchService _negotiationWorkbenchService = new();
    private ICollectionView? _renewalQueueView;
    private ICollectionView? _clientListView;
    private RenewalQueueItem? _selectedNegotiationRenewal;
    private RenewalQueueItem? _selectedUnderwritingRenewal;
    private ClientWorkspaceViewModel? _selectedClientWorkspace;
    private ClientRenewalSelectionIdentity? _clientRenewalSelectionIdentity;
    private long? _latestSavedClientId;
    private long? _latestSavedDocumentId;
    private long? _analysisClientId;
    private long? _analysisSourceDocumentId;
    private AnalysisResult? _analysis;
    private ExportResult? _export;
    private ClaimsExperienceDataModelResult? _claimsDataModelExport;
    private ClaimsExperiencePowerPointResult? _claimsPowerPointExport;
    private BookletExtractionResult? _bookletExtractionResult;
    private BatchDiscoveryResult? _batchDiscoveryResult;
    private BatchProcessingResult? _batchProcessingResult;
    private string? _batchOutputFolder;
    private List<string> _bookletExtractionInputSnapshot = [];
    private CancellationTokenSource? _bookletExtractionCancellation;
    private string? _proposalOutputFolder;
    private string? _billingOutputFolder;
    private string? _renewalQueueExportFolder;
    private long _editingPlanDesignFactorId;
    private bool _isLoadingDefaults;
    private bool _isUpdatingUnderwritingSelection;
    private bool _isUpdatingHomeDashboardFilters;
    private bool _homeDashboardFiltersInitialized;
    private bool _isLoadingPlanDesignIntelligence;
    private bool _isRefreshingClientList;
    private bool _isLoadingReportsAnalytics;
    private bool _isUpdatingReportsFilters;
    private bool _isReportsClientMode;
    private bool _isApplyingResponsiveLayout;
    private ResponsiveDisplayPreference _responsiveDisplayPreference = ResponsiveDisplayPreference.Auto;
    private ExecutiveReportingSnapshot? _reportsSnapshot;

    public MainWindow()
    {
        _builder = new ReviewPackBuilder(_databaseService);
        _clientClaimsDashboardDataService = new ClientClaimsDashboardDataService(_databaseService);
        InitializeComponent();
        ClaimsRangeReviewHost.DataContext = _claimsRangeReviewViewModel;
        _manulifeBillingImportCoordinator = new ManulifeBillingImportCoordinator(_databaseService);
        _batchDocumentProcessorService = new BatchDocumentProcessorService(databaseService: _databaseService);
        _planDesignIntelligenceService = new PlanDesignIntelligenceService(_databaseService.DatabasePath);
        _proposalPipelineOverviewService = new ProposalPipelineOverviewService(_databaseService.DatabasePath);
        SourceInitialized += MainWindow_SourceInitialized;
        DpiChanged += MainWindow_DpiChanged;
        SizeChanged += MainWindow_SizeChanged;
        Dispatcher.BeginInvoke(new Action(ApplyResponsiveLayout), DispatcherPriority.Loaded);
        AdditionalPdfsListBox.ItemsSource = _additionalPdfPaths;
        AdditionalWorkbooksListBox.ItemsSource = _additionalWorkbookPaths;
        BenefitBookletsListBox.ItemsSource = _bookletInputPaths;
        BatchDiscoveryGrid.ItemsSource = _batchDiscoveryRows;
        BillingReconciliationGrid.ItemsSource = _billingReconciliationRows;
        NetGrossRateScheduleReviewGrid.ItemsSource = _netGrossRateScheduleReviewRows;
        ResetBatchProgress();
        UpdateBatchButtons(isBusy: false);
        UnderwritingClientComboBox.ItemsSource = _renewalQueue;
        UnderwritingBenefitGrid.ItemsSource = _underwritingBenefitRows;
        AssumptionLibraryGrid.ItemsSource = _carrierAssumptionRows;
        CarrierBlockProfileGrid.ItemsSource = _carrierBlockProfiles;
        CarrierAssumptionSummaryGrid.ItemsSource = _carrierAssumptionSummaries;
        PlanDesignIntelligenceClientComboBox.ItemsSource = _planDesignIntelligenceClients;
        PlanDesignIntelligenceSummaryGrid.ItemsSource = _planDesignIntelligenceSummaryRows;
        PlanDesignIntelligenceReadinessGrid.ItemsSource = _planDesignIntelligenceReadinessRows;
        PlanDesignIntelligenceMissingGrid.ItemsSource = _planDesignIntelligenceMissingRows;
        PlanDesignIntelligenceNormalizationGrid.ItemsSource = _planDesignIntelligenceNormalizationRows;
        PlanDesignIntelligencePricingGrid.ItemsSource = _planDesignIntelligencePricingRows;
        PlanDesignIntelligenceSourcesGrid.ItemsSource = _planDesignIntelligenceSourceRows;
        PlanDesignIntelligenceFindingsGrid.ItemsSource = _planDesignIntelligenceFindingRows;
        PlanDesignFactorsGrid.ItemsSource = _planDesignFactorRows;
        CalculatorSelectedFactorComboBox.ItemsSource = _planDesignFactorOptions;
        ReportsKpiItemsControl.ItemsSource = _reportsKpiCards;
        ReportsRenewalPipelineGrid.ItemsSource = _reportsRenewalPipelineRows;
        ReportsProposalPipelineGrid.ItemsSource = _reportsProposalPipelineRows;
        ReportsQuotePipelineGrid.ItemsSource = _reportsQuotePipelineRows;
        ProposalPipelineOverviewGrid.ItemsSource = _proposalPipelineOverviewRows;
        ReportsRenewalsByMonthItemsControl.ItemsSource = _reportsRenewalsByMonthRows;
        ReportsRenewalsByCarrierItemsControl.ItemsSource = _reportsRenewalsByCarrierRows;
        ReportsRenewalsByConsultantItemsControl.ItemsSource = _reportsRenewalsByConsultantRows;
        ReportsRenewalAnalyticsGrid.ItemsSource = _reportsRenewalAnalyticsRows;
        ReportsClaimsAnalyticsGrid.ItemsSource = _reportsClaimsAnalyticsRows;
        ReportsUnderwritingAnalyticsGrid.ItemsSource = _reportsUnderwritingAnalyticsRows;
        ReportsCarrierAnalyticsGrid.ItemsSource = _reportsCarrierAnalyticsRows;
        ReportsProposalAnalyticsGrid.ItemsSource = _reportsProposalAnalyticsRows;
        ReportsPlanDesignAnalyticsGrid.ItemsSource = _reportsPlanDesignAnalyticsRows;
        ReportsBenchmarkGrid.ItemsSource = _reportsBenchmarkRows;
        ReportsAlertGrid.ItemsSource = _reportsAlertRows;
        ReportsExecutiveDecisionsItemsControl.ItemsSource = _reportsExecutiveDecisionRows;
        ReportsInsightsItemsControl.ItemsSource = _reportsInsightRows;
        ReportsDataCategoryGrid.ItemsSource = _reportsDataCategoryRows;
        ReportsDrillDownGrid.ItemsSource = _reportsDrillDownRows;
        ReportsClientSummaryItemsControl.ItemsSource = _reportsClientSummaryRows;
        ReportsClientHealthComponentsGrid.ItemsSource = _reportsClientHealthComponentRows;
        ReportsClientAnalyticsGrid.ItemsSource = _reportsClientAnalyticsRows;
        ReportsClientHistoryGrid.ItemsSource = _reportsClientHistoryRows;
        ReportsClientActionItemsControl.ItemsSource = _reportsClientActionRows;
        BookBusinessActionItemsControl.ItemsSource = _bookOfBusinessActionRows;
        BookBusinessRenewalCalendarItemsControl.ItemsSource = _bookOfBusinessRenewalCalendarRows;
        BookBusinessPipelineItemsControl.ItemsSource = _bookOfBusinessPipelineRows;
        BookBusinessCarrierItemsControl.ItemsSource = _bookOfBusinessCarrierRows;
        BookBusinessIndustryItemsControl.ItemsSource = _bookOfBusinessIndustryRows;
        BookBusinessActiveClientsGrid.ItemsSource = _bookOfBusinessClientRows;
        NegotiationRatesGrid.ItemsSource = _negotiationRates;
        NegotiationClaimsGrid.ItemsSource = _negotiationClaimsRows;
        ModelReferenceGrid.ItemsSource = _modelReferenceRows;
        NegotiationDateTextBox.Text = DateTime.Today.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        SeedRenewalQueue();
        ConfigureRenewalQueue();
        SeedClientList();
        ConfigureClientList();
        AppPaths.EnsureCreated();
        AppLogger.Info("Benefits Intelligence OS started.");
        var databaseStatus = _databaseService.Initialize();
        if (!databaseStatus.Succeeded)
        {
            SetStatus(databaseStatus.Message);
        }
        else
        {
            Loaded += MainWindow_Loaded;
        }

        WireInputChangeHandlers();
        ConfigurePlanDesignFactorInputs();
        _isLoadingDefaults = true;
        PopulateDefaults();
        _isLoadingDefaults = false;
        SetNavigationWorkspace("Home", "Home");
    }

    private void MainWindow_SourceInitialized(object? sender, EventArgs e)
    {
        FitWindowToWorkArea();
    }

    private void MainWindow_DpiChanged(object sender, DpiChangedEventArgs e)
    {
        Dispatcher.BeginInvoke(new Action(() =>
        {
            FitWindowToWorkArea();
            ApplyResponsiveLayout();
        }), DispatcherPriority.Loaded);
    }

    private void FitWindowToWorkArea()
    {
        var handle = new WindowInteropHelper(this).Handle;
        var screen = handle == IntPtr.Zero
            ? Forms.Screen.PrimaryScreen
            : Forms.Screen.FromHandle(handle);
        if (screen is null)
        {
            return;
        }

        var dpi = VisualTreeHelper.GetDpi(this);
        var dpiScaleX = dpi.DpiScaleX > 0 ? dpi.DpiScaleX : 1;
        var dpiScaleY = dpi.DpiScaleY > 0 ? dpi.DpiScaleY : 1;
        var availableWidth = Math.Max(1, screen.WorkingArea.Width / dpiScaleX);
        var availableHeight = Math.Max(1, screen.WorkingArea.Height / dpiScaleY);

        MinWidth = Math.Min(DefaultMinWindowWidth, availableWidth);
        MinHeight = Math.Min(DefaultMinWindowHeight, availableHeight);

        var restoreWidth = Math.Max(MinWidth, availableWidth * 0.94);
        var restoreHeight = Math.Max(MinHeight, availableHeight * 0.94);

        if (WindowState == WindowState.Maximized)
        {
            Width = restoreWidth;
            Height = restoreHeight;
        }
        else
        {
            Width = Math.Min(Width, restoreWidth);
            Height = Math.Min(Height, restoreHeight);
        }
    }

    private void MainWindow_SizeChanged(object sender, SizeChangedEventArgs e)
    {
        ApplyResponsiveLayout();
    }

    private void ResponsiveModeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (ResponsiveModeComboBox.SelectedItem is not ComboBoxItem item)
        {
            return;
        }

        _responsiveDisplayPreference = item.Content?.ToString() switch
        {
            "Comfortable" => ResponsiveDisplayPreference.Comfortable,
            "Compact" => ResponsiveDisplayPreference.Compact,
            _ => ResponsiveDisplayPreference.Auto
        };
        ApplyResponsiveLayout();
    }

    private void ApplyResponsiveLayout()
    {
        if (_isApplyingResponsiveLayout)
        {
            return;
        }

        if (SidebarColumn is null ||
            SidebarGrid is null ||
            SidebarLogoHost is null ||
            SidebarLogoImage is null ||
            MainHeaderBorder is null ||
            HeaderSearchColumn is null ||
            HeaderTitleTextBlock is null ||
            HeaderSubtitleTextBlock is null ||
            HeaderSearchBox is null ||
            FooterActionBar is null ||
            ExportClaimsDataModelButton is null ||
            ExportClaimsPowerPointButton is null ||
            ExtractBookletToMasterDbButton is null ||
            FooterMoreActionsButton is null)
        {
            return;
        }

        _isApplyingResponsiveLayout = true;
        try
        {
            var windowWidth = ActualWidth > 0 ? ActualWidth : Width;
            var isNarrow = windowWidth < NarrowLayoutBreakpoint;
            var isCompact = _responsiveDisplayPreference switch
            {
                ResponsiveDisplayPreference.Compact => true,
                ResponsiveDisplayPreference.Comfortable => isNarrow,
                _ => windowWidth < CompactLayoutBreakpoint
            };

            SidebarColumn.Width = new GridLength(isNarrow ? 180 : isCompact ? 212 : 232);
            SidebarGrid.Margin = isNarrow
                ? new Thickness(6)
                : isCompact
                    ? new Thickness(8)
                    : new Thickness(10);
            SidebarLogoHost.Height = isNarrow ? 58 : isCompact ? 68 : 78;
            SidebarLogoImage.MaxWidth = isNarrow ? 100 : isCompact ? 118 : 132;

            MainHeaderBorder.Padding = isNarrow
                ? new Thickness(12, 9, 12, 9)
                : isCompact
                    ? new Thickness(18, 10, 18, 10)
                    : new Thickness(28, 14, 28, 14);
            HeaderSearchColumn.Width = new GridLength(isNarrow ? 0 : isCompact ? 280 : 430);
            HeaderSearchBox.Visibility = isNarrow ? Visibility.Collapsed : Visibility.Visible;
            HeaderHelpButton.Visibility = isNarrow ? Visibility.Collapsed : Visibility.Visible;
            HeaderAlertButton.Visibility = isNarrow ? Visibility.Collapsed : Visibility.Visible;
            HeaderTitleTextBlock.FontSize = isNarrow ? 19 : isCompact ? 22 : 26;
            HeaderSubtitleTextBlock.FontSize = isNarrow ? 10.5 : isCompact ? 11.5 : 12;
            HeaderSearchBox.Margin = isCompact ? new Thickness(10, 0, 10, 0) : new Thickness(18, 0, 18, 0);

            if (HomeDashboardContentStackPanel is not null &&
                HomeDashboardWelcomeCard is not null &&
                HomeDashboardWelcomeGrid is not null &&
                HomeDashboardWelcomeTextPanel is not null &&
                HomeDashboardFiltersWrapPanel is not null)
            {
                HomeDashboardContentStackPanel.Margin = isNarrow
                    ? new Thickness(12, 10, 12, 14)
                    : isCompact
                        ? new Thickness(18, 14, 18, 16)
                        : new Thickness(24, 18, 24, 16);
                HomeDashboardWelcomeCard.Padding = isNarrow
                    ? new Thickness(14, 12, 14, 12)
                    : isCompact
                        ? new Thickness(18, 16, 18, 16)
                        : new Thickness(20, 18, 20, 18);

                if (isCompact)
                {
                    Grid.SetColumn(HomeDashboardFiltersWrapPanel, 0);
                    Grid.SetRow(HomeDashboardFiltersWrapPanel, 1);
                    HomeDashboardFiltersWrapPanel.HorizontalAlignment = HorizontalAlignment.Left;
                    HomeDashboardFiltersWrapPanel.Margin = new Thickness(0, 16, 0, 0);
                    HomeDashboardWelcomeTextPanel.Margin = new Thickness(0);

                    if (HomeDashboardWelcomeGrid.ColumnDefinitions.Count > 1)
                    {
                        HomeDashboardWelcomeGrid.ColumnDefinitions[1].Width = new GridLength(0);
                    }
                }
                else
                {
                    Grid.SetColumn(HomeDashboardFiltersWrapPanel, 1);
                    Grid.SetRow(HomeDashboardFiltersWrapPanel, 0);
                    HomeDashboardFiltersWrapPanel.HorizontalAlignment = HorizontalAlignment.Right;
                    HomeDashboardFiltersWrapPanel.Margin = new Thickness(18, 0, 0, 0);
                    HomeDashboardWelcomeTextPanel.Margin = new Thickness(0);

                    if (HomeDashboardWelcomeGrid.ColumnDefinitions.Count > 1)
                    {
                        HomeDashboardWelcomeGrid.ColumnDefinitions[1].Width = GridLength.Auto;
                    }
                }

                ApplyHomeKpiCompactSizing(isCompact);
                HomeKpiUniformGrid.Columns = isNarrow ? 2 : isCompact ? 3 : 6;
                HomeKpiUniformGrid.Rows = 0;
            }

            if (ClientsContentGrid is not null && ClientKpiUniformGrid is not null)
            {
                ClientsContentGrid.Margin = isNarrow
                    ? new Thickness(12, 10, 12, 14)
                    : isCompact
                        ? new Thickness(18, 14, 18, 16)
                        : new Thickness(24, 18, 24, 16);
                ClientKpiUniformGrid.Columns = isNarrow ? 2 : isCompact ? 3 : 5;
                ClientKpiUniformGrid.Rows = 0;
            }

            FooterActionBar.Padding = isCompact ? new Thickness(14, 9, 14, 9) : new Thickness(18, 12, 18, 12);
            ExportClaimsDataModelButton.Visibility = isCompact ? Visibility.Collapsed : Visibility.Visible;
            ExportClaimsPowerPointButton.Visibility = isCompact ? Visibility.Collapsed : Visibility.Visible;
            ExtractBookletToMasterDbButton.Visibility = isCompact ? Visibility.Collapsed : Visibility.Visible;
            FooterMoreActionsButton.Visibility = isCompact ? Visibility.Visible : Visibility.Collapsed;

            if (RenewalContentGrid is not null &&
                RenewalQueueCard is not null &&
                RenewalDetailPanel is not null &&
                RenewalDetailColumn is not null)
            {
                if (isCompact)
                {
                    RenewalDetailColumn.Width = new GridLength(0);
                    RenewalQueueCard.Margin = new Thickness(0, 0, 0, 14);
                    RenewalDetailPanel.Margin = new Thickness(0);
                    Grid.SetColumn(RenewalQueueCard, 0);
                    Grid.SetRow(RenewalQueueCard, 0);
                    Grid.SetColumn(RenewalDetailPanel, 0);
                    Grid.SetRow(RenewalDetailPanel, 1);
                }
                else
                {
                    RenewalDetailColumn.Width = new GridLength(365);
                    RenewalQueueCard.Margin = new Thickness(0, 0, 14, 0);
                    RenewalDetailPanel.Margin = new Thickness(0);
                    Grid.SetColumn(RenewalQueueCard, 0);
                    Grid.SetRow(RenewalQueueCard, 0);
                    Grid.SetColumn(RenewalDetailPanel, 1);
                    Grid.SetRow(RenewalDetailPanel, 0);
                }
            }
        }
        finally
        {
            _isApplyingResponsiveLayout = false;
        }
    }

    private void ApplyHomeKpiCompactSizing(bool isCompact)
    {
        if (HomeKpiUniformGrid is null)
        {
            return;
        }

        var titleStyle = TryFindResource("HomeKpiTitle") as Style;
        var valueStyle = TryFindResource("HomeKpiValue") as Style;
        var miniLabelStyle = TryFindResource("MiniLabel") as Style;
        var iconBadgeStyle = TryFindResource("KpiIconBadgeStyle") as Style;

        foreach (var textBlock in FindVisualChildren<TextBlock>(HomeKpiUniformGrid))
        {
            if (ReferenceEquals(textBlock.Style, titleStyle))
            {
                textBlock.FontSize = isCompact ? 10.5 : 12.5;
                textBlock.LineHeight = isCompact ? 12.5 : double.NaN;
            }
            else if (ReferenceEquals(textBlock.Style, valueStyle))
            {
                textBlock.FontSize = isCompact ? 23 : 27;
                textBlock.Margin = isCompact ? new Thickness(0, 6, 0, 0) : new Thickness(0, 10, 0, 1);
            }
            else if (ReferenceEquals(textBlock.Style, miniLabelStyle))
            {
                textBlock.FontSize = isCompact ? 9.5 : 11;
                textBlock.LineHeight = isCompact ? 11.5 : double.NaN;
            }
            else if (textBlock.FontFamily.Source == "Segoe MDL2 Assets")
            {
                textBlock.FontSize = isCompact ? 14 : 17;
            }
        }

        foreach (var border in FindVisualChildren<Border>(HomeKpiUniformGrid))
        {
            if (ReferenceEquals(border.Style, iconBadgeStyle))
            {
                border.Width = isCompact ? 30 : 38;
                border.Height = isCompact ? 30 : 38;
                border.CornerRadius = new CornerRadius(isCompact ? 15 : 19);
            }
            else if (border.Child is Grid)
            {
                border.Padding = isCompact ? new Thickness(12, 12, 10, 12) : new Thickness(18, 16, 18, 16);
                border.MinHeight = isCompact ? 104 : 116;
            }
        }

        foreach (var stackPanel in FindVisualChildren<StackPanel>(HomeKpiUniformGrid))
        {
            if (stackPanel.Parent is Grid)
            {
                stackPanel.Margin = isCompact ? new Thickness(8, 0, 0, 0) : new Thickness(12, 0, 0, 0);
            }
        }
    }

    private static IEnumerable<T> FindVisualChildren<T>(DependencyObject parent) where T : DependencyObject
    {
        if (parent is null)
        {
            yield break;
        }

        for (var i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
        {
            var child = VisualTreeHelper.GetChild(parent, i);
            if (child is T typedChild)
            {
                yield return typedChild;
            }

            foreach (var descendant in FindVisualChildren<T>(child))
            {
                yield return descendant;
            }
        }
    }

    private void FooterMoreActions_Click(object sender, RoutedEventArgs e)
    {
        if (FooterMoreActionsButton.ContextMenu is null)
        {
            return;
        }

        FooterMoreActionsButton.ContextMenu.PlacementTarget = FooterMoreActionsButton;
        FooterMoreActionsButton.ContextMenu.IsOpen = true;
    }

    private void SetOperationalFooterVisible(bool isVisible)
    {
        if (FooterActionBar is not null)
        {
            FooterActionBar.Visibility = isVisible ? Visibility.Visible : Visibility.Collapsed;
        }
    }

    private void SetNavigationWorkspace(string workspaceKey, string title, string? clientName = null)
    {
        if (ContextMainPanel is null)
        {
            return;
        }

        CollapseContextPanels();
        ContextMainPanel.Visibility = Visibility.Visible;

        ContextEyebrowTextBlock.Text = workspaceKey == "ClientWorkspace"
            ? "CLIENT WORKSPACE"
            : FriendlyWorkspaceName(workspaceKey).ToUpperInvariant();
        ContextTitleTextBlock.Text = string.IsNullOrWhiteSpace(clientName) ? title : clientName;
        ContextBackToMainButton.Visibility = workspaceKey == "Main" ? Visibility.Collapsed : Visibility.Visible;
        ContextBackToClientsButton.Visibility = workspaceKey == "ClientWorkspace" ? Visibility.Visible : Visibility.Collapsed;
        UpdateSidebarActive(workspaceKey, title);
        UpdateClientWorkspaceActiveTab(workspaceKey, title);

        var isBookOfBusinessClaims = workspaceKey == "BookOfBusiness"
            && title.Contains("Claims", StringComparison.OrdinalIgnoreCase);
        if (ClaimsExperienceDashboardScrollViewer is not null && !isBookOfBusinessClaims)
        {
            ClaimsExperienceDashboardScrollViewer.Visibility = Visibility.Collapsed;
        }

        if (ProposalPipelineOverviewScrollViewer is not null)
        {
            ProposalPipelineOverviewScrollViewer.Visibility = Visibility.Collapsed;
        }
    }

    private void UpdateClientWorkspaceActiveTab(string workspaceKey, string title)
    {
        var tabs = new (string Title, System.Windows.Controls.Button? Button)[]
        {
            ("Overview", ClientWorkspaceOverviewTabButton),
            ("Renewal", ClientWorkspaceRenewalTabButton),
            ("Plan Design", ClientWorkspacePlanDesignTabButton),
            ("Claims", ClientWorkspaceClaimsTabButton),
            ("ASO", ClientWorkspaceAsoTabButton),
            ("Billing", ClientWorkspaceBillingTabButton),
            ("Financial Intelligence", ClientWorkspaceFinancialTabButton),
            ("Proposals", ClientWorkspaceProposalsTabButton),
            ("Reports", ClientWorkspaceReportsTabButton),
            ("Documents", ClientWorkspaceDocumentsTabButton),
            ("History", ClientWorkspaceHistoryTabButton)
        };

        foreach (var (tabTitle, button) in tabs)
        {
            if (button is null)
            {
                continue;
            }

            var isActive = workspaceKey == "ClientWorkspace"
                && string.Equals(title, tabTitle, StringComparison.OrdinalIgnoreCase);
            button.Foreground = (Brush)FindResource(isActive ? "BrightBlueBrush" : "TextNavyBrush");
            button.BorderBrush = isActive
                ? (Brush)FindResource("BrightBlueBrush")
                : Brushes.Transparent;
            button.BorderThickness = isActive ? new Thickness(0, 0, 0, 2) : new Thickness(0);
            button.FontWeight = isActive ? FontWeights.Bold : FontWeights.Normal;
        }
    }

    private void UpdateSidebarActive(string workspaceKey, string title)
    {
        var normalStyle = (Style)FindResource("SidebarNavButton");
        var activeStyle = (Style)FindResource("SidebarActiveButton");
        var buttons = new[]
        {
            NavHomeButton,
            NavBookOfBusinessButton,
            NavClientsButton,
            NavRenewalsButton,
            NavProposalsButton,
            NavPlanDesignButton,
            NavClaimsButton,
            NavBenchmarkingButton,
            NavDocumentOperationsButton,
            NavBatchProcessingButton,
            NavDataExtractionButton,
            NavReportsLibraryButton,
            NavExecutiveDashboardsButton,
            NavCustomReportsButton,
            NavDataQualityButton,
            NavUsersTeamsButton,
            NavSettingsButton
        };

        foreach (var button in buttons)
        {
            button.Style = normalStyle;
        }

        var activeButton = PickActiveSidebarButton(workspaceKey, title);
        activeButton.Style = activeStyle;
    }

    private Button PickActiveSidebarButton(string workspaceKey, string title)
    {
        var normalizedTitle = title.Trim();
        if (workspaceKey == "ClientWorkspace")
        {
            if (normalizedTitle.Contains("Renewal", StringComparison.OrdinalIgnoreCase))
            {
                return NavRenewalsButton;
            }

            if (normalizedTitle.Contains("Proposal", StringComparison.OrdinalIgnoreCase))
            {
                return NavProposalsButton;
            }

            if (normalizedTitle.Contains("Plan Design", StringComparison.OrdinalIgnoreCase))
            {
                return NavPlanDesignButton;
            }

            if (normalizedTitle.Contains("Claims", StringComparison.OrdinalIgnoreCase))
            {
                return NavClaimsButton;
            }

            return NavClientsButton;
        }

        if (workspaceKey == "BookOfBusiness")
        {
            if (normalizedTitle.Contains("Renewal", StringComparison.OrdinalIgnoreCase))
            {
                return NavRenewalsButton;
            }

            if (normalizedTitle.Contains("Claims", StringComparison.OrdinalIgnoreCase))
            {
                return NavClaimsButton;
            }

            return NavBookOfBusinessButton;
        }

        return workspaceKey switch
        {
            "Home" or "Main" => NavHomeButton,
            "Clients" => NavClientsButton,
            "ExecutiveCenter" => NavExecutiveDashboardsButton,
            "Benchmarking" => NavBenchmarkingButton,
            "DocumentOperations" => NavDocumentOperationsButton,
            "Reports" => NavReportsLibraryButton,
            "Administration" => normalizedTitle.Contains("Data", StringComparison.OrdinalIgnoreCase)
                ? NavDataQualityButton
                : NavSettingsButton,
            _ => NavHomeButton
        };
    }

    private void CollapseContextPanels()
    {
        ContextMainPanel.Visibility = Visibility.Collapsed;
        ContextHomePanel.Visibility = Visibility.Collapsed;
        ContextClientsPanel.Visibility = Visibility.Collapsed;
        ContextClientWorkspacePanel.Visibility = Visibility.Collapsed;
        ContextBookBusinessPanel.Visibility = Visibility.Collapsed;
        ContextExecutiveCenterPanel.Visibility = Visibility.Collapsed;
        ContextBenchmarkingPanel.Visibility = Visibility.Collapsed;
        ContextDocumentOperationsPanel.Visibility = Visibility.Collapsed;
        ContextReportsPanel.Visibility = Visibility.Collapsed;
        ContextAdministrationPanel.Visibility = Visibility.Collapsed;
    }

    private static string FriendlyWorkspaceName(string workspaceKey)
    {
        return workspaceKey switch
        {
            "Main" => "Main Navigation",
            "BookOfBusiness" => "Book of Business",
            "ExecutiveCenter" => "Executive Center",
            "DocumentOperations" => "Document Operations",
            "ClientWorkspace" => "Client Workspace",
            _ => workspaceKey
        };
    }

    private void PopulateDefaults()
    {
        var settings = _settingsStore.Load();

        PdfPathTextBox.Text = FirstExisting(settings.PdfPath, DefaultPdfPath);
        _additionalPdfPaths.Clear();
        foreach (var path in settings.AdditionalPdfPaths.Where(path => !string.IsNullOrWhiteSpace(path)))
        {
            _additionalPdfPaths.Add(path);
        }

        WorkbookPathTextBox.Text = FirstExisting(settings.WorkbookPath, DefaultWorkbookPath);
        _additionalWorkbookPaths.Clear();
        foreach (var path in settings.AdditionalWorkbookPaths.Where(path => !string.IsNullOrWhiteSpace(path)))
        {
            _additionalWorkbookPaths.Add(path);
        }
        _bookletInputPaths.Clear();
        foreach (var path in settings.BenefitBookletPaths.Where(path => !string.IsNullOrWhiteSpace(path)))
        {
            _bookletInputPaths.Add(path);
        }

        TemplatePathTextBox.Text = FirstExisting(settings.TemplatePath, DefaultTemplatePath);
        CurrentYearFolderTextBox.Text = Directory.Exists(settings.CurrentYearFolder)
            ? settings.CurrentYearFolder
            : "";
        PriorYearFolderTextBox.Text = Directory.Exists(settings.PriorYearFolder)
            ? settings.PriorYearFolder
            : "";
        AsoBillingFolderTextBox.Text = Directory.Exists(settings.AsoBillingFolder)
            ? settings.AsoBillingFolder
            : "";
        _hcsaCalculation = settings.HcsaCalculation?.Clone() ?? new HcsaCalculation();
        _hcsaCalculation.RemoveLegacyCriticalIllnessEntries();
        UpdateHcsaCalculationSummary();
        _allstateCriticalIllnessEnrollment = settings.AllstateCriticalIllnessEnrollment?.Clone() ?? new AllstateCriticalIllnessEnrollment();
        UpdateAllstateCriticalIllnessEnrollmentSummary();
        SetReportMonthCount(settings.AutoDetectReportMonthCount ? 0 : settings.ReportMonthCount);

        var documents = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        OutputFolderTextBox.Text = !string.IsNullOrWhiteSpace(settings.OutputFolder)
            ? settings.OutputFolder
            : Path.Combine(documents, "Renewal Review Builder", "Outputs");
    }

    private void ConfigurePlanDesignFactorInputs()
    {
        SetComboBoxValues(PlanFactorBenefitCategoryFilterComboBox, PlanDesignFactorReferenceData.BenefitCategories, includeBlank: true);
        SetComboBoxValues(PlanFactorChangeTypeFilterComboBox, PlanDesignFactorReferenceData.ChangeTypes, includeBlank: true);
        SetComboBoxValues(PlanFactorConfidenceFilterComboBox, PlanDesignFactorReferenceData.ConfidenceLevels, includeBlank: true);
        SetComboBoxValues(PlanFactorBenefitCategoryComboBox, PlanDesignFactorReferenceData.BenefitCategories, includeBlank: false);
        SetComboBoxValues(PlanFactorChangeTypeComboBox, PlanDesignFactorReferenceData.ChangeTypes, includeBlank: false);
        SetComboBoxValues(PlanFactorDirectionComboBox, PlanDesignFactorReferenceData.Directions, includeBlank: false);
        SetComboBoxValues(PlanFactorConfidenceComboBox, PlanDesignFactorReferenceData.ConfidenceLevels, includeBlank: false);
        SetComboBoxValues(CalculatorBenefitCategoryComboBox, PlanDesignFactorReferenceData.BenefitCategories, includeBlank: true);
        ResetPlanDesignFactorForm();
    }

    private static void SetComboBoxValues(ComboBox comboBox, IEnumerable<string> values, bool includeBlank)
    {
        comboBox.Items.Clear();
        if (includeBlank)
        {
            comboBox.Items.Add("");
        }

        foreach (var value in values.Where(value => !string.IsNullOrWhiteSpace(value)).Distinct(StringComparer.OrdinalIgnoreCase))
        {
            comboBox.Items.Add(value);
        }

        comboBox.SelectedIndex = comboBox.Items.Count > 0 ? 0 : -1;
    }

    private void SeedRenewalQueue()
    {
        _renewalQueue.Clear();
    }

    private void ConfigureRenewalQueue()
    {
        _renewalQueueView = CollectionViewSource.GetDefaultView(_renewalQueue);
        _renewalQueueView.Filter = RenewalQueueFilter;
        _renewalQueueView.SortDescriptions.Clear();
        _renewalQueueView.SortDescriptions.Add(new SortDescription(nameof(RenewalQueueItem.StatusSortRank), ListSortDirection.Ascending));
        _renewalQueueView.SortDescriptions.Add(new SortDescription(nameof(RenewalQueueItem.RenewalDateSort), ListSortDirection.Ascending));

        RenewalQueueGrid.ItemsSource = _renewalQueueView;
        ConfigureFilterCombo(RenewalStatusFilterComboBox, "All statuses", _renewalQueue.Select(row => row.Status).Distinct().OrderBy(StatusRank));
        ConfigureFilterCombo(RenewalCarrierFilterComboBox, "All carriers", _renewalQueue.Select(row => row.Carrier).Distinct().OrderBy(value => value));
        ConfigureFilterCombo(RenewalMonthFilterComboBox, "All months", _renewalQueue.Select(row => row.RenewalMonth).Distinct().OrderBy(MonthRank));

        RenewalQueueGrid.SelectedIndex = 0;
        UpdateRenewalQueueSummary();
        UpdateSelectedRenewalDetail(RenewalQueueGrid.SelectedItem as RenewalQueueItem ?? _renewalQueue.FirstOrDefault());
    }

    private void SeedClientList()
    {
        _clientList.Clear();
    }

    private void ConfigureClientList()
    {
        _clientListView = CollectionViewSource.GetDefaultView(_clientList);
        _clientListView.Filter = ClientListFilter;
        ClientListGrid.ItemsSource = _clientListView;
        ConfigureFilterCombo(ClientStatusFilterComboBox, "All Statuses", _clientList.Select(row => row.Status).Distinct().OrderBy(value => value));
        ConfigureIndustryControls();
        UpdateClientListSummary();
    }

    private async void MainWindow_Loaded(object sender, RoutedEventArgs e)
    {
        Loaded -= MainWindow_Loaded;
        await LoadSavedRenewalQueueAsync();
        await LoadSavedClientListAsync();
        await RefreshHomeDashboardAsync();
    }

    private async Task LoadSavedClientListAsync()
    {
        try
        {
            var savedClients = await _databaseService.GetClientsAsync();
            var savedClientKeys = savedClients
                .Select(client => NormalizeClientMatchKey(client.ClientName))
                .Where(key => !string.IsNullOrWhiteSpace(key))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            var prospectRows = DiscoverGeneratedPresentationProspects(savedClientKeys);

            _clientList.Clear();
            foreach (var client in savedClients
                         .Where(client => !string.IsNullOrWhiteSpace(client.ClientName))
                         .OrderBy(client => client.ClientName, StringComparer.OrdinalIgnoreCase))
            {
                _clientList.Add(CreateClientListItem(client));
            }

            foreach (var prospect in prospectRows.OrderBy(row => row.ClientName, StringComparer.OrdinalIgnoreCase))
            {
                _clientList.Add(prospect);
            }

            RefreshClientListFilters();
            _clientListView?.Refresh();
            UpdateClientListSummary();

            SetStatus(savedClients.Count > 0 || prospectRows.Count > 0
                ? $"Loaded {_clientList.Count.ToString("#,0", CultureInfo.InvariantCulture)} client(s) from SQLite and generated presentation evidence."
                : "No saved clients were found in SQLite yet.");
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            RefreshClientListFilters();
            _clientListView?.Refresh();
            UpdateClientListSummary();
            SetStatus("Clients: could not refresh from SQLite, so the current visible list was kept.");
        }
    }

    private ClientListItem CreateClientListItem(ClientRecord client)
    {
        var renewal = FindRenewalQueueItemForClient(client.ClientName);
        var clientId = FirstNonBlank(client.ClientNumber, $"CL-{client.ClientId:0000}");
        var status = FirstNonBlank(client.Status, renewal is null ? "Active" : "Active");
        var industry = FirstNonBlank(client.Industry, "Not set");
        var employees = renewal?.Lives ?? 0;
        var renewalDate = renewal?.RenewalDateText;
        if (string.IsNullOrWhiteSpace(renewalDate) || renewalDate.Equals("N/A", StringComparison.OrdinalIgnoreCase))
        {
            renewalDate = client.EffectiveDate?.ToString("MMM d, yyyy", CultureInfo.InvariantCulture) ?? "TBD";
        }

        return new ClientListItem(
            clientId,
            ClientNameResolver.ToDisplayName(client.ClientName),
            status,
            industry,
            employees,
            renewalDate,
            FirstNonBlank(client.Carrier, renewal?.Carrier, "TBD"),
            "TBD",
            EstimateClientHealthScore(renewal, status),
            FormatClientLastUpdated(client.DateUpdated),
            client.ClientId);
    }

    private RenewalQueueItem? FindRenewalQueueItemForClient(string clientName)
    {
        var key = NormalizeClientMatchKey(clientName);
        if (string.IsNullOrWhiteSpace(key))
        {
            return null;
        }

        return _renewalQueue
            .Where(row => NormalizeClientMatchKey(row.Client).Equals(key, StringComparison.OrdinalIgnoreCase))
            .OrderByDescending(row => row.RenewalDate?.DayNumber ?? 0)
            .FirstOrDefault();
    }

    private void RefreshClientListFilters()
    {
        ConfigureFilterCombo(ClientStatusFilterComboBox, "All Statuses", _clientList.Select(row => row.Status).Distinct().OrderBy(value => value));
        ConfigureIndustryControls();
    }

    private void ConfigureIndustryControls()
    {
        var selectedFilter = SelectedFilterValue(ClientIndustryFilterComboBox, "All Industries");
        var selectedEditor = ClientIndustryEditorComboBox.SelectedItem?.ToString() ?? "";
        var savedIndustries = _clientList
            .Select(row => row.Industry)
            .Where(industry => !string.IsNullOrWhiteSpace(industry) && !industry.Equals("Not set", StringComparison.OrdinalIgnoreCase))
            .OrderBy(industry => industry, StringComparer.OrdinalIgnoreCase);
        var industries = ClientIndustryOptions
            .Concat(savedIndustries)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        ConfigureFilterCombo(
            ClientIndustryFilterComboBox,
            "All Industries",
            new[] { "Not set" }.Concat(industries));
        RestoreFilterSelection(ClientIndustryFilterComboBox, selectedFilter);

        ClientIndustryEditorComboBox.Items.Clear();
        foreach (var industry in industries)
        {
            ClientIndustryEditorComboBox.Items.Add(industry);
        }

        if (!string.IsNullOrWhiteSpace(selectedEditor))
        {
            RestoreFilterSelection(ClientIndustryEditorComboBox, selectedEditor);
        }
    }

    private List<ClientListItem> DiscoverGeneratedPresentationProspects(IReadOnlySet<string> savedClientKeys)
    {
        var evidenceByClient = new Dictionary<string, ProspectPresentationEvidence>(StringComparer.OrdinalIgnoreCase);
        foreach (var presentationPath in EnumeratePresentationFiles())
        {
            var isProposal = IsProposalPresentationPath(presentationPath);
            var isRenewal = IsRenewalPresentationPath(presentationPath);
            if (!isProposal && !isRenewal)
            {
                continue;
            }

            var clientName = InferClientNameFromPresentationPath(presentationPath);
            var clientKey = NormalizeClientMatchKey(clientName);
            if (string.IsNullOrWhiteSpace(clientKey) || savedClientKeys.Contains(clientKey))
            {
                continue;
            }

            if (!evidenceByClient.TryGetValue(clientKey, out var evidence))
            {
                evidence = new ProspectPresentationEvidence
                {
                    ClientName = ClientNameResolver.ToDisplayName(clientName),
                    LastUpdated = File.GetLastWriteTime(presentationPath)
                };
                evidenceByClient[clientKey] = evidence;
            }

            evidence.HasProposalPresentation |= isProposal;
            evidence.HasRenewalPresentation |= isRenewal;
            evidence.LastUpdated = MaxDate(evidence.LastUpdated, File.GetLastWriteTime(presentationPath));
        }

        return evidenceByClient.Values
            .Where(evidence => evidence.HasProposalPresentation && evidence.HasRenewalPresentation)
            .Select((evidence, index) => new ClientListItem(
                $"PROS-{index + 1:0000}",
                evidence.ClientName,
                "New Prospect",
                "Prospect",
                0,
                "TBD",
                "TBD",
                "TBD",
                null,
                evidence.LastUpdated == DateTime.MinValue
                    ? "Presentation found"
                    : evidence.LastUpdated.ToString("MMM d, yyyy", CultureInfo.InvariantCulture)))
            .ToList();
    }

    private IEnumerable<string> EnumeratePresentationFiles()
    {
        var roots = CandidatePresentationRoots()
            .Where(path => !string.IsNullOrWhiteSpace(path) && Directory.Exists(path))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        const int maxFiles = 3000;
        var count = 0;
        foreach (var root in roots)
        {
            IEnumerable<string> files;
            try
            {
                files = Directory.EnumerateFiles(
                    root,
                    "*.pptx",
                    new EnumerationOptions
                    {
                        RecurseSubdirectories = true,
                        IgnoreInaccessible = true,
                        ReturnSpecialDirectories = false
                    });
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                continue;
            }

            foreach (var file in files)
            {
                if (Path.GetFileName(file).StartsWith("~$", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                yield return file;
                count++;
                if (count >= maxFiles)
                {
                    yield break;
                }
            }
        }
    }

    private IEnumerable<string> CandidatePresentationRoots()
    {
        var documents = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        var selectedOutputFolder = OutputFolderTextBox.Text.Trim();
        if (SharePointWebDavPathPolicy.CanEnumeratePresentationRoot(selectedOutputFolder))
        {
            yield return selectedOutputFolder;
        }

        yield return Path.Combine(documents, "Renewal Review Builder", "Outputs");
        yield return AppPaths.DataFolder;
    }

    private static bool IsProposalPresentationPath(string path)
    {
        var text = path.ToLowerInvariant();
        return text.Contains("proposal", StringComparison.OrdinalIgnoreCase) ||
               text.Contains("quote", StringComparison.OrdinalIgnoreCase) ||
               text.Contains("carrier proposal comparison", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsRenewalPresentationPath(string path)
    {
        var fileName = Path.GetFileName(path);
        return fileName.Contains("renewal", StringComparison.OrdinalIgnoreCase) ||
               fileName.Equals("Client_Renewal_Draft.pptx", StringComparison.OrdinalIgnoreCase);
    }

    private static string InferClientNameFromPresentationPath(string path)
    {
        var fileName = Path.GetFileNameWithoutExtension(path);
        var fileCandidate = CleanPresentationClientName(fileName);
        if (IsUsableInferredClientName(fileCandidate))
        {
            return fileCandidate;
        }

        var directory = Directory.GetParent(path);
        while (directory is not null)
        {
            var candidate = CleanPresentationClientName(directory.Name);
            if (IsUsableInferredClientName(candidate))
            {
                return candidate;
            }

            directory = directory.Parent;
        }

        return "";
    }

    private static string CleanPresentationClientName(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return "";
        }

        var cleaned = value
            .Replace('_', ' ')
            .Replace('-', ' ')
            .Trim();

        var timestampIndex = cleaned.LastIndexOf(" 20", StringComparison.OrdinalIgnoreCase);
        if (timestampIndex > 0)
        {
            cleaned = cleaned[..timestampIndex].Trim();
        }

        var timestampUnderscoreIndex = value.LastIndexOf("_20", StringComparison.OrdinalIgnoreCase);
        if (timestampUnderscoreIndex > 0)
        {
            cleaned = value[..timestampUnderscoreIndex].Replace('_', ' ').Replace('-', ' ').Trim();
        }

        return cleaned;
    }

    private static bool IsUsableInferredClientName(string value)
    {
        if (string.IsNullOrWhiteSpace(value) || value.Length < 3)
        {
            return false;
        }

        var genericNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Outputs",
            "Output",
            "Renewal Review Builder",
            "Carrier Proposal Comparison",
            "Proposal Comparison",
            "Client Renewal Draft",
            "Batch Output",
            "Supporting PDFs",
            "Supporting Excel Files",
            "Executive Reports",
            "Data",
            "Westland",
            "RenewalReviewBuilder"
        };

        if (genericNames.Contains(value))
        {
            return false;
        }

        return value.Any(char.IsLetterOrDigit);
    }

    private static string NormalizeClientMatchKey(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return "";
        }

        var cleaned = new string(value
            .ToUpperInvariant()
            .Select(character => char.IsLetterOrDigit(character) ? character : ' ')
            .ToArray());
        var suffixes = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "THE",
            "LTD",
            "LIMITED",
            "INC",
            "INCORPORATED",
            "CORP",
            "CORPORATION",
            "CO",
            "COMPANY"
        };
        var tokens = cleaned
            .Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(token => !suffixes.Contains(token));

        return string.Join(" ", tokens);
    }

    private static int? EstimateClientHealthScore(RenewalQueueItem? renewal, string status)
    {
        return ClientWorkspaceRenewalHistoryPresentationService.CalculateHealthScore(
            renewal is null ? null : CreateClientWorkspaceHistoryEntry(renewal),
            status);
    }

    private static ClientWorkspaceRenewalHistoryEntry CreateClientWorkspaceHistoryEntry(RenewalQueueItem renewal)
    {
        return new ClientWorkspaceRenewalHistoryEntry
        {
            RenewalDate = renewal.RenewalDate,
            RenewalYear = renewal.RenewalDate?.Year ?? 0,
            CurrentPremium = renewal.CurrentPremium,
            RenewalPremium = renewal.RenewalPremium,
            PremiumChangePercent = renewal.PremiumChangePercent,
            EhcLossRatio = renewal.EhcLossRatio,
            DentalLossRatio = renewal.DentalLossRatio,
            IsHighRisk = renewal.Status.Equals(RenewalQueueStatuses.HighRisk, StringComparison.OrdinalIgnoreCase),
            DateLoaded = renewal.DateLoaded
        };
    }

    private static string FormatClientLastUpdated(DateTime value)
    {
        return value == DateTime.MinValue
            ? "Not saved"
            : value.ToString("MMM d, yyyy", CultureInfo.InvariantCulture);
    }

    private static DateTime MaxDate(DateTime first, DateTime second)
    {
        return first >= second ? first : second;
    }

    private async Task LoadSavedRenewalQueueAsync()
    {
        try
        {
            var savedRenewals = await _databaseService.GetRenewalQueueAsync();
            if (savedRenewals.Count > 0)
            {
                RemoveSeedRenewalQueueItems();
                RemoveSavedRenewalQueueItems();
            }

            var loaded = 0;
            foreach (var savedRenewal in savedRenewals)
            {
                var savedQueueItem = CreateRenewalQueueItem(savedRenewal);
                if (ClientNameResolver.IsInvalidClientName(savedQueueItem.Client))
                {
                    continue;
                }

                var existingDurableDuplicate = _renewalQueue.FirstOrDefault(row =>
                    IsDurableRenewalQueueItem(row) &&
                    IsSameVisibleRenewal(row, savedQueueItem));
                if (existingDurableDuplicate is not null)
                {
                    if (existingDurableDuplicate.Id.Equals("current-renewal", StringComparison.OrdinalIgnoreCase) &&
                        !IsCurrentAnalysisContext(existingDurableDuplicate))
                    {
                        _renewalQueue.Remove(existingDurableDuplicate);
                    }
                    else
                    {
                        continue;
                    }
                }

                RemoveVisibleRenewalDuplicates(savedQueueItem);
                _renewalQueue.Add(savedQueueItem);
                loaded++;
            }

            if (loaded == 0)
            {
                return;
            }

            RefreshRenewalQueueFilters();
            _renewalQueueView?.Refresh();
            UpdateRenewalQueueSummary();
            RenewalQueueGrid.SelectedItem = _clientRenewalSelectionIdentity is null
                ? _renewalQueueView?.Cast<RenewalQueueItem>().FirstOrDefault()
                : FindRenewalForClient(_clientRenewalSelectionIdentity);
            UpdateSelectedRenewalDetail(RenewalQueueGrid.SelectedItem as RenewalQueueItem);
            SetStatus($"Loaded {loaded} saved renewal(s) from Benefits Intelligence history.");
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            SetStatus("Saved renewal history could not be loaded into the queue.");
        }
    }

    private void RemoveSeedRenewalQueueItems()
    {
        var seedRows = _renewalQueue
            .Where(row => !IsDurableRenewalQueueItem(row))
            .ToList();
        foreach (var row in seedRows)
        {
            _renewalQueue.Remove(row);
        }
    }

    private void RemoveSavedRenewalQueueItems()
    {
        var savedRows = _renewalQueue
            .Where(row => row.Id.StartsWith("saved-renewal-", StringComparison.OrdinalIgnoreCase))
            .ToList();
        foreach (var row in savedRows)
        {
            _renewalQueue.Remove(row);
        }
    }

    private void RemoveVisibleRenewalDuplicates(RenewalQueueItem incoming)
    {
        var duplicates = _renewalQueue
            .Where(row => IsSameVisibleRenewal(row, incoming))
            .ToList();
        foreach (var duplicate in duplicates)
        {
            _renewalQueue.Remove(duplicate);
        }
    }

    private static bool IsDurableRenewalQueueItem(RenewalQueueItem row)
    {
        return row.Id.Equals("current-renewal", StringComparison.OrdinalIgnoreCase) ||
               row.ClientId > 0 ||
               row.SourceDocumentId.HasValue;
    }

    private static bool IsSameVisibleRenewal(RenewalQueueItem existing, RenewalQueueItem incoming)
    {
        if (existing.SourceDocumentId.HasValue &&
            incoming.SourceDocumentId.HasValue &&
            existing.SourceDocumentId == incoming.SourceDocumentId)
        {
            return true;
        }

        if (!SameRenewalMonth(existing.RenewalDate, incoming.RenewalDate))
        {
            return false;
        }

        if (existing.ClientId > 0 && incoming.ClientId > 0 && existing.ClientId == incoming.ClientId)
        {
            return true;
        }

        var existingClient = NormalizeRenewalQueueKeyText(existing.Client);
        var incomingClient = NormalizeRenewalQueueKeyText(incoming.Client);
        if (string.IsNullOrWhiteSpace(existingClient) ||
            string.IsNullOrWhiteSpace(incomingClient) ||
            !existingClient.Equals(incomingClient, StringComparison.Ordinal))
        {
            return false;
        }

        var existingCarrier = NormalizeRenewalQueueKeyText(existing.Carrier);
        var incomingCarrier = NormalizeRenewalQueueKeyText(incoming.Carrier);
        return string.IsNullOrWhiteSpace(existingCarrier) ||
               string.IsNullOrWhiteSpace(incomingCarrier) ||
               existingCarrier.Equals(incomingCarrier, StringComparison.Ordinal);
    }

    private static bool SameRenewalMonth(DateOnly? left, DateOnly? right)
    {
        return left.HasValue &&
               right.HasValue &&
               left.Value.Year == right.Value.Year &&
               left.Value.Month == right.Value.Month;
    }

    private static string NormalizeRenewalQueueKeyText(string value)
    {
        return new string((value ?? "")
            .ToUpperInvariant()
            .Where(char.IsLetterOrDigit)
            .ToArray());
    }

    private bool IsCurrentAnalysisContext(RenewalQueueItem row)
    {
        if (_analysis is null)
        {
            return false;
        }

        var analysisClient = FirstNonBlank(
            ClientNameResolver.Resolve(
                _analysis.RenewalCase.ClientName,
                _analysis.Pdf.ClientName,
                _analysis.RenewalCase.PdfPath),
            _analysis.RenewalCase.ClientName,
            _analysis.Pdf.ClientName);
        var selectedIdentity = new RenewalWorkflowIdentity(
            row.ClientId,
            row.SourceDocumentId,
            row.Client,
            row.PolicyNumber,
            row.Carrier,
            row.RenewalDate);
        var analysisIdentity = new RenewalWorkflowIdentity(
            _analysisClientId ?? 0,
            _analysisSourceDocumentId,
            analysisClient,
            CurrentAnalysisPolicyNumber(),
            DetectCarrier(_analysis),
            _analysis.RenewalCase.RenewalDate ?? _analysis.Pdf.RenewalDate);

        return RenewalWorkflowContextGuard.Matches(selectedIdentity, analysisIdentity);
    }

    private async Task<string?> EnsureAnalysisContextForExportAsync(RenewalQueueItem row)
    {
        if (IsCurrentAnalysisContext(row))
        {
            return null;
        }

        if (row.ClientId <= 0 || row.SourceDocumentId is null or <= 0)
        {
            return "This saved renewal does not have an exact source-document identity. Validate it once to enable database-only exports.";
        }

        try
        {
            SetBusy(true, "Restoring saved renewal analysis...");
            var snapshot = await _databaseService.GetRenewalAnalysisSnapshotAsync(
                row.ClientId,
                row.SourceDocumentId.Value,
                row.RenewalDate);
            if (!snapshot.Succeeded || snapshot.Analysis is null)
            {
                return snapshot.Found
                    ? snapshot.Message
                    : "This renewal was saved before reusable export snapshots were enabled. Validate it once to enable future database-only exports.";
            }

            var restoredAnalysis = snapshot.Analysis;
            var restoredClient = FirstNonBlank(
                ClientNameResolver.Resolve(
                    restoredAnalysis.RenewalCase.ClientName,
                    restoredAnalysis.Pdf.ClientName,
                    restoredAnalysis.RenewalCase.PdfPath),
                restoredAnalysis.RenewalCase.ClientName,
                restoredAnalysis.Pdf.ClientName);
            var selectedIdentity = new RenewalWorkflowIdentity(
                row.ClientId,
                row.SourceDocumentId,
                row.Client,
                row.PolicyNumber,
                row.Carrier,
                row.RenewalDate);
            var restoredIdentity = new RenewalWorkflowIdentity(
                row.ClientId,
                row.SourceDocumentId,
                restoredClient,
                FirstNonBlank(restoredAnalysis.Pdf.PolicyNumber, restoredAnalysis.RenewalCase.PolicyNumber),
                DetectCarrier(restoredAnalysis),
                restoredAnalysis.RenewalCase.RenewalDate ?? restoredAnalysis.Pdf.RenewalDate);
            if (!RenewalWorkflowContextGuard.Matches(selectedIdentity, restoredIdentity))
            {
                return "The saved analysis identity does not match the selected renewal. Validate the selected renewal once before exporting.";
            }

            var outputFolder = OutputFolderTextBox.Text.Trim();
            if (!string.IsNullOrWhiteSpace(outputFolder))
            {
                restoredAnalysis.RenewalCase.OutputFolder = outputFolder;
            }

            var templatePath = TemplatePathTextBox.Text.Trim();
            if (!string.IsNullOrWhiteSpace(templatePath) && File.Exists(templatePath))
            {
                restoredAnalysis.RenewalCase.PowerPointTemplatePath = templatePath;
            }

            restoredAnalysis.RenewalCase = restoredAnalysis.RenewalCase;
            _analysis = restoredAnalysis;
            _analysisClientId = row.ClientId;
            _analysisSourceDocumentId = row.SourceDocumentId;
            _latestSavedClientId = row.ClientId;
            _latestSavedDocumentId = row.SourceDocumentId;
            _claimsRangeReviewViewModel.ResetSession();
            UpdateReview();
            SetStatus($"Restored {row.Client}'s reviewed renewal from Benefits Intelligence.");
            return null;
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            return $"The saved renewal analysis could not be restored: {ex.Message}";
        }
        finally
        {
            SetBusy(false);
        }
    }

    private string CurrentAnalysisPolicyNumber()
    {
        return _analysis is null
            ? ""
            : FirstNonBlank(_analysis.Pdf.PolicyNumber, _analysis.RenewalCase.PolicyNumber);
    }

    private async void HomeDashboardFilter_Changed(object sender, SelectionChangedEventArgs e)
    {
        if (_isUpdatingHomeDashboardFilters || !IsLoaded)
        {
            return;
        }

        await RefreshHomeDashboardAsync(refreshClientFilter: false);
    }

    private async void HomeDashboardCustomRange_Changed(object sender, SelectionChangedEventArgs e)
    {
        if (_isUpdatingHomeDashboardFilters || !IsLoaded)
        {
            return;
        }

        await RefreshHomeDashboardAsync(refreshClientFilter: false);
    }

    private async Task RefreshHomeDashboardAsync(bool refreshClientFilter = true)
    {
        try
        {
            var savedRenewals = await _databaseService.GetRenewalQueueAsync();
            var clients = await _databaseService.GetClientsAsync();
            if (refreshClientFilter)
            {
                PopulateHomeDashboardFilters(savedRenewals, clients);
            }

            var selectedClient = SelectedHomeClientFilter();
            var selectedRenewalYear = SelectedHomeRenewalYearFilter();
            var selectedEffectiveMonth = SelectedHomeEffectiveMonthFilter();
            var dateRange = SelectedHomeDateRange();
            var filteredRenewals = savedRenewals
                .Where(row => string.IsNullOrWhiteSpace(selectedClient) ||
                              row.ClientName.Equals(selectedClient, StringComparison.OrdinalIgnoreCase))
                .Where(row => !selectedRenewalYear.HasValue ||
                              (row.RenewalDate.HasValue && row.RenewalDate.Value.Year == selectedRenewalYear.Value))
                .Where(row => !selectedEffectiveMonth.HasValue ||
                              (row.RenewalDate.HasValue && row.RenewalDate.Value.Month == selectedEffectiveMonth.Value))
                .Where(row => DateInHomeDashboardRange(row, dateRange))
                .ToList();

            var clientCount = clients
                .Where(client => string.IsNullOrWhiteSpace(selectedClient) ||
                                 client.ClientName.Equals(selectedClient, StringComparison.OrdinalIgnoreCase))
                .Select(client => client.ClientId)
                .Distinct()
                .Count();
            var avgRenewalIncrease = filteredRenewals.AverageValue(CalculateRenewalIncreasePercent);
            var avgEhcLossRatio = filteredRenewals.AverageValue(row => row.EhcLossRatio);
            var totalPremium = filteredRenewals.SumValue(row => row.RenewalPremium ?? row.CurrentPremium);
            var opportunityCount = filteredRenewals.Count(IsHighPriorityOpportunity);
            var filterLabel = BuildHomeFilterLabel(selectedRenewalYear, selectedEffectiveMonth, dateRange.Label);

            HomeClientsKpiTextBlock.Text = clientCount.ToString(CultureInfo.InvariantCulture);
            HomeRenewalsKpiTextBlock.Text = filteredRenewals.Count.ToString(CultureInfo.InvariantCulture);
            HomeAverageIncreaseKpiTextBlock.Text = FormatPercent(avgRenewalIncrease);
            HomeAverageEhcLossRatioKpiTextBlock.Text = FormatPercent(avgEhcLossRatio, includeSign: false);
            HomeTotalPremiumKpiTextBlock.Text = FormatDashboardMoney(totalPremium, emptyAsZero: true);
            HomeOpportunitiesKpiTextBlock.Text = opportunityCount.ToString(CultureInfo.InvariantCulture);

            HomeClientsKpiSubtitleTextBlock.Text = string.IsNullOrWhiteSpace(selectedClient) ? "Loaded Clients" : "Selected Client";
            HomeRenewalsKpiSubtitleTextBlock.Text = "Saved Renewals";
            HomeAverageIncreaseKpiSubtitleTextBlock.Text = filterLabel;
            HomeAverageEhcLossRatioKpiSubtitleTextBlock.Text = filterLabel;
            HomeTotalPremiumKpiSubtitleTextBlock.Text = filterLabel;
            HomeOpportunitiesKpiSubtitleTextBlock.Text = "High Priority";

            RefreshHomeRenewalPipeline(filteredRenewals);
            RefreshHomeRecentActivity(filteredRenewals);

            HomeDashboardDataSourceTextBlock.Text = savedRenewals.Count == 0
                ? "No saved renewal history yet. KPIs will populate as renewals are saved to the local database."
                : $"Based on {filteredRenewals.Count} saved renewal(s) from Benefits Intelligence history.";
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            HomeClientsKpiTextBlock.Text = "0";
            HomeRenewalsKpiTextBlock.Text = "0";
            HomeAverageIncreaseKpiTextBlock.Text = "N/A";
            HomeAverageEhcLossRatioKpiTextBlock.Text = "N/A";
            HomeTotalPremiumKpiTextBlock.Text = "$0";
            HomeOpportunitiesKpiTextBlock.Text = "0";
            RefreshHomeRenewalPipeline([]);
            RefreshHomeRecentActivity([]);
            HomeDashboardDataSourceTextBlock.Text = "Dashboard KPIs could not be loaded from saved renewal history.";
            SetStatus("Dashboard KPIs could not be loaded from saved renewal history.");
        }
    }

    private void PopulateHomeDashboardFilters(
        IReadOnlyCollection<RenewalQueueRecord> savedRenewals,
        IReadOnlyCollection<ClientRecord> clients)
    {
        var selectedClient = SelectedHomeClientFilter();
        var selectedYear = SelectedComboBoxText(HomeRenewalYearFilterComboBox);
        var selectedMonth = SelectedComboBoxText(HomeEffectiveMonthFilterComboBox);
        var selectedDateRange = SelectedComboBoxText(HomeDateRangeComboBox);

        var clientNames = clients
            .Select(client => client.ClientName)
            .Concat(savedRenewals.Select(row => row.ClientName))
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(value => value)
            .ToList();

        _isUpdatingHomeDashboardFilters = true;
        HomeClientFilterComboBox.Items.Clear();
        HomeClientFilterComboBox.Items.Add("All Clients");
        foreach (var clientName in clientNames)
        {
            HomeClientFilterComboBox.Items.Add(clientName);
        }

        HomeClientFilterComboBox.SelectedItem = !string.IsNullOrWhiteSpace(selectedClient) &&
                                                clientNames.Contains(selectedClient, StringComparer.OrdinalIgnoreCase)
            ? selectedClient
            : "All Clients";

        EnsureComboBoxSelection(
            HomeRenewalYearFilterComboBox,
            _homeDashboardFiltersInitialized && !string.IsNullOrWhiteSpace(selectedYear)
                ? selectedYear
                : DateTime.Today.Year.ToString(CultureInfo.InvariantCulture));
        EnsureComboBoxSelection(
            HomeEffectiveMonthFilterComboBox,
            _homeDashboardFiltersInitialized && !string.IsNullOrWhiteSpace(selectedMonth)
                ? selectedMonth
                : "All Months");
        EnsureComboBoxSelection(
            HomeDateRangeComboBox,
            _homeDashboardFiltersInitialized && !string.IsNullOrWhiteSpace(selectedDateRange)
                ? selectedDateRange
                : "Current Renewals");
        UpdateHomeCustomRangeVisibility();
        _homeDashboardFiltersInitialized = true;
        _isUpdatingHomeDashboardFilters = false;
    }

    private string SelectedHomeClientFilter()
    {
        var value = SelectedComboBoxText(HomeClientFilterComboBox);
        return value.Equals("All Clients", StringComparison.OrdinalIgnoreCase) ? "" : value;
    }

    private int? SelectedHomeRenewalYearFilter()
    {
        var value = SelectedComboBoxText(HomeRenewalYearFilterComboBox);
        return value.Equals("All Years", StringComparison.OrdinalIgnoreCase)
            ? null
            : int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var year)
                ? year
                : null;
    }

    private int? SelectedHomeEffectiveMonthFilter()
    {
        var value = SelectedComboBoxText(HomeEffectiveMonthFilterComboBox);
        if (value.Equals("All Months", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        return DateTime.TryParseExact(value, "MMMM", CultureInfo.InvariantCulture, DateTimeStyles.None, out var month)
            ? month.Month
            : null;
    }

    private HomeDateRangeSelection SelectedHomeDateRange()
    {
        var value = SelectedComboBoxText(HomeDateRangeComboBox);
        var today = DateOnly.FromDateTime(DateTime.Today);
        var year = today.Year;
        UpdateHomeCustomRangeVisibility();

        return value switch
        {
            "Current Renewals" => new HomeDateRangeSelection(
                new DateOnly(year, 1, 1),
                today,
                "Current Renewals",
                HomeDateRangeBasis.WorkingDate),
            "Next 30 Days" => new HomeDateRangeSelection(today, today.AddDays(30), value, HomeDateRangeBasis.EffectiveDate),
            "Next 60 Days" => new HomeDateRangeSelection(today, today.AddDays(60), value, HomeDateRangeBasis.EffectiveDate),
            "Next 90 Days" => new HomeDateRangeSelection(today, today.AddDays(90), value, HomeDateRangeBasis.EffectiveDate),
            "This Year" => new HomeDateRangeSelection(new DateOnly(year, 1, 1), new DateOnly(year, 12, 31), value, HomeDateRangeBasis.EffectiveDate),
            "Next Year" => new HomeDateRangeSelection(new DateOnly(year + 1, 1, 1), new DateOnly(year + 1, 12, 31), value, HomeDateRangeBasis.EffectiveDate),
            "Last Year" => new HomeDateRangeSelection(new DateOnly(year - 1, 1, 1), new DateOnly(year - 1, 12, 31), value, HomeDateRangeBasis.EffectiveDate),
            "Last 12 Months" => new HomeDateRangeSelection(today.AddMonths(-12), today, value, HomeDateRangeBasis.EffectiveDate),
            "Last 3 Years" => new HomeDateRangeSelection(today.AddYears(-3), today, value, HomeDateRangeBasis.EffectiveDate),
            "Last 5 Years" => new HomeDateRangeSelection(today.AddYears(-5), today, value, HomeDateRangeBasis.EffectiveDate),
            "Custom Range..." => new HomeDateRangeSelection(
                DateOnlyFromDateTime(HomeCustomStartDatePicker.SelectedDate),
                DateOnlyFromDateTime(HomeCustomEndDatePicker.SelectedDate),
                "Custom Range",
                HomeDateRangeBasis.EffectiveDate),
            _ => new HomeDateRangeSelection(null, null, "All History", HomeDateRangeBasis.EffectiveDate)
        };
    }

    private bool DateInHomeDashboardRange(RenewalQueueRecord row, HomeDateRangeSelection dateRange)
    {
        if (!dateRange.Start.HasValue && !dateRange.End.HasValue)
        {
            return true;
        }

        var date = dateRange.Basis == HomeDateRangeBasis.WorkingDate
            ? WorkingDateForRenewal(row)
            : row.RenewalDate;

        return date.HasValue &&
               (!dateRange.Start.HasValue || date.Value >= dateRange.Start.Value) &&
               (!dateRange.End.HasValue || date.Value <= dateRange.End.Value);
    }

    private static bool IsHighPriorityOpportunity(RenewalQueueRecord row)
    {
        return row.Status.Contains("High Risk", StringComparison.OrdinalIgnoreCase) ||
               row.Status.Contains("Missing Data", StringComparison.OrdinalIgnoreCase) ||
               CalculateRenewalIncreasePercent(row) is >= 10m ||
               row.EhcLossRatio is >= 85m ||
               row.DentalLossRatio is >= 85m ||
               HasMissingRequiredRenewalData(row) ||
               HasPoolingChargeIncreased(row) ||
               IsPremiumIncreaseAboveBenchmark(row);
    }

    private void UpdateHomeCustomRangeVisibility()
    {
        if (HomeCustomRangePanel is null || HomeDateRangeComboBox is null)
        {
            return;
        }

        HomeCustomRangePanel.Visibility = SelectedComboBoxText(HomeDateRangeComboBox).Equals("Custom Range...", StringComparison.OrdinalIgnoreCase)
            ? Visibility.Visible
            : Visibility.Collapsed;
    }

    private static void EnsureComboBoxSelection(ComboBox comboBox, string preferredValue)
    {
        foreach (var item in comboBox.Items)
        {
            var value = item is ComboBoxItem comboBoxItem
                ? comboBoxItem.Content?.ToString()
                : item?.ToString();
            if (value?.Equals(preferredValue, StringComparison.OrdinalIgnoreCase) == true)
            {
                comboBox.SelectedItem = item;
                return;
            }
        }

        if (comboBox.Items.Count > 0)
        {
            comboBox.SelectedIndex = 0;
        }
    }

    private static bool IsActiveClient(ClientRecord client)
    {
        return string.IsNullOrWhiteSpace(client.Status) ||
               client.Status.Equals("Active", StringComparison.OrdinalIgnoreCase);
    }

    private static decimal? CalculateRenewalIncreasePercent(RenewalQueueRecord row)
    {
        if (row.PremiumChangePercent.HasValue)
        {
            return row.PremiumChangePercent.Value;
        }

        if (row.CurrentPremium is > 0m && row.RenewalPremium.HasValue)
        {
            return Math.Round((row.RenewalPremium.Value / row.CurrentPremium.Value - 1m) * 100m, 1);
        }

        return null;
    }

    private static bool HasMissingRequiredRenewalData(RenewalQueueRecord row)
    {
        return !row.RenewalDate.HasValue ||
               !row.CurrentPremium.HasValue ||
               !row.RenewalPremium.HasValue ||
               !row.Lives.HasValue ||
               (!row.EhcLossRatio.HasValue && !row.DentalLossRatio.HasValue) ||
               row.Status.Contains("Missing Data", StringComparison.OrdinalIgnoreCase);
    }

    private static bool HasPoolingChargeIncreased(RenewalQueueRecord row)
    {
        return NormalizePercent(row.PoolingCharge) is >= 10m;
    }

    private static bool IsPremiumIncreaseAboveBenchmark(RenewalQueueRecord row)
    {
        return CalculateRenewalIncreasePercent(row) is >= 10m;
    }

    private static DateOnly? DateOnlyFromDateTime(DateTime? value)
    {
        return value.HasValue ? DateOnly.FromDateTime(value.Value) : null;
    }

    private static DateOnly? WorkingDateForRenewal(RenewalQueueRecord row)
    {
        return row.DateLoaded == DateTime.MinValue ? null : DateOnly.FromDateTime(row.DateLoaded);
    }

    private static string BuildHomeFilterLabel(int? renewalYear, int? effectiveMonth, string dateRangeLabel)
    {
        var yearLabel = renewalYear?.ToString(CultureInfo.InvariantCulture) ?? "All Years";
        var monthLabel = effectiveMonth.HasValue
            ? CultureInfo.InvariantCulture.DateTimeFormat.GetMonthName(effectiveMonth.Value)
            : "All Months";
        return $"{yearLabel} | {monthLabel} | {dateRangeLabel}";
    }

    private void RefreshHomeRenewalPipeline(IReadOnlyCollection<RenewalQueueRecord> renewals)
    {
        var stageCounts = renewals
            .GroupBy(MapHomeRenewalStage, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.Count(), StringComparer.OrdinalIgnoreCase);

        HomeRenewalPipelineTotalTextBlock.Text = renewals.Count.ToString("#,0", CultureInfo.InvariantCulture);
        HomePipelineNotStartedTextBlock.Text = $"{PipelineCount(stageCounts, "Not Started")} Not Started";
        HomePipelineDocumentsReceivedTextBlock.Text = $"{PipelineCount(stageCounts, "Documents Received")} Documents Received";
        HomePipelineValidationRequiredTextBlock.Text = $"{PipelineCount(stageCounts, "Validation Required")} Validation Required";
        HomePipelineUnderwritingReviewTextBlock.Text = $"{PipelineCount(stageCounts, "Underwriting Review")} Underwriting Review";
        HomePipelineConsultantReviewTextBlock.Text = $"{PipelineCount(stageCounts, "Consultant Strategy Review")} Consultant Strategy Review";
        HomePipelinePresentationReadyTextBlock.Text = $"{PipelineCount(stageCounts, "Presentation Ready")} Presentation Ready";
        HomePipelineDeliveredTextBlock.Text = $"{PipelineCount(stageCounts, "Delivered")} Delivered";
        HomePipelineCompletedTextBlock.Text = $"{PipelineCount(stageCounts, "Completed")} Completed";
    }

    private static int PipelineCount(IReadOnlyDictionary<string, int> stageCounts, string stage)
    {
        return stageCounts.TryGetValue(stage, out var count) ? count : 0;
    }

    private static string MapHomeRenewalStage(RenewalQueueRecord row)
    {
        var status = row.Status ?? "";
        var selectedVersion = row.SelectedRenewalVersion ?? "";
        var hasDocument = row.SourceDocumentId.HasValue || !string.IsNullOrWhiteSpace(row.SourceFilePath);
        var hasExtractedRenewalData = row.CurrentPremium.HasValue || row.RenewalPremium.HasValue || row.PremiumChangePercent.HasValue;

        if (status.Contains("Delivered", StringComparison.OrdinalIgnoreCase))
        {
            return "Delivered";
        }

        if (status.Equals(RenewalQueueStatuses.Completed, StringComparison.OrdinalIgnoreCase) ||
            status.Equals(RenewalQueueStatuses.FinalApproved, StringComparison.OrdinalIgnoreCase) ||
            selectedVersion.Equals(RenewalVersionNames.FinalApproved, StringComparison.OrdinalIgnoreCase))
        {
            return "Completed";
        }

        if (!hasDocument)
        {
            return "Not Started";
        }

        if (!hasExtractedRenewalData)
        {
            return "Documents Received";
        }

        if (HasMissingRequiredRenewalData(row))
        {
            return "Validation Required";
        }

        if (CalculateRenewalIncreasePercent(row) >= 10m ||
            row.EhcLossRatio >= 85m ||
            row.DentalLossRatio >= 85m ||
            status.Equals(RenewalQueueStatuses.HighRisk, StringComparison.OrdinalIgnoreCase) ||
            status.Equals(RenewalQueueStatuses.Watch, StringComparison.OrdinalIgnoreCase))
        {
            return "Underwriting Review";
        }

        if (selectedVersion.Equals(RenewalVersionNames.Negotiated, StringComparison.OrdinalIgnoreCase) ||
            status.Equals(RenewalQueueStatuses.Negotiated, StringComparison.OrdinalIgnoreCase))
        {
            return "Consultant Strategy Review";
        }

        if (status.Equals(RenewalQueueStatuses.ReadyToReview, StringComparison.OrdinalIgnoreCase) ||
            status.Equals(RenewalQueueStatuses.Healthy, StringComparison.OrdinalIgnoreCase))
        {
            return "Presentation Ready";
        }

        return "Presentation Ready";
    }

    private void RefreshHomeRecentActivity(IReadOnlyCollection<RenewalQueueRecord> renewals)
    {
        var activityControls = new[]
        {
            HomeRecentActivity1TextBlock,
            HomeRecentActivity2TextBlock,
            HomeRecentActivity3TextBlock,
            HomeRecentActivity4TextBlock,
            HomeRecentActivity5TextBlock
        };

        var activities = renewals
            .OrderByDescending(row => row.DateLoaded)
            .ThenByDescending(row => row.RenewalDate)
            .Take(activityControls.Length)
            .Select(row => $"{row.ClientName} renewal {row.RenewalDate?.Year.ToString(CultureInfo.InvariantCulture) ?? "date pending"} loaded")
            .ToList();

        for (var index = 0; index < activityControls.Length; index++)
        {
            activityControls[index].Text = index < activities.Count
                ? activities[index]
                : index == 0
                    ? "No recent activity for selected filters."
                    : "";
        }
    }

    private static string SelectedComboBoxText(ComboBox comboBox)
    {
        return comboBox.SelectedItem switch
        {
            ComboBoxItem item => item.Content?.ToString()?.Trim() ?? "",
            null => "",
            var item => item.ToString()?.Trim() ?? ""
        };
    }

    private static string FormatDashboardMoney(decimal? value, bool emptyAsZero = false)
    {
        if (!value.HasValue)
        {
            return emptyAsZero ? "$0" : "N/A";
        }

        var amount = value.Value;
        var absolute = Math.Abs(amount);
        if (absolute >= 1000000m)
        {
            return $"{(amount < 0 ? "-" : "")}${absolute / 1000000m:0.#}M";
        }

        if (absolute >= 1000m)
        {
            return $"{(amount < 0 ? "-" : "")}${absolute / 1000m:0.#}K";
        }

        return amount.ToString("$#,0", CultureInfo.InvariantCulture);
    }

    private static void ConfigureFilterCombo(ComboBox comboBox, string allLabel, IEnumerable<string> values)
    {
        comboBox.Items.Clear();
        comboBox.Items.Add(allLabel);
        foreach (var value in values.Where(value => !string.IsNullOrWhiteSpace(value)))
        {
            comboBox.Items.Add(value);
        }

        comboBox.SelectedIndex = 0;
    }

    private bool RenewalQueueFilter(object item)
    {
        if (item is not RenewalQueueItem row)
        {
            return false;
        }

        var search = RenewalQueueSearchTextBox?.Text?.Trim() ?? "";
        var status = SelectedFilterValue(RenewalStatusFilterComboBox, "All statuses");
        var carrier = SelectedFilterValue(RenewalCarrierFilterComboBox, "All carriers");
        var month = SelectedFilterValue(RenewalMonthFilterComboBox, "All months");

        return (string.IsNullOrWhiteSpace(search) || row.Client.Contains(search, StringComparison.OrdinalIgnoreCase)) &&
               (string.IsNullOrWhiteSpace(status) || row.Status.Equals(status, StringComparison.OrdinalIgnoreCase)) &&
               (string.IsNullOrWhiteSpace(carrier) || row.Carrier.Equals(carrier, StringComparison.OrdinalIgnoreCase)) &&
               (string.IsNullOrWhiteSpace(month) || row.RenewalMonth.Equals(month, StringComparison.OrdinalIgnoreCase));
    }

    private static string SelectedFilterValue(ComboBox comboBox, string allLabel)
    {
        var value = comboBox?.SelectedItem?.ToString() ?? "";
        return value.Equals(allLabel, StringComparison.OrdinalIgnoreCase) ? "" : value;
    }

    private bool ClientListFilter(object item)
    {
        if (item is not ClientListItem row)
        {
            return false;
        }

        var search = ClientSearchTextBox?.Text?.Trim() ?? "";
        var status = SelectedFilterValue(ClientStatusFilterComboBox, "All Statuses");
        var industry = SelectedFilterValue(ClientIndustryFilterComboBox, "All Industries");

        return row.MatchesSearch(search) &&
               (string.IsNullOrWhiteSpace(status) || row.Status.Equals(status, StringComparison.OrdinalIgnoreCase)) &&
               (string.IsNullOrWhiteSpace(industry) || row.Industry.Equals(industry, StringComparison.OrdinalIgnoreCase));
    }

    private void HeaderSearchBox_GotFocus(object sender, RoutedEventArgs e)
    {
        if (!_headerSearchShowsPlaceholder)
        {
            return;
        }

        _headerSearchShowsPlaceholder = false;
        HeaderSearchBox.Text = "";
        HeaderSearchBox.Foreground = TryFindResource("TextNavyBrush") as Brush ?? Brushes.Black;
    }

    private void HeaderSearchBox_LostFocus(object sender, RoutedEventArgs e)
    {
        if (!string.IsNullOrWhiteSpace(HeaderSearchBox.Text))
        {
            return;
        }

        _headerSearchShowsPlaceholder = true;
        HeaderSearchBox.Text = HeaderSearchPlaceholder;
        HeaderSearchBox.Foreground = TryFindResource("MutedNavyBrush") as Brush ?? Brushes.Gray;
    }

    private void HeaderSearchBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_headerSearchShowsPlaceholder || ClientSearchTextBox is null)
        {
            return;
        }

        if (ClientsScrollViewer.Visibility != Visibility.Visible || ClientListPanel.Visibility != Visibility.Visible)
        {
            ShowClientsPage();
        }

        var search = HeaderSearchBox.Text.Trim();
        if (!string.Equals(ClientSearchTextBox.Text, search, StringComparison.Ordinal))
        {
            ClientSearchTextBox.Text = search;
        }

        _clientListView?.Refresh();
        UpdateClientListSummary();
        SetStatus(string.IsNullOrWhiteSpace(search)
            ? "Client search cleared."
            : $"Client search: {ClientListShowingTextBlock.Text}");
    }

    private void ClientListFilter_Changed(object sender, RoutedEventArgs e)
    {
        _clientListView?.Refresh();
        UpdateClientListSummary();
    }

    private void ApplyClientFilters_Click(object sender, RoutedEventArgs e)
    {
        _clientListView?.Refresh();
        UpdateClientListSummary();
        SetStatus($"Clients: filters applied. {ClientListShowingTextBlock.Text}");
    }

    private void ClientListGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (ClientListGrid.SelectedItem is not ClientListItem client)
        {
            ClientIndustryEditorComboBox.SelectedItem = null;
            SaveClientIndustryButton.IsEnabled = false;
            DeleteSelectedClientButton.IsEnabled = false;
            return;
        }

        if (client.DatabaseClientId is null)
        {
            ClientIndustryEditorComboBox.SelectedItem = null;
            SaveClientIndustryButton.IsEnabled = false;
            DeleteSelectedClientButton.IsEnabled = false;
            return;
        }

        SaveClientIndustryButton.IsEnabled = true;
        DeleteSelectedClientButton.IsEnabled = true;
        if (!client.Industry.Equals("Not set", StringComparison.OrdinalIgnoreCase))
        {
            RestoreFilterSelection(ClientIndustryEditorComboBox, client.Industry);
        }
        else
        {
            ClientIndustryEditorComboBox.SelectedItem = null;
        }
    }

    private void ClientListGrid_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        var row = ItemsControl.ContainerFromElement(ClientListGrid, e.OriginalSource as DependencyObject) as DataGridRow;
        if (row?.DataContext is ClientListItem client)
        {
            OpenClientWorkspace(client);
        }
    }

    private async void DeleteSelectedClient_Click(object sender, RoutedEventArgs e)
    {
        await DeleteSelectedClientAsync(ClientListGrid.SelectedItem as ClientListItem);
    }

    private async void DeleteClientRow_Click(object sender, RoutedEventArgs e)
    {
        if ((sender as FrameworkElement)?.DataContext is ClientListItem client)
        {
            await DeleteSelectedClientAsync(client);
        }
    }

    private async Task DeleteSelectedClientAsync(ClientListItem? client)
    {
        if (client?.DatabaseClientId is null)
        {
            SetStatus("Only saved SQLite clients can be deleted. Presentation-only rows are not database clients.");
            return;
        }

        var confirmation = System.Windows.MessageBox.Show(
            this,
            $"Permanently delete {client.ClientName}?\n\nThis removes the client and all related renewal, claims, billing, plan-design, document, and history records from the Benefits Intelligence database. This cannot be undone from the app.",
            "Delete Client",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning,
            MessageBoxResult.No);
        if (confirmation != MessageBoxResult.Yes)
        {
            SetStatus($"Deletion cancelled for {client.ClientName}.");
            return;
        }

        DeleteSelectedClientButton.IsEnabled = false;
        try
        {
            var result = await _databaseService.DeleteClientAsync(client.DatabaseClientId.Value);
            if (!result.Succeeded)
            {
                System.Windows.MessageBox.Show(this, result.Message, "Delete Client", MessageBoxButton.OK, MessageBoxImage.Warning);
                SetStatus("Client was not deleted.");
                return;
            }

            await LoadSavedRenewalQueueAsync();
            await LoadSavedClientListAsync();
            await RefreshHomeDashboardAsync();
            SetStatus(result.Message);
        }
        finally
        {
            DeleteSelectedClientButton.IsEnabled = ClientListGrid.SelectedItem is ClientListItem { DatabaseClientId: not null };
        }
    }

    private async void SaveClientIndustry_Click(object sender, RoutedEventArgs e)
    {
        if (ClientListGrid.SelectedItem is not ClientListItem client || client.DatabaseClientId is null)
        {
            SetStatus("Select a saved client before setting its industry.");
            return;
        }

        var industry = ClientIndustryEditorComboBox.SelectedItem?.ToString()?.Trim() ?? "";
        if (string.IsNullOrWhiteSpace(industry))
        {
            SetStatus("Choose an industry before saving.");
            return;
        }

        SaveClientIndustryButton.IsEnabled = false;
        try
        {
            var saveResult = await _databaseService.UpdateClientIndustryAsync(client.DatabaseClientId.Value, industry);
            if (!saveResult.Succeeded)
            {
                System.Windows.MessageBox.Show(this, saveResult.Message, "Industry", MessageBoxButton.OK, MessageBoxImage.Warning);
                SetStatus("Industry was not saved.");
                return;
            }

            await LoadSavedClientListAsync();
            SetStatus($"Saved {industry} for {client.ClientName}.");
        }
        finally
        {
            SaveClientIndustryButton.IsEnabled = true;
        }
    }

    private void UpdateClientListSummary()
    {
        var visible = _clientListView?.Cast<ClientListItem>().Count() ?? _clientList.Count;
        var total = _clientList.Count;
        var active = _clientList.Count(row => HasClientStatus(row, "Active"));
        var inactive = _clientList.Count(row => HasClientStatus(row, "Inactive"));
        var prospects = _clientList.Count(row => HasClientStatus(row, "Prospect"));
        var shown = visible == 0 ? 0 : Math.Min(15, visible);

        ClientTotalKpiTextBlock.Text = total.ToString("#,0", CultureInfo.InvariantCulture);
        ClientTotalKpiSubtitleTextBlock.Text = total == 1 ? "Loaded client" : "Loaded clients";
        ClientActiveKpiTextBlock.Text = active.ToString("#,0", CultureInfo.InvariantCulture);
        ClientActiveKpiSubtitleTextBlock.Text = FormatClientShare(active, total);
        ClientInactiveKpiTextBlock.Text = inactive.ToString("#,0", CultureInfo.InvariantCulture);
        ClientInactiveKpiSubtitleTextBlock.Text = FormatClientShare(inactive, total);
        ClientProspectsKpiTextBlock.Text = prospects.ToString("#,0", CultureInfo.InvariantCulture);
        ClientProspectsKpiSubtitleTextBlock.Text = FormatClientShare(prospects, total);

        ClientListShowingTextBlock.Text = visible == 0
            ? $"Showing 0 of {total.ToString("#,0", CultureInfo.InvariantCulture)} loaded clients"
            : $"Showing 1 to {shown.ToString("#,0", CultureInfo.InvariantCulture)} of {visible.ToString("#,0", CultureInfo.InvariantCulture)} visible clients";
        ClientListFooterTextBlock.Text = visible == 0
            ? $"Showing 0 of {total.ToString("#,0", CultureInfo.InvariantCulture)} loaded results"
            : $"Showing 1 to {shown.ToString("#,0", CultureInfo.InvariantCulture)} of {visible.ToString("#,0", CultureInfo.InvariantCulture)} visible results";
    }

    private static bool HasClientStatus(ClientListItem row, string status)
    {
        if (status.Equals("Prospect", StringComparison.OrdinalIgnoreCase))
        {
            return row.Status.Contains("Prospect", StringComparison.OrdinalIgnoreCase);
        }

        return row.Status.Equals(status, StringComparison.OrdinalIgnoreCase);
    }

    private static string FormatClientShare(int count, int total)
    {
        if (total <= 0)
        {
            return "0% of total";
        }

        var percent = (decimal)count / total;
        return percent.ToString("P0", CultureInfo.InvariantCulture) + " of total";
    }

    private void RenewalQueueFilter_Changed(object sender, RoutedEventArgs e)
    {
        _renewalQueueView?.Refresh();
        UpdateRenewalQueueSummary();

        if (RenewalQueueGrid.SelectedItem is not RenewalQueueItem selected ||
            _renewalQueueView is null ||
            !_renewalQueueView.Cast<RenewalQueueItem>().Contains(selected))
        {
            RenewalQueueGrid.SelectedItem = _renewalQueueView?.Cast<RenewalQueueItem>().FirstOrDefault();
        }
    }

    private void ApplyRenewalQueueFilters_Click(object sender, RoutedEventArgs e)
    {
        _renewalQueueView?.Refresh();
        UpdateRenewalQueueSummary();
        SetStatus($"Renewal queue filters applied. {RenewalQueueCountTextBlock.Text}");
    }

    private void ExportRenewalQueue_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var visibleRows = VisibleRenewalQueueRows();
            if (visibleRows.Count == 0)
            {
                System.Windows.MessageBox.Show(
                    this,
                    "There are no renewal queue rows to export with the current filters.",
                    "Export Renewal Queue",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            var outputFolder = Path.Combine(ResolveOutputRootFolder(), "Renewal Queue Reports");
            Directory.CreateDirectory(outputFolder);
            var workbookPath = Path.Combine(outputFolder, $"Renewal_Queue_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");

            WriteRenewalQueueWorkbook(visibleRows, workbookPath);
            _renewalQueueExportFolder = outputFolder;
            ExportTextBlock.Text =
                $"Renewal Queue Excel export:{Environment.NewLine}{workbookPath}{Environment.NewLine}{Environment.NewLine}" +
                $"Rows exported: {visibleRows.Count}{Environment.NewLine}" +
                $"Filters: {BuildRenewalQueueFilterSummary()}";
            SetOpenOutputButtonsEnabled(true);
            SetStatus($"Renewal queue exported to Excel with {visibleRows.Count} visible row(s).");

            System.Windows.MessageBox.Show(
                this,
                $"Renewal Queue Excel export was created.{Environment.NewLine}{Environment.NewLine}{workbookPath}",
                "Export Renewal Queue",
                MessageBoxButton.OK,
                MessageBoxImage.Information);

            Process.Start(new ProcessStartInfo
            {
                FileName = outputFolder,
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            System.Windows.MessageBox.Show(this, ex.Message, "Export Renewal Queue failed", MessageBoxButton.OK, MessageBoxImage.Error);
            SetStatus("Renewal queue export failed.");
        }
    }

    private IReadOnlyList<RenewalQueueItem> VisibleRenewalQueueRows()
    {
        return _renewalQueueView?.Cast<RenewalQueueItem>().ToList() ?? _renewalQueue.ToList();
    }

    private void WriteRenewalQueueWorkbook(IReadOnlyList<RenewalQueueItem> rows, string workbookPath)
    {
        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Renewal Queue");
        const int columnCount = 12;

        worksheet.Cell(1, 1).Value = "Renewal Queue for Consultant Review";
        var titleRange = worksheet.Range(1, 1, 1, columnCount);
        titleRange.Merge();
        titleRange.Style.Font.Bold = true;
        titleRange.Style.Font.FontSize = 16;
        titleRange.Style.Font.FontColor = XLColor.FromHtml("#071B4D");

        worksheet.Cell(2, 1).Value = "Generated";
        worksheet.Cell(2, 2).Value = DateTime.Now;
        worksheet.Cell(2, 2).Style.DateFormat.Format = "mmm d yyyy h:mm AM/PM";
        worksheet.Cell(3, 1).Value = "Filters";
        worksheet.Cell(3, 2).Value = BuildRenewalQueueFilterSummary();
        worksheet.Range(2, 1, 3, 1).Style.Font.Bold = true;

        var headers = new[]
        {
            "Client",
            "Carrier",
            "Renewal Date",
            "Current Premium",
            "Renewal Premium",
            "Premium Change %",
            "EHC Loss Ratio",
            "Pooling Threshold",
            "Pooling Charge",
            "Dental Loss Ratio",
            "Status",
            "Next Step"
        };

        const int headerRow = 5;
        for (var column = 0; column < headers.Length; column++)
        {
            worksheet.Cell(headerRow, column + 1).Value = headers[column];
        }

        var currentRow = headerRow + 1;
        foreach (var row in rows)
        {
            worksheet.Cell(currentRow, 1).Value = row.Client;
            worksheet.Cell(currentRow, 2).Value = row.Carrier;
            if (row.RenewalDate.HasValue)
            {
                worksheet.Cell(currentRow, 3).Value = row.RenewalDate.Value.ToDateTime(TimeOnly.MinValue);
                worksheet.Cell(currentRow, 3).Style.DateFormat.Format = "mmm d yyyy";
            }
            else
            {
                worksheet.Cell(currentRow, 3).Value = "N/A";
            }

            SetMoneyCell(worksheet.Cell(currentRow, 4), row.CurrentPremium);
            SetMoneyCell(worksheet.Cell(currentRow, 5), row.RenewalPremium);
            SetPercentCell(worksheet.Cell(currentRow, 6), row.PremiumChangePercent, includeSign: true);
            SetPercentCell(worksheet.Cell(currentRow, 7), row.EhcLossRatio, includeSign: false);
            SetMoneyCell(worksheet.Cell(currentRow, 8), row.PoolingThreshold);
            SetNormalizedPercentCell(worksheet.Cell(currentRow, 9), row.PoolingCharge, includeSign: false);
            SetPercentCell(worksheet.Cell(currentRow, 10), row.DentalLossRatio, includeSign: false);
            worksheet.Cell(currentRow, 11).Value = row.Status;
            worksheet.Cell(currentRow, 11).Style.Fill.BackgroundColor = RenewalStatusExportColor(row.Status);
            worksheet.Cell(currentRow, 12).Value = row.ActionLabel;
            currentRow++;
        }

        var tableRange = worksheet.Range(headerRow, 1, currentRow - 1, columnCount);
        var table = tableRange.CreateTable("RenewalQueue");
        table.Theme = XLTableTheme.TableStyleMedium2;

        var headerRange = worksheet.Range(headerRow, 1, headerRow, columnCount);
        headerRange.Style.Font.Bold = true;
        headerRange.Style.Fill.BackgroundColor = XLColor.FromHtml("#061B3A");
        headerRange.Style.Font.FontColor = XLColor.White;

        worksheet.Range(1, 1, currentRow - 1, columnCount).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        worksheet.SheetView.FreezeRows(headerRow);
        worksheet.Columns().AdjustToContents();
        foreach (var column in worksheet.ColumnsUsed())
        {
            if (column.Width > 55)
            {
                column.Width = 55;
            }
        }

        workbook.SaveAs(workbookPath);
    }

    private string BuildRenewalQueueFilterSummary()
    {
        var search = RenewalQueueSearchTextBox?.Text?.Trim();
        var status = SelectedFilterValue(RenewalStatusFilterComboBox, "All statuses");
        var carrier = SelectedFilterValue(RenewalCarrierFilterComboBox, "All carriers");
        var month = SelectedFilterValue(RenewalMonthFilterComboBox, "All months");

        return string.Join(" | ", new[]
        {
            $"Search: {FirstNonBlank(search, "All")}",
            $"Status: {FirstNonBlank(status, "All")}",
            $"Carrier: {FirstNonBlank(carrier, "All")}",
            $"Month: {FirstNonBlank(month, "All")}"
        });
    }

    private static void SetMoneyCell(IXLCell cell, decimal? value)
    {
        if (!value.HasValue)
        {
            cell.Value = "N/A";
            return;
        }

        cell.Value = value.Value;
        cell.Style.NumberFormat.Format = "$#,##0";
    }

    private static void SetPercentCell(IXLCell cell, decimal? value, bool includeSign)
    {
        if (!value.HasValue)
        {
            cell.Value = "N/A";
            return;
        }

        cell.Value = value.Value / 100m;
        cell.Style.NumberFormat.Format = includeSign ? "+0.0%;-0.0%;0.0%" : "0.0%";
    }

    private static void SetNormalizedPercentCell(IXLCell cell, decimal? value, bool includeSign)
    {
        if (!value.HasValue)
        {
            cell.Value = "N/A";
            return;
        }

        var percent = NormalizePercent(value);
        if (!percent.HasValue)
        {
            cell.Value = "N/A";
            return;
        }

        cell.Value = percent.Value / 100m;
        cell.Style.NumberFormat.Format = includeSign ? "+0.0%;-0.0%;0.0%" : "0.0%";
    }

    private static XLColor RenewalStatusExportColor(string status)
    {
        return status switch
        {
            RenewalQueueStatuses.Healthy => XLColor.FromHtml("#E7F5ED"),
            RenewalQueueStatuses.Watch => XLColor.FromHtml("#FFF4D8"),
            RenewalQueueStatuses.HighRisk => XLColor.FromHtml("#FDE8E8"),
            RenewalQueueStatuses.MissingData => XLColor.FromHtml("#EEF1F4"),
            RenewalQueueStatuses.ReadyToReview => XLColor.FromHtml("#E7F1FB"),
            RenewalQueueStatuses.Negotiated => XLColor.FromHtml("#E8F1FF"),
            RenewalQueueStatuses.FinalApproved => XLColor.FromHtml("#D9EDE1"),
            RenewalQueueStatuses.Completed => XLColor.FromHtml("#DFF1E5"),
            _ => XLColor.White
        };
    }

    private void RenewalQueueGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        UpdateSelectedRenewalDetail(RenewalQueueGrid.SelectedItem as RenewalQueueItem);
    }

    private async void RenewalQueueAction_Click(object sender, RoutedEventArgs e)
    {
        if ((sender as FrameworkElement)?.DataContext is not RenewalQueueItem row)
        {
            return;
        }

        RenewalQueueGrid.SelectedItem = row;
        switch (row.ActionLabel)
        {
            case "Needs Data":
                SetStatus($"Needs Data: review the checklist for {row.Client} before exporting.");
                break;
            case "Export Pack" when row.Id.Equals("current-renewal", StringComparison.OrdinalIgnoreCase) &&
                                    IsCurrentAnalysisContext(row):
                Export_Click(sender, e);
                break;
            case "Export Pack":
                await OpenNegotiationWorkbenchAsync(row);
                NegotiationTabs.SelectedIndex = 4;
                NegotiationExportStatusTextBlock.Text =
                    "Choose the saved renewal version to regenerate the workbook and PowerPoint from Benefits Intelligence.";
                SetStatus($"Export Pack: opened {row.Client}'s saved renewal export workspace.");
                break;
            case "Validate":
                await ShowUnderwritingDashboardAsync(row);
                break;
            default:
                await ShowUnderwritingDashboardAsync(row);
                break;
        }
    }

    private async Task OpenNegotiationWorkbenchAsync(RenewalQueueItem row)
    {
        if (_clientRenewalSelectionIdentity is not null &&
            !ClientRenewalSelectionResolver.Matches(
                _clientRenewalSelectionIdentity,
                CreateRenewalSelectionCandidate(row, 0)))
        {
            SetStatus($"Negotiation was not opened because the selected renewal does not belong to {_clientRenewalSelectionIdentity.ClientName}.");
            return;
        }

        _selectedNegotiationRenewal = row;
        HomeDashboardScrollViewer.Visibility = Visibility.Collapsed;
        DashboardScrollViewer.Visibility = Visibility.Collapsed;
        UnderwritingWorkbenchScrollViewer.Visibility = Visibility.Collapsed;
        AssumptionLibraryScrollViewer.Visibility = Visibility.Collapsed;
        ReportsAnalyticsScrollViewer.Visibility = Visibility.Collapsed;
        BenchmarkingDashboardScrollViewer.Visibility = Visibility.Collapsed;
        ProposalComparisonScrollViewer.Visibility = Visibility.Collapsed;
        PlanDesignIntelligenceScrollViewer.Visibility = Visibility.Collapsed;
        PlanDesignFactorsScrollViewer.Visibility = Visibility.Collapsed;
        DocumentIntelligenceScrollViewer.Visibility = Visibility.Collapsed;
        ClientsScrollViewer.Visibility = Visibility.Collapsed;
        NegotiationWorkbenchScrollViewer.Visibility = Visibility.Visible;
        NegotiationTabs.SelectedIndex = 3;
        SetOperationalFooterVisible(true);

        NegotiationClientNameTextBlock.Text = row.Client;
        NegotiationCarrierTextBlock.Text = row.Carrier;
        NegotiationRenewalTitleTextBlock.Text = $"{row.RenewalDateText} Renewal";
        ApplyNegotiationStatus(row.Status);

        OverviewClientInfoTextBlock.Text =
            $"Client: {row.Client}{Environment.NewLine}" +
            $"Carrier: {row.Carrier}{Environment.NewLine}" +
            $"Renewal date: {row.RenewalDateText}{Environment.NewLine}" +
            $"Status: {row.Status}";
        OverviewMetricsTextBlock.Text =
            $"Current premium: {row.CurrentPremiumText}{Environment.NewLine}" +
            $"Proposed premium: {row.RenewalPremiumText}{Environment.NewLine}" +
            $"Premium change: {row.PremiumChangeText}{Environment.NewLine}" +
            $"EHC loss ratio: {row.EhcLossRatioText}{Environment.NewLine}" +
            $"Dental loss ratio: {row.DentalLossRatioText}{Environment.NewLine}" +
            $"Lives: {row.LivesText}";
        NegotiationQuickInsightsItemsControl.ItemsSource = row.QuickInsights.Select(insight => "- " + insight).ToList();
        NegotiationMissingDataItemsControl.ItemsSource = row.MissingDataItems.Count == 0
            ? ["- No required data gaps identified."]
            : row.MissingDataItems.Select(item => "- " + item).ToList();

        _negotiationRates.Clear();
        ResetNetGrossRateScheduleReview();
        ResetRevisedRateScheduleStatus();
        var negotiationRows = await BuildNegotiationRowsAsync(row);
        await ApplyUnderwritingRiskAsync(row, negotiationRows);
        foreach (var rate in BenefitDisplayOrder.Sort(
                     negotiationRows,
                     item => item.BenefitLine,
                     item => item.ClassName,
                     item => item.AccountClass))
        {
            _negotiationRates.Add(rate);
        }

        RefreshNetGrossRateScheduleReview();

        await LoadNegotiationClaimsRowsAsync(row);
        PopulateAssumptionsFromRenewal(row);
        RecalculateNegotiation();
        SetStatus($"Negotiation: opened final rate and export workspace for {row.Client}.");
    }

    private async Task ApplyUnderwritingRiskAsync(RenewalQueueItem renewal, IReadOnlyCollection<NegotiationRateRow> rows)
    {
        if (!IsCurrentAnalysisContext(renewal) || _analysis is null || rows.Count == 0)
        {
            return;
        }

        IReadOnlyList<UnderwritingReferenceRate> referenceRates;
        try
        {
            referenceRates = await _databaseService.GetActiveUnderwritingReferenceRatesAsync();
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            referenceRates = UnderwritingReferenceRateCatalog.Defaults;
        }

        var mortalityChange = UnderwritingRiskCalculator.CalculateChange(
            _analysis.Pdf.LifeDemographics,
            referenceRates,
            UnderwritingRateTypes.Mortality);
        var morbidityChange = UnderwritingRiskCalculator.CalculateChange(
            _analysis.Pdf.LtdDemographics,
            referenceRates,
            UnderwritingRateTypes.Morbidity);

        foreach (var row in rows)
        {
            var rateType = UnderwritingRiskCalculator.RateTypeForBenefit(row.BenefitLine);
            var changeInRisk = rateType switch
            {
                UnderwritingRateTypes.Mortality => mortalityChange,
                UnderwritingRateTypes.Morbidity => morbidityChange,
                _ => null
            };
            row.ChangeInRisk = changeInRisk;
            row.CalculatedRate = UnderwritingRiskCalculator.CalculateRate(row.CurrentRate, changeInRisk);
            row.CalculatedPremium = UnderwritingRiskCalculator.CalculatePremium(rateType, row.Volume, row.CalculatedRate);
        }
    }

    private async Task<List<NegotiationRateRow>> BuildNegotiationRowsAsync(RenewalQueueItem row)
    {
        if (IsCurrentAnalysisContext(row) || row.ClientId <= 0)
        {
            return BuildNegotiationRows(row);
        }

        try
        {
            var history = await _databaseService.GetClientHistoryAsync(row.ClientId);
            var rows = BuildNegotiationRowsFromHistory(row, history);
            if (RevisedRateNegotiationRecovery.NeedsProposalRecovery(rows))
            {
                var recoveredRows = RecoverDetailedNegotiationRows(row, history);
                if (recoveredRows.Count > 0)
                {
                    return recoveredRows;
                }
            }

            if (rows.Count > 0)
            {
                return rows;
            }
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
        }

        return BuildNegotiationRows(row);
    }

    private List<NegotiationRateRow> RecoverDetailedNegotiationRows(RenewalQueueItem row, ClientHistory history)
    {
        var supportedExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { ".pdf", ".xlsx", ".xlsm" };
        var candidates = history.Documents
            .Where(document => !string.IsNullOrWhiteSpace(document.FilePath) &&
                               supportedExtensions.Contains(Path.GetExtension(document.FilePath)) &&
                               File.Exists(document.FilePath))
            .OrderByDescending(document => row.SourceDocumentId.HasValue && document.DocumentId == row.SourceDocumentId)
            .ThenByDescending(document => document.DocumentType.Equals(DocumentTypes.Renewal, StringComparison.OrdinalIgnoreCase))
            .ThenByDescending(document => row.RenewalDate.HasValue && document.EffectiveDate == row.RenewalDate)
            .ThenByDescending(document => document.FileName.Contains("renewal", StringComparison.OrdinalIgnoreCase))
            .ThenByDescending(document => document.ProcessedDate)
            .ToList();

        foreach (var document in candidates)
        {
            try
            {
                var import = _revisedRateScheduleImporter.Import(document.FilePath);
                if (!import.Succeeded || import.Schedule is null || import.Schedule.Rows.Count < 2)
                {
                    continue;
                }

                var schedule = import.Schedule;
                if ((!string.IsNullOrWhiteSpace(schedule.ClientName) &&
                     !RevisedRateScheduleImporter.ClientMatches(schedule.ClientName, row.Client)) ||
                    (!string.IsNullOrWhiteSpace(schedule.Carrier) &&
                     !RevisedRateScheduleImporter.CarrierMatches(schedule.Carrier, row.Carrier)) ||
                    (schedule.RenewalDate.HasValue && row.RenewalDate.HasValue && schedule.RenewalDate != row.RenewalDate))
                {
                    continue;
                }

                var recovered = RevisedRateNegotiationRecovery.BuildProposalRows(
                    schedule,
                    row.CurrentPremium,
                    row.RenewalPremium);
                foreach (var rate in recovered)
                {
                    rate.CurrentClaims = EstimateBenefitClaims(row, rate.BenefitLine, rate.CurrentPremium);
                }

                if (recovered.Count > 0)
                {
                    return recovered;
                }
            }
            catch (Exception ex)
            {
                AppLogger.Error(ex);
            }
        }

        return [];
    }

    private List<NegotiationRateRow> BuildNegotiationRows(RenewalQueueItem row)
    {
        if (IsCurrentAnalysisContext(row) && _analysis is not null)
        {
            if (_analysis.RenewalVersions.Count == 0)
            {
                new RenewalEngine().BuildRenewalVersions(_analysis);
            }

            var carrier = FindRenewalVersion(_analysis, RenewalVersionNames.CarrierProposed);
            var model = FindRenewalVersion(_analysis, RenewalVersionNames.ConsultantModel);
            var negotiated = FindRenewalVersion(_analysis, RenewalVersionNames.Negotiated) ?? model ?? carrier;
            var sourceLines = negotiated?.RateLines.Count > 0
                ? negotiated.RateLines
                : carrier?.RateLines ?? [];

            var rows = new List<NegotiationRateRow>();
            foreach (var sourceLine in sourceLines)
            {
                var carrierLine = FindMatchingVersionLine(carrier?.RateLines, sourceLine);
                var modelLine = FindMatchingVersionLine(model?.RateLines, sourceLine);
                var negotiatedLine = FindMatchingVersionLine(negotiated?.RateLines, sourceLine);
                var selectedValueLine = negotiatedLine ?? sourceLine;
                var useSelectedStageRates = UsesExplicitRateValueBasis(selectedValueLine.RateValueBasis);

                rows.Add(new NegotiationRateRow
                {
                    BenefitLine = FirstNonBlank(sourceLine.Benefit, "Overall"),
                    ClassName = FirstNonBlank(sourceLine.TierOrClass, "All"),
                    AccountClass = FirstNonBlank(sourceLine.AccountClass, "All Classes"),
                    RateBasis = FirstNonBlank(sourceLine.RateBasis, "composite"),
                    Volume = sourceLine.Volume,
                    Lives = sourceLine.Lives,
                    CurrentRate = useSelectedStageRates
                        ? selectedValueLine.CurrentRate
                        : sourceLine.CurrentRate,
                    ProposedRate = selectedValueLine.ProposedRate
                        ?? carrierLine?.VersionRate ?? carrierLine?.ProposedRate,
                    ModelledRate = modelLine?.VersionRate ?? sourceLine.VersionRate,
                    NegotiatedRate = useSelectedStageRates
                        ? selectedValueLine.VersionRate
                        : negotiatedLine?.VersionRate ?? carrierLine?.VersionRate ?? sourceLine.VersionRate,
                    CurrentPremium = sourceLine.CurrentPremium,
                    ProposedPremium = selectedValueLine.ProposedPremium
                        ?? carrierLine?.VersionPremium ?? carrierLine?.ProposedPremium,
                    ModelledPremium = modelLine?.VersionPremium ?? sourceLine.VersionPremium,
                    NegotiatedPremium = negotiatedLine?.VersionPremium ?? carrierLine?.VersionPremium ?? sourceLine.VersionPremium,
                    RateValueBasis = FirstNonBlank(negotiatedLine?.RateValueBasis, sourceLine.RateValueBasis, RateValueBasisNames.CarrierStated),
                    ClientPresentationRateBasis = FirstNonBlank(negotiatedLine?.ClientPresentationRateBasis, sourceLine.ClientPresentationRateBasis, RateValueBasisNames.CarrierStated),
                    CurrentNetRate = negotiatedLine?.CurrentNetRate ?? sourceLine.CurrentNetRate,
                    CurrentGrossRate = negotiatedLine?.CurrentGrossRate ?? sourceLine.CurrentGrossRate,
                    ProposedNetRate = negotiatedLine?.ProposedNetRate ?? carrierLine?.ProposedNetRate ?? sourceLine.ProposedNetRate,
                    ProposedGrossRate = negotiatedLine?.ProposedGrossRate ?? carrierLine?.ProposedGrossRate ?? sourceLine.ProposedGrossRate,
                    VersionNetRate = negotiatedLine?.VersionNetRate ?? sourceLine.VersionNetRate,
                    VersionGrossRate = negotiatedLine?.VersionGrossRate ?? sourceLine.VersionGrossRate,
                    RateScheduleSourceFile = FirstNonBlank(negotiatedLine?.RateScheduleSourceFile, sourceLine.RateScheduleSourceFile),
                    RateScheduleSourceSheet = FirstNonBlank(negotiatedLine?.RateScheduleSourceSheet, sourceLine.RateScheduleSourceSheet),
                    RateScheduleSourceRow = negotiatedLine?.RateScheduleSourceRow ?? sourceLine.RateScheduleSourceRow,
                    SourceReference = FirstNonBlank(negotiatedLine?.SourceReference, sourceLine.SourceReference),
                    UseSavedModelValues = modelLine is not null,
                    Status = RenewalQueueStatuses.Negotiated,
                    Notes = ""
                });
            }

            AllocateBenefitClaimsToRateRows(_analysis, rows);
            return rows.Count > 0 ? rows : BuildFallbackNegotiationRows(row, isCurrentAnalysis: true);
        }

        return BuildFallbackNegotiationRows(row);
    }

    internal static List<NegotiationRateRow> BuildNegotiationRowsFromHistory(RenewalQueueItem row, ClientHistory history)
    {
        var versions = history.RenewalVersions
            .Where(version => VersionMatchesQueueRow(version, row))
            .ToList();
        if (versions.Count == 0)
        {
            versions = history.RenewalVersions;
        }

        var carrier = LatestVersion(versions, RenewalVersionNames.CarrierProposed);
        var model = LatestVersion(versions, RenewalVersionNames.ConsultantModel);
        var negotiated = LatestVersion(versions, RenewalVersionNames.Negotiated)
                         ?? LatestVersion(versions, RenewalVersionNames.FinalApproved);
        var sourceLines = MostDetailedRateRows(carrier, model, negotiated);
        if (sourceLines.Count == 0)
        {
            return [];
        }

        var sourceScale = PremiumScaleForQueue(sourceLines.SumValue(line => line.CurrentPremium), row.CurrentPremium);
        var carrierScale = PremiumScaleForVersion(carrier, row, sourceScale);
        var modelScale = PremiumScaleForVersion(model, row, sourceScale);
        var negotiatedScale = PremiumScaleForVersion(negotiated, row, sourceScale);
        var rows = new List<NegotiationRateRow>();

        foreach (var sourceLine in sourceLines)
        {
            var carrierLine = FindMatchingVersionLine(carrier?.Rates, sourceLine, allowBenefitOnly: false)
                              ?? FindMatchingVersionLine(carrier?.Rates, sourceLine, allowBenefitOnly: true);
            var modelLine = FindMatchingVersionLine(model?.Rates, sourceLine, allowBenefitOnly: false);
            var negotiatedLine = FindMatchingVersionLine(negotiated?.Rates, sourceLine, allowBenefitOnly: false);
            var selectedValueLine = negotiatedLine ?? sourceLine;
            var useSelectedStageRates = UsesExplicitRateValueBasis(selectedValueLine.RateValueBasis);

            var currentPremium = negotiatedLine?.CurrentPremium.HasValue == true
                ? ScalePremium(negotiatedLine.CurrentPremium, negotiatedScale)
                : ScalePremium(carrierLine?.CurrentPremium ?? sourceLine.CurrentPremium, carrierLine is not null ? carrierScale : sourceScale);
            var proposedPremium = ScalePremium(
                carrierLine?.VersionPremium ?? carrierLine?.ProposedPremium ?? sourceLine.ProposedPremium ?? sourceLine.VersionPremium,
                carrierLine is not null ? carrierScale : sourceScale);
            if (negotiatedLine?.ProposedPremium.HasValue == true)
            {
                proposedPremium = ScalePremium(negotiatedLine.ProposedPremium, negotiatedScale);
            }
            var modelledPremium = ScalePremium(modelLine?.VersionPremium, modelScale);
            var negotiatedPremium = ScalePremium(negotiatedLine?.VersionPremium, negotiatedScale) ?? proposedPremium;

            rows.Add(new NegotiationRateRow
            {
                BenefitLine = FirstNonBlank(sourceLine.BenefitLine, "Overall"),
                ClassName = FirstNonBlank(sourceLine.ClassName, "All"),
                AccountClass = FirstNonBlank(sourceLine.AccountClass, "All Classes"),
                RateBasis = FirstNonBlank(sourceLine.RateBasis, "composite"),
                Volume = sourceLine.Volume,
                Lives = sourceLine.Lives,
                CurrentClaims = EstimateBenefitClaims(row, sourceLine.BenefitLine, currentPremium),
                CurrentRate = negotiatedLine?.CurrentRate ?? (useSelectedStageRates
                    ? selectedValueLine.CurrentRate
                    : carrierLine?.CurrentRate ?? sourceLine.CurrentRate),
                ProposedRate = negotiatedLine?.ProposedRate ?? (useSelectedStageRates
                    ? selectedValueLine.ProposedRate
                    : carrierLine?.VersionRate ?? carrierLine?.ProposedRate ?? sourceLine.ProposedRate ?? sourceLine.VersionRate),
                ModelledRate = modelLine?.VersionRate,
                NegotiatedRate = useSelectedStageRates
                    ? selectedValueLine.VersionRate
                    : negotiatedLine?.VersionRate ?? carrierLine?.VersionRate ?? carrierLine?.ProposedRate,
                CurrentPremium = currentPremium,
                ProposedPremium = proposedPremium,
                ModelledPremium = modelledPremium,
                NegotiatedPremium = negotiatedPremium,
                RateValueBasis = FirstNonBlank(negotiatedLine?.RateValueBasis, sourceLine.RateValueBasis, RateValueBasisNames.CarrierStated),
                ClientPresentationRateBasis = FirstNonBlank(negotiatedLine?.ClientPresentationRateBasis, sourceLine.ClientPresentationRateBasis, RateValueBasisNames.CarrierStated),
                CurrentNetRate = negotiatedLine?.CurrentNetRate ?? sourceLine.CurrentNetRate,
                CurrentGrossRate = negotiatedLine?.CurrentGrossRate ?? sourceLine.CurrentGrossRate,
                ProposedNetRate = negotiatedLine?.ProposedNetRate ?? carrierLine?.ProposedNetRate ?? sourceLine.ProposedNetRate,
                ProposedGrossRate = negotiatedLine?.ProposedGrossRate ?? carrierLine?.ProposedGrossRate ?? sourceLine.ProposedGrossRate,
                VersionNetRate = negotiatedLine?.VersionNetRate ?? sourceLine.VersionNetRate,
                VersionGrossRate = negotiatedLine?.VersionGrossRate ?? sourceLine.VersionGrossRate,
                RateScheduleSourceFile = FirstNonBlank(negotiatedLine?.RateScheduleSourceFile, sourceLine.RateScheduleSourceFile),
                RateScheduleSourceSheet = FirstNonBlank(negotiatedLine?.RateScheduleSourceSheet, sourceLine.RateScheduleSourceSheet),
                RateScheduleSourceRow = negotiatedLine?.RateScheduleSourceRow ?? sourceLine.RateScheduleSourceRow,
                SourceReference = FirstNonBlank(negotiatedLine?.SourceReference, sourceLine.SourceReference),
                UseSavedModelValues = modelLine is not null,
                Status = RenewalQueueStatuses.Negotiated,
                Notes = "Loaded from saved renewal version history."
            });
        }

        return rows;
    }

    private static bool UsesExplicitRateValueBasis(string? rateValueBasis)
    {
        return rateValueBasis?.Equals(RateValueBasisNames.Net, StringComparison.OrdinalIgnoreCase) == true ||
               rateValueBasis?.Equals(RateValueBasisNames.Gross, StringComparison.OrdinalIgnoreCase) == true;
    }

    private static RenewalVersion? FindRenewalVersion(AnalysisResult analysis, string versionName)
    {
        return analysis.RenewalVersions.FirstOrDefault(version =>
            version.VersionName.Equals(versionName, StringComparison.OrdinalIgnoreCase));
    }

    private static RenewalVersionRateLine? FindMatchingVersionLine(IEnumerable<RenewalVersionRateLine>? lines, RenewalVersionRateLine source)
    {
        if (lines is null)
        {
            return null;
        }

        var sourceKey = NormalizeMatchKey(source.Benefit, source.TierOrClass, source.AccountClass);
        return lines.FirstOrDefault(line => NormalizeMatchKey(line.Benefit, line.TierOrClass, line.AccountClass) == sourceKey)
               ?? lines.FirstOrDefault(line => NormalizeMatchKey(line.Benefit, "", "") == NormalizeMatchKey(source.Benefit, "", ""));
    }

    private static RenewalRateVersionRecord? FindMatchingVersionLine(
        IEnumerable<RenewalRateVersionRecord>? lines,
        RenewalRateVersionRecord source,
        bool allowBenefitOnly)
    {
        if (lines is null)
        {
            return null;
        }

        var sourceKey = NormalizeMatchKey(source.BenefitLine, source.ClassName, source.AccountClass);
        var exactMatch = lines.FirstOrDefault(line => NormalizeMatchKey(line.BenefitLine, line.ClassName, line.AccountClass) == sourceKey);
        if (exactMatch is not null || !allowBenefitOnly)
        {
            return exactMatch;
        }

        return lines.FirstOrDefault(line =>
            NormalizeMatchKey(line.BenefitLine, "", "") == NormalizeMatchKey(source.BenefitLine, "", ""));
    }

    private static bool VersionMatchesQueueRow(RenewalVersionRecord version, RenewalQueueItem row)
    {
        if (row.SourceDocumentId.HasValue && version.SourceDocumentId == row.SourceDocumentId)
        {
            return true;
        }

        return row.RenewalDate.HasValue && version.EffectiveDate == row.RenewalDate;
    }

    private static IReadOnlyList<RenewalRateVersionRecord> MostDetailedRateRows(params RenewalVersionRecord?[] versions)
    {
        return versions
            .Select((version, index) => new
            {
                Rows = version?.Rates ?? [],
                Index = index
            })
            .Where(item => item.Rows.Count > 0)
            .OrderByDescending(item => item.Rows.Count + item.Rows.Count(HasSpecificClass))
            .ThenBy(item => item.Index)
            .Select(item => (IReadOnlyList<RenewalRateVersionRecord>)item.Rows)
            .FirstOrDefault() ?? [];
    }

    private static bool HasSpecificClass(RenewalRateVersionRecord line)
    {
        var className = FirstNonBlank(line.AccountClass, line.ClassName);
        return !string.IsNullOrWhiteSpace(className) &&
               !className.Equals("All", StringComparison.OrdinalIgnoreCase) &&
               !className.Equals("All Classes", StringComparison.OrdinalIgnoreCase) &&
               !className.Equals("Composite", StringComparison.OrdinalIgnoreCase);
    }

    private static decimal PremiumScaleForVersion(RenewalVersionRecord? version, RenewalQueueItem row, decimal fallback)
    {
        if (version is null || version.Rates.Count == 0)
        {
            return fallback;
        }

        return PremiumScaleForQueue(version.Rates.SumValue(line => line.CurrentPremium), row.CurrentPremium);
    }

    private static decimal PremiumScaleForQueue(decimal? sourceCurrentTotal, decimal? queueCurrentTotal)
    {
        if (sourceCurrentTotal is not > 0m || queueCurrentTotal is not > 0m)
        {
            return 1m;
        }

        var ratio = queueCurrentTotal.Value / sourceCurrentTotal.Value;
        return ratio is > 10m and < 14m ? 12m : 1m;
    }

    private static decimal? ScalePremium(decimal? premium, decimal scale)
    {
        return premium.HasValue ? decimal.Round(premium.Value * scale, 2) : null;
    }

    private static decimal? EstimateBenefitClaims(RenewalQueueItem row, string benefitLine, decimal? currentPremium)
    {
        if (!currentPremium.HasValue)
        {
            return null;
        }

        var normalizedBenefit = NormalizeMatchKey(benefitLine);
        var lossRatio = normalizedBenefit.Contains("dental", StringComparison.Ordinal)
            ? row.DentalLossRatio
            : normalizedBenefit.Contains("health", StringComparison.Ordinal) ||
              normalizedBenefit.Contains("ehc", StringComparison.Ordinal) ||
              normalizedBenefit.Contains("extended", StringComparison.Ordinal)
                ? row.EhcLossRatio
                : null;

        return lossRatio.HasValue ? decimal.Round(currentPremium.Value * lossRatio.Value / 100m, 2) : null;
    }

    private static string NormalizeMatchKey(params string[] values)
    {
        return new string(string.Join(" ", values)
            .Where(char.IsLetterOrDigit)
            .Select(char.ToLowerInvariant)
            .ToArray());
    }

    private static void AllocateBenefitClaimsToRateRows(AnalysisResult analysis, List<NegotiationRateRow> rows)
    {
        foreach (var benefitGroup in rows.GroupBy(row => NormalizeMatchKey(row.BenefitLine)))
        {
            var benefitRows = benefitGroup.ToList();
            if (benefitRows.Count == 0)
            {
                continue;
            }

            var experience = ExperienceLineSelector.FindBestMatch(analysis.Pdf.ExperienceLines, benefitRows[0].BenefitLine);
            var lossRatio = UnderwritingEngine.NormalizeRate(experience?.PaidLossRatio ?? experience?.IncurredLossRatio);
            if (lossRatio.HasValue)
            {
                foreach (var benefitRow in benefitRows)
                {
                    benefitRow.CurrentClaims = benefitRow.CurrentPremium.HasValue
                        ? decimal.Round(benefitRow.CurrentPremium.Value * lossRatio.Value, 2)
                        : null;
                }

                continue;
            }

            var claims = experience?.PaidClaims ?? experience?.IncurredClaims;
            if (!claims.HasValue)
            {
                continue;
            }

            if (benefitRows.Count == 1)
            {
                benefitRows[0].CurrentClaims = claims;
                continue;
            }

            var currentPremiumTotal = benefitRows.SumValue(row => row.CurrentPremium);
            if (currentPremiumTotal is not > 0m)
            {
                continue;
            }

            foreach (var benefitRow in benefitRows)
            {
                benefitRow.CurrentClaims = benefitRow.CurrentPremium is > 0m
                    ? decimal.Round(claims.Value * benefitRow.CurrentPremium.Value / currentPremiumTotal.Value, 2)
                    : null;
            }
        }
    }

    private static List<NegotiationRateRow> BuildFallbackNegotiationRows(
        RenewalQueueItem row,
        bool isCurrentAnalysis = false)
    {
        var currentTotal = row.CurrentPremium ?? 0m;
        var proposedTotal = row.RenewalPremium ?? currentTotal;
        if (currentTotal <= 0m && proposedTotal <= 0m)
        {
            return
            [
                new NegotiationRateRow
                {
                    BenefitLine = "Overall",
                    ClassName = "All",
                    AccountClass = "All Classes",
                    RateBasis = "composite",
                    CurrentPremium = row.CurrentPremium,
                    ProposedPremium = row.RenewalPremium,
                    ModelledPremium = row.RenewalPremium,
                    NegotiatedPremium = row.RenewalPremium,
                    Lives = row.Lives,
                    Status = RenewalQueueStatuses.MissingData,
                    Notes = "Load renewal files to replace this placeholder with parsed benefit lines."
                }
            ];
        }

        var benefits = new[]
        {
            new { Name = "EHC", Share = 0.48m, LossRatio = row.EhcLossRatio },
            new { Name = "Dental", Share = 0.24m, LossRatio = row.DentalLossRatio },
            new { Name = "Life / AD&D", Share = 0.14m, LossRatio = (decimal?)null },
            new { Name = "Disability", Share = 0.14m, LossRatio = (decimal?)null }
        };

        return benefits.Select(item =>
        {
            var current = currentTotal <= 0m ? (decimal?)null : decimal.Round(currentTotal * item.Share, 2);
            var proposed = proposedTotal <= 0m ? (decimal?)null : decimal.Round(proposedTotal * item.Share, 2);
            var modelled = proposed.HasValue ? decimal.Round(proposed.Value * 0.96m, 2) : proposed;
            return new NegotiationRateRow
            {
                BenefitLine = item.Name,
                ClassName = "All",
                AccountClass = "All Classes",
                RateBasis = "composite",
                Lives = row.Lives,
                CurrentClaims = current.HasValue && item.LossRatio.HasValue ? decimal.Round(current.Value * item.LossRatio.Value / 100m, 2) : null,
                CurrentPremium = current,
                ProposedPremium = proposed,
                ModelledPremium = modelled,
                NegotiatedPremium = proposed,
                Status = row.Status == RenewalQueueStatuses.MissingData ? RenewalQueueStatuses.MissingData : RenewalQueueStatuses.Negotiated,
                Notes = isCurrentAnalysis
                    ? ""
                    : "Sample queue row. Load this client's renewal files for parsed rates."
            };
        }).ToList();
    }

    private async Task LoadNegotiationClaimsRowsAsync(RenewalQueueItem row)
    {
        _negotiationClaimsRows.Clear();

        var sourceRows = new List<MonthlyClaimsRecord>();
        if (IsCurrentAnalysisContext(row) && _claimsDataModelExport is not null)
        {
            sourceRows.AddRange(_claimsDataModelExport.MonthlyClaims);
        }

        if (sourceRows.Count == 0 && row.ClientId > 0)
        {
            try
            {
                sourceRows.AddRange(await _databaseService.GetMonthlyClaimsAsync(row.ClientId));
            }
            catch (Exception ex)
            {
                AppLogger.Error(ex);
            }
        }

        if (sourceRows.Count == 0)
        {
            var month = row.RenewalDate?.AddMonths(-1) ?? DateOnly.FromDateTime(DateTime.Today);
            sourceRows.Add(new MonthlyClaimsRecord
            {
                BenefitLine = "EHC",
                Month = month,
                Premium = row.CurrentPremium.HasValue ? decimal.Round(row.CurrentPremium.Value / 12m * 0.48m, 2) : null,
                PaidClaims = EstimateClaims(row.CurrentPremium, row.EhcLossRatio, 0.48m),
                IncurredClaims = EstimateClaims(row.CurrentPremium, row.EhcLossRatio, 0.48m),
                LossRatio = row.EhcLossRatio,
                Lives = row.Lives
            });
            sourceRows.Add(new MonthlyClaimsRecord
            {
                BenefitLine = "Dental",
                Month = month,
                Premium = row.CurrentPremium.HasValue ? decimal.Round(row.CurrentPremium.Value / 12m * 0.24m, 2) : null,
                PaidClaims = EstimateClaims(row.CurrentPremium, row.DentalLossRatio, 0.24m),
                IncurredClaims = EstimateClaims(row.CurrentPremium, row.DentalLossRatio, 0.24m),
                LossRatio = row.DentalLossRatio,
                Lives = row.Lives
            });
        }

        foreach (var claim in sourceRows.OrderByDescending(item => item.Month).ThenBy(item => item.BenefitLine))
        {
            _negotiationClaimsRows.Add(NegotiationClaimsRow.FromRecord(claim));
        }
    }

    private static decimal? EstimateClaims(decimal? annualPremium, decimal? lossRatio, decimal share)
    {
        return annualPremium.HasValue && lossRatio.HasValue
            ? decimal.Round(annualPremium.Value / 12m * share * lossRatio.Value / 100m, 2)
            : null;
    }

    private void PopulateAssumptionsFromRenewal(RenewalQueueItem row)
    {
        LivesAssumptionTextBox.Text = row.Lives?.ToString(CultureInfo.InvariantCulture) ?? "0";
        VolumeAssumptionTextBox.Text = "0";
        TrendAssumptionTextBox.Text = row.EhcLossRatio >= 90m ? "8.5" : "8.0";
        TargetLossRatioTextBox.Text = "76.3";
        PoolingChargeTextBox.Text = "0.0";
        IbnrTextBox.Text = "0.0";
        AdminFeeTextBox.Text = "0.0";
        RiskChargeTextBox.Text = "0.0";
        CredibilityFactorTextBox.Text = "100.0";
        ManualAdjustmentTextBox.Text = "0.0";
    }

    private void BackToDashboard_Click(object sender, RoutedEventArgs e)
    {
        ShowHomeDashboard();
    }

    private void ShowHomeDashboard_Click(object sender, RoutedEventArgs e)
    {
        ShowHomeDashboard();
    }

    private void BreadcrumbHome_Click(object sender, RoutedEventArgs e)
    {
        ShowHomeDashboard();
    }

    private void BreadcrumbClients_Click(object sender, RoutedEventArgs e)
    {
        ShowClientsPage();
    }

    private void ProposalComparison_ClientsBreadcrumbRequested(object? sender, EventArgs e)
    {
        ShowClientsPage();
    }

    private async void ProposalComparison_ViewPlanDesignRequested(object? sender, EventArgs e)
    {
        await ShowPlanDesignIntelligenceAsync();
    }

    private async void ShowRenewalsDashboard_Click(object sender, RoutedEventArgs e)
    {
        _clientRenewalSelectionIdentity = null;
        await ShowRenewalsDashboardAsync();
    }

    private async void OpenSelectedUnderwriting_Click(object sender, RoutedEventArgs e)
    {
        var row = SelectedRenewalQueueRow();

        if (row is null)
        {
            SetStatus("Select a renewal before opening Underwriting Intelligence.");
            return;
        }

        await ShowUnderwritingDashboardAsync(row);
    }

    private async void OpenSelectedNegotiation_Click(object sender, RoutedEventArgs e)
    {
        var row = SelectedRenewalQueueRow();

        if (row is null)
        {
            SetStatus("Select a renewal before opening negotiation.");
            return;
        }

        await OpenNegotiationWorkbenchAsync(row);
    }

    private void ViewSelectedDocuments_Click(object sender, RoutedEventArgs e)
    {
        var row = SelectedRenewalQueueRow();
        if (row is null)
        {
            SetStatus("Select a renewal before opening documents.");
            return;
        }

        SetStatus($"Documents: use Workflow Setup or Document Library to review source files for {row.Client}.");
    }

    private void ExportSelectedRenewalData_Click(object sender, RoutedEventArgs e)
    {
        ExportRenewalQueue_Click(sender, e);
    }

    private void AddSelectedRenewalNote_Click(object sender, RoutedEventArgs e)
    {
        var row = SelectedRenewalQueueRow();
        if (row is null)
        {
            SetStatus("Select a renewal before adding a note.");
            return;
        }

        SetStatus($"Note placeholder ready for {row.Client}. Consultant notes can be connected to saved history next.");
    }

    private RenewalQueueItem? SelectedRenewalQueueRow()
    {
        if (_clientRenewalSelectionIdentity is not null)
        {
            if (RenewalQueueGrid.SelectedItem is RenewalQueueItem selected &&
                ClientRenewalSelectionResolver.Matches(
                    _clientRenewalSelectionIdentity,
                    CreateRenewalSelectionCandidate(selected, 0)))
            {
                return selected;
            }

            return FindRenewalForClient(_clientRenewalSelectionIdentity);
        }

        return RenewalQueueGrid.SelectedItem as RenewalQueueItem
               ?? _renewalQueueView?.Cast<RenewalQueueItem>().FirstOrDefault()
               ?? _renewalQueue.FirstOrDefault();
    }

    private void ShowHomeDashboard()
    {
        SetNavigationWorkspace("Home", "Home");
        DataSourcesScrollViewer.Visibility = Visibility.Collapsed;
        UnderwritingWorkbenchScrollViewer.Visibility = Visibility.Collapsed;
        AssumptionLibraryScrollViewer.Visibility = Visibility.Collapsed;
        ReportsAnalyticsScrollViewer.Visibility = Visibility.Collapsed;
        BenchmarkingDashboardScrollViewer.Visibility = Visibility.Collapsed;
        ProposalComparisonScrollViewer.Visibility = Visibility.Collapsed;
        PlanDesignIntelligenceScrollViewer.Visibility = Visibility.Collapsed;
        PlanDesignFactorsScrollViewer.Visibility = Visibility.Collapsed;
        DocumentIntelligenceScrollViewer.Visibility = Visibility.Collapsed;
        ClientsScrollViewer.Visibility = Visibility.Collapsed;
        NegotiationWorkbenchScrollViewer.Visibility = Visibility.Collapsed;
        DashboardScrollViewer.Visibility = Visibility.Collapsed;
        HomeDashboardScrollViewer.Visibility = Visibility.Visible;
        HomeDashboardScrollViewer.ScrollToTop();
        SetOperationalFooterVisible(true);
        _ = RefreshHomeDashboardAsync();
        SetStatus("Returned to Benefits Intelligence OS dashboard.");
    }

    private void OpenMainNavigation_Click(object sender, RoutedEventArgs e)
    {
        SetNavigationWorkspace("Main", "Main Navigation");
        SetStatus("Opened main navigation.");
    }

    private void ShowRenewalsDashboard()
    {
        SetNavigationWorkspace("BookOfBusiness", "Renewal Pipeline");
        DataSourcesScrollViewer.Visibility = Visibility.Collapsed;
        HomeDashboardScrollViewer.Visibility = Visibility.Collapsed;
        ReportsAnalyticsScrollViewer.Visibility = Visibility.Collapsed;
        BenchmarkingDashboardScrollViewer.Visibility = Visibility.Collapsed;
        UnderwritingWorkbenchScrollViewer.Visibility = Visibility.Collapsed;
        AssumptionLibraryScrollViewer.Visibility = Visibility.Collapsed;
        ProposalComparisonScrollViewer.Visibility = Visibility.Collapsed;
        PlanDesignIntelligenceScrollViewer.Visibility = Visibility.Collapsed;
        PlanDesignFactorsScrollViewer.Visibility = Visibility.Collapsed;
        DocumentIntelligenceScrollViewer.Visibility = Visibility.Collapsed;
        ClientsScrollViewer.Visibility = Visibility.Collapsed;
        NegotiationWorkbenchScrollViewer.Visibility = Visibility.Collapsed;
        DashboardScrollViewer.Visibility = Visibility.Visible;
        DashboardScrollViewer.ScrollToTop();
        SetOperationalFooterVisible(true);
        SetStatus("Opened Renewal Intelligence dashboard.");
    }

    private async Task ShowRenewalsDashboardAsync()
    {
        ShowRenewalsDashboard();
        await LoadSavedRenewalQueueAsync();
    }

    private void ShowClaimsExperienceDashboard()
    {
        SetNavigationWorkspace("BookOfBusiness", "Claims Portfolio");
        HomeDashboardScrollViewer.Visibility = Visibility.Collapsed;
        DashboardScrollViewer.Visibility = Visibility.Collapsed;
        UnderwritingWorkbenchScrollViewer.Visibility = Visibility.Collapsed;
        AssumptionLibraryScrollViewer.Visibility = Visibility.Collapsed;
        ReportsAnalyticsScrollViewer.Visibility = Visibility.Collapsed;
        BenchmarkingDashboardScrollViewer.Visibility = Visibility.Collapsed;
        ProposalComparisonScrollViewer.Visibility = Visibility.Collapsed;
        PlanDesignIntelligenceScrollViewer.Visibility = Visibility.Collapsed;
        PlanDesignFactorsScrollViewer.Visibility = Visibility.Collapsed;
        DocumentIntelligenceScrollViewer.Visibility = Visibility.Collapsed;
        ClientsScrollViewer.Visibility = Visibility.Collapsed;
        NegotiationWorkbenchScrollViewer.Visibility = Visibility.Collapsed;
        DataSourcesScrollViewer.Visibility = Visibility.Collapsed;
        ClaimsExperienceDashboardScrollViewer.Visibility = Visibility.Visible;
        ClaimsExperienceDashboardScrollViewer.ScrollToTop();
        SetOperationalFooterVisible(true);
        SetStatus("Opened Book of Business Claims Experience dashboard.");
    }

    private async void OpenReportsAnalytics_Click(object sender, RoutedEventArgs e)
    {
        await ShowReportsAnalyticsAsync();
        SetNavigationWorkspace("Reports", "Reports");
    }

    private async void OpenReportsFromGlobal_Click(object sender, RoutedEventArgs e)
    {
        await ShowReportsAnalyticsAsync();
        SetNavigationWorkspace("Reports", "Reports");
        SetStatus("Opened Reports workspace.");
    }

    private async void OpenBookOfBusiness_Click(object sender, RoutedEventArgs e)
    {
        await ShowBookOfBusinessDashboardAsync();
        SetNavigationWorkspace("BookOfBusiness", "Portfolio Dashboard");
        SetStatus("Opened Book of Business workspace.");
    }

    private async void OpenExecutiveCenter_Click(object sender, RoutedEventArgs e)
    {
        await ShowReportsAnalyticsAsync();
        SetNavigationWorkspace("ExecutiveCenter", "Executive Dashboard");
        SetStatus("Opened Executive Center workspace.");
    }

    private void OpenBenchmarking_Click(object sender, RoutedEventArgs e)
    {
        ShowBenchmarkingDashboard();
        SetNavigationWorkspace("Benchmarking", "Benchmark Dashboard");
        SetStatus("Opened Benchmarking workspace.");
    }

    private void ShowBenchmarkingDashboard()
    {
        HomeDashboardScrollViewer.Visibility = Visibility.Collapsed;
        DashboardScrollViewer.Visibility = Visibility.Collapsed;
        DataSourcesScrollViewer.Visibility = Visibility.Collapsed;
        UnderwritingWorkbenchScrollViewer.Visibility = Visibility.Collapsed;
        AssumptionLibraryScrollViewer.Visibility = Visibility.Collapsed;
        ReportsAnalyticsScrollViewer.Visibility = Visibility.Collapsed;
        ProposalComparisonScrollViewer.Visibility = Visibility.Collapsed;
        PlanDesignIntelligenceScrollViewer.Visibility = Visibility.Collapsed;
        PlanDesignFactorsScrollViewer.Visibility = Visibility.Collapsed;
        DocumentIntelligenceScrollViewer.Visibility = Visibility.Collapsed;
        ClientsScrollViewer.Visibility = Visibility.Collapsed;
        NegotiationWorkbenchScrollViewer.Visibility = Visibility.Collapsed;
        BenchmarkingDashboardScrollViewer.Visibility = Visibility.Visible;
        BenchmarkingDashboardScrollViewer.ScrollToTop();
        SetOperationalFooterVisible(false);
    }

    private async void RefreshReportsAnalytics_Click(object sender, RoutedEventArgs e)
    {
        await ShowReportsAnalyticsAsync(forceRefresh: true);
    }

    private async void RefreshBookOfBusinessDashboard_Click(object sender, RoutedEventArgs e)
    {
        await ShowBookOfBusinessDashboardAsync(forceRefresh: true);
        SetNavigationWorkspace("BookOfBusiness", "Portfolio Dashboard");
    }

    private void ReportsFilter_Changed(object sender, RoutedEventArgs e)
    {
        if (_isUpdatingReportsFilters || _reportsSnapshot is null)
        {
            return;
        }

        ApplyReportsAnalyticsFilters();
    }

    private void ReportsBookBusinessMode_Click(object sender, RoutedEventArgs e)
    {
        _isReportsClientMode = false;
        UpdateReportsModeVisibility();
        ApplyReportsAnalyticsFilters();
    }

    private void ReportsClientDashboardMode_Click(object sender, RoutedEventArgs e)
    {
        _isReportsClientMode = true;
        UpdateReportsModeVisibility();
        ApplyReportsAnalyticsFilters();
    }

    private void ReportsSelectedClient_Changed(object sender, RoutedEventArgs e)
    {
        if (_isUpdatingReportsFilters || _reportsSnapshot is null)
        {
            return;
        }

        UpdateReportsClientDashboard(ReadReportsFilter(), FilterReportsClients(_reportsSnapshot, ReadReportsFilter()));
    }

    private async Task ShowReportsAnalyticsAsync(bool forceRefresh = false)
    {
        if (_isLoadingReportsAnalytics)
        {
            return;
        }

        _isLoadingReportsAnalytics = true;
        SetBusy(true);
        try
        {
            HomeDashboardScrollViewer.Visibility = Visibility.Collapsed;
            DashboardScrollViewer.Visibility = Visibility.Collapsed;
            DataSourcesScrollViewer.Visibility = Visibility.Collapsed;
            UnderwritingWorkbenchScrollViewer.Visibility = Visibility.Collapsed;
            AssumptionLibraryScrollViewer.Visibility = Visibility.Collapsed;
            BenchmarkingDashboardScrollViewer.Visibility = Visibility.Collapsed;
            ProposalComparisonScrollViewer.Visibility = Visibility.Collapsed;
            PlanDesignIntelligenceScrollViewer.Visibility = Visibility.Collapsed;
            PlanDesignFactorsScrollViewer.Visibility = Visibility.Collapsed;
            DocumentIntelligenceScrollViewer.Visibility = Visibility.Collapsed;
            ClientsScrollViewer.Visibility = Visibility.Collapsed;
            NegotiationWorkbenchScrollViewer.Visibility = Visibility.Collapsed;
            ReportsAnalyticsScrollViewer.Visibility = Visibility.Visible;
            ReportsExecutiveDashboardPanel.Visibility = Visibility.Visible;
            BookOfBusinessDashboardPanel.Visibility = Visibility.Collapsed;
            ReportsAnalyticsScrollViewer.ScrollToTop();
            SetOperationalFooterVisible(false);

            if (forceRefresh)
            {
                _databaseService.InvalidateExecutiveReportingCache();
            }

            _reportsSnapshot = await _databaseService.GetExecutiveReportingSnapshotAsync(forceRefresh);
            PopulateReportsFilters(_reportsSnapshot);
            UpdateReportsModeVisibility();
            ApplyReportsAnalyticsFilters();
            SetStatus($"Opened Reports & Analytics Command Centre from Benefits_Intelligence.db at {_reportsSnapshot.LoadedAt:t}.");
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            SetStatus("Reports & Analytics could not be loaded.");
            System.Windows.MessageBox.Show(
                this,
                $"Reports & Analytics could not be loaded: {ex.Message}",
                "Reports & Analytics",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }
        finally
        {
            _isLoadingReportsAnalytics = false;
            SetBusy(false);
        }
    }

    private async Task ShowBookOfBusinessDashboardAsync(bool forceRefresh = false)
    {
        if (_isLoadingReportsAnalytics)
        {
            return;
        }

        _isLoadingReportsAnalytics = true;
        SetBusy(true);
        try
        {
            HomeDashboardScrollViewer.Visibility = Visibility.Collapsed;
            DashboardScrollViewer.Visibility = Visibility.Collapsed;
            DataSourcesScrollViewer.Visibility = Visibility.Collapsed;
            UnderwritingWorkbenchScrollViewer.Visibility = Visibility.Collapsed;
            AssumptionLibraryScrollViewer.Visibility = Visibility.Collapsed;
            BenchmarkingDashboardScrollViewer.Visibility = Visibility.Collapsed;
            ProposalComparisonScrollViewer.Visibility = Visibility.Collapsed;
            PlanDesignIntelligenceScrollViewer.Visibility = Visibility.Collapsed;
            PlanDesignFactorsScrollViewer.Visibility = Visibility.Collapsed;
            DocumentIntelligenceScrollViewer.Visibility = Visibility.Collapsed;
            ClientsScrollViewer.Visibility = Visibility.Collapsed;
            NegotiationWorkbenchScrollViewer.Visibility = Visibility.Collapsed;
            ReportsAnalyticsScrollViewer.Visibility = Visibility.Visible;
            ReportsExecutiveDashboardPanel.Visibility = Visibility.Collapsed;
            BookOfBusinessDashboardPanel.Visibility = Visibility.Visible;
            ReportsAnalyticsScrollViewer.ScrollToTop();
            SetOperationalFooterVisible(false);

            if (forceRefresh)
            {
                _databaseService.InvalidateExecutiveReportingCache();
            }

            var snapshot = await _databaseService.GetExecutiveReportingSnapshotAsync(forceRefresh);
            UpdateBookOfBusinessDashboard(snapshot.Clients, snapshot.Renewals, snapshot.Proposals, snapshot.QuoteScenarios);
            SetStatus($"Opened Book of Business dashboard from saved Benefits Intelligence history at {DateTime.Now:t}.");
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            SetStatus("Book of Business dashboard could not be loaded.");
            System.Windows.MessageBox.Show(
                this,
                $"Book of Business dashboard could not be loaded: {ex.Message}",
                "Book of Business",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }
        finally
        {
            _isLoadingReportsAnalytics = false;
            SetBusy(false);
        }
    }

    private void UpdateBookOfBusinessDashboard(
        IReadOnlyCollection<ClientRecord> clients,
        IReadOnlyCollection<RenewalQueueRecord> renewals,
        IReadOnlyCollection<ExecutiveProposalRecord> proposals,
        IReadOnlyCollection<ExecutiveQuoteScenarioRecord> quotes)
    {
        var activeClients = clients.Where(IsActiveClient).OrderBy(client => client.ClientName).ToList();
        var latestRenewalsByClient = renewals
            .GroupBy(row => row.ClientId)
            .ToDictionary(
                group => group.Key,
                group => group
                    .OrderByDescending(row => row.RenewalDate ?? DateOnly.MinValue)
                    .ThenByDescending(row => row.DateLoaded)
                    .First());
        var activeClientIds = activeClients.Select(client => client.ClientId).ToHashSet();
        var portfolioRenewals = latestRenewalsByClient.Values
            .Where(row => activeClientIds.Count == 0 || activeClientIds.Contains(row.ClientId))
            .ToList();

        var clientCount = activeClients.Count > 0
            ? activeClients.Count
            : portfolioRenewals.Select(row => row.ClientId).Distinct().Count();
        var coveredLives = portfolioRenewals.Sum(row => row.Lives ?? 0);
        var annualPremium = portfolioRenewals.SumValue(row => row.RenewalPremium ?? row.CurrentPremium);
        var averageIncrease = AverageRaw(portfolioRenewals.Select(CalculateRenewalIncreasePercent));
        var upcomingRenewals = portfolioRenewals.Count(row => IsDueWithin(row.RenewalDate, 90));
        var retentionRate = clients.Count == 0
            ? (decimal?)null
            : activeClients.Count / (decimal)clients.Count * 100m;

        BookBusinessTotalClientsTextBlock.Text = clientCount.ToString("#,0", CultureInfo.InvariantCulture);
        BookBusinessCoveredLivesTextBlock.Text = coveredLives.ToString("#,0", CultureInfo.InvariantCulture);
        BookBusinessAnnualPremiumTextBlock.Text = FormatMoneyCompact(annualPremium);
        BookBusinessAverageIncreaseTextBlock.Text = FormatPercent(averageIncrease);
        BookBusinessUpcomingRenewalsTextBlock.Text = upcomingRenewals.ToString("#,0", CultureInfo.InvariantCulture);
        BookBusinessRetentionTextBlock.Text = FormatPercent(retentionRate, includeSign: false);
        BookBusinessDataSourceTextBlock.Text = clients.Count == 0 && renewals.Count == 0
            ? "No saved client or renewal history yet. The dashboard will populate as records are saved."
            : $"Based on {clients.Count:#,0} saved client(s), {renewals.Count:#,0} renewal record(s), and {proposals.Count + quotes.Count:#,0} pipeline item(s).";

        var actionRows = BuildBookOfBusinessActionRows(portfolioRenewals).ToList();
        ReplaceCollection(_bookOfBusinessActionRows, actionRows);
        BookBusinessActionCountTextBlock.Text = actionRows.Sum(row => row.Count).ToString("#,0", CultureInfo.InvariantCulture);
        ReplaceCollection(_bookOfBusinessRenewalCalendarRows, BuildBookOfBusinessRenewalCalendarRows(portfolioRenewals));
        ReplaceCollection(_bookOfBusinessPipelineRows, BuildBookOfBusinessPipelineRows(activeClients, portfolioRenewals, proposals, quotes));
        ReplaceCollection(_bookOfBusinessCarrierRows, BuildBookOfBusinessCarrierRows(portfolioRenewals));
        ReplaceCollection(_bookOfBusinessIndustryRows, BuildBookOfBusinessIndustryRows(activeClients));

        var clientRows = BuildBookOfBusinessClientRows(activeClients, portfolioRenewals, latestRenewalsByClient)
            .Take(12)
            .ToList();
        ReplaceCollection(_bookOfBusinessClientRows, clientRows);
        BookBusinessClientTableCountTextBlock.Text = clientCount == 0
            ? "0 clients"
            : $"Showing {clientRows.Count:#,0} of {clientCount:#,0} clients";
    }

    private static IEnumerable<ExecutiveSummaryRow> BuildBookOfBusinessActionRows(
        IReadOnlyCollection<RenewalQueueRecord> renewals)
    {
        var due30 = renewals.Count(row => IsDueWithin(row.RenewalDate, 30));
        var missingData = renewals.Count(HasMissingRequiredRenewalData);
        var highIncrease = renewals.Count(row => CalculateRenewalIncreasePercent(row) >= 10m);
        var highLossRatio = renewals.Count(row => row.EhcLossRatio >= 85m || row.DentalLossRatio >= 85m);

        return
        [
            BookDashboardRow("Renewals due in 30 days", due30, "Client-ready review window"),
            BookDashboardRow("Renewals missing required data", missingData, "Needs source follow-up"),
            BookDashboardRow("High renewal increase signal", highIncrease, "Above 10% increase"),
            BookDashboardRow("High loss-ratio signal", highLossRatio, "EHC or Dental at 85%+")
        ];
    }

    private static IEnumerable<ExecutiveSummaryRow> BuildBookOfBusinessRenewalCalendarRows(
        IReadOnlyCollection<RenewalQueueRecord> renewals)
    {
        var today = DateOnly.FromDateTime(DateTime.Today);
        var firstMonth = new DateOnly(today.Year, today.Month, 1);
        var monthRows = Enumerable.Range(0, 6)
            .Select(offset =>
            {
                var start = firstMonth.AddMonths(offset);
                var end = start.AddMonths(1);
                var monthRenewals = renewals
                    .Where(row => row.RenewalDate.HasValue &&
                                  row.RenewalDate.Value >= start &&
                                  row.RenewalDate.Value < end)
                    .ToList();
                return new
                {
                    Month = start,
                    Count = monthRenewals.Count,
                    Premium = monthRenewals.SumValue(row => row.RenewalPremium ?? row.CurrentPremium)
                };
            })
            .ToList();
        var maxCount = Math.Max(1, monthRows.Max(row => row.Count));

        return monthRows.Select(row => new ExecutiveSummaryRow
        {
            Category = "Renewal Calendar",
            Label = row.Month.ToString("MMM", CultureInfo.InvariantCulture),
            Value = row.Count.ToString("#,0", CultureInfo.InvariantCulture),
            Status = FormatMoneyCompact(row.Premium),
            Count = PercentBar(row.Count, maxCount)
        });
    }

    private static IEnumerable<ExecutiveSummaryRow> BuildBookOfBusinessPipelineRows(
        IReadOnlyCollection<ClientRecord> clients,
        IReadOnlyCollection<RenewalQueueRecord> renewals,
        IReadOnlyCollection<ExecutiveProposalRecord> proposals,
        IReadOnlyCollection<ExecutiveQuoteScenarioRecord> quotes)
    {
        var currentYear = DateTime.Today.Year;
        var newLogoClientIds = clients
            .Where(client => IsActiveClient(client) &&
                             client.DateCreated != DateTime.MinValue &&
                             client.DateCreated.Year == currentYear)
            .Select(client => client.ClientId)
            .ToHashSet();
        var openProposals = proposals.Where(item => !IsWonPipelineStatus(item.Status) && !IsLostPipelineStatus(item.Status) && !IsClosedStatus(item.Status)).ToList();
        var openQuotes = quotes.Where(item => !IsWonPipelineStatus(item.Status) && !IsLostPipelineStatus(item.Status) && !IsClosedStatus(item.Status)).ToList();
        var outstandingProposals = proposals.Where(item => IsMarketOutstandingStatus(item.Status)).ToList();
        var outstandingQuotes = quotes.Where(item => IsMarketOutstandingStatus(item.Status)).ToList();
        var wonProposals = proposals.Where(item => IsWonPipelineStatus(item.Status)).ToList();
        var wonQuotes = quotes.Where(item => IsWonPipelineStatus(item.Status)).ToList();
        var lostProposals = proposals.Where(item => IsLostPipelineStatus(item.Status)).ToList();
        var lostQuotes = quotes.Where(item => IsLostPipelineStatus(item.Status)).ToList();

        return
        [
            BookPipelineRow(
                "New Logos",
                newLogoClientIds.Count,
                SumPipelinePremium(renewals.Where(row => newLogoClientIds.Contains(row.ClientId)).Select(row => row.RenewalPremium ?? row.CurrentPremium)),
                "\uE716",
                "+",
                "#7C3AED",
                "#F3E8FF"),
            BookPipelineRow(
                "Quotes in Progress",
                openProposals.Count + openQuotes.Count,
                SumPipelinePremium(openProposals.Select(item => item.ProposedPremium ?? item.CurrentPremium)
                    .Concat(openQuotes.Select(item => item.EstimatedPremium ?? item.CurrentPremium))),
                "\uE8A5",
                "",
                "#7C3AED",
                "#F3E8FF"),
            BookPipelineRow(
                "Markets Outstanding",
                outstandingProposals.Count + outstandingQuotes.Count,
                SumPipelinePremium(outstandingProposals.Select(item => item.ProposedPremium ?? item.CurrentPremium)
                    .Concat(outstandingQuotes.Select(item => item.EstimatedPremium ?? item.CurrentPremium))),
                "\uE823",
                "",
                "#7C3AED",
                "#F3E8FF"),
            BookPipelineRow(
                "Won",
                wonProposals.Count + wonQuotes.Count,
                SumPipelinePremium(wonProposals.Select(item => item.ProposedPremium ?? item.CurrentPremium)
                    .Concat(wonQuotes.Select(item => item.EstimatedPremium ?? item.CurrentPremium))),
                "\uE73E",
                "",
                "#059669",
                "#D1FAE5"),
            BookPipelineRow(
                "Lost",
                lostProposals.Count + lostQuotes.Count,
                SumPipelinePremium(lostProposals.Select(item => item.ProposedPremium ?? item.CurrentPremium)
                    .Concat(lostQuotes.Select(item => item.EstimatedPremium ?? item.CurrentPremium))),
                "\uE711",
                "",
                "#EF4444",
                "#FEE2E2")
        ];
    }

    private static ExecutiveSummaryRow BookPipelineRow(
        string label,
        int count,
        decimal? premium,
        string icon,
        string badge,
        string iconBrush,
        string iconBackground)
    {
        return new ExecutiveSummaryRow
        {
            Category = "Pipeline",
            Label = label,
            Value = count.ToString("#,0", CultureInfo.InvariantCulture),
            Status = $"{FormatMoneyCompact(premium)} premium",
            Count = count,
            Icon = icon,
            Badge = badge,
            IconBrush = iconBrush,
            IconBackground = iconBackground
        };
    }

    private static bool IsMarketOutstandingStatus(string status)
    {
        if (IsClosedStatus(status))
        {
            return false;
        }

        return status.Contains("market", StringComparison.OrdinalIgnoreCase) ||
               status.Contains("outstanding", StringComparison.OrdinalIgnoreCase) ||
               status.Contains("carrier", StringComparison.OrdinalIgnoreCase) ||
               status.Contains("pending", StringComparison.OrdinalIgnoreCase) ||
               status.Contains("requested", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsWonPipelineStatus(string status)
    {
        return status.Contains("won", StringComparison.OrdinalIgnoreCase) ||
               status.Contains("sold", StringComparison.OrdinalIgnoreCase) ||
               status.Contains("complete", StringComparison.OrdinalIgnoreCase) ||
               status.Contains("final", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsLostPipelineStatus(string status)
    {
        return status.Contains("lost", StringComparison.OrdinalIgnoreCase) ||
               status.Contains("cancel", StringComparison.OrdinalIgnoreCase) ||
               status.Contains("declin", StringComparison.OrdinalIgnoreCase);
    }

    private static decimal? SumPipelinePremium(IEnumerable<decimal?> values)
    {
        var savedValues = values.Where(value => value.HasValue).Select(value => value!.Value).ToList();
        return savedValues.Count == 0 ? null : savedValues.Sum();
    }

    private static IEnumerable<ExecutiveSummaryRow> BuildBookOfBusinessCarrierRows(
        IReadOnlyCollection<RenewalQueueRecord> renewals)
    {
        var premiumTotal = renewals.SumValue(row => row.RenewalPremium ?? row.CurrentPremium) ?? 0m;
        var countTotal = Math.Max(1, renewals.Count);
        var rows = renewals
            .GroupBy(row => FirstNonBlank(row.Carrier, "Unassigned Carrier"), StringComparer.OrdinalIgnoreCase)
            .Select(group =>
            {
                var premium = group.SumValue(row => row.RenewalPremium ?? row.CurrentPremium) ?? 0m;
                var percent = premiumTotal > 0m
                    ? premium / premiumTotal * 100m
                    : group.Count() / (decimal)countTotal * 100m;
                return new ExecutiveSummaryRow
                {
                    Category = "Carrier",
                    Label = group.Key,
                    Value = FormatMoneyCompact(premium),
                    Status = $"{FormatMoneyCompact(premium)} ({percent:0}%)",
                    Count = PercentBar(percent, 100m)
                };
            })
            .OrderByDescending(row => row.Count)
            .ThenBy(row => row.Label)
            .Take(6)
            .ToList();

        return rows.Count == 0
            ? [new ExecutiveSummaryRow { Category = "Carrier", Label = "No carrier data", Status = "0%", Count = 0 }]
            : rows;
    }

    private static IEnumerable<ExecutiveSummaryRow> BuildBookOfBusinessIndustryRows(
        IReadOnlyCollection<ClientRecord> clients)
    {
        var total = Math.Max(1, clients.Count);
        var rows = clients
            .GroupBy(client => FirstNonBlank(client.Industry, "Unassigned Industry"), StringComparer.OrdinalIgnoreCase)
            .Select(group =>
            {
                var percent = group.Count() / (decimal)total * 100m;
                return new ExecutiveSummaryRow
                {
                    Category = "Industry",
                    Label = group.Key,
                    Value = group.Count().ToString("#,0", CultureInfo.InvariantCulture),
                    Status = $"{group.Count():#,0} ({percent:0}%)",
                    Count = PercentBar(percent, 100m)
                };
            })
            .OrderByDescending(row => row.Count)
            .ThenBy(row => row.Label)
            .Take(6)
            .ToList();

        return rows.Count == 0
            ? [new ExecutiveSummaryRow { Category = "Industry", Label = "No industry data", Status = "0%", Count = 0 }]
            : rows;
    }

    private static IEnumerable<BookOfBusinessClientRow> BuildBookOfBusinessClientRows(
        IReadOnlyCollection<ClientRecord> activeClients,
        IReadOnlyCollection<RenewalQueueRecord> portfolioRenewals,
        IReadOnlyDictionary<long, RenewalQueueRecord> latestRenewalsByClient)
    {
        if (activeClients.Count == 0)
        {
            return portfolioRenewals
                .OrderBy(row => row.RenewalDate ?? DateOnly.MaxValue)
                .ThenBy(row => row.ClientName)
                .Select(row => new BookOfBusinessClientRow
                {
                    ClientName = FirstNonBlank(row.ClientName, $"Client #{row.ClientId}"),
                    Owner = "Unassigned",
                    Carrier = FirstNonBlank(row.Carrier, "Not set"),
                    Industry = "Not set",
                    Lives = row.Lives?.ToString("#,0", CultureInfo.InvariantCulture) ?? "N/A",
                    AnnualPremium = FormatMoney(row.RenewalPremium ?? row.CurrentPremium),
                    RenewalDate = row.RenewalDate?.ToString("MMM d, yyyy", CultureInfo.InvariantCulture) ?? "Not scheduled",
                    Stage = MapRenewalStage(row),
                    Status = FirstNonBlank(row.Status, "Saved")
                });
        }

        return activeClients
            .OrderBy(client =>
            {
                latestRenewalsByClient.TryGetValue(client.ClientId, out var renewal);
                return renewal?.RenewalDate ?? DateOnly.MaxValue;
            })
            .ThenBy(client => client.ClientName)
            .Select(client =>
            {
                latestRenewalsByClient.TryGetValue(client.ClientId, out var renewal);
                return new BookOfBusinessClientRow
                {
                    ClientName = FirstNonBlank(client.ClientName, $"Client #{client.ClientId}"),
                    Owner = FirstNonBlank(client.Advisor, "Unassigned"),
                    Carrier = FirstNonBlank(renewal?.Carrier, client.Carrier, "Not set"),
                    Industry = FirstNonBlank(client.Industry, "Not set"),
                    Lives = renewal?.Lives?.ToString("#,0", CultureInfo.InvariantCulture) ?? "N/A",
                    AnnualPremium = FormatMoney(renewal?.RenewalPremium ?? renewal?.CurrentPremium),
                    RenewalDate = renewal?.RenewalDate?.ToString("MMM d, yyyy", CultureInfo.InvariantCulture) ?? "Not scheduled",
                    Stage = renewal is null ? "No renewal saved" : MapRenewalStage(renewal),
                    Status = FirstNonBlank(client.Status, "Active")
                };
            });
    }

    private static ExecutiveSummaryRow BookDashboardRow(string label, int count, string status)
    {
        return new ExecutiveSummaryRow
        {
            Category = "Book of Business",
            Label = label,
            Value = count.ToString("#,0", CultureInfo.InvariantCulture),
            Status = status,
            Count = count
        };
    }

    private static int PercentBar(decimal value, decimal total)
    {
        if (value <= 0m || total <= 0m)
        {
            return 0;
        }

        var percent = (int)Math.Round(value / total * 100m, MidpointRounding.AwayFromZero);
        return Math.Min(100, Math.Max(4, percent));
    }

    private void UpdateReportsModeVisibility()
    {
        ReportsBookBusinessPanel.Visibility = _isReportsClientMode ? Visibility.Collapsed : Visibility.Visible;
        ReportsClientDashboardPanel.Visibility = _isReportsClientMode ? Visibility.Visible : Visibility.Collapsed;
        ReportsBookBusinessModeButton.Background = _isReportsClientMode ? Brushes.White : (Brush)FindResource("DeepBlueBrush");
        ReportsBookBusinessModeButton.Foreground = _isReportsClientMode ? (Brush)FindResource("DeepBlueBrush") : Brushes.White;
        ReportsClientDashboardModeButton.Background = _isReportsClientMode ? (Brush)FindResource("DeepBlueBrush") : Brushes.White;
        ReportsClientDashboardModeButton.Foreground = _isReportsClientMode ? Brushes.White : (Brush)FindResource("DeepBlueBrush");
    }

    private void PopulateReportsFilters(ExecutiveReportingSnapshot snapshot)
    {
        _isUpdatingReportsFilters = true;
        try
        {
            SetReportsComboBoxValues(ReportsClientFilterComboBox, "All Clients", snapshot.Clients.Select(item => item.ClientName));
            SetReportsComboBoxValues(ReportsCarrierFilterComboBox, "All Carriers", snapshot.Clients.Select(item => item.Carrier).Concat(snapshot.Renewals.Select(item => item.Carrier)).Concat(snapshot.Proposals.Select(item => item.Carrier)));
            SetReportsComboBoxValues(ReportsIndustryFilterComboBox, "All Industries", snapshot.Clients.Select(item => item.Industry));
            SetReportsComboBoxValues(ReportsConsultantFilterComboBox, "All Consultants", snapshot.Clients.Select(item => item.Advisor));
            SetReportsComboBoxValues(ReportsOfficeFilterComboBox, "All Offices", []);
            SetReportsComboBoxValues(ReportsProvinceFilterComboBox, "All Provinces", snapshot.Clients.Select(item => item.Province));
            SetReportsComboBoxValues(ReportsEffectiveDateFilterComboBox, "All Effective Dates", snapshot.Clients.Select(item => item.EffectiveDate?.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) ?? ""));
            SetReportsComboBoxValues(ReportsCalendarYearFilterComboBox, "All Calendar Years", snapshot.Renewals.Select(item => item.DateLoaded.Year.ToString(CultureInfo.InvariantCulture)));
            SetReportsComboBoxValues(ReportsRenewalYearFilterComboBox, "All Renewal Years", snapshot.Renewals.Where(item => item.RenewalYear > 0).Select(item => item.RenewalYear.ToString(CultureInfo.InvariantCulture)));
            SetReportsComboBoxValues(ReportsFundingTypeFilterComboBox, "All Funding Types", ["ASO", "Refund Accounting", "Insured", "Hybrid"]);

            ReportsSelectedClientComboBox.ItemsSource = snapshot.Clients.OrderBy(item => item.ClientName).ToList();
            if (ReportsSelectedClientComboBox.SelectedItem is null && snapshot.Clients.Count > 0)
            {
                ReportsSelectedClientComboBox.SelectedIndex = 0;
            }
        }
        finally
        {
            _isUpdatingReportsFilters = false;
        }
    }

    private static void SetReportsComboBoxValues(ComboBox comboBox, string allLabel, IEnumerable<string> values)
    {
        var selected = SelectedComboBoxText(comboBox);
        comboBox.Items.Clear();
        comboBox.Items.Add(allLabel);
        foreach (var value in values
                     .Where(value => !string.IsNullOrWhiteSpace(value))
                     .Distinct(StringComparer.OrdinalIgnoreCase)
                     .OrderBy(value => value))
        {
            comboBox.Items.Add(value);
        }

        comboBox.SelectedItem = comboBox.Items.Cast<object>().FirstOrDefault(item =>
            item?.ToString()?.Equals(selected, StringComparison.OrdinalIgnoreCase) == true) ?? allLabel;
    }

    private ReportsFilterCriteria ReadReportsFilter()
    {
        return new ReportsFilterCriteria
        {
            Client = SelectedComboBoxText(ReportsClientFilterComboBox),
            Carrier = SelectedComboBoxText(ReportsCarrierFilterComboBox),
            Industry = SelectedComboBoxText(ReportsIndustryFilterComboBox),
            Consultant = SelectedComboBoxText(ReportsConsultantFilterComboBox),
            Office = SelectedComboBoxText(ReportsOfficeFilterComboBox),
            Province = SelectedComboBoxText(ReportsProvinceFilterComboBox),
            EffectiveDate = SelectedComboBoxText(ReportsEffectiveDateFilterComboBox),
            CalendarYear = SelectedComboBoxText(ReportsCalendarYearFilterComboBox),
            RenewalYear = SelectedComboBoxText(ReportsRenewalYearFilterComboBox),
            FundingType = SelectedComboBoxText(ReportsFundingTypeFilterComboBox),
            Role = SelectedComboBoxText(ReportsRoleComboBox)
        };
    }

    private void ApplyReportsAnalyticsFilters()
    {
        if (_reportsSnapshot is null)
        {
            return;
        }

        var filter = ReadReportsFilter();
        var clients = FilterReportsClients(_reportsSnapshot, filter);
        var clientIds = clients.Select(item => item.ClientId).ToHashSet();
        var renewals = FilterReportsRenewals(_reportsSnapshot, filter, clientIds);
        var claims = _reportsSnapshot.Claims.Where(item => clientIds.Contains(item.ClientId)).ToList();
        var monthlyClaims = _reportsSnapshot.MonthlyClaims.Where(item => clientIds.Contains(item.ClientId)).ToList();
        var planDesign = _reportsSnapshot.PlanDesign.Where(item => clientIds.Contains(item.ClientId)).ToList();
        var assumptions = _reportsSnapshot.RenewalAssumptions.Where(item => clientIds.Contains(item.ClientId)).ToList();
        var proposals = FilterReportsProposals(_reportsSnapshot, filter, clients);
        var quotes = _reportsSnapshot.QuoteScenarios.Where(item => item.ClientId == 0 || clientIds.Contains(item.ClientId)).ToList();
        var versions = _reportsSnapshot.RenewalVersions.Where(item => clientIds.Contains(item.ClientId)).ToList();

        BuildReportsKpis(clients, renewals, claims, assumptions, versions, proposals, quotes);
        UpdateExecutiveCommandCentre(clients, renewals, claims, monthlyClaims, assumptions, versions, proposals, quotes);
        ReplaceCollection(_reportsRenewalPipelineRows, BuildRenewalPipelineRows(renewals));
        ReplaceCollection(_reportsProposalPipelineRows, BuildProposalPipelineRows(proposals));
        ReplaceCollection(_reportsQuotePipelineRows, BuildQuotePipelineRows(quotes));
        ReplaceCollection(_reportsRenewalsByMonthRows, BuildRenewalsByMonthRows(renewals));
        ReplaceCollection(_reportsRenewalsByCarrierRows, BuildPremiumByCarrierRows(renewals));
        ReplaceCollection(_reportsRenewalsByConsultantRows, BuildRenewalCalendarRows(renewals));
        ReplaceCollection(_reportsRenewalAnalyticsRows, BuildRenewalDistributionRows(renewals));
        ReplaceCollection(_reportsClaimsAnalyticsRows, BuildRenewalDriverRows(renewals, claims, assumptions));
        ReplaceCollection(_reportsUnderwritingAnalyticsRows, BuildUnderwritingAnalyticsRows(assumptions, versions));
        ReplaceCollection(_reportsCarrierAnalyticsRows, BuildCarrierAnalyticsRows(renewals));
        ReplaceCollection(_reportsProposalAnalyticsRows, BuildProposalAnalyticsRows(proposals));
        ReplaceCollection(_reportsPlanDesignAnalyticsRows, BuildPlanDesignAnalyticsRows(planDesign));
        ReplaceCollection(_reportsBenchmarkRows, BuildBenchmarkRows(clients, renewals, claims));
        ReplaceCollection(_reportsAlertRows, BuildExecutiveAlertRows(renewals, claims, assumptions, proposals, quotes, versions));
        ReplaceCollection(_reportsExecutiveDecisionRows, BuildExecutiveDecisionRows(renewals, claims, proposals, quotes));
        ReplaceCollection(_reportsInsightRows, BuildInsightRows(renewals, claims, proposals, clients));
        ReplaceCollection(_reportsDataCategoryRows, _reportsSnapshot.DataCategories.OrderByDescending(item => item.Records).Take(18));
        ReplaceCollection(_reportsDrillDownRows, BuildClientRiskWatchlistRows(renewals));

        if (_isReportsClientMode)
        {
            UpdateReportsClientDashboard(filter, clients);
        }
    }

    private static List<ClientRecord> FilterReportsClients(ExecutiveReportingSnapshot snapshot, ReportsFilterCriteria filter)
    {
        return snapshot.Clients
            .Where(client => FilterMatches(filter.Client, "All Clients", client.ClientName))
            .Where(client => FilterMatches(filter.Carrier, "All Carriers", client.Carrier))
            .Where(client => FilterMatches(filter.Industry, "All Industries", client.Industry))
            .Where(client => FilterMatches(filter.Consultant, "All Consultants", client.Advisor))
            .Where(client => FilterMatches(filter.Province, "All Provinces", client.Province))
            .Where(client => FilterMatches(filter.EffectiveDate, "All Effective Dates", client.EffectiveDate?.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) ?? ""))
            .ToList();
    }

    private static List<RenewalQueueRecord> FilterReportsRenewals(ExecutiveReportingSnapshot snapshot, ReportsFilterCriteria filter, HashSet<long> clientIds)
    {
        return snapshot.Renewals
            .Where(item => clientIds.Contains(item.ClientId))
            .Where(item => FilterMatches(filter.Carrier, "All Carriers", item.Carrier))
            .Where(item => FilterMatches(filter.CalendarYear, "All Calendar Years", item.DateLoaded.Year.ToString(CultureInfo.InvariantCulture)))
            .Where(item => FilterMatches(filter.RenewalYear, "All Renewal Years", item.RenewalYear > 0 ? item.RenewalYear.ToString(CultureInfo.InvariantCulture) : ""))
            .Where(item => FilterMatches(filter.EffectiveDate, "All Effective Dates", item.RenewalDate?.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) ?? ""))
            .ToList();
    }

    private static List<ExecutiveProposalRecord> FilterReportsProposals(ExecutiveReportingSnapshot snapshot, ReportsFilterCriteria filter, IReadOnlyCollection<ClientRecord> clients)
    {
        var clientNames = clients.Select(item => item.ClientName).Where(value => !string.IsNullOrWhiteSpace(value)).ToHashSet(StringComparer.OrdinalIgnoreCase);
        return snapshot.Proposals
            .Where(item => string.IsNullOrWhiteSpace(item.ClientName) || clientNames.Count == 0 || clientNames.Contains(item.ClientName))
            .Where(item => FilterMatches(filter.Carrier, "All Carriers", item.Carrier))
            .ToList();
    }

    private static bool FilterMatches(string selected, string allLabel, string value)
    {
        return string.IsNullOrWhiteSpace(selected) ||
               selected.Equals(allLabel, StringComparison.OrdinalIgnoreCase) ||
               selected.StartsWith("All ", StringComparison.OrdinalIgnoreCase) ||
               value.Equals(selected, StringComparison.OrdinalIgnoreCase);
    }

    private static void ReplaceCollection<T>(ObservableCollection<T> collection, IEnumerable<T> values)
    {
        collection.Clear();
        foreach (var value in values)
        {
            collection.Add(value);
        }
    }

    private void BuildReportsKpis(
        IReadOnlyCollection<ClientRecord> clients,
        IReadOnlyCollection<RenewalQueueRecord> renewals,
        IReadOnlyCollection<ClaimsExperienceRecord> claims,
        IReadOnlyCollection<RenewalAssumptionRecord> assumptions,
        IReadOnlyCollection<RenewalVersionRecord> versions,
        IReadOnlyCollection<ExecutiveProposalRecord> proposals,
        IReadOnlyCollection<ExecutiveQuoteScenarioRecord> quotes)
    {
        _reportsKpiCards.Clear();

        var activeClients = clients.Count(client => client.Status.Equals("Active", StringComparison.OrdinalIgnoreCase) || string.IsNullOrWhiteSpace(client.Status));
        var premiumManaged = renewals.SumValue(item => item.RenewalPremium ?? item.CurrentPremium);
        var renewalIncrease = AverageRaw(renewals.Select(item => item.PremiumChangePercent));
        var ehcLossRatio = AverageRaw(renewals.Select(item => item.EhcLossRatio).Concat(claims.Where(item => IsBenefit(item.BenefitLine, "ehc", "health")).Select(item => item.LossRatio)));
        var dentalLossRatio = AverageRaw(renewals.Select(item => item.DentalLossRatio).Concat(claims.Where(item => IsBenefit(item.BenefitLine, "dental")).Select(item => item.LossRatio)));
        var poolingLevel = AverageRaw(assumptions.Select(item => item.PoolingThreshold));
        var drugTrend = AverageNormalizedRate(assumptions.Where(item => IsBenefit(item.BenefitLine, "drug")).Select(item => item.Trend));
        var savings = CalculateNegotiatedSavings(versions);
        var opportunities = renewals.Count(item => item.PremiumChangePercent >= 10m || item.EhcLossRatio >= 100m || item.DentalLossRatio >= 100m) +
                            proposals.Count(item => !IsClosedStatus(item.Status)) +
                            quotes.Count(item => !IsClosedStatus(item.Status));
        var employerContribution = AverageRaw(proposals.Select(item => item.EmployerCost));
        var employeeContribution = AverageRaw(proposals.Select(item => item.EmployeeCost));

        AddReportKpi("Total Premium Managed", FormatMoney(premiumManaged), $"{FormatMoneyCompact(renewals.Where(item => IsDueWithin(item.RenewalDate, 90)).SumValue(item => item.RenewalPremium ?? item.CurrentPremium))} upcoming", "$");
        AddReportKpi("Forecast Commission (FY)", FormatMoneyCompact(EstimateForecastCommission(premiumManaged)), "Derived until commission table is connected", "%");
        AddReportKpi("Upcoming Renewals (90 Days)", CountDueInDays(renewals, 90).ToString("#,0", CultureInfo.InvariantCulture), $"{CountDueInDays(renewals, 30)} / {CountDueInDays(renewals, 60) - CountDueInDays(renewals, 30)} / {CountDueInDays(renewals, 90) - CountDueInDays(renewals, 60)} by 30 / 60 / 90", "90");
        AddReportKpi("Average Renewal Increase", FormatPercent(renewalIncrease), renewalIncrease.HasValue && renewalIncrease.Value > 0m ? "Executive attention signal" : "Stable", "%");
        AddReportKpi("Book Loss Ratio (EHC)", FormatPercent(ehcLossRatio, includeSign: false), ehcLossRatio >= 85m ? "Above target" : "Healthy", "E");
        AddReportKpi("Dental Loss Ratio", FormatPercent(dentalLossRatio, includeSign: false), dentalLossRatio >= 85m ? "Above target" : "Healthy", "D");
        AddReportKpi("Negotiated Savings (YTD)", FormatMoney(savings), savings.HasValue ? "Saved negotiated versions" : "No saved negotiations", "$");
    }

    private void AddReportKpi(string title, string value, string subtitle, string icon = "i")
    {
        _reportsKpiCards.Add(new ExecutiveMetricCard
        {
            Title = title,
            Value = value,
            Subtitle = subtitle,
            Icon = icon
        });
    }

    private void UpdateExecutiveCommandCentre(
        IReadOnlyCollection<ClientRecord> clients,
        IReadOnlyCollection<RenewalQueueRecord> renewals,
        IReadOnlyCollection<ClaimsExperienceRecord> claims,
        IReadOnlyCollection<MonthlyClaimsRecord> monthlyClaims,
        IReadOnlyCollection<RenewalAssumptionRecord> assumptions,
        IReadOnlyCollection<RenewalVersionRecord> versions,
        IReadOnlyCollection<ExecutiveProposalRecord> proposals,
        IReadOnlyCollection<ExecutiveQuoteScenarioRecord> quotes)
    {
        var activeClients = clients.Count(client => client.Status.Equals("Active", StringComparison.OrdinalIgnoreCase) || string.IsNullOrWhiteSpace(client.Status));
        var premiumManaged = renewals.SumValue(item => item.RenewalPremium ?? item.CurrentPremium);
        var coveredLives = renewals.Select(item => item.Lives).Where(item => item.HasValue).Select(item => item!.Value).DefaultIfEmpty(0).Sum();
        var renewalIncrease = AverageRaw(renewals.Select(item => item.PremiumChangePercent));
        var ehcLossRatio = AverageRaw(renewals.Select(item => item.EhcLossRatio).Concat(claims.Where(item => IsBenefit(item.BenefitLine, "ehc", "health")).Select(item => item.LossRatio)));
        var dentalLossRatio = AverageRaw(renewals.Select(item => item.DentalLossRatio).Concat(claims.Where(item => IsBenefit(item.BenefitLine, "dental")).Select(item => item.LossRatio)));
        var savings = CalculateNegotiatedSavings(versions);
        var upcoming90 = renewals.Where(item => IsDueWithin(item.RenewalDate, 90)).ToList();
        var upcomingPremium = upcoming90.SumValue(item => item.RenewalPremium ?? item.CurrentPremium);
        var atRisk90 = upcoming90.Count(IsExecutiveRiskRenewal);
        var potentialSavings = EstimatePotentialSavings(upcoming90, savings);
        var forecastCommission = EstimateForecastCommission(premiumManaged);
        var forecastConfidence = CalculateForecastConfidence(renewals, claims, assumptions);

        // Derived until a dedicated commission table is added. The screen remains read-only and does not change stored values.
        var financialHealth = ScoreRenewalIncrease(renewalIncrease);
        var claimsStability = ScoreLossRatio(AverageRaw(new[] { ehcLossRatio, dentalLossRatio }));
        var renewalRisk = CalculateRenewalRiskScore(renewals);
        var implementation = CalculateImplementationScore(renewals, proposals, quotes);
        var retention = CalculateRetentionScore(clients);
        var bookIndex = (financialHealth * 0.24m) + (claimsStability * 0.24m) + (renewalRisk * 0.20m) + (implementation * 0.16m) + (retention * 0.16m);

        var loadedAt = _reportsSnapshot?.LoadedAt ?? DateTime.Now;
        ReportsLastUpdatedTextBlock.Text = loadedAt.ToString("MMM d, yyyy h:mm tt", CultureInfo.InvariantCulture);
        ReportsIndexScoreTextBlock.Text = bookIndex.ToString("0", CultureInfo.InvariantCulture);
        ReportsIndexStatusTextBlock.Text = BookIndexStatus(bookIndex);
        ReportsIndexTrendTextBlock.Text = renewalIncrease.HasValue
            ? $"{FormatPercent(renewalIncrease)} average renewal increase"
            : "Trend will populate as renewal history grows";
        ReportsFinancialHealthScoreTextBlock.Text = financialHealth.ToString("0", CultureInfo.InvariantCulture);
        ReportsClaimsStabilityScoreTextBlock.Text = claimsStability.ToString("0", CultureInfo.InvariantCulture);
        ReportsRenewalRiskScoreTextBlock.Text = renewalRisk.ToString("0", CultureInfo.InvariantCulture);
        ReportsImplementationScoreTextBlock.Text = implementation.ToString("0", CultureInfo.InvariantCulture);
        ReportsRetentionScoreTextBlock.Text = retention.ToString("0", CultureInfo.InvariantCulture);
        ReportsExecutiveBriefTimestampTextBlock.Text = $"Generated {loadedAt:MMM d, yyyy h:mm tt}";
        ReportsExecutiveBriefTextBlock.Text =
            $"Your book is {BookIndexStatus(bookIndex).ToLowerInvariant()} ({bookIndex:0}/100). " +
            $"Premium under management is {FormatMoneyCompact(premiumManaged)} and {upcoming90.Count:#,0} renewals are forecast over the next 90 days. " +
            $"{atRisk90:#,0} renewals require executive attention. Estimated negotiable savings across upcoming renewals: {FormatMoneyCompact(potentialSavings)}.";

        ReportsHeroClientsTextBlock.Text = activeClients.ToString("#,0", CultureInfo.InvariantCulture);
        ReportsHeroLivesTextBlock.Text = coveredLives > 0 ? coveredLives.ToString("#,0", CultureInfo.InvariantCulture) : "N/A";
        ReportsHeroPremiumTextBlock.Text = FormatMoneyCompact(premiumManaged);
        ReportsHeroCommissionTextBlock.Text = FormatMoneyCompact(forecastCommission);

        ReportsNinetyDayExpectedPremiumTextBlock.Text = FormatMoneyCompact(upcomingPremium);
        ReportsNinetyDayCommissionTextBlock.Text = FormatMoneyCompact(EstimateForecastCommission(upcomingPremium));
        ReportsNinetyDayRenewalsTextBlock.Text = upcoming90.Count.ToString("#,0", CultureInfo.InvariantCulture);
        ReportsNinetyDayAtRiskTextBlock.Text = atRisk90.ToString("#,0", CultureInfo.InvariantCulture);
        ReportsNinetyDaySavingsTextBlock.Text = FormatMoneyCompact(potentialSavings);
        ReportsNinetyDayConfidenceTextBlock.Text = $"{forecastConfidence:0}%";
        ReportsExecutiveAlertCountTextBlock.Text = (BuildExecutiveAlertRows(renewals, claims, assumptions, proposals, quotes, versions).Sum(item => item.Count)).ToString("#,0", CultureInfo.InvariantCulture);
    }

    private static IEnumerable<ExecutiveSummaryRow> BuildExecutiveAlertRows(
        IReadOnlyCollection<RenewalQueueRecord> renewals,
        IReadOnlyCollection<ClaimsExperienceRecord> claims,
        IReadOnlyCollection<RenewalAssumptionRecord> assumptions,
        IReadOnlyCollection<ExecutiveProposalRecord> proposals,
        IReadOnlyCollection<ExecutiveQuoteScenarioRecord> quotes,
        IReadOnlyCollection<RenewalVersionRecord> versions)
    {
        var poolingGroups = renewals.Count(item => (NormalizeRatePercent(item.PoolingCharge) ?? 0m) >= 5m) +
                            assumptions.Count(item => (NormalizeRatePercent(item.PoolingCharge) ?? 0m) >= 5m);
        var highLossClients = renewals.Count(item => item.EhcLossRatio >= 100m || item.DentalLossRatio >= 100m) +
                              claims.Count(item => item.LossRatio >= 100m);
        var highIncrease = renewals.Count(item => item.PremiumChangePercent >= 20m);
        var openProposalWork = proposals.Count(item => !IsClosedStatus(item.Status)) + quotes.Count(item => !IsClosedStatus(item.Status));
        var savings = CalculateNegotiatedSavings(versions);

        return
        [
            new ExecutiveSummaryRow
            {
                Category = "Financial Risk",
                Label = $"Pooling charges increased for {poolingGroups:#,0} groups",
                Value = $"{highLossClients:#,0} EHC/Dental loss ratio records above 100%",
                Status = "Review recommended",
                Count = poolingGroups + highLossClients
            },
            new ExecutiveSummaryRow
            {
                Category = "Client Risk",
                Label = $"{highIncrease:#,0} renewals exceed 20% increase",
                Value = $"{FormatMoney(renewals.Where(item => item.PremiumChangePercent >= 20m).SumValue(item => item.RenewalPremium ?? item.CurrentPremium))} premium impacted",
                Status = "Review recommended",
                Count = highIncrease
            },
            new ExecutiveSummaryRow
            {
                Category = "Opportunity",
                Label = $"Negotiated savings YTD: {FormatMoney(savings)}",
                Value = $"Across {versions.Count(item => item.VersionName.Contains("Negotiated", StringComparison.OrdinalIgnoreCase) || item.VersionName.Contains("Final", StringComparison.OrdinalIgnoreCase)):#,0} saved renewal versions",
                Status = "Savings opportunity",
                Count = renewals.Count(item => item.PremiumChangePercent >= 10m || item.EhcLossRatio >= 85m || item.DentalLossRatio >= 85m)
            },
            new ExecutiveSummaryRow
            {
                Category = "Operations",
                Label = $"{openProposalWork:#,0} proposals or quote options waiting",
                Value = $"{renewals.Count(item => item.Status.Equals(RenewalQueueStatuses.MissingData, StringComparison.OrdinalIgnoreCase)):#,0} renewals have missing source data",
                Status = "Operational follow-up",
                Count = openProposalWork
            }
        ];
    }

    private static IEnumerable<ExecutiveSummaryRow> BuildExecutiveDecisionRows(
        IReadOnlyCollection<RenewalQueueRecord> renewals,
        IReadOnlyCollection<ClaimsExperienceRecord> claims,
        IReadOnlyCollection<ExecutiveProposalRecord> proposals,
        IReadOnlyCollection<ExecutiveQuoteScenarioRecord> quotes)
    {
        var asoCandidates = renewals.Count(item => item.Lives >= 100 && item.EhcLossRatio < 85m && item.PremiumChangePercent >= 5m);
        var highRisk = renewals.Count(IsExecutiveRiskRenewal);
        var marketSurvey = renewals.Count(item => item.PremiumChangePercent >= 15m || item.EhcLossRatio >= 100m || item.DentalLossRatio >= 100m);
        var delayed = proposals.Count(item => !IsClosedStatus(item.Status) && (DateTime.Today - item.DateLoaded.Date).Days > 7) +
                      quotes.Count(item => !IsClosedStatus(item.Status) && (DateTime.Today - item.CreatedAt.Date).Days > 7);

        return
        [
            Summary("Executive Decision", $"Approve ASO review for {asoCandidates:#,0} clients", "Action", "Potential savings review"),
            Summary("Executive Decision", $"Review {highRisk:#,0} high-risk renewals", "Action", "Executive attention"),
            Summary("Executive Decision", $"Approve market survey strategy for {marketSurvey:#,0} renewals", "Action", "Carrier strategy"),
            Summary("Executive Decision", $"Review {delayed:#,0} delayed implementation or proposal items", "Action", "Operations")
        ];
    }

    private static IEnumerable<ExecutiveSummaryRow> BuildClientRiskWatchlistRows(IReadOnlyCollection<RenewalQueueRecord> renewals)
    {
        var rows = renewals
            .Where(item => item.PremiumChangePercent >= 10m || item.EhcLossRatio >= 85m || item.DentalLossRatio >= 85m || item.Status.Equals(RenewalQueueStatuses.HighRisk, StringComparison.OrdinalIgnoreCase))
            .OrderByDescending(item => RiskWeight(item))
            .ThenBy(item => item.RenewalDate ?? DateOnly.MaxValue)
            .Take(8)
            .Select(item => new ExecutiveSummaryRow
            {
                Category = item.PremiumChangePercent >= 20m || item.EhcLossRatio >= 100m || item.DentalLossRatio >= 100m ? "High" : "Medium",
                Label = item.ClientName,
                Value = FormatMoney(item.RenewalPremium ?? item.CurrentPremium),
                Status = $"{RiskReason(item)} | EHC {FormatPercent(item.EhcLossRatio, includeSign: false)} | Dental {FormatPercent(item.DentalLossRatio, includeSign: false)}",
                Count = (int)RiskWeight(item)
            })
            .ToList();

        return rows.Count == 0 ? [Summary("Watchlist", "No high-risk records found", "N/A", "Healthy")] : rows;
    }

    private static IEnumerable<ExecutivePipelineRow> BuildRenewalPipelineRows(IReadOnlyCollection<RenewalQueueRecord> renewals)
    {
        var stages = new[]
        {
            "Not started",
            "Documents received",
            "Extracted",
            "Validation required",
            "Underwriting review",
            "Consultant strategy review",
            "Presentation ready",
            "Delivered",
            "Completed"
        };

        return BuildPipelineRows(
            stages,
            renewals,
            MapRenewalStage,
            item => item.RenewalPremium ?? item.CurrentPremium,
            item => item.DateLoaded == DateTime.MinValue ? null : item.DateLoaded,
            item => FirstNonBlank(item.ClientName, item.Carrier, "Renewal"));
    }

    private static IEnumerable<ExecutivePipelineRow> BuildProposalPipelineRows(IReadOnlyCollection<ExecutiveProposalRecord> proposals)
    {
        var stages = new[]
        {
            "Intake received",
            "Carrier quotes pending",
            "Quotes received",
            "Comparison generated",
            "Underwriting review",
            "Consultant review",
            "Presentation ready",
            "Finalist selected",
            "Sold / Lost"
        };

        return BuildPipelineRows(
            stages,
            proposals,
            item => MapProposalStage(item.Status),
            item => item.ProposedPremium ?? item.CurrentPremium,
            item => item.DateLoaded == DateTime.MinValue ? null : item.DateLoaded,
            item => FirstNonBlank(item.ClientName, item.Carrier, "Proposal"));
    }

    private static IEnumerable<ExecutivePipelineRow> BuildQuotePipelineRows(IReadOnlyCollection<ExecutiveQuoteScenarioRecord> quotes)
    {
        var stages = new[]
        {
            "Request received",
            "Carrier quote pending",
            "Factor received",
            "Cost impact calculated",
            "Underwriting review",
            "Consultant review",
            "Client decision pending",
            "Completed"
        };

        return BuildPipelineRows(
            stages,
            quotes,
            item => MapQuoteStage(item.Status),
            item => item.EstimatedPremium ?? item.CurrentPremium,
            item => item.CreatedAt == DateTime.MinValue ? null : item.CreatedAt,
            item => FirstNonBlank(item.ScenarioName, item.Carrier, "Quote / option"));
    }

    private static IEnumerable<ExecutivePipelineRow> BuildPipelineRows<T>(
        IReadOnlyList<string> stages,
        IEnumerable<T> records,
        Func<T, string> stageSelector,
        Func<T, decimal?> premiumSelector,
        Func<T, DateTime?> dateSelector,
        Func<T, string> labelSelector)
    {
        var rows = new List<ExecutivePipelineRow>();
        var materialized = records.ToList();
        foreach (var stage in stages)
        {
            var stageRecords = materialized
                .Where(item => stageSelector(item).Equals(stage, StringComparison.OrdinalIgnoreCase))
                .ToList();
            var oldest = stageRecords
                .Select(item => new { Date = dateSelector(item), Label = labelSelector(item) })
                .Where(item => item.Date.HasValue)
                .OrderBy(item => item.Date)
                .FirstOrDefault();
            var days = stageRecords
                .Select(item => dateSelector(item))
                .Where(date => date.HasValue)
                .Select(date => Math.Max(0, (DateTime.Today - date!.Value.Date).Days))
                .ToList();

            rows.Add(new ExecutivePipelineRow
            {
                Stage = stage,
                Count = stageRecords.Count,
                Premium = stageRecords.SumValue(premiumSelector),
                OldestItem = oldest is null ? "N/A" : $"{oldest.Label} ({oldest.Date:MMM d})",
                AverageDaysInStage = days.Count == 0 ? "N/A" : ((int)Math.Round(days.Average(), MidpointRounding.AwayFromZero)).ToString(CultureInfo.InvariantCulture),
                OverdueItems = days.Count(day => day > 14)
            });
        }

        return rows;
    }

    private static IEnumerable<ExecutiveChartRow> BuildRenewalsByMonthRows(IReadOnlyCollection<RenewalQueueRecord> renewals)
    {
        var grouped = renewals
            .Where(item => item.RenewalDate.HasValue)
            .GroupBy(item => new DateOnly(item.RenewalDate!.Value.Year, item.RenewalDate.Value.Month, 1))
            .OrderBy(group => group.Key)
            .Select(group => new
            {
                Month = group.Key,
                Count = group.Count(),
                Premium = group.SumValue(item => item.RenewalPremium ?? item.CurrentPremium)
            })
            .TakeLast(12)
            .ToList();
        var maxPremium = grouped.Select(item => item.Premium ?? 0m).DefaultIfEmpty(0m).Max();
        var rows = grouped
            .Select(item => new ExecutiveChartRow
            {
                Label = item.Month.ToString("MMM yyyy", CultureInfo.InvariantCulture),
                Count = item.Count,
                Value = $"{item.Count} renewal(s)",
                Status = FormatMoneyCompact(item.Premium),
                Percent = maxPremium > 0m ? Math.Clamp((item.Premium ?? 0m) / maxPremium * 100m, 4m, 100m) : 0m
            })
            .ToList();

        return rows.Count == 0 ? [ExecutiveChartRow.Empty("Renewals by Month")] : rows;
    }

    private static IEnumerable<ExecutiveChartRow> BuildPremiumByCarrierRows(IReadOnlyCollection<RenewalQueueRecord> renewals)
    {
        var totalPremium = renewals.Sum(item => item.RenewalPremium ?? item.CurrentPremium ?? 0m);
        var rows = renewals
            .GroupBy(item => FirstNonBlank(item.Carrier, "Unassigned Carrier"), StringComparer.OrdinalIgnoreCase)
            .Select(group => new
            {
                Carrier = group.Key,
                Premium = group.Sum(item => item.RenewalPremium ?? item.CurrentPremium ?? 0m),
                Count = group.Count()
            })
            .OrderByDescending(item => item.Premium)
            .Take(8)
            .Select(item =>
            {
                var percent = totalPremium > 0m ? item.Premium / totalPremium * 100m : 0m;
                return new ExecutiveChartRow
                {
                    Label = item.Carrier,
                    Count = item.Count,
                    Value = FormatMoneyCompact(item.Premium),
                    Status = $"{percent:0.0}%  {FormatMoneyCompact(item.Premium)}",
                    Percent = percent
                };
            })
            .ToList();

        return rows.Count == 0 ? [ExecutiveChartRow.Empty("Premium by Carrier")] : rows;
    }

    private static IEnumerable<ExecutiveChartRow> BuildRenewalCalendarRows(IReadOnlyCollection<RenewalQueueRecord> renewals)
    {
        var today = DateOnly.FromDateTime(DateTime.Today);
        var grouped = renewals
            .Where(item => item.RenewalDate.HasValue && item.RenewalDate.Value >= today && item.RenewalDate.Value <= today.AddDays(90))
            .GroupBy(item => new DateOnly(item.RenewalDate!.Value.Year, item.RenewalDate.Value.Month, 1))
            .OrderBy(group => group.Key)
            .Take(4)
            .Select(group => new
            {
                Month = group.Key,
                Count = group.Count(),
                Premium = group.SumValue(item => item.RenewalPremium ?? item.CurrentPremium)
            })
            .ToList();
        var maxPremium = grouped.Select(item => item.Premium ?? 0m).DefaultIfEmpty(0m).Max();
        var rows = grouped
            .Select(group => new ExecutiveChartRow
            {
                Label = group.Month.ToString("MMM", CultureInfo.InvariantCulture),
                Count = group.Count,
                Value = $"{group.Count} renewals",
                Status = FormatMoneyCompact(group.Premium),
                Percent = maxPremium > 0m ? Math.Clamp((group.Premium ?? 0m) / maxPremium * 82m, 16m, 82m) : 16m
            })
            .ToList();

        return rows.Count == 0 ? [ExecutiveChartRow.Empty("Next 90 Days")] : rows;
    }

    private static IEnumerable<ExecutiveSummaryRow> BuildRenewalDistributionRows(IReadOnlyCollection<RenewalQueueRecord> renewals)
    {
        var ranges = new (string Label, Func<decimal, bool> Predicate)[]
        {
            ("< 0%", value => value < 0m),
            ("0-5%", value => value >= 0m && value < 5m),
            ("5-10%", value => value >= 5m && value < 10m),
            ("10-15%", value => value >= 10m && value < 15m),
            ("15-20%", value => value >= 15m && value < 20m),
            ("> 20%", value => value >= 20m)
        };

        var increases = renewals
            .Select(item => item.PremiumChangePercent)
            .Where(item => item.HasValue)
            .Select(item => item!.Value)
            .ToList();
        if (increases.Count == 0)
        {
            return [Summary("Distribution", "No saved renewal increases", "N/A", "Ready")];
        }

        return ranges.Select(range =>
        {
            var count = increases.Count(range.Predicate);
            var percent = count * 100m / increases.Count;
            return new ExecutiveSummaryRow
            {
                Category = "Distribution",
                Label = range.Label,
                Value = count.ToString("#,0", CultureInfo.InvariantCulture),
                Status = $"{count:#,0}  {percent:0}%",
                Count = (int)Math.Round(percent, MidpointRounding.AwayFromZero)
            };
        }).ToList();
    }

    private static IEnumerable<ExecutiveSummaryRow> BuildRenewalDriverRows(
        IReadOnlyCollection<RenewalQueueRecord> renewals,
        IReadOnlyCollection<ClaimsExperienceRecord> claims,
        IReadOnlyCollection<RenewalAssumptionRecord> assumptions)
    {
        // Derived placeholders until explicit renewal-driver attribution is stored in SQLite.
        var drugTrend = AverageNormalizedRate(assumptions.Where(item => IsBenefit(item.BenefitLine, "drug")).Select(item => item.Trend));
        var dentalSignal = AverageRaw(renewals.Select(item => item.DentalLossRatio).Concat(claims.Where(item => IsBenefit(item.BenefitLine, "dental")).Select(item => item.LossRatio)));
        var poolingSignal = AverageRaw(renewals.Select(item => item.PoolingCharge).Concat(assumptions.Select(item => item.PoolingCharge)));
        var demographicSignal = AverageRaw(renewals.Select(item => item.Lives));

        return
        [
            new ExecutiveSummaryRow { Category = "Driver", Label = "Drug Trend", Value = drugTrend.HasValue ? FormatPercent(drugTrend, includeSign: false) : "N/A", Status = "Derived from assumptions", Count = DriverPercent(drugTrend, 46) },
            new ExecutiveSummaryRow { Category = "Driver", Label = "Dental Trend", Value = FormatPercent(dentalSignal, includeSign: false), Status = "Loss ratio signal", Count = DriverPercent(dentalSignal, 23) },
            new ExecutiveSummaryRow { Category = "Driver", Label = "Pooling Change", Value = FormatNormalizedPercent(poolingSignal, includeSign: false), Status = "Pooling signal", Count = DriverPercent(poolingSignal, 18) },
            new ExecutiveSummaryRow { Category = "Driver", Label = "Demographics", Value = demographicSignal.HasValue ? $"{demographicSignal:0} avg lives" : "N/A", Status = "Lives signal", Count = 13 }
        ];
    }

    private static IEnumerable<ExecutiveChartRow> BuildRenewalsByCarrierRows(IReadOnlyCollection<RenewalQueueRecord> renewals)
    {
        var rows = renewals
            .GroupBy(item => FirstNonBlank(item.Carrier, "Unassigned Carrier"), StringComparer.OrdinalIgnoreCase)
            .OrderByDescending(group => group.Count())
            .Select(group => new ExecutiveChartRow
            {
                Label = group.Key,
                Count = group.Count(),
                Value = $"{group.Count()} renewal(s)",
                Status = $"Avg {FormatPercent(AverageRaw(group.Select(item => item.PremiumChangePercent)))}"
            })
            .Take(12)
            .ToList();

        return rows.Count == 0 ? [ExecutiveChartRow.Empty("Renewals by Carrier")] : rows;
    }

    private static IEnumerable<ExecutiveChartRow> BuildRenewalsByConsultantRows(IReadOnlyCollection<RenewalQueueRecord> renewals, IReadOnlyCollection<ClientRecord> clients)
    {
        var consultants = clients.ToDictionary(item => item.ClientId, item => FirstNonBlank(item.Advisor, "Unassigned"), EqualityComparer<long>.Default);
        var rows = renewals
            .GroupBy(item => consultants.TryGetValue(item.ClientId, out var consultant) ? consultant : "Unassigned", StringComparer.OrdinalIgnoreCase)
            .OrderByDescending(group => group.Count())
            .Select(group => new ExecutiveChartRow
            {
                Label = group.Key,
                Count = group.Count(),
                Value = $"{group.Count()} renewal(s)",
                Status = $"{group.Count(item => item.Status.Equals(RenewalQueueStatuses.HighRisk, StringComparison.OrdinalIgnoreCase))} high risk"
            })
            .Take(12)
            .ToList();

        return rows.Count == 0 ? [ExecutiveChartRow.Empty("Renewals by Consultant")] : rows;
    }

    private static IEnumerable<ExecutiveSummaryRow> BuildRenewalAnalyticsRows(IReadOnlyCollection<RenewalQueueRecord> renewals)
    {
        var increases = renewals.Select(item => item.PremiumChangePercent).Where(item => item.HasValue).Select(item => item!.Value).ToList();
        var rows = new List<ExecutiveSummaryRow>
        {
            Summary("Renewal Analytics", "Average renewal %", FormatPercent(AverageRaw(increases)), "Book average"),
            Summary("Renewal Analytics", "Median renewal %", FormatPercent(Median(increases)), "Distribution midpoint"),
            Summary("Renewal Analytics", "Largest increase", FormatPercent(increases.Count == 0 ? null : increases.Max()), "Largest increases"),
            Summary("Renewal Analytics", "Largest decrease", FormatPercent(increases.Count == 0 ? null : increases.Min()), "Largest decreases"),
            Summary("Renewal Analytics", "Renewals requiring negotiation", renewals.Count(item => item.PremiumChangePercent >= 10m || item.EhcLossRatio >= 100m || item.DentalLossRatio >= 100m).ToString("#,0", CultureInfo.InvariantCulture), "Review required"),
            Summary("Renewal Analytics", "Renewals under rate guarantee", "N/A", "Not tracked yet"),
            Summary("Renewal Analytics", "Renewals with caps", "N/A", "Not tracked yet"),
            Summary("Renewal Analytics", "Renewals with pooling increases", renewals.Count(item => item.Status.Equals(RenewalQueueStatuses.HighRisk, StringComparison.OrdinalIgnoreCase)).ToString("#,0", CultureInfo.InvariantCulture), "Risk proxy"),
            Summary("Renewal Analytics", "Trend over 5 years", "Ready", "Uses saved history")
        };

        return rows;
    }

    private static IEnumerable<ExecutiveSummaryRow> BuildClaimsAnalyticsRows(
        IReadOnlyCollection<ClaimsExperienceRecord> claims,
        IReadOnlyCollection<MonthlyClaimsRecord> monthlyClaims)
    {
        var benefitLines = new[]
        {
            "Healthcare",
            "Dental",
            "Life",
            "LTD",
            "STD",
            "Critical Illness",
            "Drug",
            "Vision",
            "Hospital",
            "Mental Health",
            "Paramedical"
        };

        var rows = new List<ExecutiveSummaryRow>();
        foreach (var benefit in benefitLines)
        {
            var claimRows = claims.Where(item => IsBenefit(item.BenefitLine, benefit)).ToList();
            var monthlyRows = monthlyClaims.Where(item => IsBenefit(item.BenefitLine, benefit)).ToList();
            var totalClaims = claimRows.SumValue(item => item.IncurredClaims ?? item.PaidClaims) ??
                              monthlyRows.SumValue(item => item.IncurredClaims ?? item.PaidClaims);
            var totalPremium = claimRows.SumValue(item => item.Premium) ?? monthlyRows.SumValue(item => item.Premium);
            var totalLives = claimRows.Sum(item => item.Lives ?? 0);
            if (totalLives == 0)
            {
                totalLives = monthlyRows.Sum(item => item.Lives ?? 0);
            }

            var pmpm = totalClaims.HasValue && totalLives > 0 ? totalClaims / totalLives / 12m : null;
            var lossRatio = totalPremium is > 0m && totalClaims.HasValue
                ? totalClaims / totalPremium * 100m
                : AverageRaw(claimRows.Select(item => item.LossRatio).Concat(monthlyRows.Select(item => item.LossRatio)));
            var trend = CalculateClaimsTrend(monthlyRows);
            rows.Add(new ExecutiveSummaryRow
            {
                Category = "Claims Analytics",
                Label = benefit,
                Count = claimRows.Count + monthlyRows.Count,
                Value = pmpm.HasValue ? $"PMPM {FormatMoney(pmpm)}" : "No saved claims",
                Status = $"{FormatPercent(lossRatio, includeSign: false)} LR | {FormatPercent(trend)} YOY"
            });
        }

        return rows;
    }

    private static IEnumerable<ExecutiveSummaryRow> BuildUnderwritingAnalyticsRows(
        IReadOnlyCollection<RenewalAssumptionRecord> assumptions,
        IReadOnlyCollection<RenewalVersionRecord> versions)
    {
        var calculations = versions.SelectMany(item => item.Calculations).ToList();
        return
        [
            Summary("Underwriting Analytics", "Loss Ratios", FormatPercent(AverageRaw(calculations.Select(item => item.ProjectedLossRatio)), includeSign: false), "Projected average"),
            Summary("Underwriting Analytics", "Target Loss Ratios", FormatNormalizedPercent(AverageRaw(assumptions.Select(item => item.TargetLossRatio)), includeSign: false), "TLR"),
            Summary("Underwriting Analytics", "Pooling Charges", FormatNormalizedPercent(AverageRaw(assumptions.Select(item => item.PoolingCharge)), includeSign: false), "Pooling"),
            Summary("Underwriting Analytics", "IBNR", FormatNormalizedPercent(AverageRaw(assumptions.Select(item => item.Ibnr)), includeSign: false), "Reserve"),
            Summary("Underwriting Analytics", "Trend", FormatNormalizedPercent(AverageRaw(assumptions.Select(item => item.Trend)), includeSign: false), "Manual trend"),
            Summary("Underwriting Analytics", "Expense Load", FormatNormalizedPercent(AverageRaw(assumptions.Select(item => (item.AdminFee ?? 0m) + (item.RiskCharge ?? 0m))), includeSign: false), "Admin + risk"),
            Summary("Underwriting Analytics", "Retention", "N/A", "Carrier table future"),
            Summary("Underwriting Analytics", "Credibility", "N/A", "Carrier table future"),
            Summary("Underwriting Analytics", "Renewal Factors", versions.Count.ToString("#,0", CultureInfo.InvariantCulture), "Saved scenarios"),
            Summary("Underwriting Analytics", "Rate Guarantees", "N/A", "Not tracked yet"),
            Summary("Underwriting Analytics", "Scenario Outcomes", calculations.Count.ToString("#,0", CultureInfo.InvariantCulture), "Stored calculations")
        ];
    }

    private static IEnumerable<ExecutiveSummaryRow> BuildCarrierAnalyticsRows(IReadOnlyCollection<RenewalQueueRecord> renewals)
    {
        var rows = renewals
            .GroupBy(item => FirstNonBlank(item.Carrier, "Unassigned Carrier"), StringComparer.OrdinalIgnoreCase)
            .OrderByDescending(group => group.Sum(item => item.RenewalPremium ?? item.CurrentPremium ?? 0m))
            .Select(group => new ExecutiveSummaryRow
            {
                Category = "Carrier Intelligence",
                Label = group.Key,
                Count = group.Count(),
                Value = FormatPercent(AverageRaw(group.Select(item => item.PremiumChangePercent))),
                Status = $"LR {FormatPercent(AverageRaw(group.Select(item => item.EhcLossRatio).Concat(group.Select(item => item.DentalLossRatio))), includeSign: false)} | {FormatMoney(group.SumValue(item => item.RenewalPremium ?? item.CurrentPremium))}"
            })
            .ToList();

        return rows.Count == 0 ? [Summary("Carrier Intelligence", "Carrier comparison", "No saved renewals", "Ready")] : rows;
    }

    private static IEnumerable<ExecutiveSummaryRow> BuildProposalAnalyticsRows(IReadOnlyCollection<ExecutiveProposalRecord> proposals)
    {
        var sold = proposals.Count(item => item.Status.Contains("sold", StringComparison.OrdinalIgnoreCase) || item.Status.Contains("won", StringComparison.OrdinalIgnoreCase));
        var lost = proposals.Count(item => item.Status.Contains("lost", StringComparison.OrdinalIgnoreCase));
        var pending = proposals.Count(item => !IsClosedStatus(item.Status));
        var rows = new List<ExecutiveSummaryRow>
        {
            Summary("Proposal Analytics", "Proposal wins", sold.ToString("#,0", CultureInfo.InvariantCulture), "Sold / won"),
            Summary("Proposal Analytics", "Carrier success", proposals.GroupBy(item => item.Carrier).OrderByDescending(group => group.Count()).FirstOrDefault()?.Key ?? "No carrier data", "Top carrier"),
            Summary("Proposal Analytics", "Average premium difference", FormatMoney(AverageRaw(proposals.Select(item => item.Savings))), "Savings field"),
            Summary("Proposal Analytics", "Savings achieved", FormatMoney(proposals.SumValue(item => item.Savings)), "Stored proposals"),
            Summary("Proposal Analytics", "Recommendation acceptance rate", sold + lost > 0 ? $"{sold * 100m / (sold + lost):0.0}%" : "N/A", "Sold / closed"),
            Summary("Proposal Analytics", "Alternate option selection", "N/A", "Quote module future"),
            Summary("Proposal Analytics", "Quote turnaround", "N/A", "Date stamps future"),
            Summary("Proposal Analytics", "Proposal turnaround", pending.ToString("#,0", CultureInfo.InvariantCulture), "Open proposals")
        };

        return rows;
    }

    private static IEnumerable<ExecutiveSummaryRow> BuildPlanDesignAnalyticsRows(IReadOnlyCollection<PlanDesignRecord> planDesign)
    {
        return
        [
            Summary("Plan Design Analytics", "Most common deductible", CommonPlanValue(planDesign, "deductible"), "Benchmark ready"),
            Summary("Plan Design Analytics", "Most common drug reimbursement", CommonPlanValue(planDesign, "drug", preferReimbursement: true), "Benchmark ready"),
            Summary("Plan Design Analytics", "Mental Health limits", CommonPlanValue(planDesign, "mental"), "Benchmark ready"),
            Summary("Plan Design Analytics", "Paramedical limits", CommonPlanValue(planDesign, "paramedical"), "Benchmark ready"),
            Summary("Plan Design Analytics", "Vision limits", CommonPlanValue(planDesign, "vision"), "Benchmark ready"),
            Summary("Plan Design Analytics", "Dental reimbursement", CommonPlanValue(planDesign, "dental", preferReimbursement: true), "Benchmark ready"),
            Summary("Plan Design Analytics", "Hospital coverage", CommonPlanValue(planDesign, "hospital"), "Benchmark ready"),
            Summary("Plan Design Analytics", "Life Maximums", CommonPlanValue(planDesign, "life"), "Benchmark ready"),
            Summary("Plan Design Analytics", "LTD Maximums", CommonPlanValue(planDesign, "ltd", "long term disability"), "Benchmark ready")
        ];
    }

    private static IEnumerable<ExecutiveSummaryRow> BuildBenchmarkRows(
        IReadOnlyCollection<ClientRecord> clients,
        IReadOnlyCollection<RenewalQueueRecord> renewals,
        IReadOnlyCollection<ClaimsExperienceRecord> claims)
    {
        return
        [
            Summary("Market Benchmarks", "Industry", clients.GroupBy(item => FirstNonBlank(item.Industry, "Unassigned")).OrderByDescending(group => group.Count()).FirstOrDefault()?.Key ?? "No saved industry", "Average"),
            Summary("Market Benchmarks", "Province", clients.GroupBy(item => FirstNonBlank(item.Province, "Unassigned")).OrderByDescending(group => group.Count()).FirstOrDefault()?.Key ?? "No saved province", "Average"),
            Summary("Market Benchmarks", "Company Size", AverageRaw(renewals.Select(item => item.Lives)).HasValue ? $"{AverageRaw(renewals.Select(item => item.Lives)):0} employees" : "N/A", "Average"),
            Summary("Market Benchmarks", "Carrier", renewals.GroupBy(item => FirstNonBlank(item.Carrier, "Unassigned")).OrderByDescending(group => group.Count()).FirstOrDefault()?.Key ?? "No saved carrier", "Average"),
            Summary("Market Benchmarks", "Funding Arrangement", "N/A", "Awaiting saved funding type"),
            Summary("Market Benchmarks", "Above Average", CountAboveAverage(renewals).ToString("#,0", CultureInfo.InvariantCulture), "Renewal increase"),
            Summary("Market Benchmarks", "Below Average", CountBelowAverage(renewals).ToString("#,0", CultureInfo.InvariantCulture), "Renewal increase"),
            Summary("Market Benchmarks", "Benchmark Percentile", claims.Count > 0 ? "Ready" : "No saved claims", "Stored history")
        ];
    }

    private static IEnumerable<ExecutiveSummaryRow> BuildRoleActionRows(
        string role,
        IReadOnlyCollection<RenewalQueueRecord> renewals,
        IReadOnlyCollection<ClaimsExperienceRecord> claims,
        IReadOnlyCollection<ExecutiveProposalRecord> proposals,
        IReadOnlyCollection<ExecutiveQuoteScenarioRecord> quotes)
    {
        if (role.Contains("Underwriting", StringComparison.OrdinalIgnoreCase))
        {
            return
            [
                ActionRow("Validation required", renewals.Count(item => item.Status.Equals(RenewalQueueStatuses.MissingData, StringComparison.OrdinalIgnoreCase)), "Underwriting", "High"),
                ActionRow("Rate issues", renewals.Count(item => item.PremiumChangePercent >= 10m), "Underwriting", "Medium"),
                ActionRow("Claims issues", claims.Count(item => item.LossRatio >= 100m), "Underwriting", "High"),
                ActionRow("Pooling concerns", renewals.Count(item => item.Status.Equals(RenewalQueueStatuses.HighRisk, StringComparison.OrdinalIgnoreCase)), "Underwriting", "High"),
                ActionRow("Trend concerns", claims.Count(item => item.LossRatio >= 90m), "Underwriting", "Medium"),
                ActionRow("Quotes requiring review", quotes.Count(item => !IsClosedStatus(item.Status)), "Underwriting", "Medium")
            ];
        }

        if (role.Contains("Consultant", StringComparison.OrdinalIgnoreCase))
        {
            return
            [
                ActionRow("Clients requiring strategy", renewals.Count(item => item.PremiumChangePercent >= 10m), "Consultant", "High"),
                ActionRow("Client-ready presentations", renewals.Count(item => item.Status.Equals(RenewalQueueStatuses.ReadyToReview, StringComparison.OrdinalIgnoreCase)), "Consultant", "Medium"),
                ActionRow("Proposals pending", proposals.Count(item => !IsClosedStatus(item.Status)), "Consultant", "Medium"),
                ActionRow("Client decisions pending", quotes.Count(item => item.Status.Contains("decision", StringComparison.OrdinalIgnoreCase)), "Consultant", "Medium"),
                ActionRow("Upcoming renewal meetings", renewals.Count(item => IsDueWithin(item.RenewalDate, 60)), "Consultant", "High"),
                ActionRow("At-risk clients", renewals.Count(item => item.Status.Equals(RenewalQueueStatuses.HighRisk, StringComparison.OrdinalIgnoreCase)), "Consultant", "High")
            ];
        }

        return
        [
            ActionRow("Renewals due in 60 days", renewals.Count(item => IsDueWithin(item.RenewalDate, 60)), "Leadership", "High"),
            ActionRow("Proposal pipeline", proposals.Count(item => !IsClosedStatus(item.Status)), "Leadership", "Medium"),
            ActionRow("Consultant workload", renewals.Count, "Leadership", "Monitor"),
            ActionRow("Carrier performance", renewals.Select(item => item.Carrier).Distinct(StringComparer.OrdinalIgnoreCase).Count(), "Leadership", "Monitor"),
            ActionRow("Client health distribution", renewals.Count(item => item.Status.Equals(RenewalQueueStatuses.HighRisk, StringComparison.OrdinalIgnoreCase)), "Leadership", "High"),
            ActionRow("Savings generated", CalculateProposalSavingsCount(proposals), "Leadership", "Monitor")
        ];
    }

    private static IEnumerable<string> BuildInsightRows(
        IReadOnlyCollection<RenewalQueueRecord> renewals,
        IReadOnlyCollection<ClaimsExperienceRecord> claims,
        IReadOnlyCollection<ExecutiveProposalRecord> proposals,
        IReadOnlyCollection<ClientRecord> clients)
    {
        var insights = new List<string>();
        var due60NotStarted = renewals.Count(item => IsDueWithin(item.RenewalDate, 60) && MapRenewalStage(item).Equals("Not started", StringComparison.OrdinalIgnoreCase));
        insights.Add($"{due60NotStarted} renewals are due within 60 days and have not been started.");
        insights.Add($"{renewals.Count(item => item.PremiumChangePercent > 15m)} clients have renewal increases above 15%.");
        insights.Add($"{proposals.Count(item => !IsClosedStatus(item.Status) && (DateTime.Today - item.DateLoaded.Date).Days > 10)} proposals have been waiting more than 10 days.");
        insights.Add($"{renewals.Count(item => item.Status.Equals(RenewalQueueStatuses.Watch, StringComparison.OrdinalIgnoreCase))} clients are currently in Monitor status.");

        var bookAverage = AverageRaw(renewals.Select(item => item.PremiumChangePercent));
        var carrierAboveAverage = renewals
            .Where(item => item.PremiumChangePercent.HasValue && bookAverage.HasValue)
            .GroupBy(item => FirstNonBlank(item.Carrier, "Unassigned"))
            .Select(group => new { Carrier = group.Key, Average = AverageRaw(group.Select(item => item.PremiumChangePercent)) })
            .Where(item => item.Average.HasValue && bookAverage.HasValue && item.Average > bookAverage)
            .OrderByDescending(item => item.Average)
            .FirstOrDefault();
        insights.Add(carrierAboveAverage is null
            ? "Carrier performance will populate as more renewal history is saved."
            : $"{carrierAboveAverage.Carrier} renewals are averaging {carrierAboveAverage.Average - bookAverage:0.0}% higher than the book average.");

        insights.Add($"{claims.Count(item => IsBenefit(item.BenefitLine, "paramedical") && item.LossRatio >= 90m)} clients may have paramedical claims above benchmark.");
        insights.Add($"{clients.Count(client => renewals.Any(renewal => renewal.ClientId == client.ClientId && renewal.Lives >= 100 && renewal.EhcLossRatio < 80m))} clients may be good ASO candidates.");
        return insights;
    }

    private static IEnumerable<ExecutiveSummaryRow> BuildDefaultDrillRows(
        IReadOnlyCollection<RenewalQueueRecord> renewals,
        IReadOnlyCollection<ExecutiveProposalRecord> proposals,
        IReadOnlyCollection<ExecutiveQuoteScenarioRecord> quotes)
    {
        var rows = new List<ExecutiveSummaryRow>();
        rows.AddRange(renewals
            .OrderBy(item => item.RenewalDate ?? DateOnly.MaxValue)
            .Take(12)
            .Select(item => Summary("Renewal", item.ClientName, FormatMoney(item.RenewalPremium ?? item.CurrentPremium), item.Status)));
        rows.AddRange(proposals
            .OrderByDescending(item => item.DateLoaded)
            .Take(8)
            .Select(item => Summary("Proposal", FirstNonBlank(item.ClientName, item.Carrier), FormatMoney(item.ProposedPremium), item.Status)));
        rows.AddRange(quotes
            .OrderByDescending(item => item.CreatedAt)
            .Take(8)
            .Select(item => Summary("Quote / Option", FirstNonBlank(item.ScenarioName, item.Carrier), FormatMoney(item.EstimatedPremium), item.Status)));

        return rows.Count == 0 ? [Summary("Master Benefits Database", "No saved drill-down records", "N/A", "Ready")] : rows;
    }

    private void UpdateReportsClientDashboard(ReportsFilterCriteria filter, IReadOnlyCollection<ClientRecord> clients)
    {
        if (_reportsSnapshot is null)
        {
            return;
        }

        var selected = ReportsSelectedClientComboBox.SelectedItem as ClientRecord;
        if (selected is null || clients.All(item => item.ClientId != selected.ClientId))
        {
            selected = clients.OrderBy(item => item.ClientName).FirstOrDefault() ?? _reportsSnapshot.Clients.OrderBy(item => item.ClientName).FirstOrDefault();
            _isUpdatingReportsFilters = true;
            ReportsSelectedClientComboBox.SelectedItem = selected;
            _isUpdatingReportsFilters = false;
        }

        _reportsClientSummaryRows.Clear();
        _reportsClientHealthComponentRows.Clear();
        _reportsClientAnalyticsRows.Clear();
        _reportsClientHistoryRows.Clear();
        _reportsClientActionRows.Clear();

        if (selected is null)
        {
            ReportsClientHealthScoreTextBlock.Text = "N/A";
            ReportsClientHealthStatusTextBlock.Text = "No client selected";
            _reportsClientActionRows.Add("Select or save a client to build the client dashboard.");
            return;
        }

        var renewals = FilterReportsRenewals(_reportsSnapshot, filter, [selected.ClientId]);
        var clientClaims = _reportsSnapshot.Claims.Where(item => item.ClientId == selected.ClientId).ToList();
        var monthlyClaims = _reportsSnapshot.MonthlyClaims.Where(item => item.ClientId == selected.ClientId).ToList();
        var planDesign = _reportsSnapshot.PlanDesign.Where(item => item.ClientId == selected.ClientId).ToList();
        var assumptions = _reportsSnapshot.RenewalAssumptions.Where(item => item.ClientId == selected.ClientId).ToList();
        var versions = _reportsSnapshot.RenewalVersions.Where(item => item.ClientId == selected.ClientId).ToList();
        var proposals = _reportsSnapshot.Proposals.Where(item => item.ClientName.Equals(selected.ClientName, StringComparison.OrdinalIgnoreCase)).ToList();
        var latestRenewal = renewals.OrderByDescending(item => item.RenewalDate ?? DateOnly.MinValue).FirstOrDefault();

        ReplaceCollection(_reportsClientSummaryRows,
        [
            Summary("Client Summary", "Client name", selected.ClientName, selected.Status),
            Summary("Client Summary", "Industry", FirstNonBlank(selected.Industry, "N/A"), "Profile"),
            Summary("Client Summary", "Province", FirstNonBlank(selected.Province, "N/A"), "Profile"),
            Summary("Client Summary", "Consultant", FirstNonBlank(selected.Advisor, "N/A"), "Owner"),
            Summary("Client Summary", "Carrier", FirstNonBlank(latestRenewal?.Carrier, selected.Carrier, "N/A"), "Current carrier"),
            Summary("Client Summary", "Funding type", "N/A", "Not stored yet"),
            Summary("Client Summary", "Renewal date", latestRenewal?.RenewalDate?.ToString("MMM d yyyy", CultureInfo.InvariantCulture) ?? "N/A", "Next renewal"),
            Summary("Client Summary", "Number of employees", latestRenewal?.Lives?.ToString("#,0", CultureInfo.InvariantCulture) ?? "N/A", "Lives"),
            Summary("Client Summary", "Total premium", FormatMoney(latestRenewal?.RenewalPremium ?? latestRenewal?.CurrentPremium), "Annual"),
            Summary("Client Summary", "Current plan status", latestRenewal?.Status ?? "N/A", "Saved renewal")
        ]);

        var health = BuildClientHealthScore(selected, latestRenewal, renewals, clientClaims, monthlyClaims, planDesign, assumptions, proposals);
        ReportsClientHealthScoreTextBlock.Text = health.Score.ToString("0", CultureInfo.InvariantCulture);
        ReportsClientHealthStatusTextBlock.Text = health.Status;
        ReplaceCollection(_reportsClientHealthComponentRows, health.Components);

        ReplaceCollection(_reportsClientAnalyticsRows,
        [
            Summary("Client Analytics", "Premium history", renewals.Count.ToString("#,0", CultureInfo.InvariantCulture), FormatMoney(renewals.SumValue(item => item.CurrentPremium))),
            Summary("Client Analytics", "Claims history", (clientClaims.Count + monthlyClaims.Count).ToString("#,0", CultureInfo.InvariantCulture), FormatMoney(clientClaims.SumValue(item => item.IncurredClaims ?? item.PaidClaims))),
            Summary("Client Analytics", "Loss ratio history", FormatPercent(AverageRaw(clientClaims.Select(item => item.LossRatio)), includeSign: false), "Saved claims"),
            Summary("Client Analytics", "Renewal history", renewals.Count.ToString("#,0", CultureInfo.InvariantCulture), FormatPercent(AverageRaw(renewals.Select(item => item.PremiumChangePercent)))),
            Summary("Client Analytics", "Plan design history", planDesign.Count.ToString("#,0", CultureInfo.InvariantCulture), "Rows"),
            Summary("Client Analytics", "Carrier history", renewals.Select(item => item.Carrier).Distinct(StringComparer.OrdinalIgnoreCase).Count().ToString("#,0", CultureInfo.InvariantCulture), "Carriers"),
            Summary("Client Analytics", "Proposal history", proposals.Count.ToString("#,0", CultureInfo.InvariantCulture), "Proposals"),
            Summary("Client Analytics", "Quote / alternate option history", _reportsSnapshot.QuoteScenarios.Count(item => item.ClientId == selected.ClientId).ToString("#,0", CultureInfo.InvariantCulture), "Options")
        ]);

        ReplaceCollection(_reportsClientHistoryRows, BuildClientHistoryRows(selected, renewals, clientClaims, planDesign, assumptions, versions, proposals));
        ReplaceCollection(_reportsClientActionRows, BuildClientActionItems(latestRenewal, clientClaims, proposals, health));
    }

    private static ClientHealthResult BuildClientHealthScore(
        ClientRecord client,
        RenewalQueueRecord? latestRenewal,
        IReadOnlyCollection<RenewalQueueRecord> renewals,
        IReadOnlyCollection<ClaimsExperienceRecord> claims,
        IReadOnlyCollection<MonthlyClaimsRecord> monthlyClaims,
        IReadOnlyCollection<PlanDesignRecord> planDesign,
        IReadOnlyCollection<RenewalAssumptionRecord> assumptions,
        IReadOnlyCollection<ExecutiveProposalRecord> proposals)
    {
        var weights = new ExecutiveHealthWeights();
        var lossRatio = AverageRaw(claims.Select(item => item.LossRatio).Concat(renewals.Select(item => item.EhcLossRatio).Concat(renewals.Select(item => item.DentalLossRatio))));
        var renewalIncrease = latestRenewal?.PremiumChangePercent ?? AverageRaw(renewals.Select(item => item.PremiumChangePercent));
        var volatility = CalculateClaimsVolatility(monthlyClaims);
        var poolingRiskSignal = AverageRaw(assumptions.Select(item => item.PoolingCharge));
        var planScore = planDesign.Count > 0 ? 82m : 65m;
        var contributionScore = proposals.Any(item => item.EmployerCost.HasValue || item.EmployeeCost.HasValue) ? 78m : 70m;
        var marketScore = proposals.Count > 0 ? 74m : 80m;
        var consultantScore = latestRenewal?.Status switch
        {
            RenewalQueueStatuses.HighRisk => 35m,
            RenewalQueueStatuses.Watch => 65m,
            RenewalQueueStatuses.MissingData => 55m,
            _ => 85m
        };

        var lossScore = ScoreLossRatio(lossRatio);
        var renewalScore = ScoreRenewalIncrease(renewalIncrease);
        var volatilityScore = ScoreVolatility(volatility);
        var poolingScore = ScorePoolingRisk(poolingRiskSignal, lossRatio);

        var score =
            lossScore * weights.LossRatio +
            renewalScore * weights.RenewalIncrease +
            volatilityScore * weights.ClaimsVolatility +
            poolingScore * weights.PoolingLargeClaimsRisk +
            planScore * weights.PlanCompetitiveness +
            contributionScore * weights.ContributionCompetitiveness +
            marketScore * weights.ProposalMarketActivity +
            consultantScore * weights.ConsultantRiskFlag;

        var rounded = Math.Round(score, 0, MidpointRounding.AwayFromZero);
        var status = rounded >= 85m ? "Healthy" : rounded >= 70m ? "Monitor" : "At Risk";
        return new ClientHealthResult
        {
            Score = rounded,
            Status = status,
            Components =
            [
                Summary("Health Score", "Loss Ratio", "25%", $"{FormatPercent(lossRatio, includeSign: false)} | {lossScore:0}/100"),
                Summary("Health Score", "Renewal Increase", "20%", $"{FormatPercent(renewalIncrease)} | {renewalScore:0}/100"),
                Summary("Health Score", "Claims Volatility", "15%", $"{FormatPercent(volatility)} | {volatilityScore:0}/100"),
                Summary("Health Score", "Pooling / Large Claims Risk", "15%", $"{FormatNormalizedPercent(poolingRiskSignal, includeSign: false)} | {poolingScore:0}/100"),
                Summary("Health Score", "Plan Competitiveness", "10%", $"{planDesign.Count:#,0} saved rows | {planScore:0}/100"),
                Summary("Health Score", "Contribution Competitiveness", "5%", $"{(proposals.Count > 0 ? "Proposal data" : "No proposal data")} | {contributionScore:0}/100"),
                Summary("Health Score", "Proposal / Market Activity", "5%", $"{proposals.Count:#,0} proposal(s) | {marketScore:0}/100"),
                Summary("Health Score", "Consultant Risk Flag", "5%", $"{FirstNonBlank(latestRenewal?.Status, client.Status, "N/A")} | {consultantScore:0}/100")
            ]
        };
    }

    private static IEnumerable<ExecutiveSummaryRow> BuildClientHistoryRows(
        ClientRecord client,
        IReadOnlyCollection<RenewalQueueRecord> renewals,
        IReadOnlyCollection<ClaimsExperienceRecord> claims,
        IReadOnlyCollection<PlanDesignRecord> planDesign,
        IReadOnlyCollection<RenewalAssumptionRecord> assumptions,
        IReadOnlyCollection<RenewalVersionRecord> versions,
        IReadOnlyCollection<ExecutiveProposalRecord> proposals)
    {
        var rows = new List<ExecutiveSummaryRow>();
        rows.AddRange(renewals.Select(item => Summary("Renewal", item.RenewalDate?.ToString("MMM d yyyy", CultureInfo.InvariantCulture) ?? item.RenewalYear.ToString(CultureInfo.InvariantCulture), FormatMoney(item.RenewalPremium ?? item.CurrentPremium), item.Status)));
        rows.AddRange(claims.Take(15).Select(item => Summary("Claims", FirstNonBlank(item.BenefitLine, "Claims"), FormatMoney(item.IncurredClaims ?? item.PaidClaims), FormatPercent(item.LossRatio, includeSign: false))));
        rows.AddRange(planDesign.Take(15).Select(item => Summary("Plan Design", FirstNonBlank(item.BenefitName, item.BenefitCategory), FirstNonBlank(item.MaximumAmount, item.CoverageDescription, "Saved"), FormatPercent(item.ReimbursementPercent, includeSign: false))));
        rows.AddRange(assumptions.Take(12).Select(item => Summary("Underwriting", FirstNonBlank(item.BenefitLine, "Assumption"), FormatNormalizedPercent(item.TargetLossRatio, includeSign: false), FormatNormalizedPercent(item.Trend, includeSign: false))));
        rows.AddRange(versions.Take(12).Select(item => Summary("Negotiation", FirstNonBlank(item.VersionName, item.SourceModule), item.DateCreated.ToString("MMM d yyyy", CultureInfo.InvariantCulture), item.Status)));
        rows.AddRange(proposals.Take(12).Select(item => Summary("Proposal", FirstNonBlank(item.Carrier, "Carrier Proposal"), FormatMoney(item.ProposedPremium), item.Status)));

        return rows.Count == 0 ? [Summary("Client", client.ClientName, "No saved history", "Ready")] : rows;
    }

    private static IEnumerable<string> BuildClientActionItems(
        RenewalQueueRecord? latestRenewal,
        IReadOnlyCollection<ClaimsExperienceRecord> claims,
        IReadOnlyCollection<ExecutiveProposalRecord> proposals,
        ClientHealthResult health)
    {
        var actions = new List<string>();
        if (latestRenewal?.RenewalDate is not null)
        {
            actions.Add($"Upcoming renewal: {latestRenewal.RenewalDate:MMM d yyyy}.");
        }

        if (latestRenewal?.Status.Equals(RenewalQueueStatuses.MissingData, StringComparison.OrdinalIgnoreCase) == true)
        {
            actions.Add("Missing documents or validation required.");
        }

        if (claims.Any(item => item.LossRatio >= 100m))
        {
            actions.Add("High claims trend requires underwriting review.");
        }

        if (latestRenewal?.EhcLossRatio >= 100m || latestRenewal?.DentalLossRatio >= 100m)
        {
            actions.Add("Loss ratio or pooling concern needs strategy notes.");
        }

        if (proposals.Any(item => !IsClosedStatus(item.Status)))
        {
            actions.Add("Proposal follow-up is outstanding.");
        }

        if (health.Status.Equals("At Risk", StringComparison.OrdinalIgnoreCase))
        {
            actions.Add("Consultant review required before client presentation.");
        }

        if (actions.Count == 0)
        {
            actions.Add("No urgent client action items identified from saved history.");
        }

        return actions;
    }

    private static bool IsExecutiveRiskRenewal(RenewalQueueRecord item)
    {
        return item.Status.Equals(RenewalQueueStatuses.HighRisk, StringComparison.OrdinalIgnoreCase) ||
               item.PremiumChangePercent >= 10m ||
               item.EhcLossRatio >= 85m ||
               item.DentalLossRatio >= 85m ||
               (NormalizeRatePercent(item.PoolingCharge) ?? 0m) >= 5m;
    }

    private static decimal RiskWeight(RenewalQueueRecord item)
    {
        var weight = 0m;
        if (item.PremiumChangePercent.HasValue)
        {
            weight += Math.Max(0m, item.PremiumChangePercent.Value);
        }

        if (item.EhcLossRatio.HasValue)
        {
            weight += Math.Max(0m, item.EhcLossRatio.Value - 75m) / 2m;
        }

        if (item.DentalLossRatio.HasValue)
        {
            weight += Math.Max(0m, item.DentalLossRatio.Value - 75m) / 2m;
        }

        if (item.Status.Equals(RenewalQueueStatuses.HighRisk, StringComparison.OrdinalIgnoreCase))
        {
            weight += 25m;
        }

        return weight;
    }

    private static string RiskReason(RenewalQueueRecord item)
    {
        if (item.EhcLossRatio >= 100m)
        {
            return "EHC LR > 100%";
        }

        if (item.DentalLossRatio >= 100m)
        {
            return "Dental LR > 100%";
        }

        if (item.PremiumChangePercent >= 20m)
        {
            return "High renewal increase";
        }

        if ((NormalizeRatePercent(item.PoolingCharge) ?? 0m) >= 5m)
        {
            return "Pooling charge";
        }

        return "Claims trend rising";
    }

    private static decimal? EstimateForecastCommission(decimal? premium)
    {
        // Derived placeholder until the commission forecast table is connected to SQLite.
        return premium.HasValue ? premium.Value * 0.04m : null;
    }

    private static decimal? EstimatePotentialSavings(IReadOnlyCollection<RenewalQueueRecord> renewals, decimal? savedSavings)
    {
        if (savedSavings.HasValue && savedSavings.Value > 0m)
        {
            return savedSavings;
        }

        var atRiskPremium = renewals
            .Where(IsExecutiveRiskRenewal)
            .SumValue(item => item.RenewalPremium ?? item.CurrentPremium);
        return atRiskPremium.HasValue ? atRiskPremium.Value * 0.03m : null;
    }

    private static decimal CalculateRenewalRiskScore(IReadOnlyCollection<RenewalQueueRecord> renewals)
    {
        if (renewals.Count == 0)
        {
            return 82m;
        }

        var riskRate = renewals.Count(IsExecutiveRiskRenewal) / (decimal)renewals.Count;
        return Math.Clamp(100m - (riskRate * 95m), 35m, 96m);
    }

    private static decimal CalculateImplementationScore(
        IReadOnlyCollection<RenewalQueueRecord> renewals,
        IReadOnlyCollection<ExecutiveProposalRecord> proposals,
        IReadOnlyCollection<ExecutiveQuoteScenarioRecord> quotes)
    {
        var missing = renewals.Count(item => item.Status.Equals(RenewalQueueStatuses.MissingData, StringComparison.OrdinalIgnoreCase));
        var overdue = proposals.Count(item => !IsClosedStatus(item.Status) && (DateTime.Today - item.DateLoaded.Date).Days > 14) +
                      quotes.Count(item => !IsClosedStatus(item.Status) && (DateTime.Today - item.CreatedAt.Date).Days > 14);
        return Math.Clamp(96m - ((missing + overdue) * 4m), 55m, 96m);
    }

    private static decimal CalculateRetentionScore(IReadOnlyCollection<ClientRecord> clients)
    {
        if (clients.Count == 0)
        {
            return 90m;
        }

        var inactive = clients.Count(item => item.Status.Contains("Inactive", StringComparison.OrdinalIgnoreCase) ||
                                             item.Status.Contains("Lost", StringComparison.OrdinalIgnoreCase));
        return Math.Clamp(96m - (inactive / (decimal)clients.Count * 100m), 70m, 96m);
    }

    private static decimal CalculateForecastConfidence(
        IReadOnlyCollection<RenewalQueueRecord> renewals,
        IReadOnlyCollection<ClaimsExperienceRecord> claims,
        IReadOnlyCollection<RenewalAssumptionRecord> assumptions)
    {
        var signals = 0;
        if (renewals.Any(item => item.CurrentPremium.HasValue || item.RenewalPremium.HasValue))
        {
            signals++;
        }

        if (claims.Any(item => item.LossRatio.HasValue || item.PaidClaims.HasValue || item.IncurredClaims.HasValue))
        {
            signals++;
        }

        if (assumptions.Any(item => item.TargetLossRatio.HasValue || item.Trend.HasValue))
        {
            signals++;
        }

        return signals switch
        {
            >= 3 => 96m,
            2 => 88m,
            1 => 76m,
            _ => 64m
        };
    }

    private static string BookIndexStatus(decimal score)
    {
        return score switch
        {
            >= 90m => "Excellent",
            >= 80m => "Strong",
            >= 70m => "Monitor",
            _ => "At Risk"
        };
    }

    private static int DriverPercent(decimal? signal, int fallback)
    {
        if (!signal.HasValue)
        {
            return fallback;
        }

        return (int)Math.Clamp(Math.Round(Math.Abs(NormalizeRatePercent(signal) ?? signal.Value), 0, MidpointRounding.AwayFromZero), 8m, 80m);
    }

    private void ReportsDrillDown_Click(object sender, RoutedEventArgs e)
    {
        if (_reportsSnapshot is null)
        {
            return;
        }

        var metric = (sender as Button)?.Tag?.ToString() ?? "";
        var filter = ReadReportsFilter();
        var clients = FilterReportsClients(_reportsSnapshot, filter);
        var clientIds = clients.Select(item => item.ClientId).ToHashSet();
        var renewals = FilterReportsRenewals(_reportsSnapshot, filter, clientIds);
        var proposals = FilterReportsProposals(_reportsSnapshot, filter, clients);
        var quotes = _reportsSnapshot.QuoteScenarios.Where(item => item.ClientId == 0 || clientIds.Contains(item.ClientId)).ToList();

        IEnumerable<ExecutiveSummaryRow> rows = metric switch
        {
            "Total Active Clients" => clients.Select(item => Summary("Client", item.ClientName, FirstNonBlank(item.Carrier, "N/A"), item.Status)),
            "Total Premium Managed" => renewals.OrderByDescending(item => item.RenewalPremium ?? item.CurrentPremium ?? 0m).Select(item => Summary("Premium", item.ClientName, FormatMoney(item.RenewalPremium ?? item.CurrentPremium), item.Carrier)),
            "Upcoming Renewals" or "Upcoming Renewals (90 Days)" => renewals.Where(item => IsDueWithin(item.RenewalDate, 90)).OrderBy(item => item.RenewalDate).Select(item => Summary("Renewal", item.ClientName, item.RenewalDate?.ToString("MMM d yyyy", CultureInfo.InvariantCulture) ?? "N/A", item.Status)),
            "Average Renewal Increase" => renewals.OrderByDescending(item => item.PremiumChangePercent ?? decimal.MinValue).Select(item => Summary("Renewal Increase", item.ClientName, FormatPercent(item.PremiumChangePercent), item.Carrier)),
            "Average EHC Loss Ratio" or "Book Loss Ratio (EHC)" => renewals.OrderByDescending(item => item.EhcLossRatio ?? decimal.MinValue).Select(item => Summary("EHC Loss Ratio", item.ClientName, FormatPercent(item.EhcLossRatio, includeSign: false), item.Carrier)),
            "Average Dental Loss Ratio" or "Dental Loss Ratio" => renewals.OrderByDescending(item => item.DentalLossRatio ?? decimal.MinValue).Select(item => Summary("Dental Loss Ratio", item.ClientName, FormatPercent(item.DentalLossRatio, includeSign: false), item.Carrier)),
            "Total Opportunities Identified" => renewals.Where(IsExecutiveRiskRenewal).Select(item => Summary("Opportunity", item.ClientName, FormatPercent(item.PremiumChangePercent), item.Status)),
            "Negotiated Savings (YTD)" => _reportsSnapshot.RenewalVersions.Where(item => clientIds.Contains(item.ClientId) && (item.VersionName.Contains("Negotiated", StringComparison.OrdinalIgnoreCase) || item.VersionName.Contains("Final", StringComparison.OrdinalIgnoreCase))).Select(item => Summary("Negotiation", item.VersionName, item.DateCreated.ToString("MMM d yyyy", CultureInfo.InvariantCulture), item.Status)),
            _ => BuildDefaultDrillRows(renewals, proposals, quotes)
        };

        ReplaceCollection(_reportsDrillDownRows, rows.Take(200));
        SetStatus($"Drill-down updated for {FirstNonBlank(metric, "selected KPI")}.");
    }

    private void ExecutiveReportExport_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var reportName = ((sender as Button)?.Content?.ToString() ?? "Executive Report").Trim();
            var exportFolder = Path.Combine(AppPaths.DataFolder, "Executive Reports");
            Directory.CreateDirectory(exportFolder);
            var fileName = string.Concat(reportName.Select(ch => char.IsLetterOrDigit(ch) ? ch : '_')).Trim('_');
            var path = Path.Combine(exportFolder, $"{fileName}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
            IEnumerable<ExecutiveSummaryRow> rows = _reportsDrillDownRows.Count > 0
                ? _reportsDrillDownRows
                : _reportsAlertRows.Concat(_reportsRenewalAnalyticsRows).Concat(_reportsCarrierAnalyticsRows).ToList();

            SaveExecutiveReportWorkbook(path, reportName, rows);

            SetStatus($"{reportName} workbook saved to {path}.");
            System.Windows.MessageBox.Show(
                this,
                $"{reportName} workbook was saved from the read-only reporting snapshot.{Environment.NewLine}{Environment.NewLine}{path}",
                "Executive Reporting Centre",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            SetStatus("Executive report export could not be saved.");
            System.Windows.MessageBox.Show(this, ex.Message, "Executive Reporting Centre", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private void SaveExecutiveReportWorkbook(string path, string reportName, IEnumerable<ExecutiveSummaryRow> rows)
    {
        using var workbook = new XLWorkbook();
        var summarySheet = workbook.Worksheets.Add("Executive Summary");
        summarySheet.Cell(1, 1).Value = reportName;
        summarySheet.Cell(2, 1).Value = $"Generated {DateTime.Now:yyyy-MM-dd HH:mm}";
        summarySheet.Cell(3, 1).Value = "Source";
        summarySheet.Cell(3, 2).Value = "Benefits_Intelligence.db read-only reporting snapshot";
        summarySheet.Range(1, 1, 1, 6).Merge().Style.Font.SetBold().Font.SetFontSize(16).Font.SetFontColor(XLColor.FromHtml("#071B4D"));

        WriteSummaryRows(summarySheet, 5, "KPI Cards", _reportsKpiCards.Select(item => new ExecutiveSummaryRow
        {
            Category = "KPI",
            Label = item.Title,
            Value = item.Value,
            Status = item.Subtitle
        }));

        var drillSheet = workbook.Worksheets.Add("Current View");
        WriteSummaryRows(drillSheet, 1, "Current filtered rows", rows);

        var actionSheet = workbook.Worksheets.Add("Actions");
        WriteSummaryRows(actionSheet, 1, "Actionable alerts", _reportsAlertRows);

        var categorySheet = workbook.Worksheets.Add("Data Categories");
        categorySheet.Cell(1, 1).Value = "SQLite Category";
        categorySheet.Cell(1, 2).Value = "Records";
        categorySheet.Cell(1, 3).Value = "Status";
        var rowNumber = 2;
        foreach (var category in _reportsDataCategoryRows)
        {
            categorySheet.Cell(rowNumber, 1).Value = category.Category;
            categorySheet.Cell(rowNumber, 2).Value = category.Records;
            categorySheet.Cell(rowNumber, 3).Value = category.Status;
            rowNumber++;
        }

        foreach (var worksheet in workbook.Worksheets)
        {
            worksheet.Columns().AdjustToContents();
            worksheet.SheetView.FreezeRows(1);
        }

        workbook.SaveAs(path);
    }

    private static void WriteSummaryRows(IXLWorksheet worksheet, int startRow, string title, IEnumerable<ExecutiveSummaryRow> rows)
    {
        worksheet.Cell(startRow, 1).Value = title;
        worksheet.Range(startRow, 1, startRow, 6).Merge().Style.Font.SetBold().Font.SetFontColor(XLColor.FromHtml("#003A8C"));
        var headerRow = startRow + 1;
        worksheet.Cell(headerRow, 1).Value = "Category";
        worksheet.Cell(headerRow, 2).Value = "Label";
        worksheet.Cell(headerRow, 3).Value = "Value";
        worksheet.Cell(headerRow, 4).Value = "Status";
        worksheet.Cell(headerRow, 5).Value = "Owner";
        worksheet.Cell(headerRow, 6).Value = "Count";
        worksheet.Range(headerRow, 1, headerRow, 6).Style.Font.SetBold().Fill.SetBackgroundColor(XLColor.FromHtml("#DDE5F0"));

        var rowNumber = headerRow + 1;
        foreach (var row in rows)
        {
            worksheet.Cell(rowNumber, 1).Value = row.Category;
            worksheet.Cell(rowNumber, 2).Value = row.Label;
            worksheet.Cell(rowNumber, 3).Value = row.Value;
            worksheet.Cell(rowNumber, 4).Value = row.Status;
            worksheet.Cell(rowNumber, 5).Value = row.Owner;
            worksheet.Cell(rowNumber, 6).Value = row.Count == 0 ? "" : row.Count.ToString(CultureInfo.InvariantCulture);
            rowNumber++;
        }
    }

    private static ExecutiveSummaryRow Summary(string category, string label, string value, string status)
    {
        return new ExecutiveSummaryRow
        {
            Category = category,
            Label = label,
            Value = value,
            Status = status
        };
    }

    private static ExecutiveSummaryRow ActionRow(string label, int count, string owner, string status)
    {
        return new ExecutiveSummaryRow
        {
            Category = "Action",
            Label = label,
            Value = count.ToString("#,0", CultureInfo.InvariantCulture),
            Count = count,
            Owner = owner,
            Status = status
        };
    }

    private static string MapRenewalStage(RenewalQueueRecord renewal)
    {
        var status = renewal.Status ?? "";
        var selectedVersion = renewal.SelectedRenewalVersion ?? "";
        var hasDocument = renewal.SourceDocumentId.HasValue || !string.IsNullOrWhiteSpace(renewal.SourceFilePath);
        var hasExtractedPremiums = renewal.CurrentPremium.HasValue && renewal.RenewalPremium.HasValue;
        var hasExperience = renewal.EhcLossRatio.HasValue || renewal.DentalLossRatio.HasValue;

        if (status.Equals(RenewalQueueStatuses.Completed, StringComparison.OrdinalIgnoreCase) ||
            status.Equals(RenewalQueueStatuses.FinalApproved, StringComparison.OrdinalIgnoreCase) ||
            selectedVersion.Equals(RenewalVersionNames.FinalApproved, StringComparison.OrdinalIgnoreCase))
        {
            return "Completed";
        }

        if (selectedVersion.Equals(RenewalVersionNames.Negotiated, StringComparison.OrdinalIgnoreCase) ||
            status.Equals(RenewalQueueStatuses.Negotiated, StringComparison.OrdinalIgnoreCase))
        {
            return "Consultant strategy review";
        }

        if (!hasDocument)
        {
            return "Not started";
        }

        if (!hasExtractedPremiums && !hasExperience)
        {
            return "Documents received";
        }

        if (status.Equals(RenewalQueueStatuses.MissingData, StringComparison.OrdinalIgnoreCase) ||
            !hasExtractedPremiums ||
            !hasExperience)
        {
            return "Validation required";
        }

        if (status.Equals(RenewalQueueStatuses.HighRisk, StringComparison.OrdinalIgnoreCase) ||
            status.Equals(RenewalQueueStatuses.Watch, StringComparison.OrdinalIgnoreCase) ||
            renewal.PremiumChangePercent >= 10m ||
            renewal.EhcLossRatio >= 90m ||
            renewal.DentalLossRatio >= 90m)
        {
            return "Underwriting review";
        }

        if (status.Equals(RenewalQueueStatuses.ReadyToReview, StringComparison.OrdinalIgnoreCase) ||
            status.Equals(RenewalQueueStatuses.Healthy, StringComparison.OrdinalIgnoreCase))
        {
            return "Presentation ready";
        }

        return "Extracted";
    }

    private static string MapProposalStage(string status)
    {
        if (string.IsNullOrWhiteSpace(status))
        {
            return "Intake received";
        }

        if (status.Contains("pending", StringComparison.OrdinalIgnoreCase))
        {
            return "Carrier quotes pending";
        }

        if (status.Contains("comparison", StringComparison.OrdinalIgnoreCase))
        {
            return "Comparison generated";
        }

        if (status.Contains("underwriting", StringComparison.OrdinalIgnoreCase))
        {
            return "Underwriting review";
        }

        if (status.Contains("consultant", StringComparison.OrdinalIgnoreCase))
        {
            return "Consultant review";
        }

        if (status.Contains("presentation", StringComparison.OrdinalIgnoreCase))
        {
            return "Presentation ready";
        }

        if (status.Contains("finalist", StringComparison.OrdinalIgnoreCase))
        {
            return "Finalist selected";
        }

        if (IsClosedStatus(status))
        {
            return "Sold / Lost";
        }

        return "Quotes received";
    }

    private static string MapQuoteStage(string status)
    {
        if (string.IsNullOrWhiteSpace(status))
        {
            return "Request received";
        }

        if (status.Contains("carrier", StringComparison.OrdinalIgnoreCase) || status.Contains("pending", StringComparison.OrdinalIgnoreCase))
        {
            return "Carrier quote pending";
        }

        if (status.Contains("factor", StringComparison.OrdinalIgnoreCase))
        {
            return "Factor received";
        }

        if (status.Contains("cost", StringComparison.OrdinalIgnoreCase))
        {
            return "Cost impact calculated";
        }

        if (status.Contains("underwriting", StringComparison.OrdinalIgnoreCase))
        {
            return "Underwriting review";
        }

        if (status.Contains("consultant", StringComparison.OrdinalIgnoreCase))
        {
            return "Consultant review";
        }

        if (status.Contains("decision", StringComparison.OrdinalIgnoreCase))
        {
            return "Client decision pending";
        }

        return IsClosedStatus(status) ? "Completed" : "Request received";
    }

    private static bool IsClosedStatus(string status)
    {
        return status.Contains("sold", StringComparison.OrdinalIgnoreCase) ||
               status.Contains("lost", StringComparison.OrdinalIgnoreCase) ||
               status.Contains("complete", StringComparison.OrdinalIgnoreCase) ||
               status.Contains("cancel", StringComparison.OrdinalIgnoreCase) ||
               status.Contains("closed", StringComparison.OrdinalIgnoreCase) ||
               status.Contains("final", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsDueWithin(DateOnly? date, int days)
    {
        if (!date.HasValue)
        {
            return false;
        }

        var today = DateOnly.FromDateTime(DateTime.Today);
        return date.Value >= today && date.Value <= today.AddDays(days);
    }

    private static int CountDueInDays(IEnumerable<RenewalQueueRecord> renewals, int days)
    {
        return renewals.Count(item => IsDueWithin(item.RenewalDate, days));
    }

    private static bool IsBenefit(string value, params string[] needles)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }

        return needles.Any(needle => value.Contains(needle, StringComparison.OrdinalIgnoreCase));
    }

    private static decimal? AverageRaw(IEnumerable<decimal?> values)
    {
        var materialized = values.Where(item => item.HasValue).Select(item => item!.Value).ToList();
        return materialized.Count == 0 ? null : materialized.Average();
    }

    private static decimal? AverageRaw(IEnumerable<decimal> values)
    {
        var materialized = values.ToList();
        return materialized.Count == 0 ? null : materialized.Average();
    }

    private static decimal? AverageRaw(IEnumerable<int?> values)
    {
        var materialized = values.Where(item => item.HasValue).Select(item => (decimal)item!.Value).ToList();
        return materialized.Count == 0 ? null : materialized.Average();
    }

    private static decimal? Median(IReadOnlyCollection<decimal> values)
    {
        if (values.Count == 0)
        {
            return null;
        }

        var sorted = values.OrderBy(value => value).ToList();
        var mid = sorted.Count / 2;
        return sorted.Count % 2 == 0 ? (sorted[mid - 1] + sorted[mid]) / 2m : sorted[mid];
    }

    private static decimal? CalculateNegotiatedSavings(IEnumerable<RenewalVersionRecord> versions)
    {
        return versions
            .Where(version => version.VersionName.Contains("Negotiated", StringComparison.OrdinalIgnoreCase) ||
                              version.VersionName.Contains("Final", StringComparison.OrdinalIgnoreCase))
            .SelectMany(version => version.Rates)
            .SumValue(rate => rate.ProposedPremium.HasValue && rate.VersionPremium.HasValue
                ? rate.ProposedPremium - rate.VersionPremium
                : rate.SavingsFallback());
    }

    private static decimal? CalculateClaimsTrend(IReadOnlyCollection<MonthlyClaimsRecord> monthlyRows)
    {
        var rows = monthlyRows.Where(item => item.Month.HasValue).OrderByDescending(item => item.Month).ToList();
        if (rows.Count < 2)
        {
            return null;
        }

        var recentStart = rows.Max(item => item.Month!.Value).AddMonths(-11);
        var recent = rows.Where(item => item.Month!.Value >= recentStart).SumValue(item => item.IncurredClaims ?? item.PaidClaims);
        var priorStart = recentStart.AddMonths(-12);
        var prior = rows.Where(item => item.Month!.Value >= priorStart && item.Month!.Value < recentStart).SumValue(item => item.IncurredClaims ?? item.PaidClaims);
        return prior is > 0m && recent.HasValue ? (recent.Value - prior.Value) / prior.Value * 100m : null;
    }

    private static int CountAboveAverage(IReadOnlyCollection<RenewalQueueRecord> renewals)
    {
        var average = AverageRaw(renewals.Select(row => row.PremiumChangePercent));
        return average.HasValue ? renewals.Count(item => item.PremiumChangePercent.HasValue && item.PremiumChangePercent > average) : 0;
    }

    private static int CountBelowAverage(IReadOnlyCollection<RenewalQueueRecord> renewals)
    {
        var average = AverageRaw(renewals.Select(row => row.PremiumChangePercent));
        return average.HasValue ? renewals.Count(item => item.PremiumChangePercent.HasValue && item.PremiumChangePercent < average) : 0;
    }

    private static decimal? CalculateClaimsVolatility(IReadOnlyCollection<MonthlyClaimsRecord> monthlyRows)
    {
        var values = monthlyRows
            .Select(item => item.IncurredClaims ?? item.PaidClaims)
            .Where(item => item.HasValue)
            .Select(item => item!.Value)
            .ToList();
        if (values.Count < 2)
        {
            return null;
        }

        var average = values.Average();
        if (average == 0m)
        {
            return null;
        }

        var variance = values.Select(value => Math.Pow((double)(value - average), 2)).Average();
        var standardDeviation = (decimal)Math.Sqrt(variance);
        return standardDeviation / average * 100m;
    }

    private static decimal? AverageNormalizedRate(IEnumerable<decimal?> values)
    {
        var materialized = values.Select(NormalizeRatePercent).Where(item => item.HasValue).Select(item => item!.Value).ToList();
        return materialized.Count == 0 ? null : materialized.Average();
    }

    private static decimal? NormalizeRatePercent(decimal? value)
    {
        if (!value.HasValue)
        {
            return null;
        }

        return value.Value is > -1m and < 1m ? value.Value * 100m : value.Value;
    }

    private static string FormatNormalizedPercent(decimal? value, bool includeSign = true)
    {
        return FormatPercent(NormalizeRatePercent(value), includeSign);
    }

    private static string CommonPlanValue(IEnumerable<PlanDesignRecord> planDesign, params string[] terms)
    {
        return CommonPlanValue(planDesign, terms, preferReimbursement: false);
    }

    private static string CommonPlanValue(IEnumerable<PlanDesignRecord> planDesign, string term, bool preferReimbursement)
    {
        return CommonPlanValue(planDesign, [term], preferReimbursement);
    }

    private static string CommonPlanValue(IEnumerable<PlanDesignRecord> planDesign, string term1, string term2)
    {
        return CommonPlanValue(planDesign, [term1, term2], preferReimbursement: false);
    }

    private static string CommonPlanValue(IEnumerable<PlanDesignRecord> planDesign, IReadOnlyCollection<string> terms, bool preferReimbursement)
    {
        var rows = planDesign
            .Where(item => terms.Any(term =>
                item.BenefitCategory.Contains(term, StringComparison.OrdinalIgnoreCase) ||
                item.BenefitName.Contains(term, StringComparison.OrdinalIgnoreCase) ||
                item.CoverageDescription.Contains(term, StringComparison.OrdinalIgnoreCase)))
            .ToList();

        if (rows.Count == 0)
        {
            return "No saved data";
        }

        if (preferReimbursement)
        {
            var reimbursement = rows
                .Where(item => item.ReimbursementPercent.HasValue)
                .GroupBy(item => item.ReimbursementPercent!.Value)
                .OrderByDescending(group => group.Count())
                .FirstOrDefault();
            if (reimbursement is not null)
            {
                return FormatPercent(reimbursement.Key, includeSign: false);
            }
        }

        return rows
            .Select(item => FirstNonBlank(item.MaximumAmount, item.CoverageDescription, item.Frequency, "Saved"))
            .GroupBy(value => value, StringComparer.OrdinalIgnoreCase)
            .OrderByDescending(group => group.Count())
            .First()
            .Key;
    }

    private static int CalculateProposalSavingsCount(IEnumerable<ExecutiveProposalRecord> proposals)
    {
        return proposals.Count(item => item.Savings.HasValue && item.Savings.Value > 0m);
    }

    private static decimal ScoreLossRatio(decimal? lossRatio)
    {
        if (!lossRatio.HasValue)
        {
            return 75m;
        }

        return lossRatio.Value switch
        {
            <= 75m => 100m,
            <= 90m => 82m,
            <= 100m => 65m,
            <= 115m => 45m,
            _ => 25m
        };
    }

    private static decimal ScoreRenewalIncrease(decimal? renewalIncrease)
    {
        if (!renewalIncrease.HasValue)
        {
            return 75m;
        }

        return renewalIncrease.Value switch
        {
            <= 0m => 100m,
            <= 5m => 88m,
            <= 10m => 72m,
            <= 15m => 52m,
            _ => 25m
        };
    }

    private static decimal ScoreVolatility(decimal? volatility)
    {
        if (!volatility.HasValue)
        {
            return 75m;
        }

        return volatility.Value switch
        {
            <= 15m => 100m,
            <= 30m => 82m,
            <= 50m => 65m,
            <= 75m => 45m,
            _ => 25m
        };
    }

    private static decimal ScorePoolingRisk(decimal? poolingRiskSignal, decimal? lossRatio)
    {
        var signal = NormalizeRatePercent(poolingRiskSignal);
        if (!signal.HasValue && !lossRatio.HasValue)
        {
            return 75m;
        }

        var risk = Math.Max(signal ?? 0m, lossRatio.HasValue && lossRatio > 110m ? 20m : 0m);
        return risk switch
        {
            <= 2m => 92m,
            <= 5m => 78m,
            <= 10m => 62m,
            <= 20m => 42m,
            _ => 25m
        };
    }

    private async Task ShowUnderwritingDashboardAsync(RenewalQueueItem? requestedRow = null)
    {
        SetNavigationWorkspace(
            _selectedClientWorkspace is null ? "BookOfBusiness" : "ClientWorkspace",
            "Renewal",
            _selectedClientWorkspace?.ClientName);
        HomeDashboardScrollViewer.Visibility = Visibility.Collapsed;
        DashboardScrollViewer.Visibility = Visibility.Collapsed;
        AssumptionLibraryScrollViewer.Visibility = Visibility.Collapsed;
        ReportsAnalyticsScrollViewer.Visibility = Visibility.Collapsed;
        BenchmarkingDashboardScrollViewer.Visibility = Visibility.Collapsed;
        ProposalComparisonScrollViewer.Visibility = Visibility.Collapsed;
        PlanDesignIntelligenceScrollViewer.Visibility = Visibility.Collapsed;
        PlanDesignFactorsScrollViewer.Visibility = Visibility.Collapsed;
        DocumentIntelligenceScrollViewer.Visibility = Visibility.Collapsed;
        ClientsScrollViewer.Visibility = Visibility.Collapsed;
        NegotiationWorkbenchScrollViewer.Visibility = Visibility.Collapsed;
        UnderwritingWorkbenchScrollViewer.Visibility = Visibility.Visible;
        UnderwritingWorkbenchScrollViewer.ScrollToTop();
        SetOperationalFooterVisible(true);

        var row = requestedRow
                  ?? RenewalQueueGrid.SelectedItem as RenewalQueueItem
                  ?? _renewalQueueView?.Cast<RenewalQueueItem>().FirstOrDefault()
                  ?? _renewalQueue.FirstOrDefault();

        _isUpdatingUnderwritingSelection = true;
        UnderwritingClientComboBox.SelectedItem = row;
        _isUpdatingUnderwritingSelection = false;

        await LoadUnderwritingDashboardAsync(row);
    }

    private async Task LoadUnderwritingDashboardAsync(RenewalQueueItem? row)
    {
        _selectedUnderwritingRenewal = row;
        _underwritingBenefitRows.Clear();

        if (row is null)
        {
            UnderwritingClientNameTextBlock.Text = "No Renewal Selected";
            UnderwritingClientMetaTextBlock.Text = "Load or save a renewal to start underwriting.";
            UpdateUnderwritingSummary(null, [], "No data loaded.", ["- No renewal data is available."], []);
            return;
        }

        UnderwritingClientNameTextBlock.Text = row.Client;
        UnderwritingClientMetaTextBlock.Text = $"{row.Carrier} | {row.RenewalDateText} renewal | {row.LivesText} lives";

        try
        {
            SetBusy(true, "Loading underwriting intelligence...");
            var result = IsCurrentAnalysisContext(row) && _analysis is not null
                ? BuildUnderwritingRowsFromCurrentAnalysis(_analysis)
                : await BuildUnderwritingRowsFromDatabaseAsync(row);

            var rows = result.Rows.Count > 0
                ? result.Rows
                : BuildFallbackUnderwritingRows(row);

            foreach (var benefitRow in BenefitDisplayOrder.Sort(rows, item => item.BenefitLine))
            {
                _underwritingBenefitRows.Add(benefitRow);
            }

            var warnings = result.Warnings;
            if (result.Rows.Count == 0)
            {
                warnings.Add("- Detailed version data was not found, so the module is using a conservative summary estimate.");
            }

            UpdateUnderwritingSummary(row, rows, result.Source, warnings, result.Audit);
            SetStatus($"Opened Underwriting Intelligence for {row.Client}.");
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            UpdateUnderwritingSummary(row, [], "Underwriting load failed.", [$"- {ex.Message}"], []);
            SetStatus("Underwriting Intelligence could not load the selected renewal.");
        }
        finally
        {
            SetBusy(false);
        }
    }

    private (List<UnderwritingBenefitRow> Rows, string Source, List<string> Warnings, List<string> Audit) BuildUnderwritingRowsFromCurrentAnalysis(AnalysisResult analysis)
    {
        if (analysis.RenewalVersions.Count == 0)
        {
            new RenewalEngine().BuildRenewalVersions(analysis);
        }

        var carrier = analysis.RenewalVersions.FirstOrDefault(version => version.VersionName.Equals(RenewalVersionNames.CarrierProposed, StringComparison.OrdinalIgnoreCase));
        var model = analysis.RenewalVersions.FirstOrDefault(version => version.VersionName.Equals(RenewalVersionNames.ConsultantModel, StringComparison.OrdinalIgnoreCase));
        var rows = BuildUnderwritingRows(carrier, model);
        return (
            rows,
            "Current in-memory renewal analysis",
            BuildUnderwritingWarnings(rows),
            [
                "- Source: current renewal files loaded in Workflow Setup.",
                "- Calculation layer: shared UnderwritingEngine.",
                "- Version stack: Carrier Proposed compared with Consultant Model."
            ]);
    }

    private async Task<(List<UnderwritingBenefitRow> Rows, string Source, List<string> Warnings, List<string> Audit)> BuildUnderwritingRowsFromDatabaseAsync(RenewalQueueItem row)
    {
        if (row.ClientId <= 0)
        {
            return ([], "Sample queue estimate", ["- This sample row is not connected to saved SQL history."], []);
        }

        var history = await _databaseService.GetClientHistoryAsync(row.ClientId);
        var versions = history.RenewalVersions
            .Where(version => !row.SourceDocumentId.HasValue || version.SourceDocumentId == row.SourceDocumentId)
            .ToList();
        if (versions.Count == 0)
        {
            versions = history.RenewalVersions;
        }

        var carrier = LatestVersion(versions, RenewalVersionNames.CarrierProposed);
        var model = LatestVersion(versions, RenewalVersionNames.ConsultantModel);
        var rows = BuildUnderwritingRows(row, history.Renewals, carrier, model);
        return (
            rows,
            "Saved SQL renewal versions",
            BuildUnderwritingWarnings(rows),
            [
                "- Source: Benefits_Intelligence.db renewal history.",
                "- Calculation layer: shared UnderwritingEngine.",
                "- Version stack: Carrier Proposed compared with Consultant Model."
            ]);
    }

    private static RenewalVersionRecord? LatestVersion(IEnumerable<RenewalVersionRecord> versions, string versionName)
    {
        return versions
            .Where(version => version.VersionName.Equals(versionName, StringComparison.OrdinalIgnoreCase))
            .OrderByDescending(version => version.DateCreated)
            .ThenByDescending(version => version.VersionId)
            .FirstOrDefault();
    }

    private static List<UnderwritingBenefitRow> BuildUnderwritingRows(RenewalVersion? carrier, RenewalVersion? model)
    {
        var rateLines = (carrier?.RateLines.Count > 0 ? carrier.RateLines : model?.RateLines) ?? [];
        var rows = new List<UnderwritingBenefitRow>();
        foreach (var group in rateLines.GroupBy(line => UnderwritingBenefitGroupName(line.Benefit), StringComparer.OrdinalIgnoreCase))
        {
            var benefit = group.Key;
            var currentPremium = group.SumValue(line => line.CurrentPremium);
            var proposedPremium = group.SumValue(line => line.VersionPremium ?? line.ProposedPremium);
            var modelRates = model?.RateLines
                .Where(line => UnderwritingBenefitGroupName(line.Benefit).Equals(benefit, StringComparison.OrdinalIgnoreCase))
                .ToList() ?? [];
            var modelCalculations = model?.Calculations
                .Where(calculation => UnderwritingBenefitGroupName(calculation.BenefitLine).Equals(benefit, StringComparison.OrdinalIgnoreCase))
                .ToList() ?? [];
            var currentClaims = modelCalculations.SumValue(calculation => calculation.CurrentClaims);
            var projectedClaims = modelCalculations.SumValue(calculation => calculation.ProjectedClaims);
            var modelledPremium = modelRates.SumValue(line => line.VersionPremium)
                                  ?? modelCalculations.SumValue(calculation => calculation.RequiredPremium);
            var trend = FirstValue(modelCalculations.Select(calculation => calculation.Trend));
            var targetLossRatio = FirstValue(modelCalculations.Select(calculation => calculation.TargetLossRatio));

            rows.Add(CreateUnderwritingBenefitRow(
                benefit,
                currentPremium,
                currentClaims,
                proposedPremium,
                modelledPremium,
                projectedClaims,
                trend,
                targetLossRatio,
                null));
        }

        return rows;
    }

    private static List<UnderwritingBenefitRow> BuildUnderwritingRows(
        RenewalQueueItem queueRow,
        IReadOnlyCollection<RenewalRecord> renewalHistory,
        RenewalVersionRecord? carrier,
        RenewalVersionRecord? model)
    {
        var rateLines = (carrier?.Rates.Count > 0 ? carrier.Rates : model?.Rates) ?? [];
        var rows = new List<UnderwritingBenefitRow>();
        foreach (var group in rateLines.GroupBy(line => UnderwritingBenefitGroupName(line.BenefitLine), StringComparer.OrdinalIgnoreCase))
        {
            var benefit = group.Key;
            var modelRates = model?.Rates
                .Where(line => UnderwritingBenefitGroupName(line.BenefitLine).Equals(benefit, StringComparison.OrdinalIgnoreCase))
                .ToList() ?? [];
            var modelCalculations = model?.Calculations
                .Where(calculation => UnderwritingBenefitGroupName(calculation.BenefitLine).Equals(benefit, StringComparison.OrdinalIgnoreCase))
                .ToList() ?? [];
            var currentPremium = group.SumValue(line => line.CurrentPremium);
            var proposedPremium = group.SumValue(line => line.VersionPremium ?? line.ProposedPremium);
            var modelledPremium = modelRates.SumValue(line => line.VersionPremium)
                                  ?? modelCalculations.SumValue(calculation => calculation.RequiredPremium);
            var lossRatio = ResolveUnderwritingLossRatio(queueRow, renewalHistory, benefit);
            var currentClaims = currentPremium.HasValue && lossRatio.HasValue
                ? decimal.Round(currentPremium.Value * UnderwritingEngine.NormalizeRate(lossRatio).GetValueOrDefault(), 2)
                : (decimal?)null;
            var projectedClaims = modelCalculations.SumValue(calculation => calculation.ProjectedClaims);
            var trend = benefit.Equals("Dental", StringComparison.OrdinalIgnoreCase) ? 0.06m :
                benefit.Contains("Health", StringComparison.OrdinalIgnoreCase) || benefit.Equals("EHC", StringComparison.OrdinalIgnoreCase) ? 0.08m : (decimal?)null;
            var targetLossRatio = benefit.Equals("Dental", StringComparison.OrdinalIgnoreCase) ||
                                  benefit.Contains("Health", StringComparison.OrdinalIgnoreCase)
                ? 0.763m
                : 0.85m;

            rows.Add(CreateUnderwritingBenefitRow(
                benefit,
                currentPremium,
                currentClaims,
                proposedPremium,
                modelledPremium,
                projectedClaims,
                trend,
                targetLossRatio,
                lossRatio));
        }

        return rows;
    }

    private static List<UnderwritingBenefitRow> BuildFallbackUnderwritingRows(RenewalQueueItem row)
    {
        var rows = new List<UnderwritingBenefitRow>();
        AddFallbackUnderwritingRow(rows, "Extended Health Care", row.CurrentPremium, row.RenewalPremium, row.EhcLossRatio, 0.48m, 0.08m, 0.763m);
        AddFallbackUnderwritingRow(rows, "Dental", row.CurrentPremium, row.RenewalPremium, row.DentalLossRatio, 0.24m, 0.06m, 0.763m);
        AddFallbackUnderwritingRow(rows, "Life / AD&D", row.CurrentPremium, row.RenewalPremium, null, 0.14m, null, 0.85m);
        AddFallbackUnderwritingRow(rows, "Disability", row.CurrentPremium, row.RenewalPremium, null, 0.14m, null, 0.85m);
        return rows;
    }

    private static void AddFallbackUnderwritingRow(
        List<UnderwritingBenefitRow> rows,
        string benefit,
        decimal? totalCurrentPremium,
        decimal? totalProposedPremium,
        decimal? lossRatio,
        decimal share,
        decimal? trend,
        decimal targetLossRatio)
    {
        var currentPremium = totalCurrentPremium.HasValue ? decimal.Round(totalCurrentPremium.Value * share, 2) : (decimal?)null;
        var proposedPremium = totalProposedPremium.HasValue ? decimal.Round(totalProposedPremium.Value * share, 2) : (decimal?)null;
        var claims = currentPremium.HasValue && lossRatio.HasValue
            ? decimal.Round(currentPremium.Value * UnderwritingEngine.NormalizeRate(lossRatio).GetValueOrDefault(), 2)
            : (decimal?)null;
        var calculation = new UnderwritingEngine().Calculate(new UnderwritingCalculationInput
        {
            BenefitLine = benefit,
            CurrentClaims = claims,
            CurrentPremium = currentPremium,
            ProposedPremium = proposedPremium,
            Trend = trend,
            TargetLossRatio = targetLossRatio
        });

        rows.Add(CreateUnderwritingBenefitRow(
            benefit,
            currentPremium,
            claims,
            proposedPremium,
            calculation.RequiredPremium ?? proposedPremium,
            calculation.ProjectedClaims,
            trend,
            targetLossRatio,
            lossRatio));
    }

    private static UnderwritingBenefitRow CreateUnderwritingBenefitRow(
        string benefit,
        decimal? currentPremium,
        decimal? currentClaims,
        decimal? proposedPremium,
        decimal? modelledPremium,
        decimal? projectedClaims,
        decimal? trend,
        decimal? targetLossRatio,
        decimal? sourceLossRatio)
    {
        var lossRatio = sourceLossRatio ?? UnderwritingEngine.CalculateLossRatioPercent(currentClaims, currentPremium);
        var proposedChange = UnderwritingEngine.CalculatePercentChange(currentPremium, proposedPremium);
        var requiredChange = UnderwritingEngine.CalculatePercentChange(currentPremium, modelledPremium);
        var gap = proposedPremium.HasValue && modelledPremium.HasValue
            ? decimal.Round(proposedPremium.Value - modelledPremium.Value, 2)
            : (decimal?)null;

        return new UnderwritingBenefitRow
        {
            BenefitLine = benefit,
            CurrentPremium = currentPremium,
            CurrentClaims = currentClaims,
            LossRatio = lossRatio,
            Trend = trend,
            TargetLossRatio = targetLossRatio,
            ProjectedClaims = projectedClaims,
            ModelledPremium = modelledPremium,
            ProposedPremium = proposedPremium,
            ProposedChangePercent = proposedChange,
            RequiredChangePercent = requiredChange,
            GapToModel = gap,
            RiskStatus = DetermineUnderwritingBenefitRisk(lossRatio, proposedChange, gap, modelledPremium)
        };
    }

    private static decimal? ResolveUnderwritingLossRatio(RenewalQueueItem queueRow, IEnumerable<RenewalRecord> renewalHistory, string benefit)
    {
        if (benefit.Equals("Dental", StringComparison.OrdinalIgnoreCase))
        {
            return queueRow.DentalLossRatio;
        }

        if (benefit.Contains("Health", StringComparison.OrdinalIgnoreCase) || benefit.Equals("EHC", StringComparison.OrdinalIgnoreCase))
        {
            return queueRow.EhcLossRatio;
        }

        var matches = renewalHistory
            .Where(row => UnderwritingBenefitGroupName(row.BenefitLine).Equals(benefit, StringComparison.OrdinalIgnoreCase))
            .Where(row => row.LossRatio.HasValue)
            .ToList();
        return matches.Count == 0 ? null : matches.Max(row => row.LossRatio);
    }

    private static string UnderwritingBenefitGroupName(string benefit)
    {
        if (BenefitMatcher.ContainsPhrase(benefit, "dental"))
        {
            return "Dental";
        }

        if (BenefitMatcher.ContainsPhrase(benefit, "extended health care") ||
            BenefitMatcher.ContainsPhrase(benefit, "health") ||
            BenefitMatcher.ContainsPhrase(benefit, "ehc"))
        {
            return "Extended Health Care";
        }

        if (BenefitMatcher.ContainsPhrase(benefit, "accidental death") ||
            BenefitMatcher.ContainsPhrase(benefit, "ad d") ||
            BenefitMatcher.ContainsPhrase(benefit, "ad&d"))
        {
            return "AD&D";
        }

        return string.IsNullOrWhiteSpace(benefit) ? "Overall" : benefit.Trim();
    }

    private static string DetermineUnderwritingBenefitRisk(decimal? lossRatio, decimal? proposedChange, decimal? gap, decimal? modelledPremium)
    {
        if (!lossRatio.HasValue && !proposedChange.HasValue)
        {
            return RenewalQueueStatuses.MissingData;
        }

        if (lossRatio > 90m || proposedChange > 10m || (gap.HasValue && modelledPremium is > 0m && gap.Value / modelledPremium.Value > 0.10m))
        {
            return RenewalQueueStatuses.HighRisk;
        }

        if (lossRatio is >= 80m || proposedChange is >= 5m || gap > 0m)
        {
            return RenewalQueueStatuses.Watch;
        }

        return RenewalQueueStatuses.Healthy;
    }

    private static decimal? FirstValue(IEnumerable<decimal?> values)
    {
        return values.FirstOrDefault(value => value.HasValue);
    }

    private static List<string> BuildUnderwritingWarnings(IReadOnlyCollection<UnderwritingBenefitRow> rows)
    {
        var warnings = new List<string>();
        if (rows.Count == 0)
        {
            warnings.Add("- No benefit-level underwriting rows were available.");
            return warnings;
        }

        if (rows.Any(row => !row.CurrentPremium.HasValue))
        {
            warnings.Add("- One or more benefit lines are missing current premium.");
        }

        if (rows.Any(row => !row.ProposedPremium.HasValue))
        {
            warnings.Add("- One or more benefit lines are missing carrier proposed premium.");
        }

        if (rows.Any(row => !row.LossRatio.HasValue))
        {
            warnings.Add("- Some benefit lines do not have loss ratio support yet.");
        }

        return warnings.Count == 0 ? ["- No required data gaps identified."] : warnings;
    }

    private void UpdateUnderwritingSummary(
        RenewalQueueItem? row,
        IReadOnlyCollection<UnderwritingBenefitRow> rows,
        string source,
        IReadOnlyCollection<string> warnings,
        IReadOnlyCollection<string> audit)
    {
        var current = rows.SumValue(item => item.CurrentPremium) ?? row?.CurrentPremium;
        var proposed = rows.SumValue(item => item.ProposedPremium) ?? row?.RenewalPremium;
        var modelled = rows.SumValue(item => item.ModelledPremium);
        var proposedIncrease = UnderwritingEngine.CalculatePercentChange(current, proposed) ?? row?.PremiumChangePercent;
        var requiredIncrease = UnderwritingEngine.CalculatePercentChange(current, modelled);
        var gap = proposed.HasValue && modelled.HasValue ? decimal.Round(proposed.Value - modelled.Value, 2) : (decimal?)null;
        var status = DetermineUnderwritingStatus(rows, proposedIncrease, gap, modelled);

        UnderwritingCurrentPremiumKpiTextBlock.Text = FormatMoney(current);
        UnderwritingCarrierProposalKpiTextBlock.Text = FormatMoney(proposed);
        UnderwritingModelledPremiumKpiTextBlock.Text = FormatMoney(modelled);
        UnderwritingCarrierIncreaseKpiTextBlock.Text = FormatPercent(proposedIncrease);
        UnderwritingRequiredIncreaseKpiTextBlock.Text = FormatPercent(requiredIncrease);
        UnderwritingGapToModelKpiTextBlock.Text = FormatMoney(gap);
        UnderwritingEhcLossKpiTextBlock.Text = FormatPercent(row?.EhcLossRatio, includeSign: false);
        UnderwritingDentalLossKpiTextBlock.Text = FormatPercent(row?.DentalLossRatio, includeSign: false);

        ApplyUnderwritingStatus(status);
        UnderwritingStrategyItemsControl.ItemsSource = BuildUnderwritingStrategy(row, rows, proposedIncrease, requiredIncrease, gap);
        UnderwritingWarningItemsControl.ItemsSource = warnings;
        UnderwritingAuditItemsControl.ItemsSource = audit.Count == 0
            ? [$"- Source: {source}", "- Calculation layer: shared UnderwritingEngine."]
            : audit;
        UnderwritingAssumptionsTextBlock.Text = BuildUnderwritingAssumptionSummary(rows, source);
    }

    private static List<string> BuildUnderwritingStrategy(
        RenewalQueueItem? row,
        IReadOnlyCollection<UnderwritingBenefitRow> rows,
        decimal? proposedIncrease,
        decimal? requiredIncrease,
        decimal? gap)
    {
        var strategy = new List<string>();
        if (row is null || rows.Count == 0)
        {
            return ["- Load or select a renewal before preparing a strategy."];
        }

        if (gap > 0m)
        {
            strategy.Add($"- Carrier proposal is {FormatMoney(gap)} above the model. Use this as the opening negotiation ask.");
        }
        else if (gap < 0m)
        {
            strategy.Add($"- Carrier proposal is {FormatMoney(Math.Abs(gap.Value))} below the model. Validate assumptions before pushing harder.");
        }
        else
        {
            strategy.Add("- Carrier proposal is close to the modelled requirement.");
        }

        var highRiskRows = rows.Where(item => item.RiskStatus == RenewalQueueStatuses.HighRisk).Select(item => item.BenefitLine).Distinct().ToList();
        if (highRiskRows.Count > 0)
        {
            strategy.Add($"- Focus review on {string.Join(", ", highRiskRows)} before client presentation.");
        }

        if (proposedIncrease > 10m)
        {
            strategy.Add("- Carrier increase is above 10%; prepare alternate scenario and carrier challenge notes.");
        }
        else if (requiredIncrease.HasValue && proposedIncrease.HasValue && proposedIncrease.Value > requiredIncrease.Value)
        {
            strategy.Add("- Ask carrier to justify the gap between proposed and modelled required premium.");
        }

        return strategy;
    }

    private static string DetermineUnderwritingStatus(
        IReadOnlyCollection<UnderwritingBenefitRow> rows,
        decimal? proposedIncrease,
        decimal? gap,
        decimal? modelledPremium)
    {
        if (rows.Count == 0 || rows.Any(row => row.RiskStatus == RenewalQueueStatuses.MissingData))
        {
            return RenewalQueueStatuses.MissingData;
        }

        if (rows.Any(row => row.RiskStatus == RenewalQueueStatuses.HighRisk) ||
            proposedIncrease > 10m ||
            (gap.HasValue && modelledPremium is > 0m && gap.Value / modelledPremium.Value > 0.10m))
        {
            return RenewalQueueStatuses.HighRisk;
        }

        if (rows.Any(row => row.RiskStatus == RenewalQueueStatuses.Watch) || proposedIncrease >= 5m || gap > 0m)
        {
            return RenewalQueueStatuses.Watch;
        }

        return RenewalQueueStatuses.Healthy;
    }

    private static string BuildUnderwritingAssumptionSummary(IReadOnlyCollection<UnderwritingBenefitRow> rows, string source)
    {
        var ehc = rows.FirstOrDefault(row => row.BenefitLine.Contains("Health", StringComparison.OrdinalIgnoreCase));
        var dental = rows.FirstOrDefault(row => row.BenefitLine.Equals("Dental", StringComparison.OrdinalIgnoreCase));
        return
            $"Source: {source}{Environment.NewLine}" +
            $"EHC trend: {FormatPercent(ehc?.Trend is null ? null : ehc.Trend * 100m, includeSign: false)} | EHC target LR: {FormatPercent(ehc?.TargetLossRatio is null ? null : ehc.TargetLossRatio * 100m, includeSign: false)}{Environment.NewLine}" +
            $"Dental trend: {FormatPercent(dental?.Trend is null ? null : dental.Trend * 100m, includeSign: false)} | Dental target LR: {FormatPercent(dental?.TargetLossRatio is null ? null : dental.TargetLossRatio * 100m, includeSign: false)}";
    }

    private void ApplyUnderwritingStatus(string status)
    {
        var sample = new RenewalQueueItem { Status = status };
        UnderwritingStatusTextBlock.Text = status;
        UnderwritingStatusBadge.Background = sample.StatusBackground;
        UnderwritingStatusBadge.BorderBrush = sample.StatusBorder;
        UnderwritingStatusTextBlock.Foreground = sample.StatusForeground;
    }

    private void NegotiationRatesGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
    {
        if (e.EditAction != DataGridEditAction.Commit || e.Row.Item is not NegotiationRateRow row)
        {
            return;
        }

        var header = e.Column.Header?.ToString() ?? "";
        Dispatcher.BeginInvoke(() =>
        {
            if (header.Equals("Negotiated Rate", StringComparison.OrdinalIgnoreCase))
            {
                if (!row.CanEnterNegotiatedRate)
                {
                    row.NegotiatedRate = null;
                }
                else
                {
                    row.NegotiatedPremium = CalculateNegotiatedPremiumFromRate(row) ?? row.NegotiatedPremium;
                }
            }
            else if (header.Equals("Negotiated Premium", StringComparison.OrdinalIgnoreCase))
            {
                row.NegotiatedRate = row.CanEnterNegotiatedRate
                    ? CalculateNegotiatedRateFromPremium(row) ?? row.NegotiatedRate
                    : null;
            }

            RecalculateNegotiation();
        }, DispatcherPriority.Background);
    }

    private void NegotiatedRateTextBox_LostFocus(object sender, RoutedEventArgs e)
    {
        if (sender is not TextBox textBox || textBox.DataContext is not NegotiationRateRow row)
        {
            return;
        }

        textBox.GetBindingExpression(TextBox.TextProperty)?.UpdateSource();
        if (!row.CanEnterNegotiatedRate)
        {
            row.NegotiatedRate = null;
            SetStatus($"{row.BenefitLine}: this row is premium-based. Enter the negotiated premium instead of a rate.");
            RecalculateNegotiation();
            return;
        }

        row.NegotiatedPremium = CalculateNegotiatedPremiumFromRate(row) ?? row.NegotiatedPremium;
        RecalculateNegotiation();
    }

    private void BaselineRateTextBox_LostFocus(object sender, RoutedEventArgs e)
    {
        if (sender is not TextBox textBox || textBox.DataContext is not NegotiationRateRow)
        {
            return;
        }

        textBox.GetBindingExpression(TextBox.TextProperty)?.UpdateSource();
        if (Validation.GetHasError(textBox))
        {
            SetStatus("Enter a valid current or proposed rate of zero or greater.");
            return;
        }

        RecalculateNegotiation();
    }

    private void NegotiatedPremiumTextBox_LostFocus(object sender, RoutedEventArgs e)
    {
        if (sender is not TextBox textBox || textBox.DataContext is not NegotiationRateRow row)
        {
            return;
        }

        textBox.GetBindingExpression(TextBox.TextProperty)?.UpdateSource();
        row.NegotiatedRate = row.CanEnterNegotiatedRate
            ? CalculateNegotiatedRateFromPremium(row) ?? row.NegotiatedRate
            : null;
        RecalculateNegotiation();
    }

    private static decimal? CalculateNegotiatedPremiumFromRate(NegotiationRateRow row)
    {
        var premium = UnderwritingEngine.CalculatePremium(row.BenefitLine, row.Volume, row.Lives, row.NegotiatedRate, row.RateBasis);
        return premium.HasValue ? decimal.Round(premium.Value * PremiumScaleForRateCalculation(row), 2) : null;
    }

    private static decimal? CalculateNegotiatedRateFromPremium(NegotiationRateRow row)
    {
        if (!row.NegotiatedPremium.HasValue)
        {
            return null;
        }

        var premiumForRate = decimal.Round(row.NegotiatedPremium.Value / PremiumScaleForRateCalculation(row), 2);
        return UnderwritingEngine.CalculateRate(row.BenefitLine, row.Volume, row.Lives, premiumForRate, row.RateBasis);
    }

    internal static decimal PremiumScaleForRateCalculation(NegotiationRateRow row)
    {
        var scale = PremiumScaleFromComparablePremium(row.CurrentPremium, row.CurrentRate, row)
                    ?? PremiumScaleFromComparablePremium(row.ProposedPremium, row.ProposedRate, row);
        return scale ?? 1m;
    }

    private static decimal? PremiumScaleFromComparablePremium(decimal? displayedPremium, decimal? rate, NegotiationRateRow row)
    {
        var calculatedPremium = UnderwritingEngine.CalculatePremium(row.BenefitLine, row.Volume, row.Lives, rate, row.RateBasis);
        if (displayedPremium is not > 0m || calculatedPremium is not > 0m)
        {
            return null;
        }

        var ratio = displayedPremium.Value / calculatedPremium.Value;
        if (ratio is > 10m and < 14m)
        {
            return 12m;
        }

        return ratio is > 0.75m and < 1.25m ? 1m : null;
    }

    private static void CorrectSavedNegotiatedPremiumScale(NegotiationRateRow row)
    {
        if (!row.CanEnterNegotiatedRate || !row.NegotiatedRate.HasValue)
        {
            return;
        }

        var recalculatedPremium = CalculateNegotiatedPremiumFromRate(row);
        if (!recalculatedPremium.HasValue)
        {
            return;
        }

        if (row.NegotiatedPremium is not > 0m)
        {
            row.NegotiatedPremium = recalculatedPremium;
            return;
        }

        var ratio = recalculatedPremium.Value / row.NegotiatedPremium.Value;
        if (ratio is > 10m and < 14m)
        {
            row.NegotiatedPremium = recalculatedPremium;
        }
    }

    private void NegotiationAssumption_Changed(object sender, RoutedEventArgs e)
    {
        if (!IsInitialized || _negotiationRates.Count == 0)
        {
            return;
        }

        RecalculateNegotiation();
    }

    private void RecalculateNegotiation_Click(object sender, RoutedEventArgs e)
    {
        RecalculateNegotiation();
        SetStatus("Negotiation totals refreshed.");
    }

    private void RecalculateNegotiation()
    {
        if (_negotiationRates.Count == 0)
        {
            _modelReferenceRows.Clear();
            return;
        }

        var assumptions = ReadNegotiationAssumptions();
        if (assumptions.TargetLossRatio <= 0m)
        {
            _modelReferenceRows.Clear();
            UnderwritingRecommendedPositionTextBlock.Text = "Target loss ratio must be greater than 0.";
            return;
        }

        _modelReferenceRows.Clear();
        var calculations = new List<UnderwritingCalculationResult>();
        foreach (var row in _negotiationRates)
        {
            if (!row.CanEnterNegotiatedRate)
            {
                row.NegotiatedRate = null;
            }

            row.NegotiatedPremium ??= row.ProposedPremium;
            CorrectSavedNegotiatedPremiumScale(row);
            var rowLives = row.Lives ?? assumptions.Lives;
            var rowVolume = row.Volume ?? assumptions.Volume;
            var manualAdjustment = row.CurrentClaims.HasValue
                ? decimal.Round(row.CurrentClaims.Value * assumptions.ManualAdjustment, 2)
                : (decimal?)null;

            var calculation = _underwritingEngine.Calculate(new UnderwritingCalculationInput
            {
                SourceModule = BenefitsIntelligenceModules.UnderwritingIntelligence,
                VersionName = RenewalVersionNames.Negotiated,
                BenefitLine = row.BenefitLine,
                CurrentClaims = row.CurrentClaims,
                CurrentPremium = row.CurrentPremium,
                ProposedPremium = row.ProposedPremium,
                NegotiatedPremium = row.NegotiatedPremium,
                CurrentRate = row.CurrentRate,
                ProposedRate = row.ProposedRate,
                NegotiatedRate = row.CanEnterNegotiatedRate ? row.NegotiatedRate : null,
                Volume = rowVolume,
                Lives = rowLives,
                RateBasis = row.RateBasis,
                Trend = assumptions.Trend,
                TargetLossRatio = assumptions.TargetLossRatio,
                PoolingAdjustment = assumptions.PoolingCharge,
                IbnrAdjustment = assumptions.Ibnr,
                RetentionAdminFee = assumptions.AdminFee + assumptions.RiskCharge,
                ManualAdjustment = manualAdjustment
            });

            if (!row.UseSavedModelValues || !row.ModelledPremium.HasValue)
            {
                row.ModelledPremium = calculation.RequiredPremium;
            }

            if (!row.CanEnterNegotiatedRate)
            {
                row.ModelledRate = null;
            }
            else if (!row.UseSavedModelValues || !row.ModelledRate.HasValue)
            {
                row.ModelledRate = calculation.RequiredRate;
            }

            row.ProposedChangePercent = UnderwritingEngine.CalculatePercentChange(row.CurrentPremium, row.ProposedPremium);
            row.ModelledChangePercent = UnderwritingEngine.CalculatePercentChange(row.CurrentPremium, row.ModelledPremium);
            row.NegotiatedChangePercent = calculation.NegotiatedChangePercent;
            row.Savings = calculation.Savings;
            row.Status = DetermineNegotiationLineStatus(row);
            _modelReferenceRows.Add(NegotiationModelReferenceRow.From(row, calculation));
            calculations.Add(calculation);
        }

        UpdateNegotiationSummary(calculations);
    }

    private NegotiationAssumptionValues ReadNegotiationAssumptions()
    {
        return new NegotiationAssumptionValues
        {
            Trend = ReadPercentDecimal(TrendAssumptionTextBox.Text),
            TargetLossRatio = ReadPercentDecimal(TargetLossRatioTextBox.Text, 0.763m),
            PoolingCharge = ReadPercentDecimal(PoolingChargeTextBox.Text),
            PoolingThreshold = ReadMoneyDecimal(PoolingThresholdTextBox.Text),
            Ibnr = ReadPercentDecimal(IbnrTextBox.Text),
            AdminFee = ReadPercentDecimal(AdminFeeTextBox.Text),
            RiskCharge = ReadPercentDecimal(RiskChargeTextBox.Text),
            Lives = ParsingHelpers.ParseInt(LivesAssumptionTextBox.Text),
            Volume = ReadMoneyDecimal(VolumeAssumptionTextBox.Text),
            CredibilityFactor = ReadPercentDecimal(CredibilityFactorTextBox.Text, 1m),
            ManualAdjustment = ReadPercentDecimal(ManualAdjustmentTextBox.Text)
        };
    }

    private static decimal ReadPercentDecimal(string text, decimal fallback = 0m)
    {
        var value = ParsingHelpers.ParseDecimal(text);
        if (!value.HasValue)
        {
            return fallback;
        }

        return Math.Abs(value.Value) > 1m ? value.Value / 100m : value.Value;
    }

    private static decimal? ReadMoneyDecimal(string text)
    {
        var value = ParsingHelpers.ParseDecimal(text);
        return value is 0m ? null : value;
    }

    private static string DetermineNegotiationLineStatus(NegotiationRateRow row)
    {
        if (!row.CurrentPremium.HasValue || !row.ProposedPremium.HasValue)
        {
            return RenewalQueueStatuses.MissingData;
        }

        if (row.NegotiatedPremium.HasValue && row.ProposedPremium.HasValue && row.NegotiatedPremium.Value < row.ProposedPremium.Value)
        {
            return RenewalQueueStatuses.Negotiated;
        }

        if (row.ProposedChangePercent > 10m)
        {
            return RenewalQueueStatuses.HighRisk;
        }

        if (row.ProposedChangePercent is >= 5m)
        {
            return RenewalQueueStatuses.Watch;
        }

        return RenewalQueueStatuses.Healthy;
    }

    private void UpdateNegotiationSummary(IReadOnlyCollection<UnderwritingCalculationResult> calculations)
    {
        var current = _negotiationRates.SumValue(row => row.CurrentPremium);
        var proposed = _negotiationRates.SumValue(row => row.ProposedPremium);
        var modelled = _negotiationRates.SumValue(row => row.ModelledPremium);
        var negotiated = _negotiationRates.SumValue(row => row.NegotiatedPremium);
        var savings = proposed.HasValue && negotiated.HasValue ? proposed.Value - negotiated.Value : (decimal?)null;
        var gapToModel = negotiated.HasValue && modelled.HasValue ? negotiated.Value - modelled.Value : (decimal?)null;
        var proposedIncrease = UnderwritingEngine.CalculatePercentChange(current, proposed);
        var modelledIncrease = UnderwritingEngine.CalculatePercentChange(current, modelled);
        var negotiatedIncrease = UnderwritingEngine.CalculatePercentChange(current, negotiated);
        var status = DetermineNegotiationStatus(proposedIncrease, negotiatedIncrease, savings);

        NegotiationCurrentPremiumTextBlock.Text = FormatMoney(current);
        NegotiationCarrierProposalTextBlock.Text = FormatMoney(proposed);
        NegotiationModelledPremiumTextBlock.Text = FormatMoney(modelled);
        NegotiationNegotiatedPremiumTextBlock.Text = FormatMoney(negotiated);
        NegotiationSavingsTextBlock.Text = FormatMoney(savings);
        NegotiationFinalChangeTextBlock.Text = FormatPercent(negotiatedIncrease);

        CarrierProposedIncreaseTextBlock.Text = FormatPercent(proposedIncrease);
        ModelledRequiredIncreaseTextBlock.Text = FormatPercent(modelledIncrease);
        NegotiatedIncreaseTextBlock.Text = FormatPercent(negotiatedIncrease);
        SavingsAchievedTextBlock.Text = FormatMoney(savings);
        GapToModelTextBlock.Text = FormatMoney(gapToModel);
        NegotiationStatusTextBlock.Text = status;
        ApplyNegotiationStatus(status);

        UnderwritingCurrentClaimsTextBlock.Text = FormatMoney(calculations.SumValue(item => item.CurrentClaims));
        UnderwritingTrendedClaimsTextBlock.Text = FormatMoney(calculations.SumValue(item => item.TrendedClaims));
        UnderwritingProjectedClaimsTextBlock.Text = FormatMoney(calculations.SumValue(item => item.ProjectedClaims));
        UnderwritingRequiredPremiumTextBlock.Text = FormatMoney(calculations.SumValue(item => item.RequiredPremium));
        UnderwritingModelledChangeTextBlock.Text = FormatPercent(modelledIncrease);
        var projectedClaims = calculations.SumValue(item => item.ProjectedClaims);
        var projectedLossRatio = proposed is > 0m && projectedClaims.HasValue ? projectedClaims.Value / proposed.Value * 100m : (decimal?)null;
        UnderwritingProjectedLossRatioTextBlock.Text = FormatPercent(projectedLossRatio, includeSign: false);
        UnderwritingRecommendedPositionTextBlock.Text = RecommendedPositionText(status, negotiatedIncrease, savings);
    }

    private static string DetermineNegotiationStatus(decimal? proposedIncrease, decimal? negotiatedIncrease, decimal? savings)
    {
        if (!proposedIncrease.HasValue)
        {
            return RenewalQueueStatuses.MissingData;
        }

        if (savings > 0m)
        {
            return RenewalQueueStatuses.Negotiated;
        }

        if (proposedIncrease > 10m || negotiatedIncrease > 10m)
        {
            return RenewalQueueStatuses.HighRisk;
        }

        if (proposedIncrease is >= 5m || negotiatedIncrease is >= 5m)
        {
            return RenewalQueueStatuses.Watch;
        }

        return RenewalQueueStatuses.Healthy;
    }

    private static string RecommendedPositionText(string status, decimal? negotiatedIncrease, decimal? savings)
    {
        if (status == RenewalQueueStatuses.MissingData)
        {
            return "Resolve missing premium or claims data before export.";
        }

        if (status == RenewalQueueStatuses.Negotiated)
        {
            return $"Negotiated position is {FormatPercent(negotiatedIncrease)} with {FormatMoney(savings)} in savings.";
        }

        if (status == RenewalQueueStatuses.HighRisk)
        {
            return "Prepare a negotiation ask before presenting to the client.";
        }

        return "Renewal is within a normal review range.";
    }

    private void ApplyNegotiationStatus(string status)
    {
        var sample = new RenewalQueueItem { Status = status };
        NegotiationHeaderStatusTextBlock.Text = status;
        NegotiationHeaderStatusBadge.Background = sample.StatusBackground;
        NegotiationHeaderStatusBadge.BorderBrush = sample.StatusBorder;
        NegotiationHeaderStatusBadge.BorderThickness = new Thickness(1);
        NegotiationHeaderStatusTextBlock.Foreground = sample.StatusForeground;
    }

    private void LoadRevisedRateSheet_Click(object sender, RoutedEventArgs e)
    {
        if (_selectedNegotiationRenewal is null)
        {
            System.Windows.MessageBox.Show(this, "Open a client renewal in the Negotiation Workbench first.", "Load Revised Rate Sheet", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var dialog = new OpenFileDialog
        {
            Title = "Load Revised Carrier Rate Sheet",
            Filter = "Carrier rate sheets (*.pdf;*.xlsx;*.xlsm)|*.pdf;*.xlsx;*.xlsm|PDF files (*.pdf)|*.pdf|Excel workbooks (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|All files (*.*)|*.*",
            Multiselect = false,
            CheckFileExists = true
        };
        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        RevisedRateScheduleImportResult import;
        try
        {
            import = _revisedRateScheduleImporter.Import(dialog.FileName);
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            System.Windows.MessageBox.Show(this, ex.Message, "Revised rate import failed", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        if (!import.Succeeded || import.Schedule is null)
        {
            var message = import.Errors.Count == 0
                ? "The file does not contain a recognized revised rate table."
                : string.Join(Environment.NewLine, import.Errors);
            System.Windows.MessageBox.Show(this, message, "Revised rate sheet needs review", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var schedule = import.Schedule;
        var identityWarnings = new List<string>();
        var identityErrors = ValidateRevisedRateScheduleIdentity(schedule, identityWarnings);
        if (identityErrors.Count > 0)
        {
            System.Windows.MessageBox.Show(
                this,
                "The revised rate sheet was not applied because its identity does not match the selected renewal:" + Environment.NewLine +
                string.Join(Environment.NewLine, identityErrors.Select(error => $"- {error}")),
                "Client or policy mismatch",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        var mappings = BuildRevisedRateScheduleMappings(schedule, out var unmatched, out var ambiguous, out var calculationWarnings);
        if (ambiguous.Count > 0)
        {
            System.Windows.MessageBox.Show(
                this,
                "The revised rate sheet was not applied because these rows matched more than one workbench row:" + Environment.NewLine +
                string.Join(Environment.NewLine, ambiguous.Select(item => $"- {item}")),
                "Review rate mappings",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        if (mappings.Count == 0)
        {
            var details = unmatched.Count == 0 ? "No benefit rows matched the selected renewal." : string.Join(Environment.NewLine, unmatched.Select(item => $"- {item}"));
            System.Windows.MessageBox.Show(this, details, "No revised rates were mapped", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var previewLines = mappings.Take(24).Select(mapping =>
            $"{mapping.Source.SourceBenefit} -> {mapping.Target.BenefitLine} / {mapping.Target.ClassName}: " +
            $"rate {FormatImportedRate(mapping.Source.RevisedRate)}, premium {FormatImportedPremium(RevisedRateNegotiationRecovery.ScaleImportedPremium(mapping.Source, mapping.Target))}").ToList();
        if (mappings.Count > previewLines.Count)
        {
            previewLines.Add($"...and {mappings.Count - previewLines.Count} more row(s).");
        }

        var reviewNotes = identityWarnings
            .Concat(schedule.Warnings.Take(5))
            .Concat(calculationWarnings)
            .Concat(unmatched.Select(item => $"Not mapped: {item}"))
            .Distinct()
            .ToList();
        var notesText = reviewNotes.Count == 0
            ? ""
            : Environment.NewLine + Environment.NewLine + "Review notes:" + Environment.NewLine + string.Join(Environment.NewLine, reviewNotes.Select(note => $"- {note}"));
        var confirmation =
            $"Selected renewal: {_selectedNegotiationRenewal.Client} | {_selectedNegotiationRenewal.Carrier} | {_selectedNegotiationRenewal.RenewalDateText}{Environment.NewLine}" +
            $"Source: {schedule.SourceFileName} ({schedule.SourceKind}{(schedule.SheetName == "PDF" ? "" : $", {schedule.SheetName}")}){Environment.NewLine}{Environment.NewLine}" +
            string.Join(Environment.NewLine, previewLines) + notesText + Environment.NewLine + Environment.NewLine +
            "The carrier proposal will remain unchanged. These values will fill the Negotiated columns and remain manually editable. Nothing is saved until you click Save Negotiation." + Environment.NewLine + Environment.NewLine +
            "Apply these revised values?";
        if (System.Windows.MessageBox.Show(this, confirmation, "Confirm Revised Rate Sheet", MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
        {
            return;
        }

        ApplyRevisedRateSchedule(schedule, mappings);
        ExportVersionComboBox.SelectedIndex = 2;
        RecalculateNegotiation();
        if (_analysis is not null && IsCurrentAnalysisContext(_selectedNegotiationRenewal))
        {
            ApplyWorkbenchVersionToAnalysis(RenewalVersionNames.Negotiated);
        }

        NegotiationRatesGrid.Items.Refresh();
        RefreshNetGrossRateScheduleReview();
        RevisedRateScheduleStatusTextBlock.Text =
            $"Loaded {mappings.Count} revised row(s) from {schedule.SourceFileName}. Manual edits are still available; click Save Negotiation to persist.";
        SetStatus($"Loaded {mappings.Count} revised rate row(s). Save Negotiation before exporting the Negotiated version.");
    }

    private List<string> ValidateRevisedRateScheduleIdentity(RevisedRateSchedule schedule, ICollection<string> warnings)
    {
        var errors = new List<string>();
        var selected = _selectedNegotiationRenewal;
        if (selected is null)
        {
            return ["No renewal is selected."];
        }

        var targetPolicy = FirstNonBlank(
            selected.PolicyNumber,
            IsCurrentAnalysisContext(selected) ? CurrentAnalysisPolicyNumber() : "");
        var sourcePolicyKey = RevisedRateScheduleImporter.PolicyKey(schedule.PolicyNumber);
        var targetPolicyKey = RevisedRateScheduleImporter.PolicyKey(targetPolicy);
        var policiesCanBeCompared = sourcePolicyKey.Length > 0 && targetPolicyKey.Length > 0;
        var policiesMatch = policiesCanBeCompared &&
                            sourcePolicyKey.Equals(targetPolicyKey, StringComparison.Ordinal);

        if (!string.IsNullOrWhiteSpace(schedule.ClientName) &&
            !RevisedRateScheduleImporter.ClientMatches(schedule.ClientName, selected.Client))
        {
            if (policiesMatch)
            {
                warnings.Add($"Policy '{schedule.PolicyNumber}' matches, but the source client text '{schedule.ClientName}' differs from '{selected.Client}'. Confirm the client label in the mapping preview.");
            }
            else
            {
                errors.Add($"Client '{schedule.ClientName}' does not match '{selected.Client}'.");
            }
        }

        if (!string.IsNullOrWhiteSpace(schedule.Carrier) &&
            !RevisedRateScheduleImporter.CarrierMatches(schedule.Carrier, selected.Carrier))
        {
            errors.Add($"Carrier '{schedule.Carrier}' does not match '{selected.Carrier}'.");
        }

        if (policiesCanBeCompared && !policiesMatch)
        {
            errors.Add($"Policy '{schedule.PolicyNumber}' does not match '{targetPolicy}'.");
        }

        if (schedule.RenewalDate.HasValue && selected.RenewalDate.HasValue && schedule.RenewalDate != selected.RenewalDate)
        {
            errors.Add($"Effective date {schedule.RenewalDate:MMM d yyyy} does not match {selected.RenewalDate:MMM d yyyy}.");
        }

        if (string.IsNullOrWhiteSpace(schedule.ClientName)) warnings.Add("The source did not expose a client name; confirm the selected client in the mapping preview.");
        if (string.IsNullOrWhiteSpace(schedule.PolicyNumber) || string.IsNullOrWhiteSpace(targetPolicy)) warnings.Add("The policy number could not be independently cross-checked.");
        if (!schedule.RenewalDate.HasValue) warnings.Add("The source did not expose a renewal effective date; confirm the selected renewal date.");
        return errors;
    }

    private List<(RevisedRateScheduleRow Source, NegotiationRateRow Target)> BuildRevisedRateScheduleMappings(
        RevisedRateSchedule schedule,
        out List<string> unmatched,
        out List<string> ambiguous,
        out List<string> calculationWarnings)
    {
        var mappings = new List<(RevisedRateScheduleRow Source, NegotiationRateRow Target)>();
        unmatched = [];
        ambiguous = [];
        calculationWarnings = [];

        foreach (var source in schedule.Rows)
        {
            var candidates = FindRevisedRateScheduleTargets(source).ToList();
            if (candidates.Count == 0)
            {
                unmatched.Add($"{source.SourceBenefit} / {FirstNonBlank(source.TierOrClass, source.SourceClass, "All")} (no matching renewal benefit row)");
                continue;
            }

            if (candidates.Count > 1)
            {
                ambiguous.Add($"{source.SourceBenefit} / {FirstNonBlank(source.TierOrClass, source.SourceClass, "All")} -> {string.Join(", ", candidates.Select(row => $"{row.BenefitLine} / {row.ClassName}"))}");
                continue;
            }

            var target = candidates[0];
            if (source.RevisedRate.HasValue && !source.RevisedPremium.HasValue && !target.CanEnterNegotiatedRate)
            {
                unmatched.Add($"{source.SourceBenefit} / {target.ClassName} (rate-only source cannot update this premium-based row)");
                continue;
            }

            if (source.RevisedRate.HasValue && !source.RevisedPremium.HasValue && target.CanEnterNegotiatedRate)
            {
                var preview = new NegotiationRateRow
                {
                    BenefitLine = target.BenefitLine,
                    ClassName = target.ClassName,
                    AccountClass = target.AccountClass,
                    RateBasis = target.RateBasis,
                    Volume = target.Volume,
                    Lives = target.Lives,
                    CurrentRate = target.CurrentRate,
                    ProposedRate = target.ProposedRate,
                    CurrentPremium = target.CurrentPremium,
                    ProposedPremium = target.ProposedPremium,
                    NegotiatedRate = source.RevisedRate
                };
                if (!CalculateNegotiatedPremiumFromRate(preview).HasValue)
                {
                    calculationWarnings.Add($"{source.SourceBenefit} / {target.ClassName}: revised rate mapped, but premium could not be recalculated; enter the negotiated premium manually.");
                }
            }

            mappings.Add((source, target));
        }

        return mappings;
    }

    private IEnumerable<NegotiationRateRow> FindRevisedRateScheduleTargets(RevisedRateScheduleRow source)
    {
        var benefitKey = RevisedRateScheduleImporter.BenefitKey(source.CanonicalBenefit);
        var candidates = _negotiationRates
            .Where(target => RevisedRateScheduleImporter.BenefitKey(target.BenefitLine).Equals(benefitKey, StringComparison.Ordinal))
            .ToList();

        var sourceTier = RevisedRateScheduleImporter.TierKey(source.SourceBenefit, source.TierOrClass);
        if (!string.IsNullOrWhiteSpace(sourceTier))
        {
            candidates = candidates.Where(target =>
                RevisedRateScheduleImporter.TierKey(target.BenefitLine, target.ClassName, target.RateBasis).Equals(sourceTier, StringComparison.Ordinal)).ToList();
        }

        var sourceClass = source.SourceClass;
        if (!string.IsNullOrWhiteSpace(RevisedRateScheduleImporter.ClassKey(sourceClass)))
        {
            var classMatches = candidates.Where(target => RevisedRateScheduleImporter.ClassMatches(sourceClass, target.AccountClass)).ToList();
            if (classMatches.Count > 0)
            {
                return classMatches;
            }
        }

        return candidates.Count == 1 ? candidates : [];
    }

    private static void ApplyRevisedRateSchedule(
        RevisedRateSchedule schedule,
        IReadOnlyList<(RevisedRateScheduleRow Source, NegotiationRateRow Target)> mappings)
    {
        foreach (var mapping in mappings)
        {
            var source = mapping.Source;
            var target = mapping.Target;
            if (source.RevisedRate.HasValue && target.CanEnterNegotiatedRate)
            {
                target.NegotiatedRate = source.RevisedRate;
                if (target.ClientPresentationRateBasis.Equals(RateValueBasisNames.Gross, StringComparison.OrdinalIgnoreCase))
                {
                    target.VersionGrossRate = source.RevisedRate;
                }

                // Revised schedules are rate inputs for rate-based rows. Recalculate the
                // premium from the canonical exposure instead of copying a carrier premium
                // into a calculation field.
                target.NegotiatedPremium = CalculateNegotiatedPremiumFromRate(target) ?? target.NegotiatedPremium;
            }
            else
            {
                var importedPremium = RevisedRateNegotiationRecovery.ScaleImportedPremium(source, target);
                if (importedPremium.HasValue)
                {
                    target.NegotiatedPremium = importedPremium;
                    if (target.CanEnterNegotiatedRate)
                    {
                        target.NegotiatedRate = CalculateNegotiatedRateFromPremium(target) ?? target.NegotiatedRate;
                    }
                }
            }

            target.RateScheduleSourceFile = schedule.SourceFileName;
            target.RateScheduleSourceSheet = schedule.SheetName;
            target.RateScheduleSourceRow = source.SourceRow > 0 ? source.SourceRow : null;
            target.SourceReference = FirstNonBlank(source.SourceReference, $"{schedule.SourceFileName} | {schedule.SheetName}");
            target.Status = RenewalQueueStatuses.Negotiated;
            target.Notes = string.IsNullOrWhiteSpace(target.Notes)
                ? "Revised carrier rate sheet imported; manual override remains available."
                : target.Notes + " Revised carrier rate sheet imported; manual override remains available.";
        }
    }

    private static string FormatImportedRate(decimal? value) => value.HasValue ? value.Value.ToString("0.0000", CultureInfo.InvariantCulture) : "not supplied";

    private static string FormatImportedPremium(decimal? value) => value.HasValue ? value.Value.ToString("$#,0.00", CultureInfo.InvariantCulture) : "calculated from rate";

    private void LoadNetGrossRateSchedule_Click(object sender, RoutedEventArgs e)
    {
        if (_selectedNegotiationRenewal is null)
        {
            System.Windows.MessageBox.Show(this, "Open a client renewal in the Negotiation Workbench first.", "Load Net/Gross Rate Schedule", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var dialog = new OpenFileDialog
        {
            Title = "Load Net and Gross Rate Schedule",
            Filter = "Excel workbooks (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|All files (*.*)|*.*",
            Multiselect = false,
            CheckFileExists = true
        };
        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        NetGrossRateScheduleImportResult import;
        try
        {
            import = _netGrossRateScheduleImporter.Import(dialog.FileName);
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            System.Windows.MessageBox.Show(this, ex.Message, "Rate schedule import failed", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        if (!import.Succeeded || import.Schedule is null)
        {
            var message = import.Errors.Count == 0
                ? "The workbook does not contain a recognized Net and Gross rate schedule."
                : string.Join(Environment.NewLine, import.Errors);
            System.Windows.MessageBox.Show(this, message, "Rate schedule needs review", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var schedule = import.Schedule;
        var ambiguous = FindAmbiguousNetGrossScheduleMatches(schedule);
        if (ambiguous.Count > 0)
        {
            System.Windows.MessageBox.Show(
                this,
                "The schedule was not applied because these benefits matched more than one workbench row:" + Environment.NewLine +
                string.Join(Environment.NewLine, ambiguous.Select(item => $"- {item}")),
                "Review rate mappings",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        var preview = string.Join(
            Environment.NewLine,
            schedule.Rows.Select(row =>
                $"{row.SourceBenefit}: Gross {row.CurrentGrossRate:0.000} -> {row.RenewalGrossRate:0.000} -> {row.NegotiatedGrossRate:0.000}"));
        var confirmation =
            $"Dual-basis schedule detected in '{schedule.SheetName}'.{Environment.NewLine}{Environment.NewLine}" +
            $"{preview}{Environment.NewLine}{Environment.NewLine}" +
            "Excel will retain Net and Gross rates. Client PowerPoint will use Gross only." + Environment.NewLine +
            "Values will be imported exactly as displayed; the app will not recalculate Gross from percentage headings." + Environment.NewLine + Environment.NewLine +
            "Apply these rates to the open renewal?";
        if (System.Windows.MessageBox.Show(this, confirmation, "Confirm Net/Gross Rate Schedule", MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
        {
            return;
        }

        ApplyNetGrossRateSchedule(schedule);
        ExportVersionComboBox.SelectedIndex = 2;
        RecalculateNegotiation();
        if (_analysis is not null && IsCurrentAnalysisContext(_selectedNegotiationRenewal))
        {
            ApplyWorkbenchVersionToAnalysis(RenewalVersionNames.Negotiated);
        }

        NegotiationRatesGrid.Items.Refresh();
        RefreshNetGrossRateScheduleReview();
        SetStatus($"Loaded {schedule.Rows.Count} Net/Gross rate row(s). Save Negotiation to persist them.");
    }

    private List<string> FindAmbiguousNetGrossScheduleMatches(NetGrossRateSchedule schedule)
    {
        var ambiguous = new List<string>();
        foreach (var source in schedule.Rows)
        {
            var candidates = FindNetGrossScheduleTargets(source).ToList();
            if (candidates.Count > 1)
            {
                ambiguous.Add($"{source.SourceBenefit} -> {string.Join(", ", candidates.Select(row => $"{row.BenefitLine} / {row.ClassName}"))}");
            }
        }

        return ambiguous;
    }

    private IEnumerable<NegotiationRateRow> FindNetGrossScheduleTargets(NetGrossRateScheduleRow source)
    {
        var sourceBenefit = NetGrossRateScheduleImporter.BenefitKey(source.CanonicalBenefit);
        var sourceTier = NetGrossRateScheduleImporter.TierKey(source.SourceBenefit, source.TierOrClass);
        return _negotiationRates.Where(target =>
        {
            if (!NetGrossRateScheduleImporter.BenefitKey(target.BenefitLine).Equals(sourceBenefit, StringComparison.Ordinal))
            {
                return false;
            }

            var targetTier = NetGrossRateScheduleImporter.TierKey(target.BenefitLine, target.ClassName);
            return string.IsNullOrWhiteSpace(sourceTier)
                ? string.IsNullOrWhiteSpace(targetTier)
                : targetTier.Equals(sourceTier, StringComparison.Ordinal);
        });
    }

    private void ApplyNetGrossRateSchedule(NetGrossRateSchedule schedule)
    {
        foreach (var source in schedule.Rows)
        {
            var target = FindNetGrossScheduleTargets(source).SingleOrDefault();
            if (target is null)
            {
                target = new NegotiationRateRow
                {
                    BenefitLine = source.CanonicalBenefit,
                    ClassName = FirstNonBlank(source.TierOrClass, "All"),
                    AccountClass = "All Classes",
                    RateBasis = NetGrossPricingBasis(source),
                    Status = RenewalQueueStatuses.Negotiated,
                    Notes = "Imported detailed rate-only row. Premium remains sourced from the renewal report."
                };
                _negotiationRates.Add(target);
            }

            ApplyGrossPremiumBasis(target, source);
            target.CurrentNetRate = source.CurrentNetRate;
            target.CurrentGrossRate = source.CurrentGrossRate;
            target.ProposedNetRate = source.RenewalNetRate;
            target.ProposedGrossRate = source.RenewalGrossRate;
            target.VersionNetRate = source.NegotiatedNetRate;
            target.VersionGrossRate = source.NegotiatedGrossRate;
            target.CurrentRate = source.CurrentGrossRate;
            target.ProposedRate = source.RenewalGrossRate;
            target.NegotiatedRate = source.NegotiatedGrossRate;
            target.RateValueBasis = RateValueBasisNames.Gross;
            target.ClientPresentationRateBasis = RateValueBasisNames.Gross;
            target.RateScheduleSourceFile = schedule.SourceFileName;
            target.RateScheduleSourceSheet = schedule.SheetName;
            target.RateScheduleSourceRow = source.SourceRow;
            target.SourceReference = $"{schedule.SourceFileName} | {schedule.SheetName} | row {source.SourceRow}";
            target.Status = RenewalQueueStatuses.Negotiated;

            var negotiatedPremium = CalculateNegotiatedPremiumFromRate(target);
            if (negotiatedPremium.HasValue)
            {
                target.NegotiatedPremium = negotiatedPremium;
            }
        }

        SortNegotiationRates();
    }

    private void SortNegotiationRates()
    {
        var ordered = BenefitDisplayOrder.Sort(
            _negotiationRates,
            item => item.BenefitLine,
            item => item.ClassName,
            item => item.AccountClass);
        if (ordered.SequenceEqual(_negotiationRates))
        {
            return;
        }

        _negotiationRates.Clear();
        foreach (var row in ordered)
        {
            _negotiationRates.Add(row);
        }
    }

    private static void ApplyGrossPremiumBasis(NegotiationRateRow target, NetGrossRateScheduleRow source)
    {
        var premiumScale = PremiumScaleFromComparablePremium(target.CurrentPremium, source.CurrentNetRate, target)
                           ?? PremiumScaleFromComparablePremium(target.ProposedPremium, source.RenewalNetRate, target)
                           ?? 1m;

        target.CurrentPremium = CalculatePremiumAtRate(target, source.CurrentGrossRate, premiumScale)
                                ?? target.CurrentPremium;
        target.ProposedPremium = CalculatePremiumAtRate(target, source.RenewalGrossRate, premiumScale)
                                 ?? target.ProposedPremium;
        target.NegotiatedPremium = CalculatePremiumAtRate(target, source.NegotiatedGrossRate, premiumScale)
                                   ?? target.NegotiatedPremium;
    }

    private static decimal? CalculatePremiumAtRate(NegotiationRateRow row, decimal rate, decimal premiumScale)
    {
        var premium = UnderwritingEngine.CalculatePremium(row.BenefitLine, row.Volume, row.Lives, rate, row.RateBasis);
        return premium.HasValue ? decimal.Round(premium.Value * premiumScale, 2) : null;
    }

    private static string NetGrossPricingBasis(NetGrossRateScheduleRow row)
    {
        if (!string.IsNullOrWhiteSpace(row.TierOrClass))
        {
            return row.TierOrClass.ToLowerInvariant();
        }

        var benefit = NetGrossRateScheduleImporter.BenefitKey(row.CanonicalBenefit);
        return benefit is "life" or "add" ? "volume" : "per member";
    }

    private void ResetNetGrossRateScheduleReview()
    {
        _netGrossRateScheduleReviewRows.Clear();
        NetGrossRateScheduleReviewBorder.Visibility = Visibility.Collapsed;
        NetGrossRateScheduleStatusTextBlock.Text = "Use the Net/Gross loader only when the schedule has separate Net and Gross rate columns.";
        NegotiationCurrentRateColumn.Header = "Current Rate";
        NegotiationProposedRateColumn.Header = "Proposed Rate";
        NegotiationNegotiatedRateColumn.Header = "Negotiated Rate";
        NegotiationCurrentPremiumColumn.Header = "Current Premium";
        NegotiationProposedPremiumColumn.Header = "Proposed Premium";
        NegotiationModelledPremiumColumn.Header = "Modelled Premium";
        NegotiationNegotiatedPremiumColumn.Header = "Negotiated Premium";
    }

    private void ResetRevisedRateScheduleStatus()
    {
        RevisedRateScheduleStatusTextBlock.Text =
            "Optional. Import a revised carrier PDF or Excel rate sheet; manual Negotiated fields remain editable.";
    }

    private void RefreshNetGrossRateScheduleReview()
    {
        _netGrossRateScheduleReviewRows.Clear();
        foreach (var row in _negotiationRates.Where(item => item.HasNetAndGrossRates))
        {
            _netGrossRateScheduleReviewRows.Add(row);
        }

        if (_netGrossRateScheduleReviewRows.Count == 0)
        {
            ResetNetGrossRateScheduleReview();
            return;
        }

        NetGrossRateScheduleReviewBorder.Visibility = Visibility.Visible;
        NetGrossRateScheduleStatusTextBlock.Text =
            $"Dual-basis schedule active: {_netGrossRateScheduleReviewRows.Count} row(s). Premium totals and changes compare Gross to Gross; Excel = Net + Gross; PowerPoint = Gross.";
        NegotiationCurrentRateColumn.Header = "Current Gross";
        NegotiationProposedRateColumn.Header = "Renewal Gross";
        NegotiationNegotiatedRateColumn.Header = "Negotiated Gross";
        NegotiationCurrentPremiumColumn.Header = "Current Premium (Gross)";
        NegotiationProposedPremiumColumn.Header = "Proposed Premium (Gross)";
        NegotiationModelledPremiumColumn.Header = "Modelled Premium";
        NegotiationNegotiatedPremiumColumn.Header = "Negotiated Premium (Gross)";
    }

    private async void SaveNegotiation_Click(object sender, RoutedEventArgs e)
    {
        await SaveNegotiationAsync(RenewalQueueStatuses.Negotiated);
    }

    private async Task<NegotiationWorkbenchSaveRequest?> SaveNegotiationAsync(string status)
    {
        foreach (var textBox in FindVisualChildren<TextBox>(NegotiationRatesGrid))
        {
            var binding = textBox.GetBindingExpression(TextBox.TextProperty);
            if (binding?.IsDirty == true)
            {
                binding.UpdateSource();
            }
        }

        NegotiationRatesGrid.CommitEdit(DataGridEditingUnit.Cell, true);
        NegotiationRatesGrid.CommitEdit(DataGridEditingUnit.Row, true);
        var errors = ValidateNegotiation();
        if (errors.Count > 0)
        {
            System.Windows.MessageBox.Show(this, string.Join(Environment.NewLine, errors), "Review negotiation inputs", MessageBoxButton.OK, MessageBoxImage.Warning);
            SetStatus("Negotiation was not saved. Review the highlighted validation items.");
            return null;
        }

        if (_selectedNegotiationRenewal is null)
        {
            return null;
        }

        try
        {
            SetBusy(true, "Saving negotiated version...");
            RecalculateNegotiation();
            var request = BuildNegotiationSaveRequest(status);
            var result = await _databaseService.SaveNegotiationWorkbenchAsync(request);
            _latestSavedClientId = result.ClientId ?? _latestSavedClientId;
            if (!result.Succeeded)
            {
                SetStatus("Negotiation save needs review.");
                return null;
            }

            UpdateSelectedRenewalStatus(status);
            if (_analysis is not null)
            {
                _negotiationWorkbenchService.ApplySavedVersion(_analysis, request, request.VersionName);
                UpdateReview();
            }

            SetStatus($"Negotiation saved for {_selectedNegotiationRenewal.Client}.");
            return request;
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            System.Windows.MessageBox.Show(this, ex.Message, "Negotiation save failed", MessageBoxButton.OK, MessageBoxImage.Error);
            SetStatus("Negotiation save failed.");
            return null;
        }
        finally
        {
            SetBusy(false);
        }
    }

    private List<string> ValidateNegotiation()
    {
        var errors = new List<string>();
        if (FindVisualChildren<TextBox>(NegotiationRatesGrid).Any(textBox => Validation.GetHasError(textBox)))
        {
            errors.Add("Correct the invalid rate or premium entries before saving.");
        }
        if (ReadNegotiationAssumptions().TargetLossRatio <= 0m)
        {
            errors.Add("Target loss ratio must be greater than 0.");
        }

        foreach (var row in _negotiationRates)
        {
            if (row.CurrentRate < 0m || row.ProposedRate < 0m)
            {
                errors.Add($"{row.BenefitLine}: current and proposed rates cannot be negative.");
            }

            if (row.NegotiatedRate < 0m)
            {
                errors.Add($"{row.BenefitLine}: negotiated rate cannot be negative.");
            }

            if (row.NegotiatedPremium < 0m)
            {
                errors.Add($"{row.BenefitLine}: negotiated premium cannot be negative.");
            }

            if (row.HasNetAndGrossRates && !row.HasCompleteNetAndGrossRates)
            {
                errors.Add($"{row.BenefitLine}: Net/Gross schedule values are incomplete. Current, Renewal, and Negotiated Net and Gross rates are required.");
                continue;
            }

            if (row.IsCompleteImportedRateOnlyRow)
            {
                continue;
            }

            if (!row.CurrentPremium.HasValue)
            {
                errors.Add($"{row.BenefitLine}: current premium is missing.");
            }

            if (!row.ProposedPremium.HasValue)
            {
                errors.Add($"{row.BenefitLine}: proposed premium is missing.");
            }
        }

        return errors.Distinct().ToList();
    }

    private NegotiationWorkbenchSaveRequest BuildNegotiationSaveRequest(string status)
    {
        var row = _selectedNegotiationRenewal ?? new RenewalQueueItem { Client = "Unknown Client" };
        var assumptions = ReadNegotiationAssumptions();
        var calculations = _negotiationRates.Select(rate => _underwritingEngine.Calculate(new UnderwritingCalculationInput
        {
            SourceModule = BenefitsIntelligenceModules.UnderwritingIntelligence,
            VersionName = status == RenewalQueueStatuses.FinalApproved ? RenewalVersionNames.FinalApproved : RenewalVersionNames.Negotiated,
            BenefitLine = rate.BenefitLine,
            CurrentClaims = rate.CurrentClaims,
            CurrentPremium = rate.CurrentPremium,
            ProposedPremium = rate.ProposedPremium,
            NegotiatedPremium = rate.NegotiatedPremium,
            CurrentRate = rate.CurrentRate,
            ProposedRate = rate.ProposedRate,
            NegotiatedRate = rate.CanEnterNegotiatedRate ? rate.NegotiatedRate : null,
            Volume = rate.Volume ?? assumptions.Volume,
            Lives = rate.Lives ?? assumptions.Lives,
            RateBasis = rate.RateBasis,
            Trend = assumptions.Trend,
            TargetLossRatio = assumptions.TargetLossRatio,
            PoolingAdjustment = assumptions.PoolingCharge,
            IbnrAdjustment = assumptions.Ibnr,
            RetentionAdminFee = assumptions.AdminFee + assumptions.RiskCharge,
            ManualAdjustment = rate.CurrentClaims.HasValue ? decimal.Round(rate.CurrentClaims.Value * assumptions.ManualAdjustment, 2) : null
        })).ToList();

        var request = new NegotiationWorkbenchSaveRequest
        {
            ClientId = row.ClientId,
            RenewalId = row.RenewalId,
            ClientName = row.Client,
            Carrier = row.Carrier,
            EffectiveDate = row.RenewalDate,
            RenewalYear = row.RenewalDate?.Year ?? DateTime.Today.Year,
            VersionName = status == RenewalQueueStatuses.FinalApproved ? RenewalVersionNames.FinalApproved : RenewalVersionNames.Negotiated,
            Status = status,
            CreatedBy = "Consultant",
            Notes = ConsultantNotesTextBox.Text.Trim(),
            Assumptions = assumptions.ToSaveRecord(),
            Note = BuildNegotiationNote(row, status)
        };

        request.Rates.AddRange(_negotiationRates.Select(rate => rate.ToSaveRecord()));
        request.Calculations.AddRange(calculations);
        return request;
    }

    private NegotiationNoteRecord BuildNegotiationNote(RenewalQueueItem row, string status)
    {
        _ = int.TryParse(NegotiationRoundTextBox.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var round);
        return new NegotiationNoteRecord
        {
            RenewalId = row.RenewalId,
            ClientId = row.ClientId,
            RoundNumber = round <= 0 ? 1 : round,
            NoteDate = ParsingHelpers.ParseDateOnly(NegotiationDateTextBox.Text) ?? DateOnly.FromDateTime(DateTime.Today),
            CarrierResponse = CarrierResponseTextBox.Text.Trim(),
            ConsultantNotes = ConsultantNotesTextBox.Text.Trim(),
            NextAction = NextActionTextBox.Text.Trim(),
            Status = ComboBoxText(NegotiationNoteStatusComboBox, status),
            CreatedBy = "Consultant"
        };
    }

    private static string ComboBoxText(ComboBox comboBox, string fallback)
    {
        return comboBox.SelectedItem is ComboBoxItem item
            ? item.Content?.ToString() ?? fallback
            : fallback;
    }

    private void UpdateSelectedRenewalStatus(string status)
    {
        if (_selectedNegotiationRenewal is null)
        {
            return;
        }

        _selectedNegotiationRenewal.Status = status;
        _selectedNegotiationRenewal.ActionLabel = "Review";
        ApplyNegotiationStatus(status);
        _renewalQueueView?.Refresh();
        UpdateRenewalQueueSummary();
        UpdateSelectedRenewalDetail(_selectedNegotiationRenewal);
    }

    private async void GenerateRenewalWorkbook_Click(object sender, RoutedEventArgs e)
    {
        var versionName = SelectedExportVersionName();
        var savedRequest = await SaveNegotiationAsync(StatusForVersion(versionName));
        if (savedRequest is not null)
        {
            await ExportCurrentRenewalFromWorkbenchAsync("Workbook", versionName, savedRequest);
        }
    }

    private async void GenerateRenewalPowerPoint_Click(object sender, RoutedEventArgs e)
    {
        var versionName = SelectedExportVersionName();
        var savedRequest = await SaveNegotiationAsync(StatusForVersion(versionName));
        if (savedRequest is not null)
        {
            await ExportCurrentRenewalFromWorkbenchAsync("PowerPoint", versionName, savedRequest);
        }
    }

    private async void ExportNegotiationPack_Click(object sender, RoutedEventArgs e)
    {
        var versionName = SelectedExportVersionName();
        var savedRequest = await SaveNegotiationAsync(StatusForVersion(versionName));
        if (savedRequest is not null)
        {
            await ExportCurrentRenewalFromWorkbenchAsync("Export Pack", versionName, savedRequest);
        }
    }

    private async Task ExportCurrentRenewalFromWorkbenchAsync(
        string actionName,
        string versionName,
        NegotiationWorkbenchSaveRequest savedRequest)
    {
        if (_selectedNegotiationRenewal is null)
        {
            return;
        }

        var restoreError = await EnsureAnalysisContextForExportAsync(_selectedNegotiationRenewal);
        if (restoreError is not null)
        {
            NegotiationExportStatusTextBlock.Text = $"{actionName}: {restoreError}";
            SetStatus($"{actionName}: {restoreError}");
            return;
        }

        if (_negotiationRates.Any(row => row.HasNetAndGrossRates) &&
            versionName is not RenewalVersionNames.Negotiated and not RenewalVersionNames.FinalApproved)
        {
            var message = $"{actionName}: this renewal has a Net/Gross schedule. Choose Negotiated or Final Approved so the client output uses the confirmed Gross rates.";
            NegotiationExportStatusTextBlock.Text = message;
            SetStatus(message);
            System.Windows.MessageBox.Show(this, message, "Choose a client presentation version", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var exported = await ExportAnalysisAsync(versionName, savedRequest, useLoadedAnalysis: true);
        NegotiationExportStatusTextBlock.Text = exported
            ? $"{actionName}: review and PowerPoint generated with {versionName}."
            : $"{actionName}: export did not complete.";
    }

    private string SelectedExportVersionName()
    {
        var selected = ComboBoxText(ExportVersionComboBox, RenewalVersionNames.Negotiated);
        return selected switch
        {
            RenewalVersionNames.CarrierProposed => RenewalVersionNames.CarrierProposed,
            RenewalVersionNames.ConsultantModel => RenewalVersionNames.ConsultantModel,
            RenewalVersionNames.FinalApproved => RenewalVersionNames.FinalApproved,
            _ => RenewalVersionNames.Negotiated
        };
    }

    private static string StatusForVersion(string versionName)
    {
        return versionName switch
        {
            RenewalVersionNames.FinalApproved => RenewalQueueStatuses.FinalApproved,
            RenewalVersionNames.CarrierProposed => RenewalQueueStatuses.ReadyToReview,
            RenewalVersionNames.ConsultantModel => RenewalQueueStatuses.ReadyToReview,
            _ => RenewalQueueStatuses.Negotiated
        };
    }

    private void ApplyWorkbenchVersionToAnalysis(string versionName)
    {
        if (_analysis is null)
        {
            return;
        }

        if (versionName is RenewalVersionNames.CarrierProposed or RenewalVersionNames.ConsultantModel)
        {
            _analysis.SelectedRenewalVersionName = versionName;
            return;
        }

        var version = _analysis.RenewalVersions.FirstOrDefault(item => item.VersionName.Equals(versionName, StringComparison.OrdinalIgnoreCase));
        if (version is null)
        {
            version = new RenewalVersion
            {
                VersionName = versionName,
                SourceModule = BenefitsIntelligenceModules.RenewalIntelligence,
                Status = StatusForVersion(versionName)
            };
            _analysis.RenewalVersions.Add(version);
        }

        version.RateLines.Clear();
        version.Calculations.Clear();
        version.Status = StatusForVersion(versionName);
        version.CreatedAt = DateTime.Now;

        var assumptions = ReadNegotiationAssumptions();
        foreach (var row in _negotiationRates)
        {
            var versionPremium = versionName == RenewalVersionNames.FinalApproved
                ? row.NegotiatedPremium ?? row.ModelledPremium ?? row.ProposedPremium
                : row.NegotiatedPremium ?? row.ProposedPremium;
            var versionRate = versionName == RenewalVersionNames.FinalApproved
                ? (row.CanEnterNegotiatedRate ? row.NegotiatedRate ?? row.ModelledRate ?? row.ProposedRate : null)
                : (row.CanEnterNegotiatedRate ? row.NegotiatedRate ?? row.ProposedRate : null);

            version.RateLines.Add(new RenewalVersionRateLine
            {
                Benefit = row.BenefitLine,
                TierOrClass = row.ClassName,
                AccountClass = row.AccountClass,
                Volume = row.Volume,
                Lives = row.Lives,
                CurrentRate = row.CurrentRate,
                CurrentPremium = row.CurrentPremium,
                ProposedRate = row.ProposedRate,
                ProposedPremium = row.ProposedPremium,
                VersionRate = versionRate,
                VersionPremium = versionPremium,
                ChangePercent = UnderwritingEngine.CalculatePercentChange(row.CurrentPremium, versionPremium),
                RateBasis = row.RateBasis,
                RateValueBasis = row.RateValueBasis,
                ClientPresentationRateBasis = row.ClientPresentationRateBasis,
                CurrentNetRate = row.CurrentNetRate,
                CurrentGrossRate = row.CurrentGrossRate,
                ProposedNetRate = row.ProposedNetRate,
                ProposedGrossRate = row.ProposedGrossRate,
                VersionNetRate = row.VersionNetRate,
                VersionGrossRate = row.ClientPresentationRateBasis.Equals(RateValueBasisNames.Gross, StringComparison.OrdinalIgnoreCase)
                    ? versionRate
                    : row.VersionGrossRate,
                RateScheduleSourceFile = row.RateScheduleSourceFile,
                RateScheduleSourceSheet = row.RateScheduleSourceSheet,
                RateScheduleSourceRow = row.RateScheduleSourceRow,
                SourceReference = FirstNonBlank(row.SourceReference, "Renewal Negotiation Workbench")
            });

            version.Calculations.Add(_underwritingEngine.Calculate(new UnderwritingCalculationInput
            {
                SourceModule = BenefitsIntelligenceModules.RenewalIntelligence,
                VersionName = versionName,
                BenefitLine = row.BenefitLine,
                CurrentClaims = row.CurrentClaims,
                CurrentPremium = row.CurrentPremium,
                ProposedPremium = row.ProposedPremium,
                NegotiatedPremium = versionPremium,
                CurrentRate = row.CurrentRate,
                ProposedRate = row.ProposedRate,
                NegotiatedRate = row.CanEnterNegotiatedRate ? versionRate : null,
                Volume = row.Volume ?? assumptions.Volume,
                Lives = row.Lives ?? assumptions.Lives,
                RateBasis = row.RateBasis,
                Trend = assumptions.Trend,
                TargetLossRatio = assumptions.TargetLossRatio,
                PoolingAdjustment = assumptions.PoolingCharge,
                IbnrAdjustment = assumptions.Ibnr,
                RetentionAdminFee = assumptions.AdminFee + assumptions.RiskCharge,
                ManualAdjustment = row.CurrentClaims.HasValue ? decimal.Round(row.CurrentClaims.Value * assumptions.ManualAdjustment, 2) : null
            }));
        }

        _analysis.SelectedRenewalVersionName = versionName;
        UpdateReview();
    }

    private async void MarkReadyToReview_Click(object sender, RoutedEventArgs e)
    {
        await SaveNegotiationAsync(RenewalQueueStatuses.ReadyToReview);
    }

    private async void MarkFinalApproved_Click(object sender, RoutedEventArgs e)
    {
        await SaveNegotiationAsync(RenewalQueueStatuses.FinalApproved);
    }

    private void UpdateSelectedRenewalDetail(RenewalQueueItem? row)
    {
        if (row is null)
        {
            SelectedClientTitleTextBlock.Text = "Selected Client Detail";
            DetailClientTextBlock.Text = "-";
            DetailCarrierTextBlock.Text = "-";
            DetailRenewalDateTextBlock.Text = "-";
            DetailPremiumsTextBlock.Text = "-";
            DetailCurrentPremiumTextBlock.Text = "-";
            DetailRenewalPremiumTextBlock.Text = "-";
            DetailPremiumIncreaseTextBlock.Text = "-";
            DetailEhcLossRatioTextBlock.Text = "-";
            DetailDentalLossRatioTextBlock.Text = "-";
            DetailPoolingChargeTextBlock.Text = "-";
            DetailOtherMetricsTextBlock.Text = "-";
            QuickInsightsItemsControl.ItemsSource = Array.Empty<string>();
            MissingDataItemsControl.ItemsSource = Array.Empty<string>();
            return;
        }

        SelectedClientTitleTextBlock.Text = row.Client;
        SelectedStatusTextBlock.Text = row.Status;
        SelectedStatusBadge.Background = row.StatusBackground;
        SelectedStatusBadge.BorderBrush = row.StatusBorder;
        SelectedStatusBadge.BorderThickness = new Thickness(1);
        SelectedStatusTextBlock.Foreground = row.StatusForeground;

        DetailClientTextBlock.Text = row.Client;
        DetailCarrierTextBlock.Text = row.Carrier;
        DetailRenewalDateTextBlock.Text = row.RenewalDateText;
        DetailPremiumsTextBlock.Text = $"{row.CurrentPremiumText} current / {row.RenewalPremiumText} renewal";
        DetailCurrentPremiumTextBlock.Text = row.CurrentPremiumText;
        DetailRenewalPremiumTextBlock.Text = row.RenewalPremiumText;
        DetailPremiumIncreaseTextBlock.Text = row.PremiumChangeText;
        DetailEhcLossRatioTextBlock.Text = row.EhcLossRatioText;
        DetailDentalLossRatioTextBlock.Text = row.DentalLossRatioText;
        DetailPoolingChargeTextBlock.Text = row.PoolingChargeText;
        DetailOtherMetricsTextBlock.Text = $"{row.LivesText} lives; {row.TrendAssumptions}; pooling {row.PoolingCharges}";
        QuickInsightsItemsControl.ItemsSource = row.QuickInsights.Select(insight => "- " + insight).ToList();
        MissingDataItemsControl.ItemsSource = row.MissingDataItems.Count == 0
            ? ["- No required data gaps identified."]
            : row.MissingDataItems.Select(item => "- " + item).ToList();
    }

    private void UpdateRenewalQueueSummary()
    {
        var visibleRows = _renewalQueueView?.Cast<RenewalQueueItem>().ToList() ?? _renewalQueue.ToList();
        TotalRenewalsTextBlock.Text = _renewalQueue.Count.ToString(CultureInfo.InvariantCulture);
        ReadyForReviewTextBlock.Text = _renewalQueue.Count(row => row.Status == RenewalQueueStatuses.ReadyToReview).ToString(CultureInfo.InvariantCulture);
        HighRiskTextBlock.Text = _renewalQueue.Count(row => row.Status == RenewalQueueStatuses.HighRisk).ToString(CultureInfo.InvariantCulture);
        MissingDataTextBlock.Text = _renewalQueue.Count(row => row.Status == RenewalQueueStatuses.MissingData).ToString(CultureInfo.InvariantCulture);
        CompletedTextBlock.Text = _renewalQueue.Count(row => row.Status == RenewalQueueStatuses.Completed).ToString(CultureInfo.InvariantCulture);
        RenewalQueueCountTextBlock.Text = $"{visibleRows.Count} shown";
    }

    private static int StatusRank(string status)
    {
        return status switch
        {
            RenewalQueueStatuses.HighRisk => 1,
            RenewalQueueStatuses.MissingData => 2,
            RenewalQueueStatuses.ReadyToReview => 3,
            RenewalQueueStatuses.Negotiated => 4,
            RenewalQueueStatuses.Watch => 5,
            RenewalQueueStatuses.Healthy => 6,
            RenewalQueueStatuses.FinalApproved => 7,
            RenewalQueueStatuses.Completed => 8,
            _ => 9
        };
    }

    private static int MonthRank(string month)
    {
        return DateTime.TryParseExact(month, "MMMM", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsed)
            ? parsed.Month
            : 13;
    }

    private void BrowsePdf_Click(object sender, RoutedEventArgs e)
    {
        BrowseFile(PdfPathTextBox, "PDF files (*.pdf)|*.pdf|All files (*.*)|*.*");
    }

    private void BrowseWorkbook_Click(object sender, RoutedEventArgs e)
    {
        BrowseFile(WorkbookPathTextBox, "Excel files (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls|Excel workbooks (*.xlsx)|*.xlsx|Legacy Excel workbooks (*.xls)|*.xls|All files (*.*)|*.*");
    }

    private void BrowseTemplate_Click(object sender, RoutedEventArgs e)
    {
        BrowseFile(TemplatePathTextBox, "PowerPoint files (*.pptx)|*.pptx|All files (*.*)|*.*");
    }

    private void BrowseOutput_Click(object sender, RoutedEventArgs e)
    {
        BrowseFolder(OutputFolderTextBox, "Choose where the review pack should be saved");
    }

    private void BrowseCurrentYearFolder_Click(object sender, RoutedEventArgs e)
    {
        BrowseFolder(CurrentYearFolderTextBox, "Choose the claims experience folder");
    }

    private void BrowsePriorYearFolder_Click(object sender, RoutedEventArgs e)
    {
        BrowseFolder(PriorYearFolderTextBox, "Choose the prior year folder");
    }

    private void BrowseAsoBillingFolder_Click(object sender, RoutedEventArgs e)
    {
        BrowseFolder(AsoBillingFolderTextBox, "Choose the ASO billing folder with year and month subfolders");
    }

    private void OpenHcsaCalculator_Click(object sender, RoutedEventArgs e)
    {
        var window = new HcsaCalculatorWindow(_hcsaCalculation)
        {
            Owner = this
        };

        if (window.ShowDialog() != true)
        {
            return;
        }

        _hcsaCalculation = window.Calculation;
        UpdateHcsaCalculationSummary();
        SaveSettings();
        InvalidateAnalysis();
        SetStatus("HCSA funding applied. Validate the renewal to add it to the workbook and PowerPoint rate tables.");
    }

    private void UpdateHcsaCalculationSummary()
    {
        if (_hcsaCalculation.HasEntries)
        {
            HcsaCalculationSummaryTextBlock.Text = $"{_hcsaCalculation.Classes.Count(item => item.IsIncluded)} class(es) | Current {FormatCurrencyCents(_hcsaCalculation.CurrentAnnualCost)} | Renewal {FormatCurrencyCents(_hcsaCalculation.RenewalAnnualCost)} annual";
            HcsaCalculationSummaryTextBlock.Foreground = Brushes.Teal;
            return;
        }

        HcsaCalculationSummaryTextBlock.Text = "Not added";
        HcsaCalculationSummaryTextBlock.Foreground = (Brush)FindResource("MutedNavyBrush");
    }

    private void OpenAllstateCriticalIllnessEnrollment_Click(object sender, RoutedEventArgs e)
    {
        var window = new AllstateCriticalIllnessEnrollmentWindow(_allstateCriticalIllnessEnrollment)
        {
            Owner = this
        };

        if (window.ShowDialog() != true)
        {
            return;
        }

        _allstateCriticalIllnessEnrollment = window.Enrollment;
        UpdateAllstateCriticalIllnessEnrollmentSummary();
        SaveSettings();
        InvalidateAnalysis();
        SetStatus(_allstateCriticalIllnessEnrollment.HasEntries
            ? "Allstate CI enrollment counts applied. Validate the renewal to calculate the stand-alone premium."
            : "Allstate CI enrollment counts cleared.");
    }

    private void UpdateAllstateCriticalIllnessEnrollmentSummary()
    {
        if (_allstateCriticalIllnessEnrollment.HasEntries)
        {
            AllstateCriticalIllnessEnrollmentSummaryTextBlock.Text =
                $"EE {_allstateCriticalIllnessEnrollment.EmployeeOnlyCount} | Family {_allstateCriticalIllnessEnrollment.FamilyCount} | Total {_allstateCriticalIllnessEnrollment.TotalEnrolled}";
            AllstateCriticalIllnessEnrollmentSummaryTextBlock.Foreground = Brushes.Teal;
            return;
        }

        AllstateCriticalIllnessEnrollmentSummaryTextBlock.Text = "Counts required for premium";
        AllstateCriticalIllnessEnrollmentSummaryTextBlock.Foreground = (Brush)FindResource("MutedNavyBrush");
    }

    private void AddAdditionalPdfs_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "PDF files (*.pdf)|*.pdf|All files (*.*)|*.*",
            CheckFileExists = true,
            Multiselect = true
        };

        if (dialog.ShowDialog() != true)
        {
            return;
        }

        foreach (var fileName in dialog.FileNames)
        {
            if (_additionalPdfPaths.Any(path => path.Equals(fileName, StringComparison.OrdinalIgnoreCase)))
            {
                continue;
            }

            _additionalPdfPaths.Add(fileName);
        }

        InvalidateAnalysis();
    }

    private void RemoveAdditionalPdfs_Click(object sender, RoutedEventArgs e)
    {
        var selected = AdditionalPdfsListBox.SelectedItems.Cast<string>().ToList();
        foreach (var path in selected)
        {
            _additionalPdfPaths.Remove(path);
        }

        InvalidateAnalysis();
    }

    private void AddAdditionalWorkbooks_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Excel files (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls|Excel workbooks (*.xlsx)|*.xlsx|Excel macro workbooks (*.xlsm)|*.xlsm|Legacy Excel workbooks (*.xls)|*.xls|All files (*.*)|*.*",
            CheckFileExists = true,
            Multiselect = true
        };

        if (dialog.ShowDialog() != true)
        {
            return;
        }

        foreach (var fileName in dialog.FileNames)
        {
            if (_additionalWorkbookPaths.Any(path => path.Equals(fileName, StringComparison.OrdinalIgnoreCase)))
            {
                continue;
            }

            _additionalWorkbookPaths.Add(fileName);
        }

        InvalidateAnalysis();
    }

    private void RemoveAdditionalWorkbooks_Click(object sender, RoutedEventArgs e)
    {
        var selected = AdditionalWorkbooksListBox.SelectedItems.Cast<string>().ToList();
        foreach (var path in selected)
        {
            _additionalWorkbookPaths.Remove(path);
        }

        InvalidateAnalysis();
    }

    private void AddBenefitBookletPdfs_Click(object sender, RoutedEventArgs e)
    {
        PromptForBookletPdfFiles();
    }

    private bool PromptForBookletPdfFiles()
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Benefit booklet PDFs (*.pdf)|*.pdf|All files (*.*)|*.*",
            CheckFileExists = true,
            Multiselect = true
        };

        if (dialog.ShowDialog() != true)
        {
            return false;
        }

        foreach (var fileName in dialog.FileNames)
        {
            AddBookletInput(fileName);
        }

        SetStatus($"Added booklet PDF(s). Total selected PDFs/folders: {_bookletInputPaths.Count}.");
        return dialog.FileNames.Length > 0;
    }

    private void AddBenefitBookletFolder_Click(object sender, RoutedEventArgs e)
    {
        using var dialog = new Forms.FolderBrowserDialog
        {
            Description = "Choose a folder containing benefit booklet PDFs. Subfolders will be included.",
            UseDescriptionForTitle = true,
            SelectedPath = Directory.Exists(OutputFolderTextBox.Text)
                ? OutputFolderTextBox.Text
                : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        };

        if (dialog.ShowDialog() != Forms.DialogResult.OK)
        {
            return;
        }

        AddBookletInput(dialog.SelectedPath);
        var pdfCount = BookletExtractorRunner.CountPdfFiles(_bookletInputPaths);
        SetStatus($"Added booklet folder. Found {pdfCount} booklet PDF file(s) in the selected input(s).");
    }

    private void RemoveBenefitBooklets_Click(object sender, RoutedEventArgs e)
    {
        var selected = BenefitBookletsListBox.SelectedItems.Cast<string>().ToList();
        foreach (var path in selected)
        {
            _bookletInputPaths.Remove(path);
        }

        _bookletExtractionResult = null;
        _bookletExtractionInputSnapshot.Clear();
        _claimsDataModelExport = null;
        SetStatus("Removed selected booklet input(s).");
    }

    private void CancelBookletExtraction_Click(object sender, RoutedEventArgs e)
    {
        _bookletExtractionCancellation?.Cancel();
        SetStatus("Cancelling booklet extraction...");
    }

    private async void RebuildBenefitsMasterDatabase_Click(object sender, RoutedEventArgs e)
    {
        var outputFolder = OutputFolderTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(outputFolder))
        {
            var documents = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            outputFolder = Path.Combine(documents, "Renewal Review Builder", "Outputs");
            OutputFolderTextBox.Text = outputFolder;
        }

        try
        {
            SetBusy(true, "Rebuilding Benefits Master Database from SQLite...");
            SaveSettings();
            var export = await _benefitsMasterDatabaseExportService.RebuildAsync(outputFolder);
            ExportTextBlock.Text =
                $"Benefits master database: {export.MasterWorkbookPath}{Environment.NewLine}" +
                $"Source of truth: {AppPaths.BenefitsIntelligenceDatabasePath}{Environment.NewLine}" +
                $"Clients exported: {export.ClientCount}{Environment.NewLine}" +
                $"Plan Design rows exported: {export.PlanDesignRowCount}{Environment.NewLine}" +
                $"Benefit Summary sheets: {export.SummarySheetCount}";
            SetOpenOutputButtonsEnabled(Directory.Exists(outputFolder));
            System.Windows.MessageBox.Show(
                this,
                $"Benefits Master Database rebuilt from SQLite.{Environment.NewLine}{Environment.NewLine}" +
                export.MasterWorkbookPath,
                "Benefits Master Database rebuilt",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            SetStatus($"Benefits Master Database rebuilt from SQLite with {export.SummarySheetCount} Benefit Summary sheet(s).");
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            System.Windows.MessageBox.Show(
                this,
                ex.Message,
                "Benefits Master Database rebuild failed",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            SetStatus("Benefits Master Database rebuild failed. SQLite data was not changed.");
        }
        finally
        {
            SetBusy(false);
        }
    }

    private async void ExtractBookletToMasterDb_Click(object sender, RoutedEventArgs e)
    {
        if (_bookletInputPaths.Count == 0)
        {
            var selectedFiles = PromptForBookletPdfFiles();
            if (selectedFiles && _bookletInputPaths.Count > 0)
            {
                SetStatus($"Ready to extract {_bookletInputPaths.Count} selected booklet input(s).");
            }
        }

        if (_bookletInputPaths.Count == 0)
        {
            System.Windows.MessageBox.Show(
                this,
                "Choose one or more booklet PDFs, or use Add Folder in the Benefit Booklet PDFs area.",
                "No booklet PDFs selected",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        var targetClient = await PromptForBookletExtractionTargetClientAsync();
        if (targetClient is null)
        {
            SetStatus("Booklet extraction cancelled. No destination client was confirmed.");
            return;
        }

        var outputFolder = OutputFolderTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(outputFolder))
        {
            var documents = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            outputFolder = Path.Combine(documents, "Renewal Review Builder", "Outputs");
            OutputFolderTextBox.Text = outputFolder;
        }

        try
        {
            SetBusy(true, "Starting booklet extraction...");
            _bookletExtractionCancellation = new CancellationTokenSource();
            CancelBookletExtractionButton.IsEnabled = true;
            SaveSettings();

            var progress = new Progress<BookletExtractionProgress>(status => SetStatus(status.Message));
            _bookletExtractionResult = await _bookletExtractorRunner.RunAsync(
                new BookletExtractionRequest
                {
                    InputPaths = _bookletInputPaths.ToList(),
                    OutputFolder = outputFolder
                },
                progress,
                _bookletExtractionCancellation.Token);
            _bookletExtractionInputSnapshot = _bookletInputPaths.ToList();
            var databaseSave = await _databaseService.SavePlanDesignFromBookletExtractionAsync(
                _bookletExtractionResult,
                _bookletInputPaths,
                targetClient.ClientId);
            BenefitsMasterDatabaseWorkbookResult? masterRebuild = null;
            string masterRebuildError = "";
            if (databaseSave.Succeeded)
            {
                try
                {
                    masterRebuild = await _benefitsMasterDatabaseExportService.RebuildAsync(
                        outputFolder,
                        [targetClient.ClientId],
                        cancellationToken: _bookletExtractionCancellation.Token);
                }
                catch (Exception ex)
                {
                    AppLogger.Error(ex);
                    masterRebuildError = ex.Message;
                }
            }

            var targetDescription = $"{targetClient.ClientName} (client {targetClient.ClientId.ToString(CultureInfo.InvariantCulture)})";
            var masterWorkbookPath = masterRebuild?.MasterWorkbookPath ?? _bookletExtractionResult.MasterWorkbookPath;

            ExportTextBlock.Text =
                $"Benefits master database: {masterWorkbookPath}{Environment.NewLine}" +
                $"Extraction review workbook: {_bookletExtractionResult.ReviewWorkbookPath}{Environment.NewLine}" +
                $"Total PDFs found: {_bookletExtractionResult.TotalPdfsFound}{Environment.NewLine}" +
                $"Successfully processed: {_bookletExtractionResult.SuccessfullyProcessed}{Environment.NewLine}" +
                $"Errors: {_bookletExtractionResult.ErrorCount}{Environment.NewLine}" +
                $"Review required count: {_bookletExtractionResult.ReviewRequiredCount}{Environment.NewLine}" +
                $"Destination client: {targetDescription}" +
                DatabaseStatusLine(databaseSave) +
                (masterRebuild is null
                    ? string.IsNullOrWhiteSpace(masterRebuildError)
                        ? ""
                        : $"{Environment.NewLine}Master workbook rebuild: Failed - {masterRebuildError}"
                    : $"{Environment.NewLine}Master workbook rebuild: {masterRebuild.SummarySheetCount} summary sheet(s) generated from SQLite.");

            SetOpenOutputButtonsEnabled(Directory.Exists(_bookletExtractionResult.OutputFolder));
            var messageIcon = _bookletExtractionResult.ErrorCount > 0 || !databaseSave.Succeeded || !string.IsNullOrWhiteSpace(masterRebuildError)
                ? MessageBoxImage.Warning
                : MessageBoxImage.Information;
            var databaseOutcome = databaseSave.Succeeded
                ? $"Plan Design saved to SQLite:{Environment.NewLine}{targetDescription}" +
                  (masterRebuild is null
                      ? $"{Environment.NewLine}{Environment.NewLine}The workbook could not be rebuilt. Your SQLite data is safe; use Rebuild Master DB to retry.{Environment.NewLine}{masterRebuildError}"
                      : $"{Environment.NewLine}{Environment.NewLine}Benefits Master Database regenerated from SQLite.")
                : $"Plan Design was not saved to {targetDescription}.{Environment.NewLine}{databaseSave.Message}";
            System.Windows.MessageBox.Show(
                this,
                $"Booklet extraction complete.{Environment.NewLine}{Environment.NewLine}" +
                $"Benefits master database:{Environment.NewLine}{masterWorkbookPath}{Environment.NewLine}{Environment.NewLine}" +
                $"Review workbook:{Environment.NewLine}{_bookletExtractionResult.ReviewWorkbookPath}{Environment.NewLine}{Environment.NewLine}" +
                databaseOutcome,
                databaseSave.Succeeded && masterRebuild is not null
                    ? "Booklet extraction complete"
                    : databaseSave.Succeeded
                        ? "Plan Design saved - workbook rebuild needs attention"
                        : "Workbook created - database save blocked",
                MessageBoxButton.OK,
                messageIcon);
            SetStatus(databaseSave.Succeeded && masterRebuild is not null
                ? $"Booklet extraction complete. SQLite saved and master workbook rebuilt for {targetDescription}."
                : databaseSave.Succeeded
                    ? $"Plan Design saved to SQLite for {targetDescription}; workbook rebuild needs attention."
                : $"Booklet workbook created, but Plan Design was not saved. {databaseSave.Message}");
        }
        catch (OperationCanceledException)
        {
            SetStatus("Booklet extraction cancelled.");
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            System.Windows.MessageBox.Show(this, ex.Message, "Booklet extraction failed", MessageBoxButton.OK, MessageBoxImage.Error);
            SetStatus("Booklet extraction failed.");
        }
        finally
        {
            _bookletExtractionCancellation?.Dispose();
            _bookletExtractionCancellation = null;
            CancelBookletExtractionButton.IsEnabled = false;
            SetBusy(false);
        }
    }

    private async Task<ClientRecord?> PromptForBookletExtractionTargetClientAsync()
    {
        try
        {
            var clients = await _databaseService.GetClientsAsync();
            if (clients.Count == 0)
            {
                System.Windows.MessageBox.Show(
                    this,
                    "Create or import the client before extracting booklet Plan Design data.",
                    "No saved clients",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return null;
            }

            var dialog = new BookletExtractionTargetClientDialog(
                clients,
                _bookletInputPaths.Select(path => Path.GetFileName(path)).OfType<string>())
            {
                Owner = this
            };
            return dialog.ShowDialog() == true ? dialog.SelectedClient : null;
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            System.Windows.MessageBox.Show(
                this,
                $"The destination client list could not be loaded: {ex.Message}",
                "Choose booklet destination",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            return null;
        }
    }

    private void AddBookletInput(string path)
    {
        if (string.IsNullOrWhiteSpace(path) ||
            _bookletInputPaths.Any(existing => existing.Equals(path, StringComparison.OrdinalIgnoreCase)))
        {
            return;
        }

        _bookletInputPaths.Add(path);
        _bookletExtractionResult = null;
        _bookletExtractionInputSnapshot.Clear();
        _claimsDataModelExport = null;
    }

    private void ClearForm_Click(object sender, RoutedEventArgs e)
    {
        var result = System.Windows.MessageBox.Show(
            this,
            "Clear the saved file paths and current results? This will not delete any exported files.",
            "Clear form",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question);

        if (result != MessageBoxResult.Yes)
        {
            return;
        }

        _isLoadingDefaults = true;
        PdfPathTextBox.Text = "";
        WorkbookPathTextBox.Text = "";
        TemplatePathTextBox.Text = "";
        CurrentYearFolderTextBox.Text = "";
        PriorYearFolderTextBox.Text = "";
        AsoBillingFolderTextBox.Text = "";
        _hcsaCalculation = new HcsaCalculation();
        UpdateHcsaCalculationSummary();
        _allstateCriticalIllnessEnrollment = new AllstateCriticalIllnessEnrollment();
        UpdateAllstateCriticalIllnessEnrollmentSummary();
        SetReportMonthCount(0);
        _additionalPdfPaths.Clear();
        _additionalWorkbookPaths.Clear();
        _bookletInputPaths.Clear();

        var documents = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        OutputFolderTextBox.Text = Path.Combine(documents, "Renewal Review Builder", "Outputs");
        _isLoadingDefaults = false;

        _analysis = null;
        _analysisClientId = null;
        _analysisSourceDocumentId = null;
        _export = null;
        _claimsDataModelExport = null;
        _bookletExtractionResult = null;
        _bookletExtractionInputSnapshot.Clear();
        _proposalOutputFolder = null;
        _renewalQueueExportFolder = null;
        ResetReviewDisplay();
        SetOpenOutputButtonsEnabled(false);
        _settingsStore.Save(new UserSettings { OutputFolder = OutputFolderTextBox.Text.Trim() });
        SetStatus("Form cleared. Saved paths removed.");
    }

    private void OpenWestlandSharePoint_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(WestlandSharePointDocumentsUrl))
        {
            System.Windows.MessageBox.Show(this,
                "Your SharePoint location has not been configured. Use Browse to select downloaded or synced documents.",
                "SharePoint location", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = WestlandSharePointDocumentsUrl,
                UseShellExecute = true
            });
            SetStatus("Opened Westland SharePoint Documents. Choose Clients - Active, then return here and use Browse to select downloaded or synced files.");
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            System.Windows.MessageBox.Show(
                this,
                $"Westland SharePoint could not be opened. {ex.Message}",
                "Open Westland SharePoint",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            SetStatus("Westland SharePoint could not be opened.");
        }
    }

    private void StartCommand_Click(object sender, RoutedEventArgs e)
    {
        var launcherPath = Path.Combine(AppContext.BaseDirectory, "Start Renewal Review Builder.cmd");
        if (!File.Exists(launcherPath))
        {
            System.Windows.MessageBox.Show(
                this,
                "Start Renewal Review Builder.cmd was not found beside the app. Use the portable folder that contains both the app and the start command.",
                "Start command not found",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        Process.Start(new ProcessStartInfo
        {
            FileName = launcherPath,
            WorkingDirectory = AppContext.BaseDirectory,
            UseShellExecute = true
        });
    }

    private void OpenRenewalIntelligence_Click(object sender, RoutedEventArgs e)
    {
        SetNavigationWorkspace("BookOfBusiness", "Renewal Pipeline");
        UploadRenewalPackage_Click(sender, e);
    }

    private void OpenClaimsExperience_Click(object sender, RoutedEventArgs e)
    {
        ShowClaimsExperienceDashboard();
    }

    private void OpenBenefitBooklets_Click(object sender, RoutedEventArgs e)
    {
        SetNavigationWorkspace(_selectedClientWorkspace is null ? "DocumentOperations" : "ClientWorkspace", "Plan Design", _selectedClientWorkspace?.ClientName);
        if (_bookletInputPaths.Count == 0)
        {
            PromptForBookletPdfFiles();
        }

        if (_bookletInputPaths.Count > 0)
        {
            SaveSettings();
            SetStatus("Benefit booklet input selected. Use Extract Booklet to Master DB when ready.");
        }
    }

    private async void OpenProposalIntelligence_Click(object sender, RoutedEventArgs e)
    {
        SetNavigationWorkspace(_selectedClientWorkspace is null ? "BookOfBusiness" : "ClientWorkspace", "Proposals", _selectedClientWorkspace?.ClientName);
        await ShowProposalPipelineOverviewAsync();
    }

    private async void RefreshProposalPipelineOverview_Click(object sender, RoutedEventArgs e)
    {
        await ShowProposalPipelineOverviewAsync();
    }

    private void OpenQuotesAlternateOptions_Click(object sender, RoutedEventArgs e)
    {
        ShowPlaceholder(
            "Quotes / Alternate Options",
            "This will become the saved quote and alternate option workbench. Use Upload Quote / Option or Compare Carrier Proposals when you want to select files or a carrier folder.");
    }

    private void OpenDocumentIntelligence_Click(object sender, RoutedEventArgs e)
    {
        SetNavigationWorkspace("DocumentOperations", "Document Operations");
        HomeDashboardScrollViewer.Visibility = Visibility.Collapsed;
        DashboardScrollViewer.Visibility = Visibility.Collapsed;
        UnderwritingWorkbenchScrollViewer.Visibility = Visibility.Collapsed;
        AssumptionLibraryScrollViewer.Visibility = Visibility.Collapsed;
        ReportsAnalyticsScrollViewer.Visibility = Visibility.Collapsed;
        BenchmarkingDashboardScrollViewer.Visibility = Visibility.Collapsed;
        ProposalComparisonScrollViewer.Visibility = Visibility.Collapsed;
        PlanDesignIntelligenceScrollViewer.Visibility = Visibility.Collapsed;
        PlanDesignFactorsScrollViewer.Visibility = Visibility.Collapsed;
        ClientsScrollViewer.Visibility = Visibility.Collapsed;
        NegotiationWorkbenchScrollViewer.Visibility = Visibility.Collapsed;
        DataSourcesScrollViewer.Visibility = Visibility.Collapsed;
        DocumentIntelligenceScrollViewer.Visibility = Visibility.Visible;
        DocumentIntelligenceScrollViewer.ScrollToTop();
        SetOperationalFooterVisible(true);
        UpdateDocumentBatchKpis();
        SetStatus("Opened Document Intelligence batch processing.");
    }

    private void SelectBillingStatement_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Billing statement PDFs (*.pdf)|*.pdf|All files (*.*)|*.*",
            CheckFileExists = true,
            Multiselect = false,
            InitialDirectory = Directory.Exists(Path.GetDirectoryName(BillingStatementPathTextBox.Text))
                ? Path.GetDirectoryName(BillingStatementPathTextBox.Text)
                : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        };

        if (dialog.ShowDialog() == true)
        {
            BillingStatementPathTextBox.Text = dialog.FileName;
            SetStatus("Billing statement selected.");
        }
    }

    private async void ExtractBillingStatement_Click(object sender, RoutedEventArgs e)
    {
        var billingPath = BillingStatementPathTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(billingPath) || !File.Exists(billingPath))
        {
            System.Windows.MessageBox.Show(this, "Select a billing statement PDF first.", "Billing Statement", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        try
        {
            SetBusy(true, "Extracting billing premiums...");
            var outputFolder = Path.Combine(
                ResolveOutputRootFolder(),
                "Billing Statement Extraction",
                $"{SafeFolderName(Path.GetFileNameWithoutExtension(billingPath))}_{DateTime.Now:yyyyMMdd_HHmmss}");
            var import = await _manulifeBillingImportCoordinator.ImportAsync(new ManulifeBillingImportRequest
            {
                InputPath = billingPath,
                OutputFolder = outputFolder,
                SaveMode = ManulifeBillingSaveMode.Cancel,
                ExportCsv = true,
                PerformedBy = Environment.UserName
            });
            var result = import.Extraction;

            _billingReconciliationRows.Clear();
            foreach (var row in result?.ReconciliationResults.Take(500) ?? [])
            {
                _billingReconciliationRows.Add(row);
            }

            var failedCount = result?.ReconciliationResults.Count(row => row.Status == BillingReconciliationStatuses.Failed) ?? 0;
            var unavailableCount = result?.ReconciliationResults.Count(row => row.Status == BillingReconciliationStatuses.NotAvailable) ?? 0;
            var detailCount = result?.Details.Count ?? 0;
            var provinceCount = result?.ProvinceSummaries.Count ?? 0;
            BillingSummaryTextBlock.Text =
                $"{import.FinalStatus}: {detailCount:#,0} detail rows, {provinceCount:#,0} province rows, {failedCount:#,0} failures";

            if (!import.Succeeded)
            {
                _billingOutputFolder = null;
                ExportTextBlock.Text =
                    $"Billing status: {import.FinalStatus}{Environment.NewLine}" +
                    $"Database: {import.UserMessage}{Environment.NewLine}" +
                    "CSV export: Not run because the database save did not succeed.";
                SetOpenOutputButtonsEnabled(false);
                SetStatus("Billing import did not complete. No success files were exported.");
                var failureIcon = import.FinalStatus == BillingImportStatuses.NeedsReview
                    ? MessageBoxImage.Warning
                    : MessageBoxImage.Error;
                System.Windows.MessageBox.Show(
                    this,
                    import.UserMessage,
                    "Billing import did not complete",
                    MessageBoxButton.OK,
                    failureIcon);
                return;
            }

            _billingOutputFolder = import.ExportStatus == BillingExportStatuses.Completed ? outputFolder : null;
            var csvExport = import.CsvExport;

            ExportTextBlock.Text =
                $"Billing status: {import.FinalStatus}{Environment.NewLine}" +
                $"Database: {import.Persistence?.Message}{Environment.NewLine}" +
                $"CSV export: {import.ExportStatus}{Environment.NewLine}" +
                $"Detected classes CSV: {csvExport?.ClassesDetectedPath}{Environment.NewLine}" +
                $"Reconciliation CSV: {csvExport?.ReconciliationResultsPath}{Environment.NewLine}" +
                $"Validation errors CSV: {csvExport?.ValidationErrorsPath}";
            SetOpenOutputButtonsEnabled(_billingOutputFolder is not null && Directory.Exists(_billingOutputFolder));
            SetStatus(import.FinalStatus == BillingImportStatuses.NeedsReview
                ? $"Billing extraction complete. Needs review: {failedCount:#,0} reconciliation issue(s), {unavailableCount:#,0} subtotal(s) not available."
                : "Billing extraction complete.");

            var icon = import.FinalStatus == BillingImportStatuses.NeedsReview ? MessageBoxImage.Warning : MessageBoxImage.Information;
            System.Windows.MessageBox.Show(
                this,
                $"Billing import saved.{Environment.NewLine}{Environment.NewLine}Status: {import.FinalStatus}{Environment.NewLine}Detail rows: {detailCount:#,0}{Environment.NewLine}Failed reconciliation rows: {failedCount:#,0}",
                "Billing Statement",
                MessageBoxButton.OK,
                icon);
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            System.Windows.MessageBox.Show(this, ex.Message, "Billing extraction failed", MessageBoxButton.OK, MessageBoxImage.Error);
            SetStatus("Billing extraction failed.");
        }
        finally
        {
            SetBusy(false);
        }
    }

    private async void OpenRenewalWorkspaceFromSidebar_Click(object sender, RoutedEventArgs e)
    {
        _clientRenewalSelectionIdentity = null;
        await ShowRenewalsDashboardAsync();
        SetStatus("Opened Level 3 Renewal Workspace. Select a renewal row to review one client in detail.");
    }

    private async void OpenUnderwriting_Click(object sender, RoutedEventArgs e)
    {
        await ShowUnderwritingDashboardAsync(SelectedRenewalQueueRow());
    }

    private void OpenClients_Click(object sender, RoutedEventArgs e)
    {
        ShowClientsPage();
    }

    private void ShowClientsPage()
    {
        _clientRenewalSelectionIdentity = null;
        SetNavigationWorkspace("Clients", "Master Client List");
        HomeDashboardScrollViewer.Visibility = Visibility.Collapsed;
        DashboardScrollViewer.Visibility = Visibility.Collapsed;
        UnderwritingWorkbenchScrollViewer.Visibility = Visibility.Collapsed;
        AssumptionLibraryScrollViewer.Visibility = Visibility.Collapsed;
        ReportsAnalyticsScrollViewer.Visibility = Visibility.Collapsed;
        BenchmarkingDashboardScrollViewer.Visibility = Visibility.Collapsed;
        ProposalComparisonScrollViewer.Visibility = Visibility.Collapsed;
        PlanDesignIntelligenceScrollViewer.Visibility = Visibility.Collapsed;
        PlanDesignFactorsScrollViewer.Visibility = Visibility.Collapsed;
        DocumentIntelligenceScrollViewer.Visibility = Visibility.Collapsed;
        NegotiationWorkbenchScrollViewer.Visibility = Visibility.Collapsed;
        ClientListPanel.Visibility = Visibility.Visible;
        ClientWorkspacePanel.Visibility = Visibility.Collapsed;
        ClientsScrollViewer.Visibility = Visibility.Visible;
        ClientsScrollViewer.ScrollToTop();
        SetOperationalFooterVisible(true);
        SetStatus("Opened Clients master list.");
        _ = RefreshClientListWhenVisibleAsync();
    }

    private async Task RefreshClientListWhenVisibleAsync()
    {
        if (_isRefreshingClientList)
        {
            return;
        }

        try
        {
            _isRefreshingClientList = true;
            await LoadSavedClientListAsync();
        }
        finally
        {
            _isRefreshingClientList = false;
        }
    }

    private void OpenClientWorkspace(ClientListItem client)
    {
        if (_selectedClientWorkspace is null ||
            !_selectedClientWorkspace.ClientId.Equals(client.ClientId, StringComparison.OrdinalIgnoreCase))
        {
            _claimsRangeReviewViewModel.RevokeSessionEdits();
        }

        _selectedClientWorkspace = BuildClientWorkspace(client);
        _clientRenewalSelectionIdentity = CreateClientRenewalSelectionIdentity(_selectedClientWorkspace);
        SetNavigationWorkspace("ClientWorkspace", "Overview", _selectedClientWorkspace.ClientName);
        PopulateClientWorkspace(_selectedClientWorkspace);
        ShowClientWorkspaceOverview();

        ClientListPanel.Visibility = Visibility.Collapsed;
        ClientWorkspacePanel.Visibility = Visibility.Visible;
        ClientsScrollViewer.ScrollToTop();
        SetOperationalFooterVisible(true);
        SetStatus($"Opened Client Workspace: /clients/{client.ClientId}");
    }

    private ClientWorkspaceViewModel BuildClientWorkspace(ClientListItem client)
    {
        var renewalHistory = FindRenewalHistoryForClient(client);
        var renewal = renewalHistory.FirstOrDefault() ?? FindRenewalForClient(CreateClientRenewalSelectionIdentity(client));
        var historyPresentation = _clientWorkspaceRenewalHistoryPresentationService.Build(
            renewalHistory.Select(CreateClientWorkspaceHistoryEntry),
            client.Status);

        return new ClientWorkspaceViewModel
        {
            ClientId = client.ClientId,
            DatabaseClientId = client.DatabaseClientId,
            ClientName = client.ClientName,
            Status = client.Status,
            Segment = client.Industry,
            Employees = client.Employees,
            PrimaryCarrier = client.PrimaryCarrier,
            PlanType = client.PlanType,
            HealthScore = historyPresentation.CurrentHealthScore ?? client.HealthScore,
            RenewalHistory = historyPresentation,
            LastRenewalDate = renewal?.RenewalDate,
            CurrentPremium = renewal?.CurrentPremium,
            RenewalPremium = renewal?.RenewalPremium,
            PremiumChangePercent = renewal?.PremiumChangePercent,
            EhcLossRatio = renewal?.EhcLossRatio,
            DentalLossRatio = renewal?.DentalLossRatio,
            TargetLossRatio = renewal?.TargetLossRatio,
            PoolingCharge = renewal?.PoolingCharge,
            PoolingThreshold = renewal?.PoolingThreshold,
            PoolingThresholdBasis = renewal?.PoolingThreshold.HasValue == true
                ? "Renewal experience summary"
                : "No verified terms",
            AccountExecutive = "Account executive not set",
            AccountManager = "Account manager not set",
            PrimaryContactName = "Primary contact not set",
            PrimaryContactEmail = "Add email in client profile",
            PrimaryContactPhone = "Add phone in client profile",
            PlanEffectiveDate = renewal?.RenewalDate?.AddYears(-1).ToString("MMM d, yyyy", CultureInfo.InvariantCulture) ?? "Not set"
        };
    }

    private IReadOnlyList<RenewalQueueItem> FindRenewalHistoryForClient(ClientListItem client)
    {
        IEnumerable<RenewalQueueItem> matches = [];
        if (client.DatabaseClientId is > 0)
        {
            matches = _renewalQueue.Where(row => row.ClientId == client.DatabaseClientId.Value);
        }

        if (!matches.Any())
        {
            var clientKey = NormalizeClientMatchKey(client.ClientName);
            matches = _renewalQueue.Where(row =>
                NormalizeClientMatchKey(row.Client).Equals(clientKey, StringComparison.OrdinalIgnoreCase));
        }

        return matches
            .OrderByDescending(row => row.RenewalDate?.DayNumber ?? int.MinValue)
            .ThenByDescending(row => row.DateLoaded)
            .ToList();
    }

    private void PopulateClientWorkspace(ClientWorkspaceViewModel workspace)
    {
        ClientWorkspaceBreadcrumbClientTextBlock.Text = workspace.ClientName;
        ClientWorkspaceTitleTextBlock.Text = workspace.ClientName;
        ClientWorkspaceSubtitleTextBlock.Text = workspace.RouteText;
        ClientWorkspaceMetadataTextBlock.Text = workspace.MetadataText;
        ClientWorkspaceStatusTextBlock.Text = workspace.Status;
        ClientWorkspaceStatusBadge.Background = workspace.StatusBackground;
        ClientWorkspaceStatusBadge.BorderBrush = workspace.StatusBorder;
        ClientWorkspaceStatusBadge.BorderThickness = new Thickness(1);
        ClientWorkspaceStatusTextBlock.Foreground = workspace.StatusForeground;

        ClientWorkspaceHealthScoreTextBlock.Text = workspace.HealthScoreText;
        ClientWorkspaceHealthScoreTextBlock.Foreground = workspace.HealthScoreBrush;
        ClientWorkspaceHealthScoreLabelTextBlock.Text = workspace.HealthScoreLabel;
        ClientWorkspaceHealthScoreLabelTextBlock.Foreground = workspace.HealthScoreBrush;
        PopulateClientWorkspaceRenewalHistory(workspace);
        ClientWorkspaceRenewalDateTextBlock.Text = workspace.RenewalDate;
        ClientWorkspaceRenewalCountdownTextBlock.Text = workspace.RenewalCountdownText;
        ClientWorkspaceCurrentPremiumTextBlock.Text = workspace.CurrentPremiumText;
        ClientWorkspaceRenewalPremiumTextBlock.Text = workspace.RenewalPremiumText;
        ClientWorkspacePremiumIncreaseTextBlock.Text = workspace.PremiumChangeDescription;
        ClientWorkspaceSummaryPremiumChangeTextBlock.Text = workspace.PremiumChangeText;
        ClientWorkspaceEhcLossTextBlock.Text = workspace.EhcLossRatioText;
        ClientWorkspaceDentalLossTextBlock.Text = workspace.DentalLossRatioText;
        ClientWorkspaceTargetLossRatioTextBlock.Text = workspace.TargetLossRatioText;
        ClientWorkspaceTlrBasisTextBlock.Text = workspace.TargetLossRatio.HasValue
            ? "Renewal experience summary"
            : "No verified basis";
        ClientWorkspacePoolingChargeTextBlock.Text = workspace.PoolingChargeText;
        ClientWorkspacePoolingChargeBasisTextBlock.Text = workspace.PoolingCharge.HasValue
            ? "Renewal experience summary"
            : "No verified terms";
        UpdateClientWorkspacePoolingThresholdPresentation(workspace);
        ClientWorkspaceEmployeesTextBlock.Text = workspace.EmployeesText;
        ClientWorkspaceKeyRiskTextBlock.Text = workspace.KeyRisk;

        ClientWorkspacePlanTypeTextBlock.Text = workspace.PlanType;
        ClientWorkspaceAccountPlanTypeTextBlock.Text = workspace.PlanType;
        ClientWorkspaceCarrierTextBlock.Text = workspace.PrimaryCarrier;
        ClientWorkspaceAccountRenewalDateTextBlock.Text = workspace.RenewalDate;
        ClientWorkspacePlanEffectiveDateTextBlock.Text = workspace.PlanEffectiveDate;
        ClientWorkspaceAccountExecutiveTextBlock.Text = workspace.AccountExecutive;
        ClientWorkspaceAccountManagerTextBlock.Text = workspace.AccountManager;
        ClientWorkspacePrimaryContactTextBlock.Text = workspace.PrimaryContactName;
        ClientWorkspaceContactEmailTextBlock.Text = workspace.PrimaryContactEmail;
        ClientWorkspaceContactPhoneTextBlock.Text = workspace.PrimaryContactPhone;
        ClientWorkspaceFlexibleBenefitsCard.Visibility = Visibility.Collapsed;
        ClientWorkspaceHcsaSummaryTextBlock.Text = "Not confirmed";
        ClientWorkspaceLsaSummaryTextBlock.Text = "Not confirmed";
        ClientWorkspaceFlexibleAdministratorTextBlock.Text = "Not confirmed";
        ClientWorkspaceFlexibleStatusTextBlock.Text = "Not confirmed";
        ClientWorkspaceFlexibleEffectiveDateTextBlock.Text = "";
        _ = LoadClientFlexibleBenefitsOverviewAsync(workspace);
        _ = LoadClientFinancialOverviewAsync(workspace);
    }

    private void PopulateClientWorkspaceRenewalHistory(ClientWorkspaceViewModel workspace)
    {
        var history = workspace.RenewalHistory;
        ClientWorkspaceHealthScoreProgressBar.Value = workspace.HealthScore ?? 0;
        ClientWorkspaceHealthScoreProgressBar.Foreground = workspace.HealthScoreBrush;
        ClientWorkspaceHealthScoreTrendTextBlock.Text = history.HealthScoreComparison.Text;
        ClientWorkspaceHealthScoreTrendTextBlock.Foreground = ComparisonBrush(history.HealthScoreComparison.Tone);

        ClientWorkspacePremiumChangeBasisTextBlock.Text = history.Prior is null
            ? "Current to renewal"
            : $"{history.Current?.RenewalDate?.Year ?? history.Current?.RenewalYear} current to renewal";
        ClientWorkspaceEhcPriorTextBlock.Text = history.EhcComparison.Text;
        ClientWorkspaceEhcPriorTextBlock.Foreground = ComparisonBrush(history.EhcComparison.Tone);
        ClientWorkspaceDentalPriorTextBlock.Text = history.DentalComparison.Text;
        ClientWorkspaceDentalPriorTextBlock.Foreground = ComparisonBrush(history.DentalComparison.Tone);
        ClientWorkspaceSummaryPremiumChangeTextBlock.Foreground = PremiumChangeBrush(workspace.PremiumChangePercent);
        ClientWorkspaceEhcLossTextBlock.Foreground = LossRatioRiskBrush(workspace.EhcLossRatio);
        ClientWorkspaceDentalLossTextBlock.Foreground = LossRatioRiskBrush(workspace.DentalLossRatio);

        var hasPremiumHistory = history.PremiumPoints.Any(point =>
            point.CurrentPremium.HasValue || point.RenewalPremium.HasValue);
        ClientWorkspacePremiumTrendItemsControl.ItemsSource = history.PremiumPoints;
        ClientWorkspacePremiumTrendItemsControl.Visibility = hasPremiumHistory
            ? Visibility.Visible
            : Visibility.Collapsed;
        ClientWorkspacePremiumTrendEmptyTextBlock.Visibility = hasPremiumHistory
            ? Visibility.Collapsed
            : Visibility.Visible;

        var ehcPoints = history.LossRatioPoints
            .Where(point => point.EhcY.HasValue)
            .ToList();
        var dentalPoints = history.LossRatioPoints
            .Where(point => point.DentalY.HasValue)
            .ToList();
        ClientWorkspaceEhcLossTrendPolyline.Points = BuildLossRatioPoints(ehcPoints, point => point.EhcY);
        ClientWorkspaceDentalLossTrendPolyline.Points = BuildLossRatioPoints(dentalPoints, point => point.DentalY);
        ClientWorkspaceEhcLossTrendMarkers.ItemsSource = BuildLossRatioMarkers(
            ehcPoints,
            point => point.EhcY,
            point => point.EhcToolTip);
        ClientWorkspaceDentalLossTrendMarkers.ItemsSource = BuildLossRatioMarkers(
            dentalPoints,
            point => point.DentalY,
            point => point.DentalToolTip);
        var hasLossHistory = ehcPoints.Count > 0 || dentalPoints.Count > 0;
        ClientWorkspaceLossTrendChart.Visibility = hasLossHistory ? Visibility.Visible : Visibility.Collapsed;
        ClientWorkspaceLossTrendEmptyTextBlock.Visibility = hasLossHistory ? Visibility.Collapsed : Visibility.Visible;
        ClientWorkspaceLossTrendPeriodsTextBlock.Text = history.PeriodsText;
    }

    private static PointCollection BuildLossRatioPoints(
        IReadOnlyCollection<ClientWorkspaceLossRatioTrendPoint> points,
        Func<ClientWorkspaceLossRatioTrendPoint, double?> ySelector)
    {
        return new PointCollection(points
            .Select(point => new Point(point.X, ySelector(point)!.Value)));
    }

    private static IReadOnlyList<ClientWorkspaceChartMarker> BuildLossRatioMarkers(
        IReadOnlyCollection<ClientWorkspaceLossRatioTrendPoint> points,
        Func<ClientWorkspaceLossRatioTrendPoint, double?> ySelector,
        Func<ClientWorkspaceLossRatioTrendPoint, string> toolTipSelector)
    {
        return points
            .Select(point => new ClientWorkspaceChartMarker(
                point.X - 4d,
                ySelector(point)!.Value - 4d,
                toolTipSelector(point)))
            .ToList();
    }

    private static Brush ComparisonBrush(ClientWorkspaceComparisonTone tone)
    {
        return tone switch
        {
            ClientWorkspaceComparisonTone.Positive => new SolidColorBrush(Color.FromRgb(22, 163, 74)),
            ClientWorkspaceComparisonTone.Negative => new SolidColorBrush(Color.FromRgb(220, 38, 38)),
            ClientWorkspaceComparisonTone.Neutral => new SolidColorBrush(Color.FromRgb(71, 85, 105)),
            _ => new SolidColorBrush(Color.FromRgb(100, 116, 139))
        };
    }

    private static Brush PremiumChangeBrush(decimal? value)
    {
        return value switch
        {
            >= 10m => new SolidColorBrush(Color.FromRgb(220, 38, 38)),
            >= 5m => new SolidColorBrush(Color.FromRgb(217, 119, 6)),
            <= 0m => new SolidColorBrush(Color.FromRgb(22, 163, 74)),
            _ => new SolidColorBrush(Color.FromRgb(29, 78, 216))
        };
    }

    private static Brush LossRatioRiskBrush(decimal? value)
    {
        return value switch
        {
            >= 100m => new SolidColorBrush(Color.FromRgb(220, 38, 38)),
            >= 85m => new SolidColorBrush(Color.FromRgb(217, 119, 6)),
            null => new SolidColorBrush(Color.FromRgb(100, 116, 139)),
            _ => new SolidColorBrush(Color.FromRgb(22, 163, 74))
        };
    }

    private async Task LoadClientFinancialOverviewAsync(ClientWorkspaceViewModel workspace)
    {
        await LoadClientFinancialTermsAsync(workspace);
        await LoadClientSunLifeLapProjectionAsync(workspace);
    }

    private async Task LoadClientSunLifeLapProjectionAsync(ClientWorkspaceViewModel workspace)
    {
        if (workspace.DatabaseClientId is not > 0)
        {
            return;
        }

        try
        {
            var monthlyClaims = await _databaseService.GetMonthlyClaimsAsync(workspace.DatabaseClientId.Value);
            var assessment = SunLifeLapPolicy.Evaluate(workspace.PrimaryCarrier, monthlyClaims);
            if (assessment is null || !ReferenceEquals(_selectedClientWorkspace, workspace))
            {
                return;
            }

            workspace.ApplySunLifeLapAssessment(assessment);
            UpdateClientWorkspacePoolingThresholdPresentation(workspace);
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
        }
    }

    private void UpdateClientWorkspacePoolingThresholdPresentation(ClientWorkspaceViewModel workspace)
    {
        ClientWorkspacePoolingThresholdTextBlock.Text = workspace.PoolingThresholdText;
        ClientWorkspacePoolingThresholdBasisTextBlock.Text = workspace.PoolingThresholdBasis;
        ClientWorkspacePoolingThresholdBasisTextBlock.ToolTip = workspace.PoolingThresholdBasis;
    }

    private async Task LoadClientFinancialTermsAsync(ClientWorkspaceViewModel workspace)
    {
        if (workspace.DatabaseClientId is not > 0 ||
            string.IsNullOrWhiteSpace(workspace.PrimaryCarrier) ||
            workspace.PrimaryCarrier.Equals("TBD", StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        try
        {
            var terms = await _financialArrangementService.GetCurrentClientTermsAsync(
                workspace.DatabaseClientId.Value,
                workspace.PrimaryCarrier);

            if (!ReferenceEquals(_selectedClientWorkspace, workspace))
            {
                return;
            }

            if (terms is null)
            {
                await LoadClientRenewalExperienceTermsAsync(workspace);
                return;
            }

            workspace.TargetLossRatio = terms.TargetLossRatio ?? workspace.TargetLossRatio;
            workspace.TlrBasis = string.IsNullOrWhiteSpace(terms.TlrBasis) ? workspace.TlrBasis : terms.TlrBasis;
            workspace.PoolingCharge = terms.PoolingRate ?? workspace.PoolingCharge;
            workspace.PoolingThreshold = terms.PoolingThreshold ?? workspace.PoolingThreshold;
            workspace.PoolingThresholdBasis = terms.PoolingSourcePage is int poolingPage
                ? $"Verified source p. {poolingPage}"
                : workspace.PoolingThreshold.HasValue ? "Renewal experience summary" : "No verified terms";

            ClientWorkspaceTargetLossRatioTextBlock.Text = workspace.TargetLossRatioText;
            ClientWorkspaceTlrBasisTextBlock.Text = string.IsNullOrWhiteSpace(workspace.TlrBasis)
                ? workspace.TargetLossRatio.HasValue ? "Renewal experience summary" : "No verified basis"
                : workspace.TlrBasis;
            ClientWorkspacePoolingChargeTextBlock.Text = workspace.PoolingChargeText;
            ClientWorkspacePoolingChargeBasisTextBlock.Text = string.IsNullOrWhiteSpace(terms.PoolingBasis)
                ? workspace.PoolingCharge.HasValue ? "Renewal experience summary" : "No verified terms"
                : terms.PoolingBasis;
            UpdateClientWorkspacePoolingThresholdPresentation(workspace);
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
        }
    }

    private async Task LoadClientRenewalExperienceTermsAsync(ClientWorkspaceViewModel workspace)
    {
        if (workspace.DatabaseClientId is not > 0 || string.IsNullOrWhiteSpace(workspace.PrimaryCarrier))
        {
            return;
        }

        try
        {
            var terms = (await _databaseService.GetRenewalAssumptionsAsync(workspace.DatabaseClientId.Value))
                .Where(item => item.Carrier.Equals(workspace.PrimaryCarrier, StringComparison.OrdinalIgnoreCase))
                .Where(item => item.TargetLossRatio.HasValue || item.PoolingCharge.HasValue || item.PoolingThreshold.HasValue)
                .OrderByDescending(item => item.EffectiveDate)
                .ThenByDescending(item => item.BenefitLine.Equals("All", StringComparison.OrdinalIgnoreCase))
                .FirstOrDefault();

            if (terms is null || !ReferenceEquals(_selectedClientWorkspace, workspace))
            {
                return;
            }

            workspace.TargetLossRatio ??= terms.TargetLossRatio;
            workspace.PoolingCharge ??= terms.PoolingCharge;
            workspace.PoolingThreshold ??= terms.PoolingThreshold;
            workspace.PoolingThresholdBasis = workspace.PoolingThreshold.HasValue
                ? "Renewal experience summary"
                : "No verified terms";

            ClientWorkspaceTargetLossRatioTextBlock.Text = workspace.TargetLossRatioText;
            ClientWorkspaceTlrBasisTextBlock.Text = workspace.TargetLossRatio.HasValue
                ? "Renewal experience summary"
                : "No verified basis";
            ClientWorkspacePoolingChargeTextBlock.Text = workspace.PoolingChargeText;
            ClientWorkspacePoolingChargeBasisTextBlock.Text = workspace.PoolingCharge.HasValue
                ? "Renewal experience summary"
                : "No verified terms";
            UpdateClientWorkspacePoolingThresholdPresentation(workspace);
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
        }
    }

    private async Task LoadClientFlexibleBenefitsOverviewAsync(ClientWorkspaceViewModel workspace)
    {
        if (!workspace.DatabaseClientId.HasValue)
        {
            return;
        }

        try
        {
            var flexibleBenefits = await _flexibleBenefitRepository.GetWorkspaceAsync(workspace.DatabaseClientId.Value);
            if (!ReferenceEquals(_selectedClientWorkspace, workspace) || !flexibleBenefits.HasPlan)
            {
                return;
            }

            ClientWorkspaceHcsaSummaryTextBlock.Text = flexibleBenefits.HcsaSummary;
            ClientWorkspaceLsaSummaryTextBlock.Text = flexibleBenefits.LsaSummary;
            ClientWorkspaceFlexibleAdministratorTextBlock.Text = flexibleBenefits.Administrator;
            ClientWorkspaceFlexibleStatusTextBlock.Text = flexibleBenefits.ReviewStatus;
            ClientWorkspaceFlexibleEffectiveDateTextBlock.Text = $"Effective {flexibleBenefits.EffectiveDate}";
            ClientWorkspaceFlexibleBenefitsCard.Visibility = Visibility.Visible;
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            ClientWorkspaceFlexibleBenefitsCard.Visibility = Visibility.Collapsed;
        }
    }

    private RenewalQueueItem? FindRenewalForClient(ClientRenewalSelectionIdentity client)
    {
        var selected = ClientRenewalSelectionResolver.Select(
            client,
            _renewalQueue.Select(CreateRenewalSelectionCandidate));
        return selected is null ? null : _renewalQueue[selected.Index];
    }

    private bool SelectRenewalForCurrentClient()
    {
        if (_selectedClientWorkspace is null)
        {
            return false;
        }

        _clientRenewalSelectionIdentity = CreateClientRenewalSelectionIdentity(_selectedClientWorkspace);
        var row = FindRenewalForClient(_clientRenewalSelectionIdentity);
        if (row is null)
        {
            RenewalQueueGrid.SelectedItem = null;
            UpdateSelectedRenewalDetail(null);
            return false;
        }

        RenewalQueueGrid.SelectedItem = row;
        RenewalQueueGrid.ScrollIntoView(row);
        UpdateSelectedRenewalDetail(row);
        return true;
    }

    private static ClientRenewalSelectionIdentity CreateClientRenewalSelectionIdentity(ClientListItem client)
    {
        return new ClientRenewalSelectionIdentity(
            client.DatabaseClientId,
            client.ClientId,
            client.ClientName,
            client.PrimaryCarrier,
            null);
    }

    private static ClientRenewalSelectionIdentity CreateClientRenewalSelectionIdentity(ClientWorkspaceViewModel client)
    {
        return new ClientRenewalSelectionIdentity(
            client.DatabaseClientId,
            client.ClientId,
            client.ClientName,
            client.PrimaryCarrier,
            client.LastRenewalDate);
    }

    private static ClientRenewalSelectionCandidate CreateRenewalSelectionCandidate(
        RenewalQueueItem row,
        int index)
    {
        return new ClientRenewalSelectionCandidate(
            index,
            new RenewalWorkflowIdentity(
                row.ClientId,
                row.SourceDocumentId,
                row.Client,
                row.PolicyNumber,
                row.Carrier,
                row.RenewalDate),
            row.DateLoaded);
    }

    private void ShowClientWorkspaceOverview()
    {
        HideClientWorkspaceContentPanels();
        ClientWorkspaceOverviewPanel.Visibility = Visibility.Visible;
    }

    private void HideClientWorkspaceContentPanels()
    {
        ClientWorkspaceOverviewPanel.Visibility = Visibility.Collapsed;
        ClientPlanDesignWorkspaceView.Visibility = Visibility.Collapsed;
        ClientClaimsExperienceDashboard.Visibility = Visibility.Collapsed;
        ClientAsoDentalDashboard.Visibility = Visibility.Collapsed;
        ClientBillingDashboard.Visibility = Visibility.Collapsed;
        ClientFinancialIntelligenceDashboard.Visibility = Visibility.Collapsed;
        ClientHistoryWorkspace.Visibility = Visibility.Collapsed;
        ClientWorkspacePlaceholderPanel.Visibility = Visibility.Collapsed;
    }

    private void ShowClientWorkspacePlaceholder(string title, string message)
    {
        HideClientWorkspaceContentPanels();
        ClientWorkspacePlaceholderPanel.Visibility = Visibility.Visible;
        ClientWorkspacePlaceholderTitleTextBlock.Text = title;
        ClientWorkspacePlaceholderMessageTextBlock.Text = message;
        ClientsScrollViewer.ScrollToTop();
    }

    private void ClientWorkspaceOverviewTab_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsureClientWorkspaceSelected())
        {
            return;
        }

        SetNavigationWorkspace("ClientWorkspace", "Overview", _selectedClientWorkspace?.ClientName);
        ShowClientWorkspaceOverview();
        SetStatus($"Opened {_selectedClientWorkspace?.ClientName ?? "client"} overview.");
    }

    private async void OpenRenewalWorkspace_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsureClientWorkspaceSelected())
        {
            return;
        }

        ShowRenewalsDashboard();
        _clientRenewalSelectionIdentity = CreateClientRenewalSelectionIdentity(_selectedClientWorkspace!);
        SetNavigationWorkspace("ClientWorkspace", "Renewal", _selectedClientWorkspace?.ClientName);
        await LoadSavedRenewalQueueAsync();
        if (!SelectRenewalForCurrentClient())
        {
            SetStatus($"No saved renewal matches {_selectedClientWorkspace?.ClientName ?? "the selected client"}. Validate that client's renewal before opening negotiation.");
            return;
        }

        SetStatus($"Opened Renewal workspace for {_selectedClientWorkspace?.ClientName ?? "selected client"}.");
    }

    private async void OpenPlanDesign_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsureClientWorkspaceSelected())
        {
            return;
        }

        try
        {
            await OpenClientPlanDesignSubpageAsync();
            SetStatus($"Opened Plan Design workspace for {_selectedClientWorkspace?.ClientName ?? "selected client"}.");
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            SetStatus("Plan Design workspace could not load for the selected client.");
            System.Windows.MessageBox.Show(
                this,
                "The Plan Design workspace could not open for this client. The client record is still selected, but the saved plan design data could not be loaded.",
                "Plan Design Workspace",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }
    }

    private async void OpenClaimsWorkspace_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsureClientWorkspaceSelected())
        {
            return;
        }

        SetNavigationWorkspace("ClientWorkspace", "Claims", _selectedClientWorkspace?.ClientName);
        await ShowClientClaimsWorkspaceAsync();
    }

    private async void OpenClientBillingWorkspace_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsureClientWorkspaceSelected())
        {
            return;
        }

        SetNavigationWorkspace("ClientWorkspace", "Billing", _selectedClientWorkspace?.ClientName);
        await ShowClientBillingWorkspaceAsync();
    }

    private async void OpenClientAsoWorkspace_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsureClientWorkspaceSelected())
        {
            return;
        }

        SetNavigationWorkspace("ClientWorkspace", "ASO", _selectedClientWorkspace?.ClientName);
        await ShowClientAsoWorkspaceAsync();
    }

    private async Task ShowClientClaimsWorkspaceAsync()
    {
        var workspace = _selectedClientWorkspace;
        if (workspace is null)
        {
            ShowClientsPage();
            return;
        }

        HideClientWorkspaceContentPanels();
        ClientClaimsExperienceDashboard.Visibility = Visibility.Visible;
        ClientsScrollViewer.ScrollToTop();

        if (workspace.DatabaseClientId is not > 0)
        {
            ClientClaimsExperienceDashboard.LoadEmpty(workspace);
            SetStatus($"Opened Claims Experience for {workspace.ClientName}; no SQLite client id is linked yet.");
            return;
        }

        try
        {
            SetStatus($"Loading Claims Experience for {workspace.ClientName}...");
            var dashboardDataTask = _clientClaimsDashboardDataService.GetAsync(workspace.DatabaseClientId.Value);
            var assumptionsTask = _databaseService.GetRenewalAssumptionsAsync(workspace.DatabaseClientId.Value);
            var asoFinancialTask = _databaseService.GetLatestAsoFinancialActivityAsync(workspace.DatabaseClientId.Value);
            await Task.WhenAll(dashboardDataTask, assumptionsTask, asoFinancialTask);

            var renewalTerms = ClaimsExperienceRenewalTerms.FromRenewalAssumptions(
                assumptionsTask.Result,
                workspace.PrimaryCarrier);
            var dashboardData = dashboardDataTask.Result;
            ClientClaimsExperienceDashboard.LoadDashboard(
                workspace,
                dashboardData.Claims,
                dashboardData.MonthlyClaims,
                renewalTerms,
                asoFinancialTask.Result,
                dashboardData);
            SetStatus($"Opened Claims Experience dashboard for {workspace.ClientName}. {dashboardData.StatusMessage}");
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            ClientClaimsExperienceDashboard.LoadEmpty(workspace);
            SetStatus($"Claims Experience could not load for {workspace.ClientName}; showing the empty dashboard state.");
        }
    }

    private async Task ShowClientAsoWorkspaceAsync()
    {
        var workspace = _selectedClientWorkspace;
        if (workspace is null)
        {
            ShowClientsPage();
            return;
        }

        HideClientWorkspaceContentPanels();
        ClientAsoDentalDashboard.Visibility = Visibility.Visible;
        ClientsScrollViewer.ScrollToTop();

        if (workspace.DatabaseClientId is not > 0)
        {
            ClientAsoDentalDashboard.LoadEmpty(workspace);
            SetStatus($"Opened ASO Dental Cost for {workspace.ClientName}; no SQLite client id is linked yet.");
            return;
        }

        try
        {
            SetStatus($"Loading ASO Dental Cost for {workspace.ClientName}...");
            var monthlyClaimsTask = _databaseService.GetMonthlyClaimsAsync(workspace.DatabaseClientId.Value);
            var asoStatementsTask = _databaseService.GetAsoFinancialActivityHistoryAsync(workspace.DatabaseClientId.Value);
            await Task.WhenAll(monthlyClaimsTask, asoStatementsTask);
            ClientAsoDentalDashboard.LoadDashboard(workspace, monthlyClaimsTask.Result, asoStatementsTask.Result);
            SetStatus($"Opened ASO Dental Cost for {workspace.ClientName}.");
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            ClientAsoDentalDashboard.LoadEmpty(workspace);
            SetStatus($"ASO Dental Cost could not load for {workspace.ClientName}; showing the empty dashboard state.");
        }
    }

    private async void ClientClaimsExperienceDashboard_UploadClaimsExperienceRequested(object? sender, EventArgs e)
    {
        var workspace = _selectedClientWorkspace;
        if (workspace is null)
        {
            ShowClientsPage();
            return;
        }

        var folderTarget = new System.Windows.Controls.TextBox
        {
            Text = Directory.Exists(CurrentYearFolderTextBox.Text)
                ? CurrentYearFolderTextBox.Text
                : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        };

        if (!BrowseFolder(folderTarget, "Choose the parent folder containing Current, Prior, and PriorPrior"))
        {
            return;
        }

        var sourceFolder = folderTarget.Text.Trim();
        var sourceFiles = EnumerateClaimsSourceFiles(sourceFolder).ToList();
        var renewalExperiencePath = sourceFiles
            .FirstOrDefault(path => IsPdfFile(path) &&
                Path.GetFileName(path).Contains("renewal experience", StringComparison.OrdinalIgnoreCase));

        var livesPremiumClaimsPath = sourceFiles
            .Where(path => IsPdfFile(path))
            .Where(path =>
            {
                var normalizedName = Path.GetFileNameWithoutExtension(path)
                    .Replace(" ", string.Empty, StringComparison.Ordinal)
                    .Replace("_", string.Empty, StringComparison.Ordinal)
                    .Replace("-", string.Empty, StringComparison.Ordinal);
                return normalizedName.Contains("livespremiumclaims", StringComparison.OrdinalIgnoreCase) ||
                       normalizedName.Contains("premiumclaims", StringComparison.OrdinalIgnoreCase) ||
                       normalizedName.Contains("premclaims", StringComparison.OrdinalIgnoreCase) ||
                       normalizedName.Contains("insuredexperience", StringComparison.OrdinalIgnoreCase);
            })
            .OrderByDescending(path => Path.GetFileName(path).Contains("12 mos", StringComparison.OrdinalIgnoreCase) ||
                                       Path.GetFileName(path).Contains("12 month", StringComparison.OrdinalIgnoreCase))
            .ThenBy(path => Path.GetFileName(path), StringComparer.OrdinalIgnoreCase)
            .FirstOrDefault();
        if (!string.IsNullOrWhiteSpace(livesPremiumClaimsPath))
        {
            await ImportFlatClaimsExperienceAsync(workspace, sourceFolder, livesPremiumClaimsPath, sourceFiles);
            return;
        }
        if (!string.IsNullOrWhiteSpace(renewalExperiencePath))
        {
            await ImportRenewalExperienceClaimsAsync(workspace, sourceFolder, renewalExperiencePath);
            return;
        }
        var currentFolder = FindClaimsPeriodFolder(sourceFolder, "Current");
        var priorFolder = FindClaimsPeriodFolder(sourceFolder, "Prior");
        var priorPriorFolder = FindClaimsPeriodFolder(
            sourceFolder,
            "PriorPrior",
            "Prior Prior",
            "Prior_Prior",
            "Prior-Prior",
            "PriorPrior12",
            "Prior Prior 12",
            "Prior_Prior_12",
            "Prior-Prior-12");

        var missingFolders = new List<string>();
        if (currentFolder is null)
        {
            missingFolders.Add("Current");
        }

        if (priorFolder is null)
        {
            missingFolders.Add("Prior");
        }

        if (priorPriorFolder is null)
        {
            missingFolders.Add("PriorPrior");
        }

        if (missingFolders.Count > 0)
        {
            System.Windows.MessageBox.Show(
                this,
                $"The selected folder is missing: {string.Join(", ", missingFolders)}.{Environment.NewLine}{Environment.NewLine}" +
                "Choose the parent folder with this structure:" + Environment.NewLine +
                "Current" + Environment.NewLine +
                "Prior" + Environment.NewLine +
                "PriorPrior",
                "Claims Experience Folder",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        var currentFiles = EnumerateClaimsSourceFiles(currentFolder!).ToList();
        if (currentFiles.Count == 0)
        {
            System.Windows.MessageBox.Show(
                this,
                "The Current folder does not contain a supported claims PDF or Excel file.",
                "Claims Experience Folder",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        try
        {
            SetBusy(true, $"Extracting claims experience for {workspace.ClientName}...");
            CurrentYearFolderTextBox.Text = sourceFolder;
            PriorYearFolderTextBox.Text = priorFolder!;
            SaveSettings();

            var request = new ClaimsExperienceDataModelRequest
            {
                SourceFolder = sourceFolder,
                OutputFolder = OutputFolderTextBox.Text.Trim(),
                ReportMonthCount = SelectedReportMonthCount(),
                RenewalTerms = await LoadClaimsRenewalTermsAsync(
                    workspace.DatabaseClientId,
                    workspace.PrimaryCarrier)
            };

            _claimsDataModelExport = await _claimsDataModelBuilder.ExportAsync(request);

            // The selected Client Workspace is the source of truth for ownership.
            ValidateClaimsSourceIdentityForWorkspace(workspace, _claimsDataModelExport);
            _claimsDataModelExport.ClientName = workspace.ClientName;
            _claimsDataModelExport.ClientNumber = workspace.ClientId;
            if (string.IsNullOrWhiteSpace(_claimsDataModelExport.Carrier))
            {
                _claimsDataModelExport.Carrier = workspace.PrimaryCarrier;
            }

            var allSourceFiles = EnumerateClaimsSourceFiles(currentFolder!)
                .Concat(EnumerateClaimsSourceFiles(priorFolder!))
                .Concat(EnumerateClaimsSourceFiles(priorPriorFolder!))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            var databaseSave = await _databaseService.SaveClaimsExperienceAsync(
                _claimsDataModelExport,
                allSourceFiles,
                sourceFolder);

            if (!databaseSave.Succeeded)
            {
                throw new InvalidOperationException(databaseSave.Message);
            }

            if (databaseSave.ClientId is > 0)
            {
                workspace.DatabaseClientId = databaseSave.ClientId;
                _latestSavedClientId = databaseSave.ClientId;
            }

            OutputFolderTextBox.Text = _claimsDataModelExport.OutputFolder;
            ExportTextBlock.Text =
                $"Claims data model: {_claimsDataModelExport.WorkbookPath}{Environment.NewLine}" +
                $"Current files: {currentFiles.Count}{Environment.NewLine}" +
                $"Prior files: {EnumerateClaimsSourceFiles(priorFolder!).Count()}{Environment.NewLine}" +
                $"Prior-prior files: {EnumerateClaimsSourceFiles(priorPriorFolder!).Count()}{Environment.NewLine}" +
                $"Database: {databaseSave.Message}";

            SetOpenOutputButtonsEnabled(Directory.Exists(_claimsDataModelExport.OutputFolder));
            await ShowClientClaimsWorkspaceAsync();
            SetStatus($"Claims experience extracted and saved for {workspace.ClientName}.");

            System.Windows.MessageBox.Show(
                this,
                $"Claims experience was extracted from Current, Prior, and PriorPrior and saved to SQLite for {workspace.ClientName}.",
                "Claims Experience Complete",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            System.Windows.MessageBox.Show(
                this,
                ex.Message,
                "Claims Experience Extraction Failed",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            SetStatus($"Claims experience extraction failed for {workspace.ClientName}.");
        }
        finally
        {
            SetBusy(false);
        }
    }

    private async Task ImportFlatClaimsExperienceAsync(
        ClientWorkspaceViewModel workspace,
        string sourceFolder,
        string livesPremiumClaimsPath,
        IReadOnlyCollection<string> sourceFiles)
    {
        try
        {
            SetBusy(true, $"Importing claims experience for {workspace.ClientName}...");
            CurrentYearFolderTextBox.Text = sourceFolder;
            SaveSettings();

            var export = await _claimsDataModelBuilder.ExportAsync(new ClaimsExperienceDataModelRequest
            {
                SourceFolder = "",
                SourceFiles = sourceFiles.ToList(),
                OutputFolder = OutputFolderTextBox.Text.Trim(),
                ReportMonthCount = SelectedReportMonthCount(),
                RenewalTerms = await LoadClaimsRenewalTermsAsync(
                    workspace.DatabaseClientId,
                    workspace.PrimaryCarrier)
            });

            ValidateClaimsSourceIdentityForWorkspace(workspace, export);
            // The selected workspace is authoritative for client ownership.
            export.ClientName = workspace.ClientName;
            export.ClientNumber = workspace.ClientId;
            if (string.IsNullOrWhiteSpace(export.Carrier))
            {
                export.Carrier = workspace.PrimaryCarrier;
            }

            var databaseSave = await _databaseService.SaveClaimsExperienceAsync(export, sourceFiles, sourceFolder);
            if (!databaseSave.Succeeded)
            {
                throw new InvalidOperationException(databaseSave.Message);
            }

            _claimsDataModelExport = export;
            if (databaseSave.ClientId is > 0)
            {
                workspace.DatabaseClientId = databaseSave.ClientId;
                _latestSavedClientId = databaseSave.ClientId;
            }

            OutputFolderTextBox.Text = export.OutputFolder;
            ExportTextBlock.Text =
                $"Claims data model: {export.WorkbookPath}{Environment.NewLine}" +
                $"Source files: {sourceFiles.Count}{Environment.NewLine}" +
                $"Database: {databaseSave.Message}";
            SetOpenOutputButtonsEnabled(Directory.Exists(export.OutputFolder));
            await ShowClientClaimsWorkspaceAsync();
            SetStatus($"Claims experience imported and saved for {workspace.ClientName}.");
            System.Windows.MessageBox.Show(
                this,
                "Saved the Lives/Premium/Claims report from the selected folder. " +
                "If a complete renewal-experience report is available, upload that report instead because it takes precedence over interim periods.",
                "Claims Experience Complete",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            System.Windows.MessageBox.Show(
                this,
                ex.Message,
                "Claims Experience Extraction Failed",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            SetStatus($"Claims experience extraction failed for {workspace.ClientName}.");
        }
        finally
        {
            SetBusy(false);
        }
    }
    private async Task ImportRenewalExperienceClaimsAsync(
        ClientWorkspaceViewModel workspace,
        string sourceFolder,
        string renewalExperiencePath)
    {
        try
        {
            SetBusy(true, $"Importing renewal experience for {workspace.ClientName}...");
            CurrentYearFolderTextBox.Text = sourceFolder;
            SaveSettings();

            var supportingPdfPaths = EnumerateClaimsSourceFiles(sourceFolder)
                .Where(path => IsPdfFile(path) &&
                    !path.Equals(renewalExperiencePath, StringComparison.OrdinalIgnoreCase))
                .Where(path =>
                {
                    var name = Path.GetFileName(path);
                    return name.Contains("12 mos", StringComparison.OrdinalIgnoreCase) &&
                           (name.Contains("dental", StringComparison.OrdinalIgnoreCase) ||
                            name.Contains("ehc", StringComparison.OrdinalIgnoreCase));
                })
                .ToList();

            var analysis = await _builder.AnalyzeAsync(new RenewalCase
            {
                ClientName = workspace.ClientName,
                PolicyNumber = workspace.ClientId,
                PdfPath = renewalExperiencePath,
                AdditionalPdfPaths = supportingPdfPaths,
                OutputFolder = OutputFolderTextBox.Text.Trim()
            }, new Progress<string>(SetStatus));

            // The selected workspace owns the import even if the report has an older display name.
            analysis.RenewalCase.ClientName = workspace.ClientName;
            analysis.RenewalCase.PolicyNumber = workspace.ClientId;
            var databaseSave = await _databaseService.SaveRenewalAnalysisAsync(
                analysis,
                saveRenewalVersions: false);
            if (!databaseSave.Succeeded)
            {
                throw new InvalidOperationException(databaseSave.Message);
            }

            if (databaseSave.ClientId is > 0)
            {
                workspace.DatabaseClientId = databaseSave.ClientId;
                _latestSavedClientId = databaseSave.ClientId;
                _latestSavedDocumentId = databaseSave.DocumentId;
            }

            await ShowClientClaimsWorkspaceAsync();
            SetStatus($"Renewal experience imported and saved for {workspace.ClientName}.");
            System.Windows.MessageBox.Show(
                this,
                $"Saved the complete renewal-experience periods from {Path.GetFileName(renewalExperiencePath)}. " +
                "The dashboard uses those full periods ahead of shorter interim claims reports.",
                "Claims Experience Complete",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            System.Windows.MessageBox.Show(
                this,
                ex.Message,
                "Claims Experience Extraction Failed",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            SetStatus($"Claims experience extraction failed for {workspace.ClientName}.");
        }
        finally
        {
            SetBusy(false);
        }
    }
    private static string? FindClaimsPeriodFolder(string parentFolder, params string[] candidateNames)
    {
        if (!Directory.Exists(parentFolder))
        {
            return null;
        }

        var folders = Directory.EnumerateDirectories(parentFolder, "*", SearchOption.TopDirectoryOnly)
            .ToList();
        return folders.FirstOrDefault(folder => candidateNames.Any(candidate =>
            Path.GetFileName(folder).Equals(candidate, StringComparison.OrdinalIgnoreCase)));
    }

    private static IEnumerable<string> EnumerateClaimsSourceFiles(string folder)
    {
        if (!Directory.Exists(folder))
        {
            return [];
        }

        return Directory.EnumerateFiles(folder, "*.*", SearchOption.TopDirectoryOnly)
            .Where(path => IsPdfFile(path) || IsWorkbookFile(path));
    }

    private async Task ShowClientBillingWorkspaceAsync()
    {
        var workspace = _selectedClientWorkspace;
        if (workspace is null)
        {
            ShowClientsPage();
            return;
        }

        HideClientWorkspaceContentPanels();
        ClientBillingDashboard.Visibility = Visibility.Visible;
        ClientsScrollViewer.ScrollToTop();

        try
        {
            SetStatus($"Loading Billing History for {workspace.ClientName}...");
            if (workspace.DatabaseClientId is not > 0)
            {
                ClientBillingDashboard.LoadEmpty(workspace);
                SetStatus($"Opened Billing History for {workspace.ClientName}; no SQLite client id is linked yet.");
                return;
            }

            await ClientBillingDashboard.LoadDashboardAsync(workspace);
            SetStatus($"Opened Billing History for {workspace.ClientName}.");
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            ClientBillingDashboard.LoadEmpty(workspace);
            SetStatus($"Billing History could not load for {workspace.ClientName}; showing the empty dashboard state.");
        }
    }

    private async void OpenClientFinancialIntelligenceWorkspace_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsureClientWorkspaceSelected())
        {
            return;
        }

        SetNavigationWorkspace("ClientWorkspace", "Financial Intelligence", _selectedClientWorkspace?.ClientName);
        await ShowClientFinancialIntelligenceWorkspaceAsync();
    }

    private async Task ShowClientFinancialIntelligenceWorkspaceAsync()
    {
        var workspace = _selectedClientWorkspace;
        if (workspace is null)
        {
            ShowClientsPage();
            return;
        }

        HideClientWorkspaceContentPanels();
        ClientFinancialIntelligenceDashboard.Visibility = Visibility.Visible;
        ClientsScrollViewer.ScrollToTop();

        try
        {
            SetStatus($"Loading Financial Intelligence for {workspace.ClientName}...");
            if (workspace.DatabaseClientId is not > 0)
            {
                ClientFinancialIntelligenceDashboard.LoadEmpty(workspace);
                SetStatus($"Opened Financial Intelligence for {workspace.ClientName}; no SQLite client id is linked yet.");
                return;
            }

            await ClientFinancialIntelligenceDashboard.LoadDashboardAsync(workspace);
            SetStatus($"Opened Financial Intelligence for {workspace.ClientName}.");
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            ClientFinancialIntelligenceDashboard.LoadEmpty(workspace);
            SetStatus($"Financial Intelligence could not load for {workspace.ClientName}; showing the empty dashboard state.");
        }
    }

    private async void CreateProposal_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsureClientWorkspaceSelected())
        {
            return;
        }

        SetNavigationWorkspace("ClientWorkspace", "Proposals", _selectedClientWorkspace?.ClientName);
        await ShowProposalComparisonWorkspaceAsync();
        SetStatus($"Opened Proposal workspace for {_selectedClientWorkspace?.ClientName ?? "selected client"}.");
    }

    private async void ClientWorkspaceReportsTab_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsureClientWorkspaceSelected())
        {
            return;
        }

        await ShowReportsAnalyticsAsync();
        SetNavigationWorkspace("ClientWorkspace", "Reports", _selectedClientWorkspace?.ClientName);
        SetStatus($"Opened Reports for {_selectedClientWorkspace?.ClientName ?? "selected client"}.");
    }

    private void ClientWorkspaceDocumentsTab_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsureClientWorkspaceSelected())
        {
            return;
        }

        SetNavigationWorkspace("ClientWorkspace", "Documents", _selectedClientWorkspace?.ClientName);
        var clientName = _selectedClientWorkspace?.ClientName ?? "selected client";
        ShowClientWorkspacePlaceholder(
            "Documents & Extraction History",
            $"Client-filtered document history for {clientName}. This view is for reviewing documents already linked to the client. Global intake, batch discovery, extraction, validation, and saving remain in Document Intelligence at /document-intelligence.");
        SetStatus($"Opened client-filtered Documents view for {clientName}.");
    }

    private void ClientWorkspaceHistoryTab_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsureClientWorkspaceSelected())
        {
            return;
        }

        SetNavigationWorkspace("ClientWorkspace", "History", _selectedClientWorkspace?.ClientName);
        if (!_selectedClientWorkspace!.DatabaseClientId.HasValue)
        {
            System.Windows.MessageBox.Show(this, "Select a saved client before importing history.", "Client History", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        HideClientWorkspaceContentPanels();
        ClientHistoryWorkspace.Visibility = Visibility.Visible;
        _ = ClientHistoryWorkspace.LoadAsync(_selectedClientWorkspace.DatabaseClientId.Value);
        ClientsScrollViewer.ScrollToTop();
        SetStatus("Opened Client History layout.");
    }

    private bool EnsureClientWorkspaceSelected()
    {
        if (_selectedClientWorkspace is not null)
        {
            return true;
        }

        ShowClientsPage();
        SetStatus("Select a client first, then open the client workspace section.");
        return false;
    }

    private void UploadDocument_Click(object sender, RoutedEventArgs e)
    {
        OpenDocumentIntelligence_Click(sender, e);
        SetStatus("Opened global Document Intelligence for document intake. Client Documents remains a filtered history view.");
    }

    private void AddClientNote_Click(object sender, RoutedEventArgs e)
    {
        SetStatus($"Client note placeholder ready for {_selectedClientWorkspace?.ClientName ?? "selected client"}.");
    }

    private void OpenClientActions_Click(object sender, RoutedEventArgs e)
    {
        SetStatus($"Client Actions placeholder ready for {_selectedClientWorkspace?.ClientName ?? "selected client"}.");
    }

    private void NewClientWorkspace_Click(object sender, RoutedEventArgs e)
    {
        SetStatus("New Workspace placeholder ready. This can create a client-specific renewal, plan design, claims, or proposal workspace next.");
    }

    private async void StartTask_Click(object sender, RoutedEventArgs e)
    {
        var taskId = (sender as FrameworkElement)?.Tag as string ?? "";
        if (taskId.Equals("underwriting-analysis", StringComparison.OrdinalIgnoreCase))
        {
            SelectRenewalForCurrentClient();
            var row = SelectedRenewalQueueRow();
            await ShowUnderwritingDashboardAsync(row);
            return;
        }

        SetStatus($"Started task: {taskId.Replace('-', ' ')}.");
    }

    private void BackToClientsList_Click(object sender, RoutedEventArgs e)
    {
        SetNavigationWorkspace("Clients", "Master Client List");
        ClientWorkspacePanel.Visibility = Visibility.Collapsed;
        ClientListPanel.Visibility = Visibility.Visible;
        ClientsScrollViewer.ScrollToTop();
        SetStatus("Returned to Clients master list.");
    }

    private void ImportClients_Click(object sender, RoutedEventArgs e)
    {
        SetStatus("Import Clients placeholder ready. This can be connected to an Excel import next.");
    }

    private void ExportClientList_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var visibleRows = _clientListView?.Cast<ClientListItem>().ToList() ?? _clientList.ToList();
            var outputFolder = Path.Combine(ResolveOutputRootFolder(), "Client List Reports");
            Directory.CreateDirectory(outputFolder);
            var workbookPath = Path.Combine(outputFolder, $"Client_List_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");

            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Clients");
            var headers = new[]
            {
                "Client Name", "Client ID", "Status", "Industry", "Employees", "Renewal Date",
                "Primary Carrier", "Plan Type", "Health Score", "Last Updated"
            };

            for (var column = 0; column < headers.Length; column++)
            {
                worksheet.Cell(1, column + 1).Value = headers[column];
            }

            var rowNumber = 2;
            foreach (var row in visibleRows)
            {
                worksheet.Cell(rowNumber, 1).Value = row.ClientName;
                worksheet.Cell(rowNumber, 2).Value = row.ClientId;
                worksheet.Cell(rowNumber, 3).Value = row.Status;
                worksheet.Cell(rowNumber, 4).Value = row.Industry;
                worksheet.Cell(rowNumber, 5).Value = row.Employees;
                worksheet.Cell(rowNumber, 6).Value = row.RenewalDate;
                worksheet.Cell(rowNumber, 7).Value = row.PrimaryCarrier;
                worksheet.Cell(rowNumber, 8).Value = row.PlanType;
                worksheet.Cell(rowNumber, 9).Value = row.HealthScoreText;
                worksheet.Cell(rowNumber, 10).Value = row.LastUpdated;
                rowNumber++;
            }

            var range = worksheet.Range(1, 1, Math.Max(1, rowNumber - 1), headers.Length);
            range.CreateTable("Clients");
            worksheet.Row(1).Style.Font.Bold = true;
            worksheet.Row(1).Style.Fill.BackgroundColor = XLColor.FromHtml("#061B3A");
            worksheet.Row(1).Style.Font.FontColor = XLColor.White;
            worksheet.Columns().AdjustToContents();
            workbook.SaveAs(workbookPath);

            Process.Start(new ProcessStartInfo
            {
                FileName = outputFolder,
                UseShellExecute = true
            });
            SetStatus($"Client list exported to Excel with {visibleRows.Count} visible row(s).");
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            System.Windows.MessageBox.Show(this, ex.Message, "Export Client List failed", MessageBoxButton.OK, MessageBoxImage.Warning);
            SetStatus("Client list export failed.");
        }
    }

    private void CreateNewClient_Click(object sender, RoutedEventArgs e)
    {
        SetStatus("New Client placeholder ready. This can open a saved client intake form next.");
    }

    private void OpenDocumentLibrary_Click(object sender, RoutedEventArgs e)
    {
        SetNavigationWorkspace("DocumentOperations", "Document Queue");
        ShowPlaceholder(
            "Document Library",
            "This will become the searchable library for saved renewal packages, claims reports, benefit booklets, proposals, and generated outputs.");
    }

    private void OpenTasksWorkflows_Click(object sender, RoutedEventArgs e)
    {
        SetNavigationWorkspace("DocumentOperations", "Validation Queue");
        ShowPlaceholder(
            "Tasks & Workflows",
            "This will become the task queue for missing documents, validation work, consultant review, underwriting review, and client-ready presentation follow-up.");
    }

    private void OpenDataManagement_Click(object sender, RoutedEventArgs e)
    {
        SetNavigationWorkspace("Administration", "Data Management");
        ShowPlaceholder(
            "Data Management",
            "This is for maintaining the Master Benefits Database: imports, duplicate checks, data quality review, history cleanup, and database refresh tools.");
    }

    private void OpenAdministration_Click(object sender, RoutedEventArgs e)
    {
        SetNavigationWorkspace("Administration", "Administration");
        ShowPlaceholder(
            "Administration",
            "This is for setup items such as users, roles, offices, consultants, carrier settings, assumptions, and future permissions.");
    }

    private void OpenDataSources_Click(object sender, RoutedEventArgs e)
    {
        ShowDataSources();
    }

    private void ShowDataSources()
    {
        SetNavigationWorkspace("Administration", "Data Sources");
        HomeDashboardScrollViewer.Visibility = Visibility.Collapsed;
        DashboardScrollViewer.Visibility = Visibility.Collapsed;
        UnderwritingWorkbenchScrollViewer.Visibility = Visibility.Collapsed;
        AssumptionLibraryScrollViewer.Visibility = Visibility.Collapsed;
        ReportsAnalyticsScrollViewer.Visibility = Visibility.Collapsed;
        BenchmarkingDashboardScrollViewer.Visibility = Visibility.Collapsed;
        ProposalComparisonScrollViewer.Visibility = Visibility.Collapsed;
        PlanDesignIntelligenceScrollViewer.Visibility = Visibility.Collapsed;
        PlanDesignFactorsScrollViewer.Visibility = Visibility.Collapsed;
        DocumentIntelligenceScrollViewer.Visibility = Visibility.Collapsed;
        ClientsScrollViewer.Visibility = Visibility.Collapsed;
        NegotiationWorkbenchScrollViewer.Visibility = Visibility.Collapsed;
        DataSourcesScrollViewer.Visibility = Visibility.Visible;
        DataSourcesScrollViewer.ScrollToTop();
        SetOperationalFooterVisible(true);
        SetStatus("Opened Data Sources.");
    }

    private async void UnderwritingClientComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_isUpdatingUnderwritingSelection || UnderwritingWorkbenchScrollViewer.Visibility != Visibility.Visible)
        {
            return;
        }

        await LoadUnderwritingDashboardAsync(UnderwritingClientComboBox.SelectedItem as RenewalQueueItem);
    }

    private async void RefreshUnderwriting_Click(object sender, RoutedEventArgs e)
    {
        await LoadUnderwritingDashboardAsync(_selectedUnderwritingRenewal ?? UnderwritingClientComboBox.SelectedItem as RenewalQueueItem);
    }

    private void UnderwritingNewScenario_Click(object sender, RoutedEventArgs e)
    {
        SetStatus("New underwriting scenario: scenario editing belongs in Underwriting Intelligence; use negotiation only for final rates.");
        System.Windows.MessageBox.Show(
            this,
            "Scenario editing will stay in Underwriting Intelligence so the consultant model remains separate from negotiated final rates.",
            "Underwriting Scenario",
            MessageBoxButton.OK,
            MessageBoxImage.Information);
    }

    private void ExportUnderwritingMemo_Click(object sender, RoutedEventArgs e)
    {
        SetStatus("Underwriting memo: use the existing renewal workbook / PowerPoint export after selecting the modelled or negotiated version.");
        System.Windows.MessageBox.Show(
            this,
            "The underwriting memo will be produced through the existing export workflow so the workbook and PowerPoint stay tied to the selected renewal version.",
            "Underwriting Memo",
            MessageBoxButton.OK,
            MessageBoxImage.Information);
    }

    private async void UnderwritingOpenNegotiation_Click(object sender, RoutedEventArgs e)
    {
        var row = _selectedUnderwritingRenewal ?? UnderwritingClientComboBox.SelectedItem as RenewalQueueItem;
        if (row is null)
        {
            SetStatus("Select a renewal before opening negotiation.");
            return;
        }

        await OpenNegotiationWorkbenchAsync(row);
    }

    private async void OpenCarrierIntelligence_Click(object sender, RoutedEventArgs e)
    {
        SetNavigationWorkspace("BookOfBusiness", "Carrier Performance");
        await ShowAssumptionLibraryAsync();
    }

    private async void OpenPlanDesignIntelligence_Click(object sender, RoutedEventArgs e)
    {
        if (!EnsureClientWorkspaceSelected())
        {
            return;
        }

        try
        {
            await OpenClientPlanDesignSubpageAsync();
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            SetStatus("Plan Design workspace could not load.");
            System.Windows.MessageBox.Show(
                this,
                "The Plan Design workspace could not open. The app is still running, but the saved plan design data could not be loaded.",
                "Plan Design Workspace",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }
    }

    internal async Task OpenClientPlanDesignSubpageAsync(string? subpage = null)
    {
        if (!EnsureClientWorkspaceSelected())
        {
            return;
        }

        SetNavigationWorkspace("ClientWorkspace", "Plan Design", _selectedClientWorkspace?.ClientName);
        await ShowPlanDesignIntelligenceAsync(subpage);
        SetNavigationWorkspace("ClientWorkspace", "Plan Design", _selectedClientWorkspace?.ClientName);
    }

    private async void RefreshPlanDesignIntelligence_Click(object sender, RoutedEventArgs e)
    {
        await LoadPlanDesignIntelligenceAsync();
    }

    private async void PlanDesignIntelligenceClient_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_isLoadingPlanDesignIntelligence || !IsLoaded || PlanDesignIntelligenceScrollViewer.Visibility != Visibility.Visible)
        {
            return;
        }

        await LoadPlanDesignIntelligenceAsync();
    }

    private async void OpenPlanDesignFactors_Click(object sender, RoutedEventArgs e)
    {
        SetNavigationWorkspace("Administration", "Plan Design Factors");
        await ShowPlanDesignFactorsAsync();
    }

    private async void RefreshPlanDesignFactors_Click(object sender, RoutedEventArgs e)
    {
        await LoadPlanDesignFactorsAsync();
    }

    private async void ApplyPlanDesignFactorFilters_Click(object sender, RoutedEventArgs e)
    {
        await LoadPlanDesignFactorsAsync();
    }

    private void AddPlanDesignFactor_Click(object sender, RoutedEventArgs e)
    {
        ResetPlanDesignFactorForm();
        SetStatus("Ready to add a new plan design factor.");
    }

    private void EditSelectedPlanDesignFactor_Click(object sender, RoutedEventArgs e)
    {
        if (PlanDesignFactorsGrid.SelectedItem is not PlanDesignFactorRow row)
        {
            SetStatus("Select a factor to edit.");
            return;
        }

        LoadPlanDesignFactorForm(row.ToRecord());
        SetStatus($"Editing factor for {row.CarrierName} / {row.BenefitLine}.");
    }

    private void DuplicateSelectedPlanDesignFactor_Click(object sender, RoutedEventArgs e)
    {
        if (PlanDesignFactorsGrid.SelectedItem is not PlanDesignFactorRow row)
        {
            SetStatus("Select a factor to duplicate.");
            return;
        }

        var copy = row.ToRecord();
        copy.Id = 0;
        LoadPlanDesignFactorForm(copy);
        PlanFactorFormTitleTextBlock.Text = "Duplicate Factor";
        SetStatus("Factor copied into the form. Review it, then save as a new record.");
    }

    private async void ArchiveSelectedPlanDesignFactor_Click(object sender, RoutedEventArgs e)
    {
        if (PlanDesignFactorsGrid.SelectedItem is not PlanDesignFactorRow row)
        {
            SetStatus("Select a factor to archive.");
            return;
        }

        var confirm = System.Windows.MessageBox.Show(
            this,
            $"Archive this factor?\n\n{row.CarrierName} - {row.BenefitLine} - {row.CurrentDesign} to {row.ProposedDesign}",
            "Archive Plan Design Factor",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question);
        if (confirm != MessageBoxResult.Yes)
        {
            return;
        }

        var result = await _databaseService.ArchivePlanDesignFactorAsync(row.Id);
        SetStatus(result.Message);
        await LoadPlanDesignFactorsAsync();
    }

    private async void SavePlanDesignFactor_Click(object sender, RoutedEventArgs e)
    {
        var errors = new List<string>();
        var factor = new PlanDesignFactorRecord
        {
            Id = _editingPlanDesignFactorId,
            CarrierName = PlanFactorCarrierTextBox.Text,
            BenefitCategory = ComboBoxText(PlanFactorBenefitCategoryComboBox),
            BenefitLine = PlanFactorBenefitLineTextBox.Text,
            CurrentDesign = PlanFactorCurrentDesignTextBox.Text,
            ProposedDesign = PlanFactorProposedDesignTextBox.Text,
            ChangeType = ComboBoxText(PlanFactorChangeTypeComboBox),
            Factor = ParseDecimalInput(PlanFactorValueTextBox.Text, "Factor", errors) ?? 0m,
            FactorDirection = ComboBoxText(PlanFactorDirectionComboBox),
            EffectiveDate = ParseDateInput(PlanFactorEffectiveDateTextBox.Text, "Effective Date", errors),
            ExpiryDate = ParseDateInput(PlanFactorExpiryDateTextBox.Text, "Expiry Date", errors),
            SourceDocument = PlanFactorSourceDocumentTextBox.Text,
            ConfidenceLevel = ComboBoxText(PlanFactorConfidenceComboBox),
            Notes = PlanFactorNotesTextBox.Text
        };

        if (factor.ExpiryDate.HasValue &&
            factor.EffectiveDate.HasValue &&
            factor.ExpiryDate.Value < factor.EffectiveDate.Value)
        {
            errors.Add("Expiry date cannot be before effective date.");
        }

        if (errors.Count > 0)
        {
            ShowPlanDesignFactorError(errors);
            return;
        }

        var result = await _databaseService.SavePlanDesignFactorAsync(factor);
        if (!result.Succeeded)
        {
            ShowPlanDesignFactorError([result.Message]);
            return;
        }

        ResetPlanDesignFactorForm();
        await LoadPlanDesignFactorsAsync();
        SetStatus(result.Message);
    }

    private void ClearPlanDesignFactorForm_Click(object sender, RoutedEventArgs e)
    {
        ResetPlanDesignFactorForm();
    }

    private async void ImportPlanDesignFactors_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Excel factor library (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls|All files (*.*)|*.*",
            CheckFileExists = true,
            Multiselect = false
        };

        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        try
        {
            SetBusy(true, "Importing plan design factors...");
            var result = await _databaseService.ImportPlanDesignFactorsFromExcelAsync(dialog.FileName);
            await LoadPlanDesignFactorsAsync();
            SetStatus($"Imported plan design factors: {result.InsertedCount} added, {result.UpdatedCount} updated, {result.SkippedCount} skipped.");
            if (result.Warnings.Count > 0)
            {
                System.Windows.MessageBox.Show(
                    this,
                    string.Join(Environment.NewLine, result.Warnings.Take(8)),
                    "Plan Design Factor Import Warnings",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            }
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            SetStatus("Plan design factor import failed.");
            System.Windows.MessageBox.Show(this, ex.Message, "Import Plan Design Factors", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
        finally
        {
            SetBusy(false);
        }
    }

    private async void ExportPlanDesignFactors_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new SaveFileDialog
        {
            Filter = "Excel workbook (*.xlsx)|*.xlsx",
            FileName = $"Plan_Design_Factors_{DateTime.Now:yyyyMMdd-HHmm}.xlsx",
            AddExtension = true,
            DefaultExt = ".xlsx"
        };

        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        try
        {
            SetBusy(true, "Exporting plan design factors...");
            await _databaseService.ExportPlanDesignFactorsToExcelAsync(dialog.FileName, BuildPlanDesignFactorFilter());
            SetStatus($"Exported plan design factors to {dialog.FileName}.");
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            SetStatus("Plan design factor export failed.");
            System.Windows.MessageBox.Show(this, ex.Message, "Export Plan Design Factors", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
        finally
        {
            SetBusy(false);
        }
    }

    private void CalculatorSelectedFactor_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (CalculatorSelectedFactorComboBox.SelectedItem is not PlanDesignFactorRow row)
        {
            return;
        }

        CalculatorCarrierComboBox.Text = row.CarrierName;
        CalculatorBenefitCategoryComboBox.Text = row.BenefitCategory;
        CalculatorBenefitLineTextBox.Text = row.BenefitLine;
        CalculatorCurrentDesignTextBox.Text = row.CurrentDesign;
        CalculatorProposedDesignTextBox.Text = row.ProposedDesign;
        CalculatorEffectiveDateTextBox.Text = row.EffectiveDateText == "N/A" ? "" : row.EffectiveDateText;
    }

    private async void CalculateAlternateImpact_Click(object sender, RoutedEventArgs e)
    {
        var errors = new List<string>();
        var currentPremium = ParseDecimalInput(CalculatorCurrentPremiumTextBox.Text, "Current Premium", errors);
        var effectiveDate = ParseDateInput(CalculatorEffectiveDateTextBox.Text, "Effective Date", errors);
        if (currentPremium is null || currentPremium <= 0m)
        {
            errors.Add("Current premium must be greater than 0.");
        }

        if (errors.Count > 0)
        {
            ShowPlanDesignFactorError(errors);
            return;
        }

        PlanDesignFactorEstimate? estimate;
        if (CalculatorSelectedFactorComboBox.SelectedItem is PlanDesignFactorRow selectedFactor)
        {
            estimate = DatabaseService.BuildPlanDesignFactorEstimate(currentPremium!.Value, selectedFactor.ToRecord());
        }
        else
        {
            estimate = await _databaseService.EstimateAlternatePremiumAsync(
                ComboBoxText(CalculatorCarrierComboBox),
                CalculatorBenefitLineTextBox.Text,
                CalculatorCurrentDesignTextBox.Text,
                CalculatorProposedDesignTextBox.Text,
                currentPremium!.Value,
                effectiveDate);
        }

        if (estimate is null)
        {
            ShowPlanDesignFactorError(["No active factor matched that carrier, benefit line, and plan design change."]);
            return;
        }

        CalculatorRevisedPremiumTextBlock.Text = FormatCurrencyCents(estimate.EstimatedRevisedPremium);
        CalculatorDollarChangeTextBlock.Text = FormatCurrencyCents(estimate.DollarChange);
        CalculatorPercentageChangeTextBlock.Text = estimate.PercentageChange.ToString("0.00%", CultureInfo.InvariantCulture);
        CalculatorFactorSourceTextBlock.Text = estimate.FactorSource;
        CalculatorConfidenceTextBlock.Text = estimate.ConfidenceLevel;
        CalculatorConfidenceTextBlock.Foreground = estimate.ConfidenceLevel.Equals(PlanDesignFactorRecord.PlaceholderConfidence, StringComparison.OrdinalIgnoreCase)
            ? Brushes.DarkOrange
            : Brushes.Black;
        SetStatus($"Estimated revised premium: {FormatCurrencyCents(estimate.EstimatedRevisedPremium)} using factor {estimate.FactorRecord.Factor:0.0000}.");
    }

    private async Task ShowPlanDesignIntelligenceAsync(string? subpage = null)
    {
        var workspace = _selectedClientWorkspace;
        if (workspace is null)
        {
            ShowClientsPage();
            SetStatus("Select a client before opening Plan Design.");
            return;
        }

        SetBusy(true, "Loading client plan design...");
        HomeDashboardScrollViewer.Visibility = Visibility.Collapsed;
        DashboardScrollViewer.Visibility = Visibility.Collapsed;
        UnderwritingWorkbenchScrollViewer.Visibility = Visibility.Collapsed;
        AssumptionLibraryScrollViewer.Visibility = Visibility.Collapsed;
        ReportsAnalyticsScrollViewer.Visibility = Visibility.Collapsed;
        BenchmarkingDashboardScrollViewer.Visibility = Visibility.Collapsed;
        NegotiationWorkbenchScrollViewer.Visibility = Visibility.Collapsed;
        PlanDesignFactorsScrollViewer.Visibility = Visibility.Collapsed;
        DocumentIntelligenceScrollViewer.Visibility = Visibility.Collapsed;
        DataSourcesScrollViewer.Visibility = Visibility.Collapsed;
        PlanDesignIntelligenceScrollViewer.Visibility = Visibility.Collapsed;
        ProposalComparisonScrollViewer.Visibility = Visibility.Collapsed;
        ClientsScrollViewer.Visibility = Visibility.Visible;
        ClientListPanel.Visibility = Visibility.Collapsed;
        ClientWorkspacePanel.Visibility = Visibility.Visible;
        HideClientWorkspaceContentPanels();
        ClientPlanDesignWorkspaceView.Visibility = Visibility.Visible;
        ClientsScrollViewer.ScrollToTop();
        SetOperationalFooterVisible(false);

        try
        {
            await ClientPlanDesignWorkspaceView.LoadClientAsync(
                workspace.DatabaseClientId,
                workspace.ClientName,
                workspace.PrimaryCarrier,
                new PlanDesignWorkspaceClientContext
                {
                    TotalEligibleEmployees = workspace.Employees > 0 ? workspace.Employees : null,
                    Consultant = FirstNonBlank(workspace.AccountManager, workspace.AccountExecutive),
                    PlanStatus = workspace.Status,
                    PlanType = workspace.PlanType
                },
                subpage);
            ClientsScrollViewer.ScrollToTop();

            SetStatus($"Opened Plan Design Workspace for {workspace.ClientName}.");
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            SetStatus("Plan Design workspace opened, but saved plan design data could not load.");
            System.Windows.MessageBox.Show(
                this,
                "Plan Design opened, but the saved plan design data could not be loaded for this client. You can keep using the app while this clientÃ¢â‚¬â„¢s plan design data is reviewed.",
                "Plan Design Workspace",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }
        finally
        {
            SetBusy(false);
        }
    }

    private long? ResolvePlanDesignClientId(RenewalQueueItem? selectedRenewal)
    {
        if (_selectedClientWorkspace?.DatabaseClientId is > 0)
        {
            return _selectedClientWorkspace.DatabaseClientId;
        }

        return selectedRenewal?.ClientId is > 0
            ? selectedRenewal.ClientId
            : null;
    }

    private async Task EnsurePlanDesignIntelligenceClientsAsync()
    {
        if (_planDesignIntelligenceClients.Count > 0)
        {
            return;
        }

        _isLoadingPlanDesignIntelligence = true;
        try
        {
            var clients = await _planDesignIntelligenceService.GetClientOptionsAsync();
            _planDesignIntelligenceClients.Clear();
            foreach (var client in clients)
            {
                _planDesignIntelligenceClients.Add(client);
            }

            PlanDesignIntelligenceClientComboBox.SelectedItem = ResolveDefaultPlanDesignIntelligenceClient();
        }
        finally
        {
            _isLoadingPlanDesignIntelligence = false;
        }
    }

    private PlanDesignIntelligenceClientOption? ResolveDefaultPlanDesignIntelligenceClient()
    {
        var selectedClientIds = new[]
        {
            (RenewalQueueGrid.SelectedItem as RenewalQueueItem)?.ClientId,
            _selectedUnderwritingRenewal?.ClientId,
            _selectedNegotiationRenewal?.ClientId
        }
            .Where(value => value.HasValue && value.Value > 0)
            .Select(value => value!.Value)
            .ToList();

        foreach (var clientId in selectedClientIds)
        {
            var match = _planDesignIntelligenceClients.FirstOrDefault(client => client.ClientId == clientId);
            if (match is not null)
            {
                return match;
            }
        }

        return PlanDesignIntelligenceClientComboBox.SelectedItem as PlanDesignIntelligenceClientOption
               ?? _planDesignIntelligenceClients.FirstOrDefault();
    }

    private async Task LoadPlanDesignIntelligenceAsync()
    {
        if (PlanDesignIntelligenceClientComboBox.SelectedItem is not PlanDesignIntelligenceClientOption client)
        {
            ClearPlanDesignIntelligenceResult("No client selected.");
            SetStatus("Plan Design Intelligence: no saved client is available.");
            return;
        }

        try
        {
            SetBusy(true, "Loading plan design intelligence...");
            var result = await _planDesignIntelligenceService.AnalyzeClientAsync(client.ClientId);
            PopulatePlanDesignIntelligenceResult(result);
            SetStatus($"Plan Design Intelligence loaded for {result.ClientName}.");
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            ClearPlanDesignIntelligenceResult("Plan Design Intelligence could not be loaded.");
            System.Windows.MessageBox.Show(this, ex.Message, "Plan Design Intelligence", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
        finally
        {
            SetBusy(false);
        }
    }

    private void PopulatePlanDesignIntelligenceResult(PlanDesignIntelligenceResult result)
    {
        _planDesignIntelligenceSummaryRows.Clear();
        foreach (var row in result.SummaryRows)
        {
            _planDesignIntelligenceSummaryRows.Add(row);
        }

        _planDesignIntelligenceReadinessRows.Clear();
        foreach (var row in result.ReadinessRows)
        {
            _planDesignIntelligenceReadinessRows.Add(row);
        }

        _planDesignIntelligenceMissingRows.Clear();
        foreach (var finding in result.MissingDataFindings.OrderBy(PlanDesignFindingRiskRank).ThenBy(finding => finding.Benefit))
        {
            _planDesignIntelligenceMissingRows.Add(finding);
        }

        _planDesignIntelligenceNormalizationRows.Clear();
        foreach (var finding in result.NormalizationSuggestions.OrderBy(finding => finding.Benefit))
        {
            _planDesignIntelligenceNormalizationRows.Add(finding);
        }

        _planDesignIntelligencePricingRows.Clear();
        foreach (var finding in result.PricingImpactOpportunities.OrderBy(PlanDesignFindingRiskRank).ThenBy(finding => finding.Benefit))
        {
            _planDesignIntelligencePricingRows.Add(finding);
        }

        _planDesignIntelligenceFindingRows.Clear();
        foreach (var finding in result.Findings.OrderBy(PlanDesignFindingRiskRank).ThenBy(finding => finding.FindingType).ThenBy(finding => finding.Benefit))
        {
            _planDesignIntelligenceFindingRows.Add(finding);
        }

        _planDesignIntelligenceSourceRows.Clear();
        foreach (var row in result.SourceTables)
        {
            _planDesignIntelligenceSourceRows.Add(row);
        }

        PlanDesignIntelligenceSelectedClientTextBlock.Text = FirstNonBlank(result.ClientName, $"Client #{result.ClientId}");
        var usedTables = result.SourceTables.Count(row => row.Used);
        var itemCount = result.SummaryRows.Sum(row => row.FeatureCount);
        PlanDesignIntelligenceAvailabilityTextBlock.Text = $"{itemCount} item(s) from {usedTables} table(s)";
        PlanDesignIntelligenceReadinessTextBlock.Text = result.OverallReadinessStatus;
        PlanDesignIntelligenceFindingCountTextBlock.Text = result.Findings.Count.ToString(CultureInfo.InvariantCulture);
    }

    private void ClearPlanDesignIntelligenceResult(string status)
    {
        _planDesignIntelligenceSummaryRows.Clear();
        _planDesignIntelligenceReadinessRows.Clear();
        _planDesignIntelligenceMissingRows.Clear();
        _planDesignIntelligenceNormalizationRows.Clear();
        _planDesignIntelligencePricingRows.Clear();
        _planDesignIntelligenceFindingRows.Clear();
        _planDesignIntelligenceSourceRows.Clear();
        PlanDesignIntelligenceSelectedClientTextBlock.Text = "-";
        PlanDesignIntelligenceAvailabilityTextBlock.Text = status;
        PlanDesignIntelligenceReadinessTextBlock.Text = "-";
        PlanDesignIntelligenceFindingCountTextBlock.Text = "0";
    }

    private static int PlanDesignFindingRiskRank(PlanDesignIntelligenceFinding finding)
    {
        return finding.RiskLevel switch
        {
            PlanDesignRiskLevels.High => 0,
            PlanDesignRiskLevels.Medium => 1,
            _ => 2
        };
    }

    private async Task ShowPlanDesignFactorsAsync()
    {
        HomeDashboardScrollViewer.Visibility = Visibility.Collapsed;
        DashboardScrollViewer.Visibility = Visibility.Collapsed;
        UnderwritingWorkbenchScrollViewer.Visibility = Visibility.Collapsed;
        AssumptionLibraryScrollViewer.Visibility = Visibility.Collapsed;
        ReportsAnalyticsScrollViewer.Visibility = Visibility.Collapsed;
        BenchmarkingDashboardScrollViewer.Visibility = Visibility.Collapsed;
        NegotiationWorkbenchScrollViewer.Visibility = Visibility.Collapsed;
        ProposalComparisonScrollViewer.Visibility = Visibility.Collapsed;
        PlanDesignIntelligenceScrollViewer.Visibility = Visibility.Collapsed;
        DocumentIntelligenceScrollViewer.Visibility = Visibility.Collapsed;
        ClientsScrollViewer.Visibility = Visibility.Collapsed;
        DataSourcesScrollViewer.Visibility = Visibility.Collapsed;
        PlanDesignFactorsScrollViewer.Visibility = Visibility.Visible;
        PlanDesignFactorsScrollViewer.ScrollToTop();
        SetOperationalFooterVisible(true);

        await LoadPlanDesignFactorsAsync();
    }

    private async Task LoadPlanDesignFactorsAsync()
    {
        try
        {
            SetBusy(true, "Loading plan design factors...");
            var filter = BuildPlanDesignFactorFilter();
            var records = await _databaseService.GetPlanDesignFactorsAsync(filter);
            var allRecords = await _databaseService.GetPlanDesignFactorsAsync();
            var summary = await _databaseService.GetPlanDesignFactorSummaryAsync();

            _planDesignFactorRows.Clear();
            foreach (var row in records.Select(record => new PlanDesignFactorRow(record)))
            {
                _planDesignFactorRows.Add(row);
            }

            var today = DateOnly.FromDateTime(DateTime.Today);
            _planDesignFactorOptions.Clear();
            foreach (var option in allRecords.Where(record => record.IsActiveOn(today)).Select(record => new PlanDesignFactorRow(record)))
            {
                _planDesignFactorOptions.Add(option);
            }

            PopulatePlanDesignFactorPicklists(allRecords);
            PlanFactorTotalTextBlock.Text = summary.TotalFactors.ToString(CultureInfo.InvariantCulture);
            PlanFactorCarriersTextBlock.Text = summary.CarriersCovered.ToString(CultureInfo.InvariantCulture);
            PlanFactorActiveTextBlock.Text = summary.ActiveFactors.ToString(CultureInfo.InvariantCulture);
            PlanFactorPlaceholderTextBlock.Text = summary.EstimatedPlaceholderFactors.ToString(CultureInfo.InvariantCulture);
            PlanFactorExpiringTextBlock.Text = summary.ExpiringSoon.ToString(CultureInfo.InvariantCulture);
            PlanFactorLibraryStatusTextBlock.Text = $"{records.Count} factor(s) shown. Estimated placeholders must be replaced with carrier-approved factors before client reliance.";
            SetStatus($"Opened Plan Design Factors with {records.Count} visible factor(s).");
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            SetStatus("Plan Design Factors could not be loaded.");
            System.Windows.MessageBox.Show(this, ex.Message, "Plan Design Factors", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
        finally
        {
            SetBusy(false);
        }
    }

    private void PopulatePlanDesignFactorPicklists(IReadOnlyCollection<PlanDesignFactorRecord> records)
    {
        var selectedCarrierFilter = ComboBoxText(PlanFactorCarrierFilterComboBox);
        var selectedCalculatorCarrier = ComboBoxText(CalculatorCarrierComboBox);
        var carriers = records
            .Select(record => record.CarrierName)
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(value => value)
            .ToList();

        SetComboBoxValues(PlanFactorCarrierFilterComboBox, carriers, includeBlank: true);
        PlanFactorCarrierFilterComboBox.Text = selectedCarrierFilter;
        SetComboBoxValues(CalculatorCarrierComboBox, carriers, includeBlank: true);
        CalculatorCarrierComboBox.Text = selectedCalculatorCarrier;
    }

    private PlanDesignFactorFilter BuildPlanDesignFactorFilter()
    {
        return new PlanDesignFactorFilter
        {
            SearchTerm = PlanFactorSearchTextBox.Text,
            CarrierName = ComboBoxText(PlanFactorCarrierFilterComboBox),
            BenefitCategory = ComboBoxText(PlanFactorBenefitCategoryFilterComboBox),
            ChangeType = ComboBoxText(PlanFactorChangeTypeFilterComboBox),
            ConfidenceLevel = ComboBoxText(PlanFactorConfidenceFilterComboBox),
            EffectiveDate = TryParseDate(PlanFactorEffectiveDateFilterTextBox.Text)
        };
    }

    private void LoadPlanDesignFactorForm(PlanDesignFactorRecord factor)
    {
        _editingPlanDesignFactorId = factor.Id;
        PlanFactorFormTitleTextBlock.Text = factor.Id > 0 ? "Edit Factor" : "Add Factor";
        PlanFactorCarrierTextBox.Text = factor.CarrierName;
        PlanFactorBenefitCategoryComboBox.Text = factor.BenefitCategory;
        PlanFactorBenefitLineTextBox.Text = factor.BenefitLine;
        PlanFactorCurrentDesignTextBox.Text = factor.CurrentDesign;
        PlanFactorProposedDesignTextBox.Text = factor.ProposedDesign;
        PlanFactorChangeTypeComboBox.Text = factor.ChangeType;
        PlanFactorValueTextBox.Text = factor.Factor.ToString("0.0000", CultureInfo.InvariantCulture);
        PlanFactorDirectionComboBox.Text = factor.FactorDirection;
        PlanFactorEffectiveDateTextBox.Text = ToDateInputText(factor.EffectiveDate);
        PlanFactorExpiryDateTextBox.Text = ToDateInputText(factor.ExpiryDate);
        PlanFactorSourceDocumentTextBox.Text = factor.SourceDocument;
        PlanFactorConfidenceComboBox.Text = factor.ConfidenceLevel;
        PlanFactorNotesTextBox.Text = factor.Notes;
    }

    private void ResetPlanDesignFactorForm()
    {
        _editingPlanDesignFactorId = 0;
        PlanFactorFormTitleTextBlock.Text = "Add Factor";
        PlanFactorCarrierTextBox.Text = "";
        PlanFactorBenefitCategoryComboBox.Text = PlanDesignFactorReferenceData.BenefitCategories.First();
        PlanFactorBenefitLineTextBox.Text = "";
        PlanFactorCurrentDesignTextBox.Text = "";
        PlanFactorProposedDesignTextBox.Text = "";
        PlanFactorChangeTypeComboBox.Text = PlanDesignFactorReferenceData.ChangeTypes.First();
        PlanFactorValueTextBox.Text = "1.0000";
        PlanFactorDirectionComboBox.Text = "Neutral";
        PlanFactorEffectiveDateTextBox.Text = DateTime.Today.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        PlanFactorExpiryDateTextBox.Text = "";
        PlanFactorSourceDocumentTextBox.Text = "";
        PlanFactorConfidenceComboBox.Text = "Carrier Provided";
        PlanFactorNotesTextBox.Text = "";
    }

    private static string ComboBoxText(ComboBox comboBox)
    {
        return !string.IsNullOrWhiteSpace(comboBox.Text)
            ? comboBox.Text.Trim()
            : SelectedComboBoxText(comboBox);
    }

    private static decimal? ParseDecimalInput(string value, string label, List<string> errors)
    {
        var cleaned = (value ?? "").Trim().Replace("$", "", StringComparison.Ordinal).Replace(",", "", StringComparison.Ordinal);
        if (decimal.TryParse(cleaned, NumberStyles.Any, CultureInfo.InvariantCulture, out var parsed) ||
            decimal.TryParse(cleaned, NumberStyles.Any, CultureInfo.CurrentCulture, out parsed))
        {
            return parsed;
        }

        errors.Add($"{label} must be a valid number.");
        return null;
    }

    private static DateOnly? ParseDateInput(string value, string label, List<string> errors)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        var parsed = TryParseDate(value);
        if (parsed.HasValue)
        {
            return parsed;
        }

        errors.Add($"{label} must use a valid date such as 2026-01-01.");
        return null;
    }

    private static DateOnly? TryParseDate(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        if (DateOnly.TryParseExact(value.Trim(), "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var date))
        {
            return date;
        }

        return DateOnly.TryParse(value, CultureInfo.CurrentCulture, DateTimeStyles.None, out date) ? date : null;
    }

    private static string ToDateInputText(DateOnly? date)
    {
        return date?.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) ?? "";
    }

    private void ShowPlanDesignFactorError(IReadOnlyCollection<string> errors)
    {
        var message = string.Join(Environment.NewLine, errors.Where(error => !string.IsNullOrWhiteSpace(error)));
        SetStatus(message);
        System.Windows.MessageBox.Show(this, message, "Plan Design Factors", MessageBoxButton.OK, MessageBoxImage.Warning);
    }

    private static string FormatCurrencyCents(decimal value)
    {
        return value.ToString("$#,##0.00", CultureInfo.InvariantCulture);
    }

    private async void RefreshAssumptionLibrary_Click(object sender, RoutedEventArgs e)
    {
        await LoadAssumptionLibraryAsync(forceSeed: false);
    }

    private void NavigateAssumptionSection_Click(object sender, RoutedEventArgs e)
    {
        NavigateAssumptionSection((sender as FrameworkElement)?.Tag?.ToString());
    }

    private void NavigateAssumptionSection(string? sectionKey)
    {
        var target = sectionKey switch
        {
            "CarrierSummary" => CarrierSummarySectionBorder,
            "BlockProfiles" => BlockProfilesSectionBorder,
            "BenefitLineAssumptions" => BenefitLineAssumptionsSectionBorder,
            "TlrBasisOptions" => TlrBasisOptionsSectionBorder,
            "ManualSections" => ManualSectionsSectionBorder,
            _ => null
        };

        var label = sectionKey switch
        {
            "CarrierSummary" => "Carrier Summary",
            "BlockProfiles" => "Block Profiles",
            "BenefitLineAssumptions" => "Benefit Line Assumptions",
            "TlrBasisOptions" => "TLR Basis Options",
            "ManualSections" => "Manual Sections",
            _ => "Assumption Library"
        };

        if (target is null)
        {
            SetStatus("Assumption Library section could not be found.");
            return;
        }

        DocumentIntelligenceScrollViewer.Visibility = Visibility.Collapsed;
        ClientsScrollViewer.Visibility = Visibility.Collapsed;
        AssumptionLibraryScrollViewer.Visibility = Visibility.Visible;
        SetOperationalFooterVisible(true);
        target.Dispatcher.BeginInvoke(
            () =>
            {
                target.BringIntoView();
                target.Focus();
            },
            DispatcherPriority.Background);
        SetStatus($"Opened {label} in the Assumption Library.");
    }

    private async Task ShowAssumptionLibraryAsync()
    {
        HomeDashboardScrollViewer.Visibility = Visibility.Collapsed;
        DashboardScrollViewer.Visibility = Visibility.Collapsed;
        UnderwritingWorkbenchScrollViewer.Visibility = Visibility.Collapsed;
        ReportsAnalyticsScrollViewer.Visibility = Visibility.Collapsed;
        BenchmarkingDashboardScrollViewer.Visibility = Visibility.Collapsed;
        ProposalComparisonScrollViewer.Visibility = Visibility.Collapsed;
        PlanDesignIntelligenceScrollViewer.Visibility = Visibility.Collapsed;
        PlanDesignFactorsScrollViewer.Visibility = Visibility.Collapsed;
        DocumentIntelligenceScrollViewer.Visibility = Visibility.Collapsed;
        ClientsScrollViewer.Visibility = Visibility.Collapsed;
        NegotiationWorkbenchScrollViewer.Visibility = Visibility.Collapsed;
        DataSourcesScrollViewer.Visibility = Visibility.Collapsed;
        AssumptionLibraryScrollViewer.Visibility = Visibility.Visible;
        AssumptionLibraryScrollViewer.ScrollToTop();
        SetOperationalFooterVisible(true);

        await LoadAssumptionLibraryAsync(forceSeed: true);
    }

    private async Task LoadAssumptionLibraryAsync(bool forceSeed)
    {
        try
        {
            SetBusy(true, "Loading carrier underwriting assumptions...");
            var records = await _databaseService.GetCarrierAssumptionsAsync();
            var seededDefaults = false;

            if (forceSeed && records.Count == 0)
            {
                var seed = BuildDefaultCarrierAssumptions();
                var save = await _databaseService.SaveCarrierAssumptionsAsync(seed);
                seededDefaults = save.Succeeded;
                records = await _databaseService.GetCarrierAssumptionsAsync();
            }

            if (NormalizeCarrierAssumptionTlrBasis(records))
            {
                await _databaseService.SaveCarrierAssumptionsAsync(records);
                records = await _databaseService.GetCarrierAssumptionsAsync();
            }

            var displayRows = BuildCarrierAssumptionLibraryRows(records);

            _carrierAssumptionRows.Clear();
            foreach (var row in displayRows)
            {
                _carrierAssumptionRows.Add(row);
            }

            _carrierBlockProfiles.Clear();
            foreach (var profile in BuildCarrierBlockProfiles(records))
            {
                _carrierBlockProfiles.Add(profile);
            }

            _carrierAssumptionSummaries.Clear();
            foreach (var summary in BuildCarrierAssumptionSummaries(records))
            {
                _carrierAssumptionSummaries.Add(summary);
            }

            AssumptionCarrierCountTextBlock.Text = records.Select(item => item.Carrier).Where(value => !string.IsNullOrWhiteSpace(value)).Distinct(StringComparer.OrdinalIgnoreCase).Count().ToString(CultureInfo.InvariantCulture);
            AssumptionBlockCountTextBlock.Text = _carrierBlockProfiles.Count.ToString(CultureInfo.InvariantCulture);
            AssumptionBenefitCountTextBlock.Text = displayRows.Count.ToString(CultureInfo.InvariantCulture);
            AssumptionEffectiveDateTextBlock.Text = FormatLatestDate(records.Select(item => item.EffectiveDate));
            AssumptionLibrarySeedStatusTextBlock.Text = seededDefaults
                ? "Starter Westland carrier assumptions were added to the local database."
                : "Loaded from the local Benefits Intelligence database.";
            SetStatus($"Opened Assumption Library with {displayRows.Count} grouped manual rule(s).");
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            SetStatus("Assumption Library could not be loaded.");
            System.Windows.MessageBox.Show(
                this,
                $"The Assumption Library could not be loaded: {ex.Message}",
                "Assumption Library",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }
        finally
        {
            SetBusy(false);
        }
    }

    private static List<CarrierBlockProfileRow> BuildCarrierBlockProfiles(IReadOnlyCollection<CarrierAssumptionRecord> records)
    {
        return records
            .GroupBy(item => string.IsNullOrWhiteSpace(item.BlockName) ? "Standard Manual" : item.BlockName, StringComparer.OrdinalIgnoreCase)
            .OrderBy(group => group.Key)
            .Select(group => new CarrierBlockProfileRow(group.Key, group.ToList()))
            .ToList();
    }

    private static List<CarrierAssumptionLibraryRow> BuildCarrierAssumptionLibraryRows(IReadOnlyCollection<CarrierAssumptionRecord> records)
    {
        return records
            .GroupBy(item => new
            {
                BlockName = item.BlockName,
                BenefitLine = item.BenefitLine,
                FundingType = item.FundingType,
                TargetLossRatio = item.TargetLossRatio,
                TlrBasis = TlrBasisOptions.Normalize(item.TlrBasis),
                PremiumTax = item.PremiumTax,
                Commission = item.Commission,
                AdminMargin = item.AdminMargin,
                RiskMargin = item.RiskMargin,
                Trend = item.Trend,
                CredibilityFactor = item.CredibilityFactor,
                PoolingThreshold = item.PoolingThreshold,
                PoolingCharge = item.PoolingCharge,
                ReserveIbnr = item.ReserveIbnr,
                EffectiveDate = item.EffectiveDate,
                Status = item.Status,
                Notes = item.Notes
            })
            .Select(group => new CarrierAssumptionLibraryRow(group.Select(item => item).ToList()))
            .OrderBy(row => row.BlockName)
            .ThenBy(row => row.BenefitLine)
            .ThenBy(row => row.Carrier)
            .ToList();
    }

    private static List<CarrierAssumptionSummaryRow> BuildCarrierAssumptionSummaries(IReadOnlyCollection<CarrierAssumptionRecord> records)
    {
        return records
            .GroupBy(item => item.Carrier, StringComparer.OrdinalIgnoreCase)
            .OrderBy(group => group.Key)
            .Select(group => new CarrierAssumptionSummaryRow
            {
                Carrier = group.Key,
                Blocks = group.Select(item => item.BlockName).Where(value => !string.IsNullOrWhiteSpace(value)).Distinct(StringComparer.OrdinalIgnoreCase).Count(),
                BenefitRules = group.Count(),
                TargetLossRatioRange = FormatPercentRange(group.Select(item => item.TargetLossRatio)),
                LatestEffectiveDate = FormatLatestDate(group.Select(item => item.EffectiveDate)),
                Status = group.All(item => item.Status.Equals("Active", StringComparison.OrdinalIgnoreCase)) ? "Active" : "Review"
            })
            .ToList();
    }

    private static bool NormalizeCarrierAssumptionTlrBasis(IReadOnlyCollection<CarrierAssumptionRecord> records)
    {
        var changed = false;
        foreach (var record in records)
        {
            var normalized = TlrBasisOptions.Normalize(record.TlrBasis);
            if (!record.TlrBasis.Equals(normalized, StringComparison.Ordinal))
            {
                record.TlrBasis = normalized;
                changed = true;
            }
        }

        return changed;
    }

    private static string FormatLatestDate(IEnumerable<DateOnly?> dates)
    {
        var values = dates.Where(date => date.HasValue).Select(date => date!.Value).ToList();
        return values.Count == 0
            ? "N/A"
            : values.Max().ToString("MMM d, yyyy", CultureInfo.InvariantCulture);
    }

    private static string FormatPercentRange(IEnumerable<decimal?> values)
    {
        var percentages = values.Where(value => value.HasValue).Select(value => value!.Value).ToList();
        if (percentages.Count == 0)
        {
            return "N/A";
        }

        var min = percentages.Min();
        var max = percentages.Max();
        return min == max
            ? $"{min:0.0}%"
            : $"{min:0.0}% - {max:0.0}%";
    }

    private static List<CarrierAssumptionRecord> BuildDefaultCarrierAssumptions()
    {
        var effectiveDate = new DateOnly(2026, 1, 1);
        var records = new List<CarrierAssumptionRecord>();

        foreach (var carrier in new[] { "Equitable", "RBC", "Empire" })
        {
            records.Add(CreateCarrierAssumption(carrier, "Westland Partnership Block", "Extended Health Care", "Experience-rated", 90m, TlrBasisOptions.NetOfPremiumTaxOnly, 2m, null, 8m, 100m, 100000m, 0m, 0m, effectiveDate, "Partnership block. Confirm commission schedule before modelling."));
            records.Add(CreateCarrierAssumption(carrier, "Westland Partnership Block", "Dental", "Experience-rated", 90m, TlrBasisOptions.NetOfPremiumTaxOnly, 2m, null, 6m, 100m, null, 0m, 0m, effectiveDate, "Partnership block. Confirm commission schedule before modelling."));
            records.Add(CreateCarrierAssumption(carrier, "Westland Partnership Block", "Life / AD&D", "Pooled", 85m, TlrBasisOptions.GrossBilledPremium, 2m, null, 0m, 100m, null, 0m, 0m, effectiveDate, "Use as a starting pooled benefit target."));
            records.Add(CreateCarrierAssumption(carrier, "Westland Partnership Block", "Disability", "Pooled", 85m, TlrBasisOptions.GrossBilledPremium, 2m, null, 3m, 100m, null, 0m, 0m, effectiveDate, "Confirm LTD/STD reserve basis by carrier."));
        }

        records.Add(CreateCarrierAssumption("Manulife", "Standard Renewal Manual", "Extended Health Care", "Experience-rated", 76.3m, TlrBasisOptions.GrossBilledPremium, 2m, null, 8m, 100m, 100000m, 0m, 0m, effectiveDate, "Starter default from current underwriting model."));
        records.Add(CreateCarrierAssumption("Manulife", "Standard Renewal Manual", "Dental", "Experience-rated", 76.3m, TlrBasisOptions.GrossBilledPremium, 2m, null, 6m, 100m, null, 0m, 0m, effectiveDate, "Starter default from current underwriting model."));
        records.Add(CreateCarrierAssumption("Canada Life", "Standard Renewal Manual", "Extended Health Care", "Experience-rated", 78m, TlrBasisOptions.GrossBilledPremium, 2m, null, 7m, 100m, 100000m, 0m, 0m, effectiveDate, "Canada Life underwriting condition for independent contractors (50% of group): the client must keep enrollment records current and accurate, and each covered independent contractor must have a minimum one-year contract with the client."));
        records.Add(CreateCarrierAssumption("Sun Life", "Standard Renewal Manual", "Extended Health Care", "Experience-rated", 78m, TlrBasisOptions.GrossBilledPremium, 2m, null, 7m, 100m, 100000m, 0m, 0m, effectiveDate, "Replace with official Sun Life block assumption when available."));

        return records;
    }

    private static CarrierAssumptionRecord CreateCarrierAssumption(
        string carrier,
        string blockName,
        string benefitLine,
        string fundingType,
        decimal targetLossRatio,
        string tlrBasis,
        decimal premiumTax,
        decimal? commission,
        decimal trend,
        decimal credibilityFactor,
        decimal? poolingThreshold,
        decimal poolingCharge,
        decimal reserveIbnr,
        DateOnly effectiveDate,
        string notes)
    {
        return new CarrierAssumptionRecord
        {
            Carrier = carrier,
            BlockName = blockName,
            BenefitLine = benefitLine,
            FundingType = fundingType,
            TargetLossRatio = targetLossRatio,
            TlrBasis = tlrBasis,
            PremiumTax = premiumTax,
            Commission = commission,
            AdminMargin = null,
            RiskMargin = null,
            Trend = trend,
            CredibilityFactor = credibilityFactor,
            PoolingThreshold = poolingThreshold,
            PoolingCharge = poolingCharge,
            ReserveIbnr = reserveIbnr,
            EffectiveDate = effectiveDate,
            Status = "Active",
            Notes = notes,
            DateLoaded = DateTime.Now
        };
    }

    private void UploadRenewalPackage_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Renewal package files (*.pdf;*.xlsx;*.xlsm;*.xls)|*.pdf;*.xlsx;*.xlsm;*.xls|PDF files (*.pdf)|*.pdf|Excel files (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls|All files (*.*)|*.*",
            CheckFileExists = true,
            Multiselect = true
        };

        if (dialog.ShowDialog() != true)
        {
            return;
        }

        var pdfs = dialog.FileNames.Where(IsPdfFile).ToList();
        var workbooks = dialog.FileNames.Where(IsWorkbookFile).ToList();

        if (pdfs.Count > 0)
        {
            PdfPathTextBox.Text = pdfs[0];
            foreach (var path in pdfs.Skip(1))
            {
                AddDistinctPath(_additionalPdfPaths, path, PdfPathTextBox.Text);
            }
        }

        if (workbooks.Count > 0)
        {
            WorkbookPathTextBox.Text = workbooks[0];
            foreach (var path in workbooks.Skip(1))
            {
                AddDistinctPath(_additionalWorkbookPaths, path, WorkbookPathTextBox.Text);
            }
        }

        SaveSettings();
        InvalidateAnalysis();
        SetStatus($"Renewal package selected: {pdfs.Count} PDF file(s), {workbooks.Count} Excel file(s). Click Validate, then Export Pack.");
    }

    private void UploadClaimsReport_Click(object sender, RoutedEventArgs e)
    {
        var choice = System.Windows.MessageBox.Show(
            this,
            "Choose a claims experience folder? Select No to add individual claims report files instead.",
            "Upload Claims Report",
            MessageBoxButton.YesNoCancel,
            MessageBoxImage.Question);

        if (choice == MessageBoxResult.Cancel)
        {
            return;
        }

        if (choice == MessageBoxResult.Yes)
        {
            if (BrowseFolder(CurrentYearFolderTextBox, "Choose the current year or claims experience folder"))
            {
                SaveSettings();
                SetStatus("Claims experience folder selected. Use Export Claims Model or Export Claims PPT when ready.");
            }

            return;
        }

        var dialog = new OpenFileDialog
        {
            Filter = "Claims report files (*.pdf;*.xlsx;*.xlsm;*.xls)|*.pdf;*.xlsx;*.xlsm;*.xls|PDF files (*.pdf)|*.pdf|Excel files (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls|All files (*.*)|*.*",
            CheckFileExists = true,
            Multiselect = true
        };

        if (dialog.ShowDialog() != true)
        {
            return;
        }

        foreach (var fileName in dialog.FileNames)
        {
            if (IsPdfFile(fileName))
            {
                AddDistinctPath(_additionalPdfPaths, fileName, PdfPathTextBox.Text);
            }
            else if (IsWorkbookFile(fileName))
            {
                AddDistinctPath(_additionalWorkbookPaths, fileName, WorkbookPathTextBox.Text);
            }
        }

        CurrentYearFolderTextBox.Text = Path.GetDirectoryName(dialog.FileNames[0]) ?? CurrentYearFolderTextBox.Text;
        SaveSettings();
        InvalidateAnalysis();
        SetStatus($"Claims report file(s) selected: {dialog.FileNames.Length}. Use Export Claims Model or Export Claims PPT when ready.");
    }

    private void UploadBenefitBooklet_Click(object sender, RoutedEventArgs e)
    {
        if (PromptForBookletPdfFiles())
        {
            SaveSettings();
        }
    }

    private void UploadProposal_Click(object sender, RoutedEventArgs e)
    {
        TestCarrierProposal_Click(sender, e);
    }

    private async void CompareCarrierProposals_Click(object sender, RoutedEventArgs e)
    {
        await ShowProposalComparisonWorkspaceAsync();
    }

    private async Task ShowProposalPipelineOverviewAsync()
    {
        try
        {
            var overview = await _proposalPipelineOverviewService.BuildAsync();
            ReplaceCollection(_proposalPipelineOverviewRows, overview.Stages);
            UpdateProposalPipelineOverviewMetrics(overview);
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            SetStatus("Proposal pipeline could not be loaded. Opening an empty pipeline overview.");
            ReplaceCollection(_proposalPipelineOverviewRows, []);
            UpdateProposalPipelineOverviewMetrics(new ProposalPipelineOverview());
        }

        HomeDashboardScrollViewer.Visibility = Visibility.Collapsed;
        DashboardScrollViewer.Visibility = Visibility.Collapsed;
        UnderwritingWorkbenchScrollViewer.Visibility = Visibility.Collapsed;
        AssumptionLibraryScrollViewer.Visibility = Visibility.Collapsed;
        ReportsAnalyticsScrollViewer.Visibility = Visibility.Collapsed;
        BenchmarkingDashboardScrollViewer.Visibility = Visibility.Collapsed;
        PlanDesignIntelligenceScrollViewer.Visibility = Visibility.Collapsed;
        PlanDesignFactorsScrollViewer.Visibility = Visibility.Collapsed;
        DocumentIntelligenceScrollViewer.Visibility = Visibility.Collapsed;
        ClientsScrollViewer.Visibility = Visibility.Collapsed;
        NegotiationWorkbenchScrollViewer.Visibility = Visibility.Collapsed;
        DataSourcesScrollViewer.Visibility = Visibility.Collapsed;
        ClaimsExperienceDashboardScrollViewer.Visibility = Visibility.Collapsed;
        ProposalComparisonScrollViewer.Visibility = Visibility.Collapsed;
        ProposalPipelineOverviewScrollViewer.Visibility = Visibility.Visible;
        ProposalPipelineOverviewScrollViewer.ScrollToTop();
        SetOperationalFooterVisible(true);
        SetStatus("Opened Proposals / Pipeline overview from saved proposal data.");
    }

    private void UpdateProposalPipelineOverviewMetrics(ProposalPipelineOverview overview)
    {
        ProposalPipelineTotalStagesTextBlock.Text = overview.TotalStages.ToString("#,0", CultureInfo.InvariantCulture);
        ProposalPipelineOpenItemsTextBlock.Text = overview.OpenItems.ToString("#,0", CultureInfo.InvariantCulture);
        ProposalPipelinePremiumTextBlock.Text = overview.TotalPremium.ToString("$#,0", CultureInfo.InvariantCulture);
        ProposalPipelineOverdueTextBlock.Text = overview.OverdueItems.ToString("#,0", CultureInfo.InvariantCulture);
    }

    private async Task ShowProposalComparisonWorkspaceAsync()
    {
        HomeDashboardScrollViewer.Visibility = Visibility.Collapsed;
        DashboardScrollViewer.Visibility = Visibility.Collapsed;
        UnderwritingWorkbenchScrollViewer.Visibility = Visibility.Collapsed;
        AssumptionLibraryScrollViewer.Visibility = Visibility.Collapsed;
        ReportsAnalyticsScrollViewer.Visibility = Visibility.Collapsed;
        BenchmarkingDashboardScrollViewer.Visibility = Visibility.Collapsed;
        PlanDesignIntelligenceScrollViewer.Visibility = Visibility.Collapsed;
        PlanDesignFactorsScrollViewer.Visibility = Visibility.Collapsed;
        DocumentIntelligenceScrollViewer.Visibility = Visibility.Collapsed;
        ClientsScrollViewer.Visibility = Visibility.Collapsed;
        NegotiationWorkbenchScrollViewer.Visibility = Visibility.Collapsed;
        DataSourcesScrollViewer.Visibility = Visibility.Collapsed;
        ClaimsExperienceDashboardScrollViewer.Visibility = Visibility.Collapsed;
        ProposalPipelineOverviewScrollViewer.Visibility = Visibility.Collapsed;
        ProposalComparisonScrollViewer.Visibility = Visibility.Visible;
        ProposalComparisonScrollViewer.ScrollToTop();
        SetOperationalFooterVisible(true);

        var selectedRenewal = RenewalQueueGrid.SelectedItem as RenewalQueueItem;
        var clientId = ResolvePlanDesignClientId(selectedRenewal);
        var clientName = FirstNonBlank(_selectedClientWorkspace?.ClientName, selectedRenewal?.Client);
        var carrier = FirstNonBlank(_selectedClientWorkspace?.PrimaryCarrier, selectedRenewal?.Carrier);
        await ProposalComparisonWorkspace.LoadClientAsync(clientId, clientName, carrier);
        SetStatus(clientId is > 0
            ? $"Opened Current vs Quote plan design for {FirstNonBlank(clientName, $"Client #{clientId.Value}")}."
            : "Opened Current vs Quote plan design. Select a saved client to load comparison data.");
    }

    private async void ProposalComparison_RunExistingComparisonRequested(object sender, EventArgs e)
    {
        await RunExistingCarrierProposalComparisonAsync();
    }

    private async Task RunExistingCarrierProposalComparisonAsync()
    {
        var proposalFolder = PromptForFolder(
            "Choose the folder containing the carrier proposal files. Subfolders will be included.",
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
        if (string.IsNullOrWhiteSpace(proposalFolder))
        {
            return;
        }

        var bookletPath = PromptForOptionalBookletPath();
        var outputFolder = ResolveProposalOutputFolder(proposalFolder);
        var databasePath = ResolveProposalDatabasePath();
        var folderClientName = Path.GetFileName(proposalFolder.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
        var selectedRenewal = RenewalQueueGrid.SelectedItem as RenewalQueueItem;
        var selectedClientId = ResolvePlanDesignClientId(selectedRenewal);
        var selectedClientName = FirstNonBlank(_selectedClientWorkspace?.ClientName, selectedRenewal?.Client, folderClientName);
        var persistenceClientId = selectedClientId?.ToString(CultureInfo.InvariantCulture) ?? selectedClientName;

        try
        {
            SetBusy(true, "Comparing carrier proposals...");
            SaveSettings();
            var progress = new Progress<string>(SetStatus);
            var result = await _proposalEngineRunner.RunComparisonAsync(
                [proposalFolder],
                persistenceClientId,
                bookletPath,
                outputFolder,
                databasePath,
                progress,
                selectedClientName);

            ShowProposalEngineResult(result, "Carrier proposal comparison");
            if (result.Succeeded)
            {
                await ProposalComparisonWorkspace.RefreshAsync();
            }
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            System.Windows.MessageBox.Show(this, ex.Message, "Carrier proposal comparison failed", MessageBoxButton.OK, MessageBoxImage.Error);
            SetStatus("Carrier proposal comparison failed.");
        }
        finally
        {
            SetBusy(false);
        }
    }

    private async void TestCarrierProposal_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Proposal files (*.pdf;*.xlsx;*.xlsm;*.xls;*.csv)|*.pdf;*.xlsx;*.xlsm;*.xls;*.csv|PDF files (*.pdf)|*.pdf|Excel and CSV files (*.xlsx;*.xlsm;*.xls;*.csv)|*.xlsx;*.xlsm;*.xls;*.csv|All files (*.*)|*.*",
            CheckFileExists = true,
            Multiselect = false
        };

        if (dialog.ShowDialog() != true)
        {
            return;
        }

        var outputFolder = ResolveProposalOutputFolder();
        var databasePath = ResolveProposalDatabasePath();

        try
        {
            SetBusy(true, "Testing carrier proposal extraction...");
            SaveSettings();
            var progress = new Progress<string>(SetStatus);
            var result = await _proposalEngineRunner.RunTestAsync(dialog.FileName, outputFolder, databasePath, progress);
            ShowProposalEngineResult(result, "Carrier proposal test");
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            System.Windows.MessageBox.Show(this, ex.Message, "Carrier proposal test failed", MessageBoxButton.OK, MessageBoxImage.Error);
            SetStatus("Carrier proposal test failed.");
        }
        finally
        {
            SetBusy(false);
        }
    }

    private string? PromptForFolder(string description, string? selectedPath = null)
    {
        using var dialog = new Forms.FolderBrowserDialog
        {
            Description = description,
            UseDescriptionForTitle = true,
            SelectedPath = Directory.Exists(selectedPath)
                ? selectedPath
                : Directory.Exists(OutputFolderTextBox.Text)
                    ? OutputFolderTextBox.Text
                    : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        };

        return dialog.ShowDialog() == Forms.DialogResult.OK ? dialog.SelectedPath : null;
    }

    private IReadOnlyList<string> PromptForProposalComparisonInputs(string clientFolder)
    {
        var choice = System.Windows.MessageBox.Show(
            this,
            "Choose a proposal folder? Select No to choose individual proposal files instead.",
            "Compare Carrier Proposals",
            MessageBoxButton.YesNoCancel,
            MessageBoxImage.Question);

        if (choice == MessageBoxResult.Cancel)
        {
            return [];
        }

        if (choice == MessageBoxResult.Yes)
        {
            using var folderDialog = new Forms.FolderBrowserDialog
            {
                Description = "Choose the folder containing carrier proposal PDFs or Excel/Word quote packages",
                UseDescriptionForTitle = true,
                SelectedPath = Directory.Exists(clientFolder)
                    ? clientFolder
                    : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            };

            return folderDialog.ShowDialog() == Forms.DialogResult.OK
                ? [folderDialog.SelectedPath]
                : [];
        }

        var fileDialog = new OpenFileDialog
        {
            Filter = "Proposal files (*.pdf;*.xlsx;*.xlsm;*.docx)|*.pdf;*.xlsx;*.xlsm;*.docx|PDF files (*.pdf)|*.pdf|Excel workbooks (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|Word proposal files (*.docx)|*.docx|All files (*.*)|*.*",
            CheckFileExists = true,
            Multiselect = true,
            InitialDirectory = Directory.Exists(clientFolder)
                ? clientFolder
                : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        };

        return fileDialog.ShowDialog() == true ? fileDialog.FileNames.ToList() : [];
    }

    private string? PromptForOptionalBookletPath()
    {
        var choice = System.Windows.MessageBox.Show(
            this,
            "Do you want to include a benefits booklet PDF for plan design validation?",
            "Benefits Booklet",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question);

        if (choice != MessageBoxResult.Yes)
        {
            return null;
        }

        var dialog = new OpenFileDialog
        {
            Filter = "Benefit booklet PDFs (*.pdf)|*.pdf|All files (*.*)|*.*",
            CheckFileExists = true,
            Multiselect = false
        };

        return dialog.ShowDialog() == true ? dialog.FileName : null;
    }

    private string ResolveProposalOutputFolder(string? proposalSourceFolder = null)
    {
        var outputRoot = ResolveOutputRootFolder();

        var runName = SafeFolderName(Path.GetFileName((proposalSourceFolder ?? "").TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)));
        if (string.IsNullOrWhiteSpace(runName))
        {
            runName = "Proposal_Comparison";
        }

        var outputFolder = Path.Combine(
            outputRoot,
            "Carrier Proposal Comparison",
            $"{runName}_{DateTime.Now:yyyyMMdd_HHmmss}");
        Directory.CreateDirectory(outputFolder);
        return outputFolder;
    }

    private string ResolveOutputRootFolder()
    {
        var outputRoot = OutputFolderTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(outputRoot))
        {
            var documents = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            outputRoot = Path.Combine(
                documents,
                "Renewal Review Builder",
                "Outputs");
            OutputFolderTextBox.Text = outputRoot;
        }

        Directory.CreateDirectory(outputRoot);
        return outputRoot;
    }

    private static string SafeFolderName(string value)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var cleaned = new string(value.Select(character => invalid.Contains(character) ? '_' : character).ToArray()).Trim();
        return string.IsNullOrWhiteSpace(cleaned) ? "" : cleaned;
    }

    private static string ResolveProposalDatabasePath()
    {
        AppPaths.EnsureCreated();
        return AppPaths.BenefitsIntelligenceDatabasePath;
    }

    private void ShowProposalEngineResult(ProposalEngineRunResult result, string title)
    {
        _proposalOutputFolder = result.OutputFolder;
        ExportTextBlock.Text =
            $"{title}:{Environment.NewLine}{result.Message}{Environment.NewLine}{Environment.NewLine}" +
            $"Output folder: {result.OutputFolder}{Environment.NewLine}" +
            $"SQLite database: {result.DatabasePath}";
        SetOpenOutputButtonsEnabled(result.Succeeded && Directory.Exists(result.OutputFolder));
        SetStatus(result.Succeeded
            ? $"{title} complete."
            : $"{title} needs review. See exported details and status messages.");

        var icon = result.Succeeded ? MessageBoxImage.Information : MessageBoxImage.Warning;
        System.Windows.MessageBox.Show(this, result.Message, title, MessageBoxButton.OK, icon);
        if (result.Succeeded && Directory.Exists(result.OutputFolder))
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = result.OutputFolder,
                UseShellExecute = true
            });
        }
    }

    private void NewClient_Click(object sender, RoutedEventArgs e)
    {
        ShowPlaceholder("New Client");
    }

    private void ViewRenewalPipeline_Click(object sender, RoutedEventArgs e)
    {
        ShowPlaceholder("Renewal Pipeline");
    }

    private async void QuickReport_Click(object sender, RoutedEventArgs e)
    {
        await ShowReportsAnalyticsAsync();
        SetStatus($"{(sender as Button)?.Content}: opened Reports & Analytics Command Centre.");
    }

    private void AskAssistant_Click(object sender, RoutedEventArgs e)
    {
        ShowPlaceholder("AI Assistant (Copilot)");
    }

    private void OpenHelpSupport_Click(object sender, RoutedEventArgs e)
    {
        ShowPlaceholder(
            "Help & Support",
            "Help & Support will contain user guides, workflow support, release notes, and escalation contacts for Benefits Intelligence OS.");
    }

    private void ShowPlaceholder(string featureName)
    {
        ShowPlaceholder(featureName, $"{featureName} is set up on the dashboard and can be connected to a full workflow next.");
    }

    private void ShowPlaceholder(string featureName, string description)
    {
        SetStatus($"{featureName} is ready as a dashboard placeholder.");
        System.Windows.MessageBox.Show(
            this,
            description,
            featureName,
            MessageBoxButton.OK,
            MessageBoxImage.Information);
    }

    private static bool IsPdfFile(string path)
    {
        return string.Equals(Path.GetExtension(path), ".pdf", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsWorkbookFile(string path)
    {
        var extension = Path.GetExtension(path);
        return extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) ||
               extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase) ||
               extension.Equals(".xls", StringComparison.OrdinalIgnoreCase);
    }

    private static void AddDistinctPath(ObservableCollection<string> paths, string path, string primaryPath = "")
    {
        if (string.IsNullOrWhiteSpace(path) ||
            path.Equals(primaryPath, StringComparison.OrdinalIgnoreCase) ||
            paths.Any(existing => existing.Equals(path, StringComparison.OrdinalIgnoreCase)))
        {
            return;
        }

        paths.Add(path);
    }

    private static void BrowseFile(System.Windows.Controls.TextBox textBox, string filter)
    {
        var dialog = new OpenFileDialog
        {
            Filter = filter,
            CheckFileExists = true
        };

        if (dialog.ShowDialog() == true)
        {
            textBox.Text = dialog.FileName;
        }
    }

    private void WireInputChangeHandlers()
    {
        PdfPathTextBox.TextChanged += Input_Changed;
        WorkbookPathTextBox.TextChanged += Input_Changed;
        TemplatePathTextBox.TextChanged += Input_Changed;
        OutputFolderTextBox.TextChanged += OutputDestination_Changed;
        CurrentYearFolderTextBox.TextChanged += Input_Changed;
        PriorYearFolderTextBox.TextChanged += Input_Changed;
        AsoBillingFolderTextBox.TextChanged += Input_Changed;
    }

    private void ReportMonthCount_Changed(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        InvalidateAnalysis();
    }

    private void Input_Changed(object sender, System.Windows.Controls.TextChangedEventArgs e)
    {
        InvalidateAnalysis();
    }

    private void OutputDestination_Changed(object sender, System.Windows.Controls.TextChangedEventArgs e)
    {
        if (_isLoadingDefaults)
        {
            return;
        }

        _export = null;
        SetOpenOutputButtonsEnabled(false);
        if (SharePointWebDavPathPolicy.IsDirectSharePointWebDavPath(OutputFolderTextBox.Text))
        {
            SetStatus("SharePoint destination selected. The review will be generated locally first, then published to SharePoint.");
            return;
        }

        SetStatus(_analysis is null
            ? "Output folder changed. Click Validate, then Export Pack."
            : "Output folder changed. The validated source review is retained for Export Pack.");
    }

    private void InvalidateAnalysis()
    {
        if (_isLoadingDefaults)
        {
            return;
        }

        _analysis = null;
        _analysisClientId = null;
        _analysisSourceDocumentId = null;
        _export = null;
        _claimsDataModelExport = null;
        _claimsRangeReviewViewModel.ResetSession();
        SetOpenOutputButtonsEnabled(false);
        SetStatus("Files changed. Click Validate, then Export Pack.");
    }

    private async void Validate_Click(object sender, RoutedEventArgs e)
    {
        await AnalyzeAsync();
    }

    private async void Export_Click(object sender, RoutedEventArgs e)
    {
        await ExportAnalysisAsync();
    }

    private async Task<bool> ExportAnalysisAsync(
        string? requestedVersionName = null,
        NegotiationWorkbenchSaveRequest? savedRequest = null,
        bool useLoadedAnalysis = false)
    {
        if (!useLoadedAnalysis)
        {
            var renewalCase = CreateRenewalCase();
            var sourceInputsChanged = _analysis is not null &&
                                      (!_analysis.MatchesExportInputs(renewalCase) ||
                                       ReviewPackBuilder.SourceInputsChangedSince(renewalCase, _analysis.CreatedAt));
            if (_analysis is null || sourceInputsChanged)
            {
                if (sourceInputsChanged)
                {
                    SetStatus("Source files changed since validation. Refreshing analysis before export...");
                }

                await AnalyzeAsync();
            }
        }

        if (_analysis is null)
        {
            return false;
        }

        if (!string.IsNullOrWhiteSpace(requestedVersionName))
        {
            if (requestedVersionName is RenewalVersionNames.CarrierProposed or RenewalVersionNames.ConsultantModel)
            {
                _analysis.SelectedRenewalVersionName = requestedVersionName;
            }
            else if (savedRequest is not null)
            {
                _negotiationWorkbenchService.ApplySavedVersion(_analysis, savedRequest, requestedVersionName);
            }
            else
            {
                ApplyWorkbenchVersionToAnalysis(requestedVersionName);
            }

            UpdateReview();
        }

        ApplyAssurePlusCommentaryEdits();
        _claimsRangeReviewViewModel.ApplySessionEdits();

        try
        {
            SetBusy(true, "Exporting review pack...");
            SaveSettings();
            var selectedOutputFolder = _analysis.RenewalCase.OutputFolder;
            var isDirectSharePointOutput = SharePointWebDavPathPolicy.IsDirectSharePointWebDavPath(selectedOutputFolder);
            string? localSafetyFolder = null;
            bool? sharePointPublishSucceeded = null;

            try
            {
                if (isDirectSharePointOutput)
                {
                    localSafetyFolder = SharePointWebDavPathPolicy.CreateLocalStagingFolder(
                        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                        DateTime.Now);
                    _analysis.RenewalCase.OutputFolder = localSafetyFolder;
                    SetStatus("Generating the review locally before publishing it to SharePoint...");
                }

                _export = await _builder.ExportAsync(_analysis, new Progress<string>(SetStatus));
            }
            finally
            {
                _analysis.RenewalCase.OutputFolder = selectedOutputFolder;
            }

            if (isDirectSharePointOutput && _export is not null && localSafetyFolder is not null)
            {
                try
                {
                    var localExportFolder = _export.OutputFolder;
                    var publish = await _sharePointExportPublisher.PublishAsync(
                        localExportFolder,
                        selectedOutputFolder,
                        new Progress<string>(SetStatus));

                    _export.WorkbookPath = SharePointStagedExportPublisher.RebasePublishedPath(
                        _export.WorkbookPath,
                        localExportFolder,
                        publish.DestinationFolder);
                    _export.PowerPointPath = SharePointStagedExportPublisher.RebasePublishedPath(
                        _export.PowerPointPath,
                        localExportFolder,
                        publish.DestinationFolder);
                    _export.SupportingPdfFolder = SharePointStagedExportPublisher.RebasePublishedPath(
                        _export.SupportingPdfFolder,
                        localExportFolder,
                        publish.DestinationFolder);
                    _export.SupportingWorkbookFolder = SharePointStagedExportPublisher.RebasePublishedPath(
                        _export.SupportingWorkbookFolder,
                        localExportFolder,
                        publish.DestinationFolder);
                    _export.OutputFolder = publish.DestinationFolder;
                    _export.DeckMessage =
                        $"{_export.DeckMessage}{Environment.NewLine}" +
                        $"Published {publish.PublishedFileCount} file(s) to SharePoint.{Environment.NewLine}" +
                        $"Local safety copy: {localExportFolder}";
                    sharePointPublishSucceeded = true;
                }
                catch (Exception publishException)
                {
                    AppLogger.Error(publishException);
                    _export.DeckMessage =
                        $"{_export.DeckMessage}{Environment.NewLine}" +
                        $"SharePoint upload needs review: {publishException.Message}{Environment.NewLine}" +
                        $"The completed local safety copy is available at: {_export.OutputFolder}";
                    sharePointPublishSucceeded = false;
                }
            }

            var completedExport = _export ?? throw new InvalidOperationException("The review export did not return an output result.");
            var supportingPdfText = completedExport.SupportingPdfCount > 0
                ? $"{Environment.NewLine}Supporting PDFs: {completedExport.SupportingPdfCount} copied to {completedExport.SupportingPdfFolder}{Environment.NewLine}"
                : "";
            var supportingWorkbookText = completedExport.SupportingWorkbookCount > 0
                ? $"Supporting Excel files: {completedExport.SupportingWorkbookCount} copied to {completedExport.SupportingWorkbookFolder}{Environment.NewLine}"
                : "";
            ExportTextBlock.Text =
                $"Workbook: {completedExport.WorkbookPath}{Environment.NewLine}" +
                $"PowerPoint: {completedExport.PowerPointPath}{Environment.NewLine}" +
                supportingPdfText +
                supportingWorkbookText +
                $"{completedExport.DeckMessage}";
            RefreshSourceChecksGrid();
            SetOpenOutputButtonsEnabled(IsOutputFolderAvailableForUi(completedExport.OutputFolder));
            if (!useLoadedAnalysis)
            {
                UpsertCurrentRenewalQueueItem(completed: sharePointPublishSucceeded != false);
            }
            SetStatus(sharePointPublishSucceeded == false
                ? "Review and PowerPoint are ready locally. SharePoint upload needs review."
                : sharePointPublishSucceeded == true
                    ? "Review and PowerPoint are ready in SharePoint."
                    : "Review and PowerPoint ready.");
            return true;
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            System.Windows.MessageBox.Show(this, ex.Message, "Export failed", MessageBoxButton.OK, MessageBoxImage.Error);
            SetStatus("Export failed.");
            return false;
        }
        finally
        {
            SetBusy(false);
        }
    }

    private async Task<ClaimsExperienceDataModelRequest> CreateClaimsDataModelRequestAsync(
        bool includeBookletSchedule = true)
    {
        var bookletPaths = includeBookletSchedule
            ? ResolveClaimsBenefitBookletPaths()
            : new List<string>();
        var bookletMasterPath = includeBookletSchedule
            ? await EnsureBookletMasterForClaimsAsync(bookletPaths)
            : "";
        var selectedRenewal = RenewalQueueGrid.SelectedItem as RenewalQueueItem;
        var clientId = ResolvePlanDesignClientId(selectedRenewal);
        var carrier = FirstNonBlank(_selectedClientWorkspace?.PrimaryCarrier, selectedRenewal?.Carrier);
        return new ClaimsExperienceDataModelRequest
        {
            SourceFolder = ResolveClaimsDataModelSourceFolder(),
            SourceFiles = ResolveClaimsDataModelSourceFiles(),
            OutputFolder = OutputFolderTextBox.Text.Trim(),
            ReportMonthCount = SelectedReportMonthCount(),
            BookletMasterWorkbookPath = bookletMasterPath,
            BenefitBookletPaths = bookletPaths,
            RenewalTerms = await LoadClaimsRenewalTermsAsync(clientId, carrier)
        };
    }

    private async Task<ClaimsExperienceRenewalTerms> LoadClaimsRenewalTermsAsync(
        long? clientId,
        string? carrier)
    {
        if (clientId is not > 0)
        {
            return new ClaimsExperienceRenewalTerms();
        }

        var assumptions = await _databaseService.GetRenewalAssumptionsAsync(clientId.Value);
        return ClaimsExperienceRenewalTerms.FromRenewalAssumptions(assumptions, carrier);
    }

    private async Task<string> EnsureBookletMasterForClaimsAsync(IReadOnlyList<string> bookletPaths)
    {
        if (bookletPaths.Count == 0)
        {
            _bookletExtractionInputSnapshot.Clear();
            return "";
        }

        if (_bookletExtractionResult is not null &&
            File.Exists(_bookletExtractionResult.MasterWorkbookPath) &&
            SamePaths(_bookletExtractionInputSnapshot, bookletPaths))
        {
            return _bookletExtractionResult.MasterWorkbookPath;
        }

        var outputFolder = OutputFolderTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(outputFolder))
        {
            var documents = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            outputFolder = Path.Combine(documents, "Renewal Review Builder", "Outputs");
            OutputFolderTextBox.Text = outputFolder;
        }

        SetStatus("Extracting benefit booklet schedule for claims export...");
        _bookletExtractionCancellation = new CancellationTokenSource();
        CancelBookletExtractionButton.IsEnabled = true;

        try
        {
            var progress = new Progress<BookletExtractionProgress>(status => SetStatus(status.Message));
            _bookletExtractionResult = await _bookletExtractorRunner.RunAsync(
                new BookletExtractionRequest
                {
                    InputPaths = bookletPaths.ToList(),
                    OutputFolder = outputFolder
                },
                progress,
                _bookletExtractionCancellation.Token);
            _bookletExtractionInputSnapshot = bookletPaths.ToList();

            var databaseSave = await _databaseService.SavePlanDesignFromBookletExtractionAsync(
                _bookletExtractionResult,
                bookletPaths,
                ResolvePlanDesignClientId(RenewalQueueGrid.SelectedItem as RenewalQueueItem));

            SetStatus(databaseSave.Succeeded
                ? "Benefit booklet schedule extracted for claims export."
                : "Benefit booklet schedule extracted. Database save needs review.");
            return _bookletExtractionResult.MasterWorkbookPath;
        }
        finally
        {
            _bookletExtractionCancellation?.Dispose();
            _bookletExtractionCancellation = null;
            CancelBookletExtractionButton.IsEnabled = false;
        }
    }

    private async void ExportClaimsDataModel_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            SetBusy(true, "Exporting claims experience data model...");
            SaveSettings();

            _claimsDataModelExport = await _claimsDataModelBuilder.ExportAsync(await CreateClaimsDataModelRequestAsync());
            var databaseSave = await SaveClaimsDataModelToDatabaseAsync();

            var warningText = _claimsDataModelExport.Warnings.Count > 0
                ? $"{Environment.NewLine}Warnings: {string.Join(" ", _claimsDataModelExport.Warnings)}"
                : "";
            var scheduleText = !string.IsNullOrWhiteSpace(_claimsDataModelExport.ScheduleOfBenefitsSourcePath)
                ? $"{Environment.NewLine}Schedule of benefits source: {_claimsDataModelExport.ScheduleOfBenefitsSourcePath}"
                : "";

            ExportTextBlock.Text =
                $"Claims data model: {_claimsDataModelExport.WorkbookPath}{Environment.NewLine}" +
                $"Tables created: {_claimsDataModelExport.TableCount}{Environment.NewLine}" +
                $"Source rows read: {_claimsDataModelExport.SourceRowCount}" +
                scheduleText +
                warningText +
                DatabaseStatusLine(databaseSave);

            SetOpenOutputButtonsEnabled(Directory.Exists(_claimsDataModelExport.OutputFolder));
            SetStatus(databaseSave?.Succeeded == true
                ? "Claims experience data model exported and saved to the Benefits Intelligence database."
                : "Claims experience data model exported. Database save needs review.");
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            System.Windows.MessageBox.Show(this, ex.Message, "Claims data model export failed", MessageBoxButton.OK, MessageBoxImage.Error);
            SetStatus("Claims data model export failed.");
        }
        finally
        {
            SetBusy(false);
        }
    }

    private async void ExportClaimsPowerPoint_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            SetBusy(true, "Exporting claims experience PowerPoint...");
            SaveSettings();
            DatabaseSaveResult? databaseSave = null;
            var claimsRequest = await CreateClaimsDataModelRequestAsync();

            if (_claimsDataModelExport is null ||
                !File.Exists(_claimsDataModelExport.WorkbookPath) ||
                !string.IsNullOrWhiteSpace(claimsRequest.BookletMasterWorkbookPath))
            {
                SetStatus("Creating claims experience data model first...");
                _claimsDataModelExport = await _claimsDataModelBuilder.ExportAsync(claimsRequest);
                databaseSave = await SaveClaimsDataModelToDatabaseAsync();
            }

            var templatePath = ResolveClaimsPowerPointTemplatePath();
            _claimsPowerPointExport = await _claimsPowerPointExporter.ExportAsync(new ClaimsExperiencePowerPointRequest
            {
                WorkbookPath = _claimsDataModelExport.WorkbookPath,
                TemplatePath = templatePath,
                OutputFolder = OutputFolderTextBox.Text.Trim()
            });

            var warningText = _claimsPowerPointExport.Warnings.Count > 0
                ? $"{Environment.NewLine}Warnings: {string.Join(" ", _claimsPowerPointExport.Warnings)}"
                : "";

            ExportTextBlock.Text =
                $"Claims PowerPoint: {_claimsPowerPointExport.PowerPointPath}{Environment.NewLine}" +
                $"Claims data model: {_claimsPowerPointExport.SourceWorkbookPath}{Environment.NewLine}" +
                $"Slides updated: {_claimsPowerPointExport.SlideCountUpdated}" +
                warningText +
                DatabaseStatusLine(databaseSave);

            SetOpenOutputButtonsEnabled(Directory.Exists(_claimsPowerPointExport.OutputFolder));
            SetStatus("Claims experience PowerPoint exported.");
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            System.Windows.MessageBox.Show(this, ex.Message, "Claims PowerPoint export failed", MessageBoxButton.OK, MessageBoxImage.Error);
            SetStatus("Claims PowerPoint export failed.");
        }
        finally
        {
            SetBusy(false);
        }
    }

    private async Task AnalyzeAsync()
    {
        try
        {
            SetBusy(true, "Validating files...");
            _claimsRangeReviewViewModel.ResetSession();
            SaveSettings();
            _analysis = await _builder.AnalyzeAsync(CreateRenewalCase(), new Progress<string>(SetStatus));
            var databaseSave = await _databaseService.SaveRenewalAnalysisAsync(_analysis);
            SetStatus("Refreshing monthly claims history...");
            var claimsDatabaseSave = await SaveClaimsExperienceDuringValidationAsync(_analysis);
            _latestSavedClientId = databaseSave.ClientId;
            _latestSavedDocumentId = databaseSave.DocumentId;
            _analysisClientId = databaseSave.ClientId;
            _analysisSourceDocumentId = databaseSave.DocumentId;
            UpdateReview();
            UpsertCurrentRenewalQueueItem(completed: false);
            if (databaseSave.Succeeded && databaseSave.ClientId is > 0)
            {
                await RefreshClientWorkspaceAfterRenewalSaveAsync(_analysis, databaseSave.ClientId.Value);
            }
            var claimsStatus = claimsDatabaseSave switch
            {
                { Succeeded: true } => " Monthly claims history was refreshed.",
                { Succeeded: false } => " Monthly claims history needs review.",
                _ => ""
            };
            SetStatus(databaseSave.Succeeded
                ? $"Validation complete. Renewal saved to the Benefits Intelligence database.{claimsStatus}"
                : $"Validation complete. Database save needs review.{claimsStatus}");
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            System.Windows.MessageBox.Show(this, ex.Message, "Validation failed", MessageBoxButton.OK, MessageBoxImage.Error);
            SetStatus("Validation failed.");
        }
        finally
        {
            SetBusy(false);
        }
    }

    private async Task RefreshClientWorkspaceAfterRenewalSaveAsync(AnalysisResult analysis, long clientId)
    {
        var selectedWorkspace = _selectedClientWorkspace;
        var resolvedClientName = ClientNameResolver.Resolve(
            analysis.RenewalCase.ClientName,
            analysis.Pdf.ClientName,
            analysis.RenewalCase.PdfPath);
        var policyNumber = FirstNonBlank(analysis.Pdf.PolicyNumber, analysis.RenewalCase.PolicyNumber);
        var selectedWorkspaceMatches = selectedWorkspace is not null &&
            (selectedWorkspace.DatabaseClientId == clientId ||
             NormalizeClientMatchKey(selectedWorkspace.ClientName)
                 .Equals(NormalizeClientMatchKey(resolvedClientName), StringComparison.OrdinalIgnoreCase) ||
             (!string.IsNullOrWhiteSpace(policyNumber) &&
              selectedWorkspace.ClientId.Equals(policyNumber, StringComparison.OrdinalIgnoreCase)));

        await LoadSavedClientListAsync();
        if (!selectedWorkspaceMatches)
        {
            return;
        }

        var refreshedClient = _clientList.FirstOrDefault(item => item.DatabaseClientId == clientId);
        if (refreshedClient is null)
        {
            return;
        }

        _selectedClientWorkspace = BuildClientWorkspace(refreshedClient);
        _clientRenewalSelectionIdentity = CreateClientRenewalSelectionIdentity(_selectedClientWorkspace);
        if (ClientClaimsExperienceDashboard.Visibility == Visibility.Visible)
        {
            await ShowClientClaimsWorkspaceAsync();
        }
    }

    private RenewalCase CreateRenewalCase()
    {
        return new RenewalCase
        {
            PdfPath = PdfPathTextBox.Text.Trim(),
            AdditionalPdfPaths = _additionalPdfPaths.ToList(),
            WorkbookPath = WorkbookPathTextBox.Text.Trim(),
            AdditionalWorkbookPaths = _additionalWorkbookPaths.ToList(),
            CurrentYearFolder = CurrentYearFolderTextBox.Text.Trim(),
            PriorYearFolder = PriorYearFolderTextBox.Text.Trim(),
            AsoBillingFolder = AsoBillingFolderTextBox.Text.Trim(),
            HcsaCalculation = _hcsaCalculation.Clone(),
            AllstateCriticalIllnessEnrollment = _allstateCriticalIllnessEnrollment.Clone(),
            PowerPointTemplatePath = TemplatePathTextBox.Text.Trim(),
            OutputFolder = OutputFolderTextBox.Text.Trim()
        };
    }

    private void UpsertCurrentRenewalQueueItem(bool completed)
    {
        if (_analysis is null)
        {
            return;
        }

        var current = CreateRenewalQueueItem(_analysis, completed);
        var staleCurrentRows = _renewalQueue
            .Where(row => row.Id.Equals("current-renewal", StringComparison.OrdinalIgnoreCase))
            .ToList();
        foreach (var staleCurrentRow in staleCurrentRows)
        {
            _renewalQueue.Remove(staleCurrentRow);
        }

        RemoveVisibleRenewalDuplicates(current);
        _renewalQueue.Insert(0, current);
        RefreshRenewalQueueFilters();
        _renewalQueueView?.Refresh();
        UpdateRenewalQueueSummary();
        RenewalQueueGrid.SelectedItem = current;
    }

    private void RefreshRenewalQueueFilters()
    {
        var selectedStatus = SelectedFilterValue(RenewalStatusFilterComboBox, "All statuses");
        var selectedCarrier = SelectedFilterValue(RenewalCarrierFilterComboBox, "All carriers");
        var selectedMonth = SelectedFilterValue(RenewalMonthFilterComboBox, "All months");

        ConfigureFilterCombo(RenewalStatusFilterComboBox, "All statuses", _renewalQueue.Select(row => row.Status).Distinct().OrderBy(StatusRank));
        ConfigureFilterCombo(RenewalCarrierFilterComboBox, "All carriers", _renewalQueue.Select(row => row.Carrier).Distinct().OrderBy(value => value));
        ConfigureFilterCombo(RenewalMonthFilterComboBox, "All months", _renewalQueue.Select(row => row.RenewalMonth).Distinct().OrderBy(MonthRank));

        RestoreFilterSelection(RenewalStatusFilterComboBox, selectedStatus);
        RestoreFilterSelection(RenewalCarrierFilterComboBox, selectedCarrier);
        RestoreFilterSelection(RenewalMonthFilterComboBox, selectedMonth);
    }

    private static void RestoreFilterSelection(ComboBox comboBox, string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            comboBox.SelectedIndex = 0;
            return;
        }

        foreach (var item in comboBox.Items)
        {
            if (string.Equals(item?.ToString(), value, StringComparison.OrdinalIgnoreCase))
            {
                comboBox.SelectedItem = item;
                return;
            }
        }

        comboBox.SelectedIndex = 0;
    }

    private RenewalQueueItem CreateRenewalQueueItem(AnalysisResult analysis, bool completed)
    {
        var ehcLossRatio = LossRatioFor(analysis, ["Extended", "EHC", "Health"]);
        var dentalLossRatio = LossRatioFor(analysis, ["Dental"]);
        var status = DetermineRenewalQueueStatus(analysis, completed, ehcLossRatio);
        var client = FirstNonBlank(
            ClientNameResolver.Resolve(
                analysis.RenewalCase.ClientName,
                analysis.Pdf.ClientName,
                analysis.RenewalCase.PdfPath),
            "Current Renewal");

        return new RenewalQueueItem
        {
            Id = "current-renewal",
            ClientId = _analysisClientId ?? 0,
            SourceDocumentId = _analysisSourceDocumentId,
            SelectedRenewalVersion = RenewalVersionNames.Negotiated,
            Client = client,
            PolicyNumber = FirstNonBlank(analysis.Pdf.PolicyNumber, analysis.RenewalCase.PolicyNumber),
            Carrier = DetectCarrier(analysis),
            RenewalDate = analysis.RenewalCase.RenewalDate ?? analysis.Pdf.RenewalDate,
            DateLoaded = DateTime.Now,
            CurrentPremium = analysis.CurrentMonthlyPremium * 12m,
            RenewalPremium = analysis.RenewalMonthlyPremium * 12m,
            PremiumChangePercent = analysis.MonthlyPremiumChangePercent,
            EhcLossRatio = ehcLossRatio,
            TargetLossRatio = FindRenewalExperienceTargetLossRatio(analysis),
            PoolingThreshold = analysis.Pdf.PoolingThreshold,
            PoolingCharge = analysis.Pdf.PoolingCharge,
            DentalLossRatio = dentalLossRatio,
            Lives = ResolveLives(analysis),
            TrendAssumptions = analysis.Workbook.ClaimTrends.Count > 0 ? "Workbook trend data read" : "Trend data not read",
            PoolingCharges = BuildPoolingSummary(analysis.Pdf.PoolingThreshold, analysis.Pdf.PoolingCharge),
            Status = status,
            ActionLabel = ActionForStatus(status),
            QuickInsights = BuildRenewalInsights(analysis.MonthlyPremiumChangePercent, ehcLossRatio, dentalLossRatio),
            MissingDataItems = MissingDataItems(analysis)
        };
    }

    private static RenewalQueueItem CreateRenewalQueueItem(RenewalQueueRecord record)
    {
        var status = DetermineRenewalQueueStatus(record);
        var client = FirstNonBlank(
            ClientNameResolver.Resolve(record.ClientName, null, record.SourceFilePath),
            record.ClientName);
        return new RenewalQueueItem
        {
            Id = $"saved-renewal-{record.ClientId}-{record.SourceDocumentId?.ToString(CultureInfo.InvariantCulture) ?? record.RenewalDate?.DayNumber.ToString(CultureInfo.InvariantCulture) ?? "unknown"}",
            ClientId = record.ClientId,
            RenewalId = record.RenewalId,
            SourceDocumentId = record.SourceDocumentId,
            SelectedRenewalVersion = string.IsNullOrWhiteSpace(record.SelectedRenewalVersion)
                ? RenewalVersionNames.CarrierProposed
                : record.SelectedRenewalVersion,
            Client = client,
            PolicyNumber = record.PolicyNumber,
            Carrier = FirstNonBlank(record.Carrier, "Carrier"),
            RenewalDate = record.RenewalDate ?? (record.RenewalYear > 0 ? new DateOnly(record.RenewalYear, 1, 1) : null),
            DateLoaded = record.DateLoaded,
            CurrentPremium = record.CurrentPremium,
            RenewalPremium = record.RenewalPremium,
            PremiumChangePercent = record.PremiumChangePercent,
            EhcLossRatio = record.EhcLossRatio,
            PoolingThreshold = record.PoolingThreshold,
            PoolingCharge = record.PoolingCharge,
            DentalLossRatio = record.DentalLossRatio,
            Lives = record.Lives,
            TrendAssumptions = "Loaded from Benefits Intelligence history",
            PoolingCharges = BuildPoolingSummary(record.PoolingThreshold, record.PoolingCharge),
            Status = status,
            ActionLabel = ActionForStatus(status),
            QuickInsights = BuildRenewalInsights(record.PremiumChangePercent, record.EhcLossRatio, record.DentalLossRatio),
            MissingDataItems = MissingDataItems(record)
        };
    }

    private static decimal? FindRenewalExperienceTargetLossRatio(AnalysisResult analysis)
    {
        var healthTarget = analysis.Pdf.ExperienceLines
            .Where(row => row.Benefit.Contains("extended health", StringComparison.OrdinalIgnoreCase) ||
                          row.Benefit.Contains("health care", StringComparison.OrdinalIgnoreCase) ||
                          row.Benefit.Equals("EHC", StringComparison.OrdinalIgnoreCase))
            .Select(row => row.TargetLossRatio)
            .FirstOrDefault(value => value.HasValue);

        return healthTarget ?? analysis.Pdf.ExperienceLines
            .Select(row => row.TargetLossRatio)
            .FirstOrDefault(value => value.HasValue);
    }

    private static string BuildPoolingSummary(decimal? threshold, decimal? charge)
    {
        if (!threshold.HasValue && !charge.HasValue)
        {
            return "N/A";
        }

        return $"threshold {FormatMoney(threshold)}; charge {FormatPercent(NormalizePercent(charge), includeSign: false)}";
    }

    private static string DetermineRenewalQueueStatus(RenewalQueueRecord record)
    {
        if (IsKnownRenewalQueueStatus(record.Status))
        {
            return record.Status;
        }

        if (!record.CurrentPremium.HasValue || !record.RenewalPremium.HasValue)
        {
            return RenewalQueueStatuses.MissingData;
        }

        if (record.PremiumChangePercent > 10m || record.EhcLossRatio > 90m)
        {
            return RenewalQueueStatuses.HighRisk;
        }

        if (record.PremiumChangePercent is >= 5m and <= 10m || record.EhcLossRatio is >= 80m and <= 90m)
        {
            return RenewalQueueStatuses.Watch;
        }

        return RenewalQueueStatuses.ReadyToReview;
    }

    private static bool IsKnownRenewalQueueStatus(string status)
    {
        return status is RenewalQueueStatuses.Healthy
            or RenewalQueueStatuses.Watch
            or RenewalQueueStatuses.HighRisk
            or RenewalQueueStatuses.MissingData
            or RenewalQueueStatuses.ReadyToReview
            or RenewalQueueStatuses.Negotiated
            or RenewalQueueStatuses.FinalApproved
            or RenewalQueueStatuses.Completed;
    }

    private static List<string> MissingDataItems(RenewalQueueRecord record)
    {
        var missing = new List<string>();
        if (!record.CurrentPremium.HasValue)
        {
            missing.Add("Current premium");
        }

        if (!record.RenewalPremium.HasValue)
        {
            missing.Add("Renewal premium");
        }

        if (!record.EhcLossRatio.HasValue)
        {
            missing.Add("EHC loss ratio");
        }

        if (!record.DentalLossRatio.HasValue)
        {
            missing.Add("Dental loss ratio");
        }

        return missing;
    }

    private static string DetermineRenewalQueueStatus(AnalysisResult analysis, bool completed, decimal? ehcLossRatio)
    {
        if (completed)
        {
            return RenewalQueueStatuses.Completed;
        }

        if (analysis.SourceChecks.Any(check => check.Status.Equals("Missing", StringComparison.OrdinalIgnoreCase)))
        {
            return RenewalQueueStatuses.MissingData;
        }

        if (analysis.Findings.Any(finding => finding.Severity == FindingSeverity.Critical) ||
            analysis.MonthlyPremiumChangePercent > 10m ||
            ehcLossRatio > 90m)
        {
            return RenewalQueueStatuses.HighRisk;
        }

        if (analysis.Findings.Any(finding => finding.Severity == FindingSeverity.Warning) ||
            analysis.MonthlyPremiumChangePercent is >= 5m and <= 10m ||
            ehcLossRatio is >= 80m and <= 90m)
        {
            return RenewalQueueStatuses.Watch;
        }

        return RenewalQueueStatuses.ReadyToReview;
    }

    private static string ActionForStatus(string status)
    {
        return status switch
        {
            RenewalQueueStatuses.MissingData => "Needs Data",
            RenewalQueueStatuses.Completed => "Review",
            RenewalQueueStatuses.FinalApproved => "Review",
            _ => "Underwrite"
        };
    }

    private static string DetectCarrier(AnalysisResult analysis)
    {
        if (!string.IsNullOrWhiteSpace(analysis.Pdf.Carrier))
        {
            return analysis.Pdf.Carrier.Trim();
        }

        var paths = new[] { analysis.RenewalCase.PdfPath, analysis.RenewalCase.WorkbookPath };
        return paths.Any(path => path.Contains("Manulife", StringComparison.OrdinalIgnoreCase) ||
                                 path.Contains("Client_Copy", StringComparison.OrdinalIgnoreCase))
            ? "Manulife"
            : "Carrier";
    }

    private static int? ResolveLives(AnalysisResult analysis)
    {
        return analysis.Pdf.RateLines
            .Select(line => line.Lives)
            .Concat(analysis.Pdf.ExperienceLines.Select(line => line.AverageLives))
            .Where(value => value.HasValue)
            .Max();
    }

    private static decimal? LossRatioFor(AnalysisResult analysis, string[] benefitTokens)
    {
        var matches = analysis.Pdf.ExperienceLines
            .Where(item => benefitTokens.Any(token => BenefitMatcher.ContainsPhrase(item.Benefit, token)))
            .Select(item => new
            {
                Line = item,
                StartDate = ExperiencePeriodStart(item)
            })
            .OrderByDescending(item => item.StartDate ?? DateOnly.MinValue)
            .ToList();

        if (matches.Count == 0)
        {
            return null;
        }

        var newestStartDate = matches[0].StartDate;
        var newestRows = matches
            .Where(item => Nullable.Equals(item.StartDate, newestStartDate))
            .Select(item => item.Line)
            .ToList();

        var preferred = newestRows
            .OrderByDescending(row => ExperienceLinePriority(row, benefitTokens))
            .FirstOrDefault(row => ExperienceLinePriority(row, benefitTokens) >= 90);
        var preferredRatio = preferred?.PaidLossRatio ?? preferred?.IncurredLossRatio;
        if (preferredRatio.HasValue)
        {
            return NormalizePercent(preferredRatio);
        }

        var combinedRatio = CombinedLossRatio(newestRows);
        if (combinedRatio.HasValue)
        {
            return NormalizePercent(combinedRatio);
        }

        var line = newestRows
            .OrderByDescending(row => ExperienceLinePriority(row, benefitTokens))
            .FirstOrDefault();

        return NormalizePercent(line?.PaidLossRatio ?? line?.IncurredLossRatio);
    }

    private static DateOnly? ExperiencePeriodStart(ExperienceLine line)
    {
        var periodStart = line.Period.Split(" - ", StringSplitOptions.None).FirstOrDefault();
        return ParsingHelpers.ParseDateOnly(periodStart);
    }

    private static int ExperienceLinePriority(ExperienceLine line, string[] benefitTokens)
    {
        if (BenefitMatcher.ContainsPhrase(line.Benefit, "subtotal"))
        {
            return 100;
        }

        if (benefitTokens.Any(token => BenefitMatcher.NormalizeWords(line.Benefit).Equals(BenefitMatcher.NormalizeWords(token), StringComparison.OrdinalIgnoreCase)))
        {
            return 90;
        }

        if (benefitTokens.Any(token => BenefitMatcher.IsMatch(line.Benefit, token)))
        {
            return 80;
        }

        return 0;
    }

    private static decimal? CombinedLossRatio(IEnumerable<ExperienceLine> rows)
    {
        var premium = rows.Sum(row => row.BilledPremium ?? 0m);
        var claims = rows.Sum(row => row.PaidClaims ?? row.IncurredClaims ?? 0m);
        return premium > 0m && claims > 0m
            ? decimal.Round(claims / premium * 100m, 1)
            : null;
    }

    private static decimal? NormalizePercent(decimal? value)
    {
        if (!value.HasValue)
        {
            return null;
        }

        return value.Value is > 0m and <= 1m ? value.Value * 100m : value.Value;
    }

    private static List<string> BuildRenewalInsights(decimal? premiumIncrease, decimal? ehcLossRatio, decimal? dentalLossRatio)
    {
        var insights = new List<string>();

        if (ehcLossRatio > 90m)
        {
            insights.Add("EHC experience is running above target.");
        }
        else if (ehcLossRatio is >= 80m)
        {
            insights.Add("EHC experience is approaching the watch threshold.");
        }
        else if (ehcLossRatio.HasValue)
        {
            insights.Add("EHC experience is within the expected range.");
        }

        if (premiumIncrease > 10m)
        {
            insights.Add("Premium increase exceeds expected trend.");
        }
        else if (premiumIncrease is >= 5m)
        {
            insights.Add("Premium increase should be reviewed before client presentation.");
        }
        else if (premiumIncrease.HasValue)
        {
            insights.Add("Premium movement is inside the expected range.");
        }

        if (dentalLossRatio.HasValue)
        {
            insights.Add(dentalLossRatio > 85m ? "Dental experience needs review." : "Dental experience remains stable.");
        }

        return insights.Count == 0
            ? ["Run validation to generate client-specific insights."]
            : insights;
    }

    private static List<string> MissingDataItems(AnalysisResult analysis)
    {
        return analysis.SourceChecks
            .Where(check => check.Status.Equals("Missing", StringComparison.OrdinalIgnoreCase))
            .Select(check => check.Area)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private async Task<DatabaseSaveResult?> SaveClaimsDataModelToDatabaseAsync()
    {
        if (_claimsDataModelExport is null)
        {
            return null;
        }

        return await _databaseService.SaveClaimsExperienceAsync(
            _claimsDataModelExport,
            ResolveClaimsDataModelSourceFiles(),
            ResolveClaimsDataModelSourceFolder());
    }

    private async Task<DatabaseSaveResult?> SaveClaimsExperienceDuringValidationAsync(
        AnalysisResult analysis)
    {
        try
        {
            var request = await CreateClaimsDataModelRequestAsync(includeBookletSchedule: false);
            var result = await _claimsDataModelBuilder
                .ExtractAsync(request)
                .WaitAsync(ClaimsRefreshTimeout);

            if (result.Summaries.Count == 0 &&
                result.MonthlyClaims.Count == 0 &&
                result.AllPeriodMonthlyClaims.Count == 0)
            {
                return null;
            }

            result.ClientName = FirstNonBlank(
                ClientNameResolver.Resolve(
                    analysis.RenewalCase.ClientName,
                    analysis.Pdf.ClientName,
                    analysis.RenewalCase.PdfPath),
                result.ClientName);
            result.ClientNumber = FirstNonBlank(
                analysis.Pdf.PolicyNumber,
                analysis.RenewalCase.PolicyNumber,
                result.ClientNumber);
            result.Carrier = FirstNonBlank(
                analysis.Pdf.Carrier,
                result.Carrier,
                DetectCarrier(analysis));

            return await _databaseService
                .SaveClaimsExperienceAsync(
                    result,
                    request.SourceFiles,
                    request.SourceFolder,
                    monthlyOnly: true)
                .WaitAsync(ClaimsRefreshTimeout);
        }
        catch (TimeoutException exception)
        {
            AppLogger.Error(exception);
            return DatabaseSaveResult.Failure(
                $"Monthly claims refresh timed out after {ClaimsRefreshTimeout.TotalMinutes:0} minutes. " +
                "Validation and export can continue; use locally synced source files before retrying the claims refresh.",
                _databaseService.DatabasePath);
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            return DatabaseSaveResult.Failure(
                $"Monthly claims extraction needs review: {ex.Message}",
                _databaseService.DatabasePath);
        }
    }

    private static string DatabaseStatusLine(DatabaseSaveResult? result)
    {
        if (result is null)
        {
            return "";
        }

        return $"{Environment.NewLine}Database: {result.Message}";
    }

    private void SaveSettings()
    {
        _settingsStore.Save(new UserSettings
        {
            PdfPath = PdfPathTextBox.Text.Trim(),
            AdditionalPdfPaths = _additionalPdfPaths.ToList(),
            WorkbookPath = WorkbookPathTextBox.Text.Trim(),
            AdditionalWorkbookPaths = _additionalWorkbookPaths.ToList(),
            BenefitBookletPaths = _bookletInputPaths.ToList(),
            CurrentYearFolder = CurrentYearFolderTextBox.Text.Trim(),
            PriorYearFolder = PriorYearFolderTextBox.Text.Trim(),
            AsoBillingFolder = AsoBillingFolderTextBox.Text.Trim(),
            HcsaCalculation = _hcsaCalculation.Clone(),
            AllstateCriticalIllnessEnrollment = _allstateCriticalIllnessEnrollment.Clone(),
            ReportMonthCount = SelectedReportMonthCount(),
            AutoDetectReportMonthCount = SelectedReportMonthCount() == 0,
            TemplatePath = TemplatePathTextBox.Text.Trim(),
            OutputFolder = OutputFolderTextBox.Text.Trim()
        });
    }

    private int SelectedReportMonthCount()
    {
        if (ReportMonthCountComboBox.SelectedItem is System.Windows.Controls.ComboBoxItem item)
        {
            if (int.TryParse(item.Tag?.ToString(), out var tagged))
            {
                return NormalizeReportMonthCount(tagged);
            }

            if (int.TryParse(item.Content?.ToString(), out var selected))
            {
                return NormalizeReportMonthCount(selected);
            }
        }

        if (int.TryParse(ReportMonthCountComboBox.Text, out var typed))
        {
            return NormalizeReportMonthCount(typed);
        }

        return 0;
    }

    private void SetReportMonthCount(int value)
    {
        var normalized = NormalizeReportMonthCount(value);
        foreach (var item in ReportMonthCountComboBox.Items.OfType<System.Windows.Controls.ComboBoxItem>())
        {
            var itemValue = int.TryParse(item.Tag?.ToString(), out var tagged)
                ? tagged
                : int.TryParse(item.Content?.ToString(), out var contentValue)
                    ? contentValue
                    : -1;
            if (itemValue == normalized)
            {
                ReportMonthCountComboBox.SelectedItem = item;
                return;
            }
        }

        ReportMonthCountComboBox.SelectedIndex = 0;
    }

    private static int NormalizeReportMonthCount(int value)
    {
        return value is >= 1 and <= 36 ? value : 0;
    }

    private void UpdateReview()
    {
        if (_analysis is null)
        {
            return;
        }

        CurrentPremiumTextBlock.Text = ParsingHelpers.Money(_analysis.CurrentMonthlyPremium);
        RenewalPremiumTextBlock.Text = ParsingHelpers.Money(_analysis.RenewalMonthlyPremium);
        ChangeTextBlock.Text = $"{ParsingHelpers.Money(_analysis.MonthlyPremiumChange)} ({ParsingHelpers.Percent(_analysis.MonthlyPremiumChangePercent)})";

        var critical = _analysis.Findings.Count(finding => finding.Severity == FindingSeverity.Critical);
        var warnings = _analysis.Findings.Count(finding => finding.Severity == FindingSeverity.Warning);
        FindingCountTextBlock.Text = $"{critical} / {warnings}";

        FindingsGrid.ItemsSource = null;
        FindingsGrid.ItemsSource = _analysis.Findings.ToList();

        RefreshSourceChecksGrid();

        UpdateAssurePlusReview();
        _claimsRangeReviewViewModel.Load(_analysis.ClaimsRange);
    }

    private void RefreshSourceChecksGrid()
    {
        SourceChecksGrid.ItemsSource = null;
        SourceChecksGrid.ItemsSource = _analysis?.SourceChecks.ToList();
    }

    private void UpdateAssurePlusReview()
    {
        var summary = _analysis?.AssurePlusExperience;
        if (summary is null)
        {
            ResetAssurePlusReview();
            return;
        }

        var currentYear = summary.EffectiveDate?.Year ??
                          summary.AnnualRatings.Select(row => row.Year).DefaultIfEmpty(DateTime.Today.Year).Max();
        var firstYear = currentYear - 2;
        var history = new List<AssurePlusHistoryReviewRow>();
        for (var year = firstYear; year <= currentYear; year++)
        {
            var health = summary.AnnualRatings.FirstOrDefault(row =>
                row.Year == year && row.Benefit.Equals("Health", StringComparison.OrdinalIgnoreCase));
            var dental = summary.AnnualRatings.FirstOrDefault(row =>
                row.Year == year && row.Benefit.Equals("Dental", StringComparison.OrdinalIgnoreCase));
            var validation = health is null || dental is null
                ? "Missing"
                : health.Conflict || dental.Conflict
                    ? "Conflict - review required"
                    : "Validated";
            history.Add(new AssurePlusHistoryReviewRow(
                year,
                health?.Rating ?? "Not available",
                dental?.Rating ?? "Not available",
                validation));
        }

        AssurePlusHistoryGrid.ItemsSource = history;
        AssurePlusPricingGrid.ItemsSource = _analysis!.Pdf.AssurePlusPricingReports
            .Where(report => report.EffectiveDate.HasValue)
            .OrderBy(report => report.EffectiveDate)
            .SelectMany(report => report.Benefits.Select(benefit => new AssurePlusPricingReviewRow(
                report.EffectiveDate!.Value.ToString("MMM d, yyyy", CultureInfo.InvariantCulture),
                benefit.Benefit,
                $"{report.EffectiveDate.Value.Year - 1}: {benefit.PreviousYearRating}",
                $"{report.EffectiveDate.Value.Year}: {benefit.RecentYearRating}",
                FormatAssurePlusPercent(benefit.StandardChangePercent, includeSign: false),
                FormatAssurePlusPercent(benefit.ExperienceAdjustmentPercent),
                FormatAssurePlusPercent(benefit.TotalChangePercent),
                report.SourceReference)))
            .ToList();
        AssurePlusDistributionGrid.ItemsSource = _analysis.Pdf.AssurePlusPricingReports
            .Where(report => report.EffectiveDate.HasValue)
            .OrderBy(report => report.EffectiveDate)
            .SelectMany(report => report.DistributionRows.Select(row => new AssurePlusDistributionReviewRow(
                report.EffectiveDate!.Value.Year,
                row.Benefit,
                row.Category,
                FormatAssurePlusPercent(row.YourClaimsPercent, includeSign: false),
                FormatAssurePlusPercent(row.BenchmarkClaimsPercent, includeSign: false),
                row.SourceReference)))
            .ToList();

        SetAssurePlusOutcome(
            summary.HealthOutcome,
            AssurePlusHealthStatusTextBlock,
            AssurePlusHealthTotalTextBlock,
            AssurePlusHealthDetailTextBlock);
        SetAssurePlusOutcome(
            summary.DentalOutcome,
            AssurePlusDentalStatusTextBlock,
            AssurePlusDentalTotalTextBlock,
            AssurePlusDentalDetailTextBlock);

        AssurePlusDentalCommentaryTextBox.Text = summary.DentalCommentary;
        AssurePlusHealthCommentaryTextBox.Text = summary.HealthCommentary;
        AssurePlusTakeawayTextBox.Text = summary.Takeaway;
        AssurePlusDentalCommentaryTextBox.IsEnabled = true;
        AssurePlusHealthCommentaryTextBox.IsEnabled = true;
        AssurePlusTakeawayTextBox.IsEnabled = true;

        var validationMessages = summary.Warnings.Count == 0
            ? new[] { $"Validated: complete {firstYear}-{currentYear} Health and Dental history. Ready for workbook and PowerPoint export." }
            : summary.Warnings.ToArray();
        AssurePlusValidationItemsControl.ItemsSource = validationMessages;
        AssurePlusAvailabilityTextBlock.Text = summary.Warnings.Count == 0
            ? $"AssurePLUS data found and validated for {firstYear}-{currentYear}. Commentary edits below will be used in Export Pack."
            : $"AssurePLUS data found with {summary.Warnings.Count} item(s) requiring review before the slide can be client-ready.";
    }

    private static void SetAssurePlusOutcome(
        AssurePlusBenefitPricing? outcome,
        TextBlock statusTextBlock,
        TextBlock totalTextBlock,
        TextBlock detailTextBlock)
    {
        if (outcome is null)
        {
            statusTextBlock.Text = "Not available";
            totalTextBlock.Text = "-";
            detailTextBlock.Text = "No current pricing outcome was extracted.";
            return;
        }

        statusTextBlock.Text = string.IsNullOrWhiteSpace(outcome.RecentYearRating)
            ? "Not available"
            : outcome.RecentYearRating;
        totalTextBlock.Text = FormatAssurePlusPercent(outcome.TotalChangePercent);
        detailTextBlock.Text =
            $"Standard change: {FormatAssurePlusPercent(outcome.StandardChangePercent, includeSign: false)} | " +
            $"Experience adjustment: {FormatAssurePlusPercent(outcome.ExperienceAdjustmentPercent)}";
    }

    private void ApplyAssurePlusCommentaryEdits()
    {
        var summary = _analysis?.AssurePlusExperience;
        if (summary is null)
        {
            return;
        }

        summary.DentalCommentary = AssurePlusDentalCommentaryTextBox.Text.Trim();
        summary.HealthCommentary = AssurePlusHealthCommentaryTextBox.Text.Trim();
        summary.Takeaway = AssurePlusTakeawayTextBox.Text.Trim();
    }

    private static string FormatAssurePlusPercent(decimal? value, bool includeSign = true)
    {
        if (!value.HasValue)
        {
            return "Not available";
        }

        var sign = includeSign
            ? value.Value > 0m ? "+" : value.Value < 0m ? "\u2212" : ""
            : value.Value < 0m ? "\u2212" : "";
        return $"{sign}{Math.Abs(value.Value).ToString("0.00", CultureInfo.InvariantCulture)}%";
    }

    private void ResetAssurePlusReview()
    {
        AssurePlusAvailabilityTextBlock.Text =
            "No AssurePLUS pricing page has been validated. Select the current renewal PDF and the prior-year PDF, then click Validate.";
        AssurePlusHistoryGrid.ItemsSource = null;
        AssurePlusPricingGrid.ItemsSource = null;
        AssurePlusDistributionGrid.ItemsSource = null;
        AssurePlusHealthStatusTextBlock.Text = "Not available";
        AssurePlusHealthTotalTextBlock.Text = "-";
        AssurePlusHealthDetailTextBlock.Text = "Validate the source PDFs to populate this outcome.";
        AssurePlusDentalStatusTextBlock.Text = "Not available";
        AssurePlusDentalTotalTextBlock.Text = "-";
        AssurePlusDentalDetailTextBlock.Text = "Validate the source PDFs to populate this outcome.";
        AssurePlusDentalCommentaryTextBox.Text = "";
        AssurePlusHealthCommentaryTextBox.Text = "";
        AssurePlusTakeawayTextBox.Text = "";
        AssurePlusDentalCommentaryTextBox.IsEnabled = false;
        AssurePlusHealthCommentaryTextBox.IsEnabled = false;
        AssurePlusTakeawayTextBox.IsEnabled = false;
        AssurePlusValidationItemsControl.ItemsSource = new[] { "Awaiting AssurePLUS source validation." };
    }

    private void ResetReviewDisplay()
    {
        CurrentPremiumTextBlock.Text = "-";
        RenewalPremiumTextBlock.Text = "-";
        ChangeTextBlock.Text = "-";
        FindingCountTextBlock.Text = "-";
        FindingsGrid.ItemsSource = null;
        SourceChecksGrid.ItemsSource = null;
        ResetAssurePlusReview();
        _claimsRangeReviewViewModel.Load(null);
        ExportTextBlock.Text = "Nothing exported yet.";
    }

    private void SelectBatchParentFolder_Click(object sender, RoutedEventArgs e)
    {
        var folder = PromptForFolder("Choose one folder per client, or a flat testing folder containing only booklets");
        if (string.IsNullOrWhiteSpace(folder))
        {
            return;
        }

        BatchParentFolderTextBox.Text = folder;
        _batchOutputFolder = Path.Combine(folder, "Batch Output");
        BatchOutputFolderTextBox.Text = _batchOutputFolder;
        _batchDiscoveryResult = null;
        _batchProcessingResult = null;
        _batchDiscoveryRows.Clear();
        ResetBatchProgress();
        BatchSummaryTextBlock.Text = "Parent folder selected. Run Document Discovery.";
        UpdateBatchButtons(isBusy: false);
        SetStatus("Batch parent folder selected. Run Document Discovery.");
    }

    private async void DocumentDiscovery_Click(object sender, RoutedEventArgs e)
    {
        var request = CreateBatchProcessingRequest();
        if (!Directory.Exists(request.ParentFolder))
        {
            System.Windows.MessageBox.Show(
                this,
                "Select a parent folder before running Document Discovery.",
                "Document Discovery",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        try
        {
            SetBusy(true, "Running document discovery...");
            ResetBatchProgress();
            _batchProcessingResult = null;
            _batchDiscoveryRows.Clear();

            var progress = new Progress<BatchProcessingProgress>(UpdateBatchProgress);
            _batchDiscoveryResult = await _batchDocumentProcessorService.DiscoverAsync(request, progress);
            _batchOutputFolder = _batchDiscoveryResult.OutputFolder;
            BatchOutputFolderTextBox.Text = _batchOutputFolder;

            foreach (var client in _batchDiscoveryResult.Clients)
            {
                _batchDiscoveryRows.Add(client);
            }

            var ready = _batchDiscoveryResult.Clients.Count(client => client.Status == BatchProcessingStatuses.ReadyToProcess);
            var needsReview = _batchDiscoveryResult.Clients.Count(client => client.Status == BatchProcessingStatuses.NeedsReview);
            var warningText = _batchDiscoveryResult.Warnings.Count == 0
                ? ""
                : $" Warnings: {string.Join(" ", _batchDiscoveryResult.Warnings)}";
            BatchSummaryTextBlock.Text =
                $"Discovered {_batchDiscoveryResult.Clients.Count} client group(s). Ready to Process: {ready}. Needs Review: {needsReview}.{warningText}";
            UpdateDocumentBatchKpis();
            SetStatus("Document discovery complete.");
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            System.Windows.MessageBox.Show(this, ex.Message, "Document discovery failed", MessageBoxButton.OK, MessageBoxImage.Error);
            SetStatus("Document discovery failed.");
        }
        finally
        {
            SetBusy(false);
        }
    }

    private async void ProcessBatch_Click(object sender, RoutedEventArgs e)
    {
        if (_batchDiscoveryResult is null)
        {
            System.Windows.MessageBox.Show(
                this,
                "Run Document Discovery first. Batch Processing uses those discovery results and does not scan the folders again.",
                "Process Batch",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        try
        {
            SetBusy(true, "Processing batch...");
            var progress = new Progress<BatchProcessingProgress>(UpdateBatchProgress);
            _batchProcessingResult = await _batchDocumentProcessorService.ProcessBatchAsync(
                _batchDiscoveryResult,
                CreateBatchProcessingRequest(),
                progress);

            _batchOutputFolder = _batchProcessingResult.OutputFolder;
            BatchOutputFolderTextBox.Text = _batchOutputFolder;
            ApplyBatchProcessingResultsToGrid(_batchProcessingResult);
            BatchSummaryTextBlock.Text = BuildBatchSummary(_batchProcessingResult);
            UpdateDocumentBatchKpis();
            await LoadSavedRenewalQueueAsync();
            await LoadSavedClientListAsync();
            await RefreshHomeDashboardAsync();
            SetStatus("Batch processing complete. Batch log saved.");
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            System.Windows.MessageBox.Show(this, ex.Message, "Batch processing failed", MessageBoxButton.OK, MessageBoxImage.Error);
            SetStatus("Batch processing failed.");
        }
        finally
        {
            SetBusy(false);
        }
    }

    private void OpenBatchLog_Click(object sender, RoutedEventArgs e)
    {
        var logPath = _batchProcessingResult?.LogPath ?? "";
        if (!File.Exists(logPath))
        {
            return;
        }

        Process.Start(new ProcessStartInfo
        {
            FileName = logPath,
            UseShellExecute = true
        });
    }

    private void OpenBatchOutputFolder_Click(object sender, RoutedEventArgs e)
    {
        var folder = FirstNonBlank(_batchProcessingResult?.OutputFolder, _batchOutputFolder, BatchOutputFolderTextBox.Text);
        if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
        {
            return;
        }

        Process.Start(new ProcessStartInfo
        {
            FileName = folder,
            UseShellExecute = true
        });
    }

    private BatchProcessingRequest CreateBatchProcessingRequest()
    {
        var parentFolder = BatchParentFolderTextBox.Text.Trim();
        var outputFolder = FirstNonBlank(
            _batchOutputFolder,
            BatchOutputFolderTextBox.Text,
            string.IsNullOrWhiteSpace(parentFolder) ? "" : Path.Combine(parentFolder, "Batch Output"));

        return new BatchProcessingRequest
        {
            ParentFolder = parentFolder,
            OutputFolder = outputFolder,
            RenewalPowerPointTemplatePath = TemplatePathTextBox.Text.Trim(),
            ClaimsPowerPointTemplatePath = SafeResolveClaimsPowerPointTemplatePath(),
            ReportMonthCount = SelectedReportMonthCount(),
            DatabasePath = _databaseService.DatabasePath
        };
    }

    private string SafeResolveClaimsPowerPointTemplatePath()
    {
        try
        {
            return ResolveClaimsPowerPointTemplatePath();
        }
        catch
        {
            return "";
        }
    }

    private void UpdateBatchProgress(BatchProcessingProgress progress)
    {
        BatchCurrentClientTextBlock.Text = string.IsNullOrWhiteSpace(progress.CurrentClient) ? "-" : progress.CurrentClient;
        BatchCurrentWorkflowTextBlock.Text = string.IsNullOrWhiteSpace(progress.CurrentWorkflow) ? "-" : progress.CurrentWorkflow;
        BatchProgressCountTextBlock.Text = $"{progress.Completed} / {progress.Total}";
        BatchCompletedTextBlock.Text = progress.Completed.ToString(CultureInfo.InvariantCulture);
        BatchWarningFailedTextBlock.Text = $"{progress.WarningCount} / {progress.FailedCount}";
        BatchEtaTextBlock.Text = progress.EstimatedTimeRemaining == TimeSpan.Zero
            ? "-"
            : FormatBatchDuration(progress.EstimatedTimeRemaining);
        BatchProgressBar.Value = progress.Total <= 0
            ? 0
            : Math.Min(100, Math.Max(0, progress.Completed * 100d / progress.Total));

        if (!string.IsNullOrWhiteSpace(progress.CurrentClient) &&
            !progress.CurrentWorkflow.Equals("Document Discovery", StringComparison.OrdinalIgnoreCase))
        {
            var row = _batchDiscoveryRows.FirstOrDefault(client =>
                client.ClientName.Equals(progress.CurrentClient, StringComparison.OrdinalIgnoreCase));
            if (row is not null &&
                row.Status is not BatchProcessingStatuses.Completed and
                    not BatchProcessingStatuses.CompletedWithWarnings and
                    not BatchProcessingStatuses.Failed and
                    not BatchProcessingStatuses.Skipped)
            {
                row.Status = BatchProcessingStatuses.Processing;
                BatchDiscoveryGrid.Items.Refresh();
            }
        }

        if (!string.IsNullOrWhiteSpace(progress.Message))
        {
            SetStatus(progress.Message);
        }

        UpdateDocumentBatchKpis();
    }

    private void ResetBatchProgress()
    {
        BatchCurrentClientTextBlock.Text = "-";
        BatchCurrentWorkflowTextBlock.Text = "-";
        BatchProgressBar.Value = 0;
        BatchProgressCountTextBlock.Text = "0 / 0";
        BatchCompletedTextBlock.Text = "0";
        BatchWarningFailedTextBlock.Text = "0 / 0";
        BatchEtaTextBlock.Text = "-";
        UpdateDocumentBatchKpis();
    }

    private void ApplyBatchProcessingResultsToGrid(BatchProcessingResult result)
    {
        foreach (var discoveryRow in _batchDiscoveryRows)
        {
            var processedRow = result.Clients.FirstOrDefault(client =>
                client.ClientName.Equals(discoveryRow.ClientName, StringComparison.OrdinalIgnoreCase));
            if (processedRow is null)
            {
                continue;
            }

            discoveryRow.Status = processedRow.Status;
            discoveryRow.Warnings.Clear();
            discoveryRow.Warnings.AddRange(processedRow.Warnings);
            discoveryRow.Warnings.AddRange(processedRow.Errors);
        }

        BatchDiscoveryGrid.Items.Refresh();
    }

    private void UpdateDocumentBatchKpis()
    {
        var discoveryClients = _batchDiscoveryResult?.Clients ?? _batchDiscoveryRows.ToList();
        var totalDocuments = discoveryClients.Sum(client => client.DocumentDiscovery.Documents.Count);
        var needsReviewDocuments = discoveryClients
            .Where(client => client.Status == BatchProcessingStatuses.NeedsReview)
            .Sum(client => Math.Max(1, client.DocumentDiscovery.Documents.Count));

        var processedDocuments = 0;
        var failedDocuments = 0;
        var savedDocuments = 0;
        if (_batchProcessingResult is not null)
        {
            foreach (var client in _batchProcessingResult.Clients)
            {
                var documentCount = discoveryClients
                    .FirstOrDefault(discovery => discovery.ClientName.Equals(client.ClientName, StringComparison.OrdinalIgnoreCase))
                    ?.DocumentDiscovery.Documents.Count ?? 1;

                if (client.Status is BatchProcessingStatuses.Completed or BatchProcessingStatuses.CompletedWithWarnings)
                {
                    processedDocuments += Math.Max(1, documentCount);
                }

                if (client.Status == BatchProcessingStatuses.Failed)
                {
                    failedDocuments += Math.Max(1, documentCount);
                }

                savedDocuments += client.Workflows.Count(workflow => workflow.SavedToDatabase);
                if (client.RenewalSavedToDatabase)
                {
                    savedDocuments++;
                }
            }
        }

        DocumentBatchTotalTextBlock.Text = totalDocuments.ToString("#,0", CultureInfo.InvariantCulture);
        DocumentBatchProcessedTextBlock.Text = processedDocuments.ToString("#,0", CultureInfo.InvariantCulture);
        DocumentBatchNeedsReviewTextBlock.Text = needsReviewDocuments.ToString("#,0", CultureInfo.InvariantCulture);
        DocumentBatchFailedTextBlock.Text = failedDocuments.ToString("#,0", CultureInfo.InvariantCulture);
        DocumentBatchSavedTextBlock.Text = savedDocuments.ToString("#,0", CultureInfo.InvariantCulture);
    }

    private void PauseBatchProcessing_Click(object sender, RoutedEventArgs e)
    {
        SetStatus("Pause Processing is ready for the future queue controller. Current batches finish the active processing run once started.");
    }

    private void RetryFailedDocuments_Click(object sender, RoutedEventArgs e)
    {
        var failed = _batchProcessingResult?.Clients.Count(client => client.Status == BatchProcessingStatuses.Failed) ?? 0;
        SetStatus(failed == 0
            ? "No failed documents are available to retry."
            : $"Retry Failed is ready for {failed} failed item(s). Re-run Process Batch to process the current discovered set.");
    }

    private void ReviewExtractionIssues_Click(object sender, RoutedEventArgs e)
    {
        var issueCount = _batchDiscoveryRows.Count(row => row.Status == BatchProcessingStatuses.NeedsReview || row.Warnings.Count > 0);
        SetStatus(issueCount == 0
            ? "No extraction issues are currently flagged."
            : $"Review Issues: {issueCount} row(s) need review in the discovery grid.");
    }

    private void SaveApprovedBatchResults_Click(object sender, RoutedEventArgs e)
    {
        SetStatus("Approved batch results are saved by the existing processing workflows when each renewal, claims, booklet, or plan design extraction succeeds.");
    }

    private void ExportBatchResults_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var rows = _batchProcessingResult?.Clients
                .Select(client => new
                {
                    client.ClientName,
                    client.Carrier,
                    client.Status,
                    client.DocumentTypesFound,
                    client.WorkflowsExecuted,
                    client.RenewalSavedToDatabase,
                    client.PlanDesignRowsSaved,
                    client.WarningsText,
                    client.ErrorsText
                })
                .ToList();

            if (rows is null || rows.Count == 0)
            {
                rows = _batchDiscoveryRows
                    .Select(client => new
                    {
                        client.ClientName,
                        client.Carrier,
                        client.Status,
                        DocumentTypesFound = client.DocumentTypesText,
                        WorkflowsExecuted = client.WorkflowsText,
                        RenewalSavedToDatabase = false,
                        PlanDesignRowsSaved = 0,
                        client.WarningsText,
                        ErrorsText = ""
                    })
                    .ToList();
            }

            if (rows.Count == 0)
            {
                SetStatus("Run Document Discovery before exporting batch results.");
                return;
            }

            var outputFolder = FirstNonBlank(_batchProcessingResult?.OutputFolder, _batchOutputFolder, ResolveOutputRootFolder());
            Directory.CreateDirectory(outputFolder);
            var workbookPath = Path.Combine(outputFolder, $"Document_Intelligence_Batch_Results_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");

            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Batch Results");
            var headers = new[]
            {
                "Client", "Carrier", "Status", "Document Types", "Workflows", "Renewal Saved",
                "Plan Design Rows Saved", "Warnings", "Errors"
            };

            for (var column = 0; column < headers.Length; column++)
            {
                worksheet.Cell(1, column + 1).Value = headers[column];
            }

            var rowNumber = 2;
            foreach (var row in rows)
            {
                worksheet.Cell(rowNumber, 1).Value = row.ClientName;
                worksheet.Cell(rowNumber, 2).Value = row.Carrier;
                worksheet.Cell(rowNumber, 3).Value = row.Status;
                worksheet.Cell(rowNumber, 4).Value = row.DocumentTypesFound;
                worksheet.Cell(rowNumber, 5).Value = row.WorkflowsExecuted;
                worksheet.Cell(rowNumber, 6).Value = row.RenewalSavedToDatabase ? "Yes" : "No";
                worksheet.Cell(rowNumber, 7).Value = row.PlanDesignRowsSaved;
                worksheet.Cell(rowNumber, 8).Value = row.WarningsText;
                worksheet.Cell(rowNumber, 9).Value = row.ErrorsText;
                rowNumber++;
            }

            worksheet.Range(1, 1, 1, headers.Length).Style.Font.Bold = true;
            worksheet.Columns().AdjustToContents();
            workbook.SaveAs(workbookPath);
            SetStatus($"Batch results exported: {workbookPath}");
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            System.Windows.MessageBox.Show(this, ex.Message, "Export Batch Results", MessageBoxButton.OK, MessageBoxImage.Error);
            SetStatus("Batch results export failed.");
        }
    }

    private static string BuildBatchSummary(BatchProcessingResult result)
    {
        var completed = result.Clients.Count(client => client.Status == BatchProcessingStatuses.Completed);
        var warnings = result.Clients.Count(client => client.Status == BatchProcessingStatuses.CompletedWithWarnings);
        var failed = result.Clients.Count(client => client.Status == BatchProcessingStatuses.Failed);
        var skipped = result.Clients.Count(client => client.Status == BatchProcessingStatuses.Skipped);
        return $"Batch complete. Completed: {completed}. Warnings: {warnings}. Failed: {failed}. Skipped: {skipped}. Log: {result.LogPath}";
    }

    private void UpdateBatchButtons(bool isBusy)
    {
        var parentFolder = BatchParentFolderTextBox.Text.Trim();
        var outputFolder = FirstNonBlank(_batchProcessingResult?.OutputFolder, _batchOutputFolder, BatchOutputFolderTextBox.Text);

        SelectBatchParentFolderButton.IsEnabled = !isBusy;
        DocumentDiscoveryButton.IsEnabled = !isBusy && Directory.Exists(parentFolder);
        ProcessBatchButton.IsEnabled = !isBusy && _batchDiscoveryResult is not null;
        OpenBatchLogButton.IsEnabled = !isBusy && File.Exists(_batchProcessingResult?.LogPath ?? "");
        OpenBatchOutputFolderButton.IsEnabled = !isBusy && Directory.Exists(outputFolder);
    }

    private static string FormatBatchDuration(TimeSpan duration)
    {
        return duration.TotalSeconds < 1
            ? "<1 sec"
            : duration.TotalMinutes < 1
                ? $"{duration.TotalSeconds:0} sec"
                : $"{duration.TotalMinutes:0.0} min";
    }

    private void OpenOutputFolder_Click(object sender, RoutedEventArgs e)
    {
        var folder = _billingOutputFolder ?? _renewalQueueExportFolder ?? _proposalOutputFolder ?? _bookletExtractionResult?.OutputFolder ?? _claimsPowerPointExport?.OutputFolder ?? _claimsDataModelExport?.OutputFolder ?? _export?.OutputFolder;
        if (!IsOutputFolderAvailableForUi(folder))
        {
            return;
        }

        Process.Start(new ProcessStartInfo
        {
            FileName = folder,
            UseShellExecute = true
        });
    }

    private void SetBusy(bool isBusy, string? status = null)
    {
        BusyProgressBar.Visibility = isBusy ? Visibility.Visible : Visibility.Collapsed;
        AddAdditionalPdfsButton.IsEnabled = !isBusy;
        RemoveAdditionalPdfsButton.IsEnabled = !isBusy;
        AddAdditionalWorkbooksButton.IsEnabled = !isBusy;
        RemoveAdditionalWorkbooksButton.IsEnabled = !isBusy;
        AddBenefitBookletPdfsButton.IsEnabled = !isBusy;
        AddBenefitBookletFolderButton.IsEnabled = !isBusy;
        RemoveBenefitBookletsButton.IsEnabled = !isBusy;
        CancelBookletExtractionButton.IsEnabled = _bookletExtractionCancellation is not null;
        ExtractBillingStatementButton.IsEnabled = !isBusy;
        ClearFormButton.IsEnabled = !isBusy;
        ValidateButton.IsEnabled = !isBusy;
        ExportClaimsDataModelButton.IsEnabled = !isBusy;
        ExportClaimsPowerPointButton.IsEnabled = !isBusy;
        ExtractBookletToMasterDbButton.IsEnabled = !isBusy;
        RebuildBenefitsMasterDatabaseButton.IsEnabled = !isBusy;
        ExportButton.IsEnabled = !isBusy;
        FooterMoreActionsButton.IsEnabled = !isBusy;
        FooterExportClaimsDataModelMenuItem.IsEnabled = !isBusy;
        FooterExportClaimsPowerPointMenuItem.IsEnabled = !isBusy;
        FooterExtractBookletToMasterDbMenuItem.IsEnabled = !isBusy;
        SetOpenOutputButtonsEnabled(!isBusy &&
                                    ((_bookletExtractionResult is not null && IsOutputFolderAvailableForUi(_bookletExtractionResult.OutputFolder)) ||
                                     (_export is not null && IsOutputFolderAvailableForUi(_export.OutputFolder)) ||
                                     (_claimsDataModelExport is not null && IsOutputFolderAvailableForUi(_claimsDataModelExport.OutputFolder)) ||
                                     (_claimsPowerPointExport is not null && IsOutputFolderAvailableForUi(_claimsPowerPointExport.OutputFolder)) ||
                                     IsOutputFolderAvailableForUi(_renewalQueueExportFolder) ||
                                     IsOutputFolderAvailableForUi(_billingOutputFolder) ||
                                     IsOutputFolderAvailableForUi(_proposalOutputFolder)));
        UpdateBatchButtons(isBusy);

        if (!string.IsNullOrWhiteSpace(status))
        {
            SetStatus(status);
        }
    }

    private void SetOpenOutputButtonsEnabled(bool isEnabled)
    {
        OpenOutputFolderButton.IsEnabled = isEnabled;
        OpenOutputFolderFooterButton.IsEnabled = isEnabled;
    }

    private static bool IsOutputFolderAvailableForUi(string? folder)
    {
        if (string.IsNullOrWhiteSpace(folder))
        {
            return false;
        }

        return SharePointWebDavPathPolicy.IsDirectSharePointWebDavPath(folder) || Directory.Exists(folder);
    }

    private void SetStatus(string status)
    {
        StatusTextBlock.Text = status;
        AppLogger.Info(status);
    }

    private static string FirstExisting(params string[] paths)
    {
        return paths.FirstOrDefault(path => !string.IsNullOrWhiteSpace(path) && File.Exists(path)) ?? "";
    }

    private static string FirstNonBlank(params string?[] values)
    {
        return values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value))?.Trim() ?? "";
    }

    private string ResolveClaimsPowerPointTemplatePath()
    {
        var selectedTemplate = TemplatePathTextBox.Text.Trim();
        var selectedClaimsTemplate = Path.GetFileNameWithoutExtension(selectedTemplate)
            .Contains("Experience Report", StringComparison.OrdinalIgnoreCase)
            ? selectedTemplate
            : "";
        var candidates = new[]
        {
            selectedClaimsTemplate,
            MapSharePointTemplateUrl(selectedTemplate),
            DefaultClaimsPowerPointTemplatePath,
            selectedTemplate,
            DefaultTemplatePath
        };

        foreach (var candidate in candidates)
        {
            if (!string.IsNullOrWhiteSpace(candidate) && File.Exists(candidate))
            {
                return candidate;
            }
        }

        throw new FileNotFoundException("Select the claims experience PowerPoint template, such as Experience ReportV1.pptx.", selectedTemplate);
    }

    private static string MapSharePointTemplateUrl(string value)
    {
        if (!Uri.TryCreate(value, UriKind.Absolute, out var uri) ||
            !uri.Host.Contains("sharepoint.com", StringComparison.OrdinalIgnoreCase))
        {
            return "";
        }

        var decodedPath = Uri.UnescapeDataString(uri.AbsolutePath);
        var marker = "/Documents/";
        var markerIndex = decodedPath.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
        if (markerIndex < 0)
        {
            return "";
        }

        var relativePath = decodedPath[(markerIndex + marker.Length)..]
            .Replace('/', Path.DirectorySeparatorChar);
        return string.IsNullOrWhiteSpace(WestlandOneDriveRoot)
            ? ""
            : Path.Combine(WestlandOneDriveRoot, relativePath);
    }

    private static string ResolveWestlandOneDriveRoot()
    {
        var configuredRoot = Environment.GetEnvironmentVariable("WESTLAND_ONEDRIVE_ROOT");
        var candidates = new[]
        {
            configuredRoot,
            Environment.GetEnvironmentVariable("OneDriveCommercial"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "OneDrive - WESTLANDCORP")
        };

        return candidates.FirstOrDefault(path => !string.IsNullOrWhiteSpace(path) && Directory.Exists(path)) ?? "";
    }

    private static string ResolveWestlandPath(params string[] parts) =>
        string.IsNullOrWhiteSpace(WestlandOneDriveRoot)
            ? ""
            : Path.Combine([WestlandOneDriveRoot, .. parts]);

    private string ResolveClaimsDataModelSourceFolder() =>
        ResolveClaimsDataModelSourceFolder(CurrentYearFolderTextBox.Text);

    internal static string ResolveClaimsDataModelSourceFolder(string? currentYearFolder)
    {
        var selectedFolder = currentYearFolder?.Trim() ?? "";
        return Directory.Exists(selectedFolder) ? selectedFolder : "";
    }

    private List<string> ResolveClaimsDataModelSourceFiles()
    {
        var files = new List<string>
        {
            PdfPathTextBox.Text.Trim(),
            WorkbookPathTextBox.Text.Trim()
        };

        files.AddRange(_additionalPdfPaths);
        files.AddRange(_additionalWorkbookPaths);

        return files
            .Where(path => !string.IsNullOrWhiteSpace(path))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private List<string> ResolveClaimsBenefitBookletPaths()
    {
        var paths = new List<string>();
        foreach (var path in _bookletInputPaths)
        {
            AddClaimsBookletPath(paths, path);
        }

        foreach (var path in _additionalPdfPaths.Where(IsLikelyBenefitBookletPdf))
        {
            AddClaimsBookletPath(paths, path);
        }

        var sourceFolder = ResolveClaimsDataModelSourceFolder();
        if (!string.IsNullOrWhiteSpace(sourceFolder) && Directory.Exists(sourceFolder))
        {
            AddLikelyBookletInputsFromFolder(paths, sourceFolder);
        }

        return paths;
    }

    private static void AddLikelyBookletInputsFromFolder(List<string> paths, string sourceFolder)
    {
        foreach (var pdfPath in Directory.EnumerateFiles(sourceFolder, "*.pdf", SearchOption.TopDirectoryOnly).Where(IsLikelyBenefitBookletPdf))
        {
            AddClaimsBookletPath(paths, pdfPath);
        }

        foreach (var childFolder in Directory.EnumerateDirectories(sourceFolder, "*", SearchOption.TopDirectoryOnly))
        {
            if (IsLikelyBenefitBookletFolder(childFolder))
            {
                AddClaimsBookletPath(paths, childFolder);
                continue;
            }

            foreach (var nestedFolder in Directory.EnumerateDirectories(childFolder, "*", SearchOption.TopDirectoryOnly).Where(IsLikelyBenefitBookletFolder))
            {
                AddClaimsBookletPath(paths, nestedFolder);
            }

            foreach (var pdfPath in Directory.EnumerateFiles(childFolder, "*.pdf", SearchOption.TopDirectoryOnly).Where(IsLikelyBenefitBookletPdf))
            {
                AddClaimsBookletPath(paths, pdfPath);
            }
        }
    }

    private static bool IsLikelyBenefitBookletFolder(string path)
    {
        var name = Path.GetFileName(path).ToLowerInvariant();
        return name.Contains("booklet", StringComparison.OrdinalIgnoreCase) ||
               name.Contains("benefit booklet", StringComparison.OrdinalIgnoreCase) ||
               name.Contains("schedule of benefits", StringComparison.OrdinalIgnoreCase) ||
               name.Contains("plan design", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsLikelyBenefitBookletPdf(string path)
    {
        if (string.IsNullOrWhiteSpace(path) ||
            !File.Exists(path) ||
            !Path.GetExtension(path).Equals(".pdf", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        var name = Path.GetFileNameWithoutExtension(path).ToLowerInvariant();
        if (name.Contains("billing", StringComparison.OrdinalIgnoreCase) ||
            name.Contains("billings", StringComparison.OrdinalIgnoreCase) ||
            name.Contains("claims", StringComparison.OrdinalIgnoreCase) ||
            name.Contains("experience", StringComparison.OrdinalIgnoreCase) ||
            name.Contains("insured", StringComparison.OrdinalIgnoreCase) ||
            name.Contains("statement", StringComparison.OrdinalIgnoreCase) ||
            name.Contains("renewal", StringComparison.OrdinalIgnoreCase) ||
            name.Contains("package", StringComparison.OrdinalIgnoreCase) ||
            name.Contains("ehc", StringComparison.OrdinalIgnoreCase) ||
            name.Contains("dental", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        return name.Contains("booklet", StringComparison.OrdinalIgnoreCase) ||
               name.Contains("benefit", StringComparison.OrdinalIgnoreCase) ||
               name.Contains("schedule of benefits", StringComparison.OrdinalIgnoreCase) ||
               name.Contains("plan design", StringComparison.OrdinalIgnoreCase);
    }

    private static void AddClaimsBookletPath(List<string> paths, string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return;
        }

        var normalized = Path.GetFullPath(path.Trim());
        if ((File.Exists(normalized) || Directory.Exists(normalized)) &&
            !paths.Any(existing => existing.Equals(normalized, StringComparison.OrdinalIgnoreCase)))
        {
            paths.Add(normalized);
        }
    }

    private static bool SamePaths(IReadOnlyList<string> left, IReadOnlyList<string> right)
    {
        return left.Count == right.Count &&
               left.Select(Path.GetFullPath).OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
                   .SequenceEqual(
                       right.Select(Path.GetFullPath).OrderBy(path => path, StringComparer.OrdinalIgnoreCase),
                       StringComparer.OrdinalIgnoreCase);
    }

    private static string FormatMoney(decimal? value)
    {
        return value.HasValue
            ? value.Value.ToString("$#,0", CultureInfo.InvariantCulture)
            : "N/A";
    }

    private static string FormatMoneyCompact(decimal? value)
    {
        if (!value.HasValue)
        {
            return "N/A";
        }

        var amount = value.Value;
        var absolute = Math.Abs(amount);
        if (absolute >= 1_000_000m)
        {
            return "$" + (amount / 1_000_000m).ToString("0.##", CultureInfo.InvariantCulture) + "M";
        }

        if (absolute >= 1_000m)
        {
            return "$" + (amount / 1_000m).ToString("0.#", CultureInfo.InvariantCulture) + "K";
        }

        return FormatMoney(amount);
    }

    private static string FormatPercent(decimal? value, bool includeSign = true)
    {
        if (!value.HasValue)
        {
            return "N/A";
        }

        return includeSign && value.Value >= 0m
            ? $"+{value.Value:0.0}%"
            : $"{value.Value:0.0}%";
    }

    private static bool BrowseFolder(System.Windows.Controls.TextBox textBox, string description)
    {
        using var dialog = new Forms.FolderBrowserDialog
        {
            Description = description,
            UseDescriptionForTitle = true,
            SelectedPath = Directory.Exists(textBox.Text) ? textBox.Text : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        };

        if (dialog.ShowDialog() == Forms.DialogResult.OK)
        {
            textBox.Text = dialog.SelectedPath;
            return true;
        }

        return false;
    }
}

public sealed class ExecutiveMetricCard
{
    public string Title { get; init; } = "";
    public string Value { get; init; } = "N/A";
    public string Subtitle { get; init; } = "Drill down";
    public string Icon { get; init; } = "i";
}

public sealed class ExecutivePipelineRow
{
    public string Stage { get; init; } = "";
    public int Count { get; init; }
    public decimal? Premium { get; init; }
    public string PremiumText => Premium.HasValue ? Premium.Value.ToString("$#,0", CultureInfo.InvariantCulture) : "N/A";
    public string OldestItem { get; init; } = "N/A";
    public string AverageDaysInStage { get; init; } = "N/A";
    public int OverdueItems { get; init; }
}

public sealed class ExecutiveChartRow
{
    public string Label { get; init; } = "";
    public int Count { get; init; }
    public string Value { get; init; } = "";
    public string Status { get; init; } = "";
    public decimal? Percent { get; init; }

    public static ExecutiveChartRow Empty(string label)
    {
        return new ExecutiveChartRow
        {
            Label = label,
            Value = "No saved records yet",
            Status = "Ready"
        };
    }

    public override string ToString()
    {
        return string.IsNullOrWhiteSpace(Status)
            ? $"{Label}: {Value}"
            : $"{Label}: {Value} - {Status}";
    }
}

public sealed class ExecutiveSummaryRow
{
    public string Category { get; init; } = "";
    public string Label { get; init; } = "";
    public string Value { get; init; } = "";
    public string Status { get; init; } = "";
    public string Owner { get; init; } = "";
    public int Count { get; init; }
    public string Icon { get; init; } = "";
    public string Badge { get; init; } = "";
    public string IconBrush { get; init; } = "#6D28D9";
    public string IconBackground { get; init; } = "#F3E8FF";

    public override string ToString()
    {
        return string.IsNullOrWhiteSpace(Status)
            ? $"{Label}: {Value}"
            : $"{Label}: {Value} - {Status}";
    }
}

public sealed class BookOfBusinessClientRow
{
    public string ClientName { get; init; } = "";
    public string Owner { get; init; } = "";
    public string Carrier { get; init; } = "";
    public string Industry { get; init; } = "";
    public string Lives { get; init; } = "N/A";
    public string AnnualPremium { get; init; } = "N/A";
    public string RenewalDate { get; init; } = "Not scheduled";
    public string Stage { get; init; } = "";
    public string Status { get; init; } = "";
}

public sealed class ClientWorkspaceViewModel
{
    public string ClientId { get; init; } = "";
    public long? DatabaseClientId { get; set; }
    public string ClientName { get; init; } = "";
    public string Status { get; init; } = "";
    public string Segment { get; init; } = "";
    public int Employees { get; init; }
    public string PrimaryCarrier { get; init; } = "";
    public string PlanType { get; init; } = "";
    public int? HealthScore { get; init; }
    public ClientWorkspaceRenewalHistoryPresentation RenewalHistory { get; init; } = new();
    public DateOnly? LastRenewalDate { get; init; }
    public decimal? CurrentPremium { get; init; }
    public decimal? RenewalPremium { get; init; }
    public decimal? PremiumChangePercent { get; init; }
    public decimal? EhcLossRatio { get; init; }
    public decimal? DentalLossRatio { get; init; }
    public decimal? TargetLossRatio { get; set; }
    public string TlrBasis { get; set; } = "";
    public decimal? PoolingCharge { get; set; }
    public decimal? PoolingThreshold { get; set; }
    public decimal? CurrentContractPoolingThreshold { get; private set; }
    public decimal? ProjectedPoolingThreshold { get; private set; }
    public int? LatestEhcInsuredLives { get; private set; }
    public DateOnly? LatestEhcLivesMonth { get; private set; }
    public string PoolingThresholdBasis { get; set; } = "No verified terms";
    public string AccountExecutive { get; init; } = "";
    public string AccountManager { get; init; } = "";
    public string PrimaryContactName { get; init; } = "";
    public string PrimaryContactEmail { get; init; } = "";
    public string PrimaryContactPhone { get; init; } = "";
    public string PlanEffectiveDate { get; init; } = "";

    public string RouteText => $"/clients/{ClientId}";
    public DateOnly? NextRenewalDate => LastRenewalDate.HasValue
        ? ProjectNextAnnualRenewal(LastRenewalDate.Value, DateOnly.FromDateTime(DateTime.Today))
        : null;
    public string RenewalDate => NextRenewalDate?.ToString("MMM d, yyyy", CultureInfo.InvariantCulture) ?? "TBD";
    public string RenewalCountdownText
    {
        get
        {
            if (!NextRenewalDate.HasValue)
            {
                return "Renewal date not set";
            }

            var days = NextRenewalDate.Value.DayNumber - DateOnly.FromDateTime(DateTime.Today).DayNumber;
            return days == 0 ? "Renews today" : $"{days:#,0} days remaining";
        }
    }
    public string EmployeesText => Employees.ToString("#,0", CultureInfo.InvariantCulture);
    public string HealthScoreText => HealthScore?.ToString(CultureInfo.InvariantCulture) ?? "N/A";
    public string HealthScoreLabel => HealthScore switch
    {
        >= 85 => "Healthy",
        >= 75 => "Good",
        >= 50 => "Monitor",
        null => "N/A",
        _ => "At Risk"
    };

    public string MetadataText => $"Policy # {ClientId}  -  {PrimaryCarrier}  -  Effective {PlanEffectiveDate}  -  {EmployeesText} Employees  -  {PlanType} Plan";
    public string CurrentPremiumText => FormatMoney(CurrentPremium);
    public string RenewalPremiumText => FormatMoney(RenewalPremium);
    public string PremiumChangeText => FormatPercent(PremiumChangePercent, includeSign: true);
    public string PremiumChangeDescription => PremiumChangePercent switch
    {
        > 0m => $"{PremiumChangeText} increase",
        < 0m => $"{PremiumChangeText} decrease",
        0m => "No change",
        _ => "N/A"
    };

    public string EhcLossRatioText => FormatPercent(EhcLossRatio, includeSign: false);
    public string DentalLossRatioText => FormatPercent(DentalLossRatio, includeSign: false);
    public string TargetLossRatioText => FormatPercent(NormalizePercent(TargetLossRatio), includeSign: false);
    public string PoolingChargeText => FormatPercent(NormalizePercent(PoolingCharge), includeSign: false);
    public string PoolingThresholdText => FormatMoney(PoolingThreshold);

    public void ApplySunLifeLapAssessment(SunLifeLapAssessment assessment)
    {
        var currentThreshold = PoolingThreshold;
        CurrentContractPoolingThreshold = currentThreshold;
        ProjectedPoolingThreshold = assessment.ProjectedLapLevel;
        LatestEhcInsuredLives = assessment.LatestEhcLives;
        LatestEhcLivesMonth = assessment.LatestMonth;
        PoolingThreshold = assessment.ProjectedLapLevel;

        if (!currentThreshold.HasValue)
        {
            PoolingThresholdBasis = $"Projected Sun Life LAP at renewal; {assessment.ParticipationBasisText}";
            return;
        }

        PoolingThresholdBasis = currentThreshold.Value == assessment.ProjectedLapLevel
            ? $"Projected Sun Life LAP at renewal; {assessment.ParticipationBasisText}; current contract matches"
            : $"Projected Sun Life LAP at renewal; {assessment.ParticipationBasisText}; current contract {FormatMoney(currentThreshold)}";
    }
    public string KeyRisk => PremiumChangePercent >= 10m || EhcLossRatio >= 90m || DentalLossRatio >= 100m || HealthScore < 70
        ? "High Risk"
        : HealthScore >= 75
            ? "Healthy"
            : "Monitor";

    public Brush StatusBackground => Status switch
    {
        "Active" => new SolidColorBrush(Color.FromRgb(231, 245, 237)),
        "Inactive" => new SolidColorBrush(Color.FromRgb(238, 241, 244)),
        "Prospect" => new SolidColorBrush(Color.FromRgb(231, 241, 251)),
        "New Prospect" => new SolidColorBrush(Color.FromRgb(231, 241, 251)),
        _ => Brushes.White
    };

    public Brush StatusBorder => Status switch
    {
        "Active" => new SolidColorBrush(Color.FromRgb(183, 222, 199)),
        "Inactive" => new SolidColorBrush(Color.FromRgb(207, 214, 222)),
        "Prospect" => new SolidColorBrush(Color.FromRgb(185, 213, 238)),
        "New Prospect" => new SolidColorBrush(Color.FromRgb(185, 213, 238)),
        _ => Brushes.LightGray
    };

    public Brush StatusForeground => Status switch
    {
        "Active" => new SolidColorBrush(Color.FromRgb(31, 111, 67)),
        "Inactive" => new SolidColorBrush(Color.FromRgb(72, 85, 99)),
        "Prospect" => new SolidColorBrush(Color.FromRgb(31, 95, 150)),
        "New Prospect" => new SolidColorBrush(Color.FromRgb(31, 95, 150)),
        _ => Brushes.Black
    };

    public Brush HealthScoreBrush => HealthScore switch
    {
        >= 75 => new SolidColorBrush(Color.FromRgb(22, 163, 74)),
        >= 50 => new SolidColorBrush(Color.FromRgb(245, 158, 11)),
        null => new SolidColorBrush(Color.FromRgb(148, 163, 184)),
        _ => new SolidColorBrush(Color.FromRgb(220, 38, 38))
    };

    internal static DateOnly ProjectNextAnnualRenewal(DateOnly renewalDate, DateOnly today)
    {
        while (renewalDate < today)
        {
            renewalDate = renewalDate.AddYears(1);
        }

        return renewalDate;
    }

    private static string FormatMoney(decimal? value)
    {
        return value.HasValue
            ? value.Value.ToString("$#,0", CultureInfo.InvariantCulture)
            : "N/A";
    }

    private static string FormatPercent(decimal? value, bool includeSign)
    {
        if (!value.HasValue)
        {
            return "N/A";
        }

        var format = includeSign ? "+0.0;-0.0;0.0" : "0.0";
        return value.Value.ToString(format, CultureInfo.InvariantCulture) + "%";
    }

    private static decimal? NormalizePercent(decimal? value)
    {
        if (!value.HasValue)
        {
            return null;
        }

        return value.Value is > 0m and <= 1m ? value.Value * 100m : value.Value;
    }
}

public sealed class ClientListItem
{
    public ClientListItem(
        string clientId,
        string clientName,
        string status,
        string industry,
        int employees,
        string renewalDate,
        string primaryCarrier,
        string planType,
        int? healthScore,
        string lastUpdated,
        long? databaseClientId = null)
    {
        ClientId = clientId;
        DatabaseClientId = databaseClientId;
        ClientName = clientName;
        Status = status;
        Industry = industry;
        Employees = employees;
        RenewalDate = renewalDate;
        PrimaryCarrier = primaryCarrier;
        PlanType = planType;
        HealthScore = healthScore;
        LastUpdated = lastUpdated;
    }

    public string ClientId { get; }
    public long? DatabaseClientId { get; }
    public bool IsSavedClient => DatabaseClientId is not null;
    public string ClientName { get; }
    public string Status { get; }
    public string Industry { get; }
    public int Employees { get; }
    public string RenewalDate { get; }
    public string PrimaryCarrier { get; }
    public string PlanType { get; }
    public int? HealthScore { get; }
    public string LastUpdated { get; }
    public string EmployeesText => Employees.ToString("#,0", CultureInfo.InvariantCulture);
    public string HealthScoreText => HealthScore.HasValue ? HealthScore.Value.ToString(CultureInfo.InvariantCulture) : "N/A";
    public int HealthScoreSort => HealthScore ?? -1;

    public bool MatchesSearch(string? search)
    {
        if (string.IsNullOrWhiteSpace(search))
        {
            return true;
        }

        var terms = search.Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        return terms.All(term =>
            ClientName.Contains(term, StringComparison.OrdinalIgnoreCase) ||
            ClientId.Contains(term, StringComparison.OrdinalIgnoreCase) ||
            PrimaryCarrier.Contains(term, StringComparison.OrdinalIgnoreCase) ||
            Status.Contains(term, StringComparison.OrdinalIgnoreCase) ||
            Industry.Contains(term, StringComparison.OrdinalIgnoreCase) ||
            PlanType.Contains(term, StringComparison.OrdinalIgnoreCase) ||
            RenewalDate.Contains(term, StringComparison.OrdinalIgnoreCase));
    }

    public Brush StatusBackground => Status switch
    {
        "Active" => new SolidColorBrush(Color.FromRgb(231, 245, 237)),
        "Inactive" => new SolidColorBrush(Color.FromRgb(238, 241, 244)),
        "Prospect" => new SolidColorBrush(Color.FromRgb(231, 241, 251)),
        _ => Brushes.White
    };

    public Brush StatusBorder => Status switch
    {
        "Active" => new SolidColorBrush(Color.FromRgb(183, 222, 199)),
        "Inactive" => new SolidColorBrush(Color.FromRgb(207, 214, 222)),
        "Prospect" => new SolidColorBrush(Color.FromRgb(185, 213, 238)),
        _ => Brushes.LightGray
    };

    public Brush StatusForeground => Status switch
    {
        "Active" => new SolidColorBrush(Color.FromRgb(31, 111, 67)),
        "Inactive" => new SolidColorBrush(Color.FromRgb(72, 85, 99)),
        "Prospect" => new SolidColorBrush(Color.FromRgb(31, 95, 150)),
        _ => Brushes.Black
    };

    public Brush HealthScoreBrush => HealthScore switch
    {
        >= 75 => new SolidColorBrush(Color.FromRgb(22, 163, 74)),
        >= 50 => new SolidColorBrush(Color.FromRgb(245, 158, 11)),
        null => new SolidColorBrush(Color.FromRgb(148, 163, 184)),
        _ => new SolidColorBrush(Color.FromRgb(220, 38, 38))
    };
}

public sealed class ReportsFilterCriteria
{
    public string Client { get; init; } = "";
    public string Carrier { get; init; } = "";
    public string Industry { get; init; } = "";
    public string Consultant { get; init; } = "";
    public string Office { get; init; } = "";
    public string Province { get; init; } = "";
    public string EffectiveDate { get; init; } = "";
    public string CalendarYear { get; init; } = "";
    public string RenewalYear { get; init; } = "";
    public string FundingType { get; init; } = "";
    public string Role { get; init; } = "";
}

public sealed class ClientHealthResult
{
    public decimal Score { get; init; }
    public string Status { get; init; } = "Monitor";
    public List<ExecutiveSummaryRow> Components { get; init; } = [];
}

public static class RenewalQueueStatuses
{
    public const string Healthy = "Healthy";
    public const string Watch = "Watch";
    public const string HighRisk = "High Risk";
    public const string MissingData = "Missing Data";
    public const string ReadyToReview = "Ready to Review";
    public const string Negotiated = "Negotiated";
    public const string FinalApproved = "Final Approved";
    public const string Completed = "Completed";
}

public sealed class RenewalQueueItem
{
    public string Id { get; init; } = Guid.NewGuid().ToString("N");
    public long ClientId { get; init; }
    public long RenewalId { get; init; }
    public long? SourceDocumentId { get; init; }
    public string SelectedRenewalVersion { get; init; } = RenewalVersionNames.CarrierProposed;
    public string Client { get; init; } = "";
    public string PolicyNumber { get; init; } = "";
    public string Carrier { get; init; } = "";
    public DateOnly? RenewalDate { get; init; }
    public DateTime DateLoaded { get; init; }
    public decimal? CurrentPremium { get; init; }
    public decimal? RenewalPremium { get; init; }
    public decimal? PremiumChangePercent { get; init; }
    public decimal? EhcLossRatio { get; init; }
    public decimal? TargetLossRatio { get; init; }
    public decimal? PoolingThreshold { get; init; }
    public decimal? PoolingCharge { get; init; }
    public decimal? DentalLossRatio { get; init; }
    public int? Lives { get; init; }
    public string TrendAssumptions { get; init; } = "";
    public string PoolingCharges { get; init; } = "";
    public string Status { get; set; } = RenewalQueueStatuses.MissingData;
    public string ActionLabel { get; set; } = "Review";
    public List<string> QuickInsights { get; init; } = [];
    public List<string> MissingDataItems { get; init; } = [];

    public int StatusSortRank => Status switch
    {
        RenewalQueueStatuses.HighRisk => 1,
        RenewalQueueStatuses.MissingData => 2,
        RenewalQueueStatuses.ReadyToReview => 3,
        RenewalQueueStatuses.Negotiated => 4,
        RenewalQueueStatuses.Watch => 5,
        RenewalQueueStatuses.Healthy => 6,
        RenewalQueueStatuses.FinalApproved => 7,
        RenewalQueueStatuses.Completed => 8,
        _ => 9
    };

    public int RenewalDateSort => RenewalDate?.DayNumber ?? int.MaxValue;
    public string RenewalMonth => RenewalDate?.ToString("MMMM", CultureInfo.InvariantCulture) ?? "Unscheduled";
    public string RenewalDateText => RenewalDate?.ToString("MMM d yyyy", CultureInfo.InvariantCulture) ?? "N/A";
    public string CurrentPremiumText => FormatMoney(CurrentPremium);
    public string RenewalPremiumText => FormatMoney(RenewalPremium);
    public string PremiumChangeText => FormatPercent(PremiumChangePercent);
    public string EhcLossRatioText => FormatPercent(EhcLossRatio, includeSign: false);
    public string PoolingThresholdText => FormatMoney(PoolingThreshold);
    public string PoolingChargeText => FormatPercent(NormalizePercent(PoolingCharge), includeSign: false);
    public string DentalLossRatioText => FormatPercent(DentalLossRatio, includeSign: false);
    public string LivesText => Lives?.ToString("#,0", CultureInfo.InvariantCulture) ?? "N/A";

    public Brush StatusBackground => Status switch
    {
        RenewalQueueStatuses.Healthy => new SolidColorBrush(Color.FromRgb(231, 245, 237)),
        RenewalQueueStatuses.Watch => new SolidColorBrush(Color.FromRgb(255, 244, 216)),
        RenewalQueueStatuses.HighRisk => new SolidColorBrush(Color.FromRgb(253, 232, 232)),
        RenewalQueueStatuses.MissingData => new SolidColorBrush(Color.FromRgb(238, 241, 244)),
        RenewalQueueStatuses.ReadyToReview => new SolidColorBrush(Color.FromRgb(231, 241, 251)),
        RenewalQueueStatuses.Negotiated => new SolidColorBrush(Color.FromRgb(232, 241, 255)),
        RenewalQueueStatuses.FinalApproved => new SolidColorBrush(Color.FromRgb(217, 237, 225)),
        RenewalQueueStatuses.Completed => new SolidColorBrush(Color.FromRgb(223, 241, 229)),
        _ => Brushes.White
    };

    public Brush StatusBorder => Status switch
    {
        RenewalQueueStatuses.Healthy => new SolidColorBrush(Color.FromRgb(183, 222, 199)),
        RenewalQueueStatuses.Watch => new SolidColorBrush(Color.FromRgb(232, 207, 131)),
        RenewalQueueStatuses.HighRisk => new SolidColorBrush(Color.FromRgb(242, 183, 183)),
        RenewalQueueStatuses.MissingData => new SolidColorBrush(Color.FromRgb(207, 214, 222)),
        RenewalQueueStatuses.ReadyToReview => new SolidColorBrush(Color.FromRgb(185, 213, 238)),
        RenewalQueueStatuses.Negotiated => new SolidColorBrush(Color.FromRgb(173, 202, 244)),
        RenewalQueueStatuses.FinalApproved => new SolidColorBrush(Color.FromRgb(136, 194, 157)),
        RenewalQueueStatuses.Completed => new SolidColorBrush(Color.FromRgb(159, 208, 177)),
        _ => Brushes.LightGray
    };

    public Brush StatusForeground => Status switch
    {
        RenewalQueueStatuses.Healthy => new SolidColorBrush(Color.FromRgb(31, 111, 67)),
        RenewalQueueStatuses.Watch => new SolidColorBrush(Color.FromRgb(122, 86, 0)),
        RenewalQueueStatuses.HighRisk => new SolidColorBrush(Color.FromRgb(165, 35, 35)),
        RenewalQueueStatuses.MissingData => new SolidColorBrush(Color.FromRgb(72, 85, 99)),
        RenewalQueueStatuses.ReadyToReview => new SolidColorBrush(Color.FromRgb(31, 95, 150)),
        RenewalQueueStatuses.Negotiated => new SolidColorBrush(Color.FromRgb(30, 83, 145)),
        RenewalQueueStatuses.FinalApproved => new SolidColorBrush(Color.FromRgb(19, 87, 51)),
        RenewalQueueStatuses.Completed => new SolidColorBrush(Color.FromRgb(24, 91, 54)),
        _ => Brushes.Black
    };

    private static string FormatMoney(decimal? value)
    {
        return value.HasValue
            ? value.Value.ToString("$#,0", CultureInfo.InvariantCulture)
            : "N/A";
    }

    private static string FormatPercent(decimal? value, bool includeSign = true)
    {
        if (!value.HasValue)
        {
            return "N/A";
        }

        return includeSign && value.Value >= 0m
            ? $"+{value.Value:0.0}%"
            : $"{value.Value:0.0}%";
    }

    private static decimal? NormalizePercent(decimal? value)
    {
        if (!value.HasValue)
        {
            return null;
        }

        return value.Value is > 0m and <= 1m ? value.Value * 100m : value.Value;
    }
}

public sealed class UnderwritingBenefitRow
{
    public string BenefitLine { get; init; } = "";
    public decimal? CurrentPremium { get; init; }
    public decimal? CurrentClaims { get; init; }
    public decimal? LossRatio { get; init; }
    public decimal? Trend { get; init; }
    public decimal? TargetLossRatio { get; init; }
    public decimal? ProjectedClaims { get; init; }
    public decimal? ModelledPremium { get; init; }
    public decimal? ProposedPremium { get; init; }
    public decimal? ProposedChangePercent { get; init; }
    public decimal? RequiredChangePercent { get; init; }
    public decimal? GapToModel { get; init; }
    public string RiskStatus { get; init; } = RenewalQueueStatuses.MissingData;

    public string CurrentPremiumText => FormatMoney(CurrentPremium);
    public string CurrentClaimsText => FormatMoney(CurrentClaims);
    public string LossRatioText => FormatPercent(LossRatio, includeSign: false);
    public string TrendText => FormatRatePercent(Trend);
    public string TargetLossRatioText => FormatRatePercent(TargetLossRatio);
    public string ModelledPremiumText => FormatMoney(ModelledPremium);
    public string ProposedPremiumText => FormatMoney(ProposedPremium);
    public string GapToModelText => FormatMoney(GapToModel);

    private static string FormatMoney(decimal? value)
    {
        return value.HasValue ? value.Value.ToString("$#,0", CultureInfo.InvariantCulture) : "N/A";
    }

    private static string FormatPercent(decimal? value, bool includeSign = true)
    {
        if (!value.HasValue)
        {
            return "N/A";
        }

        var percent = value.Value is > 0m and <= 1m ? value.Value * 100m : value.Value;
        return includeSign && percent >= 0m ? $"+{percent:0.0}%" : $"{percent:0.0}%";
    }

    private static string FormatRatePercent(decimal? value)
    {
        return value.HasValue ? $"{UnderwritingEngine.NormalizeRate(value).GetValueOrDefault() * 100m:0.0}%" : "N/A";
    }
}

public sealed class NegotiationModelReferenceRow
{
    public string BenefitLine { get; init; } = "";
    public decimal? CurrentPremium { get; init; }
    public decimal? CurrentClaims { get; init; }
    public decimal? TrendedClaims { get; init; }
    public decimal? ProjectedClaims { get; init; }
    public decimal? RequiredPremium { get; init; }
    public decimal? ModelledChangePercent { get; init; }
    public decimal? ProjectedLossRatio { get; init; }

    public string CurrentPremiumText => FormatMoney(CurrentPremium);
    public string CurrentClaimsText => FormatMoney(CurrentClaims);
    public string TrendedClaimsText => FormatMoney(TrendedClaims);
    public string ProjectedClaimsText => FormatMoney(ProjectedClaims);
    public string RequiredPremiumText => FormatMoney(RequiredPremium);
    public string ModelledChangeText => FormatPercent(ModelledChangePercent);
    public string ProjectedLossRatioText => FormatPercent(ProjectedLossRatio, includeSign: false);

    public static NegotiationModelReferenceRow From(NegotiationRateRow row, UnderwritingCalculationResult calculation)
    {
        return new NegotiationModelReferenceRow
        {
            BenefitLine = row.BenefitLine,
            CurrentPremium = row.CurrentPremium,
            CurrentClaims = calculation.CurrentClaims,
            TrendedClaims = calculation.TrendedClaims,
            ProjectedClaims = calculation.ProjectedClaims,
            RequiredPremium = calculation.RequiredPremium,
            ModelledChangePercent = UnderwritingEngine.CalculatePercentChange(row.CurrentPremium, calculation.RequiredPremium),
            ProjectedLossRatio = UnderwritingEngine.CalculateLossRatioPercent(calculation.ProjectedClaims, row.ProposedPremium)
        };
    }

    private static string FormatMoney(decimal? value)
    {
        return value.HasValue ? value.Value.ToString("$#,0", CultureInfo.InvariantCulture) : "N/A";
    }

    private static string FormatPercent(decimal? value, bool includeSign = true)
    {
        if (!value.HasValue)
        {
            return "N/A";
        }

        var percent = value.Value is > 0m and <= 1m ? value.Value * 100m : value.Value;
        return includeSign && percent >= 0m ? $"+{percent:0.0}%" : $"{percent:0.0}%";
    }
}

public sealed partial class NegotiationRateRow : INotifyPropertyChanged
{
    private decimal? _currentRate;
    private decimal? _proposedRate;
    private decimal? _currentPremium;
    private decimal? _proposedPremium;
    private decimal? _modelledRate;
    private decimal? _negotiatedRate;
    private decimal? _modelledPremium;
    private decimal? _negotiatedPremium;
    private decimal? _proposedChangePercent;
    private decimal? _modelledChangePercent;
    private decimal? _negotiatedChangePercent;
    private decimal? _savings;
    private string _status = RenewalQueueStatuses.Negotiated;
    private string _notes = "";

    public event PropertyChangedEventHandler? PropertyChanged;

    public string BenefitLine { get; init; } = "";
    public string ClassName { get; init; } = "";
    public string AccountClass { get; init; } = "";
    public string RateBasis { get; init; } = "";
    public decimal? Volume { get; init; }
    public int? Lives { get; init; }
    public decimal? CurrentClaims { get; set; }
    public decimal? CurrentRate
    {
        get => _currentRate;
        set => SetValue(ref _currentRate, value);
    }
    public decimal? ProposedRate
    {
        get => _proposedRate;
        set => SetValue(ref _proposedRate, value);
    }
    public decimal? ChangeInRisk { get; set; }
    public decimal? CalculatedRate { get; set; }
    public decimal? CalculatedPremium { get; set; }
    public decimal? CurrentPremium
    {
        get => _currentPremium;
        set => SetValue(ref _currentPremium, value);
    }
    public decimal? ProposedPremium
    {
        get => _proposedPremium;
        set => SetValue(ref _proposedPremium, value);
    }
    public bool UseSavedModelValues { get; init; }
    public string RateValueBasis { get; set; } = RateValueBasisNames.CarrierStated;
    public string ClientPresentationRateBasis { get; set; } = RateValueBasisNames.CarrierStated;
    public decimal? CurrentNetRate { get; set; }
    public decimal? CurrentGrossRate { get; set; }
    public decimal? ProposedNetRate { get; set; }
    public decimal? ProposedGrossRate { get; set; }
    public decimal? VersionNetRate { get; set; }
    public decimal? VersionGrossRate { get; set; }
    public string RateScheduleSourceFile { get; set; } = "";
    public string RateScheduleSourceSheet { get; set; } = "";
    public int? RateScheduleSourceRow { get; set; }
    public string SourceReference { get; set; } = "";

    public decimal? ModelledRate
    {
        get => _modelledRate;
        set => SetValue(ref _modelledRate, value);
    }

    public decimal? NegotiatedRate
    {
        get => _negotiatedRate;
        set => SetValue(ref _negotiatedRate, value);
    }

    public decimal? ModelledPremium
    {
        get => _modelledPremium;
        set => SetValue(ref _modelledPremium, value);
    }

    public decimal? NegotiatedPremium
    {
        get => _negotiatedPremium;
        set => SetValue(ref _negotiatedPremium, value);
    }

    public decimal? ProposedChangePercent
    {
        get => _proposedChangePercent;
        set => SetValue(ref _proposedChangePercent, value);
    }

    public decimal? ModelledChangePercent
    {
        get => _modelledChangePercent;
        set => SetValue(ref _modelledChangePercent, value);
    }

    public decimal? NegotiatedChangePercent
    {
        get => _negotiatedChangePercent;
        set => SetValue(ref _negotiatedChangePercent, value);
    }

    public decimal? Savings
    {
        get => _savings;
        set => SetValue(ref _savings, value);
    }

    public string Status
    {
        get => _status;
        set => SetValue(ref _status, value);
    }

    public string Notes
    {
        get => _notes;
        set => SetValue(ref _notes, value);
    }

    public string CurrentRateText => FormatRate(CurrentRate);
    public string ChangeInRiskText => ChangeInRisk.HasValue
        ? ChangeInRisk.Value.ToString("0.00%", CultureInfo.InvariantCulture)
        : "—";
    public string CalculatedRateText => CalculatedRate.HasValue
        ? CalculatedRate.Value.ToString("0.000", CultureInfo.InvariantCulture)
        : "—";
    public string CalculatedPremiumText => CalculatedPremium.HasValue
        ? CalculatedPremium.Value.ToString("$#,##0.00", CultureInfo.InvariantCulture)
        : "—";
    public string ProposedRateText => FormatRate(ProposedRate);
    public string ModelledRateText => CanEnterNegotiatedRate ? FormatRate(ModelledRate) : "N/A";
    public string CurrentPremiumText => FormatMoney(CurrentPremium);
    public string ProposedPremiumText => FormatMoney(ProposedPremium);
    public string ModelledPremiumText => FormatMoney(ModelledPremium);
    public string ProposedChangeText => FormatPercent(ProposedChangePercent);
    public string ModelledChangeText => FormatPercent(ModelledChangePercent);
    public string NegotiatedChangeText => FormatPercent(NegotiatedChangePercent);
    public string SavingsText => FormatMoney(Savings);
    public bool HasNetAndGrossRates =>
        CurrentNetRate.HasValue || CurrentGrossRate.HasValue ||
        ProposedNetRate.HasValue || ProposedGrossRate.HasValue ||
        VersionNetRate.HasValue || VersionGrossRate.HasValue;
    public bool HasCompleteNetAndGrossRates =>
        CurrentNetRate.HasValue && CurrentGrossRate.HasValue &&
        ProposedNetRate.HasValue && ProposedGrossRate.HasValue &&
        VersionNetRate.HasValue && (VersionGrossRate.HasValue || NegotiatedRate.HasValue);
    public bool IsCompleteImportedRateOnlyRow => HasCompleteNetAndGrossRates && !CurrentPremium.HasValue && !ProposedPremium.HasValue;
    public string RateScheduleSourceDisplay => string.IsNullOrWhiteSpace(RateScheduleSourceFile)
        ? SourceReference
        : $"{RateScheduleSourceFile} | {RateScheduleSourceSheet} | row {RateScheduleSourceRow}";
    public bool CanEnterNegotiatedRate => HasUsableRateBasis(RateBasis, CurrentRate, ProposedRate, Volume);
    public bool CannotEnterNegotiatedRate => !CanEnterNegotiatedRate;
    public string NegotiatedRateHelp => CanEnterNegotiatedRate
        ? "Enter the final negotiated rate for this benefit line."
        : "This row only has composite premium data. Enter the final negotiated premium instead.";

    public NegotiationRateSaveRecord ToSaveRecord()
    {
        return new NegotiationRateSaveRecord
        {
            BenefitLine = BenefitLine,
            ClassName = ClassName,
            AccountClass = AccountClass,
            RateBasis = RateBasis,
            CurrentRate = CurrentRate,
            ProposedRate = ProposedRate,
            ModelledRate = CanEnterNegotiatedRate ? ModelledRate : null,
            NegotiatedRate = CanEnterNegotiatedRate ? NegotiatedRate : null,
            CurrentPremium = CurrentPremium,
            ProposedPremium = ProposedPremium,
            ModelledPremium = ModelledPremium,
            NegotiatedPremium = NegotiatedPremium,
            ProposedChangePct = PercentToDecimal(ProposedChangePercent),
            ModelledChangePct = PercentToDecimal(ModelledChangePercent),
            NegotiatedChangePct = PercentToDecimal(NegotiatedChangePercent),
            Savings = Savings,
            Status = Status,
            Notes = Notes,
            Source = string.IsNullOrWhiteSpace(SourceReference) ? "Renewal Negotiation Workbench" : SourceReference,
            RateValueBasis = RateValueBasis,
            ClientPresentationRateBasis = ClientPresentationRateBasis,
            CurrentNetRate = CurrentNetRate,
            CurrentGrossRate = CurrentGrossRate,
            ProposedNetRate = ProposedNetRate,
            ProposedGrossRate = ProposedGrossRate,
            VersionNetRate = VersionNetRate,
            VersionGrossRate = ClientPresentationRateBasis.Equals(RateValueBasisNames.Gross, StringComparison.OrdinalIgnoreCase)
                ? NegotiatedRate ?? VersionGrossRate
                : VersionGrossRate,
            RateScheduleSourceFile = RateScheduleSourceFile,
            RateScheduleSourceSheet = RateScheduleSourceSheet,
            RateScheduleSourceRow = RateScheduleSourceRow,
            Volume = Volume,
            Lives = Lives
        };
    }

    private void SetValue<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value))
        {
            return;
        }

        field = value;
        OnPropertyChanged(propertyName);
        if (propertyName == nameof(CurrentRate) || propertyName == nameof(ProposedRate))
        {
            OnPropertyChanged(propertyName == nameof(CurrentRate) ? nameof(CurrentRateText) : nameof(ProposedRateText));
            OnPropertyChanged(propertyName == nameof(CurrentRate) ? nameof(EditableCurrentRate) : nameof(EditableProposedRate));
            OnPropertyChanged(nameof(CanEnterNegotiatedRate));
            OnPropertyChanged(nameof(CannotEnterNegotiatedRate));
        }
        if (propertyName == nameof(CurrentPremium) || propertyName == nameof(ProposedPremium))
        {
            OnPropertyChanged(propertyName == nameof(CurrentPremium) ? nameof(CurrentPremiumText) : nameof(ProposedPremiumText));
        }
        OnPropertyChanged(nameof(ModelledRateText));
        OnPropertyChanged(nameof(ModelledPremiumText));
        OnPropertyChanged(nameof(ProposedChangeText));
        OnPropertyChanged(nameof(ModelledChangeText));
        OnPropertyChanged(nameof(NegotiatedChangeText));
        OnPropertyChanged(nameof(SavingsText));
        OnPropertyChanged(nameof(NegotiatedRateHelp));
    }

    private void OnPropertyChanged(string? propertyName)
    {
        if (!string.IsNullOrWhiteSpace(propertyName))
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    private static decimal? PercentToDecimal(decimal? value)
    {
        return value.HasValue ? decimal.Round(value.Value / 100m, 6) : null;
    }

    private static string FormatRate(decimal? value)
    {
        return value.HasValue ? value.Value.ToString("0.0000", CultureInfo.InvariantCulture) : "N/A";
    }

    private static bool HasUsableRateBasis(string rateBasis, decimal? currentRate, decimal? proposedRate, decimal? volume)
    {
        var normalized = (rateBasis ?? "").Trim().ToLowerInvariant();
        if (string.IsNullOrWhiteSpace(normalized) ||
            normalized.Contains("composite", StringComparison.Ordinal) ||
            normalized.Contains("premium", StringComparison.Ordinal))
        {
            return false;
        }

        return currentRate.HasValue ||
               proposedRate.HasValue ||
               volume.HasValue ||
               normalized.Contains("single", StringComparison.Ordinal) ||
               normalized.Contains("family", StringComparison.Ordinal) ||
               normalized.Contains("volume", StringComparison.Ordinal) ||
               normalized.Contains("salary", StringComparison.Ordinal);
    }

    private static string FormatMoney(decimal? value)
    {
        return value.HasValue ? value.Value.ToString("$#,0", CultureInfo.InvariantCulture) : "N/A";
    }

    private static string FormatPercent(decimal? value)
    {
        if (!value.HasValue)
        {
            return "N/A";
        }

        return value.Value >= 0m ? $"+{value.Value:0.0}%" : $"{value.Value:0.0}%";
    }
}

public sealed class NegotiationClaimsRow
{
    public string MonthText { get; init; } = "";
    public string BenefitLine { get; init; } = "";
    public string PremiumText { get; init; } = "";
    public string PaidClaimsText { get; init; } = "";
    public string IncurredClaimsText { get; init; } = "";
    public string LossRatioText { get; init; } = "";
    public string LivesText { get; init; } = "";

    public static NegotiationClaimsRow FromRecord(MonthlyClaimsRecord record)
    {
        return new NegotiationClaimsRow
        {
            MonthText = record.Month?.ToString("MMM yyyy", CultureInfo.InvariantCulture) ?? "N/A",
            BenefitLine = string.IsNullOrWhiteSpace(record.BenefitLine) ? "Overall" : record.BenefitLine,
            PremiumText = FormatMoney(record.Premium),
            PaidClaimsText = FormatMoney(record.PaidClaims),
            IncurredClaimsText = FormatMoney(record.IncurredClaims),
            LossRatioText = FormatPercent(record.LossRatio),
            LivesText = record.Lives?.ToString("#,0", CultureInfo.InvariantCulture) ?? "N/A"
        };
    }

    private static string FormatMoney(decimal? value)
    {
        return value.HasValue ? value.Value.ToString("$#,0", CultureInfo.InvariantCulture) : "N/A";
    }

    private static string FormatPercent(decimal? value)
    {
        if (!value.HasValue)
        {
            return "N/A";
        }

        var percent = value.Value is > 0m and <= 1m ? value.Value * 100m : value.Value;
        return $"{percent:0.0}%";
    }
}

public sealed class NegotiationAssumptionValues
{
    public decimal Trend { get; init; }
    public decimal TargetLossRatio { get; init; }
    public decimal PoolingCharge { get; init; }
    public decimal? PoolingThreshold { get; init; }
    public decimal Ibnr { get; init; }
    public decimal AdminFee { get; init; }
    public decimal RiskCharge { get; init; }
    public int? Lives { get; init; }
    public decimal? Volume { get; init; }
    public decimal CredibilityFactor { get; init; }
    public decimal ManualAdjustment { get; init; }

    public NegotiationAssumptionSaveRecord ToSaveRecord()
    {
        return new NegotiationAssumptionSaveRecord
        {
            TrendPct = Trend,
            TargetLossRatioPct = TargetLossRatio,
            PoolingChargePct = PoolingCharge,
            PoolingThreshold = PoolingThreshold,
            IbnrPct = Ibnr,
            AdminFeePct = AdminFee,
            RiskChargePct = RiskCharge,
            Lives = Lives,
            Volume = Volume,
            CredibilityFactor = CredibilityFactor,
            ManualAdjustmentPct = ManualAdjustment
        };
    }
}

public sealed class CarrierAssumptionLibraryRow
{
    public CarrierAssumptionLibraryRow(IReadOnlyCollection<CarrierAssumptionRecord> records)
    {
        var record = records.First();
        Carrier = FormatCarriers(records.Select(item => item.Carrier));
        BlockName = string.IsNullOrWhiteSpace(record.BlockName) ? "Standard Manual" : record.BlockName;
        BenefitLine = record.BenefitLine;
        FundingType = record.FundingType;
        TargetLossRatio = record.TargetLossRatio;
        TlrBasis = TlrBasisOptions.Normalize(record.TlrBasis);
        PremiumTax = record.PremiumTax;
        Commission = record.Commission;
        AdminMargin = record.AdminMargin;
        RiskMargin = record.RiskMargin;
        Trend = record.Trend;
        CredibilityFactor = record.CredibilityFactor;
        PoolingThreshold = record.PoolingThreshold;
        PoolingCharge = record.PoolingCharge;
        ReserveIbnr = record.ReserveIbnr;
        EffectiveDate = LatestDate(records.Select(item => item.EffectiveDate));
        Status = records.All(item => item.Status.Equals("Active", StringComparison.OrdinalIgnoreCase)) ? "Active" : "Review";
        Notes = FormatNotes(records.Select(item => item.Notes));
    }

    public string Carrier { get; }
    public string BlockName { get; }
    public string BenefitLine { get; }
    public string FundingType { get; }
    public decimal? TargetLossRatio { get; }
    public string TargetLossRatioText => FormatPercent(TargetLossRatio);
    public string TlrBasis { get; }
    public decimal? PremiumTax { get; }
    public string PremiumTaxText => FormatPercent(PremiumTax);
    public decimal? Commission { get; }
    public string CommissionText => Commission.HasValue ? FormatPercent(Commission) : "Review";
    public decimal? AdminMargin { get; }
    public decimal? RiskMargin { get; }
    public string TotalLoadText => FormatPercent(SumNullable(PremiumTax, Commission, AdminMargin, RiskMargin));
    public decimal? Trend { get; }
    public string TrendText => FormatPercent(Trend);
    public decimal? CredibilityFactor { get; }
    public string CredibilityText => FormatPercent(CredibilityFactor);
    public decimal? PoolingThreshold { get; }
    public string PoolingThresholdText => PoolingThreshold.HasValue ? PoolingThreshold.Value.ToString("$#,##0", CultureInfo.InvariantCulture) : "N/A";
    public decimal? PoolingCharge { get; }
    public string PoolingChargeText => FormatPercent(PoolingCharge);
    public decimal? ReserveIbnr { get; }
    public string ReserveIbnrText => FormatPercent(ReserveIbnr);
    public DateOnly? EffectiveDate { get; }
    public string EffectiveDateText => EffectiveDate?.ToString("MMM d, yyyy", CultureInfo.InvariantCulture) ?? "N/A";
    public string Status { get; }
    public string Notes { get; }

    private static decimal? SumNullable(params decimal?[] values)
    {
        return values.Any(value => value.HasValue)
            ? values.Sum(value => value ?? 0m)
            : null;
    }

    private static string FormatPercent(decimal? value)
    {
        return value.HasValue ? $"{value.Value:0.0}%" : "N/A";
    }

    private static string FormatCarriers(IEnumerable<string> carriers)
    {
        var values = carriers
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(CarrierSortOrder)
            .ThenBy(value => value)
            .ToList();

        return values.Count == 0 ? "N/A" : string.Join(", ", values);
    }

    private static int CarrierSortOrder(string carrier)
    {
        return carrier.ToUpperInvariant() switch
        {
            "EQUITABLE" => 0,
            "RBC" => 1,
            "EMPIRE" => 2,
            _ => 100
        };
    }

    private static DateOnly? LatestDate(IEnumerable<DateOnly?> dates)
    {
        var values = dates.Where(date => date.HasValue).Select(date => date!.Value).ToList();
        return values.Count == 0 ? null : values.Max();
    }

    private static string FormatNotes(IEnumerable<string> notes)
    {
        var values = notes
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        return values.Count == 0 ? "" : string.Join(" | ", values);
    }
}

public sealed class CarrierBlockProfileRow
{
    public CarrierBlockProfileRow(string blockProfile, IReadOnlyCollection<CarrierAssumptionRecord> records)
    {
        BlockProfile = blockProfile;
        Carriers = FormatValues(records.Select(item => item.Carrier), orderPartnershipCarriers: true);
        Funding = FormatValues(records.Select(item => item.FundingType), orderPartnershipCarriers: false);
        DefaultTlrBasis = FormatDefaultTlrBasis(records.Select(item => TlrBasisOptions.Normalize(item.TlrBasis)));
        BenefitRules = records.Select(item => item.BenefitLine).Where(value => !string.IsNullOrWhiteSpace(value)).Distinct(StringComparer.OrdinalIgnoreCase).Count();
        TargetLossRatioRange = FormatPercentRange(records.Select(item => item.TargetLossRatio));
        Notes = BuildNotes(records);
    }

    public string BlockProfile { get; }
    public string Carriers { get; }
    public string Funding { get; }
    public string DefaultTlrBasis { get; }
    public int BenefitRules { get; }
    public string TargetLossRatioRange { get; }
    public string Notes { get; }

    private static string FormatValues(IEnumerable<string> values, bool orderPartnershipCarriers)
    {
        var distinct = values
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Distinct(StringComparer.OrdinalIgnoreCase);

        var sorted = orderPartnershipCarriers
            ? distinct.OrderBy(CarrierSortOrder).ThenBy(value => value).ToList()
            : distinct.OrderBy(value => value).ToList();

        return sorted.Count == 0 ? "N/A" : string.Join(", ", sorted);
    }

    private static string FormatDefaultTlrBasis(IEnumerable<string> values)
    {
        var bases = values
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(value => value)
            .ToList();

        if (bases.Count == 0)
        {
            return "N/A";
        }

        return bases.Count == 1 ? bases[0] : "Mixed by benefit";
    }

    private static string BuildNotes(IReadOnlyCollection<CarrierAssumptionRecord> records)
    {
        var bases = records
            .Select(item => TlrBasisOptions.Normalize(item.TlrBasis))
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (bases.Count > 1)
        {
            return "Benefit-specific TLR basis; review benefit line rules below.";
        }

        var notes = records
            .Select(item => item.Notes)
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        return notes.Count == 0 ? "" : string.Join(" | ", notes.Take(2));
    }

    private static string FormatPercentRange(IEnumerable<decimal?> values)
    {
        var percentages = values.Where(value => value.HasValue).Select(value => value!.Value).ToList();
        if (percentages.Count == 0)
        {
            return "N/A";
        }

        var min = percentages.Min();
        var max = percentages.Max();
        return min == max
            ? $"{min:0.0}%"
            : $"{min:0.0}% - {max:0.0}%";
    }

    private static int CarrierSortOrder(string carrier)
    {
        return carrier.ToUpperInvariant() switch
        {
            "EQUITABLE" => 0,
            "RBC" => 1,
            "EMPIRE" => 2,
            _ => 100
        };
    }
}

public sealed class CarrierAssumptionSummaryRow
{
    public string Carrier { get; init; } = "";
    public int Blocks { get; init; }
    public int BenefitRules { get; init; }
    public string TargetLossRatioRange { get; init; } = "N/A";
    public string LatestEffectiveDate { get; init; } = "N/A";
    public string Status { get; init; } = "Active";
}

public sealed class PlanDesignFactorRow
{
    private readonly PlanDesignFactorRecord _record;

    public PlanDesignFactorRow(PlanDesignFactorRecord record)
    {
        _record = record;
    }

    public long Id => _record.Id;
    public string CarrierName => _record.CarrierName;
    public string BenefitCategory => _record.BenefitCategory;
    public string BenefitLine => _record.BenefitLine;
    public string CurrentDesign => _record.CurrentDesign;
    public string ProposedDesign => _record.ProposedDesign;
    public string ChangeType => _record.ChangeType;
    public string FactorText => _record.Factor.ToString("0.0000", CultureInfo.InvariantCulture);
    public string FactorDirection => _record.FactorDirection;
    public string EffectiveDateText => _record.EffectiveDate?.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) ?? "N/A";
    public string ExpiryDateText => _record.ExpiryDate?.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) ?? "";
    public string SourceDocument => _record.SourceDocument;
    public string Notes => _record.Notes;
    public string ConfidenceDisplay => _record.IsEstimatedPlaceholder
        ? "Estimated Placeholder - not carrier-approved"
        : _record.ConfidenceLevel;

    public string CalculatorOptionText =>
        $"{CarrierName} | {BenefitLine} | {CurrentDesign} to {ProposedDesign} | {FactorText}";

    public PlanDesignFactorRecord ToRecord()
    {
        return new PlanDesignFactorRecord
        {
            Id = _record.Id,
            CarrierName = _record.CarrierName,
            BenefitCategory = _record.BenefitCategory,
            BenefitLine = _record.BenefitLine,
            CurrentDesign = _record.CurrentDesign,
            ProposedDesign = _record.ProposedDesign,
            ChangeType = _record.ChangeType,
            Factor = _record.Factor,
            FactorDirection = _record.FactorDirection,
            EffectiveDate = _record.EffectiveDate,
            ExpiryDate = _record.ExpiryDate,
            SourceDocument = _record.SourceDocument,
            Notes = _record.Notes,
            ConfidenceLevel = _record.ConfidenceLevel,
            CreatedAt = _record.CreatedAt,
            UpdatedAt = _record.UpdatedAt
        };
    }
}

public static class NegotiationEnumerableExtensions
{
    public static decimal? SavingsFallback(this RenewalRateVersionRecord rate)
    {
        return rate.ProposedPremium.HasValue && rate.VersionPremium.HasValue
            ? rate.ProposedPremium.Value - rate.VersionPremium.Value
            : null;
    }

    public static decimal? SumValue<T>(this IEnumerable<T> source, Func<T, decimal?> selector)
    {
        var hasAny = false;
        var total = 0m;
        foreach (var value in source.Select(selector).Where(value => value.HasValue))
        {
            hasAny = true;
            total += value!.Value;
        }

        return hasAny ? decimal.Round(total, 2) : null;
    }

    public static decimal? AverageValue<T>(this IEnumerable<T> source, Func<T, decimal?> selector)
    {
        var values = source
            .Select(selector)
            .Where(value => value.HasValue)
            .Select(value => value!.Value)
            .ToList();

        return values.Count == 0 ? null : decimal.Round(values.Average(), 1);
    }
}

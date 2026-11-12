using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using ClosedXML.Excel;
using Microsoft.Win32;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;

namespace RenewalReviewBuilder.App.Views;

public partial class ClientBillingDashboardView : UserControl
{
    private readonly DatabaseService _databaseService = new();
    private readonly GroupHealthBillingStatementParser _parser = new();
    private readonly ManulifeExcelBillingImportService _manulifeExcelBillingImportService;
    private readonly ObservableCollection<ClientBillingHistoryRow> _historyRows = [];
    private readonly ObservableCollection<ClientBillingDetailRow> _detailRows = [];
    private readonly ObservableCollection<ClientBillingValidationRow> _validationRows = [];
    private readonly ObservableCollection<ClientBillingRateRow> _rateRows = [];
    private readonly ObservableCollection<ClientBillingChangeRow> _changeRows = [];
    private readonly List<string> _selectedFiles = [];
    private readonly List<GroupHealthBillingRow> _allDetailRows = [];
    private readonly List<ClientBillingDetailRow> _allManulifeDetailRows = [];
    private readonly List<GroupHealthBillingEmployeeChange> _allChangeRows = [];
    private ClientWorkspaceViewModel? _workspace;
    private long? _clientIdOverride;
    private bool _isLoading;

    public ClientBillingDashboardView()
    {
        InitializeComponent();
        _manulifeExcelBillingImportService = new ManulifeExcelBillingImportService(_databaseService);
        HistoryGrid.ItemsSource = _historyRows;
        DetailGrid.ItemsSource = _detailRows;
        ValidationGrid.ItemsSource = _validationRows;
        RateGrid.ItemsSource = _rateRows;
        ChangeGrid.ItemsSource = _changeRows;
        ConfigureFilters();
        ResetSummaryCards();
    }

    public async Task LoadDashboardAsync(ClientWorkspaceViewModel workspace)
    {
        _workspace = workspace;
        _clientIdOverride = workspace.DatabaseClientId;
        HeaderSubtitleTextBlock.Text = $"{workspace.ClientName} | Carrier: {Display(workspace.PrimaryCarrier)} | Route: {workspace.RouteText}";
        SelectedFilesTextBlock.Text = "No billing file selected.";
        ImportStatusTextBlock.Text = "Ready";
        await RefreshHistoryAsync(selectMostRecent: true);
    }

    public void LoadEmpty(ClientWorkspaceViewModel workspace)
    {
        _workspace = workspace;
        _clientIdOverride = workspace.DatabaseClientId;
        HeaderSubtitleTextBlock.Text = $"{workspace.ClientName} | Carrier: {Display(workspace.PrimaryCarrier)} | Route: {workspace.RouteText}";
        _historyRows.Clear();
        ClearStatementDetail();
        ResetSummaryCards();
        SelectedFilesTextBlock.Text = "No billing file selected.";
        ImportStatusTextBlock.Text = "No saved client billing yet.";
    }

    private void ConfigureFilters()
    {
        ClassFilterComboBox.ItemsSource = new[] { "All classes" };
        ClassFilterComboBox.SelectedIndex = 0;
        BenefitFilterComboBox.ItemsSource = new[]
        {
            "All benefits",
            "Life",
            "AD&D",
            "LTD",
            "Health",
            "Dental",
            "EFAP",
            "Virtual Healthcare",
            "Flex Admin",
            "Member AD&D"
        };
        BenefitFilterComboBox.SelectedIndex = 0;
        ValidationFilterComboBox.ItemsSource = new[] { "All statuses", "Passed", "Warning", "Failed", "Information" };
        ValidationFilterComboBox.SelectedIndex = 0;
        ChangeFilterComboBox.ItemsSource = new[]
        {
            "All changes",
            "New Employee",
            "Termination",
            "Salary Change",
            "Class Change",
            "Health Tier Change",
            "Dental Tier Change",
            "Coverage Change",
            "Premium Change"
        };
        ChangeFilterComboBox.SelectedIndex = 0;
    }

    private async Task RefreshHistoryAsync(bool selectMostRecent)
    {
        _historyRows.Clear();
        ClearStatementDetail();
        ResetSummaryCards();

        var clientId = CurrentClientId();
        if (clientId is not > 0)
        {
            ImportStatusTextBlock.Text = "This workspace is not linked to a saved SQLite client yet.";
            return;
        }

        try
        {
            var groupHealthHistoryTask = _databaseService.GetGroupHealthBillingHistoryAsync(clientId.Value);
            var manulifeHistoryTask = _databaseService.GetManulifeExcelBillingImportHistoryAsync(clientId.Value);
            await Task.WhenAll(groupHealthHistoryTask, manulifeHistoryTask);

            var history = groupHealthHistoryTask.Result
                .Select(ClientBillingHistoryRow.FromRecord)
                .Concat(manulifeHistoryTask.Result.Select(ClientBillingHistoryRow.FromManulifeRecord))
                .OrderByDescending(row => row.StatementYear)
                .ThenByDescending(row => row.StatementMonth)
                .ThenByDescending(row => row.ImportVersion)
                .ThenByDescending(row => row.BillingStatementId)
                .ToList();
            foreach (var row in history)
            {
                _historyRows.Add(row);
            }

            UpdateSummaryCards(_historyRows.FirstOrDefault());
            ImportStatusTextBlock.Text = _historyRows.Count == 0
                ? "No billing imports have been saved for this client yet."
                : $"{_historyRows.Count:#,0} billing statement/version row(s) loaded.";

            if (selectMostRecent && _historyRows.Count > 0)
            {
                HistoryGrid.SelectedIndex = 0;
            }
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            ImportStatusTextBlock.Text = "Billing history could not load.";
        }
    }

    private async Task LoadStatementDetailsAsync(ClientBillingHistoryRow selected)
    {
        if (selected.BillingStatementId <= 0)
        {
            ClearStatementDetail();
            return;
        }

        if (selected.IsManulifeExcel)
        {
            await LoadManulifeStatementDetailsAsync(selected);
            return;
        }

        try
        {
            var rowsTask = _databaseService.GetGroupHealthBillingRowsAsync(selected.BillingStatementId);
            var ratesTask = _databaseService.GetGroupHealthRateTableLinesAsync(selected.BillingStatementId);
            var validationTask = _databaseService.GetGroupHealthValidationResultsAsync(selected.BillingStatementId);
            var changesTask = _databaseService.GetGroupHealthEmployeeChangesAsync(selected.BillingStatementId);
            await Task.WhenAll(rowsTask, ratesTask, validationTask, changesTask);

            _allDetailRows.Clear();
            _allDetailRows.AddRange(rowsTask.Result);
            _allChangeRows.Clear();
            _allChangeRows.AddRange(changesTask.Result);

            _rateRows.Clear();
            AddDistinctRateRows(ratesTask.Result.Select(ClientBillingRateRow.FromRecord));

            _validationRows.Clear();
            foreach (var row in validationTask.Result.Select(ClientBillingValidationRow.FromRecord))
            {
                _validationRows.Add(row);
            }

            _changeRows.Clear();
            foreach (var row in changesTask.Result.Select(ClientBillingChangeRow.FromRecord))
            {
                _changeRows.Add(row);
            }

            UpdateDetailFilters();
            ApplyDetailFilters();
            StatementDetailTabControl.SelectedIndex = 0;
            UpdateIntegrityReport(validationTask.Result, changesTask.Result, rowsTask.Result);
            ImportStatusTextBlock.Text = $"Loaded {rowsTask.Result.Count:#,0} billing row(s), {validationTask.Result.Count:#,0} validation check(s), and {ratesTask.Result.Count:#,0} rate line(s).";
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            ClearStatementDetail();
            ImportStatusTextBlock.Text = "Statement detail could not load.";
        }
    }

    private async Task LoadManulifeStatementDetailsAsync(ClientBillingHistoryRow selected)
    {
        try
        {
            var employeesTask = _databaseService.GetManulifeExcelBillingEmployeeRecordsAsync(selected.BillingStatementId);
            var exceptionsTask = _databaseService.GetManulifeExcelBillingExceptionsAsync(selected.BillingStatementId);
            var ratesTask = _databaseService.GetManulifeExcelBillingRateSummariesAsync(selected.BillingStatementId);
            await Task.WhenAll(employeesTask, exceptionsTask, ratesTask);

            _allDetailRows.Clear();
            _allManulifeDetailRows.Clear();
            _allManulifeDetailRows.AddRange(employeesTask.Result.Select(ClientBillingDetailRow.FromManulifeRecord));
            _allChangeRows.Clear();
            _rateRows.Clear();
            AddDistinctRateRows(ratesTask.Result.Select(ClientBillingRateRow.FromManulifeSummary));
            _changeRows.Clear();
            _validationRows.Clear();
            foreach (var row in exceptionsTask.Result.Select(ClientBillingValidationRow.FromManulifeRecord))
            {
                _validationRows.Add(row);
            }

            UpdateDetailFilters();
            ApplyDetailFilters();
            StatementDetailTabControl.SelectedIndex = 0;
            UpdateManulifeIntegrityReport(selected, exceptionsTask.Result);
            ImportStatusTextBlock.Text = $"Loaded {employeesTask.Result.Count:#,0} Manulife employee row(s), {ratesTask.Result.Count:#,0} distinct rate(s), and {exceptionsTask.Result.Count:#,0} validation item(s).";
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            ClearStatementDetail();
            ImportStatusTextBlock.Text = "Manulife billing detail could not load.";
        }
    }

    private void UpdateDetailFilters()
    {
        _isLoading = true;
        var classes = _allDetailRows.Select(row => row.ClassCode)
            .Concat(_allManulifeDetailRows.Select(row => row.ClassCode))
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
            .ToList();
        classes.Insert(0, "All classes");
        ClassFilterComboBox.ItemsSource = classes;
        ClassFilterComboBox.SelectedIndex = 0;
        BenefitFilterComboBox.SelectedIndex = 0;
        ValidationFilterComboBox.SelectedIndex = 0;
        ChangeFilterComboBox.SelectedIndex = 0;
        _isLoading = false;
    }

    private void ApplyDetailFilters()
    {
        if (_isLoading)
        {
            return;
        }

        var classFilter = ClassFilterComboBox.SelectedItem as string ?? "All classes";
        var benefitFilter = BenefitFilterComboBox.SelectedItem as string ?? "All benefits";
        var validationFilter = ValidationFilterComboBox.SelectedItem as string ?? "All statuses";
        var changeFilter = ChangeFilterComboBox.SelectedItem as string ?? "All changes";
        var changedCertificates = changeFilter.Equals("All changes", StringComparison.OrdinalIgnoreCase)
            ? new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            : _allChangeRows
                .Where(change => change.ChangeType.Equals(changeFilter, StringComparison.OrdinalIgnoreCase))
                .Select(change => change.CertificateId)
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

        if (_allManulifeDetailRows.Count > 0)
        {
            var manulifeRows = _allManulifeDetailRows.Where(row =>
                (classFilter == "All classes" || row.ClassCode.Equals(classFilter, StringComparison.OrdinalIgnoreCase)) &&
                (validationFilter == "All statuses" || row.ValidationStatus.Equals(validationFilter, StringComparison.OrdinalIgnoreCase)) &&
                (changeFilter == "All changes" || changedCertificates.Contains(row.CertificateId)));
            _detailRows.Clear();
            foreach (var row in manulifeRows)
            {
                _detailRows.Add(row);
            }
            return;
        }

        var rows = _allDetailRows.Where(row =>
            (classFilter == "All classes" || row.ClassCode.Equals(classFilter, StringComparison.OrdinalIgnoreCase)) &&
            (benefitFilter == "All benefits" || RowHasBenefit(row, benefitFilter)) &&
            (validationFilter == "All statuses" || row.ValidationStatus.Equals(validationFilter, StringComparison.OrdinalIgnoreCase)) &&
            (changeFilter == "All changes" || changedCertificates.Contains(row.CertificateId)));

        _detailRows.Clear();
        foreach (var row in rows.Select(ClientBillingDetailRow.FromRecord))
        {
            _detailRows.Add(row);
        }
    }

    private static bool RowHasBenefit(GroupHealthBillingRow row, string benefit)
    {
        return benefit switch
        {
            "Life" => HasValue(row.LifeVolume) || HasValue(row.LifePremium),
            "AD&D" => HasValue(row.AdAndDVolume) || HasValue(row.AdAndDPremium),
            "LTD" => HasValue(row.LtdVolume) || HasValue(row.LtdPremium),
            "Health" => !string.IsNullOrWhiteSpace(row.HealthTier) || HasValue(row.HealthPremium),
            "Dental" => !string.IsNullOrWhiteSpace(row.DentalTier) || HasValue(row.DentalPremium),
            "EFAP" => HasValue(row.EfapPremium),
            "Virtual Healthcare" => HasValue(row.VirtualHealthcarePremium),
            "Flex Admin" => HasValue(row.FlexAdminFee),
            "Member AD&D" => HasValue(row.MemberAdAndDVolume) || HasValue(row.MemberAdAndDPremium) ||
                              HasValue(row.MemberAdAndDSpVolume) || HasValue(row.MemberAdAndDSpPremium) ||
                              HasValue(row.MemberAdAndDChVolume) || HasValue(row.MemberAdAndDChPremium),
            _ => true
        };
    }

    private static bool HasValue(decimal? value)
    {
        return value.HasValue && value.Value != 0m;
    }

    private void AddDistinctRateRows(IEnumerable<ClientBillingRateRow> rows)
    {
        foreach (var group in rows.GroupBy(row => new
                 {
                     row.Benefit,
                     row.ClassCode,
                     row.PlanCode,
                     row.Tier,
                     row.VolumeBasis,
                     row.EmployeeRate,
                     row.EmployerRate,
                     row.CarrierRate,
                     row.RateStatus,
                     row.Source,
                     row.FundingSplitUnavailable
                 }))
        {
            _rateRows.Add(group.First().WithRecordCount(group.Sum(item => item.RecordCount)));
        }
    }

    private void UpdateSummaryCards(ClientBillingHistoryRow? latest)
    {
        StatementCountTextBlock.Text = _historyRows.Count.ToString("#,0", CultureInfo.InvariantCulture);
        EmployeeCountTextBlock.Text = latest?.EmployeeCountText ?? "0";
        EePremiumTextBlock.Text = latest is null || string.IsNullOrWhiteSpace(latest.TotalEmployeePremiumText) ? "N/A" : latest.TotalEmployeePremiumText;
        ErPremiumTextBlock.Text = latest is null || string.IsNullOrWhiteSpace(latest.TotalEmployerPremiumText) ? "N/A" : latest.TotalEmployerPremiumText;
        TotalPremiumTextBlock.Text = latest?.TotalPremiumText ?? "$0.00";
        IntegrityScoreTextBlock.Text = latest?.IntegrityScoreText ?? "N/A";
        IntegrityStatusTextBlock.Text = latest?.ImportStatus ?? "No billing imported";
    }

    private void UpdateIntegrityReport(
        IReadOnlyCollection<GroupHealthBillingValidationResult> validation,
        IReadOnlyCollection<GroupHealthBillingEmployeeChange> changes,
        IReadOnlyCollection<GroupHealthBillingRow> rows)
    {
        var passed = validation.Count(row => row.Status.Equals(GroupHealthBillingValidationStatuses.Passed, StringComparison.OrdinalIgnoreCase));
        var warnings = validation.Count(row => row.Status.Equals(GroupHealthBillingValidationStatuses.Warning, StringComparison.OrdinalIgnoreCase));
        var failed = validation.Count(row => row.Status.Equals(GroupHealthBillingValidationStatuses.Failed, StringComparison.OrdinalIgnoreCase));
        var duplicates = validation.Count(row => row.CheckName.Contains("Duplicate", StringComparison.OrdinalIgnoreCase));
        var selected = HistoryGrid.SelectedItem as ClientBillingHistoryRow;

        IntegrityReportStatusTextBlock.Text = selected is null
            ? "No statement selected"
            : $"{selected.ImportStatus} | {selected.IntegrityScoreText} | {rows.Count:#,0} row(s)";
        PassedChecksTextBlock.Text = passed.ToString("#,0", CultureInfo.InvariantCulture);
        WarningChecksTextBlock.Text = warnings.ToString("#,0", CultureInfo.InvariantCulture);
        FailedChecksTextBlock.Text = failed.ToString("#,0", CultureInfo.InvariantCulture);
        DuplicateRowsTextBlock.Text = duplicates.ToString("#,0", CultureInfo.InvariantCulture);
        NewHireTextBlock.Text = CountChanges(changes, "New Employee");
        TerminationTextBlock.Text = CountChanges(changes, "Termination");
        SalaryChangeTextBlock.Text = CountChanges(changes, "Salary Change");
        var coverageTier = changes.Count(row =>
            row.ChangeType.Contains("Coverage", StringComparison.OrdinalIgnoreCase) ||
            row.ChangeType.Contains("Tier", StringComparison.OrdinalIgnoreCase));
        CoverageChangeTextBlock.Text = coverageTier.ToString("#,0", CultureInfo.InvariantCulture);
    }

    private static string CountChanges(IReadOnlyCollection<GroupHealthBillingEmployeeChange> changes, string changeType)
    {
        return changes.Count(row => row.ChangeType.Equals(changeType, StringComparison.OrdinalIgnoreCase))
            .ToString("#,0", CultureInfo.InvariantCulture);
    }

    private void UpdateManulifeIntegrityReport(ClientBillingHistoryRow selected, IReadOnlyCollection<CarrierBillingException> exceptions)
    {
        var warnings = exceptions.Count(row => row.Severity.Equals(CarrierBillingExceptionSeverities.Warning, StringComparison.OrdinalIgnoreCase));
        var errors = exceptions.Count(row => row.Severity.Equals(CarrierBillingExceptionSeverities.Error, StringComparison.OrdinalIgnoreCase));
        IntegrityReportStatusTextBlock.Text = $"{selected.ImportStatus} | {selected.EmployeeCount:#,0} employee row(s)";
        PassedChecksTextBlock.Text = errors == 0 && warnings == 0 ? "1" : "0";
        WarningChecksTextBlock.Text = warnings.ToString("#,0", CultureInfo.InvariantCulture);
        FailedChecksTextBlock.Text = errors.ToString("#,0", CultureInfo.InvariantCulture);
        DuplicateRowsTextBlock.Text = exceptions.Count(row => row.ExceptionType.Contains("duplicate", StringComparison.OrdinalIgnoreCase)).ToString("#,0", CultureInfo.InvariantCulture);
        NewHireTextBlock.Text = "N/A";
        TerminationTextBlock.Text = "N/A";
        SalaryChangeTextBlock.Text = "N/A";
        CoverageChangeTextBlock.Text = "N/A";
    }

    private void ResetSummaryCards()
    {
        StatementCountTextBlock.Text = "0";
        EmployeeCountTextBlock.Text = "0";
        EePremiumTextBlock.Text = "$0.00";
        ErPremiumTextBlock.Text = "$0.00";
        TotalPremiumTextBlock.Text = "$0.00";
        IntegrityScoreTextBlock.Text = "N/A";
        IntegrityStatusTextBlock.Text = "No billing imported";
    }

    private void ClearStatementDetail()
    {
        _allDetailRows.Clear();
        _allManulifeDetailRows.Clear();
        _allChangeRows.Clear();
        _detailRows.Clear();
        _validationRows.Clear();
        _rateRows.Clear();
        _changeRows.Clear();
        IntegrityReportStatusTextBlock.Text = "No statement selected";
        PassedChecksTextBlock.Text = "0";
        WarningChecksTextBlock.Text = "0";
        FailedChecksTextBlock.Text = "0";
        DuplicateRowsTextBlock.Text = "0";
        NewHireTextBlock.Text = "0";
        TerminationTextBlock.Text = "0";
        SalaryChangeTextBlock.Text = "0";
        CoverageChangeTextBlock.Text = "0";
    }

    private void SelectBillingPdfs_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Title = "Select billing files",
            Filter = "Billing files (*.pdf;*.xlsx;*.xlsm;*.csv)|*.pdf;*.xlsx;*.xlsm;*.csv|All files (*.*)|*.*",
            Multiselect = true
        };

        if (dialog.ShowDialog() != true)
        {
            return;
        }

        _selectedFiles.Clear();
        _selectedFiles.AddRange(dialog.FileNames);
        SelectedFilesTextBlock.Text = _selectedFiles.Count == 1
            ? dialog.FileNames[0]
            : $"{_selectedFiles.Count:#,0} billing statement file(s) selected.";
        ImportStatusTextBlock.Text = "Ready to import selected billing.";
    }

    private async void ImportSelected_Click(object sender, RoutedEventArgs e)
    {
        if (_workspace is null)
        {
            MessageBox.Show("Select a client before importing billing.", "Billing Import", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        if (_selectedFiles.Count == 0)
        {
            MessageBox.Show("Select one or more billing files first.", "Billing Import", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var imported = 0;
        var skipped = 0;
        var failed = 0;
        IsEnabled = false;
        try
        {
            foreach (var file in _selectedFiles)
            {
                ImportStatusTextBlock.Text = $"Extracting {Path.GetFileName(file)}...";
                if (ManulifeExcelBillingImportService.IsSupportedFile(file))
                {
                    var outcome = await ImportManulifeExcelBillingAsync(file);
                    imported += outcome == BillingImportOutcome.Saved ? 1 : 0;
                    skipped += outcome == BillingImportOutcome.Skipped ? 1 : 0;
                    failed += outcome == BillingImportOutcome.Failed ? 1 : 0;
                    continue;
                }

                var result = _parser.Extract(file);
                AttachSelectedClient(result);

                var approval = MessageBox.Show(
                    IntegrityPrompt(result),
                    "Billing Integrity Report",
                    MessageBoxButton.YesNoCancel,
                    result.ImportStatus.Equals(GroupHealthBillingImportStatuses.Failed, StringComparison.OrdinalIgnoreCase)
                        ? MessageBoxImage.Warning
                        : MessageBoxImage.Information);

                if (approval == MessageBoxResult.Cancel)
                {
                    skipped++;
                    break;
                }

                if (approval == MessageBoxResult.No)
                {
                    skipped++;
                    continue;
                }

                var save = await SaveWithDuplicateWorkflowAsync(result);
                if (save.Succeeded)
                {
                    imported++;
                    continue;
                }

                failed++;
                MessageBox.Show(save.Message, "Billing Import", MessageBoxButton.OK, MessageBoxImage.Warning);
            }

            ImportStatusTextBlock.Text = $"Import complete. Saved {imported:#,0}, skipped {skipped:#,0}, failed {failed:#,0}.";
            await RefreshHistoryAsync(selectMostRecent: true);
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            ImportStatusTextBlock.Text = "Billing import failed before completion.";
            MessageBox.Show(ex.Message, "Billing Import", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            IsEnabled = true;
        }
    }

    private async Task<BillingImportOutcome> ImportManulifeExcelBillingAsync(string file)
    {
        if (_workspace is null)
        {
            return BillingImportOutcome.Failed;
        }

        var preview = await _manulifeExcelBillingImportService.CreatePreviewAsync(new ManulifeExcelBillingImportRequest
        {
            InputPath = file,
            SelectedClientId = CurrentClientId(),
            SelectedClientName = _workspace.ClientName
        });
        if (!preview.CanConfirm)
        {
            MessageBox.Show(
                ManulifePreviewSummary(preview, includeConfirmationPrompt: false),
                "Manulife Billing Preview",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return BillingImportOutcome.Failed;
        }

        var approval = MessageBox.Show(
            ManulifePreviewSummary(preview, includeConfirmationPrompt: true),
            "Confirm Manulife Billing Import",
            MessageBoxButton.YesNo,
            preview.WarningCount > 0 ? MessageBoxImage.Warning : MessageBoxImage.Information);
        if (approval != MessageBoxResult.Yes)
        {
            return BillingImportOutcome.Skipped;
        }

        var save = await _manulifeExcelBillingImportService.ConfirmAsync(preview, new ManulifeExcelBillingImportRequest
        {
            InputPath = file,
            SelectedClientId = CurrentClientId(),
            SelectedClientName = _workspace.ClientName,
            BillingPeriod = preview.BillingPeriod
        });
        if (!save.Succeeded)
        {
            MessageBox.Show(save.Message, "Manulife Billing Import", MessageBoxButton.OK, MessageBoxImage.Warning);
            return BillingImportOutcome.Failed;
        }

        _clientIdOverride = save.ClientId;
        return BillingImportOutcome.Saved;
    }

    private static string ManulifePreviewSummary(ManulifeExcelBillingImportPreview preview, bool includeConfirmationPrompt)
    {
        var summary = string.Join(
            Environment.NewLine,
            $"File: {preview.SourceFileName}",
            $"Period: {preview.BillingPeriod}",
            $"Employees: {preview.EmployeeRecordCount:#,0}",
            $"Benefit records: {preview.BenefitRecordCount:#,0}",
            $"Current premium: {preview.TotalCurrentMonthPremium:C2}",
            $"Adjustments: {preview.TotalPriorPeriodAdjustment:C2}",
            $"Tax: {preview.TotalSalesTax:C2}",
            $"Premium and tax: {preview.TotalPremiumAndTax:C2}",
            $"Reconciliation difference: {preview.ReconciliationDifference:C2}",
            $"Warnings: {preview.WarningCount:#,0}",
            $"Errors: {preview.ErrorCount:#,0}");
        return includeConfirmationPrompt ? summary + Environment.NewLine + Environment.NewLine + "Save this billing import?" : summary;
    }

    private enum BillingImportOutcome
    {
        Saved,
        Skipped,
        Failed
    }

    private void AttachSelectedClient(GroupHealthBillingExtractionResult result)
    {
        if (_workspace is null)
        {
            return;
        }

        var clientId = CurrentClientId();
        if (clientId is > 0)
        {
            result.ClientId = clientId.Value;
        }

        result.ClientName = _workspace.ClientName;
        if (string.IsNullOrWhiteSpace(result.ClientNumber))
        {
            result.ClientNumber = _workspace.ClientId;
        }
    }

    private long? CurrentClientId()
    {
        return _workspace?.DatabaseClientId is > 0
            ? _workspace.DatabaseClientId
            : _clientIdOverride;
    }

    private async Task<DatabaseSaveResult> SaveWithDuplicateWorkflowAsync(GroupHealthBillingExtractionResult result)
    {
        var save = await _databaseService.SaveGroupHealthBillingStatementAsync(result);
        if (save.Succeeded ||
            !save.Message.Equals("Existing billing found. Replace, save as new version, or cancel.", StringComparison.OrdinalIgnoreCase))
        {
            if (save.Succeeded && save.ClientId is > 0)
            {
                _clientIdOverride = save.ClientId;
            }

            return save;
        }

        var choice = MessageBox.Show(
            "Existing billing found for this client, carrier, and month.\n\nYes = Replace the active statement\nNo = Save as a new version\nCancel = Do not import this file",
            "Existing Billing Found",
            MessageBoxButton.YesNoCancel,
            MessageBoxImage.Question);

        var selectedSave = choice switch
        {
            MessageBoxResult.Yes => await _databaseService.SaveGroupHealthBillingStatementAsync(result, GroupHealthBillingSaveMode.Replace),
            MessageBoxResult.No => await _databaseService.SaveGroupHealthBillingStatementAsync(result, GroupHealthBillingSaveMode.NewVersion),
            _ => DatabaseSaveResult.Failure("Import cancelled.", _databaseService.DatabasePath)
        };
        if (selectedSave.Succeeded && selectedSave.ClientId is > 0)
        {
            _clientIdOverride = selectedSave.ClientId;
        }

        return selectedSave;
    }

    private static string IntegrityPrompt(GroupHealthBillingExtractionResult result)
    {
        var report = result.IntegrityReport;
        return string.Join(
            Environment.NewLine,
            $"File: {result.SourceFileName}",
            $"Statement: {result.StatementMonth:00}/{result.StatementYear}",
            $"Rows: {result.Rows.Count:#,0}",
            $"Employees: {report.TotalEmployeesImported:#,0}",
            $"Total Premium: {FormatMoney(result.StatementGrandTotal)}",
            $"Integrity Score: {result.BillingIntegrityScore:0.##}%",
            $"Status: {result.ImportStatus}",
            $"Passed: {report.PassedChecks:#,0} | Warnings: {report.Warnings:#,0} | Failed: {report.FailedChecks:#,0}",
            "",
            "Approve and save this billing import?");
    }

    private async void DeleteSelected_Click(object sender, RoutedEventArgs e)
    {
        if (HistoryGrid.SelectedItem is not ClientBillingHistoryRow selected)
        {
            MessageBox.Show("Select a billing statement first.", "Delete Billing", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        if (selected.IsManulifeExcel)
        {
            MessageBox.Show("Manulife Excel billing deletion is not available from this screen yet. It remains protected from the GroupHEALTH delete action.", "Delete Billing", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var confirm = MessageBox.Show(
            $"Delete billing history for {selected.StatementPeriodText} version {selected.VersionText}?\n\nThis marks the statement deleted in SQLite and leaves the rest of the database unchanged.",
            "Delete Billing",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning);
        if (confirm != MessageBoxResult.Yes)
        {
            return;
        }

        var delete = await _databaseService.DeleteGroupHealthBillingStatementAsync(selected.BillingStatementId);
        ImportStatusTextBlock.Text = delete.Message;
        await RefreshHistoryAsync(selectMostRecent: true);
    }

    private void ExportHistory_Click(object sender, RoutedEventArgs e)
    {
        if (_historyRows.Count == 0)
        {
            MessageBox.Show("There is no billing history to export yet.", "Billing Export", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var dialog = new SaveFileDialog
        {
            Title = "Export billing history",
            Filter = "Excel workbook (*.xlsx)|*.xlsx",
            FileName = $"{SafeFileName(_workspace?.ClientName ?? "Client")}_Billing_History.xlsx"
        };

        if (dialog.ShowDialog() != true)
        {
            return;
        }

        using var workbook = new XLWorkbook();
        WriteHistorySheet(workbook);
        WriteDetailSheet(workbook);
        WriteValidationSheet(workbook);
        WriteRateSheet(workbook);
        WriteChangeSheet(workbook);
        workbook.SaveAs(dialog.FileName);
        ImportStatusTextBlock.Text = $"Exported billing workbook: {dialog.FileName}";
    }

    private void WriteHistorySheet(XLWorkbook workbook)
    {
        var sheet = workbook.Worksheets.Add("Billing History");
        var headers = new[]
        {
            "Month", "Carrier", "Version", "EE Premium", "ER Premium", "Total Premium", "Employees", "Integrity Score", "Status", "Source File"
        };
        WriteHeaders(sheet, headers);
        for (var i = 0; i < _historyRows.Count; i++)
        {
            var row = _historyRows[i];
            var excelRow = i + 2;
            sheet.Cell(excelRow, 1).Value = row.StatementPeriodText;
            sheet.Cell(excelRow, 2).Value = row.Carrier;
            sheet.Cell(excelRow, 3).Value = row.VersionText;
            sheet.Cell(excelRow, 4).Value = row.TotalEmployeePremium;
            sheet.Cell(excelRow, 5).Value = row.TotalEmployerPremium;
            sheet.Cell(excelRow, 6).Value = row.TotalPremium;
            sheet.Cell(excelRow, 7).Value = row.EmployeeCount;
            sheet.Cell(excelRow, 8).Value = row.BillingIntegrityScore;
            sheet.Cell(excelRow, 9).Value = row.ImportStatus;
            sheet.Cell(excelRow, 10).Value = row.SourceFileName;
        }

        FormatWorksheet(sheet);
    }

    private void WriteDetailSheet(XLWorkbook workbook)
    {
        var sheet = workbook.Worksheets.Add("Employee Rows");
        var headers = new[]
        {
            "Status", "Class", "Certificate", "Employee", "Annual Earnings", "Cost Division", "Retro", "Life Volume", "Life Premium",
            "AD&D Premium", "LTD Premium", "Health", "Dental", "EFAP", "Virtual Healthcare", "Flex", "Member AD&D", "EE", "ER", "GST/HST", "Row Total", "Flags"
        };
        WriteHeaders(sheet, headers);
        for (var i = 0; i < _detailRows.Count; i++)
        {
            var row = _detailRows[i];
            var excelRow = i + 2;
            sheet.Cell(excelRow, 1).Value = row.ValidationStatus;
            sheet.Cell(excelRow, 2).Value = row.ClassCode;
            sheet.Cell(excelRow, 3).Value = row.CertificateId;
            sheet.Cell(excelRow, 4).Value = row.EmployeeName;
            sheet.Cell(excelRow, 5).Value = row.AnnualEarnings;
            sheet.Cell(excelRow, 6).Value = row.CostDivision;
            sheet.Cell(excelRow, 7).Value = row.RetroMonth;
            sheet.Cell(excelRow, 8).Value = row.LifeVolume;
            sheet.Cell(excelRow, 9).Value = row.LifePremium;
            sheet.Cell(excelRow, 10).Value = row.AdAndDPremium;
            sheet.Cell(excelRow, 11).Value = row.LtdPremium;
            sheet.Cell(excelRow, 12).Value = row.HealthPremium;
            sheet.Cell(excelRow, 13).Value = row.DentalPremium;
            sheet.Cell(excelRow, 14).Value = row.EfapPremium;
            sheet.Cell(excelRow, 15).Value = row.VirtualHealthcarePremium;
            sheet.Cell(excelRow, 16).Value = row.FlexAdminFee;
            sheet.Cell(excelRow, 17).Value = row.MemberAdAndDPremium;
            sheet.Cell(excelRow, 18).Value = row.EmployeePremium;
            sheet.Cell(excelRow, 19).Value = row.EmployerPremium;
            sheet.Cell(excelRow, 20).Value = row.GstHst;
            sheet.Cell(excelRow, 21).Value = row.RowTotal;
            sheet.Cell(excelRow, 22).Value = row.ValidationFlags;
        }

        FormatWorksheet(sheet);
    }

    private void WriteValidationSheet(XLWorkbook workbook)
    {
        var sheet = workbook.Worksheets.Add("Validation");
        var headers = new[] { "Status", "Level", "Check", "Expected", "Actual", "Difference", "Class", "Certificate", "Message" };
        WriteHeaders(sheet, headers);
        for (var i = 0; i < _validationRows.Count; i++)
        {
            var row = _validationRows[i];
            var excelRow = i + 2;
            sheet.Cell(excelRow, 1).Value = row.Status;
            sheet.Cell(excelRow, 2).Value = row.ValidationLevel;
            sheet.Cell(excelRow, 3).Value = row.CheckName;
            sheet.Cell(excelRow, 4).Value = row.ExpectedAmount;
            sheet.Cell(excelRow, 5).Value = row.ActualAmount;
            sheet.Cell(excelRow, 6).Value = row.DifferenceAmount;
            sheet.Cell(excelRow, 7).Value = row.ClassCode;
            sheet.Cell(excelRow, 8).Value = row.CertificateId;
            sheet.Cell(excelRow, 9).Value = row.Message;
        }

        FormatWorksheet(sheet);
    }

    private void WriteRateSheet(XLWorkbook workbook)
    {
        var sheet = workbook.Worksheets.Add("Rate Table");
        var headers = new[] { "Benefit", "Class", "Plan", "Coverage", "Basis", "EE Rate", "ER Rate", "Carrier Rate", "Records", "Status", "Source" };
        WriteHeaders(sheet, headers);
        for (var i = 0; i < _rateRows.Count; i++)
        {
            var row = _rateRows[i];
            var excelRow = i + 2;
            sheet.Cell(excelRow, 1).Value = row.Benefit;
            sheet.Cell(excelRow, 2).Value = row.ClassCode;
            sheet.Cell(excelRow, 3).Value = row.PlanCode;
            sheet.Cell(excelRow, 4).Value = row.Tier;
            sheet.Cell(excelRow, 5).Value = row.VolumeBasis;
            sheet.Cell(excelRow, 6).Value = row.EmployeeRateText;
            sheet.Cell(excelRow, 7).Value = row.EmployerRateText;
            sheet.Cell(excelRow, 8).Value = row.CarrierRateText;
            sheet.Cell(excelRow, 9).Value = row.RecordCount;
            sheet.Cell(excelRow, 10).Value = row.RateStatus;
            sheet.Cell(excelRow, 11).Value = row.Source;
        }

        FormatWorksheet(sheet);
    }

    private void WriteChangeSheet(XLWorkbook workbook)
    {
        var sheet = workbook.Worksheets.Add("Employee Changes");
        var headers = new[] { "Severity", "Change", "Employee", "Certificate", "Field", "Previous", "Current", "Difference", "Change %", "Source" };
        WriteHeaders(sheet, headers);
        for (var i = 0; i < _changeRows.Count; i++)
        {
            var row = _changeRows[i];
            var excelRow = i + 2;
            sheet.Cell(excelRow, 1).Value = row.Severity;
            sheet.Cell(excelRow, 2).Value = row.ChangeType;
            sheet.Cell(excelRow, 3).Value = row.EmployeeName;
            sheet.Cell(excelRow, 4).Value = row.CertificateId;
            sheet.Cell(excelRow, 5).Value = row.FieldName;
            sheet.Cell(excelRow, 6).Value = row.PreviousValue;
            sheet.Cell(excelRow, 7).Value = row.CurrentValue;
            sheet.Cell(excelRow, 8).Value = row.DifferenceAmount;
            sheet.Cell(excelRow, 9).Value = row.ChangePercent;
            sheet.Cell(excelRow, 10).Value = row.Source;
        }

        FormatWorksheet(sheet);
    }

    private static void WriteHeaders(IXLWorksheet sheet, IReadOnlyList<string> headers)
    {
        for (var i = 0; i < headers.Count; i++)
        {
            sheet.Cell(1, i + 1).Value = headers[i];
        }

        var range = sheet.Range(1, 1, 1, headers.Count);
        range.Style.Font.Bold = true;
        range.Style.Fill.BackgroundColor = XLColor.FromHtml("#EAF2FF");
    }

    private static void FormatWorksheet(IXLWorksheet sheet)
    {
        sheet.Columns().AdjustToContents();
        sheet.SheetView.FreezeRows(1);
        foreach (var column in sheet.ColumnsUsed().Where(column => column.CellsUsed().Skip(1).Any(cell => cell.DataType == XLDataType.Number)))
        {
            column.Style.NumberFormat.Format = "$#,##0.00;-$#,##0.00;0";
        }
    }

    private async void HistoryGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (HistoryGrid.SelectedItem is not ClientBillingHistoryRow selected)
        {
            ClearStatementDetail();
            return;
        }

        UpdateSummaryCards(selected);
        await LoadStatementDetailsAsync(selected);
    }

    private async void HistoryGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (HistoryGrid.SelectedItem is ClientBillingHistoryRow selected)
        {
            await LoadStatementDetailsAsync(selected);
        }
    }

    private void DetailFilter_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        ApplyDetailFilters();
    }

    private static string Display(string value)
    {
        return string.IsNullOrWhiteSpace(value) ? "Not set" : value;
    }

    private static string FormatMoney(decimal? value)
    {
        return value.HasValue
            ? value.Value.ToString("$#,0.00;-$#,0.00;$0.00", CultureInfo.InvariantCulture)
            : "";
    }

    private static string FormatNumber(decimal? value)
    {
        return value.HasValue
            ? value.Value.ToString("#,0.##", CultureInfo.InvariantCulture)
            : "";
    }

    private static string FormatPercent(decimal? value)
    {
        return value.HasValue
            ? value.Value.ToString("0.##", CultureInfo.InvariantCulture) + "%"
            : "";
    }

    private static string SafeFileName(string value)
    {
        var invalid = Path.GetInvalidFileNameChars();
        return string.Concat(value.Select(ch => invalid.Contains(ch) ? '_' : ch));
    }
}

public sealed class ClientBillingHistoryRow
{
    public long BillingStatementId { get; init; }
    public bool IsManulifeExcel { get; init; }
    public string Carrier { get; init; } = "";
    public int StatementMonth { get; init; }
    public int StatementYear { get; init; }
    public int ImportVersion { get; init; }
    public string ImportStatus { get; init; } = "";
    public decimal BillingIntegrityScore { get; init; }
    public decimal? TotalEmployeePremium { get; init; }
    public decimal? TotalEmployerPremium { get; init; }
    public decimal? TotalPremium { get; init; }
    public int EmployeeCount { get; init; }
    public string SourceFileName { get; init; } = "";
    public int ValidationWarningCount { get; init; }
    public int ValidationErrorCount { get; init; }

    public string StatementPeriodText => StatementMonth > 0 && StatementYear > 0
        ? $"{StatementYear}-{StatementMonth:00}"
        : "";
    public string VersionText => $"v{ImportVersion:#,0}";
    public string IntegrityScoreText => IsManulifeExcel
        ? ValidationErrorCount > 0 ? "Error" : ValidationWarningCount > 0 ? "Review" : "Passed"
        : BillingIntegrityScore > 0m ? $"{BillingIntegrityScore:0.##}%" : "N/A";
    public string TotalEmployeePremiumText => IsManulifeExcel && !TotalEmployeePremium.HasValue ? "N/A" : FormatMoney(TotalEmployeePremium);
    public string TotalEmployerPremiumText => IsManulifeExcel && !TotalEmployerPremium.HasValue ? "N/A" : FormatMoney(TotalEmployerPremium);
    public string TotalPremiumText => FormatMoney(TotalPremium);
    public string EmployeeCountText => EmployeeCount.ToString("#,0", CultureInfo.InvariantCulture);

    public static ClientBillingHistoryRow FromRecord(GroupHealthBillingHistoryRecord row)
    {
        return new ClientBillingHistoryRow
        {
            BillingStatementId = row.BillingStatementId,
            Carrier = row.Carrier,
            StatementMonth = row.StatementMonth,
            StatementYear = row.StatementYear,
            ImportVersion = row.ImportVersion,
            ImportStatus = row.ImportStatus,
            BillingIntegrityScore = row.BillingIntegrityScore,
            TotalEmployeePremium = row.TotalEmployeePremium,
            TotalEmployerPremium = row.TotalEmployerPremium,
            TotalPremium = row.TotalPremium,
            EmployeeCount = row.EmployeeCount,
            SourceFileName = row.SourceFileName
        };
    }

    public static ClientBillingHistoryRow FromManulifeRecord(CarrierBillingImportHistoryRecord row)
    {
        var month = 0;
        var year = 0;
        if (DateTime.TryParseExact(row.BillingPeriod + "-01", "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var period))
        {
            month = period.Month;
            year = period.Year;
        }

        return new ClientBillingHistoryRow
        {
            BillingStatementId = row.CarrierBillingImportId,
            IsManulifeExcel = true,
            Carrier = row.CarrierName,
            StatementMonth = month,
            StatementYear = year,
            ImportVersion = row.VersionNumber,
            ImportStatus = row.ImportStatus,
            TotalPremium = row.TotalPremiumAndTax,
            EmployeeCount = row.EmployeeRecordCount,
            SourceFileName = row.SourceFileName,
            ValidationWarningCount = row.WarningCount,
            ValidationErrorCount = row.ErrorCount
        };
    }

    private static string FormatMoney(decimal? value)
    {
        return value.HasValue
            ? value.Value.ToString("$#,0.00;-$#,0.00;$0.00", CultureInfo.InvariantCulture)
            : "";
    }
}

public sealed class ClientBillingDetailRow
{
    public string ValidationStatus { get; init; } = "";
    public string ClassCode { get; init; } = "";
    public string CertificateId { get; init; } = "";
    public string EmployeeName { get; init; } = "";
    public decimal? AnnualEarnings { get; init; }
    public string PlanCode { get; init; } = "";
    public decimal? CurrentMonthPremium { get; init; }
    public decimal? PriorPeriodAdjustment { get; init; }
    public string CostDivision { get; init; } = "";
    public string RetroMonth { get; init; } = "";
    public decimal? LifeVolume { get; init; }
    public decimal? LifePremium { get; init; }
    public decimal? AdAndDPremium { get; init; }
    public decimal? LtdPremium { get; init; }
    public decimal? HealthPremium { get; init; }
    public decimal? DentalPremium { get; init; }
    public decimal? EfapPremium { get; init; }
    public decimal? VirtualHealthcarePremium { get; init; }
    public decimal? FlexAdminFee { get; init; }
    public decimal? MemberAdAndDPremium { get; init; }
    public decimal? EmployeePremium { get; init; }
    public decimal? EmployerPremium { get; init; }
    public decimal? GstHst { get; init; }
    public decimal? RowTotal { get; init; }
    public bool FundingSplitUnavailable { get; init; }
    public string ValidationFlags { get; init; } = "";

    public string AnnualEarningsText => FormatNumber(AnnualEarnings);
    public string CurrentMonthPremiumText => FormatMoney(CurrentMonthPremium);
    public string PriorPeriodAdjustmentText => FormatMoney(PriorPeriodAdjustment);
    public string LifeVolumeText => FormatNumber(LifeVolume);
    public string LifePremiumText => FormatMoney(LifePremium);
    public string AdAndDPremiumText => FormatMoney(AdAndDPremium);
    public string LtdPremiumText => FormatMoney(LtdPremium);
    public string HealthPremiumText => FormatMoney(HealthPremium);
    public string DentalPremiumText => FormatMoney(DentalPremium);
    public string EfapPremiumText => FormatMoney(EfapPremium);
    public string VirtualHealthcarePremiumText => FormatMoney(VirtualHealthcarePremium);
    public string FlexAdminFeeText => FormatMoney(FlexAdminFee);
    public string MemberAdAndDPremiumText => FormatMoney(MemberAdAndDPremium);
    public string EmployeePremiumText => FundingSplitUnavailable ? "N/A" : FormatMoney(EmployeePremium);
    public string EmployerPremiumText => FundingSplitUnavailable ? "N/A" : FormatMoney(EmployerPremium);
    public string GstHstText => FormatMoney(GstHst);
    public string RowTotalText => FormatMoney(RowTotal);

    public static ClientBillingDetailRow FromRecord(GroupHealthBillingRow row)
    {
        return new ClientBillingDetailRow
        {
            ValidationStatus = row.ValidationStatus,
            ClassCode = row.ClassCode,
            CertificateId = row.CertificateId,
            EmployeeName = row.EmployeeName,
            AnnualEarnings = row.AnnualEarnings,
            CostDivision = row.CostDivision,
            RetroMonth = row.RetroMonth,
            LifeVolume = row.LifeVolume,
            LifePremium = row.LifePremium,
            AdAndDPremium = row.AdAndDPremium,
            LtdPremium = row.LtdPremium,
            HealthPremium = row.HealthPremium,
            DentalPremium = row.DentalPremium,
            EfapPremium = row.EfapPremium,
            VirtualHealthcarePremium = row.VirtualHealthcarePremium,
            FlexAdminFee = row.FlexAdminFee,
            MemberAdAndDPremium = row.MemberAdAndDPremium + row.MemberAdAndDSpPremium + row.MemberAdAndDChPremium,
            CurrentMonthPremium = row.EmployeePremium.HasValue || row.EmployerPremium.HasValue
                ? (row.EmployeePremium ?? 0m) + (row.EmployerPremium ?? 0m)
                : null,
            EmployeePremium = row.EmployeePremium,
            EmployerPremium = row.EmployerPremium,
            GstHst = row.GstHst,
            RowTotal = row.RowTotal,
            ValidationFlags = row.ValidationFlags
        };
    }

    public static ClientBillingDetailRow FromManulifeRecord(CarrierBillingEmployeeRecord row)
    {
        var flags = row.ReconciliationDifference is { } difference && difference != 0m
            ? $"Reconciliation difference: {difference:$#,0.00;-$#,0.00;$0.00}"
            : "";
        return new ClientBillingDetailRow
        {
            ValidationStatus = row.ValidationStatus,
            ClassCode = row.ClassCode,
            PlanCode = row.PlanCode,
            CertificateId = row.CertificateNumber,
            EmployeeName = string.Join(" ", new[] { row.FirstName, row.LastName }.Where(value => !string.IsNullOrWhiteSpace(value))),
            CurrentMonthPremium = row.CurrentMonthPremium,
            PriorPeriodAdjustment = row.PriorPeriodAdjustment,
            CostDivision = row.DivisionalBillingCode,
            GstHst = row.TotalSalesTax,
            RowTotal = row.TotalPremiumAndTax,
            FundingSplitUnavailable = true,
            ValidationFlags = flags
        };
    }

    private static string FormatMoney(decimal? value)
    {
        return value.HasValue
            ? value.Value.ToString("$#,0.00;-$#,0.00;$0.00", CultureInfo.InvariantCulture)
            : "";
    }

    private static string FormatNumber(decimal? value)
    {
        return value.HasValue
            ? value.Value.ToString("#,0.##", CultureInfo.InvariantCulture)
            : "";
    }
}

public sealed class ClientBillingValidationRow
{
    public string Status { get; init; } = "";
    public string ValidationLevel { get; init; } = "";
    public string CheckName { get; init; } = "";
    public decimal? ExpectedAmount { get; init; }
    public decimal? ActualAmount { get; init; }
    public decimal? DifferenceAmount { get; init; }
    public string ClassCode { get; init; } = "";
    public string CertificateId { get; init; } = "";
    public string Message { get; init; } = "";

    public string ExpectedAmountText => FormatMoney(ExpectedAmount);
    public string ActualAmountText => FormatMoney(ActualAmount);
    public string DifferenceAmountText => FormatMoney(DifferenceAmount);

    public static ClientBillingValidationRow FromRecord(GroupHealthBillingValidationResult row)
    {
        return new ClientBillingValidationRow
        {
            Status = row.Status,
            ValidationLevel = row.ValidationLevel,
            CheckName = row.CheckName,
            ExpectedAmount = row.ExpectedAmount,
            ActualAmount = row.ActualAmount,
            DifferenceAmount = row.DifferenceAmount,
            ClassCode = row.ClassCode,
            CertificateId = row.CertificateId,
            Message = row.Message
        };
    }

    public static ClientBillingValidationRow FromManulifeRecord(CarrierBillingException row)
    {
        return new ClientBillingValidationRow
        {
            Status = row.Severity.Equals(CarrierBillingExceptionSeverities.Error, StringComparison.OrdinalIgnoreCase) ? "Failed" : "Warning",
            ValidationLevel = string.IsNullOrWhiteSpace(row.BenefitCode) ? "File" : row.BenefitCode,
            CheckName = row.ExceptionType,
            CertificateId = row.CertificateNumber,
            Message = row.Message
        };
    }

    private static string FormatMoney(decimal? value)
    {
        return value.HasValue
            ? value.Value.ToString("#,0.00;-#,0.00;0.00", CultureInfo.InvariantCulture)
            : "";
    }
}

public sealed class ClientBillingRateRow
{
    public string Benefit { get; init; } = "";
    public string ClassCode { get; init; } = "";
    public string PlanCode { get; init; } = "";
    public string Tier { get; init; } = "";
    public string VolumeBasis { get; init; } = "";
    public decimal? EmployeeRate { get; init; }
    public decimal? EmployerRate { get; init; }
    public decimal? CarrierRate { get; init; }
    public int RecordCount { get; init; } = 1;
    public bool FundingSplitUnavailable { get; init; }
    public string RateStatus { get; init; } = "";
    public string Source { get; init; } = "";

    public string EmployeeRateText => FundingSplitUnavailable ? "N/A" : FormatRate(EmployeeRate);
    public string EmployerRateText => FundingSplitUnavailable ? "N/A" : FormatRate(EmployerRate);
    public string CarrierRateText => FormatRate(CarrierRate);

    public static ClientBillingRateRow FromRecord(GroupHealthClientRateTableLine row)
    {
        return new ClientBillingRateRow
        {
            Benefit = row.Benefit,
            ClassCode = row.ClassCode,
            Tier = row.Tier,
            VolumeBasis = row.VolumeBasis,
            EmployeeRate = row.EmployeeRate,
            EmployerRate = row.EmployerRate,
            CarrierRate = row.TotalRate ?? row.CalculatedRate,
            RateStatus = row.RateStatus,
            Source = row.Source
        };
    }

    public static ClientBillingRateRow FromManulifeSummary(CarrierBillingRateSummary row)
    {
        return new ClientBillingRateRow
        {
            Benefit = row.BenefitCode,
            ClassCode = row.ClassCode,
            PlanCode = row.PlanCode,
            Tier = row.Coverage,
            VolumeBasis = row.RateBasis,
            CarrierRate = row.Rate,
            RecordCount = row.RecordCount,
            FundingSplitUnavailable = true,
            RateStatus = "Calculated",
            Source = row.Source
        };
    }

    public ClientBillingRateRow WithRecordCount(int recordCount)
    {
        return new ClientBillingRateRow
        {
            Benefit = Benefit,
            ClassCode = ClassCode,
            PlanCode = PlanCode,
            Tier = Tier,
            VolumeBasis = VolumeBasis,
            EmployeeRate = EmployeeRate,
            EmployerRate = EmployerRate,
            CarrierRate = CarrierRate,
            RecordCount = recordCount,
            FundingSplitUnavailable = FundingSplitUnavailable,
            RateStatus = RateStatus,
            Source = Source
        };
    }

    private static string FormatRate(decimal? value)
    {
        return value.HasValue
            ? value.Value.ToString("#,0.####", CultureInfo.InvariantCulture)
            : "";
    }
}

public sealed class ClientBillingChangeRow
{
    public string Severity { get; init; } = "";
    public string ChangeType { get; init; } = "";
    public string EmployeeName { get; init; } = "";
    public string CertificateId { get; init; } = "";
    public string FieldName { get; init; } = "";
    public string PreviousValue { get; init; } = "";
    public string CurrentValue { get; init; } = "";
    public decimal? DifferenceAmount { get; init; }
    public decimal? ChangePercent { get; init; }
    public string Source { get; init; } = "";

    public string DifferenceAmountText => FormatNumber(DifferenceAmount);
    public string ChangePercentText => FormatPercent(ChangePercent);

    public static ClientBillingChangeRow FromRecord(GroupHealthBillingEmployeeChange row)
    {
        return new ClientBillingChangeRow
        {
            Severity = row.Severity,
            ChangeType = row.ChangeType,
            EmployeeName = row.EmployeeName,
            CertificateId = row.CertificateId,
            FieldName = row.FieldName,
            PreviousValue = row.PreviousValue,
            CurrentValue = row.CurrentValue,
            DifferenceAmount = row.DifferenceAmount,
            ChangePercent = row.ChangePercent,
            Source = row.Source
        };
    }

    private static string FormatNumber(decimal? value)
    {
        return value.HasValue
            ? value.Value.ToString("#,0.##;-#,0.##;0", CultureInfo.InvariantCulture)
            : "";
    }

    private static string FormatPercent(decimal? value)
    {
        return value.HasValue
            ? value.Value.ToString("#,0.##", CultureInfo.InvariantCulture) + "%"
            : "";
    }
}

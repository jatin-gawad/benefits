using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;

namespace RenewalReviewBuilder.App.Views;

public partial class ControlledBookletReprocessingView : UserControl
{
    private readonly CanonicalControlledBookletReprocessingService _service = new();
    private readonly ObservableCollection<ControlledBookletDocumentOption> _documents = [];
    private ControlledBookletReprocessingRun? _currentRun;
    private CancellationTokenSource? _stagingCancellation;
    private bool _loaded;
    private bool _isBusy;
    private bool _isLoadingDocuments;

    public ControlledBookletReprocessingView()
    {
        InitializeComponent();
        DocumentComboBox.ItemsSource = _documents;
    }

    private async void ControlledBookletReprocessingView_Loaded(object sender, RoutedEventArgs e)
    {
        if (_loaded)
        {
            return;
        }

        _loaded = true;
        await LoadDocumentsAsync();
    }

    private async void RefreshDocuments_Click(object sender, RoutedEventArgs e)
    {
        await LoadDocumentsAsync();
    }

    private async void DocumentComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_isLoadingDocuments)
        {
            ApplySelectedDocument(DocumentComboBox.SelectedItem as ControlledBookletDocumentOption);
            return;
        }

        await LoadSelectedDocumentRunAsync(
            DocumentComboBox.SelectedItem as ControlledBookletDocumentOption);
    }

    private async void StageBooklet_Click(object sender, RoutedEventArgs e)
    {
        if (DocumentComboBox.SelectedItem is not ControlledBookletDocumentOption selected)
        {
            MessageBox.Show(
                "Select one registered benefit booklet first.",
                "Controlled Booklet Reprocessing",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        try
        {
            SetBusy(true, "Preparing the selected booklet for review. Live Plan Design data is unchanged.");
            _stagingCancellation = new CancellationTokenSource();
            CancelStagingButton.IsEnabled = true;
            var progress = new Progress<BookletExtractionProgress>(status =>
            {
                WorkflowStatusTextBlock.Text = status.Message;
            });
            _currentRun = await _service.StageAsync(
                new ControlledBookletReprocessingRequest
                {
                    SourceDocumentId = selected.DocumentId,
                    RequestedBy = Environment.UserName
                },
                progress,
                _stagingCancellation.Token);
            ApplyRun(_currentRun);
            WorkflowStatusTextBlock.Text = _currentRun.Status == ControlledBookletReprocessingStatuses.ReadyForApproval
                ? "The review is ready. Check the changes before approval."
                : _currentRun.Status == ControlledBookletReprocessingStatuses.Blocked
                    ? "The review needs corrections. Select a plan row to edit or review the items below."
                    : $"Review preparation finished with status: {_currentRun.Status}.";
        }
        catch (OperationCanceledException)
        {
            WorkflowStatusTextBlock.Text = "Review preparation was cancelled. No live Plan Design data was changed.";
        }
        catch (InvalidOperationException ex)
        {
            AppLogger.Error(ex);
            var activeRun = await _service.GetActiveRunForDocumentAsync(selected.DocumentId);
            if (activeRun is not null)
            {
                _currentRun = activeRun;
                ApplyRun(_currentRun);
                WorkflowStatusTextBlock.Text =
                    $"Loaded active run {_currentRun.RunId}. Review it before starting another reprocess.";
            }
            else
            {
                WorkflowStatusTextBlock.Text = "The booklet review could not be prepared.";
                MessageBox.Show(
                    ex.Message,
                    "Controlled booklet staging failed",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            WorkflowStatusTextBlock.Text = "The booklet review could not be prepared.";
            MessageBox.Show(
                ex.Message,
                "Controlled booklet staging failed",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        finally
        {
            _stagingCancellation?.Dispose();
            _stagingCancellation = null;
            SetBusy(false);
        }
    }

    private void CancelStaging_Click(object sender, RoutedEventArgs e)
    {
        _stagingCancellation?.Cancel();
        WorkflowStatusTextBlock.Text = "Cancelling review preparation...";
    }

    private async void EditSelected_Click(object sender, RoutedEventArgs e)
    {
        await EditSelectedRowAsync();
    }

    private void ChangesGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        UpdateButtons();
    }

    private async void ChangesGrid_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        if (EditSelectedButton.IsEnabled)
        {
            await EditSelectedRowAsync();
        }
    }

    private async Task EditSelectedRowAsync()
    {
        if (_currentRun is null ||
            ChangesGrid.SelectedItem is not ControlledBookletReprocessingChange selectedChange)
        {
            return;
        }

        var stagedRow = _currentRun.StagedRows.FirstOrDefault(row =>
            PlanDesignCanonicalMapper.BuildProvisionComparisonKey(row)
                .Equals(selectedChange.ComparisonKey, StringComparison.OrdinalIgnoreCase));
        if (stagedRow is null)
        {
            MessageBox.Show(
                "This line cannot be edited because it is not present in the review copy.",
                "Edit Plan Provision",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        var dialog = new ControlledBookletPlanEditDialog(stagedRow, selectedChange.ComparisonKey)
        {
            Owner = Window.GetWindow(this)
        };
        if (dialog.ShowDialog() != true || dialog.Request is null)
        {
            return;
        }

        try
        {
            SetBusy(true, "Saving the correction and rerunning the plan checks...");
            _currentRun = await _service.ApplyStagedEditAsync(_currentRun.RunId, dialog.Request);
            ApplyRun(_currentRun);
            var refreshedChange = _currentRun.Changes.FirstOrDefault(change =>
                change.ComparisonKey.Equals(selectedChange.ComparisonKey, StringComparison.OrdinalIgnoreCase));
            if (refreshedChange is not null)
            {
                ShowUnchangedRowsCheckBox.IsChecked = refreshedChange.ChangeType.Equals(
                    ControlledBookletReprocessingChangeTypes.Unchanged,
                    StringComparison.OrdinalIgnoreCase);
                ApplyChangeFilter();
                ChangesGrid.SelectedItem = refreshedChange;
                ChangesGrid.ScrollIntoView(refreshedChange);
            }

            WorkflowStatusTextBlock.Text = _currentRun.CanApprove
                ? "Correction saved and all checks passed. The review is ready for approval."
                : "Correction saved. Review the remaining items to resolve.";
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            MessageBox.Show(
                ex.Message,
                "Plan correction could not be saved",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        finally
        {
            SetBusy(false);
        }
    }

    private async void ApproveRun_Click(object sender, RoutedEventArgs e)
    {
        if (_currentRun is null || !_currentRun.CanApprove)
        {
            return;
        }

        var confirmation = MessageBox.Show(
            $"Apply staged run {_currentRun.RunId} to {_currentRun.ClientName}?\n\n" +
            $"Current rows: {_currentRun.BeforeRowCount}\n" +
            $"Staged rows: {_currentRun.StagedRowCount}\n" +
            $"{_currentRun.DifferenceSummary}\n\n" +
            "A verified database backup will be created before the document-scoped replacement.",
            "Approve controlled booklet replacement",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning,
            MessageBoxResult.No);
        if (confirmation != MessageBoxResult.Yes)
        {
            return;
        }

        try
        {
            SetBusy(true, "Creating a database backup and applying the approved Plan Design replacement...");
            _currentRun = await _service.ApproveAndCommitAsync(
                _currentRun.RunId,
                Environment.UserName);
            ApplyRun(_currentRun);
            WorkflowStatusTextBlock.Text = _currentRun.Status == ControlledBookletReprocessingStatuses.Committed
                ? "Approved replacement committed and audited successfully."
                : $"The replacement did not complete. Status: {_currentRun.Status}.";
            await LoadDocumentsAsync(_currentRun.SourceDocumentId, preserveRun: true);
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            WorkflowStatusTextBlock.Text = "Approval could not be completed.";
            MessageBox.Show(
                ex.Message,
                "Controlled booklet approval failed",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        finally
        {
            SetBusy(false);
        }
    }

    private async void RejectRun_Click(object sender, RoutedEventArgs e)
    {
        if (_currentRun is null || !_currentRun.CanReject)
        {
            return;
        }

        var confirmation = MessageBox.Show(
            $"Reject staged run {_currentRun.RunId}? No live Plan Design data will be changed.",
            "Reject controlled booklet run",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question,
            MessageBoxResult.No);
        if (confirmation != MessageBoxResult.Yes)
        {
            return;
        }

        try
        {
            SetBusy(true, "Discarding the review...");
            _currentRun = await _service.RejectAsync(
                _currentRun.RunId,
                Environment.UserName,
                ReviewNoteTextBox.Text.Trim());
            ApplyRun(_currentRun);
            WorkflowStatusTextBlock.Text = "The review was discarded. Live Plan Design data was unchanged.";
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            MessageBox.Show(
                ex.Message,
                "Controlled booklet rejection failed",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        finally
        {
            SetBusy(false);
        }
    }

    private void OpenStagingFolder_Click(object sender, RoutedEventArgs e)
    {
        OpenFolder(_currentRun?.OutputFolder);
    }

    private void OpenBackup_Click(object sender, RoutedEventArgs e)
    {
        var backupPath = _currentRun?.BackupPath ?? "";
        OpenFolder(string.IsNullOrWhiteSpace(backupPath) ? "" : Path.GetDirectoryName(backupPath));
    }

    private void ShowUnchangedRows_Click(object sender, RoutedEventArgs e)
    {
        ApplyChangeFilter();
    }

    private async Task LoadDocumentsAsync(long? selectedDocumentId = null, bool preserveRun = false)
    {
        if (_isBusy && !preserveRun)
        {
            return;
        }

        try
        {
            var preservedRun = preserveRun ? _currentRun : null;
            var currentDocumentId = selectedDocumentId ??
                                    (DocumentComboBox.SelectedItem as ControlledBookletDocumentOption)?.DocumentId;
            if (!preserveRun)
            {
                SetBusy(true, "Loading registered benefit booklets...");
            }

            var documents = await _service.GetDocumentsAsync();
            _isLoadingDocuments = true;
            try
            {
                _documents.Clear();
                foreach (var document in documents)
                {
                    _documents.Add(document);
                }

                DocumentComboBox.SelectedItem = currentDocumentId.HasValue
                    ? _documents.FirstOrDefault(document => document.DocumentId == currentDocumentId.Value)
                    : _documents.FirstOrDefault(document => document.OutOfRangePercentageCount > 0) ?? _documents.FirstOrDefault();
                _currentRun = preservedRun;
                ApplySelectedDocument(DocumentComboBox.SelectedItem as ControlledBookletDocumentOption);
            }
            finally
            {
                _isLoadingDocuments = false;
            }

            if (preserveRun)
            {
                _currentRun = preservedRun;
                ApplyRun(_currentRun);
            }
            else
            {
                await LoadSelectedDocumentRunAsync(
                    DocumentComboBox.SelectedItem as ControlledBookletDocumentOption,
                    manageBusyState: false);
            }
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            WorkflowStatusTextBlock.Text = $"Registered booklets could not load: {ex.Message}";
        }
        finally
        {
            if (!preserveRun)
            {
                SetBusy(false);
            }
            UpdateButtons();
        }
    }

    private async Task LoadSelectedDocumentRunAsync(
        ControlledBookletDocumentOption? document,
        bool manageBusyState = true)
    {
        if (manageBusyState)
        {
            SetBusy(true, "Checking for an existing controlled reprocessing run...");
        }

        try
        {
            _currentRun = null;
            ApplySelectedDocument(document);
            ApplyRun(null);

            if (document is null)
            {
                WorkflowStatusTextBlock.Text = "No registered benefit booklets were found.";
                return;
            }

            var activeRun = await _service.GetActiveRunForDocumentAsync(document.DocumentId);
            if ((DocumentComboBox.SelectedItem as ControlledBookletDocumentOption)?.DocumentId != document.DocumentId)
            {
                return;
            }

            _currentRun = activeRun;
            ApplySelectedDocument(document);
            ApplyRun(_currentRun);
            WorkflowStatusTextBlock.Text = _currentRun is null
                ? "No active run exists for this booklet. It is ready for controlled staging."
                : $"Loaded active run {_currentRun.RunId}. Review it before starting another reprocess.";
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            _currentRun = null;
            ApplyRun(null);
            WorkflowStatusTextBlock.Text = $"The booklet's active run could not load: {ex.Message}";
        }
        finally
        {
            if (manageBusyState)
            {
                SetBusy(false);
            }
        }
    }

    private void ApplySelectedDocument(ControlledBookletDocumentOption? document)
    {
        SelectedClientTextBlock.Text = document?.ClientName ?? "-";
        SelectedDocumentIdTextBlock.Text = document is null ? "-" : document.DocumentId.ToString();
        SelectedHashTextBlock.Text = document?.HashPreview ?? "-";
        SelectedFileTextBlock.Text = document?.FileName ?? "-";
        SelectedFileTextBlock.ToolTip = document?.FilePath ?? "";
        BeforeRowsTextBlock.Text = document?.CurrentRowCount.ToString() ?? "0";
        StagedRowsTextBlock.Text = "0";
    }

    private void ApplyRun(ControlledBookletReprocessingRun? run)
    {
        var unchangedCount = run?.Changes.Count(change =>
            change.ChangeType.Equals(
                ControlledBookletReprocessingChangeTypes.Unchanged,
                StringComparison.OrdinalIgnoreCase)) ?? 0;
        ShowUnchangedRowsCheckBox.IsChecked = false;
        ShowUnchangedRowsCheckBox.IsEnabled = unchangedCount > 0;
        ShowUnchangedRowsCheckBox.Content = unchangedCount > 0
            ? $"Show unchanged ({unchangedCount})"
            : "Show unchanged";
        ApplyChangeFilter();
        IssuesGrid.ItemsSource = run?.Issues;
        RunStatusTextBlock.Text = run?.Status ?? "Not staged";
        RunIdentityTextBlock.Text = run is null
            ? ""
            : $"Run {run.RunId} | Document {run.SourceDocumentId} | {run.WorkflowVersion}";
        BeforeRowsTextBlock.Text = run?.BeforeRowCount.ToString() ??
                                   (DocumentComboBox.SelectedItem as ControlledBookletDocumentOption)?.CurrentRowCount.ToString() ?? "0";
        StagedRowsTextBlock.Text = run?.StagedRowCount.ToString() ?? "0";
        ChangesSummaryTextBlock.Text = run?.DifferenceSummary ?? "-";
        if (run is not null && run.Edits.Count > 0)
        {
            ChangesSummaryTextBlock.Text += $" | Corrections {run.Edits.Count}";
        }
        ValidationSummaryTextBlock.Text = run?.ValidationSummary ?? "-";
        ReviewNoteTextBox.Text = run?.Notes ?? "";
        OpenStagingFolderButton.IsEnabled = run is not null && Directory.Exists(run.OutputFolder);
        OpenBackupButton.IsEnabled = run is not null && File.Exists(run.BackupPath);
        UpdateButtons();
    }

    private void ApplyChangeFilter()
    {
        if (_currentRun is null)
        {
            ChangesGrid.ItemsSource = null;
            return;
        }

        ChangesGrid.ItemsSource = ShowUnchangedRowsCheckBox.IsChecked == true
            ? _currentRun.Changes
            : _currentRun.Changes
                .Where(change => !change.ChangeType.Equals(
                    ControlledBookletReprocessingChangeTypes.Unchanged,
                    StringComparison.OrdinalIgnoreCase))
                .ToList();
    }

    private void SetBusy(bool isBusy, string? status = null)
    {
        _isBusy = isBusy;
        if (!string.IsNullOrWhiteSpace(status))
        {
            WorkflowStatusTextBlock.Text = status;
        }

        UpdateButtons();
    }

    private void UpdateButtons()
    {
        var hasActiveRun = _currentRun?.Status is
            ControlledBookletReprocessingStatuses.Extracting or
            ControlledBookletReprocessingStatuses.Blocked or
            ControlledBookletReprocessingStatuses.ReadyForApproval or
            ControlledBookletReprocessingStatuses.Committing;
        DocumentComboBox.IsEnabled = !_isBusy;
        RefreshDocumentsButton.IsEnabled = !_isBusy;
        StageBookletButton.IsEnabled = !_isBusy &&
                                               !hasActiveRun &&
                                               DocumentComboBox.SelectedItem is ControlledBookletDocumentOption;
        CancelStagingButton.IsEnabled = _isBusy && _stagingCancellation is not null;
        ApproveRunButton.IsEnabled = !_isBusy && _currentRun?.CanApprove == true;
        RejectRunButton.IsEnabled = !_isBusy && _currentRun?.CanReject == true;
        EditSelectedButton.IsEnabled = !_isBusy &&
                                       _currentRun?.CanEdit == true &&
                                       ChangesGrid.SelectedItem is ControlledBookletReprocessingChange change &&
                                       !change.ChangeType.Equals(
                                           ControlledBookletReprocessingChangeTypes.Removed,
                                           StringComparison.OrdinalIgnoreCase) &&
                                       _currentRun.StagedRows.Any(row =>
                                           PlanDesignCanonicalMapper.BuildProvisionComparisonKey(row)
                                               .Equals(change.ComparisonKey, StringComparison.OrdinalIgnoreCase));
        OpenStagingFolderButton.IsEnabled = !_isBusy && _currentRun is not null && Directory.Exists(_currentRun.OutputFolder);
        OpenBackupButton.IsEnabled = !_isBusy && _currentRun is not null && File.Exists(_currentRun.BackupPath);
    }

    private static void OpenFolder(string? folder)
    {
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
}

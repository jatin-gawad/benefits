using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using Microsoft.VisualBasic;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;

namespace RenewalReviewBuilder.App.Views;

public partial class ClientFinancialIntelligenceDashboardView : UserControl
{
    private static readonly CultureInfo CanadianCulture = CultureInfo.GetCultureInfo("en-CA");
    private readonly DatabaseService _databaseService = new();
    private readonly FinancialArrangementService _financialArrangementService = new(AppPaths.BenefitsIntelligenceDatabasePath);
    private readonly RenewalFinancialArrangementComparisonDataService _financialArrangementComparisonDataService = new(AppPaths.BenefitsIntelligenceDatabasePath);
    private readonly AsoFinancialActivityForecastService _forecastService = new();
    private readonly FinancialProjectionCalculationService _financialProjectionCalculationService = new();
    private readonly AsoFinancialActivityImportService _importService;
    private ClientWorkspaceViewModel? _workspace;
    private bool _isBusy;
    private bool _loadingStatementHistory;
    private IReadOnlyList<AsoFinancialActivityStatement> _statementHistory = [];
    private AsoExperiencePeriodReference? _renewalExperiencePeriod;

    public ClientFinancialIntelligenceDashboardView()
    {
        _importService = new AsoFinancialActivityImportService(_databaseService);
        InitializeComponent();
    }

    public async Task LoadDashboardAsync(ClientWorkspaceViewModel workspace)
    {
        _workspace = workspace;
        HeaderSubtitleTextBlock.Text =
            $"{workspace.ClientName} | Carrier: {Display(workspace.PrimaryCarrier)} | SQLite financial history";

        if (workspace.DatabaseClientId is not > 0)
        {
            ShowEmpty(
                "This client is not linked to SQLite yet.",
                "Import an ASO Activity workbook to link and save financial activity for this client.");
            return;
        }

        try
        {
            SetBusy(true, "Loading saved ASO financial activity...");
            _statementHistory = await _databaseService.GetAsoFinancialActivityHistoryAsync(workspace.DatabaseClientId.Value);
            _renewalExperiencePeriod = await _databaseService.GetRenewalExperiencePeriodAsync(workspace.DatabaseClientId.Value);
            if (_renewalExperiencePeriod is null && await _databaseService.BackfillRenewalExperiencePeriodAsync(workspace.DatabaseClientId.Value))
                _renewalExperiencePeriod = await _databaseService.GetRenewalExperiencePeriodAsync(workspace.DatabaseClientId.Value);
            await LoadFinancialArrangementComparisonAsync(workspace);
            var statement = _statementHistory.FirstOrDefault();
            if (statement is null)
            {
                ShowEmpty(
                    "No saved ASO financial activity for this client yet.",
                    "Import an ASO Activity workbook, or process it through Document Operations. Approved results will appear here automatically.");
                return;
            }

            LoadStatementHistory(statement);
        }
        catch (Exception exception)
        {
            AppLogger.Error(exception);
            ShowEmpty(
                "Financial activity could not be loaded.",
                "The client remains selected. Refresh the page or import the source workbook again.");
        }
        finally
        {
            SetBusy(false);
        }
    }

    public void LoadEmpty(ClientWorkspaceViewModel workspace)
    {
        _workspace = workspace;
        HeaderSubtitleTextBlock.Text =
            $"{workspace.ClientName} | Carrier: {Display(workspace.PrimaryCarrier)} | SQLite financial history";
        ShowEmpty(
            "No saved ASO financial activity for this client yet.",
            "Import an ASO Activity workbook, or process it through Document Operations. Approved results will appear here automatically.");
        _ = LoadFinancialArrangementComparisonAsync(workspace);
    }

    private async void ImportAsoActivity_Click(object sender, RoutedEventArgs e)
    {
        if (_isBusy)
        {
            return;
        }

        if (_workspace is null)
        {
            MessageBox.Show(
                "Select a client before importing ASO financial activity.",
                "Financial Intelligence",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        var dialog = new OpenFileDialog
        {
            Title = "Select ASO Activity workbook",
            Filter = "ASO Activity workbooks (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|Excel workbooks (*.xlsx)|*.xlsx|All files (*.*)|*.*",
            Multiselect = false
        };
        if (dialog.ShowDialog() != true)
        {
            return;
        }

        try
        {
            SetBusy(true, $"Reading {Path.GetFileName(dialog.FileName)}...");
            var request = new AsoFinancialActivityImportRequest
            {
                InputPath = dialog.FileName,
                SelectedClientId = _workspace.DatabaseClientId,
                SelectedCarrier = _workspace.PrimaryCarrier,
                ImportedBy = "Client Financial Intelligence"
            };
            var import = await _importService.ImportAsync(request);
            var save = import.Persistence;
            if (save is null)
            {
                MessageBox.Show(
                    import.Message,
                    "Financial Intelligence Import",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                ImportStatusTextBlock.Text = import.Message;
                return;
            }

            if (!save.Succeeded &&
                save.RequiresVerifiedAsoPolicyLink &&
                import.Statement is not null &&
                _workspace.DatabaseClientId is > 0)
            {
                var confirmation = MessageBox.Show(
                    $"This ASO workbook is for policy {save.WorkbookContractNumber}. The selected Client Workspace uses policy {save.SelectedClientNumber}.\n\n" +
                    $"Link {save.WorkbookContractNumber} as a verified ASO policy for {_workspace.ClientName}, then import this activity into the same client history?",
                    "Confirm related ASO policy",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);

                if (confirmation == MessageBoxResult.Yes)
                {
                    var policyLink = await _databaseService.LinkVerifiedAsoPolicyAsync(
                        _workspace.DatabaseClientId.Value,
                        import.Statement.ContractNumber,
                        import.Statement.ClientName,
                        import.Statement.Carrier,
                        Path.GetFileName(dialog.FileName),
                        "Client Financial Intelligence");
                    if (policyLink.Succeeded)
                    {
                        import = await _importService.ImportAsync(request);
                        save = import.Persistence;
                    }
                    else
                    {
                        MessageBox.Show(
                            policyLink.Message,
                            "Financial Intelligence Import",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning);
                        ImportStatusTextBlock.Text = policyLink.Message;
                        return;
                    }
                }
            }

            if (save is null || !save.Succeeded)
            {
                MessageBox.Show(
                    save?.Message ?? import.Message,
                    "Financial Intelligence Import",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                ImportStatusTextBlock.Text = save?.Message ?? import.Message;
                return;
            }

            if (_workspace.DatabaseClientId is not > 0 && save.ClientId is > 0)
            {
                _workspace.DatabaseClientId = save.ClientId;
            }

            await LoadDashboardAsync(_workspace);
            MessageBox.Show(
                save.WasDuplicate
                    ? save.Message
                    : $"{save.Message}{Environment.NewLine}{Environment.NewLine}The saved activity is now available in this client's Financial Intelligence workspace.",
                "Financial Intelligence Import",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
        catch (Exception exception)
        {
            AppLogger.Error(exception);
            MessageBox.Show(
                exception.Message,
                "Financial Intelligence Import Failed",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        finally
        {
            SetBusy(false);
        }
    }

    private async void Refresh_Click(object sender, RoutedEventArgs e)
    {
        if (_workspace is not null)
        {
            await LoadDashboardAsync(_workspace);
        }
    }

    private async void RefreshFinancialArrangementComparison_Click(object sender, RoutedEventArgs e)
    {
        if (_workspace is not null)
        {
            await RefreshFinancialArrangementSnapshotAsync(_workspace);
            await LoadFinancialArrangementComparisonAsync(_workspace);
        }
    }

    private async void ImportFinancialTerms_Click(object sender, RoutedEventArgs e)
    {
        if (_isBusy || _workspace?.DatabaseClientId is not > 0) return;

        var scenarioChoice = MessageBox.Show(
            "Is this PDF the current active financial arrangement?\n\nChoose Yes for Current terms, No for a Renewal proposal, or Cancel.",
            "Financial Terms Document", MessageBoxButton.YesNoCancel, MessageBoxImage.Question);
        if (scenarioChoice == MessageBoxResult.Cancel) return;

        var dialog = new OpenFileDialog { Title = "Select renewal or financial arrangement PDF", Filter = "PDF files (*.pdf)|*.pdf", Multiselect = false };
        if (dialog.ShowDialog() != true) return;

        var scenario = scenarioChoice == MessageBoxResult.Yes ? FinancialArrangementScenario.Current : FinancialArrangementScenario.InitialRenewal;
        var effectiveDate = DateOnly.TryParse(_workspace.RenewalDate, out var renewalDate) ? renewalDate : DateOnly.FromDateTime(DateTime.Today);
        if (scenario == FinancialArrangementScenario.Current) effectiveDate = effectiveDate.AddYears(-1);

        try
        {
            SetBusy(true, "Extracting financial terms from the selected PDF...");
            var extraction = _financialArrangementService.ExtractSourceTerms(dialog.FileName, effectiveDate);
            var summary = string.Join(Environment.NewLine, extraction.Charges.Select(charge => $"• {charge.ChargeType}: {(charge.Rate.HasValue ? Percent(charge.Rate) : charge.Formula)}")) +
                          string.Join(Environment.NewLine, extraction.Pooling.Select(pool => $"• Pooling {pool.Benefit}: {Percent(pool.Rate)} / {Money(pool.Threshold)}"));
            if (MessageBox.Show($"The following terms were extracted:\n\n{summary}\n\nSave these as {scenario}?", "Confirm extracted financial terms", MessageBoxButton.YesNo, MessageBoxImage.Information) != MessageBoxResult.Yes) return;

            var arrangement = new FinancialArrangement
            {
                ClientId = _workspace.DatabaseClientId.Value,
                Carrier = _workspace.PrimaryCarrier,
                EffectiveDate = effectiveDate,
                Source = Path.GetFileName(dialog.FileName),
                TargetLossRatio = extraction.TargetLossRatio,
                TlrBasis = extraction.TlrBasis,
                TargetLossRatioSourceText = extraction.TargetLossRatioSourceText,
                TargetLossRatioSourcePage = extraction.TargetLossRatioSourcePage,
                TargetLossRatioExtractionConfidence = extraction.TargetLossRatioExtractionConfidence
            };
            arrangement.Charges.AddRange(extraction.Charges);

            if (extraction.HasIncompleteTargetLossRatioEvidence)
            {
                await _financialArrangementService.SaveFinancialArrangementForReviewAsync(
                    arrangement,
                    "Target loss ratio evidence is incomplete and requires confirmation.");
                await LoadFinancialArrangementComparisonAsync(_workspace);
                MessageBox.Show(
                    "The incomplete TLR evidence was saved for review. It did not replace the client's verified financial terms.",
                    "Financial Terms Requires Review",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            var authorization = FinancialArrangementAuthorizationContext.Allow("Desktop user", FinancialArrangementPermissions.FinancialArrangementConfirm, FinancialArrangementPermissions.PoolingConfirm);
            arrangement = await _financialArrangementService.ConfirmFinancialArrangementAsync(arrangement, authorization, "Confirmed from financial arrangement PDF");
            PoolingArrangement? pooling = extraction.Pooling.FirstOrDefault();
            if (pooling is not null)
            {
                pooling.ClientId = _workspace.DatabaseClientId.Value; pooling.Carrier = _workspace.PrimaryCarrier; pooling.EffectiveDate = effectiveDate;
                pooling = await _financialArrangementService.ConfirmPoolingArrangementAsync(pooling, authorization, "Confirmed from financial arrangement PDF");
            }
            var version = await _financialArrangementService.CreateVersionAsync(new FinancialArrangementVersion { ClientId = _workspace.DatabaseClientId.Value, Carrier = _workspace.PrimaryCarrier, Scenario = scenario, EffectiveDate = effectiveDate, FinancialArrangementId = arrangement.FinancialArrangementId, PoolingArrangementId = pooling?.PoolingArrangementId });
            var current = await _financialArrangementService.GetLatestVersionAsync(_workspace.DatabaseClientId.Value, _workspace.PrimaryCarrier, FinancialArrangementScenario.Current);
            var proposed = await _financialArrangementService.GetLatestVersionAsync(_workspace.DatabaseClientId.Value, _workspace.PrimaryCarrier, FinancialArrangementScenario.InitialRenewal);
            var renewal = (await _databaseService.GetRenewalQueueAsync()).Where(item => item.ClientId == _workspace.DatabaseClientId.Value).OrderByDescending(item => item.RenewalDate).FirstOrDefault();
            if (current is not null && proposed is not null && renewal?.RenewalId > 0)
                await _financialArrangementService.CreateSnapshotAsync(new RenewalFinancialArrangementSnapshot { RenewalId = renewal.RenewalId, ClientId = _workspace.DatabaseClientId.Value, Carrier = _workspace.PrimaryCarrier, Scenario = FinancialArrangementScenario.InitialRenewal, AsOfUtc = DateTime.UtcNow, CurrentArrangementVersionId = current.FinancialArrangementVersionId, ProposedArrangementVersionId = proposed.FinancialArrangementVersionId });
            await LoadFinancialArrangementComparisonAsync(_workspace);
            MessageBox.Show(current is not null && proposed is not null ? "Terms were confirmed and the comparison snapshot was created." : "Terms were confirmed. Import the other document type to create the comparison snapshot.", "Financial Terms", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception exception) { AppLogger.Error(exception); MessageBox.Show(exception.Message, "Financial Terms Import", MessageBoxButton.OK, MessageBoxImage.Error); }
        finally { SetBusy(false); }
    }

    private async void SetCommission_Click(object sender, RoutedEventArgs e)
    {
        if (_isBusy || _workspace?.DatabaseClientId is not > 0) return;
        var choice = MessageBox.Show("Set commission for the current active terms?\n\nChoose Yes for Current, No for Renewal, or Cancel.", "Set commission", MessageBoxButton.YesNoCancel, MessageBoxImage.Question);
        if (choice == MessageBoxResult.Cancel) return;
        var scenario = choice == MessageBoxResult.Yes ? FinancialArrangementScenario.Current : FinancialArrangementScenario.InitialRenewal;
        var version = await _financialArrangementService.GetLatestVersionAsync(_workspace.DatabaseClientId.Value, _workspace.PrimaryCarrier, scenario);
        if (version is null) { MessageBox.Show("Import or confirm the carrier financial terms for this scenario first.", "Set commission", MessageBoxButton.OK, MessageBoxImage.Information); return; }
        var effective = version.EffectiveDate;
        var imported = scenario == FinancialArrangementScenario.InitialRenewal
            ? await _financialArrangementComparisonDataService.FindLatestSoldProposalCommissionAsync(_workspace.DatabaseClientId.Value)
            : null;
        decimal? rate = imported;
        var source = imported.HasValue ? "Sold proposal" : "Manual confirmation";
        if (imported.HasValue && MessageBox.Show($"A {Percent(imported)} commission was found in the sold proposal. Use it?", "Sold proposal commission", MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes) rate = null;
        if (!rate.HasValue)
        {
            var text = Interaction.InputBox("Enter commission as a percent, for example 5 or 5.0.", "Set commission", "");
            if (string.IsNullOrWhiteSpace(text)) return;
            if (!decimal.TryParse(text, NumberStyles.Number, CanadianCulture, out var percent) || percent < 0m || percent > 100m) { MessageBox.Show("Enter a commission between 0 and 100.", "Set commission", MessageBoxButton.OK, MessageBoxImage.Warning); return; }
            rate = percent / 100m;
        }
        var basis = Interaction.InputBox("Enter the commission basis.", "Set commission", "Total benefit payments");
        if (string.IsNullOrWhiteSpace(basis)) return;
        try
        {
            SetBusy(true, "Saving commission...");
            var authorization = FinancialArrangementAuthorizationContext.Allow("Desktop user", FinancialArrangementPermissions.CommissionConfirm);
            var commission = await _financialArrangementService.ConfirmCommissionArrangementAsync(new CommissionArrangement { ClientId=_workspace.DatabaseClientId.Value, Carrier=_workspace.PrimaryCarrier, Rate=rate, Basis=basis, Source=source, ConfirmationStatus="Confirmed", EffectiveDate=effective }, authorization, $"{source} commission confirmed in Financial Arrangement Comparison");
            await _financialArrangementService.CreateVersionAsync(new FinancialArrangementVersion { ClientId=version.ClientId, Carrier=version.Carrier, Scenario=scenario, EffectiveDate=version.EffectiveDate, FinancialArrangementId=version.FinancialArrangementId, CommissionArrangementId=commission.CommissionArrangementId, PoolingArrangementId=version.PoolingArrangementId });
            await RefreshFinancialArrangementSnapshotAsync(_workspace);
            await LoadFinancialArrangementComparisonAsync(_workspace);
            MessageBox.Show($"{Percent(rate)} commission was saved for {scenario}.", "Set commission", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception exception) { AppLogger.Error(exception); MessageBox.Show(exception.Message, "Set commission", MessageBoxButton.OK, MessageBoxImage.Error); }
        finally { SetBusy(false); }
    }

    private async Task RefreshFinancialArrangementSnapshotAsync(ClientWorkspaceViewModel workspace)
    {
        if (workspace.DatabaseClientId is not > 0) return;
        var current = await _financialArrangementService.GetLatestVersionAsync(workspace.DatabaseClientId.Value, workspace.PrimaryCarrier, FinancialArrangementScenario.Current);
        var proposed = await _financialArrangementService.GetLatestVersionAsync(workspace.DatabaseClientId.Value, workspace.PrimaryCarrier, FinancialArrangementScenario.InitialRenewal);
        if (current is null || proposed is null) return;
        var renewal = (await _databaseService.GetRenewalQueueAsync()).Where(item => item.ClientId == workspace.DatabaseClientId.Value).OrderByDescending(item => item.RenewalDate).FirstOrDefault();
        if (renewal?.RenewalId is not > 0) return;
        var latest = (await _financialArrangementService.GetSnapshotsAsync(renewal.RenewalId)).OrderByDescending(item => item.AsOfUtc).FirstOrDefault();
        var authorization = FinancialArrangementAuthorizationContext.Allow("Desktop user", FinancialArrangementPermissions.RenewalFinancialComparisonRefresh);
        if (latest is null)
            await _financialArrangementService.CreateSnapshotAsync(new RenewalFinancialArrangementSnapshot { RenewalId=renewal.RenewalId, ClientId=workspace.DatabaseClientId.Value, Carrier=workspace.PrimaryCarrier, Scenario=FinancialArrangementScenario.InitialRenewal, AsOfUtc=DateTime.UtcNow, CurrentArrangementVersionId=current.FinancialArrangementVersionId, ProposedArrangementVersionId=proposed.FinancialArrangementVersionId });
        else if (latest.SnapshotStatus.Equals("Finalized", StringComparison.OrdinalIgnoreCase))
            return;
        else
            await _financialArrangementService.RefreshSnapshotAsync(latest.RenewalFinancialArrangementSnapshotId, current.FinancialArrangementVersionId, proposed.FinancialArrangementVersionId, authorization);
    }

    private async Task LoadFinancialArrangementComparisonAsync(ClientWorkspaceViewModel workspace)
    {
        FinancialArrangementComparisonPanel.Visibility = Visibility.Visible;
        ArrangementAdminTextBlock.Text = "N/A";
        ArrangementCommissionTextBlock.Text = "N/A";
        ArrangementLoadTextBlock.Text = "N/A";
        ArrangementPoolingTextBlock.Text = "N/A";
        ArrangementCostTextBlock.Text = "N/A";
        ArrangementCostBasisTextBlock.Text = "Select a forecast window";
        FinancialArrangementStaleBorder.Visibility = Visibility.Collapsed;
        if (workspace.DatabaseClientId is not > 0)
        {
            FinancialArrangementComparisonMessageTextBlock.Text = "Link this client to SQLite before creating a financial arrangement comparison snapshot.";
            FinancialArrangementSnapshotGrid.ItemsSource = null;
            return;
        }
        var snapshots = await _financialArrangementService.GetSnapshotsForClientAsync(workspace.DatabaseClientId.Value);
        FinancialArrangementSnapshotGrid.ItemsSource = snapshots.Select(FinancialArrangementSnapshotRow.FromSnapshot).ToList();
        var latest = snapshots.FirstOrDefault();
        FinancialArrangementComparisonMessageTextBlock.Text = latest is null
            ? "No financial arrangement snapshot exists yet. Confirm canonical terms, then create an Initial Renewal snapshot."
            : $"Showing {snapshots.Count:#,0} historical financial arrangement snapshot(s). Select a renewal workflow to create or refresh comparison terms.";
        FinancialArrangementStaleBorder.Visibility = snapshots.Any(snapshot => snapshot.IsStale) ? Visibility.Visible : Visibility.Collapsed;
        FinancialArrangementLoadGrid.ItemsSource = null;
        if (latest is null) return;
        var comparison = await _financialArrangementComparisonDataService.LoadAsync(latest.RenewalFinancialArrangementSnapshotId);
        if (comparison is null)
        {
            FinancialArrangementComparisonMessageTextBlock.Text = "The latest snapshot references records that are no longer available. Its history remains visible, but the comparison needs review.";
            return;
        }
        ArrangementAdminTextBlock.Text = DisplayRange(comparison.AdministrativeLoadComparisons.FirstOrDefault()?.CurrentCarrierAdministrationRate, comparison.AdministrativeLoadComparisons.FirstOrDefault()?.ProposedCarrierAdministrationRate, Percent);
        ArrangementCommissionTextBlock.Text = DisplayRange(comparison.CurrentCommissionRate, comparison.ProposedCommissionRate, Percent);
        ArrangementLoadTextBlock.Text = DisplayRange(comparison.AdministrativeLoadComparisons.FirstOrDefault()?.CurrentTotalAdministrativeLoadRate, comparison.AdministrativeLoadComparisons.FirstOrDefault()?.ProposedTotalAdministrativeLoadRate, Percent);
        ArrangementPoolingTextBlock.Text = DisplayRange(comparison.CurrentPoolingRate, comparison.ProposedPoolingRate, Percent);
        FinancialArrangementLoadGrid.ItemsSource = comparison.AdministrativeLoadComparisons.Select(FinancialArrangementLoadRow.FromComparison).ToList();

        var forecast = BuildAsoForecastWindow();
        if (!forecast.HasSourceMonths)
        {
            ArrangementCostTextBlock.Text = "No ASO activity";
            ArrangementCostBasisTextBlock.Text = ForecastWindowLabel(forecast.RequestedMonths);
            FinancialArrangementComparisonMessageTextBlock.Text += " Projected annual cost is waiting for saved ASO activity.";
            return;
        }

        ArrangementCostBasisTextBlock.Text = $"{ForecastWindowLabel(forecast.RequestedMonths)} ending {forecast.EndMonth!.Value.ToString("MMM yyyy", CanadianCulture)}";
        if (forecast.MissingMonths.Count > 0)
        {
            ArrangementCostTextBlock.Text = "Missing ASO months";
            FinancialArrangementComparisonMessageTextBlock.Text += $" {ForecastWindowLabel(forecast.RequestedMonths)} cannot be annualized because it is missing: {string.Join(", ", forecast.MissingMonths.Select(month => month.ToString("MMMM yyyy", CanadianCulture)))}.";
            return;
        }

        if (forecast.MissingBenefitPaymentMonths.Count > 0)
        {
            ArrangementCostTextBlock.Text = "ASO totals incomplete";
            FinancialArrangementComparisonMessageTextBlock.Text += $" {ForecastWindowLabel(forecast.RequestedMonths)} cannot be annualized because Total Benefit Payment is missing for: {string.Join(", ", forecast.MissingBenefitPaymentMonths.Select(month => month.ToString("MMMM yyyy", CanadianCulture)))}.";
            return;
        }

        var request = await _financialArrangementComparisonDataService.LoadProjectionRequestAsync(
            latest.RenewalFinancialArrangementSnapshotId,
            BuildHistoricalPeriodSources(forecast),
            AnnualizedAsoInputs(forecast));
        if (request is null)
        {
            ArrangementCostTextBlock.Text = "Input review required";
            FinancialArrangementComparisonMessageTextBlock.Text += " The referenced pricing terms could not be loaded.";
            return;
        }

        ShowFinancialProjections(_financialProjectionCalculationService.Calculate(request));
    }

    private AsoAnnualForecastWindowResult BuildAsoForecastWindow()
    {
        var observations = _statementHistory
            .SelectMany(statement => statement.Periods
                .Where(period => !period.IsTotal && period.ActivityMonth.HasValue)
                .Select(period => new AsoFinancialActivityMonthObservation(
                    period.ActivityMonth!.Value,
                    period.TotalBenefitPayment,
                    statement.SourceFileName,
                    statement.ImportedAt)));
        return _forecastService.BuildRollingWindow(observations, SelectedForecastMonths());
    }

    private static FinancialCalculationInputSet AnnualizedAsoInputs(AsoAnnualForecastWindowResult forecast)
    {
        var inputs = new FinancialCalculationInputSet();
        if (forecast.AnnualizedBenefitPayments is not { } benefitPayments || benefitPayments <= 0m) return inputs;
        inputs.ValidatedAsoActivity.Add(new FinancialCalculationInput
        {
            Basis = $"Annualized total benefit payments ({forecast.RequestedMonths}-month rolling window)",
            Value = benefitPayments,
            Source = $"Combined saved ASO Activity: {string.Join(", ", forecast.Observations.Select(item => item.SourceFileName).Where(name => !string.IsNullOrWhiteSpace(name)).Distinct())}",
            Period = $"{Date(forecast.StartMonth)} to {Date(forecast.EndMonth)}",
            IsValidated = true
        });
        return inputs;
    }

    private IReadOnlyList<FinancialHistoricalPeriodSource> BuildHistoricalPeriodSources(AsoAnnualForecastWindowResult forecast)
    {
        var sourceByMonth = _statementHistory
            .SelectMany(statement => statement.Periods
                .Where(period => !period.IsTotal && period.ActivityMonth.HasValue)
                .Select(period => new FinancialHistoricalPeriodSource
                {
                    Period = period,
                    StatementId = statement.StatementId,
                    SourceDocumentId = statement.SourceDocumentId,
                    SourceVersionId = "ASO statement " + statement.StatementId + " v" + statement.VersionNumber,
                    EvidenceId = statement.SourceFileName
                }))
            .GroupBy(source => source.Period.ActivityMonth!.Value)
            .ToDictionary(group => group.Key, group => group
                .OrderByDescending(source => _statementHistory.First(statement => statement.StatementId == source.StatementId).ImportedAt)
                .ThenByDescending(source => source.EvidenceId, StringComparer.OrdinalIgnoreCase)
                .First());
        return forecast.Observations
            .Where(observation => sourceByMonth.ContainsKey(observation.ActivityMonth))
            .Select(observation => sourceByMonth[observation.ActivityMonth])
            .ToList();
    }

    private void ShowFinancialProjections(IReadOnlyList<FinancialProjectionResult> projections)
    {
        var historical = projections.Single(result => result.ProjectionType == FinancialProjectionType.HistoricalAnnualizedRunRate);
        var current = projections.Single(result => result.ProjectionType == FinancialProjectionType.CurrentTermsProjectedCost);
        var renewal = projections.Single(result => result.ProjectionType == FinancialProjectionType.RenewalTermsProjectedCost);
        var projectedClaims = current.Components.FirstOrDefault(component => component.ComponentType == FinancialCostComponentType.Claims)?.CalculatedAmount;
        ProjectedClaimsTextBlock.Text = Money(projectedClaims);
        HistoricalRunRateTextBlock.Text = DisplayProjectionTotal(historical);
        CurrentTermsProjectedCostTextBlock.Text = DisplayProjectionTotal(current);
        RenewalTermsProjectedCostTextBlock.Text = DisplayProjectionTotal(renewal);
        decimal? change = current.TotalClientCost.HasValue && renewal.TotalClientCost.HasValue ? renewal.TotalClientCost.Value - current.TotalClientCost.Value : null;
        RenewalCostChangeTextBlock.Text = Money(change);
        RenewalCostChangePercentTextBlock.Text = current.TotalClientCost is { } currentTotal && currentTotal > 0m && change.HasValue ? Percent(change.Value / currentTotal) : "Not Available";

        ArrangementCostTextBlock.Text = "See terms cost";
        FinancialProjectionBreakdownGrid.ItemsSource = FinancialProjectionBreakdownRow.Build(historical, current, renewal);
        var warnings = projections.SelectMany(result => result.Validations).Where(validation => validation.RequiresReview).Select(validation => validation.Message).Distinct(StringComparer.Ordinal).ToList();
        FinancialArrangementComparisonMessageTextBlock.Text += " Historical Annualized Run Rate annualizes actual recurring ASO activity from the selected completed-month window. Current and renewal projections apply the applicable pricing terms to projected claims.";
        if (warnings.Count > 0) FinancialArrangementComparisonMessageTextBlock.Text += " Input review required: " + string.Join(" ", warnings);
    }

    private static string DisplayProjectionTotal(FinancialProjectionResult result) =>
        result.TotalClientCost.HasValue ? Money(result.TotalClientCost) + (result.RequiresInputReview ? " (review)" : "") : "Not Available";

    private int SelectedForecastMonths() =>
        int.TryParse(ForecastWindowComboBox?.SelectedValue?.ToString(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var months) && months is 6 or 9 or 12
            ? months
            : 12;

    private static string ForecastWindowLabel(int months) => $"Last {months} months";

    private async void ForecastWindowComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (!IsLoaded || _isBusy || _workspace is null || FinancialArrangementComparisonPanel is null || !FinancialArrangementComparisonPanel.IsVisible) return;
        await LoadFinancialArrangementComparisonAsync(_workspace);
    }

    private void LoadStatementHistory(AsoFinancialActivityStatement statement)
    {
        _loadingStatementHistory = true;
        StatementHistoryComboBox.ItemsSource = _statementHistory.Select(item => new StatementHistoryItem(item)).ToList();
        StatementHistoryComboBox.SelectedItem = StatementHistoryComboBox.Items.Cast<StatementHistoryItem>()
            .FirstOrDefault(item => item.Statement.StatementId == statement.StatementId);
        _loadingStatementHistory = false;
        ShowStatement(statement);
    }

    private void StatementHistoryComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_loadingStatementHistory || StatementHistoryComboBox.SelectedItem is not StatementHistoryItem item) return;
        ShowStatement(item.Statement);
    }

    private void ShowStatement(AsoFinancialActivityStatement statement)
    {
        EmptyStateBorder.Visibility = Visibility.Collapsed;
        DashboardContentPanel.Visibility = Visibility.Visible;

        var latestPeriod = statement.Periods
            .OrderByDescending(period => period.ActivityMonth)
            .ThenByDescending(period => period.DisplayOrder)
            .FirstOrDefault();
        var total = statement.Total;

        CurrentBalanceTextBlock.Text = Money(latestPeriod?.CurrentBalance);
        BalanceStatusTextBlock.Text = Display(latestPeriod?.BalanceStatus);
        CurrentBalanceTextBlock.Foreground = latestPeriod?.BalanceStatus.Equals("Deficit", StringComparison.OrdinalIgnoreCase) == true
            ? (System.Windows.Media.Brush)FindResource("FinancialRedBrush")
            : (System.Windows.Media.Brush)FindResource("FinancialGreenBrush");
        LossRatioTextBlock.Text = Percent(total?.LossRatio);
        BenefitPaymentsTextBlock.Text = Money(total?.TotalBenefitPayment);
        AdminTaxesTextBlock.Text = Money((total?.AdminFees ?? 0m) + (total?.TotalTaxes ?? 0m));
        InterestTextBlock.Text = Money(total?.InterestCreditCharge);
        InterestStatusTextBlock.Text = total?.InterestCreditCharge < 0 ? "Credit" : "Charge";
        InterestTextBlock.Foreground = total?.InterestCreditCharge < 0
            ? (System.Windows.Media.Brush)FindResource("FinancialGreenBrush")
            : (System.Windows.Media.Brush)FindResource("FinancialTextBrush");
        RequiredFloatTextBlock.Text = Money(latestPeriod?.RequiredFloat);

        PopulateExecutiveFundingSummary(statement, latestPeriod, total);

        var sourceMonths = statement.Periods.Select(period => period.ActivityMonth).Where(month => month.HasValue).Select(month => month!.Value).Distinct().Count();
        PeriodTextBlock.Text = $"{Date(statement.PeriodStart)} to {Date(statement.PeriodEnd)} ({sourceMonths:#,0} monthly rows)";
        ExperiencePeriodStatusTextBlock.Text = ExperiencePeriodStatus(statement, sourceMonths);
        HistoricalComparisonTextBlock.Text = HistoricalComparison(statement);
        SourceTextBlock.Text = $"Source: {Display(statement.SourceFileName)} | Sheet: {Display(statement.SourceSheet)}";
        ContractTextBlock.Text =
            $"Contract: {Display(statement.ContractNumber)} | Carrier: {Display(statement.Carrier)} | " +
            $"Carrier funding: {Display(statement.SourceFundingStatus)} | Statement: {Date(statement.StatementDate)}";
        VersionStatusTextBlock.Text =
            $"Version {statement.VersionNumber} | {statement.ImportStatus} | Imported {Imported(statement.ImportedAt)}";
        ValidationTextBlock.Text = string.IsNullOrWhiteSpace(statement.ValidationSummary)
            ? "Validation complete"
            : statement.ValidationSummary;
        ImportStatusTextBlock.Text = $"{statement.Periods.Count:#,0} monthly rows loaded from SQLite.";

        MonthlyActivityGrid.ItemsSource = statement.Periods
            .OrderBy(period => period.DisplayOrder)
            .Select(FinancialActivityRow.FromPeriod)
            .ToList();
        RowCountTextBlock.Text = $"{statement.Periods.Count:#,0} monthly rows";
    }

    private void PopulateExecutiveFundingSummary(
        AsoFinancialActivityStatement statement,
        AsoFinancialActivityPeriod? latestPeriod,
        AsoFinancialActivityPeriod? total)
    {
        var balance = latestPeriod?.CurrentBalance;
        var sourceMonths = statement.Periods.Count(period => !period.IsTotal && period.ActivityMonth.HasValue);
        var lossRatio = total?.LossRatio;
        var fundingStatus = AsoFundingStatusService.Evaluate(latestPeriod);

        FundingStatusTextBlock.Text = fundingStatus.StatusTitle;
        FundingStatusTextBlock.Foreground = (System.Windows.Media.Brush)FindResource(
            fundingStatus.IsPositive ? "FinancialGreenBrush" : "FinancialPurpleBrush");
        FundingStatusDetailTextBlock.Text = fundingStatus.Detail;
        FundingInsightOneTextBlock.Text = fundingStatus.Insight;
        FundingInsightTwoTextBlock.Text = $"\u2022 {sourceMonths:#,0} saved activity month(s); incomplete windows are not annualized for pricing.";
        FundingRecommendationTextBlock.Text = fundingStatus.Recommendation;
        ForecastBalanceTextBlock.Text = Money(balance);
        ForecastCoverageTextBlock.Text = $"{sourceMonths:#,0} month(s)";
        ForecastStatusTextBlock.Text = "Latest confirmed balance only. A projected balance is shown only when a complete saved ASO window supports it.";
        ClaimsRiskTextBlock.Text = lossRatio.HasValue
            ? lossRatio <= 0.75m ? "Lower claims pressure" : lossRatio <= 1m ? "Claims within funding range" : "Claims require review"
            : "Loss ratio unavailable";
        ClaimsRiskDetailTextBlock.Text = lossRatio.HasValue
            ? $"Saved ASO loss ratio: {Percent(lossRatio)}. Review the monthly activity and source evidence below."
            : "Load a saved ASO statement to assess the claims-to-deposits relationship.";
    }

    private string ExperiencePeriodStatus(AsoFinancialActivityStatement statement, int sourceMonths)
    {
        if (_renewalExperiencePeriod?.PeriodStart is not { } start || _renewalExperiencePeriod.PeriodEnd is not { } end)
            return $"Coverage: {sourceMonths:#,0} source month(s). Renewal experience period is not recorded in the processed renewal document, so this period is not confirmed.";

        var matches = statement.PeriodStart == start && statement.PeriodEnd == end;
        return matches
            ? $"Coverage: {sourceMonths:#,0} source month(s). Confirmed against {_renewalExperiencePeriod.SourceDescription}: {Date(start)} to {Date(end)}."
            : $"Coverage: {sourceMonths:#,0} source month(s). Does not match {_renewalExperiencePeriod.SourceDescription}: {Date(start)} to {Date(end)}.";
    }

    private string HistoricalComparison(AsoFinancialActivityStatement statement)
    {
        var prior = _statementHistory
            .Where(item => item.StatementId != statement.StatementId && item.PeriodEnd < statement.PeriodStart)
            .OrderByDescending(item => item.PeriodEnd)
            .FirstOrDefault();
        if (prior?.Total?.TotalBenefitPayment is not { } priorPaid || statement.Total?.TotalBenefitPayment is not { } paid)
            return "Historical comparison: select another saved statement to review its period and totals.";

        var difference = paid - priorPaid;
        return $"Historical comparison: benefit payments are {Money(Math.Abs(difference))} {(difference >= 0 ? "higher" : "lower")} than {Date(prior.PeriodStart)} to {Date(prior.PeriodEnd)}.";
    }

    private sealed class StatementHistoryItem(AsoFinancialActivityStatement statement)
    {
        public AsoFinancialActivityStatement Statement { get; } = statement;
        public string DisplayName => $"{Date(Statement.PeriodStart)} to {Date(Statement.PeriodEnd)} | {Statement.SourceFileName}";
    }

    private void ShowEmpty(string title, string message)
    {
        DashboardContentPanel.Visibility = Visibility.Collapsed;
        EmptyStateBorder.Visibility = Visibility.Visible;
        EmptyStateTitleTextBlock.Text = title;
        EmptyStateMessageTextBlock.Text = message;
    }

    private void SetBusy(bool busy, string? status = null)
    {
        _isBusy = busy;
        if (!string.IsNullOrWhiteSpace(status))
        {
            ImportStatusTextBlock.Text = status;
        }
    }

    private static string Money(decimal? value)
    {
        return value.HasValue
            ? value.Value.ToString("$#,0.00;($#,0.00);$0.00", CanadianCulture)
            : "N/A";
    }

    private static string Percent(decimal? value)
    {
        return value.HasValue
            ? value.Value.ToString("P1", CanadianCulture)
            : "N/A";
    }

    private static string Date(DateOnly? value)
    {
        return value?.ToString("MMM d, yyyy", CanadianCulture) ?? "Not available";
    }

    private static string Imported(DateTime value)
    {
        return value == default
            ? "date unavailable"
            : value.ToLocalTime().ToString("MMM d, yyyy h:mm tt", CanadianCulture);
    }

    private static string Display(string? value)
    {
        return string.IsNullOrWhiteSpace(value) ? "Not available" : value.Trim();
    }

    private static string DisplayRange(decimal? current, decimal? proposed, Func<decimal?, string> formatter) => $"{formatter(current)} → {formatter(proposed)}";

    private sealed class FinancialActivityRow
    {
        public string Month { get; init; } = "";
        public string OpeningBalance { get; init; } = "";
        public string TotalDeposit { get; init; } = "";
        public string UnpooledBenefits { get; init; } = "";
        public string PooledBenefits { get; init; } = "";
        public string TotalBenefits { get; init; } = "";
        public string LossRatio { get; init; } = "";
        public string PoolingCharges { get; init; } = "";
        public string AdminFees { get; init; } = "";
        public string OtherCharges { get; init; } = "";
        public string TotalTaxes { get; init; } = "";
        public string Interest { get; init; } = "";
        public string TotalCharges { get; init; } = "";
        public string CurrentBalance { get; init; } = "";
        public string Status { get; init; } = "";

        public static FinancialActivityRow FromPeriod(AsoFinancialActivityPeriod period)
        {
            return new FinancialActivityRow
            {
                Month = period.ActivityMonth?.ToString("MMM yyyy", CanadianCulture) ?? period.PeriodLabel,
                OpeningBalance = Money(period.OpeningBalance),
                TotalDeposit = Money(period.TotalDeposit),
                UnpooledBenefits = Money(period.UnpooledBenefitPayments),
                PooledBenefits = Money(period.PooledBenefitPayments),
                TotalBenefits = Money(period.TotalBenefitPayment),
                LossRatio = Percent(period.LossRatio),
                PoolingCharges = Money(period.PoolingCharges),
                AdminFees = Money(period.AdminFees),
                OtherCharges = Money(period.OtherCharges),
                TotalTaxes = Money(period.TotalTaxes),
                Interest = Money(period.InterestCreditCharge),
                TotalCharges = Money(period.TotalCharges),
                CurrentBalance = Money(period.CurrentBalance),
                Status = Display(period.BalanceStatus)
            };
        }
    }

    private sealed class FinancialArrangementSnapshotRow
    {
        public string Renewal { get; init; } = "";
        public string Scenario { get; init; } = "";
        public string SnapshotType { get; init; } = "";
        public string Status { get; init; } = "";
        public string CurrentVersion { get; init; } = "";
        public string ProposedVersion { get; init; } = "";
        public string Created { get; init; } = "";
        public string FinalizedBy { get; init; } = "";

        public static FinancialArrangementSnapshotRow FromSnapshot(RenewalFinancialArrangementSnapshot snapshot) => new()
        {
            Renewal = snapshot.RenewalId.ToString(CanadianCulture),
            Scenario = snapshot.Scenario.ToString(),
            SnapshotType = snapshot.SnapshotType,
            Status = snapshot.IsStale ? "Needs review" : snapshot.SnapshotStatus,
            CurrentVersion = snapshot.CurrentArrangementVersionId.ToString(CanadianCulture),
            ProposedVersion = snapshot.ProposedArrangementVersionId.ToString(CanadianCulture),
            Created = snapshot.AsOfUtc.ToLocalTime().ToString("MMM d, yyyy h:mm tt", CanadianCulture),
            FinalizedBy = Display(snapshot.FinalizedBy)
        };
    }

    private sealed class FinancialArrangementLoadRow
    {
        public string Benefit { get; init; } = "";
        public string CurrentAdministration { get; init; } = "";
        public string ProposedAdministration { get; init; } = "";
        public string CurrentCommission { get; init; } = "";
        public string ProposedCommission { get; init; } = "";
        public string CurrentLoad { get; init; } = "";
        public string ProposedLoad { get; init; } = "";
        public string Warning { get; init; } = "";
        public static FinancialArrangementLoadRow FromComparison(AdministrativeLoadComparison comparison) => new()
        {
            Benefit = comparison.Benefit,
            CurrentAdministration = Percent(comparison.CurrentCarrierAdministrationRate),
            ProposedAdministration = Percent(comparison.ProposedCarrierAdministrationRate),
            CurrentCommission = Percent(comparison.CurrentCommissionRate),
            ProposedCommission = Percent(comparison.ProposedCommissionRate),
            CurrentLoad = Percent(comparison.CurrentTotalAdministrativeLoadRate),
            ProposedLoad = Percent(comparison.ProposedTotalAdministrativeLoadRate),
            Warning = string.Join(" ", comparison.CalculationWarnings)
        };
    }

    private sealed class FinancialProjectionBreakdownRow
    {
        public string Component { get; init; } = "";
        public string Historical { get; init; } = "";
        public string Current { get; init; } = "";
        public string Renewal { get; init; } = "";
        public string Change { get; init; } = "";

        public static IReadOnlyList<FinancialProjectionBreakdownRow> Build(
            FinancialProjectionResult historical,
            FinancialProjectionResult current,
            FinancialProjectionResult renewal)
        {
            var rows = Enum.GetValues<FinancialCostComponentType>()
                .Select(component =>
                {
                    var historicalAmount = Amount(historical, component);
                    var currentAmount = Amount(current, component);
                    var renewalAmount = Amount(renewal, component);
                    return new FinancialProjectionBreakdownRow
                    {
                        Component = Label(component),
                        Historical = DisplayAmount(historicalAmount),
                        Current = DisplayAmount(currentAmount),
                        Renewal = DisplayAmount(renewalAmount),
                        Change = currentAmount.HasValue && renewalAmount.HasValue ? Money(renewalAmount.Value - currentAmount.Value) : "Not Available"
                    };
                })
                .ToList();
            rows.Add(new FinancialProjectionBreakdownRow
            {
                Component = "Total Client Cost",
                Historical = DisplayAmount(historical.TotalClientCost),
                Current = DisplayAmount(current.TotalClientCost),
                Renewal = DisplayAmount(renewal.TotalClientCost),
                Change = current.TotalClientCost.HasValue && renewal.TotalClientCost.HasValue ? Money(renewal.TotalClientCost.Value - current.TotalClientCost.Value) : "Not Available"
            });
            return rows;
        }

        private static decimal? Amount(FinancialProjectionResult result, FinancialCostComponentType type)
        {
            var components = result.Components.Where(component => component.ComponentType == type).ToList();
            return components.Count == 0 || components.All(component => !component.CalculatedAmount.HasValue)
                ? null
                : components.Where(component => component.CalculatedAmount.HasValue).Sum(component => component.CalculatedAmount!.Value);
        }

        private static string DisplayAmount(decimal? value) => value.HasValue ? Money(value) : "Not Available";

        private static string Label(FinancialCostComponentType component) => component switch
        {
            FinancialCostComponentType.OtherCharges => "Other Charges",
            _ => component.ToString()
        };
    }
}

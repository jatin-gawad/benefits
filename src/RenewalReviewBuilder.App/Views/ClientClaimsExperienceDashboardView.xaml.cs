using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;

namespace RenewalReviewBuilder.App.Views;

public partial class ClientClaimsExperienceDashboardView : UserControl
{
    public event EventHandler? UploadClaimsExperienceRequested;

    private ClientWorkspaceViewModel? _workspace;
    private IReadOnlyCollection<ClaimsExperienceRecord> _claims = [];
    private IReadOnlyCollection<MonthlyClaimsRecord> _monthlyClaims = [];
    private ClaimsExperienceRenewalTerms _renewalTerms = new();
    private ClientClaimsDashboardDataResult? _dashboardData;
    private bool _updatingPeriodOptions;

    public ClientClaimsExperienceDashboardView()
    {
        InitializeComponent();
        DataContext = ClientClaimsDashboardModel.Empty();
    }

    public void LoadDashboard(
        ClientWorkspaceViewModel workspace,
        IReadOnlyCollection<ClaimsExperienceRecord> claims,
        IReadOnlyCollection<MonthlyClaimsRecord> monthlyClaims,
        ClaimsExperienceRenewalTerms? renewalTerms = null,
        AsoFinancialActivityStatement? asoFinancialActivity = null,
        ClientClaimsDashboardDataResult? dashboardData = null)
    {
        var dashboardSources = AsoDentalClaimsDashboardAdapter.Merge(
            claims,
            monthlyClaims,
            asoFinancialActivity);
        _workspace = workspace;
        _claims = dashboardSources.Claims;
        _monthlyClaims = dashboardSources.MonthlyClaims;
        _renewalTerms = renewalTerms ?? new ClaimsExperienceRenewalTerms();
        _dashboardData = dashboardData;
        RefreshPeriodOptions();
        ApplyFilters();
    }

    public void LoadEmpty(ClientWorkspaceViewModel workspace)
    {
        _workspace = workspace;
        _claims = [];
        _monthlyClaims = [];
        _renewalTerms = new ClaimsExperienceRenewalTerms();
        _dashboardData = null;
        RefreshPeriodOptions();
        ApplyFilters();
    }

    private void Filters_Changed(object sender, SelectionChangedEventArgs e)
    {
        if (_updatingPeriodOptions)
        {
            return;
        }

        ApplyFilters();
    }

    private void ApplyFilters_Click(object sender, RoutedEventArgs e)
    {
        ApplyFilters();
    }

    private void ApplyFilters()
    {
        if (_workspace is null)
        {
            return;
        }

        var benefit = (BenefitFilterComboBox?.SelectedItem as ComboBoxItem)?.Tag as string ?? "all";
        var period = PeriodFilterComboBox?.SelectedItem as ClientClaimsPeriodFilterOption
                     ?? ClientClaimsPeriodFilterOption.All("All saved periods");
        var claims = _claims.Where(row => MatchesBenefit(row.BenefitLine, benefit)).ToList();
        var monthly = _monthlyClaims.Where(row => MatchesBenefit(row.BenefitLine, benefit)).ToList();

        if (period.Kind == ClientClaimsPeriodFilterKind.ExactPeriod &&
            period.Start.HasValue &&
            period.End.HasValue)
        {
            claims = claims
                .Where(row => row.ExperiencePeriodStart == period.Start &&
                              row.ExperiencePeriodEnd == period.End)
                .ToList();
            monthly = monthly
                .Where(row => row.Month >= period.Start && row.Month <= period.End)
                .ToList();

            var expectedMonths = MonthsInclusive(period.Start.Value, period.End.Value);
            var savedMonths = monthly
                .Where(row => row.Month.HasValue)
                .Select(row => row.Month!.Value)
                .Distinct()
                .Count();
            if (savedMonths < expectedMonths)
            {
                monthly = [];
            }
        }
        else if (period.Kind == ClientClaimsPeriodFilterKind.LatestMonthly &&
                 period.Start.HasValue &&
                 period.End.HasValue)
        {
            claims = claims
                .Where(row => row.ExperiencePeriodStart >= period.Start &&
                              row.ExperiencePeriodEnd <= period.End)
                .ToList();
            monthly = monthly
                .Where(row => row.Month >= period.Start && row.Month <= period.End)
                .ToList();
        }
        else if (claims.Count > 0)
        {
            // All-period comparisons use one authoritative aggregate per period.
            // Mixing them with newer monthly rows would hide the older periods.
            monthly = [];
        }

        var averageSpendComparison = ClientClaimsDashboardModel.BuildAverageSpendComparison(
            period,
            benefit,
            _claims,
            _monthlyClaims);

        DataContext = ClientClaimsDashboardModel.FromClaims(
            _workspace, claims, monthly, averageSpendComparison, _renewalTerms, _dashboardData);
    }

    private void RefreshPeriodOptions()
    {
        if (PeriodFilterComboBox is null)
        {
            return;
        }

        var options = BuildPeriodOptions(_claims, _monthlyClaims);
        _updatingPeriodOptions = true;
        try
        {
            PeriodFilterComboBox.ItemsSource = options;
            PeriodFilterComboBox.SelectedIndex = options.Count > 0 ? 0 : -1;
        }
        finally
        {
            _updatingPeriodOptions = false;
        }
    }

    private static IReadOnlyList<ClientClaimsPeriodFilterOption> BuildPeriodOptions(
        IReadOnlyCollection<ClaimsExperienceRecord> claims,
        IReadOnlyCollection<MonthlyClaimsRecord> monthlyClaims)
    {
        var exactPeriods = ClientClaimsDashboardModel.SelectAuthoritativeClaimsRows(claims)
            .Where(row => row.ExperiencePeriodStart.HasValue && row.ExperiencePeriodEnd.HasValue)
            .Select(row => new
            {
                Start = row.ExperiencePeriodStart!.Value,
                End = row.ExperiencePeriodEnd!.Value
            })
            .Distinct()
            .OrderByDescending(period => period.End)
            .ThenByDescending(period => period.Start)
            .ToList();

        var options = new List<ClientClaimsPeriodFilterOption>();
        for (var index = 0; index < exactPeriods.Count; index++)
        {
            var period = exactPeriods[index];
            options.Add(ClientClaimsPeriodFilterOption.Exact(
                RankLabel(index),
                period.Start,
                period.End));
        }

        var monthlyDates = monthlyClaims
            .Where(row => row.Month.HasValue)
            .Select(row => row.Month!.Value)
            .Distinct()
            .OrderBy(month => month)
            .ToList();
        if (monthlyDates.Count > 0)
        {
            var latestMonth = monthlyDates[^1];
            var rollingStart = latestMonth.AddMonths(-11);
            var availableDates = monthlyDates.Where(month => month >= rollingStart).ToList();
            var actualStart = availableDates[0];
            var displayEnd = new DateOnly(
                latestMonth.Year,
                latestMonth.Month,
                DateTime.DaysInMonth(latestMonth.Year, latestMonth.Month));
            var label = availableDates.Count >= 12
                ? "Last 12 monthly months"
                : $"Latest monthly data ({availableDates.Count} months)";
            options.Add(ClientClaimsPeriodFilterOption.LatestMonthly(
                $"{label}: {FormatRange(actualStart, displayEnd)}",
                rollingStart,
                displayEnd));
        }

        var allStarts = exactPeriods.Select(period => period.Start)
            .Concat(monthlyDates)
            .ToList();
        var allEnds = exactPeriods.Select(period => period.End)
            .Concat(monthlyDates)
            .ToList();
        if (allStarts.Count > 0 && allEnds.Count > 0)
        {
            options.Add(ClientClaimsPeriodFilterOption.All(
                $"All saved periods: {FormatRange(allStarts.Min(), allEnds.Max())}"));
        }
        else
        {
            options.Add(ClientClaimsPeriodFilterOption.All("All saved periods"));
        }

        return options;
    }

    private static string RankLabel(int index)
    {
        return index switch
        {
            0 => "Current",
            1 => "Prior",
            2 => "Prior Prior",
            3 => "Prior Prior Prior",
            _ => $"Older period {index - 2}"
        };
    }

    private static string FormatRange(DateOnly start, DateOnly end)
    {
        return $"{start.ToString("MMM d, yyyy", CultureInfo.InvariantCulture)} - {end.ToString("MMM d, yyyy", CultureInfo.InvariantCulture)}";
    }

    private static int MonthsInclusive(DateOnly start, DateOnly end)
    {
        return Math.Max(0, ((end.Year - start.Year) * 12) + end.Month - start.Month + 1);
    }

    private static bool MatchesBenefit(string benefitLine, string selectedBenefit)
    {
        return selectedBenefit.Equals("all", StringComparison.OrdinalIgnoreCase) ||
               (selectedBenefit.Equals("ehc", StringComparison.OrdinalIgnoreCase) &&
                (benefitLine.Contains("health", StringComparison.OrdinalIgnoreCase) || benefitLine.Contains("ehc", StringComparison.OrdinalIgnoreCase))) ||
               (selectedBenefit.Equals("dental", StringComparison.OrdinalIgnoreCase) &&
                benefitLine.Contains("dental", StringComparison.OrdinalIgnoreCase));
    }

    private void UploadClaimsExperience_Click(object sender, RoutedEventArgs e)
    {
        UploadClaimsExperienceRequested?.Invoke(this, EventArgs.Empty);
    }

    private void Export_Click(object sender, RoutedEventArgs e)
    {
        MessageBox.Show(
            "Client claims export is ready to connect to the existing Claims Model and Claims PowerPoint export workflow. The bottom action bar still controls the live export actions.",
            "Client Claims Experience",
            MessageBoxButton.OK,
            MessageBoxImage.Information);
    }
}

public enum ClientClaimsPeriodFilterKind
{
    ExactPeriod,
    LatestMonthly,
    All
}

public sealed record ClientClaimsPeriodFilterOption(
    string DisplayLabel,
    ClientClaimsPeriodFilterKind Kind,
    DateOnly? Start,
    DateOnly? End)
{
    public static ClientClaimsPeriodFilterOption Exact(string rank, DateOnly start, DateOnly end)
    {
        return new ClientClaimsPeriodFilterOption(
            $"{rank}: {start.ToString("MMM d, yyyy", CultureInfo.InvariantCulture)} - {end.ToString("MMM d, yyyy", CultureInfo.InvariantCulture)}",
            ClientClaimsPeriodFilterKind.ExactPeriod,
            start,
            end);
    }

    public static ClientClaimsPeriodFilterOption LatestMonthly(string label, DateOnly start, DateOnly end)
    {
        return new ClientClaimsPeriodFilterOption(label, ClientClaimsPeriodFilterKind.LatestMonthly, start, end);
    }

    public static ClientClaimsPeriodFilterOption All(string label)
    {
        return new ClientClaimsPeriodFilterOption(label, ClientClaimsPeriodFilterKind.All, null, null);
    }
}

public sealed class ClientClaimsDashboardModel
{
    private static readonly Brush Purple = CreateFrozenBrush(109, 40, 217);
    private static readonly Brush Green = CreateFrozenBrush(22, 163, 74);
    private static readonly Brush Orange = CreateFrozenBrush(249, 115, 22);
    private static readonly Brush Red = CreateFrozenBrush(220, 38, 38);
    private static readonly Brush Muted = CreateFrozenBrush(100, 116, 139);

    public string HeaderSubtitle { get; init; } = "";
    public string ExperiencePeriodLabel { get; init; } = "All saved claims periods";
    public string TrendTitle { get; init; } = "Monthly Claims Trend (Paid)";
    public string TotalClaimsPaidText { get; init; } = "$0";
    public string TotalClaimsSubtext { get; init; } = "No saved claims";
    public string ClaimsPerEmployeeText { get; init; } = "N/A";
    public string MembersText { get; init; } = "No members saved";
    public string LossRatioText { get; init; } = "N/A";
    public string LossRatioSubtext { get; init; } = "Premium unavailable";
    public Brush LossRatioBrush { get; init; } = Muted;
    public string PlanMembersText { get; init; } = "N/A";
    public string PlanText { get; init; } = "Plan not set";
    public string AverageClaimText { get; init; } = "N/A";
    public string ClaimantsText { get; init; } = "N/A";
    public string ClaimantDataStatus { get; init; } = "Claimant table not saved";
    public string SourceText { get; init; } = "Source: SQLite ClaimsExperience and MonthlyClaims";
    public string DataAccuracyText { get; init; } = "Data accuracy: High";
    public Brush DataAccuracyBrush { get; init; } = Green;
    public bool HasData { get; init; }
    public bool IsEmpty => !HasData;
    public IReadOnlyList<ClientClaimsCategoryRow> CategoryRows { get; init; } = [];
    public IReadOnlyList<ClientClaimsBarRow> MonthlyTrendRows { get; init; } = [];
    public IReadOnlyList<ClientClaimsBarRow> LossRatioTrendRows { get; init; } = [];
    public IReadOnlyList<ClientClaimsMetricRow> UtilizationRows { get; init; } = [];
    public IReadOnlyList<ClientClaimsInsightRow> InsightRows { get; init; } = [];
    public IReadOnlyList<string> SummaryBullets { get; init; } = [];
    public ClientClaimsAverageSpendComparison AverageSpendComparison { get; init; } = ClientClaimsAverageSpendComparison.Empty();

    public static ClientClaimsDashboardModel Empty()
    {
        return new ClientClaimsDashboardModel
        {
            HeaderSubtitle = "Select a client to load client-specific claims experience.",
            HasData = false
        };
    }

    public static ClientClaimsDashboardModel FromClaims(
        ClientWorkspaceViewModel workspace,
        IReadOnlyCollection<ClaimsExperienceRecord> claims,
        IReadOnlyCollection<MonthlyClaimsRecord> monthlyClaims,
        ClientClaimsAverageSpendComparison? averageSpendComparison = null,
        ClaimsExperienceRenewalTerms? renewalTerms = null,
        ClientClaimsDashboardDataResult? dashboardData = null)
    {
        renewalTerms ??= new ClaimsExperienceRenewalTerms();
        var claimsRows = SelectAuthoritativeClaimsRows(claims.Where(row => row.ClientId > 0));
        var monthlyRows = SelectAuthoritativeMonthlyRows(monthlyClaims.Where(row => row.ClientId > 0));
        claimsRows = ExcludeAggregateRowsReplacedByMonthlyClaims(claimsRows, monthlyRows);
        var hasData = claimsRows.Count > 0 || monthlyRows.Count > 0;
        var latestInsuredLives = ResolveLatestLives(workspace, claimsRows, monthlyRows);
        var averageInsuredLives = ResolveAverageInsuredLives(workspace, claimsRows, monthlyRows);
        var paidClaims = SumPaidClaims(claimsRows, monthlyRows);
        var premium = SumPremium(claimsRows, monthlyRows, renewalTerms);
        var hasCompletePremium = HasCompletePremiumForPaidClaims(claimsRows, monthlyRows);
        var hasCompleteLives = HasCompleteLivesForPaidClaims(claimsRows, monthlyRows);
        var lossRatio = premium > 0m && hasCompletePremium
            ? paidClaims / premium * 100m
            : (decimal?)null;
        var targetLossRatio = ResolveTargetLossRatioPercent(claimsRows, monthlyRows, renewalTerms);
        var claimsPerCurrentEmployee = hasCompleteLives && latestInsuredLives > 0m && paidClaims > 0m
            ? paidClaims / latestInsuredLives
            : (decimal?)null;
        var averageSpendPerInsuredLife = hasCompleteLives && averageInsuredLives > 0m && paidClaims > 0m
            ? paidClaims / averageInsuredLives
            : (decimal?)null;
        var categoryRows = BuildCategoryRows(claimsRows, monthlyRows, paidClaims);
        var topCategory = categoryRows.FirstOrDefault();
        var period = ResolvePeriod(claimsRows, monthlyRows);
        var includesAsoOperatingAccounts = claimsRows.Any(IsAsoDashboardRow) || monthlyRows.Any(IsAsoDashboardRow);
        var latestLoaded = claimsRows.Select(row => row.DateLoaded)
            .Concat(monthlyRows.Select(row => row.DateLoaded))
            .Where(date => date > DateTime.MinValue)
            .DefaultIfEmpty(DateTime.MinValue)
            .Max();

        var model = new ClientClaimsDashboardModel
        {
            HeaderSubtitle = $"{workspace.ClientName} | Carrier: {DisplayOrDefault(workspace.PrimaryCarrier)} | Effective date: {DisplayOrDefault(workspace.PlanEffectiveDate)} | Plan: {DisplayOrDefault(workspace.PlanType)}",
            ExperiencePeriodLabel = period,
            TrendTitle = monthlyRows.Count > 0 && claimsRows.Count > 0
                ? "Monthly Claims Trend (Paid; monthly sources only)"
                : monthlyRows.Count > 0
                    ? "Monthly Claims Trend (Paid)"
                    : "Period Claims (Paid)",
            TotalClaimsPaidText = FormatMoney(paidClaims),
            TotalClaimsSubtext = hasData ? $"{claimsRows.Count + monthlyRows.Count:#,0} saved claims row(s)" : "No saved claims",
            ClaimsPerEmployeeText = claimsPerCurrentEmployee.HasValue ? FormatAverageMoney(claimsPerCurrentEmployee.Value) : "N/A",
            MembersText = latestInsuredLives > 0m
                ? hasCompleteLives
                    ? $"{FormatLives(latestInsuredLives)} latest insured lives"
                    : $"{FormatLives(latestInsuredLives)} latest insured lives (partial)"
                : "Insured lives not saved",
            LossRatioText = lossRatio.HasValue ? FormatPercent(lossRatio.Value) : "N/A",
            LossRatioSubtext = lossRatio.HasValue
                ? RiskLabel(lossRatio.Value, targetLossRatio)
                : includesAsoOperatingAccounts && !hasCompletePremium
                    ? "Mixed funding: ASO Dental premium unavailable"
                : hasCompletePremium
                    ? "Premium unavailable"
                    : "Premium unavailable for part of paid claims",
            LossRatioBrush = !lossRatio.HasValue
                ? Muted
                : lossRatio.Value >= 100m
                    ? Red
                    : lossRatio.Value > targetLossRatio
                        ? Orange
                        : Green,
            PlanMembersText = latestInsuredLives > 0m ? FormatLives(latestInsuredLives) : "N/A",
            PlanText = string.IsNullOrWhiteSpace(workspace.PlanType) ? "Plan not set" : workspace.PlanType,
            AverageClaimText = averageSpendPerInsuredLife.HasValue ? FormatAverageMoney(averageSpendPerInsuredLife.Value) : "N/A",
            ClaimantsText = hasCompleteLives && averageInsuredLives > 0m ? FormatLives(averageInsuredLives) : "N/A",
            ClaimantDataStatus = hasCompleteLives && averageInsuredLives > 0m
                ? "Average across selected period"
                : hasCompleteLives
                    ? "Insured lives not saved"
                    : "Lives unavailable for part of paid claims",
            SourceText = ResolveSourceText(dashboardData, includesAsoOperatingAccounts),
            DataAccuracyText = includesAsoOperatingAccounts && (!hasCompletePremium || !hasCompleteLives)
                ? "Data scope: paid claims loaded; premium/lives partial"
                : dashboardData?.DataStatusText ?? "Data accuracy: High",
            DataAccuracyBrush = (dashboardData?.NeedsReview ?? false) ||
                                includesAsoOperatingAccounts && (!hasCompletePremium || !hasCompleteLives)
                ? Orange
                : Green,
            HasData = hasData,
            CategoryRows = categoryRows,
            MonthlyTrendRows = BuildMonthlyTrendRows(claimsRows, monthlyRows),
            LossRatioTrendRows = BuildLossRatioTrendRows(claimsRows, monthlyRows, renewalTerms),
            UtilizationRows = BuildUtilizationRows(claimsRows, monthlyRows, premium, paidClaims, lossRatio, latestLoaded),
            SummaryBullets = BuildSummaryBullets(paidClaims, lossRatio, topCategory, monthlyRows.Count, claimsRows.Count),
            InsightRows = BuildInsights(workspace, paidClaims, lossRatio, targetLossRatio, categoryRows),
            AverageSpendComparison = averageSpendComparison ?? ClientClaimsAverageSpendComparison.Empty(period)
        };

        return model;
    }

    private static string ResolveSourceText(
        ClientClaimsDashboardDataResult? dashboardData,
        bool includesAsoOperatingAccounts)
    {
        var source = dashboardData?.SourceText ?? "Source: SQLite ClaimsExperience and MonthlyClaims";
        return includesAsoOperatingAccounts
            ? $"{source}; verified ASO operating accounts included"
            : source;
    }

    internal static ClientClaimsAverageSpendComparison BuildAverageSpendComparison(
        ClientClaimsPeriodFilterOption selectedPeriod,
        string selectedBenefit,
        IReadOnlyCollection<ClaimsExperienceRecord> claims,
        IReadOnlyCollection<MonthlyClaimsRecord> monthlyClaims)
    {
        var claimsRows = claims.Where(row => row.ClientId > 0).ToList();
        var monthlyRows = SelectAuthoritativeMonthlyRows(monthlyClaims.Where(row => row.ClientId > 0));
        var window = ResolveAverageSpendWindow(selectedPeriod, claimsRows, monthlyRows);
        if (!window.HasValue)
        {
            return ClientClaimsAverageSpendComparison.Empty(selectedPeriod.DisplayLabel);
        }

        var currentWindow = window.Value;
        var priorWindow = (Start: currentWindow.Start.AddYears(-1), End: currentWindow.End.AddYears(-1));
        var current = BuildAverageSpendSnapshot(claimsRows, monthlyRows, currentWindow.Start, currentWindow.End);
        var prior = BuildAverageSpendSnapshot(claimsRows, monthlyRows, priorWindow.Start, priorWindow.End);
        var currentCombined = CombineAverageSpend(current.Ehc, current.Dental);
        var priorCombined = CombineAverageSpend(prior.Ehc, prior.Dental);
        var measures = new[]
        {
            (Coverage: "EHC", Current: current.Ehc, Prior: prior.Ehc),
            (Coverage: "Dental", Current: current.Dental, Prior: prior.Dental)
        };
        var maxAverageSpend = measures
            .SelectMany(item => new[] { item.Current.AverageSpend, item.Prior.AverageSpend })
            .Where(value => value.HasValue)
            .Select(value => value!.Value)
            .DefaultIfEmpty(0m)
            .Max();
        var rows = measures
            .Select(item => CreateAverageSpendRow(item.Coverage, item.Current, item.Prior, maxAverageSpend))
            .ToList();

        var selectedMeasure = selectedBenefit switch
        {
            "ehc" => (Label: "EHC", Current: current.Ehc, Prior: prior.Ehc),
            "dental" => (Label: "Dental", Current: current.Dental, Prior: prior.Dental),
            _ when currentCombined.HasSource =>
                (Label: "Combined EHC + Dental", Current: currentCombined, Prior: priorCombined),
            _ when current.Ehc.HasSource && !current.Dental.HasSource =>
                (Label: "EHC only — Dental lives unavailable", Current: current.Ehc, Prior: prior.Ehc),
            _ when current.Dental.HasSource && !current.Ehc.HasSource =>
                (Label: "Dental only — EHC lives unavailable", Current: current.Dental, Prior: prior.Dental),
            _ => (Label: "Combined EHC + Dental", Current: currentCombined, Prior: priorCombined)
        };
        var selectedChange = CalculateChange(selectedMeasure.Current.AverageSpend, selectedMeasure.Prior.AverageSpend);
        var largestChange = rows
            .Where(row => row.ChangePercentValue.HasValue)
            .OrderByDescending(row => row.ChangePercentValue)
            .FirstOrDefault();

        return new ClientClaimsAverageSpendComparison
        {
            PeriodContextText = $"Selected reporting period: {FormatAverageSpendRange(currentWindow.Start, currentWindow.End)} | Paid claims",
            CurrentPeriodText = $"Current: {FormatAverageSpendRange(currentWindow.Start, currentWindow.End)}",
            PriorPeriodText = $"Prior: {FormatAverageSpendRange(priorWindow.Start, priorWindow.End)}",
            ScopeLabel = selectedMeasure.Label,
            CurrentValueText = FormatAverageSpendValue(selectedMeasure.Current.AverageSpend),
            PriorValueText = selectedMeasure.Prior.AverageSpend.HasValue
                ? $"Prior: {FormatAverageSpendValue(selectedMeasure.Prior.AverageSpend)} per insured life"
                : "Comparable prior period not available",
            ChangeText = FormatAverageSpendChange(selectedChange.Percent),
            ChangeAmountText = FormatAverageSpendDelta(selectedChange.Amount),
            ChangeBrush = ResolveAverageSpendChangeBrush(selectedChange.Amount),
            Rows = rows,
            LargestChangeText = largestChange is null
                ? "No comparable prior-period average is stored."
                : $"Largest increase: {largestChange.Coverage}",
            ComparisonBasisText = priorCombined.AverageSpend.HasValue || prior.Ehc.AverageSpend.HasValue || prior.Dental.AverageSpend.HasValue
                ? "Same-period comparison | Paid claims / average monthly insured lives"
                : "Prior-period comparison unavailable | Paid claims / average monthly insured lives"
        };
    }

    private static (DateOnly Start, DateOnly End)? ResolveAverageSpendWindow(
        ClientClaimsPeriodFilterOption selectedPeriod,
        IReadOnlyCollection<ClaimsExperienceRecord> claims,
        IReadOnlyCollection<MonthlyClaimsRecord> monthlyClaims)
    {
        if (selectedPeriod.Start.HasValue && selectedPeriod.End.HasValue)
        {
            if (selectedPeriod.Kind == ClientClaimsPeriodFilterKind.LatestMonthly)
            {
                var selectedDates = monthlyClaims
                    .Where(row => row.Month >= selectedPeriod.Start && row.Month <= selectedPeriod.End)
                    .Select(row => row.Month)
                    .Where(month => month.HasValue)
                    .Select(month => month!.Value)
                    .Distinct()
                    .OrderBy(month => month)
                    .ToList();
                if (selectedDates.Count > 0)
                {
                    return (selectedDates[0], EndOfMonth(selectedDates[^1]));
                }
            }

            return (selectedPeriod.Start.Value, selectedPeriod.End.Value);
        }

        var latestExactPeriod = claims
            .Where(row => row.ExperiencePeriodStart.HasValue && row.ExperiencePeriodEnd.HasValue)
            .Select(row => (Start: row.ExperiencePeriodStart!.Value, End: row.ExperiencePeriodEnd!.Value))
            .Distinct()
            .OrderByDescending(period => period.End)
            .ThenByDescending(period => period.Start)
            .FirstOrDefault();
        if (latestExactPeriod != default)
        {
            return latestExactPeriod;
        }

        var monthlyDates = monthlyClaims
            .Where(row => row.Month.HasValue)
            .Select(row => row.Month!.Value)
            .Distinct()
            .OrderBy(month => month)
            .ToList();
        if (monthlyDates.Count == 0)
        {
            return null;
        }

        var end = monthlyDates[^1];
        var rollingStart = end.AddMonths(-11);
        return (monthlyDates.First(month => month >= rollingStart), EndOfMonth(end));
    }

    private static AverageSpendSnapshot BuildAverageSpendSnapshot(
        IReadOnlyCollection<ClaimsExperienceRecord> claims,
        IReadOnlyCollection<MonthlyClaimsRecord> monthlyClaims,
        DateOnly start,
        DateOnly end)
    {
        return new AverageSpendSnapshot(
            BuildAverageSpendMeasure("Extended Health", claims, monthlyClaims, start, end),
            BuildAverageSpendMeasure("Dental", claims, monthlyClaims, start, end));
    }

    private static AverageSpendMeasure BuildAverageSpendMeasure(
        string coverage,
        IReadOnlyCollection<ClaimsExperienceRecord> claims,
        IReadOnlyCollection<MonthlyClaimsRecord> monthlyClaims,
        DateOnly start,
        DateOnly end)
    {
        var expectedMonths = CountAverageSpendMonths(start, end);
        var selectedMonthly = monthlyClaims
            .Where(row => row.Month >= start &&
                          row.Month <= end &&
                          NormalizeBenefitLine(row.BenefitLine).Equals(coverage, StringComparison.OrdinalIgnoreCase))
            .ToList();
        var savedMonths = selectedMonthly
            .Where(row => row.Month.HasValue)
            .Select(row => row.Month!.Value)
            .Distinct()
            .Count();
        var validLives = selectedMonthly
            .Where(row => row.Lives > 0)
            .Select(row => (decimal)row.Lives!.Value)
            .ToList();
        if (savedMonths >= expectedMonths && validLives.Count >= expectedMonths)
        {
            return new AverageSpendMeasure(
                selectedMonthly.Sum(row => row.PaidClaims ?? 0m),
                validLives.Average(),
                true);
        }

        var selectedClaims = claims
            .Where(row => row.ExperiencePeriodStart == start &&
                          row.ExperiencePeriodEnd == end &&
                          NormalizeBenefitLine(row.BenefitLine).Equals(coverage, StringComparison.OrdinalIgnoreCase))
            .GroupBy(row => NormalizeBenefitLine(row.BenefitLine))
            .Select(group => group.OrderByDescending(row => row.DateLoaded).ThenByDescending(row => row.ClaimsId).First())
            .ToList();
        var aggregateLives = selectedClaims
            .Where(row => row.Lives > 0)
            .Select(row => (decimal)row.Lives!.Value)
            .DefaultIfEmpty(0m)
            .Max();
        return selectedClaims.Count > 0 && aggregateLives > 0m
            ? new AverageSpendMeasure(
                selectedClaims.Sum(row => row.PaidClaims ?? 0m),
                aggregateLives,
                true)
            : AverageSpendMeasure.Empty;
    }

    private static AverageSpendMeasure CombineAverageSpend(AverageSpendMeasure ehc, AverageSpendMeasure dental)
    {
        var averageLives = Math.Max(ehc.AverageLives, dental.AverageLives);
        return ehc.HasSource && dental.HasSource && averageLives > 0m
            ? new AverageSpendMeasure(ehc.PaidClaims + dental.PaidClaims, averageLives, true)
            : AverageSpendMeasure.Empty;
    }

    private static ClientClaimsAverageSpendRow CreateAverageSpendRow(
        string coverage,
        AverageSpendMeasure current,
        AverageSpendMeasure prior,
        decimal maxAverageSpend)
    {
        var change = CalculateChange(current.AverageSpend, prior.AverageSpend);
        return new ClientClaimsAverageSpendRow
        {
            Coverage = coverage,
            CurrentValueText = FormatAverageSpendValue(current.AverageSpend),
            PriorValueText = FormatAverageSpendValue(prior.AverageSpend),
            ChangeText = FormatAverageSpendChange(change.Percent),
            ChangeAmountText = FormatAverageSpendDelta(change.Amount),
            ChangeBrush = ResolveAverageSpendChangeBrush(change.Amount),
            CurrentBarWidth = CalculateAverageSpendBarWidth(current.AverageSpend, maxAverageSpend),
            PriorBarWidth = CalculateAverageSpendBarWidth(prior.AverageSpend, maxAverageSpend),
            ChangePercentValue = change.Percent
        };
    }

    private static (decimal? Percent, decimal? Amount) CalculateChange(decimal? current, decimal? prior)
    {
        if (!current.HasValue || !prior.HasValue || prior.Value == 0m)
        {
            return (null, null);
        }

        return ((current.Value / prior.Value - 1m) * 100m, current.Value - prior.Value);
    }

    private static string FormatAverageSpendValue(decimal? value)
        => value.HasValue ? FormatAverageMoney(value.Value) : "N/A";

    private static string FormatAverageSpendChange(decimal? value)
        => !value.HasValue ? "N/A" : value.Value >= 0m
            ? $"Up {value.Value:0.0}%"
            : $"Down {Math.Abs(value.Value):0.0}%";

    private static string FormatAverageSpendDelta(decimal? value)
        => !value.HasValue ? "" : value.Value >= 0m
            ? $"+{FormatAverageMoney(value.Value)} per insured life"
            : $"-{FormatAverageMoney(Math.Abs(value.Value))} per insured life";

    private static Brush ResolveAverageSpendChangeBrush(decimal? value)
        => !value.HasValue ? Muted : value.Value > 0m ? Orange : value.Value < 0m ? Green : Muted;

    private static double CalculateAverageSpendBarWidth(decimal? value, decimal max)
        => !value.HasValue || value.Value <= 0m || max <= 0m
            ? 0d
            : (double)Math.Max(8m, value.Value / max * 300m);

    private static int CountAverageSpendMonths(DateOnly start, DateOnly end)
        => Math.Max(0, ((end.Year - start.Year) * 12) + end.Month - start.Month + 1);

    private static DateOnly EndOfMonth(DateOnly month)
        => new(month.Year, month.Month, DateTime.DaysInMonth(month.Year, month.Month));

    private static string FormatAverageSpendRange(DateOnly start, DateOnly end)
        => $"{start.ToString("MMM yyyy", CultureInfo.InvariantCulture)} - {end.ToString("MMM yyyy", CultureInfo.InvariantCulture)}";

    private sealed record AverageSpendSnapshot(AverageSpendMeasure Ehc, AverageSpendMeasure Dental);

    private sealed record AverageSpendMeasure(decimal PaidClaims, decimal AverageLives, bool HasSource)
    {
        public static AverageSpendMeasure Empty { get; } = new(0m, 0m, false);
        public decimal? AverageSpend => HasSource && AverageLives > 0m ? PaidClaims / AverageLives : null;
    }

    internal static List<ClaimsExperienceRecord> SelectAuthoritativeClaimsRows(IEnumerable<ClaimsExperienceRecord> source)
    {
        var latestExactPeriods = source
            .Where(row => row.ExperiencePeriodStart.HasValue && row.ExperiencePeriodEnd.HasValue)
            .GroupBy(row => new
            {
                Benefit = NormalizeBenefitLine(row.BenefitLine),
                Start = row.ExperiencePeriodStart!.Value,
                End = row.ExperiencePeriodEnd!.Value
            })
            .Select(group => group.OrderByDescending(row => row.DateLoaded).ThenByDescending(row => row.ClaimsId).First())
            .ToList();

        return latestExactPeriods
            .Where(row => !latestExactPeriods.Any(other =>
                !ReferenceEquals(other, row) &&
                NormalizeBenefitLine(other.BenefitLine).Equals(NormalizeBenefitLine(row.BenefitLine), StringComparison.OrdinalIgnoreCase) &&
                other.ExperiencePeriodStart <= row.ExperiencePeriodStart &&
                other.ExperiencePeriodEnd >= row.ExperiencePeriodEnd &&
                (other.ExperiencePeriodStart < row.ExperiencePeriodStart || other.ExperiencePeriodEnd > row.ExperiencePeriodEnd)))
            .ToList();
    }

    internal static List<MonthlyClaimsRecord> SelectAuthoritativeMonthlyRows(IEnumerable<MonthlyClaimsRecord> source)
    {
        var rows = source.ToList();
        var datedRows = rows
            .Where(row => row.Month.HasValue)
            .GroupBy(row => new
            {
                Benefit = NormalizeBenefitLine(row.BenefitLine),
                Month = row.Month!.Value
            })
            .Select(group => group
                .OrderByDescending(row => row.DateLoaded)
                .ThenByDescending(row => row.ClaimMonthId)
                .First());

        return datedRows
            .Concat(rows.Where(row => !row.Month.HasValue))
            .ToList();
    }

    private static List<ClaimsExperienceRecord> ExcludeAggregateRowsReplacedByMonthlyClaims(
        IEnumerable<ClaimsExperienceRecord> claims,
        IReadOnlyCollection<MonthlyClaimsRecord> monthlyClaims)
    {
        var monthlyBenefits = monthlyClaims
            .Select(row => NormalizeBenefitLine(row.BenefitLine))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        return claims
            .Where(row => !monthlyBenefits.Contains(NormalizeBenefitLine(row.BenefitLine)))
            .ToList();
    }

    private static decimal ResolveLatestLives(
        ClientWorkspaceViewModel workspace,
        IReadOnlyCollection<ClaimsExperienceRecord> claims,
        IReadOnlyCollection<MonthlyClaimsRecord> monthlyClaims)
    {
        var latestMonthlyLivesDate = monthlyClaims
            .Where(row => row.Month.HasValue && row.Lives > 0)
            .Select(row => row.Month!.Value)
            .DefaultIfEmpty(DateOnly.MinValue)
            .Max();
        if (latestMonthlyLivesDate != DateOnly.MinValue)
        {
            return monthlyClaims
                .Where(row => row.Month == latestMonthlyLivesDate && row.Lives > 0)
                .Select(row => row.Lives!.Value)
                .Max();
        }

        var latestClaimsLivesDate = claims
            .Where(row => row.ExperiencePeriodEnd.HasValue && row.Lives > 0)
            .Select(row => row.ExperiencePeriodEnd!.Value)
            .DefaultIfEmpty(DateOnly.MinValue)
            .Max();
        if (latestClaimsLivesDate != DateOnly.MinValue)
        {
            return claims
                .Where(row => row.ExperiencePeriodEnd == latestClaimsLivesDate && row.Lives > 0)
                .Select(row => row.Lives!.Value)
                .Max();
        }

        return Math.Max(workspace.Employees, 0);
    }

    private static decimal ResolveAverageInsuredLives(
        ClientWorkspaceViewModel workspace,
        IReadOnlyCollection<ClaimsExperienceRecord> claims,
        IReadOnlyCollection<MonthlyClaimsRecord> monthlyClaims)
    {
        var monthlyLives = monthlyClaims
            .Where(row => row.Month.HasValue && row.Lives > 0)
            .GroupBy(row => row.Month!.Value)
            .Select(group => group.Max(row => (decimal)row.Lives!.Value))
            .ToList();
        if (monthlyLives.Count > 0)
        {
            return monthlyLives.Average();
        }

        var periodLives = claims
            .Where(row => row.ExperiencePeriodEnd.HasValue && row.Lives > 0)
            .GroupBy(row => new { row.ExperiencePeriodStart, row.ExperiencePeriodEnd })
            .Select(group => group.Max(row => (decimal)row.Lives!.Value))
            .ToList();
        if (periodLives.Count > 0)
        {
            return periodLives.Average();
        }

        return Math.Max(workspace.Employees, 0);
    }

    private static decimal SumPaidClaims(
        IReadOnlyCollection<ClaimsExperienceRecord> claims,
        IReadOnlyCollection<MonthlyClaimsRecord> monthlyClaims)
    {
        return claims.Sum(PaidClaims) + monthlyClaims.Sum(PaidClaims);
    }

    private static decimal SumPremium(
        IReadOnlyCollection<ClaimsExperienceRecord> claims,
        IReadOnlyCollection<MonthlyClaimsRecord> monthlyClaims,
        ClaimsExperienceRenewalTerms renewalTerms)
    {
        return claims.Sum(row => CalculatePremium(row.BenefitLine, row.Premium, renewalTerms)) +
               monthlyClaims.Sum(row => CalculatePremium(row.BenefitLine, row.Premium, renewalTerms));
    }

    private static bool HasCompletePremiumForPaidClaims(
        IReadOnlyCollection<ClaimsExperienceRecord> claims,
        IReadOnlyCollection<MonthlyClaimsRecord> monthlyClaims)
    {
        return claims.Where(row => PaidClaims(row) > 0m).All(row => row.Premium > 0m) &&
               monthlyClaims.Where(row => PaidClaims(row) > 0m).All(row => row.Premium > 0m);
    }

    private static bool HasCompleteLivesForPaidClaims(
        IReadOnlyCollection<ClaimsExperienceRecord> claims,
        IReadOnlyCollection<MonthlyClaimsRecord> monthlyClaims)
    {
        return claims.Where(row => PaidClaims(row) > 0m).All(row => row.Lives > 0) &&
               monthlyClaims.Where(row => PaidClaims(row) > 0m).All(row => row.Lives > 0);
    }

    private static decimal PaidClaims(ClaimsExperienceRecord row) =>
        row.PaidClaims ?? row.IncurredClaims ?? 0m;

    private static decimal PaidClaims(MonthlyClaimsRecord row) =>
        row.PaidClaims ?? row.IncurredClaims ?? row.ClaimsPlusFees ?? 0m;

    private static bool IsAsoDashboardRow(ClaimsExperienceRecord row) =>
        row.BenefitLine.Contains("ASO", StringComparison.OrdinalIgnoreCase);

    private static bool IsAsoDashboardRow(MonthlyClaimsRecord row) =>
        row.BenefitLine.Contains("ASO", StringComparison.OrdinalIgnoreCase);

    private static IReadOnlyList<ClientClaimsCategoryRow> BuildCategoryRows(
        IReadOnlyCollection<ClaimsExperienceRecord> claims,
        IReadOnlyCollection<MonthlyClaimsRecord> monthlyClaims,
        decimal totalPaid)
    {
        var rows = claims
            .Select(row => new
            {
                Category = NormalizeBenefitLine(row.BenefitLine),
                Paid = PaidClaims(row)
            })
            .Concat(monthlyClaims.Select(row => new
            {
                Category = NormalizeBenefitLine(row.BenefitLine),
                Paid = PaidClaims(row)
            }))
            .GroupBy(row => row.Category, StringComparer.OrdinalIgnoreCase)
            .Select(group => new
            {
                Category = group.Key,
                Paid = group.Sum(row => row.Paid)
            });

        return rows
            .OrderByDescending(row => row.Paid)
            .Take(8)
            .Select(row => new ClientClaimsCategoryRow
            {
                Category = row.Category,
                PaidClaimsText = FormatMoney(row.Paid),
                PercentText = totalPaid > 0m ? FormatPercent(row.Paid / totalPaid * 100m) : "N/A"
            })
            .ToList();
    }

    private static IReadOnlyList<ClientClaimsBarRow> BuildMonthlyTrendRows(
        IReadOnlyCollection<ClaimsExperienceRecord> claims,
        IReadOnlyCollection<MonthlyClaimsRecord> monthlyClaims)
    {
        var points = monthlyClaims.Count > 0
            ? monthlyClaims
                .Where(row => row.Month.HasValue)
                .GroupBy(row => row.Month!.Value)
                .Select(group => (Date: group.Key, Value: group.Sum(row => row.PaidClaims ?? row.IncurredClaims ?? row.ClaimsPlusFees ?? 0m)))
            : claims
                .Where(row => row.ExperiencePeriodEnd.HasValue)
                .GroupBy(row => row.ExperiencePeriodEnd!.Value)
                .Select(group => (Date: group.Key, Value: group.Sum(row => row.PaidClaims ?? row.IncurredClaims ?? 0m)));

        var ordered = points.OrderBy(point => point.Date).TakeLast(12).ToList();
        var max = ordered.Select(point => point.Value).DefaultIfEmpty(0m).Max();
        return ordered.Select(point => new ClientClaimsBarRow
        {
            Label = point.Date.ToString("MMM yy", CultureInfo.InvariantCulture),
            ValueText = FormatMoney(point.Value),
            BarWidth = CalculateBarWidth(point.Value, max),
            BarBrush = Purple
        }).ToList();
    }

    private static IReadOnlyList<ClientClaimsBarRow> BuildLossRatioTrendRows(
        IReadOnlyCollection<ClaimsExperienceRecord> claims,
        IReadOnlyCollection<MonthlyClaimsRecord> monthlyClaims,
        ClaimsExperienceRenewalTerms renewalTerms)
    {
        var points = monthlyClaims.Count > 0
            ? monthlyClaims
                .Where(row => row.Month.HasValue)
                .GroupBy(row => row.Month!.Value)
                .Select(group =>
                {
                    var paid = group.Sum(row => row.PaidClaims ?? row.IncurredClaims ?? row.ClaimsPlusFees ?? 0m);
                    var premium = group.Sum(row => CalculatePremium(row.BenefitLine, row.Premium, renewalTerms));
                    var target = ResolveTargetLossRatioPercent(
                        Array.Empty<ClaimsExperienceRecord>(), group.ToList(), renewalTerms);
                    return (Date: group.Key, Ratio: premium > 0m ? paid / premium * 100m : (decimal?)null, Target: target);
                })
            : claims
                .Where(row => row.ExperiencePeriodEnd.HasValue)
                .GroupBy(row => row.ExperiencePeriodEnd!.Value)
                .Select(group =>
                {
                    var paid = group.Sum(row => row.PaidClaims ?? row.IncurredClaims ?? 0m);
                    var premium = group.Sum(row => CalculatePremium(row.BenefitLine, row.Premium, renewalTerms));
                    var target = ResolveTargetLossRatioPercent(
                        group.ToList(), Array.Empty<MonthlyClaimsRecord>(), renewalTerms);
                    return (Date: group.Key, Ratio: premium > 0m ? paid / premium * 100m : group.Average(row => row.LossRatio), Target: target);
                });

        return points
            .Where(point => point.Ratio.HasValue)
            .OrderBy(point => point.Date)
            .TakeLast(8)
            .Select(point => new ClientClaimsBarRow
            {
                Label = point.Date.ToString("MMM yy", CultureInfo.InvariantCulture),
                ValueText = FormatPercent(point.Ratio!.Value),
                BarWidth = CalculateBarWidth(point.Ratio.Value, 150m),
                BarBrush = point.Ratio.Value >= 100m ? Red : point.Ratio.Value > point.Target ? Orange : Green
            })
            .ToList();
    }

    private static IReadOnlyList<ClientClaimsMetricRow> BuildUtilizationRows(
        IReadOnlyCollection<ClaimsExperienceRecord> claims,
        IReadOnlyCollection<MonthlyClaimsRecord> monthlyClaims,
        decimal premium,
        decimal paidClaims,
        decimal? lossRatio,
        DateTime latestLoaded)
    {
        return
        [
            new ClientClaimsMetricRow("Saved claims records", (claims.Count + monthlyClaims.Count).ToString("#,0", CultureInfo.InvariantCulture)),
            new ClientClaimsMetricRow("Benefit categories", claims.Select(row => NormalizeBenefitLine(row.BenefitLine)).Concat(monthlyClaims.Select(row => NormalizeBenefitLine(row.BenefitLine))).Distinct(StringComparer.OrdinalIgnoreCase).Count().ToString("#,0", CultureInfo.InvariantCulture)),
            new ClientClaimsMetricRow("Premium captured", FormatMoney(premium)),
            new ClientClaimsMetricRow("Paid claims captured", FormatMoney(paidClaims)),
            new ClientClaimsMetricRow("Calculated loss ratio", lossRatio.HasValue ? FormatPercent(lossRatio.Value) : "N/A"),
            new ClientClaimsMetricRow("Latest data load", latestLoaded > DateTime.MinValue ? latestLoaded.ToString("MMM d, yyyy", CultureInfo.InvariantCulture) : "N/A")
        ];
    }

    private static IReadOnlyList<string> BuildSummaryBullets(
        decimal paidClaims,
        decimal? lossRatio,
        ClientClaimsCategoryRow? topCategory,
        int monthlyRows,
        int claimsRows)
    {
        if (paidClaims <= 0m && claimsRows == 0 && monthlyRows == 0)
        {
            return ["No claims experience has been saved for this client yet."];
        }

        var bullets = new List<string>
        {
            $"Paid claims captured: {FormatMoney(paidClaims)}.",
            lossRatio.HasValue
                ? $"Calculated paid loss ratio is {FormatPercent(lossRatio.Value)}."
                : "Premium is not fully captured, so loss ratio is unavailable.",
            topCategory is not null
                ? $"{topCategory.Category} is the largest saved claims category."
                : "Benefit category mix will appear once categories are saved.",
            monthlyRows > 0
                ? $"{monthlyRows:#,0} monthly claims row(s) are available for trend analysis."
                : $"{claimsRows:#,0} claims experience row(s) are available."
        };

        return bullets;
    }

    private static IReadOnlyList<ClientClaimsInsightRow> BuildInsights(
        ClientWorkspaceViewModel workspace,
        decimal paidClaims,
        decimal? lossRatio,
        decimal targetLossRatio,
        IReadOnlyList<ClientClaimsCategoryRow> categoryRows)
    {
        if (paidClaims <= 0m)
        {
            return
            [
                ClientClaimsInsightRow.Info("Ready for data", "Process a claims report for this client and save it to SQLite to populate this dashboard.")
            ];
        }

        var insights = new List<ClientClaimsInsightRow>();
        if (lossRatio >= 100m)
        {
            insights.Add(ClientClaimsInsightRow.Critical("High loss ratio", $"Claims are running above premium for {workspace.ClientName}."));
        }
        else if (lossRatio > targetLossRatio)
        {
            insights.Add(ClientClaimsInsightRow.Warning("Monitor closely", $"Loss ratio is above the {FormatPercent(targetLossRatio)} renewal target."));
        }
        else if (lossRatio.HasValue)
        {
            insights.Add(ClientClaimsInsightRow.Positive("Positive trend", $"Claims are currently at or below the {FormatPercent(targetLossRatio)} renewal target."));
        }

        var top = categoryRows.FirstOrDefault();
        if (top is not null)
        {
            insights.Add(ClientClaimsInsightRow.Info("Largest category", $"{top.Category} is the highest paid saved category at {top.PaidClaimsText}."));
        }

        return insights;
    }

    private static string ResolvePeriod(
        IReadOnlyCollection<ClaimsExperienceRecord> claims,
        IReadOnlyCollection<MonthlyClaimsRecord> monthlyClaims)
    {
        var starts = claims.Select(row => row.ExperiencePeriodStart)
            .Concat(monthlyClaims.Select(row => row.Month))
            .Where(date => date.HasValue)
            .Select(date => date!.Value)
            .ToList();
        var ends = claims.Select(row => row.ExperiencePeriodEnd)
            .Concat(monthlyClaims.Select(row => row.Month.HasValue
                ? new DateOnly(
                    row.Month.Value.Year,
                    row.Month.Value.Month,
                    DateTime.DaysInMonth(row.Month.Value.Year, row.Month.Value.Month))
                : (DateOnly?)null))
            .Where(date => date.HasValue)
            .Select(date => date!.Value)
            .ToList();

        if (starts.Count == 0 || ends.Count == 0)
        {
            return "All saved claims periods";
        }

        return $"{starts.Min():MMM d, yyyy} - {ends.Max():MMM d, yyyy}";
    }

    private static string NormalizeBenefitLine(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return "Other";
        }

        var normalized = value.Trim();
        if (normalized.Contains("drug", StringComparison.OrdinalIgnoreCase))
        {
            return "Drugs";
        }

        if (normalized.Contains("paramed", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("massage", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("chiro", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("physio", StringComparison.OrdinalIgnoreCase))
        {
            return "Paramedical";
        }

        if (normalized.Contains("dental", StringComparison.OrdinalIgnoreCase))
        {
            return "Dental";
        }

        if (normalized.Contains("vision", StringComparison.OrdinalIgnoreCase))
        {
            return "Vision";
        }

        if (normalized.Contains("hospital", StringComparison.OrdinalIgnoreCase))
        {
            return "Hospital";
        }

        if (normalized.Contains("health", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("ehc", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("extended", StringComparison.OrdinalIgnoreCase))
        {
            return "Extended Health";
        }

        return normalized;
    }

    private static double CalculateBarWidth(decimal value, decimal max)
    {
        if (max <= 0m || value <= 0m)
        {
            return 0d;
        }

        var width = (double)Math.Min(220m, Math.Max(8m, value / max * 220m));
        return width;
    }

    private static Brush CreateFrozenBrush(byte red, byte green, byte blue)
    {
        var brush = new SolidColorBrush(Color.FromRgb(red, green, blue));
        brush.Freeze();
        return brush;
    }

    private static string FormatMoney(decimal value)
    {
        return value >= 1000000m
            ? value.ToString("$0.##M", CultureInfo.InvariantCulture)
            : value.ToString("$#,0", CultureInfo.InvariantCulture);
    }

    private static string FormatAverageMoney(decimal value)
    {
        return value.ToString("$#,0.00", CultureInfo.InvariantCulture);
    }

    private static string FormatLives(decimal value)
    {
        return value == decimal.Truncate(value)
            ? value.ToString("#,0", CultureInfo.InvariantCulture)
            : value.ToString("#,0.0", CultureInfo.InvariantCulture);
    }

    private static string FormatPercent(decimal value)
    {
        return value.ToString("0.0'%' ", CultureInfo.InvariantCulture).Trim();
    }

    private static decimal CalculatePremium(
        string benefitLine,
        decimal? billedPremium,
        ClaimsExperienceRenewalTerms renewalTerms)
    {
        var premium = billedPremium ?? 0m;
        return IsEhcBenefit(benefitLine)
            ? premium * (1m - renewalTerms.EffectiveEhcPoolingCharge)
            : premium;
    }

    private static decimal ResolveTargetLossRatioPercent(
        IReadOnlyCollection<ClaimsExperienceRecord> claims,
        IReadOnlyCollection<MonthlyClaimsRecord> monthlyClaims,
        ClaimsExperienceRenewalTerms renewalTerms)
    {
        decimal ehcPremium;
        decimal dentalPremium;
        bool hasEhc;
        bool hasDental;

        if (monthlyClaims.Count > 0)
        {
            hasEhc = monthlyClaims.Any(row => IsEhcBenefit(row.BenefitLine));
            hasDental = monthlyClaims.Any(row => IsDentalBenefit(row.BenefitLine));
            ehcPremium = monthlyClaims
                .Where(row => IsEhcBenefit(row.BenefitLine))
                .Sum(row => CalculatePremium(row.BenefitLine, row.Premium, renewalTerms));
            dentalPremium = monthlyClaims
                .Where(row => IsDentalBenefit(row.BenefitLine))
                .Sum(row => row.Premium ?? 0m);
        }
        else
        {
            hasEhc = claims.Any(row => IsEhcBenefit(row.BenefitLine));
            hasDental = claims.Any(row => IsDentalBenefit(row.BenefitLine));
            ehcPremium = claims
                .Where(row => IsEhcBenefit(row.BenefitLine))
                .Sum(row => CalculatePremium(row.BenefitLine, row.Premium, renewalTerms));
            dentalPremium = claims
                .Where(row => IsDentalBenefit(row.BenefitLine))
                .Sum(row => row.Premium ?? 0m);
        }

        var totalPremium = ehcPremium + dentalPremium;
        if (totalPremium > 0m)
        {
            return ((ehcPremium * renewalTerms.EffectiveEhcTargetLossRatio) +
                    (dentalPremium * renewalTerms.EffectiveDentalTargetLossRatio)) /
                   totalPremium * 100m;
        }

        if (hasDental && !hasEhc)
        {
            return renewalTerms.EffectiveDentalTargetLossRatio * 100m;
        }

        return renewalTerms.EffectiveEhcTargetLossRatio * 100m;
    }

    private static bool IsEhcBenefit(string value)
    {
        var normalized = NormalizeBenefitLine(value);
        return normalized.Equals("Extended Health", StringComparison.OrdinalIgnoreCase) ||
               normalized.Equals("Drugs", StringComparison.OrdinalIgnoreCase) ||
               normalized.Equals("Paramedical", StringComparison.OrdinalIgnoreCase) ||
               normalized.Equals("Vision", StringComparison.OrdinalIgnoreCase) ||
               normalized.Equals("Hospital", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsDentalBenefit(string value)
    {
        return NormalizeBenefitLine(value).Equals("Dental", StringComparison.OrdinalIgnoreCase);
    }

    private static string RiskLabel(decimal value, decimal targetLossRatio)
    {
        if (value >= 100m)
        {
            return "Above premium";
        }

        return value > targetLossRatio
            ? $"Above {FormatPercent(targetLossRatio)} target"
            : $"At or below {FormatPercent(targetLossRatio)} target";
    }

    private static string DisplayOrDefault(string value)
    {
        return string.IsNullOrWhiteSpace(value) ? "Not set" : value;
    }
}

public sealed class ClientClaimsAverageSpendComparison
{
    public string PeriodContextText { get; init; } = "Selected reporting period | Paid claims";
    public string CurrentPeriodText { get; init; } = "Current period";
    public string PriorPeriodText { get; init; } = "Prior period";
    public string ScopeLabel { get; init; } = "Combined EHC + Dental";
    public string CurrentValueText { get; init; } = "N/A";
    public string PriorValueText { get; init; } = "Comparable prior period not available";
    public string ChangeText { get; init; } = "N/A";
    public string ChangeAmountText { get; init; } = "";
    public Brush ChangeBrush { get; init; } = Brushes.Gray;
    public IReadOnlyList<ClientClaimsAverageSpendRow> Rows { get; init; } = [];
    public string LargestChangeText { get; init; } = "No comparable prior-period average is stored.";
    public string ComparisonBasisText { get; init; } = "Paid claims / average monthly insured lives";

    public static ClientClaimsAverageSpendComparison Empty(string period = "Selected reporting period")
    {
        return new ClientClaimsAverageSpendComparison
        {
            PeriodContextText = $"{period} | Paid claims"
        };
    }
}

public sealed class ClientClaimsAverageSpendRow
{
    public string Coverage { get; init; } = "";
    public string CurrentValueText { get; init; } = "N/A";
    public string PriorValueText { get; init; } = "N/A";
    public string ChangeText { get; init; } = "N/A";
    public string ChangeAmountText { get; init; } = "";
    public Brush ChangeBrush { get; init; } = Brushes.Gray;
    public double CurrentBarWidth { get; init; }
    public double PriorBarWidth { get; init; }
    internal decimal? ChangePercentValue { get; init; }
}
public sealed class ClientClaimsCategoryRow
{
    public string Category { get; init; } = "";
    public string PaidClaimsText { get; init; } = "";
    public string PercentText { get; init; } = "";
}

public sealed class ClientClaimsBarRow
{
    public string Label { get; init; } = "";
    public string ValueText { get; init; } = "";
    public double BarWidth { get; init; }
    public Brush BarBrush { get; init; } = Brushes.Gray;
}

public sealed record ClientClaimsMetricRow(string Metric, string Value);

public sealed class ClientClaimsInsightRow
{
    public string Title { get; init; } = "";
    public string Detail { get; init; } = "";
    public Brush Background { get; init; } = Brushes.White;
    public Brush BorderBrush { get; init; } = Brushes.LightGray;
    public Brush Foreground { get; init; } = Brushes.Black;

    public static ClientClaimsInsightRow Positive(string title, string detail)
    {
        return Create(title, detail, "#ECFDF5", "#A7F3D0", "#047857");
    }

    public static ClientClaimsInsightRow Warning(string title, string detail)
    {
        return Create(title, detail, "#FFF7ED", "#FED7AA", "#C2410C");
    }

    public static ClientClaimsInsightRow Critical(string title, string detail)
    {
        return Create(title, detail, "#FEF2F2", "#FECACA", "#B91C1C");
    }

    public static ClientClaimsInsightRow Info(string title, string detail)
    {
        return Create(title, detail, "#EEF2FF", "#C7D2FE", "#4338CA");
    }

    private static ClientClaimsInsightRow Create(string title, string detail, string background, string border, string foreground)
    {
        return new ClientClaimsInsightRow
        {
            Title = title,
            Detail = detail,
            Background = (Brush)new BrushConverter().ConvertFromString(background)!,
            BorderBrush = (Brush)new BrushConverter().ConvertFromString(border)!,
            Foreground = (Brush)new BrushConverter().ConvertFromString(foreground)!
        };
    }
}

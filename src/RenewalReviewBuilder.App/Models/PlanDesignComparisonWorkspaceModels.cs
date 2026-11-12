using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Media;

namespace RenewalReviewBuilder.App.Models;

public static class PlanDesignComparisonStatuses
{
    public const string Improved = "Improved";
    public const string Reduced = "Reduced";
    public const string Changed = "Changed";
    public const string Same = "Same";
    public const string Added = "Added";
    public const string Removed = "Removed";
    public const string ReviewRequired = "Review Required";
    public const string NotComparable = "Not Comparable";
}

public static class PlanDesignComparisonViewModes
{
    public const string AllBenefits = "All Benefits";
    public const string ChangedOnly = "Changed Only";
    public const string KeyBenefitsOnly = "Key Benefits Only";
}

public sealed class PlanDesignComparisonOption
{
    public string Key { get; init; } = "";
    public string DisplayName { get; init; } = "";
    public string Detail { get; init; } = "";
    public bool IsBaseline { get; init; }
}

public sealed class PlanDesignComparisonMetric
{
    public string Label { get; init; } = "";
    public string Value { get; init; } = "";
    public string Detail { get; init; } = "";
    public Brush ForegroundBrush { get; init; } = Brushes.Gray;
}

public sealed class PlanDesignComparisonEvidence
{
    public string ReferenceText { get; init; } = "Not available";
    public string SourceText { get; init; } = "";
    public string SourceDocument { get; init; } = "";
    public int? SourcePage { get; init; }
    public bool IsAvailable { get; init; }
}

public sealed class PlanDesignComparisonProvisionRow : INotifyPropertyChanged
{
    private bool _isDetailsExpanded;

    public event PropertyChangedEventHandler? PropertyChanged;

    public string Category { get; init; } = "Other";
    public string ScenarioKey { get; init; } = "";
    public string BenefitName { get; init; } = "";
    public string ProvisionName { get; init; } = "";
    public string CurrentValue { get; init; } = "Not available";
    public string SelectedValue { get; init; } = "Not available";
    public string DifferenceStatus { get; init; } = PlanDesignComparisonStatuses.NotComparable;
    public string DifferenceRationale { get; init; } = "";
    public string CostImpactText { get; init; } = "Not available";
    public string ValueRecommendation { get; init; } = "Not available";
    public string ValidationWarning { get; init; } = "";
    public bool IsKeyBenefit { get; init; }
    public PlanDesignComparisonEvidence CurrentEvidence { get; init; } = new();
    public PlanDesignComparisonEvidence SelectedEvidence { get; init; } = new();

    public bool IsDetailsExpanded
    {
        get => _isDetailsExpanded;
        set
        {
            if (_isDetailsExpanded == value)
            {
                return;
            }

            _isDetailsExpanded = value;
            OnPropertyChanged();
        }
    }

    public bool HasEvidence => CurrentEvidence.IsAvailable || SelectedEvidence.IsAvailable;
    public string EvidenceText => SelectedEvidence.IsAvailable
        ? SelectedEvidence.ReferenceText
        : CurrentEvidence.ReferenceText;
    public string SourceWording => FirstNonBlank(SelectedEvidence.SourceText, CurrentEvidence.SourceText, "Source wording is not available.");
    public string ValidationWarningText => FirstNonBlank(ValidationWarning, "No validation warning recorded.");
    public Brush StatusForegroundBrush => PlanDesignComparisonBrushes.Foreground(DifferenceStatus);
    public Brush StatusBackgroundBrush => PlanDesignComparisonBrushes.Background(DifferenceStatus);

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

    private static string FirstNonBlank(params string?[] values) =>
        values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value))?.Trim() ?? "";
}

public sealed class PlanDesignComparisonGroup
{
    public string Category { get; init; } = "Other";
    public int Order { get; init; }
    public bool IsExpanded { get; set; } = true;
    public ObservableCollection<PlanDesignComparisonProvisionRow> Rows { get; } = [];
    public int ChangedCount => Rows.Count(row => !row.DifferenceStatus.Equals(PlanDesignComparisonStatuses.Same, StringComparison.OrdinalIgnoreCase));
    public string CostImpactText => "Cost data not available";
}

public sealed class PlanDesignComparisonDisplayRow
{
    public PlanDesignComparisonGroup? Group { get; init; }
    public PlanDesignComparisonProvisionRow? Provision { get; init; }
    public bool IsGroupHeader => Group is not null;
    public bool IsProvision => Provision is not null;
    public string GroupName => Group?.Category ?? "";
    public int ChangedCount { get; init; }
    public string GroupChangedText => Group is null ? "" : $"{ChangedCount} changed";
    public string GroupCostText { get; init; } = "";
    public bool GroupIsExpanded => Group?.IsExpanded ?? false;
}

public sealed class PlanDesignComparisonWorkspaceViewModel : INotifyPropertyChanged
{
    private string _selectedBaselineKey = "current";
    private string _selectedScenarioKey = "";
    private string _viewMode = PlanDesignComparisonViewModes.AllBenefits;
    private string _searchText = "";
    private bool _highlightDifferences = true;

    public event PropertyChangedEventHandler? PropertyChanged;

    public ObservableCollection<PlanDesignComparisonOption> BaselineOptions { get; } = [];
    public ObservableCollection<PlanDesignComparisonOption> ScenarioOptions { get; } = [];
    public ObservableCollection<PlanDesignComparisonGroup> Groups { get; } = [];
    public ObservableCollection<PlanDesignComparisonDisplayRow> DisplayRows { get; } = [];
    public ObservableCollection<PlanDesignComparisonMetric> SummaryMetrics { get; } = [];

    public string SelectedBaselineKey
    {
        get => _selectedBaselineKey;
        set => SetField(ref _selectedBaselineKey, value);
    }

    public string SelectedScenarioKey
    {
        get => _selectedScenarioKey;
        set
        {
            if (SetField(ref _selectedScenarioKey, value))
            {
                OnPropertyChanged(nameof(SelectedScenarioName));
            }
                RefreshPresentation();
        }
    }

    public string ViewMode
    {
        get => _viewMode;
        set
        {
            if (SetField(ref _viewMode, value))
            {
                RebuildDisplayRows();
            }
        }
    }

    public string SearchText
    {
        get => _searchText;
        set
        {
            if (SetField(ref _searchText, value))
            {
                RebuildDisplayRows();
            }
        }
    }

    public bool HighlightDifferences
    {
        get => _highlightDifferences;
        set => SetField(ref _highlightDifferences, value);
    }

    public bool HasComparisonScenario => ScenarioOptions.Count > 0 && !string.IsNullOrWhiteSpace(SelectedScenarioKey);
    public string SelectedScenarioName => ScenarioOptions.FirstOrDefault(option => option.Key.Equals(SelectedScenarioKey, StringComparison.OrdinalIgnoreCase))?.DisplayName
        ?? "No comparison scenario selected";
    public bool HasDifferences => Groups.SelectMany(group => group.Rows)
        .Where(row => row.ScenarioKey.Equals(SelectedScenarioKey, StringComparison.OrdinalIgnoreCase))
        .Any(row => !row.DifferenceStatus.Equals(PlanDesignComparisonStatuses.Same, StringComparison.OrdinalIgnoreCase));
    public bool HasVisibleRows => DisplayRows.Any(row => row.IsProvision);
    public Visibility ComparisonContentVisibility => HasComparisonScenario && HasVisibleRows ? Visibility.Visible : Visibility.Collapsed;
    public Visibility NoScenarioVisibility => HasComparisonScenario ? Visibility.Collapsed : Visibility.Visible;
    public Visibility NoMatchesVisibility => HasComparisonScenario && !HasVisibleRows ? Visibility.Visible : Visibility.Collapsed;
    public string NoScenarioMessage => "Select a saved quote or scenario to compare it with the current plan.";
    public string NoMatchesMessage => HasDifferences && ViewMode.Equals(PlanDesignComparisonViewModes.ChangedOnly, StringComparison.OrdinalIgnoreCase)
        ? "The selected plan matches the current plan across all comparable provisions."
        : "No benefits match the selected filters.";

    public void ToggleGroup(PlanDesignComparisonGroup? group)
    {
        if (group is null)
        {
            return;
        }

        group.IsExpanded = !group.IsExpanded;
        RebuildDisplayRows();
    }

    public void SetAllGroupsExpanded(bool expanded)
    {
        foreach (var group in Groups)
        {
            group.IsExpanded = expanded;
        }

        RebuildDisplayRows();
    }

    public void RefreshPresentation()
    {
        OnPropertyChanged(nameof(HasComparisonScenario));
        OnPropertyChanged(nameof(HasDifferences));
        OnPropertyChanged(nameof(SelectedScenarioName));
        RebuildSummaryMetrics();
        RebuildDisplayRows();
    }

    public void RebuildDisplayRows()
    {
        DisplayRows.Clear();
        foreach (var group in Groups.OrderBy(group => group.Order).ThenBy(group => group.Category, StringComparer.OrdinalIgnoreCase))
        {
            var scenarioRows = group.Rows.Where(row => row.ScenarioKey.Equals(SelectedScenarioKey, StringComparison.OrdinalIgnoreCase));
            var rows = scenarioRows.Where(MatchesCurrentFilter).ToList();
            if (rows.Count == 0)
            {
                continue;
            }

            DisplayRows.Add(new PlanDesignComparisonDisplayRow
            {
                Group = group,
                ChangedCount = scenarioRows.Count(row => !row.DifferenceStatus.Equals(PlanDesignComparisonStatuses.Same, StringComparison.OrdinalIgnoreCase)),
                GroupCostText = "Cost data not available"
            });
            if (group.IsExpanded)
            {
                foreach (var row in rows)
                {
                    DisplayRows.Add(new PlanDesignComparisonDisplayRow { Provision = row });
                }
            }
        }

        OnPropertyChanged(nameof(HasVisibleRows));
        OnPropertyChanged(nameof(ComparisonContentVisibility));
        OnPropertyChanged(nameof(NoScenarioVisibility));
        OnPropertyChanged(nameof(NoMatchesVisibility));
        OnPropertyChanged(nameof(NoMatchesMessage));
    }

    private void RebuildSummaryMetrics()
    {
        SummaryMetrics.Clear();
        var rows = Groups.SelectMany(group => group.Rows).Where(row => row.ScenarioKey.Equals(SelectedScenarioKey, StringComparison.OrdinalIgnoreCase)).ToList();
        AddMetric("Benefits Changed", rows.Count(row => !row.DifferenceStatus.Equals(PlanDesignComparisonStatuses.Same, StringComparison.OrdinalIgnoreCase)).ToString(), "Comparable provisions", Brushes.SlateGray);
        AddMetric("Coverage Improved", rows.Count(row => row.DifferenceStatus == PlanDesignComparisonStatuses.Improved).ToString(), "Domain-supported", Brushes.SeaGreen);
        AddMetric("Coverage Reduced", rows.Count(row => row.DifferenceStatus == PlanDesignComparisonStatuses.Reduced).ToString(), "Domain-supported", Brushes.Firebrick);
        AddMetric("Administrative Changes", rows.Count(row => row.Category == "Eligibility & Administration" && !row.DifferenceStatus.Equals(PlanDesignComparisonStatuses.Same, StringComparison.OrdinalIgnoreCase)).ToString(), "Review as applicable", Brushes.DarkOrange);
        AddMetric("Estimated Cost Impact", "Not available", "No mapped scenario cost", Brushes.SlateGray);
        AddMetric("Value Score Impact", "Not available", "No traceable score", Brushes.SlateGray);
        AddMetric("Compliance Score", rows.Any(row => row.DifferenceStatus == PlanDesignComparisonStatuses.ReviewRequired) ? "Review required" : "Not available", "No mapped compliance score", Brushes.SlateGray);

        void AddMetric(string label, string value, string detail, Brush foreground) => SummaryMetrics.Add(new PlanDesignComparisonMetric
        {
            Label = label,
            Value = value,
            Detail = detail,
            ForegroundBrush = foreground
        });
    }

    private bool MatchesCurrentFilter(PlanDesignComparisonProvisionRow row)
    {
        if (ViewMode.Equals(PlanDesignComparisonViewModes.ChangedOnly, StringComparison.OrdinalIgnoreCase) &&
            row.DifferenceStatus.Equals(PlanDesignComparisonStatuses.Same, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        if (ViewMode.Equals(PlanDesignComparisonViewModes.KeyBenefitsOnly, StringComparison.OrdinalIgnoreCase) && !row.IsKeyBenefit)
        {
            return false;
        }

        if (string.IsNullOrWhiteSpace(SearchText))
        {
            return true;
        }

        return $"{row.Category} {row.BenefitName} {row.ProvisionName} {row.CurrentValue} {row.SelectedValue} {row.DifferenceStatus}"
            .Contains(SearchText.Trim(), StringComparison.OrdinalIgnoreCase);
    }

    private bool SetField<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value))
        {
            return false;
        }

        field = value;
        OnPropertyChanged(propertyName);
        return true;
    }

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
}

public static class PlanDesignComparisonBrushes
{
    public static Brush Background(string status) => status switch
    {
        PlanDesignComparisonStatuses.Improved or PlanDesignComparisonStatuses.Added => new SolidColorBrush(Color.FromRgb(236, 253, 245)),
        PlanDesignComparisonStatuses.Reduced or PlanDesignComparisonStatuses.Removed => new SolidColorBrush(Color.FromRgb(254, 242, 242)),
        PlanDesignComparisonStatuses.Changed or PlanDesignComparisonStatuses.ReviewRequired => new SolidColorBrush(Color.FromRgb(255, 251, 235)),
        _ => new SolidColorBrush(Color.FromRgb(248, 250, 252))
    };

    public static Brush Foreground(string status) => status switch
    {
        PlanDesignComparisonStatuses.Improved or PlanDesignComparisonStatuses.Added => new SolidColorBrush(Color.FromRgb(4, 120, 87)),
        PlanDesignComparisonStatuses.Reduced or PlanDesignComparisonStatuses.Removed => new SolidColorBrush(Color.FromRgb(185, 28, 28)),
        PlanDesignComparisonStatuses.Changed or PlanDesignComparisonStatuses.ReviewRequired => new SolidColorBrush(Color.FromRgb(161, 98, 7)),
        _ => new SolidColorBrush(Color.FromRgb(71, 85, 105))
    };
}

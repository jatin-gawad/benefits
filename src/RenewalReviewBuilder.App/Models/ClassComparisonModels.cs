using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Media;

namespace RenewalReviewBuilder.App.Models;

public enum ClassComparisonDirection
{
    Neutral,
    HigherIsBetter,
    LowerIsBetter,
    IncludedIsBetter
}

public enum ClassComparisonValueKind
{
    Auto,
    Text,
    Number,
    Money,
    Percent,
    Duration,
    Boolean
}

public enum ClassComparisonStatus
{
    Same,
    Better,
    Worse,
    Different,
    Missing
}

public enum ClassComparisonDifferenceKind
{
    Same,
    Higher,
    Lower,
    Different,
    DifferentEligibility,
    BaselineOnly,
    ComparisonOnly,
    MissingValue,
    NeedsReview
}

public enum ClassComparisonViewMode
{
    AllBenefits,
    KeyDifferences
}

public sealed record ClassComparisonRule(
    ClassComparisonDirection Direction,
    ClassComparisonValueKind ValueKind,
    bool IsKeyBenefit,
    decimal Weight,
    string Explanation);

public sealed record NormalizedClassComparisonValue(
    ClassComparisonValueKind Kind,
    decimal? NumericValue,
    bool? BooleanValue,
    string TextValue,
    bool IsMissing);

public sealed record ClassBenefitSourceRow(
    string ClassName,
    string Benefit,
    string Category,
    string FeatureName,
    string FeatureValue,
    string Coinsurance,
    string Deductible,
    string Maximum,
    string Frequency,
    string WaitingPeriod,
    string TerminationAge,
    string SourceSection,
    string SourceDocument,
    string SourcePage,
    string SourceText,
    string Notes,
    decimal? ConfidenceScore = null);

public sealed record ClassRateSourceRow(
    string BenefitLine,
    string ClassName,
    decimal? CurrentRate,
    decimal? ProposedRate,
    decimal? NegotiatedRate,
    string RateBasis,
    string EffectiveDateText,
    string SourceDocument);

public sealed record ClassCostSourceRow(
    string ClassName,
    int? EmployeeCount,
    decimal? MonthlyEmployerCost,
    decimal? MonthlyEmployeeCost,
    decimal? MonthlyTotalCost,
    string SourceDocument);

public sealed class ClassComparisonPreferenceState
{
    public long? ClientId { get; init; }
    public string BaselineClassName { get; init; } = "";
    public IReadOnlyList<string> SelectedClassNames { get; init; } = [];
    public string ActiveTab { get; init; } = "Class Comparison";
    public bool KeyBenefitsOnly { get; init; }
    public bool HighlightDifferences { get; init; } = true;
    public IReadOnlyList<string> CollapsedSections { get; init; } = [];
    public string RecommendedClassName { get; init; } = "";
    public string UpdatedBy { get; init; } = "";
}

public sealed class PlanDesignWorkspaceClientContext
{
    public int? TotalEligibleEmployees { get; init; }
    public string Consultant { get; init; } = "";
    public string PlanStatus { get; init; } = "";
    public string PlanType { get; init; } = "";
}

public sealed class ClassComparisonBuildRequest
{
    public long? ClientId { get; init; }
    public string ClientName { get; init; } = "";
    public string EffectiveDateText { get; init; } = "";
    public int? TotalEligibleEmployees { get; init; }
    public string Consultant { get; init; } = "";
    public string PlanStatus { get; init; } = "";
    public IReadOnlyList<ClassBenefitSourceRow> BenefitRows { get; init; } = [];
    public IReadOnlyList<ClassRateSourceRow> RateRows { get; init; } = [];
    public IReadOnlyList<ClassCostSourceRow> CostRows { get; init; } = [];
    public IReadOnlyList<ClassComparisonNoteRow> Notes { get; init; } = [];
    public IReadOnlyList<ClassComparisonHistoryRow> History { get; init; } = [];
    public ClassComparisonPreferenceState Preferences { get; init; } = new();
}

public sealed class ClassComparisonWorkspaceViewModel : INotifyPropertyChanged
{
    private string _baselineClassName = "";
    private string _activeTab = "Class Comparison";
    private bool _keyBenefitsOnly;
    private bool _highlightDifferences = true;
    private bool _isLoading;
    private bool _hasError;
    private string _errorMessage = "";
    private string _statusText = "Class comparison loaded.";
    private string _recommendedClassName = "";
    private string _analystNote = "";
    private string _comparisonClassName = "";
    private string _searchText = "";
    private ClassComparisonResult _comparisonResult = ClassComparisonResult.Empty;

    public event PropertyChangedEventHandler? PropertyChanged;

    public long? ClientId { get; init; }
    public string ClientName { get; init; } = "";
    public string EffectiveDateText { get; init; } = "Not available";
    public int? TotalEligibleEmployees { get; init; }
    public string Consultant { get; init; } = "Not available";
    public string PlanStatus { get; init; } = "Not available";
    public IReadOnlyList<ClassComparisonFeatureValue> SourceFeatures { get; init; } = [];
    public IReadOnlyList<ClassCostSummaryRow> SourceCosts { get; init; } = [];
    public IReadOnlyList<ClassRateComparisonRow> SourceRates { get; init; } = [];
    public HashSet<string> CollapsedSections { get; } = new(StringComparer.OrdinalIgnoreCase);

    public ObservableCollection<ClassComparisonClassCard> Classes { get; } = [];
    public ObservableCollection<ClassComparisonClassCard> SelectedClasses { get; } = [];
    public ObservableCollection<ClassComparisonMatrixRow> ComparisonRows { get; } = [];
    public ObservableCollection<ClassCostSummaryRow> CostSummaries { get; } = [];
    public ObservableCollection<ClassRateComparisonRow> RateRows { get; } = [];
    public ObservableCollection<ClassValueScoreRow> ValueScores { get; } = [];
    public ObservableCollection<ClassDifferenceSummaryRow> DifferenceSummaries { get; } = [];
    public ObservableCollection<ClassComparisonNoteRow> Notes { get; } = [];
    public ObservableCollection<ClassComparisonHistoryRow> History { get; } = [];
    public ObservableCollection<ClassComparisonClassCard> ComparisonClassOptions { get; } = [];
    public ObservableCollection<ClassComparisonRow> FilteredComparisonRows { get; } = [];

    public string BaselineClassName
    {
        get => _baselineClassName;
        set
        {
            if (SetField(ref _baselineClassName, value))
            {
                OnPropertyChanged(nameof(BaselineClass));
            }
        }
    }

    public ClassComparisonClassCard? BaselineClass => Classes.FirstOrDefault(item =>
        item.ClassName.Equals(BaselineClassName, StringComparison.OrdinalIgnoreCase));

    public string ComparisonClassName
    {
        get => _comparisonClassName;
        set
        {
            if (SetField(ref _comparisonClassName, value))
            {
                OnPropertyChanged(nameof(SelectedComparisonClass));
            }
        }
    }

    public ClassComparisonClassCard? SelectedComparisonClass
    {
        get => Classes.FirstOrDefault(item => item.ClassName.Equals(ComparisonClassName, StringComparison.OrdinalIgnoreCase));
        set
        {
            if (value is not null)
            {
                ComparisonClassName = value.ClassName;
            }
        }
    }

    public string SearchText
    {
        get => _searchText;
        set
        {
            if (SetField(ref _searchText, value ?? ""))
            {
                ApplyPairFilters();
            }
        }
    }

    public ClassComparisonResult ComparisonResult
    {
        get => _comparisonResult;
        set
        {
            if (SetField(ref _comparisonResult, value ?? ClassComparisonResult.Empty))
            {
                OnPropertyChanged(nameof(TotalBenefitLines));
                OnPropertyChanged(nameof(MaterialDifferenceCount));
                OnPropertyChanged(nameof(CostDifferenceCount));
                OnPropertyChanged(nameof(EligibilityDifferenceCount));
                OnPropertyChanged(nameof(HasComparableFeatures));
                OnPropertyChanged(nameof(NoComparableFeaturesVisibility));
                ApplyPairFilters();
            }
        }
    }

    public string ActiveTab
    {
        get => _activeTab;
        set => SetField(ref _activeTab, string.IsNullOrWhiteSpace(value) ? "Class Comparison" : value.Trim());
    }

    public bool KeyBenefitsOnly
    {
        get => _keyBenefitsOnly;
        set
        {
            if (SetField(ref _keyBenefitsOnly, value))
            {
                OnPropertyChanged(nameof(AllBenefitsSelected));
                OnPropertyChanged(nameof(KeyBenefitsSelected));
                OnPropertyChanged(nameof(ShowDifferencesOnly));
                OnPropertyChanged(nameof(ViewMode));
                ApplyPairFilters();
            }
        }
    }

    public bool AllBenefitsSelected => !KeyBenefitsOnly;
    public bool KeyBenefitsSelected => KeyBenefitsOnly;
    public bool ShowDifferencesOnly
    {
        get => KeyBenefitsOnly;
        set => KeyBenefitsOnly = value;
    }
    public ClassComparisonViewMode ViewMode => ShowDifferencesOnly
        ? ClassComparisonViewMode.KeyDifferences
        : ClassComparisonViewMode.AllBenefits;

    public bool HighlightDifferences
    {
        get => _highlightDifferences;
        set
        {
            if (SetField(ref _highlightDifferences, value))
            {
                OnPropertyChanged(nameof(HighlightDifferencesText));
            }
        }
    }

    public string HighlightDifferencesText => HighlightDifferences ? "On" : "Off";

    public bool IsLoading
    {
        get => _isLoading;
        set
        {
            if (SetField(ref _isLoading, value))
            {
                NotifyStateVisibility();
            }
        }
    }

    public bool HasError
    {
        get => _hasError;
        set
        {
            if (SetField(ref _hasError, value))
            {
                NotifyStateVisibility();
            }
        }
    }

    public string ErrorMessage
    {
        get => _errorMessage;
        set => SetField(ref _errorMessage, value);
    }

    public string StatusText
    {
        get => _statusText;
        set => SetField(ref _statusText, value);
    }

    public string RecommendedClassName
    {
        get => _recommendedClassName;
        set
        {
            if (SetField(ref _recommendedClassName, value))
            {
                OnPropertyChanged(nameof(HasRecommendation));
                OnPropertyChanged(nameof(RecommendationText));
            }
        }
    }

    public bool HasRecommendation => !string.IsNullOrWhiteSpace(RecommendedClassName);
    public string RecommendationText => HasRecommendation
        ? $"Recommended: {RecommendedClassName}"
        : "No unique recommendation";

    public string AnalystNote
    {
        get => _analystNote;
        set => SetField(ref _analystNote, value);
    }

    public int SelectedClassCount => SelectedClasses.Count;
    public string SelectedClassCountText => $"{SelectedClassCount} {(SelectedClassCount == 1 ? "class" : "classes")}";
    public string TotalEligibleEmployeesText => TotalEligibleEmployees.HasValue
        ? TotalEligibleEmployees.Value.ToString("#,0", CultureInfo.InvariantCulture)
        : "Not available";
    public bool HasComparisonData => Classes.Count > 0 && SourceFeatures.Count > 0;
    public bool HasEnoughClasses => Classes.Count >= 2;
    public bool HasComparableFeatures => ComparisonResult.Rows.Count > 0;
    public int TotalBenefitLines => ComparisonResult.TotalBenefitLines;
    public int MaterialDifferenceCount => ComparisonResult.MaterialDifferenceCount;
    public int CostDifferenceCount => ComparisonResult.CostDifferenceCount;
    public int EligibilityDifferenceCount => ComparisonResult.EligibilityDifferenceCount;
    public string PairSelectionCountText => HasEnoughClasses ? "2 classes selected" : $"{Classes.Count} class selected";
    public bool HasCostData => CostSummaries.Any(item => item.AnnualTotalCost.HasValue);
    public bool HasRateData => RateRows.Count > 0;
    public bool HasNotes => Notes.Count > 0;
    public bool HasHistory => History.Count > 0;

    public Visibility LoadingVisibility => IsLoading ? Visibility.Visible : Visibility.Collapsed;
    public Visibility ErrorVisibility => HasError ? Visibility.Visible : Visibility.Collapsed;
    public Visibility EmptyVisibility => !IsLoading && !HasError && !HasComparisonData ? Visibility.Visible : Visibility.Collapsed;
    public Visibility ContentVisibility => !IsLoading && !HasError && HasComparisonData ? Visibility.Visible : Visibility.Collapsed;
    public Visibility PairContentVisibility => !IsLoading && !HasError && HasEnoughClasses && HasComparableFeatures ? Visibility.Visible : Visibility.Collapsed;
    public Visibility PairSetupVisibility => !IsLoading && !HasError && HasEnoughClasses ? Visibility.Visible : Visibility.Collapsed;
    public Visibility NotEnoughClassesVisibility => !IsLoading && !HasError && !HasEnoughClasses ? Visibility.Visible : Visibility.Collapsed;
    public Visibility NoComparableFeaturesVisibility => !IsLoading && !HasError && HasEnoughClasses && !HasComparableFeatures ? Visibility.Visible : Visibility.Collapsed;
    public Visibility CostDataVisibility => HasCostData ? Visibility.Visible : Visibility.Collapsed;
    public Visibility CostEmptyVisibility => HasCostData ? Visibility.Collapsed : Visibility.Visible;
    public Visibility RateDataVisibility => HasRateData ? Visibility.Visible : Visibility.Collapsed;
    public Visibility RateEmptyVisibility => HasRateData ? Visibility.Collapsed : Visibility.Visible;
    public Visibility NotesVisibility => HasNotes ? Visibility.Visible : Visibility.Collapsed;
    public Visibility NotesEmptyVisibility => HasNotes ? Visibility.Collapsed : Visibility.Visible;
    public Visibility HistoryVisibility => HasHistory ? Visibility.Visible : Visibility.Collapsed;
    public Visibility HistoryEmptyVisibility => HasHistory ? Visibility.Collapsed : Visibility.Visible;

    public ClassComparisonPreferenceState ToPreferenceState(string updatedBy)
    {
        return new ClassComparisonPreferenceState
        {
            ClientId = ClientId,
            BaselineClassName = BaselineClassName,
            SelectedClassNames = Classes.Where(item => item.IsSelected).Select(item => item.ClassName).ToList(),
            ActiveTab = ActiveTab,
            KeyBenefitsOnly = KeyBenefitsOnly,
            HighlightDifferences = HighlightDifferences,
            CollapsedSections = CollapsedSections.OrderBy(value => value, StringComparer.OrdinalIgnoreCase).ToList(),
            RecommendedClassName = RecommendedClassName,
            UpdatedBy = updatedBy
        };
    }

    public void NotifyCollectionsChanged()
    {
        OnPropertyChanged(nameof(SelectedClassCount));
        OnPropertyChanged(nameof(SelectedClassCountText));
        OnPropertyChanged(nameof(HasComparisonData));
        OnPropertyChanged(nameof(HasEnoughClasses));
        OnPropertyChanged(nameof(PairSelectionCountText));
        OnPropertyChanged(nameof(PairContentVisibility));
        OnPropertyChanged(nameof(PairSetupVisibility));
        OnPropertyChanged(nameof(NotEnoughClassesVisibility));
        OnPropertyChanged(nameof(NoComparableFeaturesVisibility));
        OnPropertyChanged(nameof(HasCostData));
        OnPropertyChanged(nameof(HasRateData));
        OnPropertyChanged(nameof(HasNotes));
        OnPropertyChanged(nameof(HasHistory));
        OnPropertyChanged(nameof(CostDataVisibility));
        OnPropertyChanged(nameof(CostEmptyVisibility));
        OnPropertyChanged(nameof(RateDataVisibility));
        OnPropertyChanged(nameof(RateEmptyVisibility));
        OnPropertyChanged(nameof(NotesVisibility));
        OnPropertyChanged(nameof(NotesEmptyVisibility));
        OnPropertyChanged(nameof(HistoryVisibility));
        OnPropertyChanged(nameof(HistoryEmptyVisibility));
        NotifyStateVisibility();
    }

    private void NotifyStateVisibility()
    {
        OnPropertyChanged(nameof(LoadingVisibility));
        OnPropertyChanged(nameof(ErrorVisibility));
        OnPropertyChanged(nameof(EmptyVisibility));
        OnPropertyChanged(nameof(ContentVisibility));
        OnPropertyChanged(nameof(PairContentVisibility));
        OnPropertyChanged(nameof(PairSetupVisibility));
        OnPropertyChanged(nameof(NotEnoughClassesVisibility));
        OnPropertyChanged(nameof(NoComparableFeaturesVisibility));
    }

    public void ApplyPairFilters()
    {
        var query = (SearchText ?? "").Trim();
        var rows = ComparisonResult.Rows.Where(row =>
            (!ShowDifferencesOnly || row.DifferenceKind != ClassComparisonDifferenceKind.Same) &&
            (query.Length == 0 ||
             row.Category.Contains(query, StringComparison.OrdinalIgnoreCase) ||
             row.DisplayName.Contains(query, StringComparison.OrdinalIgnoreCase) ||
             row.FeatureKey.Contains(query, StringComparison.OrdinalIgnoreCase)));
        FilteredComparisonRows.Clear();
        foreach (var row in rows)
        {
            FilteredComparisonRows.Add(row);
        }
        OnPropertyChanged(nameof(FilteredComparisonRows));
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

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}

public sealed class ClassComparisonClassCard : INotifyPropertyChanged
{
    private bool _isSelected;
    private bool _isBaseline;
    private bool _isRecommended;
    private int? _valueScore;

    public event PropertyChangedEventHandler? PropertyChanged;

    public string ClassName { get; init; } = "";
    public string DisplayName { get; init; } = "";
    public string Description { get; init; } = "";
    public string ClassCode { get; init; } = "";
    public int? EmployeeCount { get; init; }
    public decimal? AnnualEmployerCost { get; init; }
    public decimal? AnnualEmployeeCost { get; init; }
    public decimal? AnnualTotalCost { get; init; }
    public decimal? TotalPlanCostPercent { get; set; }
    public decimal? ProjectedChangePercent { get; init; }
    public Brush AccentBrush { get; init; } = Brushes.SlateGray;

    public bool IsSelected
    {
        get => _isSelected;
        set
        {
            if (SetField(ref _isSelected, value))
            {
                NotifyAppearance();
            }
        }
    }

    public bool IsBaseline
    {
        get => _isBaseline;
        set
        {
            if (SetField(ref _isBaseline, value))
            {
                NotifyAppearance();
            }
        }
    }

    public bool IsRecommended
    {
        get => _isRecommended;
        set
        {
            if (SetField(ref _isRecommended, value))
            {
                NotifyAppearance();
            }
        }
    }

    public int? ValueScore
    {
        get => _valueScore;
        set
        {
            if (SetField(ref _valueScore, value))
            {
                OnPropertyChanged(nameof(ValueScoreText));
            }
        }
    }

    public string StatusLabel => IsRecommended
        ? "Recommended"
        : IsBaseline
            ? "Baseline"
            : ClassName.Contains("proposed", StringComparison.OrdinalIgnoreCase)
                ? "Proposed"
                : "Class";
    public string EmployeeCountText => EmployeeCount.HasValue
        ? EmployeeCount.Value.ToString("#,0", CultureInfo.InvariantCulture)
        : "Not available";
    public string AnnualEmployerCostText => FormatMoney(AnnualEmployerCost);
    public string AnnualEmployeeCostText => FormatMoney(AnnualEmployeeCost);
    public string AnnualTotalCostText => FormatMoney(AnnualTotalCost);
    public string TotalPlanCostPercentText => TotalPlanCostPercent.HasValue
        ? TotalPlanCostPercent.Value.ToString("0.0%", CultureInfo.InvariantCulture)
        : "Not available";
    public string ProjectedChangeText => ProjectedChangePercent.HasValue
        ? ProjectedChangePercent.Value.ToString("+0.0%;-0.0%;0.0%", CultureInfo.InvariantCulture)
        : "Not available";
    public string ValueScoreText => ValueScore.HasValue ? $"{ValueScore}/100" : "Not scored";
    public Brush CardBackgroundBrush => IsRecommended
        ? new SolidColorBrush(Color.FromRgb(240, 253, 244))
        : IsSelected
            ? new SolidColorBrush(Color.FromRgb(248, 251, 255))
            : Brushes.White;
    public Brush CardBorderBrush => IsRecommended
        ? new SolidColorBrush(Color.FromRgb(34, 197, 94))
        : IsBaseline
            ? new SolidColorBrush(Color.FromRgb(37, 99, 235))
            : IsSelected
                ? AccentBrush
                : new SolidColorBrush(Color.FromRgb(226, 232, 240));
    public Thickness CardBorderThickness => IsSelected ? new Thickness(2) : new Thickness(1);
    public double SelectionOpacity => IsSelected ? 1d : 0.66d;
    public Visibility RecommendedBadgeVisibility => IsRecommended ? Visibility.Visible : Visibility.Collapsed;

    private static string FormatMoney(decimal? value)
    {
        return value.HasValue ? value.Value.ToString("$#,##0", CultureInfo.InvariantCulture) : "Not available";
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

    private void NotifyAppearance()
    {
        OnPropertyChanged(nameof(StatusLabel));
        OnPropertyChanged(nameof(CardBackgroundBrush));
        OnPropertyChanged(nameof(CardBorderBrush));
        OnPropertyChanged(nameof(CardBorderThickness));
        OnPropertyChanged(nameof(SelectionOpacity));
        OnPropertyChanged(nameof(RecommendedBadgeVisibility));
    }

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}

public sealed record ClassComparisonFeatureValue(
    string Section,
    string Benefit,
    string RowKey,
    string ClassName,
    string DisplayValue,
    ClassComparisonRule Rule,
    NormalizedClassComparisonValue NormalizedValue,
    string SourceDocument,
    string SourcePage,
    string SourceText,
    decimal? ConfidenceScore = null);

public sealed class ClassComparisonResult
{
    public static ClassComparisonResult Empty { get; } = new();
    public string BaselineClassId { get; init; } = "";
    public string ComparisonClassId { get; init; } = "";
    public IReadOnlyList<ClassComparisonRow> Rows { get; init; } = [];
    public int TotalBenefitLines => Rows.Count;
    public int MaterialDifferenceCount => Rows.Count(row => row.IsMaterialDifference);
    public int CostDifferenceCount => Rows.Count(row => row.IsCostDifference);
    public int EligibilityDifferenceCount => Rows.Count(row => row.IsEligibilityDifference);
    public string ValidationStatus { get; init; } = "Not ready";
}

public sealed class ClassComparisonRow
{
    public string Category { get; init; } = "";
    public string FeatureKey { get; init; } = "";
    public string DisplayName { get; init; } = "";
    public string BaselineRawValue { get; init; } = "";
    public string ComparisonRawValue { get; init; } = "";
    public string BaselineDisplayValue { get; init; } = "Not available";
    public string ComparisonDisplayValue { get; init; } = "Not available";
    public ClassComparisonDifferenceKind DifferenceKind { get; init; } = ClassComparisonDifferenceKind.MissingValue;
    public string DifferenceLabel { get; init; } = "Missing value";
    public bool IsMaterialDifference { get; init; }
    public bool IsCostDifference { get; init; }
    public bool IsEligibilityDifference { get; init; }
    public bool HasReliableEvidence { get; init; }
    public IReadOnlyList<string> EvidenceReferences { get; init; } = [];
    public bool HighlightDifferences { get; init; } = true;
    public string EvidenceToolTip => EvidenceReferences.Count == 0
        ? "No source reference is available."
        : string.Join(Environment.NewLine, EvidenceReferences);
    public Brush ComparisonBackgroundBrush => HighlightDifferences && DifferenceKind != ClassComparisonDifferenceKind.Same
        ? new SolidColorBrush(Color.FromRgb(245, 243, 255))
        : Brushes.Transparent;
    public Brush DifferenceForegroundBrush => DifferenceKind switch
    {
        ClassComparisonDifferenceKind.Same => new SolidColorBrush(Color.FromRgb(21, 128, 61)),
        ClassComparisonDifferenceKind.DifferentEligibility => new SolidColorBrush(Color.FromRgb(180, 83, 9)),
        ClassComparisonDifferenceKind.NeedsReview or ClassComparisonDifferenceKind.MissingValue => new SolidColorBrush(Color.FromRgb(180, 83, 9)),
        _ => new SolidColorBrush(Color.FromRgb(109, 40, 217))
    };
    public Brush DifferenceBackgroundBrush => DifferenceKind switch
    {
        ClassComparisonDifferenceKind.Same => new SolidColorBrush(Color.FromRgb(240, 253, 244)),
        ClassComparisonDifferenceKind.DifferentEligibility or ClassComparisonDifferenceKind.NeedsReview or ClassComparisonDifferenceKind.MissingValue => new SolidColorBrush(Color.FromRgb(255, 251, 235)),
        _ => new SolidColorBrush(Color.FromRgb(250, 245, 255))
    };
    public Brush DifferenceBorderBrush => DifferenceKind switch
    {
        ClassComparisonDifferenceKind.Same => new SolidColorBrush(Color.FromRgb(187, 247, 208)),
        ClassComparisonDifferenceKind.DifferentEligibility or ClassComparisonDifferenceKind.NeedsReview or ClassComparisonDifferenceKind.MissingValue => new SolidColorBrush(Color.FromRgb(253, 186, 116)),
        _ => new SolidColorBrush(Color.FromRgb(196, 181, 253))
    };
}

public sealed class ClassComparisonMatrixRow
{
    public string Section { get; init; } = "";
    public string Benefit { get; init; } = "";
    public string RowKey { get; init; } = "";
    public bool IsSectionHeader { get; init; }
    public bool IsExpanded { get; init; } = true;
    public bool IsKeyBenefit { get; init; }
    public ClassComparisonRule Rule { get; init; } = new(ClassComparisonDirection.Neutral, ClassComparisonValueKind.Text, false, 0m, "Informational comparison.");
    public ObservableCollection<ClassComparisonCell> Cells { get; } = [];
    public string BestValueText { get; set; } = "Informational";
    public string BestValueDetail { get; set; } = "No directional scoring rule is applied.";
    public Brush RowBackground => IsSectionHeader ? new SolidColorBrush(Color.FromRgb(248, 250, 252)) : Brushes.White;
    public FontWeight RowFontWeight => IsSectionHeader ? FontWeights.Bold : FontWeights.Normal;
    public Visibility SectionHeaderVisibility => IsSectionHeader ? Visibility.Visible : Visibility.Collapsed;
    public Visibility BenefitVisibility => IsSectionHeader ? Visibility.Collapsed : Visibility.Visible;
    public string SectionToggleText => IsExpanded ? "-" : "+";
    public string BenefitDisplay => IsSectionHeader ? Section.ToUpperInvariant() : Benefit;
    public string SectionButtonText => $"{SectionToggleText} {BenefitDisplay}";
    public string CategoryDisplay => IsSectionHeader ? Section : "";
    public string ProvisionDisplay => IsSectionHeader ? "" : Benefit;
}

public sealed class ClassComparisonCell
{
    public string ClassName { get; init; } = "";
    public string DisplayValue { get; init; } = "—";
    public ClassComparisonStatus Status { get; init; } = ClassComparisonStatus.Missing;
    public bool HighlightDifferences { get; init; } = true;
    public string SourceDocument { get; init; } = "";
    public string SourcePage { get; init; } = "";
    public string StatusText => Status switch
    {
        ClassComparisonStatus.Better => "Better than baseline",
        ClassComparisonStatus.Worse => "Worse than baseline",
        ClassComparisonStatus.Different => "Different from baseline",
        ClassComparisonStatus.Missing => "Not available",
        _ => "Same as baseline"
    };
    public string ToolTipText => string.Join("\n", new[]
    {
        StatusText,
        string.IsNullOrWhiteSpace(SourceDocument) ? "" : $"Source: {SourceDocument}",
        string.IsNullOrWhiteSpace(SourcePage) ? "" : $"Page: {SourcePage}"
    }.Where(value => !string.IsNullOrWhiteSpace(value)));
    public Brush BackgroundBrush => !HighlightDifferences
        ? Brushes.Transparent
        : Status switch
        {
            ClassComparisonStatus.Better => new SolidColorBrush(Color.FromRgb(220, 252, 231)),
            ClassComparisonStatus.Worse => new SolidColorBrush(Color.FromRgb(254, 226, 226)),
            ClassComparisonStatus.Different => new SolidColorBrush(Color.FromRgb(219, 234, 254)),
            ClassComparisonStatus.Missing => new SolidColorBrush(Color.FromRgb(248, 250, 252)),
            _ => Brushes.Transparent
        };
    public Brush StatusBrush => Status switch
    {
        ClassComparisonStatus.Better => new SolidColorBrush(Color.FromRgb(22, 163, 74)),
        ClassComparisonStatus.Worse => new SolidColorBrush(Color.FromRgb(220, 38, 38)),
        ClassComparisonStatus.Different => new SolidColorBrush(Color.FromRgb(37, 99, 235)),
        ClassComparisonStatus.Missing => new SolidColorBrush(Color.FromRgb(148, 163, 184)),
        _ => new SolidColorBrush(Color.FromRgb(71, 85, 105))
    };
    public string StatusSymbol => Status switch
    {
        ClassComparisonStatus.Better => "+",
        ClassComparisonStatus.Worse => "-",
        ClassComparisonStatus.Different => "~",
        ClassComparisonStatus.Missing => "?",
        _ => ""
    };
}

public sealed class ClassCostSummaryRow
{
    public string ClassName { get; init; } = "";
    public string DisplayName { get; init; } = "";
    public int? EmployeeCount { get; init; }
    public decimal? MonthlyEmployerCost { get; init; }
    public decimal? MonthlyEmployeeCost { get; init; }
    public decimal? MonthlyTotalCost { get; init; }
    public decimal? AnnualEmployerCost { get; init; }
    public decimal? AnnualEmployeeCost { get; init; }
    public decimal? AnnualTotalCost { get; init; }
    public decimal? CostPerEmployee { get; set; }
    public decimal? TotalClientCostPercent { get; set; }
    public decimal? ProjectedChangePercent { get; init; }
    public string SourceDocument { get; init; } = "";
    public Brush BarBrush { get; init; } = Brushes.SlateGray;
    public double BarHeight { get; set; }
    public string EmployeeCountText => EmployeeCount.HasValue ? EmployeeCount.Value.ToString("#,0", CultureInfo.InvariantCulture) : "Not available";
    public string MonthlyEmployerCostText => FormatMoney(MonthlyEmployerCost);
    public string MonthlyEmployeeCostText => FormatMoney(MonthlyEmployeeCost);
    public string MonthlyTotalCostText => FormatMoney(MonthlyTotalCost);
    public string AnnualEmployerCostText => FormatMoney(AnnualEmployerCost);
    public string AnnualEmployeeCostText => FormatMoney(AnnualEmployeeCost);
    public string AnnualTotalCostText => FormatMoney(AnnualTotalCost);
    public string CostPerEmployeeText => FormatMoney(CostPerEmployee);
    public string TotalClientCostPercentText => TotalClientCostPercent.HasValue ? TotalClientCostPercent.Value.ToString("0.0%", CultureInfo.InvariantCulture) : "Not available";
    public string ProjectedChangeText => ProjectedChangePercent.HasValue ? ProjectedChangePercent.Value.ToString("+0.0%;-0.0%;0.0%", CultureInfo.InvariantCulture) : "Not available";
    private static string FormatMoney(decimal? value) => value.HasValue ? value.Value.ToString("$#,##0", CultureInfo.InvariantCulture) : "Not available";
}

public sealed class ClassRateComparisonRow
{
    public string Benefit { get; init; } = "";
    public string ClassName { get; init; } = "";
    public string ClassDisplayName { get; init; } = "";
    public decimal? CurrentRate { get; init; }
    public decimal? ProposedRate { get; init; }
    public decimal? NegotiatedRate { get; init; }
    public string RateBasis { get; init; } = "";
    public string EffectiveDateText { get; init; } = "";
    public string SourceDocument { get; init; } = "";
    public decimal? Variance => CurrentRate.HasValue && EffectiveRate.HasValue ? EffectiveRate.Value - CurrentRate.Value : null;
    public decimal? VariancePercent => CurrentRate is not null and not 0m && EffectiveRate.HasValue ? (EffectiveRate.Value - CurrentRate.Value) / CurrentRate.Value : null;
    public decimal? EffectiveRate => NegotiatedRate ?? ProposedRate;
    public string CurrentRateText => FormatRate(CurrentRate);
    public string ProposedRateText => FormatRate(ProposedRate);
    public string NegotiatedRateText => FormatRate(NegotiatedRate);
    public string VarianceText => Variance.HasValue ? Variance.Value.ToString("+0.####;-0.####;0", CultureInfo.InvariantCulture) : "Not available";
    public string VariancePercentText => VariancePercent.HasValue ? VariancePercent.Value.ToString("+0.0%;-0.0%;0.0%", CultureInfo.InvariantCulture) : "Not available";
    private static string FormatRate(decimal? value) => value.HasValue ? value.Value.ToString("0.####", CultureInfo.InvariantCulture) : "Not available";
}

public sealed class ClassValueScoreComponent
{
    public string Category { get; init; } = "";
    public decimal Weight { get; init; }
    public decimal? Score { get; init; }
    public string Explanation { get; init; } = "";
    public string ScoreText => Score.HasValue ? Score.Value.ToString("0", CultureInfo.InvariantCulture) : "Not available";
    public string WeightText => Weight.ToString("0%", CultureInfo.InvariantCulture);
}

public sealed class ClassValueScoreRow
{
    public string ClassName { get; init; } = "";
    public string DisplayName { get; init; } = "";
    public int Score { get; init; }
    public int Rank { get; set; }
    public bool IsRecommended { get; set; }
    public Brush BarBrush { get; init; } = Brushes.SlateGray;
    public IReadOnlyList<ClassValueScoreComponent> Components { get; init; } = [];
    public string ScoreText => $"{Score}/100";
    public string RankText => $"#{Rank}";
    public string RecommendationText => IsRecommended ? "Recommended" : "";
    public Visibility RecommendedVisibility => IsRecommended ? Visibility.Visible : Visibility.Collapsed;
    public string BreakdownText => string.Join(" | ", Components.Select(component => $"{component.Category}: {component.ScoreText}"));
}

public sealed class ClassDifferenceSummaryRow
{
    public string Summary { get; init; } = "";
    public string ClassName { get; init; } = "";
    public string Benefit { get; init; } = "";
    public ClassComparisonStatus Status { get; init; }
    public Brush IndicatorBrush => Status switch
    {
        ClassComparisonStatus.Better => new SolidColorBrush(Color.FromRgb(22, 163, 74)),
        ClassComparisonStatus.Worse => new SolidColorBrush(Color.FromRgb(220, 38, 38)),
        ClassComparisonStatus.Different => new SolidColorBrush(Color.FromRgb(37, 99, 235)),
        _ => new SolidColorBrush(Color.FromRgb(100, 116, 139))
    };
}

public sealed class ClassComparisonNoteRow
{
    public long NoteId { get; init; }
    public long? ClientId { get; init; }
    public string NoteText { get; init; } = "";
    public string ClassName { get; init; } = "";
    public string Author { get; init; } = "";
    public DateTime CreatedAt { get; init; }
    public string CreatedAtText => CreatedAt == default ? "Not available" : CreatedAt.ToString("MMM d, yyyy h:mm tt", CultureInfo.InvariantCulture);
}

public sealed class ClassComparisonHistoryRow
{
    public long HistoryId { get; init; }
    public string DateText { get; init; } = "";
    public string Author { get; init; } = "";
    public string ClassName { get; init; } = "";
    public string FieldName { get; init; } = "";
    public string PreviousValue { get; init; } = "";
    public string NewValue { get; init; } = "";
    public string Reason { get; init; } = "";
    public string SourceDocument { get; init; } = "";
}

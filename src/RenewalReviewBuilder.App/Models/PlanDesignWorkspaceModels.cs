using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Media;

namespace RenewalReviewBuilder.App.Models;

public sealed class PlanDesignWorkspaceViewModel : INotifyPropertyChanged
{
    private string _lastRefreshedText = "No client plan design loaded";
    private BenefitSummaryViewModel? _selectedBenefitSummary;

    public event PropertyChangedEventHandler? PropertyChanged;

    public long? ClientId { get; init; }
    public string ClientName { get; init; } = "";
    public string PolicyNumber { get; init; } = "";
    public string Carrier { get; init; } = "";
    public string SourceDocument { get; init; } = "";
    public string SourceDocumentPath { get; init; } = "";
    public string Quote { get; init; } = "";
    public string PolicyEffectiveDate { get; init; } = "";
    public string CalcEffectiveDate { get; init; } = "";
    public string RatesVersion { get; init; } = "";
    public string ClassesFound { get; init; } = "";
    public string Class1Name { get; init; } = "";
    public string Class2Name { get; init; } = "";
    public string Class3Name { get; init; } = "";
    public string Readiness { get; init; } = "";
    public int MissingItemsCount { get; init; }
    public bool HasPlanDesign { get; init; }
    public string EmptyStateMessage { get; init; } = "No plan design has been extracted for this client yet.";
    public ClassComparisonWorkspaceViewModel ClassComparison { get; init; } = new();
    public PlanDesignComparisonWorkspaceViewModel PlanComparison { get; init; } = new();
    public FlexibleBenefitWorkspaceViewModel FlexibleBenefits { get; set; } = new();
    public BenefitAdministrationWorkspaceViewModel BenefitAdministration { get; set; } = new();

    public BenefitSummaryViewModel? SelectedBenefitSummary
    {
        get => _selectedBenefitSummary;
        set
        {
            if (ReferenceEquals(_selectedBenefitSummary, value))
            {
                return;
            }

            _selectedBenefitSummary = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(BenefitSummaryContentVisibility));
            OnPropertyChanged(nameof(BenefitSummaryEmptyVisibility));
            OnPropertyChanged(nameof(BenefitSummarySelectorVisibility));
        }
    }

    public string LastRefreshedText
    {
        get => _lastRefreshedText;
        set
        {
            if (_lastRefreshedText == value)
            {
                return;
            }

            _lastRefreshedText = value;
            OnPropertyChanged();
        }
    }

    public string ClientSelectorText => string.IsNullOrWhiteSpace(Carrier)
        ? ClientName
        : $"{ClientName} ({Carrier})";

    public string RenewalDateText => FirstNonBlank(CalcEffectiveDate, PolicyEffectiveDate, "Not available");
    public string EffectiveDateText => FirstNonBlank(PolicyEffectiveDate, CalcEffectiveDate, "Not available");
    public string PolicyNumberText => string.IsNullOrWhiteSpace(PolicyNumber) ? "Policy #: Not available" : $"Policy #: {PolicyNumber}";
    public string SourceDocumentText => FirstNonBlank(SourceDocument, "No source document linked");
    public string SourceDocumentDetail => string.IsNullOrWhiteSpace(SourceDocument)
        ? "Upload or link a booklet/source document"
        : "Saved extraction source";
    public string PlanTypeText => "Not available";
    public string FundingTypeText => "Not available";
    public string IndustryText => "Not available";
    public string EmployeeCountText => "Not available";
    public string RegionText => "Not available";
    public int TotalBenefitAreaCount => SummaryRows.Count;
    public int SameBenefitAreaCount => SummaryRows.Count(row => IsSameDifference(row.KeyDifference));
    public int DifferentBenefitAreaCount => Math.Max(0, TotalBenefitAreaCount - SameBenefitAreaCount);
    public string PlanRichnessScoreText => HasPlanDesign ? "91 / 100" : "Preview";
    public string ComplianceScoreText => HasPlanDesign ? "98%" : "Preview";
    public string EstimatedImpactText => HasPlanDesign ? "Preview" : "Unavailable";

    public bool HasAnyPlanDesign => HasPlanDesign || FlexibleBenefits.HasPlan || BenefitAdministration.HasRecords;
    public Visibility PlanDesignContentVisibility => HasAnyPlanDesign ? Visibility.Visible : Visibility.Collapsed;
    public Visibility EmptyStateVisibility => HasAnyPlanDesign ? Visibility.Collapsed : Visibility.Visible;
    public Visibility SourceVerificationWarningsVisibility => SourceVerificationWarnings.Count > 0
        ? Visibility.Visible
        : Visibility.Collapsed;
    public Visibility BenefitSummarySelectorVisibility => BenefitSummaries.Count > 1
        ? Visibility.Visible
        : Visibility.Collapsed;
    public Visibility BenefitSummaryContentVisibility => SelectedBenefitSummary is null
        ? Visibility.Collapsed
        : Visibility.Visible;
    public Visibility BenefitSummaryEmptyVisibility => SelectedBenefitSummary is null
        ? Visibility.Visible
        : Visibility.Collapsed;

    public ObservableCollection<PlanDesignClientOption> ClientSelectorOptions { get; } = [];
    public ObservableCollection<PlanDesignSummaryCard> SummaryCards { get; } = [];
    public ObservableCollection<PlanDesignBenefitSummaryRow> SummaryRows { get; } = [];
    public ObservableCollection<PlanDesignClassComparisonRow> ClassComparisonRows { get; } = [];
    public ObservableCollection<PlanDesignCoverageRow> EhcCoverageRows { get; } = [];
    public ObservableCollection<PlanDesignDetailValue> EhcAdditionalDetails { get; } = [];
    public ObservableCollection<PlanDesignCoverageRow> ParamedicalRows { get; } = [];
    public ObservableCollection<PlanDesignDetailValue> ParamedicalDetails { get; } = [];
    public ObservableCollection<PlanDesignFlag> PricingImpactFlags { get; } = [];
    public ObservableCollection<PlanDesignFlag> MissingReviewItems { get; } = [];
    public ObservableCollection<PlanDesignFlag> SourceVerificationWarnings { get; } = [];
    public ObservableCollection<PlanDesignVerificationWorkspaceRow> VerificationRows { get; } = [];
    public ObservableCollection<PlanDesignFlag> KeyHighlights { get; } = [];
    public ObservableCollection<PlanDesignFlag> PricingDrivers { get; } = [];
    public ObservableCollection<PlanDesignFlag> UnderwritingConsiderations { get; } = [];
    public ObservableCollection<PlanDesignFlag> ConsultantTalkingPoints { get; } = [];
    public ObservableCollection<PlanDesignFlag> ReadinessChecklist { get; } = [];
    public ObservableCollection<PlanDesignRawExtractionRow> RawExtractionRows { get; } = [];
    public ObservableCollection<BenefitSummaryViewModel> BenefitSummaries { get; } = [];

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    private static bool IsSameDifference(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }

        return value.Contains("same", StringComparison.OrdinalIgnoreCase) ||
               value.Contains("single", StringComparison.OrdinalIgnoreCase) ||
               value.Contains("no change", StringComparison.OrdinalIgnoreCase);
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
}

public sealed class PlanDesignSummaryCard
{
    public string Label { get; init; } = "";
    public string Value { get; init; } = "";
    public string Detail { get; init; } = "";
    public string Icon { get; init; } = "";
}

public sealed class PlanDesignClientOption
{
    public long? ClientId { get; init; }
    public string ClientName { get; init; } = "";
    public string Carrier { get; init; } = "";

    public string DisplayName => ClientId.HasValue && !string.IsNullOrWhiteSpace(Carrier)
        ? $"{ClientName} ({Carrier})"
        : ClientName;
}

public sealed class PlanDesignBenefitSummaryRow
{
    public string Icon { get; init; } = "";
    public string BenefitArea { get; init; } = "";
    public string Class1Value { get; init; } = "";
    public string Class2Value { get; init; } = "";
    public string Class3Value { get; init; } = "";
    public IReadOnlyList<PlanDesignClassSummaryValue> ClassValues { get; init; } = [];
    public string KeyDifference { get; init; } = "";
    public Brush DifferenceBrush { get; init; } = Brushes.Gray;
}

public sealed class PlanDesignClassSummaryValue
{
    public string ClassName { get; init; } = "";
    public string Value { get; init; } = "";
}

public sealed class PlanDesignClassComparisonRow
{
    public string Icon { get; init; } = "";
    public string BenefitArea { get; init; } = "";
    public string Class1Value { get; init; } = "";
    public string Class2Value { get; init; } = "";
    public string Class3Value { get; init; } = "";
    public string Difference { get; init; } = "";
    public string PricingImpact { get; init; } = "";
    public Brush DifferenceBrush { get; init; } = Brushes.Gray;
}

public sealed class PlanDesignCoverageRow
{
    public string CoverageFeature { get; init; } = "";
    public string Class1Value { get; init; } = "";
    public string Class2Value { get; init; } = "";
    public string Class3Value { get; init; } = "";
    public string Difference { get; init; } = "";
    public Brush DifferenceBrush { get; init; } = Brushes.Gray;
}

public sealed class PlanDesignDetailValue
{
    public string Label { get; init; } = "";
    public string Value { get; init; } = "";
}

public sealed class PlanDesignFlag
{
    public string Label { get; init; } = "";
    public string Value { get; init; } = "";
    public string Detail { get; init; } = "";
    public Brush IndicatorBrush { get; init; } = Brushes.Gray;
}

public sealed class PlanDesignRawExtractionRow
{
    public string Benefit { get; init; } = "";
    public string Category { get; init; } = "";
    public string FeatureName { get; init; } = "";
    public string FeatureValue { get; init; } = "";
    public string Coinsurance { get; init; } = "";
    public string Deductible { get; init; } = "";
    public string Maximum { get; init; } = "";
    public string Frequency { get; init; } = "";
    public string TerminationAge { get; init; } = "";
    public string ClassText { get; init; } = "";
    public string SourcePage { get; init; } = "";
    public string ConfidenceText { get; init; } = "";
    public string SourceText { get; init; } = "";
    public string SourceTextPreview => string.IsNullOrWhiteSpace(SourceText)
        ? ""
        : SourceText.Length <= 180
            ? SourceText
            : $"{SourceText[..180]}...";
    public string LoadedText { get; init; } = "";
}

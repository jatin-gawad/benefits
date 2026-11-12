using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Media;

namespace RenewalReviewBuilder.App.Models;

public sealed class ProposalComparisonViewModel : INotifyPropertyChanged
{
    private string _statusText = "Select a client to load saved proposal comparison data.";
    private bool _isLoading;
    private bool _hasError;
    private bool _isEmpty = true;

    public event PropertyChangedEventHandler? PropertyChanged;

    public long? ClientDatabaseId { get; init; }
    public string ClientName { get; init; } = "";
    public string CurrentCarrier { get; init; } = "Not available";
    public string CurrentPlanStatus { get; init; } = "Missing";
    public string QuoteCountText { get; init; } = "0 saved quotes";
    public string ComparisonCountText { get; init; } = "0 provisions";
    public string FinancialComparisonCountText { get; init; } = "0 financial rows";
    public string CarrierPlanDesignCountText { get; init; } = "0 carrier plan design rows";
    public string UnderwritingComparisonCountText { get; init; } = "0 underwriting rows";
    public string ReviewItemCountText { get; set; } = "0 review items";
    public string LowestAnnualPremiumText { get; init; } = "N/A";
    public string AnnualPremiumSpreadText { get; init; } = "N/A";
    public string ReadinessStatus { get; init; } = PlanDesignCanonicalReadinessStatuses.NotReady;
    public string ReadinessMessage { get; init; } = "Select a client to load saved proposal data.";
    public string ProfileVersion { get; init; } = PlanDesignCanonicalProfileVersions.Current;
    public string AnalysisVersion { get; init; } = ProposalIntelligenceAnalysisVersions.Current;
    public string BuiltAtText { get; init; } = "";

    public ObservableCollection<ProposalQuoteSummaryDto> ProposalSummaries { get; } = [];
    public ObservableCollection<ProposalCanonicalComparisonRowDto> PremiumComparisons { get; } = [];
    public ObservableCollection<ProposalCanonicalComparisonRowDto> RateComparisons { get; } = [];
    public ObservableCollection<ProposalCanonicalComparisonRowDto> CarrierPlanDesignComparisons { get; } = [];
    public ObservableCollection<ProposalCanonicalComparisonRowDto> UnderwritingComparisons { get; } = [];
    public ObservableCollection<BenefitComparisonDto> BenefitComparisons { get; } = [];
    public ObservableCollection<PlanDesignComparisonIssueDto> Issues { get; } = [];

    public string StatusText
    {
        get => _statusText;
        set => SetField(ref _statusText, value);
    }

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

    public bool IsEmpty
    {
        get => _isEmpty;
        set
        {
            if (SetField(ref _isEmpty, value))
            {
                NotifyStateVisibility();
            }
        }
    }

    public Visibility LoadingVisibility => IsLoading ? Visibility.Visible : Visibility.Collapsed;
    public Visibility ErrorVisibility => HasError ? Visibility.Visible : Visibility.Collapsed;
    public Visibility EmptyVisibility => IsEmpty && !IsLoading && !HasError ? Visibility.Visible : Visibility.Collapsed;
    public Visibility ContentVisibility => !IsLoading && !HasError && !IsEmpty ? Visibility.Visible : Visibility.Collapsed;
    public Visibility IssuesVisibility => Issues.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
    public bool HasCanonicalProposals => ProposalSummaries.Count > 0;
    public Brush ReadinessBackgroundBrush => ReadinessStatus switch
    {
        PlanDesignCanonicalReadinessStatuses.Ready => new SolidColorBrush(Color.FromRgb(236, 253, 245)),
        PlanDesignCanonicalReadinessStatuses.NeedsReview => new SolidColorBrush(Color.FromRgb(255, 247, 237)),
        PlanDesignCanonicalReadinessStatuses.Partial => new SolidColorBrush(Color.FromRgb(255, 251, 235)),
        _ => new SolidColorBrush(Color.FromRgb(248, 250, 252))
    };

    public Brush ReadinessForegroundBrush => ReadinessStatus switch
    {
        PlanDesignCanonicalReadinessStatuses.Ready => new SolidColorBrush(Color.FromRgb(4, 120, 87)),
        PlanDesignCanonicalReadinessStatuses.NeedsReview => new SolidColorBrush(Color.FromRgb(194, 65, 12)),
        PlanDesignCanonicalReadinessStatuses.Partial => new SolidColorBrush(Color.FromRgb(161, 98, 7)),
        _ => new SolidColorBrush(Color.FromRgb(71, 85, 105))
    };

    private void NotifyStateVisibility()
    {
        OnPropertyChanged(nameof(LoadingVisibility));
        OnPropertyChanged(nameof(ErrorVisibility));
        OnPropertyChanged(nameof(EmptyVisibility));
        OnPropertyChanged(nameof(ContentVisibility));
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

public sealed class ProposalQuoteSummaryDto
{
    public long ProposalId { get; init; }
    public string CarrierName { get; init; } = "";
    public string DisplayName { get; init; } = "";
    public string AnnualPremiumText { get; init; } = "N/A";
    public string MonthlyPremiumText { get; init; } = "N/A";
    public string DifferenceFromLowestText { get; init; } = "N/A";
    public string ReadinessStatus { get; init; } = ProposalCanonicalReadinessStatuses.NotReady;
    public string FinancialStatusText { get; init; } = "Missing";
    public string PlanDesignStatusText { get; init; } = "Missing";
    public string ValidationStatusText { get; init; } = "Missing";
    public string ReviewStatusText { get; init; } = "No review items";
    public string SourceText { get; init; } = "No source evidence";
    public bool IsLowestAnnualPremium { get; init; }
    public string LowestPremiumLabel => IsLowestAnnualPremium ? "Lowest saved premium" : "";
    public Brush ReadinessBackgroundBrush => ComparisonBrushes.ReadinessBackground(ReadinessStatus);
    public Brush ReadinessForegroundBrush => ComparisonBrushes.ReadinessForeground(ReadinessStatus);
}

public sealed class ProposalCanonicalComparisonRowDto
{
    public string Category { get; init; } = "";
    public string Subject { get; init; } = "";
    public string Detail { get; init; } = "";
    public string ComparisonStatus { get; init; } = ProposalIntelligenceComparisonStatuses.NotComparable;
    public string MissingText { get; init; } = "";
    public ObservableCollection<ProposalCanonicalComparisonValueDto> Values { get; } = [];

    public Brush StatusBackgroundBrush => ComparisonBrushes.Background(ComparisonStatus);
    public Brush StatusForegroundBrush => ComparisonBrushes.Foreground(ComparisonStatus);
}

public sealed class ProposalCanonicalComparisonValueDto
{
    public long ProposalId { get; init; }
    public string CarrierName { get; init; } = "";
    public string Value { get; init; } = "";
    public string Detail { get; init; } = "";
    public string SourceText { get; init; } = "No source evidence";
}

public sealed class BenefitComparisonDto
{
    public string BenefitCategory { get; init; } = "";
    public string BenefitName { get; init; } = "";
    public string ProvisionName { get; init; } = "";
    public string ClassName { get; init; } = "";
    public string CurrentValue { get; init; } = "";
    public string CurrentSourceText { get; init; } = "";
    public string ComparisonStatus { get; init; } = CurrentVsQuotePlanDesignStatuses.Unavailable;
    public ObservableCollection<QuoteComparisonValueDto> QuoteValues { get; } = [];

    public Brush StatusBackgroundBrush => ComparisonBrushes.Background(ComparisonStatus);
    public Brush StatusForegroundBrush => ComparisonBrushes.Foreground(ComparisonStatus);
}

public sealed class QuoteComparisonValueDto
{
    public string CarrierName { get; init; } = "";
    public string Reference { get; init; } = "";
    public string Value { get; init; } = "";
    public string Status { get; init; } = CurrentVsQuotePlanDesignStatuses.Unavailable;
    public string SourceText { get; init; } = "";

    public Brush BackgroundBrush => ComparisonBrushes.Background(Status);
    public Brush ForegroundBrush => ComparisonBrushes.Foreground(Status);
}

public sealed class PlanDesignComparisonIssueDto
{
    public string Code { get; init; } = "";
    public string Severity { get; init; } = "";
    public string Category { get; init; } = "";
    public string Carrier { get; init; } = "";
    public string Message { get; init; } = "";
    public string Context { get; init; } = "";
}

internal static class ComparisonBrushes
{
    public static Brush Background(string status)
    {
        return status switch
        {
            CurrentVsQuotePlanDesignStatuses.Same => new SolidColorBrush(Color.FromRgb(220, 252, 231)),
            CurrentVsQuotePlanDesignStatuses.Different => new SolidColorBrush(Color.FromRgb(254, 249, 195)),
            CurrentVsQuotePlanDesignStatuses.NeedsReview => new SolidColorBrush(Color.FromRgb(255, 237, 213)),
            CurrentVsQuotePlanDesignStatuses.CurrentOnly => new SolidColorBrush(Color.FromRgb(239, 246, 255)),
            CurrentVsQuotePlanDesignStatuses.QuoteOnly => new SolidColorBrush(Color.FromRgb(240, 253, 250)),
            ProposalIntelligenceComparisonStatuses.Incomplete => new SolidColorBrush(Color.FromRgb(239, 246, 255)),
            _ => new SolidColorBrush(Color.FromRgb(248, 250, 252))
        };
    }

    public static Brush Foreground(string status)
    {
        return status switch
        {
            CurrentVsQuotePlanDesignStatuses.Same => new SolidColorBrush(Color.FromRgb(21, 128, 61)),
            CurrentVsQuotePlanDesignStatuses.Different => new SolidColorBrush(Color.FromRgb(161, 98, 7)),
            CurrentVsQuotePlanDesignStatuses.NeedsReview => new SolidColorBrush(Color.FromRgb(194, 65, 12)),
            CurrentVsQuotePlanDesignStatuses.CurrentOnly => new SolidColorBrush(Color.FromRgb(29, 78, 216)),
            CurrentVsQuotePlanDesignStatuses.QuoteOnly => new SolidColorBrush(Color.FromRgb(15, 118, 110)),
            ProposalIntelligenceComparisonStatuses.Incomplete => new SolidColorBrush(Color.FromRgb(29, 78, 216)),
            _ => new SolidColorBrush(Color.FromRgb(71, 85, 105))
        };
    }

    public static Brush ReadinessBackground(string status)
    {
        return status switch
        {
            ProposalCanonicalReadinessStatuses.Ready => new SolidColorBrush(Color.FromRgb(236, 253, 245)),
            ProposalCanonicalReadinessStatuses.NeedsReview => new SolidColorBrush(Color.FromRgb(255, 237, 213)),
            ProposalCanonicalReadinessStatuses.Partial => new SolidColorBrush(Color.FromRgb(255, 251, 235)),
            _ => new SolidColorBrush(Color.FromRgb(248, 250, 252))
        };
    }

    public static Brush ReadinessForeground(string status)
    {
        return status switch
        {
            ProposalCanonicalReadinessStatuses.Ready => new SolidColorBrush(Color.FromRgb(4, 120, 87)),
            ProposalCanonicalReadinessStatuses.NeedsReview => new SolidColorBrush(Color.FromRgb(194, 65, 12)),
            ProposalCanonicalReadinessStatuses.Partial => new SolidColorBrush(Color.FromRgb(161, 98, 7)),
            _ => new SolidColorBrush(Color.FromRgb(71, 85, 105))
        };
    }
}

using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using System.Windows.Media;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.ViewModels;

public sealed class ClaimsRangeReviewViewModel : INotifyPropertyChanged
{
    private static readonly Brush ReadyBackground = FrozenBrush(220, 246, 232);
    private static readonly Brush ReadyForeground = FrozenBrush(8, 122, 102);
    private static readonly Brush ReviewBackground = FrozenBrush(255, 243, 205);
    private static readonly Brush ReviewForeground = FrozenBrush(110, 83, 0);
    private readonly DelegateCommand _confirmIdentityCommand;
    private ClaimsRangeComparison? _comparison;
    private string _availabilityText = "Validate a current and prior Claims Range report to populate this review.";
    private string _statusText = ClaimsRangeComparisonStatuses.NeedsReview;
    private Brush _statusBackground = ReviewBackground;
    private Brush _statusForeground = ReviewForeground;
    private string _priorClientName = "Not available";
    private string _priorPolicyNumber = "Not available";
    private string _priorPeriodText = "Not available";
    private string _currentClientName = "Not available";
    private string _currentPolicyNumber = "Not available";
    private string _currentPeriodText = "Not available";
    private string _consultantTakeaway = "";
    private bool _identityConfirmationRequired;
    private bool _identityConfirmedForSession;

    public ClaimsRangeReviewViewModel()
    {
        _confirmIdentityCommand = new DelegateCommand(ConfirmIdentityForSession, () => CanConfirmIdentity);
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    public ObservableCollection<ClaimsRangeReviewRowViewModel> Rows { get; } = [];
    public ObservableCollection<string> ValidationIssues { get; } = [];
    public ICommand ConfirmIdentityCommand => _confirmIdentityCommand;

    public string AvailabilityText
    {
        get => _availabilityText;
        private set => SetField(ref _availabilityText, value);
    }

    public string StatusText
    {
        get => _statusText;
        private set => SetField(ref _statusText, value);
    }

    public Brush StatusBackground
    {
        get => _statusBackground;
        private set => SetField(ref _statusBackground, value);
    }

    public Brush StatusForeground
    {
        get => _statusForeground;
        private set => SetField(ref _statusForeground, value);
    }

    public string PriorClientName
    {
        get => _priorClientName;
        private set => SetField(ref _priorClientName, value);
    }

    public string PriorPolicyNumber
    {
        get => _priorPolicyNumber;
        private set => SetField(ref _priorPolicyNumber, value);
    }

    public string PriorPeriodText
    {
        get => _priorPeriodText;
        private set => SetField(ref _priorPeriodText, value);
    }

    public string CurrentClientName
    {
        get => _currentClientName;
        private set => SetField(ref _currentClientName, value);
    }

    public string CurrentPolicyNumber
    {
        get => _currentPolicyNumber;
        private set => SetField(ref _currentPolicyNumber, value);
    }

    public string CurrentPeriodText
    {
        get => _currentPeriodText;
        private set => SetField(ref _currentPeriodText, value);
    }

    public string ConsultantTakeaway
    {
        get => _consultantTakeaway;
        set => SetField(ref _consultantTakeaway, value ?? "");
    }

    public bool IdentityConfirmationRequired
    {
        get => _identityConfirmationRequired;
        private set
        {
            if (SetField(ref _identityConfirmationRequired, value))
            {
                OnPropertyChanged(nameof(CanConfirmIdentity));
                _confirmIdentityCommand.RaiseCanExecuteChanged();
            }
        }
    }

    public bool IdentityConfirmedForSession
    {
        get => _identityConfirmedForSession;
        private set
        {
            if (SetField(ref _identityConfirmedForSession, value))
            {
                OnPropertyChanged(nameof(CanConfirmIdentity));
                _confirmIdentityCommand.RaiseCanExecuteChanged();
            }
        }
    }

    public bool CanConfirmIdentity => IdentityConfirmationRequired && !IdentityConfirmedForSession;

    public void Load(ClaimsRangeComparison? comparison)
    {
        if (!ReferenceEquals(_comparison, comparison))
        {
            ClearSessionState(_comparison);
        }

        _comparison = comparison;
        Rows.Clear();
        ValidationIssues.Clear();

        if (comparison is null)
        {
            PriorClientName = "Not available";
            PriorPolicyNumber = "Not available";
            PriorPeriodText = "Not available";
            CurrentClientName = "Not available";
            CurrentPolicyNumber = "Not available";
            CurrentPeriodText = "Not available";
            ConsultantTakeaway = "";
            IdentityConfirmationRequired = false;
            IdentityConfirmedForSession = false;
            AvailabilityText = "No Claims Range comparison has been validated. Select the current and prior reports, then click Validate.";
            StatusText = ClaimsRangeComparisonStatuses.NeedsReview;
            StatusBackground = ReviewBackground;
            StatusForeground = ReviewForeground;
            ValidationIssues.Add("Awaiting current and prior Claims Range report validation.");
            return;
        }

        PriorClientName = Display(comparison.PriorReport?.ClientName);
        PriorPolicyNumber = Display(comparison.PriorReport?.PolicyNumber);
        PriorPeriodText = Display(comparison.PriorReport?.PeriodLabel);
        CurrentClientName = Display(comparison.CurrentReport?.ClientName);
        CurrentPolicyNumber = Display(comparison.CurrentReport?.PolicyNumber);
        CurrentPeriodText = Display(comparison.CurrentReport?.PeriodLabel);
        ConsultantTakeaway = comparison.EffectiveTakeaway;
        IdentityConfirmationRequired = comparison.IdentityConfirmationRequired;
        IdentityConfirmedForSession = comparison.IdentityConfirmedForSession;

        AddReportRows("Prior", comparison.PriorReport);
        AddReportRows("Current", comparison.CurrentReport);
        RefreshStatusAndValidation();
    }

    public void ResetSession() => Load(null);

    public void RevokeSessionEdits()
    {
        if (_comparison is null)
        {
            return;
        }

        ClearSessionState(_comparison);
        ConsultantTakeaway = _comparison.GeneratedTakeaway;
        IdentityConfirmationRequired = _comparison.IdentityConfirmationRequired;
        IdentityConfirmedForSession = false;
        RefreshStatusAndValidation();
    }

    public void ApplySessionEdits()
    {
        if (_comparison is null)
        {
            return;
        }

        var edited = ConsultantTakeaway.Trim();
        var generated = _comparison.GeneratedTakeaway.Trim();
        _comparison.TakeawayOverride = string.IsNullOrWhiteSpace(edited) ||
                                       edited.Equals(generated, StringComparison.Ordinal)
            ? null
            : edited;
    }

    private void AddReportRows(string periodRole, ClaimsRangeReport? report)
    {
        if (report is null)
        {
            return;
        }

        foreach (var row in report.Rows
                     .OrderBy(row => BenefitSortOrder(row.Benefit))
                     .ThenBy(row => row.SortOrder))
        {
            Rows.Add(new ClaimsRangeReviewRowViewModel(
                periodRole,
                report.PeriodLabel,
                BenefitDisplayName(row.Benefit),
                row.RangeLabel,
                row.CertificateCount,
                row.SourcePage,
                report.SourceFileName));
        }
    }

    private void ConfirmIdentityForSession()
    {
        if (_comparison is null || !CanConfirmIdentity)
        {
            return;
        }

        _comparison.ConfirmIdentityForSession(
            Environment.UserName,
            "Confirmed in the Claims Range Renewal Review for this export session.");
        IdentityConfirmedForSession = _comparison.IdentityConfirmedForSession;
        RefreshStatusAndValidation();
    }

    private static void ClearSessionState(ClaimsRangeComparison? comparison)
    {
        if (comparison is null)
        {
            return;
        }

        comparison.ClearIdentityConfirmation();
        comparison.TakeawayOverride = null;
    }

    private void RefreshStatusAndValidation()
    {
        if (_comparison is null)
        {
            return;
        }

        StatusText = _comparison.Status;
        var ready = _comparison.CanUseForClientOutput;
        StatusBackground = ready ? ReadyBackground : ReviewBackground;
        StatusForeground = ready ? ReadyForeground : ReviewForeground;
        AvailabilityText = ready
            ? $"Claims Range comparison validated with {Rows.Count} source rows. Commentary edits below will be used for this export."
            : $"Claims Range data found with {_comparison.ValidationIssues.Count} item(s) requiring review before client output.";

        ValidationIssues.Clear();
        foreach (var issue in _comparison.ValidationIssues)
        {
            ValidationIssues.Add(_comparison.IdentityConfirmedForSession && issue.CanConfirmForSession
                ? $"Confirmed for this session: {issue.Message}"
                : $"{issue.Severity}: {issue.Message}");
        }

        if (_comparison.IdentityConfirmedForSession)
        {
            var confirmedBy = Display(_comparison.IdentityConfirmedBy);
            ValidationIssues.Add($"Confirmed for this session by {confirmedBy}: the two report identities are intentionally paired.");
        }

        if (ValidationIssues.Count == 0)
        {
            ValidationIssues.Add("Validated: both reports contain complete Health and Dental claims-range rows.");
        }
    }

    private static int BenefitSortOrder(string benefit) =>
        benefit.Equals(ClaimsRangeBenefitKeys.Health, StringComparison.OrdinalIgnoreCase) ? 0 : 1;

    private static string BenefitDisplayName(string benefit) =>
        benefit.Equals(ClaimsRangeBenefitKeys.Health, StringComparison.OrdinalIgnoreCase)
            ? "Extended Health Care"
            : benefit;

    private static string Display(string? value) =>
        string.IsNullOrWhiteSpace(value) ? "Not available" : value.Trim();

    private static Brush FrozenBrush(byte red, byte green, byte blue)
    {
        var brush = new SolidColorBrush(Color.FromRgb(red, green, blue));
        brush.Freeze();
        return brush;
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

    private sealed class DelegateCommand(Action execute, Func<bool> canExecute) : ICommand
    {
        public event EventHandler? CanExecuteChanged;

        public bool CanExecute(object? parameter) => canExecute();

        public void Execute(object? parameter) => execute();

        public void RaiseCanExecuteChanged() => CanExecuteChanged?.Invoke(this, EventArgs.Empty);
    }
}

public sealed record ClaimsRangeReviewRowViewModel(
    string PeriodRole,
    string PeriodText,
    string Benefit,
    string BandLabel,
    int CertificateCount,
    int SourcePage,
    string SourceFile);

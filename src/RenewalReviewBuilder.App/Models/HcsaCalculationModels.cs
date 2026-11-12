using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace RenewalReviewBuilder.App.Models;

/// <summary>
/// Manual HCSA funding assumptions entered for a renewal. HCSA is funded as an
/// annual employer credit per eligible employee, plus the carrier's administration fee.
/// </summary>
public sealed class HcsaCalculation
{
    public List<HcsaClassCalculation> Classes { get; set; } = [];

    public bool HasEntries => IncludedClasses().Any();

    public decimal CurrentAnnualCost => IncludedClasses().Sum(item => item.CurrentAnnualCost);

    public decimal RenewalAnnualCost => IncludedClasses().Sum(item => item.RenewalAnnualCost);

    public IEnumerable<HcsaClassCalculation> IncludedClasses() =>
        Classes.Where(item => item.IsIncluded && !IsLegacyCriticalIllnessEntry(item));

    public int RemoveLegacyCriticalIllnessEntries() =>
        Classes.RemoveAll(IsLegacyCriticalIllnessEntry);

    public HcsaCalculation Clone() => new()
    {
        Classes = Classes.Select(item => item.Clone()).ToList()
    };

    private static bool IsLegacyCriticalIllnessEntry(HcsaClassCalculation item)
    {
        var name = (item.ClassName ?? "").Trim();
        return name.Contains("critical illness", StringComparison.OrdinalIgnoreCase) ||
               name.Equals("CI", StringComparison.OrdinalIgnoreCase) ||
               name.StartsWith("CI ", StringComparison.OrdinalIgnoreCase) ||
               name.StartsWith("CI-", StringComparison.OrdinalIgnoreCase);
    }
}

public sealed class HcsaClassCalculation : INotifyPropertyChanged
{
    private string _className = "";
    private int _eligibleEmployees;
    private decimal _currentAnnualCredit;
    private decimal _renewalAnnualCredit;
    private decimal _usagePercent = 100m;
    private decimal _administrationFeePercent = 10m;

    public event PropertyChangedEventHandler? PropertyChanged;

    public string ClassName
    {
        get => _className;
        set => SetField(ref _className, value);
    }

    public int EligibleEmployees
    {
        get => _eligibleEmployees;
        set
        {
            if (SetField(ref _eligibleEmployees, value))
            {
                NotifyCalculatedValuesChanged();
            }
        }
    }

    public decimal CurrentAnnualCredit
    {
        get => _currentAnnualCredit;
        set
        {
            if (SetField(ref _currentAnnualCredit, value))
            {
                NotifyCalculatedValuesChanged();
            }
        }
    }

    public decimal RenewalAnnualCredit
    {
        get => _renewalAnnualCredit;
        set
        {
            if (SetField(ref _renewalAnnualCredit, value))
            {
                NotifyCalculatedValuesChanged();
            }
        }
    }

    public decimal UsagePercent
    {
        get => _usagePercent;
        set
        {
            if (SetField(ref _usagePercent, value))
            {
                NotifyCalculatedValuesChanged();
            }
        }
    }

    public decimal AdministrationFeePercent
    {
        get => _administrationFeePercent;
        set
        {
            if (SetField(ref _administrationFeePercent, value))
            {
                NotifyCalculatedValuesChanged();
            }
        }
    }

    public bool IsIncluded => EligibleEmployees > 0 &&
                              (CurrentAnnualCredit > 0m || RenewalAnnualCredit > 0m);

    public decimal CurrentAnnualCost => CalculateAnnualCost(CurrentAnnualCredit);

    public decimal RenewalAnnualCost => CalculateAnnualCost(RenewalAnnualCredit);

    public decimal CurrentMonthlyCost => CurrentAnnualCost / 12m;

    public decimal RenewalMonthlyCost => RenewalAnnualCost / 12m;

    public HcsaClassCalculation Clone() => new()
    {
        ClassName = ClassName,
        EligibleEmployees = EligibleEmployees,
        CurrentAnnualCredit = CurrentAnnualCredit,
        RenewalAnnualCredit = RenewalAnnualCredit,
        UsagePercent = UsagePercent,
        AdministrationFeePercent = AdministrationFeePercent
    };

    private decimal CalculateAnnualCost(decimal annualCredit) =>
        decimal.Round(
            EligibleEmployees * annualCredit * (UsagePercent / 100m) * (1m + (AdministrationFeePercent / 100m)),
            2,
            MidpointRounding.AwayFromZero);

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

    private void NotifyCalculatedValuesChanged()
    {
        OnPropertyChanged(nameof(IsIncluded));
        OnPropertyChanged(nameof(CurrentAnnualCost));
        OnPropertyChanged(nameof(RenewalAnnualCost));
        OnPropertyChanged(nameof(CurrentMonthlyCost));
        OnPropertyChanged(nameof(RenewalMonthlyCost));
    }

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
}

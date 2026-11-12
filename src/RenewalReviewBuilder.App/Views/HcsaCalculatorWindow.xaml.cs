using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Views;

public partial class HcsaCalculatorWindow : Window
{
    public ObservableCollection<HcsaClassCalculation> Classes { get; }

    public HcsaCalculation Calculation { get; private set; } = new();

    public HcsaCalculatorWindow(HcsaCalculation? existingCalculation)
    {
        Classes = new ObservableCollection<HcsaClassCalculation>(
            existingCalculation?.Classes.Select(item => item.Clone()) ?? []);
        InitializeComponent();
        DataContext = this;

        if (Classes.Count == 0)
        {
            Classes.Add(new HcsaClassCalculation { ClassName = "Class 1" });
        }

        foreach (var item in Classes)
        {
            SubscribeToChanges(item);
        }

        UpdateTotals();
    }

    private void AddClass_Click(object sender, RoutedEventArgs e)
    {
        var item = new HcsaClassCalculation { ClassName = $"Class {Classes.Count + 1}" };
        SubscribeToChanges(item);
        Classes.Add(item);
        ClassesGrid.SelectedItem = item;
    }

    private void RemoveSelected_Click(object sender, RoutedEventArgs e)
    {
        if (ClassesGrid.SelectedItem is HcsaClassCalculation selected)
        {
            selected.PropertyChanged -= Class_PropertyChanged;
            Classes.Remove(selected);
            UpdateTotals();
        }
    }

    private void Apply_Click(object sender, RoutedEventArgs e)
    {
        ClassesGrid.CommitEdit();
        ClassesGrid.CommitEdit();

        var invalidRow = Classes.FirstOrDefault(item =>
            item.EligibleEmployees < 0 ||
            item.CurrentAnnualCredit < 0m ||
            item.RenewalAnnualCredit < 0m ||
            item.UsagePercent < 0m || item.UsagePercent > 100m ||
            item.AdministrationFeePercent < 0m);
        if (invalidRow is not null)
        {
            MessageBox.Show(
                this,
                "Eligible employees, annual credits, and administration fees cannot be negative. Usage must be from 0% to 100%.",
                "Check HCSA inputs",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        Calculation = new HcsaCalculation { Classes = Classes.Select(item => item.Clone()).ToList() };
        DialogResult = true;
    }

    private void SubscribeToChanges(HcsaClassCalculation item)
    {
        item.PropertyChanged -= Class_PropertyChanged;
        item.PropertyChanged += Class_PropertyChanged;
    }

    private void Class_PropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        UpdateTotals();
    }

    protected override void OnActivated(EventArgs e)
    {
        base.OnActivated(e);
        UpdateTotals();
    }

    private void UpdateTotals()
    {
        CurrentAnnualTotalTextBlock.Text = Classes.Sum(item => item.CurrentAnnualCost).ToString("C2");
        RenewalAnnualTotalTextBlock.Text = Classes.Sum(item => item.RenewalAnnualCost).ToString("C2");
    }

    protected override void OnClosed(EventArgs e)
    {
        foreach (var item in Classes)
        {
            item.PropertyChanged -= Class_PropertyChanged;
        }

        base.OnClosed(e);
    }
}

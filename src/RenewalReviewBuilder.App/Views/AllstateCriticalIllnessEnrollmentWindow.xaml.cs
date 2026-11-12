using System.Globalization;
using System.Windows;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Views;

public partial class AllstateCriticalIllnessEnrollmentWindow : Window
{
    public AllstateCriticalIllnessEnrollment Enrollment { get; private set; }

    public AllstateCriticalIllnessEnrollmentWindow(AllstateCriticalIllnessEnrollment enrollment)
    {
        InitializeComponent();
        Enrollment = enrollment.Clone();
        EmployeeOnlyCountTextBox.Text = Enrollment.EmployeeOnlyCount.ToString(CultureInfo.InvariantCulture);
        FamilyCountTextBox.Text = Enrollment.FamilyCount.ToString(CultureInfo.InvariantCulture);
    }

    private void Apply_Click(object sender, RoutedEventArgs e)
    {
        if (!TryReadCount(EmployeeOnlyCountTextBox.Text, "Employee-only", out var employeeOnly) ||
            !TryReadCount(FamilyCountTextBox.Text, "Family", out var family))
        {
            return;
        }

        if (employeeOnly + family == 0)
        {
            MessageBox.Show(this, "Enter at least one enrolled employee, or choose Clear to remove the counts.", "Enrollment Counts", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        Enrollment = new AllstateCriticalIllnessEnrollment
        {
            EmployeeOnlyCount = employeeOnly,
            FamilyCount = family
        };
        DialogResult = true;
    }

    private void Clear_Click(object sender, RoutedEventArgs e)
    {
        Enrollment = new AllstateCriticalIllnessEnrollment();
        DialogResult = true;
    }

    private bool TryReadCount(string text, string label, out int value)
    {
        if (int.TryParse(text.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out value) && value >= 0)
        {
            return true;
        }

        MessageBox.Show(this, $"{label} enrollment must be a whole number of zero or more.", "Enrollment Counts", MessageBoxButton.OK, MessageBoxImage.Warning);
        return false;
    }
}

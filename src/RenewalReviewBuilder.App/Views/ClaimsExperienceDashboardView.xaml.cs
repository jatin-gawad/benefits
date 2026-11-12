using System.Windows;
using System.Windows.Controls;

namespace RenewalReviewBuilder.App.Views;

public partial class ClaimsExperienceDashboardView : UserControl
{
    public ClaimsExperienceDashboardView()
    {
        InitializeComponent();
    }

    private void Export_Click(object sender, RoutedEventArgs e)
    {
        MessageBox.Show(
            "Claims dashboard export will connect to the existing Claims Model and Claims PowerPoint export workflow. For now, use the bottom Export Claims Model or Export Claims PPT buttons.",
            "Claims Experience Dashboard",
            MessageBoxButton.OK,
            MessageBoxImage.Information);
    }
}

using System.Threading;
using System.Windows.Controls;
using RenewalReviewBuilder.App.Views;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class HcsaCalculatorWindowTests
{
    [Fact]
    public void FundingGrid_Overrides_GlobalReadOnlyStyle_ForInputColumns()
    {
        Exception? failure = null;
        var completed = false;
        var thread = new Thread(() =>
        {
            try
            {
                var window = new HcsaCalculatorWindow(null);
                var grid = Assert.IsType<DataGrid>(window.FindName("ClassesGrid"));

                Assert.False(grid.IsReadOnly);
                Assert.False(grid.CanUserSortColumns);
                Assert.False(grid.CanUserReorderColumns);
                Assert.All(grid.Columns.Take(6), column => Assert.False(column.IsReadOnly));
                Assert.All(grid.Columns.Skip(6), column => Assert.True(column.IsReadOnly));
                completed = true;
                window.Close();
            }
            catch (Exception exception)
            {
                failure = exception;
            }
        });
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();

        Assert.True(thread.Join(TimeSpan.FromSeconds(20)), "HCSA calculator smoke test timed out.");
        Assert.Null(failure);
        Assert.True(completed);
    }

    [Fact]
    public void FundingGrid_RecalculatesCostsAndTotals_WithoutRefreshingTheCollection()
    {
        Exception? failure = null;
        var thread = new Thread(() =>
        {
            try
            {
                var window = new HcsaCalculatorWindow(null);
                var grid = Assert.IsType<DataGrid>(window.FindName("ClassesGrid"));
                var currentTotal = Assert.IsType<TextBlock>(window.FindName("CurrentAnnualTotalTextBlock"));
                var renewalTotal = Assert.IsType<TextBlock>(window.FindName("RenewalAnnualTotalTextBlock"));
                var row = Assert.Single(window.Classes);

                row.EligibleEmployees = 10;
                row.CurrentAnnualCredit = 500m;
                row.RenewalAnnualCredit = 600m;
                row.UsagePercent = 80m;
                row.AdministrationFeePercent = 10m;

                Assert.Equal(4400m, row.CurrentAnnualCost);
                Assert.Equal(5280m, row.RenewalAnnualCost);
                Assert.Equal(4400m.ToString("C2"), currentTotal.Text);
                Assert.Equal(5280m.ToString("C2"), renewalTotal.Text);
                Assert.True(grid.CommitEdit(DataGridEditingUnit.Cell, true));
                window.Close();
            }
            catch (Exception exception)
            {
                failure = exception;
            }
        });
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();

        Assert.True(thread.Join(TimeSpan.FromSeconds(20)), "HCSA calculator edit test timed out.");
        Assert.Null(failure);
    }
}

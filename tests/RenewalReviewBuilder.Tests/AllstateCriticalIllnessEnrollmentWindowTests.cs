using System;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Views;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class AllstateCriticalIllnessEnrollmentWindowTests
{
    [Fact]
    public void EnrollmentWindow_Loads_Both_Tier_Counts_And_Arranges_Without_Layout_Errors()
    {
        Exception? failure = null;
        var completed = false;
        var thread = new Thread(() =>
        {
            try
            {
                var window = new AllstateCriticalIllnessEnrollmentWindow(new AllstateCriticalIllnessEnrollment
                {
                    EmployeeOnlyCount = 10,
                    FamilyCount = 6
                });

                var employee = Assert.IsType<TextBox>(window.FindName("EmployeeOnlyCountTextBox"));
                var family = Assert.IsType<TextBox>(window.FindName("FamilyCountTextBox"));
                Assert.Equal("10", employee.Text);
                Assert.Equal("6", family.Text);

                window.ShowInTaskbar = false;
                window.Opacity = 0;
                window.Show();
                window.UpdateLayout();
                Assert.True(employee.ActualWidth > 0);
                Assert.True(family.ActualWidth > 0);
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

        Assert.True(thread.Join(TimeSpan.FromSeconds(20)), "Allstate CI enrollment window smoke test timed out.");
        Assert.Null(failure);
        Assert.True(completed);
    }
}

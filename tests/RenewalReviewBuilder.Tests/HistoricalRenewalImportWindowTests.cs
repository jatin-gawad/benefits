using System.Reflection;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using RenewalReviewBuilder.App.Views;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class HistoricalRenewalImportWindowTests
{
    [Fact]
    public void Repeated_Render_Review_And_Back_Reuses_Inputs_Without_Logical_Child_Errors()
    {
        Exception? failure = null;
        var completed = false;
        var thread = new Thread(() =>
        {
            HistoricalRenewalImportWindow? window = null;
            try
            {
                window = new HistoricalRenewalImportWindow(42);
                var windowType = typeof(HistoricalRenewalImportWindow);
                var renderStep = PrivateMethod(windowType, "RenderStep");
                var moveToStep = PrivateMethod(windowType, "MoveToStep");
                var backClick = PrivateMethod(windowType, "Back_Click");

                renderStep.Invoke(window, null);
                AssertSharedInputsHaveSinglePanelParent(window, "_pdfPath", "_workbookPath");

                moveToStep.Invoke(window, [1]);
                renderStep.Invoke(window, null);
                AssertSharedInputsHaveSinglePanelParent(
                    window,
                    "_carrier",
                    "_effectiveDate",
                    "_renewalYear",
                    "_requestedIncrease",
                    "_finalAcceptedIncrease",
                    "_annualSavings");

                var carrier = PrivateField<TextBox>(window, "_carrier");
                var effectiveDate = PrivateField<TextBox>(window, "_effectiveDate");
                var renewalYear = PrivateField<TextBox>(window, "_renewalYear");
                var requestedIncrease = PrivateField<TextBox>(window, "_requestedIncrease");
                var finalAcceptedIncrease = PrivateField<TextBox>(window, "_finalAcceptedIncrease");
                var annualSavings = PrivateField<TextBox>(window, "_annualSavings");
                carrier.Text = "Sun Life";
                effectiveDate.Text = "2023-10-01";
                renewalYear.Text = "2023";
                requestedIncrease.Text = "12.5";
                finalAcceptedIncrease.Text = "7.0";
                annualSavings.Text = "25000";

                backClick.Invoke(window, [window, new RoutedEventArgs()]);
                Assert.Equal(0, PrivateField<int>(window, "_step"));
                AssertSharedInputsHaveSinglePanelParent(window, "_pdfPath", "_workbookPath");

                moveToStep.Invoke(window, [1]);
                Assert.Equal("Sun Life", carrier.Text);
                Assert.Equal("2023-10-01", effectiveDate.Text);
                Assert.Equal("2023", renewalYear.Text);
                Assert.Equal("12.5", requestedIncrease.Text);
                Assert.Equal("7.0", finalAcceptedIncrease.Text);
                Assert.Equal("25000", annualSavings.Text);
                AssertSharedInputsHaveSinglePanelParent(
                    window,
                    "_carrier",
                    "_effectiveDate",
                    "_renewalYear",
                    "_requestedIncrease",
                    "_finalAcceptedIncrease",
                    "_annualSavings");

                moveToStep.Invoke(window, [3]);
                renderStep.Invoke(window, null);
                AssertSinglePanelParent(PrivateField<CheckBox>(window, "_confirm"));

                moveToStep.Invoke(window, [1]);
                moveToStep.Invoke(window, [0]);
                renderStep.Invoke(window, null);
                AssertSharedInputsHaveSinglePanelParent(window, "_pdfPath", "_workbookPath");
                completed = true;
            }
            catch (Exception exception)
            {
                failure = exception;
            }
            finally
            {
                window?.Close();
            }
        });
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();

        Assert.True(thread.Join(TimeSpan.FromSeconds(20)), "Historical renewal import render test timed out.");
        Assert.Null(failure);
        Assert.True(completed);
    }

    private static MethodInfo PrivateMethod(Type type, string name) =>
        type.GetMethod(name, BindingFlags.Instance | BindingFlags.NonPublic)
        ?? throw new MissingMethodException(type.FullName, name);

    private static T PrivateField<T>(object instance, string name) =>
        (T)(instance.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic)?.GetValue(instance)
            ?? throw new MissingFieldException(instance.GetType().FullName, name));

    private static void AssertSharedInputsHaveSinglePanelParent(object window, params string[] fieldNames)
    {
        foreach (var fieldName in fieldNames)
        {
            var input = PrivateField<TextBox>(window, fieldName);
            AssertSinglePanelParent(input);
        }
    }

    private static void AssertSinglePanelParent(FrameworkElement element) =>
        Assert.IsAssignableFrom<Panel>(LogicalTreeHelper.GetParent(element));
}

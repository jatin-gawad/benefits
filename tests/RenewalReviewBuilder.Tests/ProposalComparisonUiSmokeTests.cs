using System.Threading;
using System.Windows;
using System.Windows.Controls;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Views;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class ProposalComparisonUiSmokeTests
{
    [Fact]
    public void ProposalComparisonWorkspace_OpensCanonicalTabsInEmptyClientStateWithoutCrashing()
    {
        Exception? failure = null;
        var completed = false;
        var thread = new Thread(() =>
        {
            try
            {
                if (Application.Current is null)
                {
                    var application = new RenewalReviewBuilder.App.App();
                    application.InitializeComponent();
                }

                var view = new ProposalComparisonWorkspaceView();
                view.LoadClientAsync(null, "No Client Selected", "").GetAwaiter().GetResult();

                var model = Assert.IsType<ProposalComparisonViewModel>(view.DataContext);
                Assert.True(model.IsEmpty);
                Assert.Equal("No Client Selected", model.ClientName);
                Assert.DoesNotContain("mock", model.StatusText, StringComparison.OrdinalIgnoreCase);
                var tabs = Assert.IsType<TabControl>(view.FindName("ProposalComparisonTabs"));
                Assert.Equal(6, tabs.Items.Count);
                Assert.NotNull(view.FindName("ProposalSummaryGrid"));
                Assert.NotNull(view.FindName("ProposalPremiumGrid"));
                Assert.NotNull(view.FindName("ProposalRateGrid"));
                Assert.NotNull(view.FindName("ProposalCurrentVsQuoteGrid"));
                Assert.NotNull(view.FindName("ProposalCarrierPlanDesignGrid"));
                Assert.NotNull(view.FindName("ProposalUnderwritingGrid"));
                Assert.NotNull(view.FindName("ProposalFindingsItemsControl"));
                completed = true;
            }
            catch (Exception exception)
            {
                failure = exception;
            }
        });
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();

        Assert.True(thread.Join(TimeSpan.FromSeconds(20)), "Current vs Quote workspace smoke test timed out.");
        Assert.Null(failure);
        Assert.True(completed);
    }
}

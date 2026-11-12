using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Windows.Controls;
using System.Xml.Linq;
using RenewalReviewBuilder.App;
using RenewalReviewBuilder.App.Models;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class AssurePlusUiSmokeTests
{
    private static readonly IReadOnlyDictionary<string, string> RequiredControls =
        new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["AssurePlusAvailabilityTextBlock"] = "TextBlock",
            ["AssurePlusHistoryGrid"] = "DataGrid",
            ["AssurePlusHealthStatusTextBlock"] = "TextBlock",
            ["AssurePlusHealthTotalTextBlock"] = "TextBlock",
            ["AssurePlusHealthDetailTextBlock"] = "TextBlock",
            ["AssurePlusDentalStatusTextBlock"] = "TextBlock",
            ["AssurePlusDentalTotalTextBlock"] = "TextBlock",
            ["AssurePlusDentalDetailTextBlock"] = "TextBlock",
            ["AssurePlusDentalCommentaryTextBox"] = "TextBox",
            ["AssurePlusHealthCommentaryTextBox"] = "TextBox",
            ["AssurePlusTakeawayTextBox"] = "TextBox",
            ["AssurePlusPricingGrid"] = "DataGrid",
            ["AssurePlusDistributionGrid"] = "DataGrid",
            ["AssurePlusValidationItemsControl"] = "ItemsControl"
        };

    [Fact]
    public void MainWindow_Xaml_Contains_Visible_AssurePlus_Tab_And_Required_Review_Controls()
    {
        XNamespace presentation = "http://schemas.microsoft.com/winfx/2006/xaml/presentation";
        XNamespace xaml = "http://schemas.microsoft.com/winfx/2006/xaml";
        var path = Path.Combine(
            FindRepositoryRoot(),
            "src",
            "RenewalReviewBuilder.App",
            "MainWindow.xaml");
        var document = XDocument.Load(path, LoadOptions.SetLineInfo);

        var tab = Assert.Single(
            document.Descendants(presentation + "TabItem"),
            element => string.Equals(
                (string?)element.Attribute("Header"),
                "AssurePLUS",
                StringComparison.Ordinal));

        var tabControl = Assert.IsType<XElement>(tab.Parent);
        Assert.Equal(presentation + "TabControl", tabControl.Name);
        Assert.False(IsCollapsed(tab), "The AssurePLUS TabItem is collapsed in MainWindow.xaml.");
        Assert.False(IsCollapsed(tabControl), "The AssurePLUS TabControl is collapsed in MainWindow.xaml.");

        foreach (var (controlName, expectedType) in RequiredControls)
        {
            var control = Assert.Single(
                tab.Descendants(),
                element => string.Equals(
                    (string?)element.Attribute(xaml + "Name"),
                    controlName,
                    StringComparison.Ordinal));
            Assert.Equal(expectedType, control.Name.LocalName);
        }

        foreach (var commentaryName in new[]
                 {
                     "AssurePlusDentalCommentaryTextBox",
                     "AssurePlusHealthCommentaryTextBox",
                     "AssurePlusTakeawayTextBox"
                 })
        {
            var commentary = tab.Descendants().Single(element => string.Equals(
                (string?)element.Attribute(xaml + "Name"),
                commentaryName,
                StringComparison.Ordinal));
            Assert.Equal("True", (string?)commentary.Attribute("AcceptsReturn"));
            Assert.NotEqual("True", (string?)commentary.Attribute("IsReadOnly"));
        }
    }

    [Fact]
    public void MainWindow_Source_Wires_AssurePlus_Analysis_Reset_And_Commentary_Into_Export()
    {
        var path = Path.Combine(
            FindRepositoryRoot(),
            "src",
            "RenewalReviewBuilder.App",
            "MainWindow.xaml.cs");
        var source = File.ReadAllText(path);
        var updateReview = SliceBetween(
            source,
            "private void UpdateReview()",
            "private void UpdateAssurePlusReview()");
        var updateAssurePlus = SliceBetween(
            source,
            "private void UpdateAssurePlusReview()",
            "private static void SetAssurePlusOutcome(");
        var applyCommentary = SliceBetween(
            source,
            "private void ApplyAssurePlusCommentaryEdits()",
            "private static string FormatAssurePlusPercent(");
        var resetAssurePlus = SliceBetween(
            source,
            "private void ResetAssurePlusReview()",
            "private void ResetReviewDisplay()");
        var resetReview = SliceBetween(
            source,
            "private void ResetReviewDisplay()",
            "private void SelectBatchParentFolder_Click(");
        var export = SliceBetween(
            source,
            "private async void Export_Click(",
            "private async Task<ClaimsExperienceDataModelRequest> CreateClaimsDataModelRequestAsync(");

        Assert.Contains("UpdateAssurePlusReview();", updateReview, StringComparison.Ordinal);
        Assert.Contains("_analysis?.AssurePlusExperience", updateAssurePlus, StringComparison.Ordinal);
        foreach (var controlName in RequiredControls.Keys)
        {
            Assert.Contains(controlName, updateAssurePlus, StringComparison.Ordinal);
        }

        Assert.Contains("summary.DentalCommentary = AssurePlusDentalCommentaryTextBox.Text.Trim();", applyCommentary, StringComparison.Ordinal);
        Assert.Contains("summary.HealthCommentary = AssurePlusHealthCommentaryTextBox.Text.Trim();", applyCommentary, StringComparison.Ordinal);
        Assert.Contains("summary.Takeaway = AssurePlusTakeawayTextBox.Text.Trim();", applyCommentary, StringComparison.Ordinal);
        Assert.Contains("ResetAssurePlusReview();", resetReview, StringComparison.Ordinal);
        Assert.Contains("AssurePlusHistoryGrid.ItemsSource = null;", resetAssurePlus, StringComparison.Ordinal);
        Assert.Contains("AssurePlusPricingGrid.ItemsSource = null;", resetAssurePlus, StringComparison.Ordinal);
        Assert.Contains("AssurePlusDistributionGrid.ItemsSource = null;", resetAssurePlus, StringComparison.Ordinal);

        var commentaryIndex = export.IndexOf("ApplyAssurePlusCommentaryEdits();", StringComparison.Ordinal);
        var builderExportIndex = export.IndexOf("_builder.ExportAsync(_analysis", StringComparison.Ordinal);
        Assert.True(commentaryIndex >= 0, "Export_Click does not apply the consultant's AssurePLUS commentary edits.");
        Assert.True(builderExportIndex > commentaryIndex, "AssurePLUS commentary edits must be applied before the review pack is exported.");
    }

    [Fact]
    public void AssurePlus_Review_Projects_A_Populated_Analysis_And_Applies_Commentary_Edits()
    {
        Exception? failure = null;
        var completed = false;
        var thread = new Thread(() =>
        {
            try
            {
                var analysis = BuildPopulatedAnalysis();
                var window = (MainWindow)RuntimeHelpers.GetUninitializedObject(typeof(MainWindow));
                SetField(window, "_analysis", analysis);

                var availability = SetControl<TextBlock>(window, "AssurePlusAvailabilityTextBlock");
                var history = SetControl<DataGrid>(window, "AssurePlusHistoryGrid");
                var healthStatus = SetControl<TextBlock>(window, "AssurePlusHealthStatusTextBlock");
                var healthTotal = SetControl<TextBlock>(window, "AssurePlusHealthTotalTextBlock");
                var healthDetail = SetControl<TextBlock>(window, "AssurePlusHealthDetailTextBlock");
                var dentalStatus = SetControl<TextBlock>(window, "AssurePlusDentalStatusTextBlock");
                var dentalTotal = SetControl<TextBlock>(window, "AssurePlusDentalTotalTextBlock");
                var dentalDetail = SetControl<TextBlock>(window, "AssurePlusDentalDetailTextBlock");
                var dentalCommentary = SetControl<TextBox>(window, "AssurePlusDentalCommentaryTextBox");
                var healthCommentary = SetControl<TextBox>(window, "AssurePlusHealthCommentaryTextBox");
                var takeaway = SetControl<TextBox>(window, "AssurePlusTakeawayTextBox");
                var pricing = SetControl<DataGrid>(window, "AssurePlusPricingGrid");
                var distribution = SetControl<DataGrid>(window, "AssurePlusDistributionGrid");
                var validation = SetControl<ItemsControl>(window, "AssurePlusValidationItemsControl");

                InvokePrivate(window, "UpdateAssurePlusReview");

                var historyRows = Assert.IsAssignableFrom<IEnumerable>(history.ItemsSource).Cast<object>().ToList();
                Assert.Equal(3, historyRows.Count);
                Assert.Equal(2024, Property<int>(historyRows[0], "Year"));
                Assert.Equal("Performing", Property<string>(historyRows[0], "Health"));
                Assert.Equal("Performing", Property<string>(historyRows[0], "Dental"));
                Assert.Equal(2025, Property<int>(historyRows[1], "Year"));
                Assert.Equal("Top Performing", Property<string>(historyRows[1], "Dental"));
                Assert.Equal(2026, Property<int>(historyRows[2], "Year"));
                Assert.Equal("Validated", Property<string>(historyRows[2], "Validation"));

                Assert.Equal(4, Assert.IsAssignableFrom<IEnumerable>(pricing.ItemsSource).Cast<object>().Count());
                var distributionRows = Assert.IsAssignableFrom<IEnumerable>(distribution.ItemsSource).Cast<object>().ToList();
                Assert.Equal(20, distributionRows.Count);
                Assert.Equal(2025, Property<int>(distributionRows[0], "RenewalYear"));
                Assert.Equal("Health", Property<string>(distributionRows[0], "Benefit"));
                Assert.Equal("Drugs", Property<string>(distributionRows[0], "Category"));
                Assert.Equal("47.28%", Property<string>(distributionRows[0], "YourClaims"));
                Assert.Equal("47.22%", Property<string>(distributionRows[0], "BenchmarkClaims"));
                Assert.Equal("Performing", healthStatus.Text);
                Assert.Equal("+6.90%", healthTotal.Text);
                Assert.Equal("Standard change: 6.90% | Experience adjustment: 0.00%", healthDetail.Text);
                Assert.Equal("Top Performing", dentalStatus.Text);
                Assert.Equal("\u221210.10%", dentalTotal.Text);
                Assert.Equal("Standard change: 4.90% | Experience adjustment: \u221215.00%", dentalDetail.Text);
                Assert.Equal("Dental commentary from analysis", dentalCommentary.Text);
                Assert.Equal("Health commentary from analysis", healthCommentary.Text);
                Assert.Equal("Client-facing takeaway", takeaway.Text);
                Assert.True(dentalCommentary.IsEnabled);
                Assert.True(healthCommentary.IsEnabled);
                Assert.True(takeaway.IsEnabled);
                Assert.Contains("validated for 2024-2026", availability.Text, StringComparison.OrdinalIgnoreCase);
                Assert.Contains(
                    Assert.IsAssignableFrom<IEnumerable>(validation.ItemsSource).Cast<object>().Select(item => item.ToString()),
                    item => item?.Contains("Ready for workbook and PowerPoint export", StringComparison.Ordinal) == true);

                dentalCommentary.Text = "  Consultant dental edit  ";
                healthCommentary.Text = "  Consultant health edit  ";
                takeaway.Text = "  Consultant takeaway edit  ";
                InvokePrivate(window, "ApplyAssurePlusCommentaryEdits");

                Assert.Equal("Consultant dental edit", analysis.AssurePlusExperience!.DentalCommentary);
                Assert.Equal("Consultant health edit", analysis.AssurePlusExperience.HealthCommentary);
                Assert.Equal("Consultant takeaway edit", analysis.AssurePlusExperience.Takeaway);
                completed = true;
            }
            catch (Exception exception)
            {
                failure = exception;
            }
        });
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();

        Assert.True(thread.Join(TimeSpan.FromSeconds(20)), "AssurePLUS review projection smoke test timed out.");
        Assert.Null(failure);
        Assert.True(completed);
    }

    private static AnalysisResult BuildPopulatedAnalysis()
    {
        var summary = new AssurePlusExperienceSummary
        {
            EffectiveDate = new DateOnly(2026, 10, 1),
            HealthOutcome = Benefit("Health", "Performing", "Performing", 6.90m, 0m, 6.90m),
            DentalOutcome = Benefit("Dental", "Top Performing", "Top Performing", 4.90m, -15m, -10.10m),
            DentalCommentary = "Dental commentary from analysis",
            HealthCommentary = "Health commentary from analysis",
            Takeaway = "Client-facing takeaway"
        };
        AddRating(summary, 2024, "Health", "Performing");
        AddRating(summary, 2024, "Dental", "Performing");
        AddRating(summary, 2025, "Health", "Performing");
        AddRating(summary, 2025, "Dental", "Top Performing");
        AddRating(summary, 2026, "Health", "Performing");
        AddRating(summary, 2026, "Dental", "Top Performing");

        var analysis = new AnalysisResult { AssurePlusExperience = summary };
        analysis.Pdf.AssurePlusPricingReports.Add(Report(
            2025,
            Benefit("Health", "Performing", "Performing", 8.90m, 0m, 8.90m),
            Benefit("Dental", "Performing", "Top Performing", 6.90m, 0m, 6.90m)));
        analysis.Pdf.AssurePlusPricingReports.Add(Report(
            2026,
            Benefit("Health", "Performing", "Performing", 6.90m, 0m, 6.90m),
            Benefit("Dental", "Top Performing", "Top Performing", 4.90m, -15m, -10.10m)));
        return analysis;
    }

    private static AssurePlusPricingReport Report(
        int year,
        AssurePlusBenefitPricing health,
        AssurePlusBenefitPricing dental)
    {
        var report = new AssurePlusPricingReport
        {
            SourcePath = $"October{year}_AssurePlus.pdf",
            SourcePage = 9,
            EffectiveDate = new DateOnly(year, 10, 1)
        };
        report.Benefits.Add(health);
        report.Benefits.Add(dental);
        AddDistributionRows(report, year);
        return report;
    }

    private static void AddDistributionRows(AssurePlusPricingReport report, int year)
    {
        var healthYourClaims = year == 2025
            ? new[] { 47.28m, 49.47m, 0.00m, 2.12m, 1.13m }
            : new[] { 63.84m, 30.33m, 2.68m, 2.02m, 1.13m };
        var healthBenchmark = year == 2025
            ? new[] { 47.22m, 35.77m, 5.00m, 10.48m, 1.53m }
            : new[] { 46.16m, 36.83m, 5.42m, 10.57m, 1.02m };
        var dentalYourClaims = year == 2025
            ? new[] { 76.57m, 6.24m, 3.45m, 13.74m, 0.00m }
            : new[] { 69.96m, 0.00m, 6.28m, 23.76m, 0.00m };
        var dentalBenchmark = year == 2025
            ? new[] { 87.24m, 3.57m, 3.86m, 4.47m, 0.86m }
            : new[] { 88.00m, 3.19m, 3.78m, 4.36m, 0.67m };

        AddDistributionRows(
            report,
            "Health",
            ["Drugs", "Paramedicals", "Supplies", "Vision", "Other"],
            healthYourClaims,
            healthBenchmark);
        AddDistributionRows(
            report,
            "Dental",
            ["Basic Other", "Periodontics", "Endodontics", "Major", "Ortho"],
            dentalYourClaims,
            dentalBenchmark);
    }

    private static void AddDistributionRows(
        AssurePlusPricingReport report,
        string benefit,
        IReadOnlyList<string> categories,
        IReadOnlyList<decimal> yourClaims,
        IReadOnlyList<decimal> benchmarkClaims)
    {
        for (var index = 0; index < categories.Count; index++)
        {
            report.DistributionRows.Add(new AssurePlusDistributionRow
            {
                Benefit = benefit,
                Category = categories[index],
                YourClaimsPercent = yourClaims[index],
                BenchmarkClaimsPercent = benchmarkClaims[index],
                SourceReference = report.SourceReference
            });
        }
    }

    private static AssurePlusBenefitPricing Benefit(
        string benefit,
        string previousRating,
        string recentRating,
        decimal standard,
        decimal adjustment,
        decimal total) =>
        new()
        {
            Benefit = benefit,
            PreviousYearRating = previousRating,
            RecentYearRating = recentRating,
            StandardChangePercent = standard,
            ExperienceAdjustmentPercent = adjustment,
            TotalChangePercent = total,
            SourceReference = "AssurePLUS PDF page 9"
        };

    private static void AddRating(AssurePlusExperienceSummary summary, int year, string benefit, string rating)
    {
        summary.AnnualRatings.Add(new AssurePlusAnnualRating
        {
            Year = year,
            Benefit = benefit,
            Rating = rating,
            SourceReference = "AssurePLUS PDF page 9"
        });
    }

    private static T SetControl<T>(MainWindow window, string fieldName)
        where T : class, new()
    {
        var control = new T();
        SetField(window, fieldName, control);
        return control;
    }

    private static void SetField(object target, string fieldName, object? value)
    {
        var field = target.GetType().GetField(
            fieldName,
            BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
        Assert.NotNull(field);
        field!.SetValue(target, value);
    }

    private static void InvokePrivate(object target, string methodName)
    {
        var method = target.GetType().GetMethod(
            methodName,
            BindingFlags.Instance | BindingFlags.NonPublic);
        Assert.NotNull(method);
        method!.Invoke(target, null);
    }

    private static T Property<T>(object target, string propertyName)
    {
        var property = target.GetType().GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public);
        Assert.NotNull(property);
        return Assert.IsType<T>(property!.GetValue(target));
    }

    private static string SliceBetween(string source, string startMarker, string endMarker)
    {
        var start = source.IndexOf(startMarker, StringComparison.Ordinal);
        Assert.True(start >= 0, $"Source marker was not found: {startMarker}");
        var end = source.IndexOf(endMarker, start + startMarker.Length, StringComparison.Ordinal);
        Assert.True(end > start, $"Source marker was not found after {startMarker}: {endMarker}");
        return source[start..end];
    }

    private static bool IsCollapsed(XElement element) => string.Equals(
        (string?)element.Attribute("Visibility"),
        "Collapsed",
        StringComparison.OrdinalIgnoreCase);

    private static string FindRepositoryRoot()
    {
        var folder = new DirectoryInfo(AppContext.BaseDirectory);
        while (folder is not null && !File.Exists(Path.Combine(folder.FullName, "RenewalReviewBuilder.sln")))
        {
            folder = folder.Parent;
        }

        return folder?.FullName ?? throw new DirectoryNotFoundException("Repository root was not found.");
    }
}

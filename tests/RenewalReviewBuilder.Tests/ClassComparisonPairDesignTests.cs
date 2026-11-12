using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class ClassComparisonPairDesignTests
{
    [Fact]
    public void Identical_Normalized_Values_Return_Same()
    {
        var model = Build(Benefit("Class A", "Extended Health Care", "Maximum", "$1,000"), Benefit("Class B", "Extended Health Care", "Maximum", "1000"));
        Assert.Equal(ClassComparisonDifferenceKind.Same, Assert.Single(model.ComparisonResult.Rows).DifferenceKind);
    }

    [Fact]
    public void Numeric_Maximum_Direction_And_Label_Are_Deterministic()
    {
        var model = Build(Benefit("Class A", "Life Insurance", "Maximum", "$150,000"), Benefit("Class B", "Life Insurance", "Maximum", "$300,000"));
        var row = Assert.Single(model.ComparisonResult.Rows);
        Assert.Equal(ClassComparisonDifferenceKind.Higher, row.DifferenceKind);
        Assert.Equal("Higher maximum", row.DifferenceLabel);
    }

    [Fact]
    public void Eligibility_Differences_Are_Counted()
    {
        var model = Build(Benefit("Class A", "Eligibility", "Waiting Period", "3 months"), Benefit("Class B", "Eligibility", "Waiting Period", "1 month"));
        Assert.Equal(1, model.EligibilityDifferenceCount);
        Assert.Equal(ClassComparisonDifferenceKind.DifferentEligibility, Assert.Single(model.ComparisonResult.Rows).DifferenceKind);
    }

    [Fact]
    public void Missing_Value_Is_Not_Same()
    {
        var model = Build(
            Benefit("Class A", "Extended Health Care", "Vision Maximum", "$250"),
            Benefit("Class B", "Dental Care", "Coinsurance", "80%"));
        var vision = Assert.Single(model.ComparisonResult.Rows, row => row.DisplayName == "Vision Maximum");
        Assert.Equal(ClassComparisonDifferenceKind.BaselineOnly, vision.DifferenceKind);
    }

    [Fact]
    public void Low_Confidence_Value_Returns_NeedsReview()
    {
        var model = Build(
            Benefit("Class A", "Life Insurance", "Maximum", "$150,000", 0.95m),
            Benefit("Class B", "Life Insurance", "Maximum", "$300,000", 0.60m));
        var row = Assert.Single(model.ComparisonResult.Rows);
        Assert.Equal(ClassComparisonDifferenceKind.NeedsReview, row.DifferenceKind);
        Assert.False(row.HasReliableEvidence);
    }

    [Fact]
    public void Swapping_Classes_Reverses_Directional_Comparison()
    {
        var model = Build(Benefit("Class A", "Life Insurance", "Maximum", "$150,000"), Benefit("Class B", "Life Insurance", "Maximum", "$300,000"));
        Assert.True(ClassComparisonEngine.SwapPairClasses(model));
        var row = Assert.Single(model.ComparisonResult.Rows);
        Assert.Equal(ClassComparisonDifferenceKind.Lower, row.DifferenceKind);
        Assert.Equal("Lower maximum", row.DifferenceLabel);
    }

    [Fact]
    public void Key_Differences_Filter_Removes_Identical_Rows()
    {
        var model = Build(
            Benefit("Class A", "Dental Care", "Coinsurance", "80%"), Benefit("Class B", "Dental Care", "Coinsurance", "80%"),
            Benefit("Class A", "Life Insurance", "Maximum", "$100,000"), Benefit("Class B", "Life Insurance", "Maximum", "$200,000"));
        ClassComparisonEngine.SetShowDifferencesOnly(model, true);
        Assert.Single(model.FilteredComparisonRows);
        Assert.DoesNotContain(model.FilteredComparisonRows, row => row.DifferenceKind == ClassComparisonDifferenceKind.Same);
    }

    [Fact]
    public void Search_Filters_By_Benefit_And_Feature_Name()
    {
        var model = Build(
            Benefit("Class A", "Dental Care", "Coinsurance", "80%"), Benefit("Class B", "Dental Care", "Coinsurance", "90%"),
            Benefit("Class A", "Life Insurance", "Maximum", "$100,000"), Benefit("Class B", "Life Insurance", "Maximum", "$200,000"));
        ClassComparisonEngine.SetSearchText(model, "Dental");
        Assert.Single(model.FilteredComparisonRows);
        ClassComparisonEngine.SetSearchText(model, "maximum");
        Assert.Single(model.FilteredComparisonRows);
        Assert.Equal("Maximum", model.FilteredComparisonRows[0].DisplayName);
    }

    [Fact]
    public void Summary_Counts_Match_Comparison_Result()
    {
        var model = Build(
            Benefit("Class A", "Eligibility", "Waiting Period", "3 months"), Benefit("Class B", "Eligibility", "Waiting Period", "1 month"),
            Benefit("Class A", "Cost Sharing", "Employee Contribution", "$25"), Benefit("Class B", "Cost Sharing", "Employee Contribution", "$50"),
            Benefit("Class A", "Dental Care", "Coinsurance", "80%"), Benefit("Class B", "Dental Care", "Coinsurance", "80%"));
        Assert.Equal(model.ComparisonResult.Rows.Count, model.TotalBenefitLines);
        Assert.Equal(model.ComparisonResult.Rows.Count(row => row.IsMaterialDifference), model.MaterialDifferenceCount);
        Assert.Equal(1, model.CostDifferenceCount);
        Assert.Equal(1, model.EligibilityDifferenceCount);
    }

    [Fact]
    public void Selecting_The_Same_Class_Is_Prevented()
    {
        var model = Build(Benefit("Class A", "Dental Care", "Coinsurance", "80%"), Benefit("Class B", "Dental Care", "Coinsurance", "90%"));
        Assert.False(ClassComparisonEngine.TrySetPairClasses(model, "Class A", "Class A", out var message));
        Assert.Contains("different", message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Class_Comparison_Remains_Inside_The_Client_Workspace_Shell()
    {
        var root = FindRepositoryRoot();
        var mainXaml = File.ReadAllText(Path.Combine(root, "src", "RenewalReviewBuilder.App", "MainWindow.xaml"));
        var mainCode = File.ReadAllText(Path.Combine(root, "src", "RenewalReviewBuilder.App", "MainWindow.xaml.cs"));
        var planXaml = File.ReadAllText(Path.Combine(root, "src", "RenewalReviewBuilder.App", "Views", "PlanDesignWorkspaceView.xaml"));

        Assert.Contains("ClientWorkspacePanel", mainXaml, StringComparison.Ordinal);
        Assert.Contains("ClientPlanDesignWorkspaceView", mainXaml, StringComparison.Ordinal);
        Assert.Contains("ClassComparisonWorkspaceView", planXaml, StringComparison.Ordinal);
        Assert.DoesNotContain("Breadcrumb", planXaml, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("ClientSelector", planXaml, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("UpdateClientWorkspaceActiveTab", mainCode, StringComparison.Ordinal);
        Assert.Contains("(\"Plan Design\", ClientWorkspacePlanDesignTabButton)", mainCode, StringComparison.Ordinal);
    }

    private static ClassComparisonWorkspaceViewModel Build(params ClassBenefitSourceRow[] rows)
    {
        return ClassComparisonEngine.Build(new ClassComparisonBuildRequest
        {
            ClientId = 1,
            ClientName = "Comparison Test Client",
            BenefitRows = rows,
            Preferences = new ClassComparisonPreferenceState { BaselineClassName = "Class A", SelectedClassNames = ["Class A", "Class B"] }
        });
    }

    private static ClassBenefitSourceRow Benefit(string className, string category, string feature, string value, decimal? confidence = 0.95m) =>
        new(className, category, category, feature, value, "", "", "", "", "", "", category, "source.pdf", "1", value, "", confidence);

    private static string FindRepositoryRoot()
    {
        var folder = new DirectoryInfo(AppContext.BaseDirectory);
        while (folder is not null)
        {
            if (File.Exists(Path.Combine(folder.FullName, "RenewalReviewBuilder.sln"))) return folder.FullName;
            folder = folder.Parent;
        }
        throw new DirectoryNotFoundException("Repository root was not found.");
    }
}

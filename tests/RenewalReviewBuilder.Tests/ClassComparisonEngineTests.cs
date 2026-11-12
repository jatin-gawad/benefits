using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class ClassComparisonEngineTests
{
    [Fact]
    public void Normalize_Treats_Formatted_And_Unformatted_Money_As_Equal()
    {
        var formatted = ClassComparisonValueNormalizer.Normalize("$1,000", ClassComparisonValueKind.Money);
        var plain = ClassComparisonValueNormalizer.Normalize("1000", ClassComparisonValueKind.Money);

        Assert.True(ClassComparisonValueNormalizer.AreEquivalent(formatted, plain));
        Assert.Equal(1000m, formatted.NumericValue);
    }

    [Fact]
    public void Normalize_Treats_Percent_Fraction_And_Percent_Display_As_Equal()
    {
        var percent = ClassComparisonValueNormalizer.Normalize("80%", ClassComparisonValueKind.Percent);
        var fraction = ClassComparisonValueNormalizer.Normalize("0.80", ClassComparisonValueKind.Percent);

        Assert.True(ClassComparisonValueNormalizer.AreEquivalent(percent, fraction));
        Assert.Equal(0.8m, percent.NumericValue);
    }

    [Fact]
    public void Normalize_Treats_TwentyFour_Months_And_Two_Years_As_Equal()
    {
        var months = ClassComparisonValueNormalizer.Normalize("24 months", ClassComparisonValueKind.Duration);
        var years = ClassComparisonValueNormalizer.Normalize("2 years", ClassComparisonValueKind.Duration);

        Assert.True(ClassComparisonValueNormalizer.AreEquivalent(months, years));
    }

    [Fact]
    public void Normalize_Treats_Included_And_True_As_Equal_But_Not_As_Missing()
    {
        var included = ClassComparisonValueNormalizer.Normalize("Included", ClassComparisonValueKind.Boolean);
        var trueValue = ClassComparisonValueNormalizer.Normalize("true", ClassComparisonValueKind.Boolean);
        var notIncluded = ClassComparisonValueNormalizer.Normalize("Not Included", ClassComparisonValueKind.Boolean);
        var missing = ClassComparisonValueNormalizer.Normalize("", ClassComparisonValueKind.Boolean);

        Assert.True(ClassComparisonValueNormalizer.AreEquivalent(included, trueValue));
        Assert.False(notIncluded.IsMissing);
        Assert.False(ClassComparisonValueNormalizer.AreEquivalent(notIncluded, missing));
    }

    [Fact]
    public void Build_Uses_Directional_Rules_And_Selects_Row_Best_Value()
    {
        var model = ClassComparisonEngine.Build(new ClassComparisonBuildRequest
        {
            ClientId = 1,
            ClientName = "Comparison Test Client",
            BenefitRows =
            [
                Benefit("Class A", "Extended Health Care", "Paramedical Maximum", "$500"),
                Benefit("Class B", "Extended Health Care", "Paramedical Maximum", "$1,000"),
                Benefit("Class A", "Eligibility", "Waiting Period", "3 months"),
                Benefit("Class B", "Eligibility", "Waiting Period", "1 month"),
                Benefit("Class A", "Long Term Disability", "LTD Taxability", "Non-taxable"),
                Benefit("Class B", "Long Term Disability", "LTD Taxability", "Taxable")
            ],
            Preferences = new ClassComparisonPreferenceState
            {
                BaselineClassName = "Class A",
                SelectedClassNames = ["Class A", "Class B"]
            }
        });

        var paramedical = Assert.Single(model.ComparisonRows, row => row.Benefit == "Paramedical Maximum");
        Assert.Equal(ClassComparisonStatus.Same, paramedical.Cells[0].Status);
        Assert.Equal(ClassComparisonStatus.Better, paramedical.Cells[1].Status);
        Assert.Equal("Class B", paramedical.BestValueText);

        var waiting = Assert.Single(model.ComparisonRows, row => row.Benefit == "Waiting Period");
        Assert.Equal(ClassComparisonStatus.Better, waiting.Cells[1].Status);
        Assert.Equal("Class B", waiting.BestValueText);

        var taxability = Assert.Single(model.ComparisonRows, row => row.Benefit == "LTD Taxability");
        Assert.Equal(ClassComparisonStatus.Different, taxability.Cells[1].Status);
        Assert.Equal("Informational", taxability.BestValueText);
    }

    [Fact]
    public void Selection_Requires_One_Class_And_Supports_Six_Classes()
    {
        var benefits = Enumerable.Range(1, 6)
            .Select(index => Benefit($"Class {index}", "Extended Health Care", "Vision Maximum", $"${index * 100}"))
            .ToList();
        var model = ClassComparisonEngine.Build(new ClassComparisonBuildRequest
        {
            ClientId = 1,
            ClientName = "Selection Test Client",
            BenefitRows = benefits
        });

        Assert.Equal(6, model.Classes.Count);
        Assert.Equal(6, model.SelectedClassCount);
        Assert.All(model.Classes, item => Assert.True(item.IsSelected));

        foreach (var selected in model.Classes.Where(item => item.IsSelected).Skip(1).ToList())
        {
            Assert.True(ClassComparisonEngine.TrySetClassSelection(model, selected.ClassName, false, out _));
        }

        Assert.False(ClassComparisonEngine.TrySetClassSelection(model, model.Classes.Single(item => item.IsSelected).ClassName, false, out var minimumMessage));
        Assert.Contains("at least one", minimumMessage, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(3)]
    [InlineData(6)]
    public void Build_Projects_Each_Selected_Class_As_A_Column_Source(int classCount)
    {
        var model = ClassComparisonEngine.Build(new ClassComparisonBuildRequest
        {
            BenefitRows = Enumerable.Range(1, classCount)
                .Select(index => Benefit($"Class {index}", "Health", "Vision Maximum", $"${index * 100}"))
                .ToList()
        });

        Assert.Equal(classCount, model.SelectedClasses.Count);
        var row = Assert.Single(model.ComparisonRows, item => !item.IsSectionHeader);
        Assert.Equal(classCount, row.Cells.Count);
    }

    [Fact]
    public void Sections_Expand_And_Collapse_All_Without_Changing_Canonical_Source_Features()
    {
        var model = ClassComparisonEngine.Build(new ClassComparisonBuildRequest
        {
            BenefitRows =
            [
                Benefit("Class A", "Life Insurance", "Employee Basic Life", "$50,000"),
                Benefit("Class A", "Dental Care", "Basic Coinsurance", "80%")
            ]
        });

        var featureCount = model.SourceFeatures.Count;
        ClassComparisonEngine.SetAllSectionsExpanded(model, false);
        Assert.All(model.ComparisonRows, row => Assert.True(row.IsSectionHeader));
        ClassComparisonEngine.SetAllSectionsExpanded(model, true);
        Assert.Contains(model.ComparisonRows, row => !row.IsSectionHeader && row.Benefit == "Employee Basic Life");
        Assert.Equal(featureCount, model.SourceFeatures.Count);
    }

    [Theory]
    [InlineData("N/A", "N/A")]
    [InlineData("Not covered", "Not covered")]
    [InlineData("Not available", "—")]
    public void Build_Preserves_Canonical_Missing_Value_Presentation(string source, string expected)
    {
        var model = ClassComparisonEngine.Build(new ClassComparisonBuildRequest { BenefitRows = [Benefit("Class A", "Health", "Virtual Care", source)] });
        var row = Assert.Single(model.ComparisonRows, item => !item.IsSectionHeader);
        Assert.Equal(expected, Assert.Single(row.Cells).DisplayValue);
    }

    [Fact]
    public void Build_Aggregates_Monthly_Costs_And_Uses_Available_Score_Components()
    {
        var model = ClassComparisonEngine.Build(new ClassComparisonBuildRequest
        {
            ClientId = 1,
            ClientName = "Cost Test Client",
            BenefitRows =
            [
                Benefit("Class A", "Extended Health Care", "Vision Maximum", "$250"),
                Benefit("Class B", "Extended Health Care", "Vision Maximum", "$350")
            ],
            CostRows =
            [
                new ClassCostSourceRow("Class A", 10, 1000m, 100m, 1100m, "billing.pdf"),
                new ClassCostSourceRow("Class B", 20, 1800m, 200m, 2000m, "billing.pdf")
            ]
        });

        var classA = Assert.Single(model.CostSummaries, row => row.ClassName == "Class A");
        Assert.Equal(12000m, classA.AnnualEmployerCost);
        Assert.Equal(13200m, classA.AnnualTotalCost);
        Assert.Equal(1320m, classA.CostPerEmployee);
        Assert.True(classA.TotalClientCostPercent is > 0m and < 1m);

        Assert.Equal(2, model.ValueScores.Count);
        Assert.All(model.ValueScores, score => Assert.InRange(score.Score, 0, 100));
        Assert.All(model.ValueScores, score => Assert.Contains(score.Components, component => component.Category == "Administration complexity" && component.Score is null));
    }

    [Fact]
    public void Build_Leaves_Unavailable_Class_Value_Distinct_From_NotIncluded()
    {
        var model = ClassComparisonEngine.Build(new ClassComparisonBuildRequest
        {
            BenefitRows =
            [
                Benefit("Class A", "Extended Health Care", "Virtual Care", "Included"),
                Benefit("Class B", "Extended Health Care", "Virtual Care", "Not Included"),
                Benefit("Class C", "Extended Health Care", "Vision Maximum", "$250")
            ]
        });

        var virtualCare = Assert.Single(model.ComparisonRows, row => row.Benefit == "Virtual Care");
        Assert.Equal(ClassComparisonStatus.Worse, virtualCare.Cells.Single(cell => cell.ClassName == "Class B").Status);
        Assert.Equal(ClassComparisonStatus.Missing, virtualCare.Cells.Single(cell => cell.ClassName == "Class C").Status);
    }

    [Fact]
    public void Build_Uses_BenefitSummary_Value_Normalization_Without_Changing_Comparison_Identity()
    {
        var model = ClassComparisonEngine.Build(new ClassComparisonBuildRequest
        {
            BenefitRows =
            [
                Benefit("Class A", "Drug Plan", "Maximum", "1500 annual"),
                Benefit("Class A", "Dental", "coinsurance", "80%")
            ]
        });

        var drugRow = Assert.Single(model.ComparisonRows, row => !row.IsSectionHeader && row.Benefit == "Maximum");
        Assert.Equal("$1,500 per calendar year", Assert.Single(drugRow.Cells).DisplayValue);
        Assert.Contains(model.ComparisonRows, row => row.IsSectionHeader && row.Section == "Extended Health Care");
        Assert.Contains(model.ComparisonRows, row => !row.IsSectionHeader && row.Benefit == "Coinsurance");
    }

    [Fact]
    public void Build_Preserves_Tiered_Ltd_Formula_Instead_Of_The_Middle_Coinsurance_Tier()
    {
        const string source =
            "Long Term Disability Benefit Amount: 66.7% of first $2,500 of monthly Earnings, plus 50% of the next $2,500 of monthly Earnings, plus 40% of any excess amount, up to a maximum benefit of $10,000.";
        var model = ClassComparisonEngine.Build(new ClassComparisonBuildRequest
        {
            BenefitRows =
            [
                new ClassBenefitSourceRow(
                    "Class A",
                    "LTD",
                    "LTD",
                    "Plan Summary",
                    source,
                    "50%",
                    "",
                    "",
                    "",
                    "",
                    "",
                    "LTD",
                    "source.pdf",
                    "1",
                    source,
                    "")
            ]
        });

        var row = Assert.Single(model.ComparisonRows, item => !item.IsSectionHeader && item.Benefit == "Plan Summary");
        var displayValue = Assert.Single(row.Cells).DisplayValue;

        Assert.Equal(
            "66.7% of first $2,500 of monthly earnings; 50% of next $2,500 of monthly earnings; 40% of excess (maximum $10,000)",
            displayValue);
        Assert.NotEqual("50%", displayValue);
    }

    [Fact]
    public void Build_Excludes_Policy_Narrative_And_Retains_A_Structured_Supplemental_Value()
    {
        var narrative = "Accidental Death and Dismemberment coverage is subject to the terms and conditions of this benefit and the policy. Upon a loss, the insurer will pay according to the schedule of benefits.";
        var model = ClassComparisonEngine.Build(new ClassComparisonBuildRequest
        {
            BenefitRows =
            [
                new ClassBenefitSourceRow("Class A", "Life Insurance", "Life Insurance", "AD&D", narrative, "", "", "$50,000", "", "", "", "", "source.pdf", "1", narrative, "")
            ]
        });

        var row = Assert.Single(model.ComparisonRows, item => !item.IsSectionHeader);
        Assert.Equal("AD&D", row.Benefit);
        Assert.Equal("$50,000", Assert.Single(row.Cells).DisplayValue);
        Assert.DoesNotContain(narrative, row.Cells[0].DisplayValue, StringComparison.Ordinal);
    }

    [Fact]
    public void Filters_Baseline_And_Section_Collapse_Refresh_The_Visible_Matrix()
    {
        var model = ClassComparisonEngine.Build(new ClassComparisonBuildRequest
        {
            BenefitRows =
            [
                Benefit("Class A", "Eligibility", "Waiting Period", "3 months"),
                Benefit("Class B", "Eligibility", "Waiting Period", "1 month"),
                Benefit("Class A", "Other", "Administrative Provision", "Manual"),
                Benefit("Class B", "Other", "Administrative Provision", "Automatic")
            ],
            Preferences = new ClassComparisonPreferenceState
            {
                BaselineClassName = "Class A",
                SelectedClassNames = ["Class A", "Class B"]
            }
        });

        Assert.True(ClassComparisonEngine.TrySetBaseline(model, "Class B"));
        var waiting = Assert.Single(model.ComparisonRows, row => row.Benefit == "Waiting Period");
        Assert.Equal(ClassComparisonStatus.Worse, waiting.Cells.Single(cell => cell.ClassName == "Class A").Status);

        ClassComparisonEngine.ToggleSection(model, "Eligibility");
        var eligibilityHeader = Assert.Single(model.ComparisonRows, row => row.IsSectionHeader && row.Section == "Eligibility");
        Assert.False(eligibilityHeader.IsExpanded);
        Assert.DoesNotContain(model.ComparisonRows, row => row.Benefit == "Waiting Period");

        ClassComparisonEngine.SetKeyBenefitsOnly(model, true);
        Assert.DoesNotContain(model.ComparisonRows, row => row.Benefit == "Administrative Provision");
    }

    private static ClassBenefitSourceRow Benefit(string className, string category, string feature, string value)
    {
        return new ClassBenefitSourceRow(
            className,
            category,
            category,
            feature,
            value,
            "",
            "",
            "",
            "",
            "",
            "",
            category,
            "source.pdf",
            "1",
            value,
            "");
    }
}

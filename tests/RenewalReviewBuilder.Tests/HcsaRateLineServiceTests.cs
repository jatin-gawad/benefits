using System.Reflection;
using System.Collections.Generic;
using ClosedXML.Excel;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class HcsaRateLineServiceTests
{
    [Fact]
    public void Legacy_Critical_Illness_Rows_Are_Excluded_From_Hcsa_Lines_And_Totals()
    {
        var analysis = new AnalysisResult();
        var calculation = new HcsaCalculation
        {
            Classes =
            [
                new HcsaClassCalculation
                {
                    ClassName = "Class 1",
                    EligibleEmployees = 2,
                    CurrentAnnualCredit = 5_945.454545m,
                    RenewalAnnualCredit = 5_945.454545m,
                    AdministrationFeePercent = 10m
                },
                new HcsaClassCalculation
                {
                    ClassName = "Class 2",
                    EligibleEmployees = 14,
                    CurrentAnnualCredit = 594.545455m,
                    RenewalAnnualCredit = 594.545455m,
                    AdministrationFeePercent = 10m
                },
                new HcsaClassCalculation
                {
                    ClassName = "Critical Illness - Single",
                    EligibleEmployees = 5,
                    CurrentAnnualCredit = 100m,
                    RenewalAnnualCredit = 100m,
                    AdministrationFeePercent = 10m
                },
                new HcsaClassCalculation
                {
                    ClassName = "Critical Illness - Couple",
                    EligibleEmployees = 11,
                    CurrentAnnualCredit = 100m,
                    RenewalAnnualCredit = 100m,
                    AdministrationFeePercent = 10m
                }
            ]
        };

        new HcsaRateLineService().Apply(analysis, calculation);

        Assert.Equal(2, analysis.Pdf.RateLines.Count);
        Assert.Contains(analysis.Pdf.RateLines, line => line.Benefit == "HCSA" && line.TierOrClass == "Class 1");
        Assert.Contains(analysis.Pdf.RateLines, line => line.Benefit == "HCSA" && line.TierOrClass == "Class 2");
        Assert.DoesNotContain(analysis.Pdf.RateLines, line =>
            line.TierOrClass.Contains("Critical Illness", StringComparison.OrdinalIgnoreCase));
        Assert.Equal(22_236m, calculation.CurrentAnnualCost);
        Assert.Equal(22_236m, calculation.RenewalAnnualCost);
        Assert.Contains(analysis.SourceChecks, check =>
            check.Area == "HCSA funding" &&
            check.Detail.Contains("2 class line(s)", StringComparison.OrdinalIgnoreCase));

        Assert.Equal(2, calculation.RemoveLegacyCriticalIllnessEntries());
        Assert.Equal(2, calculation.Classes.Count);
    }

    [Fact]
    public void Apply_Calculates_ClassFunding_As_HcsaMonthlyRateLines()
    {
        var analysis = new AnalysisResult();
        analysis.Pdf.RateLines.Add(new RateLine { Benefit = "Health Spending Account", CurrentPremium = 10m });
        var calculation = new HcsaCalculation
        {
            Classes =
            [
                new HcsaClassCalculation
                {
                    ClassName = "Class 2",
                    EligibleEmployees = 6,
                    CurrentAnnualCredit = 1_250m,
                    RenewalAnnualCredit = 1_500m,
                    UsagePercent = 100m,
                    AdministrationFeePercent = 10m
                }
            ]
        };

        new HcsaRateLineService().Apply(analysis, calculation);

        var line = Assert.Single(analysis.Pdf.RateLines);
        Assert.Equal("HCSA", line.Benefit);
        Assert.Equal("Class 2", line.TierOrClass);
        Assert.Equal(6, line.Lives);
        Assert.Equal(687.50m, line.CurrentPremium);
        Assert.Equal(825.00m, line.RenewalPremium);
        Assert.Equal(114.583333m, line.CurrentRate);
        Assert.Equal(137.5m, line.RenewalRate);
        Assert.Equal("Employer-funded HCSA", line.FundingArrangement);
        Assert.Contains(analysis.SourceChecks, check => check.Area == "HCSA funding" && check.Status == "Entered");
    }

    [Fact]
    public void HcsaRateLines_Appear_In_Workbook_RateTables_And_PowerPointTokens()
    {
        var analysis = new AnalysisResult();
        new HcsaRateLineService().Apply(analysis, new HcsaCalculation
        {
            Classes =
            [
                new HcsaClassCalculation
                {
                    ClassName = "Class 4",
                    EligibleEmployees = 2,
                    CurrentAnnualCredit = 2_000m,
                    RenewalAnnualCredit = 2_500m,
                    UsagePercent = 100m,
                    AdministrationFeePercent = 10m
                }
            ]
        });

        var outputFolder = Path.Combine(Path.GetTempPath(), "RenewalReviewBuilderTests", $"HcsaRateLines-{Guid.NewGuid():N}");
        var path = new ExcelReportWriter().Write(analysis, outputFolder);

        using var workbook = new XLWorkbook(path);
        var sourceTables = workbook.Worksheet("PowerPoint Export Tables");
        Assert.Contains(sourceTables.CellsUsed(), cell => cell.GetString().Equals("HCSA", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(sourceTables.CellsUsed(), cell => cell.GetString().Equals("Class 4", StringComparison.OrdinalIgnoreCase));

        var formattedRates = workbook.Worksheet("RateTable_LineRates_Formatted");
        Assert.Contains(formattedRates.CellsUsed(), cell => cell.GetString().Equals("Health Spending Account", StringComparison.OrdinalIgnoreCase));

        var tokens = InvokeTokenMap(analysis);
        Assert.Equal("$366.67", tokens["{{HCSA_CurrentPremium}}"]);
        Assert.Equal("$458.33", tokens["{{HCSA_RenewalPremium}}"]);
    }

    private static Dictionary<string, string> InvokeTokenMap(AnalysisResult analysis)
    {
        var method = typeof(PowerPointDeckWriter).GetMethod("BuildTokenMap", BindingFlags.NonPublic | BindingFlags.Static);
        Assert.NotNull(method);
        return Assert.IsType<Dictionary<string, string>>(method!.Invoke(null, [analysis]));
    }
}

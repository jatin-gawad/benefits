using ClosedXML.Excel;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class ManulifeRenewalAnalysisBlockTests
{
    [Fact]
    public void Parser_Uses_Detail_Rate_Page_Without_Duplicating_Summary_Rows()
    {
        var result = ManulifePdfParser.ParseExtractedPages(
            [
                Page(
                    3,
                    "Renewal Rate Illustration",
                    "ACCOUNT/CLASS: ALL",
                    "Life 60,000 4 1.718 103.08 1.890 113.40 10.0 %",
                    "Accidental Death and Dismemberment 50,000 2 0.067 3.35 0.067 3.35 0.0 %",
                    "Extended Health Care",
                    "Single 1 425.24 425.24 499.58 499.58",
                    "Family 2 623.17 1,246.34 732.16 1,464.32",
                    "*Total Extended Health Care $1,671.58 $1,963.90 17.5 %",
                    "Dental",
                    "Single 1 40.37 40.37 50.87 50.87",
                    "Family 2 106.72 213.44 134.47 268.94",
                    "Total Dental $253.81 $319.81 26.0 %",
                    "Total Monthly Premium $2,031.82 $2,400.46"),
                Page(
                    4,
                    "Renewal Rate Illustration Summary",
                    "Extended Health Care",
                    "000/A Single 1 425.24 425.24 499.58 499.58",
                    "Family 2 623.17 1,246.34 732.16 1,464.32",
                    "*Total Extended Health Care Single $425.24 $499.58 Family $1,246.34 $1,464.32 Total $1,671.58 $1,963.90",
                    "Dental",
                    "ALL Single 1 40.37 40.37 50.87 50.87",
                    "Family 2 106.72 213.44 134.47 268.94",
                    "Total Dental Single 1 $40.37 $50.87 Family 2 $213.44 $268.94 Total 3 $253.81 $319.81",
                    "Total Monthly Premium $2,031.82 $2,400.46")
            ],
            "dj-martin-renewal.pdf");

        Assert.Equal(6, result.RateLines.Count);
        Assert.DoesNotContain(result.RateLines, line => line.SourceReference == "PDF page 4");
        Assert.Equal(2031.82m, result.RateLines.Sum(line => line.CurrentPremium ?? 0m));
        Assert.Equal(2400.46m, result.RateLines.Sum(line => line.RenewalPremium ?? 0m));
        Assert.Equal(2031.82m, result.CurrentMonthlyPremium);
        Assert.Equal(2400.46m, result.RenewalMonthlyPremium);
        Assert.Equal(1671.58m, result.CurrentEhcPremium);
        Assert.Equal(1963.90m, result.RenewalEhcPremium);
        Assert.Equal(253.81m, result.CurrentDentalPremium);
        Assert.Equal(319.81m, result.RenewalDentalPremium);
    }

    [Fact]
    public void Parser_Sums_Source_Totals_Across_Distinct_Account_Class_Pages()
    {
        var result = ManulifePdfParser.ParseExtractedPages(
            [
                Page(
                    3,
                    "Renewal Rate Illustration",
                    "ACCOUNT/CLASS: 001/01, 002/03",
                    "Life 11,757,000 102 0.119 1,399.08 0.119 1,399.08 0.0 %",
                    "Extended Health Care",
                    "Single 54 126.32 6,821.28 116.33 6,281.82",
                    "Family 41 269.49 11,049.09 247.34 10,140.94",
                    "Total Extended Health Care $17,870.37 $16,422.76 -8.1 %",
                    "Dental",
                    "Single 51 52.94 2,699.94 51.67 2,635.17",
                    "Family 44 140.93 6,200.92 137.54 6,051.76",
                    "Total Dental $8,900.86 $8,686.93 -2.4 %",
                    "Total Monthly Premium $38,197.71 $35,135.29"),
                Page(
                    4,
                    "Renewal Rate Illustration",
                    "ACCOUNT/CLASS: 001/02",
                    "Life 518,000 3 0.118 61.12 0.118 61.12 0.0 %",
                    "Extended Health Care",
                    "Single 0 174.93 0.00 160.69 0.00",
                    "Family 3 423.38 1,270.14 386.23 1,158.69",
                    "Total Extended Health Care $1,270.14 $1,158.69 -8.8 %",
                    "Dental",
                    "Single 0 58.24 0.00 56.84 0.00",
                    "Family 3 152.85 458.55 149.17 447.51",
                    "Total Dental $458.55 $447.51 -2.4 %",
                    "Total Monthly Premium $2,052.91 $1,938.86")
            ],
            "groundworks-renewal.pdf");

        Assert.Equal(40250.62m, result.CurrentMonthlyPremium);
        Assert.Equal(37074.15m, result.RenewalMonthlyPremium);
        Assert.Equal(19140.51m, result.CurrentEhcPremium);
        Assert.Equal(17581.45m, result.RenewalEhcPremium);
        Assert.Equal(9359.41m, result.CurrentDentalPremium);
        Assert.Equal(9134.44m, result.RenewalDentalPremium);
    }

    [Fact]
    public void Parser_Maps_Ltd_Experience_With_Average_Lives_And_No_Weighting()
    {
        var result = ManulifePdfParser.ParseExtractedPages(
            [
                Page(
                    9,
                    "Long Term Disability",
                    "Experience History",
                    "EXPERIENCE PERIOD FROM TO AVERAGE NUMBER OF LIVES BILLED PREMIUM PAID CLAIMS INCURRED CLAIMS INCURRED LOSS RATIO",
                    "01-Oct-2024 31-May-2025 75 28,851 0 14,753 51.1 %",
                    "01-Jun-2025 31-May-2026 78 47,196 47,138 49,845 105.6 %")
            ],
            "manulife-renewal.pdf");

        var current = Assert.Single(
            result.ExperienceLines,
            line => line.Benefit.Equals("Long Term Disability", StringComparison.OrdinalIgnoreCase) &&
                    line.Period.Equals("01-Jun-2025 - 31-May-2026", StringComparison.OrdinalIgnoreCase));

        Assert.Equal(78, current.AverageLives);
        Assert.Equal(47196m, current.BilledPremium);
        Assert.Equal(47138m, current.PaidClaims);
        Assert.Equal(49845m, current.IncurredClaims);
        Assert.Equal(99.9m, current.PaidLossRatio);
        Assert.Equal(105.6m, current.IncurredLossRatio);
        Assert.Null(current.Weighting);
    }

    [Fact]
    public void Parser_Extracts_Current_Health_Dental_And_Pooled_Renewal_Blocks()
    {
        var result = ManulifePdfParser.ParseExtractedPages(
            [
                Page(
                    4,
                    "Extended Health Care",
                    "Experience History Non Pooled Extended Health Care",
                    "ACCOUNT/CLASS: 001/A,B, 002/A,B",
                    "EXPERIENCE PERIOD FROM TO BILLED PREMIUM EXPERIENCE RATED PAID CLAIMS INCURRED CLAIMS INCURRED LOSS RATIO WEIGHTING",
                    "01-Jun-2025 31-May-2026 177,001 107,864 107,864 60.9 % 1",
                    "Renewal Factors",
                    "Trend Factor 11.80 %",
                    "Credibility 100.0 %",
                    "Target Loss Ratio 86.8 %",
                    "Rate Adjustment -26.2 %",
                    "Pooled Extended Health Care",
                    "EXPERIENCE PERIOD FROM TO POOLED PREMIUM DRUG STOP LOSS CLAIMS OTHER POOLED CLAIMS",
                    "01-Jun-2025 31-May-2026 12,229 0 406"),
                Page(
                    5,
                    "Dental",
                    "Experience History",
                    "ACCOUNT/CLASS: 001/A,B, 002/A,B",
                    "EXPERIENCE PERIOD FROM TO AVERAGE NUMBER OF LIVES BILLED PREMIUM PAID CLAIMS INCURRED CLAIMS INCURRED LOSS RATIO WEIGHTING",
                    "01-Jun-2025 31-May-2026 Single 24 138,701 130,991 130,991 94.4 % 1",
                    "Family 56",
                    "Renewal Factors",
                    "Trend Factor 8.90 %",
                    "Credibility 100.0 %",
                    "Target Loss Ratio 86.8 %",
                    "Rate Adjustment 13.2 %")
            ],
            "manulife-renewal.pdf");

        var health = Assert.Single(
            result.ExperienceLines,
            line => line.Benefit.Equals("Non Pooled Extended Health Care - Subtotal", StringComparison.OrdinalIgnoreCase));
        Assert.Equal("01-Jun-2025 - 31-May-2026", health.Period);
        Assert.Equal(177001m, health.BilledPremium);
        Assert.Equal(107864m, health.PaidClaims);
        Assert.Equal(107864m, health.IncurredClaims);
        Assert.Equal(60.9m, health.IncurredLossRatio);
        Assert.Equal(11.8m, health.TrendFactor);
        Assert.Equal(100m, health.Credibility);
        Assert.Equal(86.8m, health.TargetLossRatio);
        Assert.Equal(-26.2m, health.ProposedAdjustment);

        var dental = Assert.Single(
            result.ExperienceLines,
            line => line.Benefit.Equals("Dental", StringComparison.OrdinalIgnoreCase));
        Assert.Equal(24, dental.AverageLives);
        Assert.Equal(138701m, dental.BilledPremium);
        Assert.Equal(130991m, dental.PaidClaims);
        Assert.Equal(130991m, dental.IncurredClaims);
        Assert.Equal(94.4m, dental.IncurredLossRatio);
        Assert.Equal(8.9m, dental.TrendFactor);
        Assert.Equal(13.2m, dental.ProposedAdjustment);

        Assert.Equal(12229m, result.PooledPremium);
        Assert.Equal(0m, result.DrugStopLossClaims);
        Assert.Equal(406m, result.OtherPooledClaims);
        Assert.Equal(6.91m, result.PoolingCharge);
    }

    [Fact]
    public void Parser_Uses_Latest_Pooled_And_NonPooled_Period_To_Calculate_Pooling_Charge()
    {
        var result = ManulifePdfParser.ParseExtractedPages(
            [
                Page(
                    7,
                    "ACCOUNT: ALL",
                    "NON POOLED EXTENDED",
                    "HEALTH CARE - SUBTOTAL",
                    "01-Oct-2023 31-May-2024 31,182 28,242 91 % 30,267 97 %",
                    "01-Jun-2024 31-May-2025 70,072 39,640 57 % 41,763 60 %",
                    "01-Jul-2025 31-Mar-2026 172,688 103,597 60 % 106,105 61 %",
                    "POOLED",
                    "EXPERIENCE PERIOD FROM TO POOLED PREMIUM DRUG STOP LOSS CLAIMS OTHER POOLED CLAIMS",
                    "01-Oct-2023 31-May-2024 8,239 0 1,129",
                    "01-Jun-2024 31-May-2025 18,515 0 1,164",
                    "01-Jul-2025 31-Mar-2026 33,083 0 1,671",
                    "DENTAL",
                    "01-Jul-2025 31-Mar-2026 172,913 102,589 59 % 111,642 64 %")
            ],
            "groundworks-renewal.pdf");

        Assert.Equal(33083m, result.PooledPremium);
        Assert.Equal(0m, result.DrugStopLossClaims);
        Assert.Equal(1671m, result.OtherPooledClaims);
        Assert.Equal(19.16m, result.PoolingCharge);
    }

    [Fact]
    public void AnalysisEngine_Adds_Pooled_Premium_Cross_Check_When_Estimate_Differs()
    {
        var analysis = new AnalysisResult();
        analysis.Pdf.PooledPremium = 12229m;
        analysis.Pdf.RateLines.Add(new RateLine
        {
            Benefit = "Extended Health Care - Pooled",
            TierOrClass = "Single",
            RenewalPremium = 1000m,
            SourceReference = "PDF page 2"
        });

        new AnalysisEngine().Analyze(analysis);

        var check = Assert.Single(
            analysis.SourceChecks,
            item => item.Area.Equals("Pooled premium cross-check", StringComparison.OrdinalIgnoreCase));
        Assert.Equal("Mismatch", check.Status);
        Assert.Contains("$12,000.00", check.Detail);
        Assert.Contains("$12,229.00", check.Detail);
    }

    [Fact]
    public void Workbook_Writes_Formula_Values_And_Pooled_Cross_Check()
    {
        var outputFolder = Path.Combine(Path.GetTempPath(), "RenewalReviewBuilderDesktopFormulaTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(outputFolder);

        try
        {
            var analysis = new AnalysisResult();
            analysis.Pdf.ExperienceLines.Add(new ExperienceLine
            {
                Benefit = "Non Pooled Extended Health Care - Subtotal",
                Period = "01-Jun-2025 - 31-May-2026",
                BilledPremium = 177001m,
                PaidClaims = 107864m,
                IncurredClaims = 107864m,
                IncurredLossRatio = 60.9m,
                Weighting = 1m,
                TrendFactor = 11.8m,
                Credibility = 100m,
                TargetLossRatio = 86.8m,
                ProposedAdjustment = -26.2m,
                SourceReference = "PDF page 4"
            });
            analysis.Pdf.ExperienceLines.Add(new ExperienceLine
            {
                Benefit = "Dental",
                Period = "01-Jun-2025 - 31-May-2026",
                AverageLives = 24,
                BilledPremium = 138701m,
                PaidClaims = 130991m,
                IncurredClaims = 130991m,
                IncurredLossRatio = 94.4m,
                Weighting = 1m,
                TrendFactor = 8.9m,
                Credibility = 100m,
                TargetLossRatio = 86.8m,
                ProposedAdjustment = 13.2m,
                SourceReference = "PDF page 5"
            });
            analysis.Pdf.PooledPremium = 12229m;
            analysis.Pdf.DrugStopLossClaims = 0m;
            analysis.Pdf.OtherPooledClaims = 406m;
            analysis.Pdf.RateLines.Add(new RateLine
            {
                Benefit = "Extended Health Care - Pooled",
                TierOrClass = "Single",
                RenewalPremium = 1000m
            });

            var workbookPath = new ExcelReportWriter().Write(analysis, outputFolder);

            using var workbook = new XLWorkbook(workbookPath);
            var health = workbook.Worksheet("HealthRenewal_FormulaAnalysis");
            Assert.Equal(177001m, health.Cell("E8").GetValue<decimal>());
            Assert.Equal(107864m, health.Cell("E9").GetValue<decimal>());
            Assert.Equal(0.118m, health.Cell("B26").GetValue<decimal>());
            Assert.Equal(1m, health.Cell("B27").GetValue<decimal>());
            Assert.Equal(0.868m, health.Cell("B28").GetValue<decimal>());
            Assert.Equal(-0.262m, health.Cell("I12").GetValue<decimal>());
            Assert.Equal(12229m, health.Cell("I16").GetValue<decimal>());
            Assert.Equal(12000m, health.Cell("I17").GetValue<decimal>());
            Assert.Equal(-229m, health.Cell("I18").GetValue<decimal>());
            Assert.Equal("Review", health.Cell("I20").GetString());
            Assert.Contains("Other pooled claims", health.Cell("I21").GetString());

            var dental = workbook.Worksheet("DentalRenewal_FormulaAnalysis");
            Assert.Equal(138701m, dental.Cell("E8").GetValue<decimal>());
            Assert.Equal(130991m, dental.Cell("E9").GetValue<decimal>());
            Assert.Equal(0.089m, dental.Cell("B26").GetValue<decimal>());
            Assert.Equal(1m, dental.Cell("B27").GetValue<decimal>());
            Assert.Equal(0.868m, dental.Cell("B28").GetValue<decimal>());
            Assert.Equal(0.132m, dental.Cell("I12").GetValue<decimal>());
        }
        finally
        {
            if (Directory.Exists(outputFolder))
            {
                Directory.Delete(outputFolder, recursive: true);
            }
        }
    }

    private static PdfPageText Page(int pageNumber, params string[] lines)
    {
        return new PdfPageText
        {
            PageNumber = pageNumber,
            Lines = lines.ToList(),
            Text = string.Join(" ", lines)
        };
    }
}

using ClosedXML.Excel;
using System.Collections.Generic;
using System.Linq;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class BenefitDisplayOrderTests
{
    [Fact]
    public void Sort_Uses_StandardBenefitOrder_And_SingleParentBelowSingle()
    {
        var rows = BuildShuffledRows();

        var ordered = BenefitDisplayOrder.Sort(
            rows,
            row => row.Benefit,
            row => row.TierOrClass,
            row => row.AccountClass);

        Assert.Equal(ExpectedOrder, ordered.Select(RowKey));
    }

    [Fact]
    public void RenewalReview_RateReconciliation_Uses_StandardDisplayOrder()
    {
        var analysis = new AnalysisResult();
        analysis.Pdf.RateLines.AddRange(BuildShuffledRows());
        analysis.Pdf.CurrentMonthlyPremium = analysis.Pdf.RateLines.Sum(row => row.CurrentPremium ?? 0m);
        analysis.Pdf.RenewalMonthlyPremium = analysis.Pdf.RateLines.Sum(row => row.RenewalPremium ?? 0m);

        var outputFolder = Path.Combine(
            Path.GetTempPath(),
            "RenewalReviewBuilderTests",
            $"BenefitDisplayOrder-{Guid.NewGuid():N}");
        var path = new ExcelReportWriter().Write(analysis, outputFolder);

        using var workbook = new XLWorkbook(path);
        var sheet = workbook.Worksheet("Rate Reconciliation");
        var actual = sheet.RowsUsed()
            .Skip(1)
            .Select(row => $"{row.Cell(1).GetString()}|{row.Cell(2).GetString()}")
            .Where(ExpectedOrder.Contains)
            .ToList();

        Assert.Equal(ExpectedOrder, actual);
    }

    private static readonly string[] ExpectedOrder =
    [
        "Basic Life|All",
        "AD&D|All",
        "Dependent Life|All",
        "Critical Illness|All",
        "Short Term Disability|All",
        "Long Term Disability|All",
        "Extended Health Care|Single",
        "Extended Health Care|Single Parent",
        "Extended Health Care|Couple",
        "Extended Health Care|Family",
        "Dental Care|Single",
        "Dental Care|Couple",
        "Dental Care|Family",
        "Vision|Family"
    ];

    private static List<RateLine> BuildShuffledRows()
    {
        return
        [
            Line("Vision", "Family"),
            Line("Dental Care", "Family"),
            Line("Extended Health Care", "Couple"),
            Line("Long Term Disability", "All"),
            Line("AD&D", "All"),
            Line("Extended Health Care", "Family"),
            Line("Critical Illness", "All"),
            Line("Dental Care", "Single"),
            Line("Basic Life", "All"),
            Line("Extended Health Care", "Single Parent"),
            Line("Dependent Life", "All"),
            Line("Dental Care", "Couple"),
            Line("Short Term Disability", "All"),
            Line("Extended Health Care", "Single")
        ];
    }

    private static RateLine Line(string benefit, string tier)
    {
        return new RateLine
        {
            Benefit = benefit,
            TierOrClass = tier,
            AccountClass = "All Classes",
            Lives = 1,
            CurrentRate = 1m,
            CurrentPremium = 1m,
            RenewalRate = 1m,
            RenewalPremium = 1m,
            SourceReference = "Benefit display order fixture"
        };
    }

    private static string RowKey(RateLine row) => $"{row.Benefit}|{row.TierOrClass}";
}

using ClosedXML.Excel;
using RenewalReviewBuilder.App;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class ClaimsExperienceSourceSelectionTests
{
    [Fact]
    public void ClaimsExport_UsesOnlyAnExplicitlySelectedSourceFolder()
    {
        var existingFolder = Path.GetFullPath(AppContext.BaseDirectory);
        var missingFolder = Path.Combine(existingFolder, $"missing-{Guid.NewGuid():N}");

        Assert.Equal("", MainWindow.ResolveClaimsDataModelSourceFolder(""));
        Assert.Equal("", MainWindow.ResolveClaimsDataModelSourceFolder(missingFolder));
        Assert.Equal(existingFolder, MainWindow.ResolveClaimsDataModelSourceFolder($"  {existingFolder}  "));
    }

    [Fact]
    public void ClaimsExport_SelectsExplicitWorkbooksByStructureAndMatchingPeriod()
    {
        var testFolder = Path.Combine(
            AppContext.BaseDirectory,
            "buildtmp",
            $"claims-source-selection-{Guid.NewGuid():N}");
        var outputFolder = Path.Combine(testFolder, "output");
        var experiencePath = Path.Combine(testFolder, "Selected Insured Experience.xlsx");
        var matchingClaimsPath = Path.Combine(testFolder, "ClaimsServiceDetailReport - Health.xlsx");
        var wrongPeriodPath = Path.Combine(testFolder, "EHC and Dental 2025.xlsx");
        var unselectedPdfPath = Path.Combine(testFolder, "Monthly Unselected Premium Claims.pdf");

        try
        {
            Directory.CreateDirectory(testFolder);
            WriteExperienceWorkbook(experiencePath);
            WriteClaimsSummaryWorkbook(
                matchingClaimsPath,
                "Reporting period: Jan 1, 2026 to Feb 28, 2026",
                150m,
                300m);
            WriteClaimsSummaryWorkbook(
                wrongPeriodPath,
                "Reporting period: Jan 1, 2025 to Dec 31, 2025",
                9_999m,
                9_999m);
            File.WriteAllText(unselectedPdfPath, "This file was not selected.");

            var result = new ClaimsExperienceDataModelBuilder().Export(new ClaimsExperienceDataModelRequest
            {
                SourceFolder = "",
                SourceFiles = [experiencePath, wrongPeriodPath, matchingClaimsPath],
                OutputFolder = outputFolder,
                ReportMonthCount = 0
            });

            Assert.Equal(new DateOnly(2026, 1, 1), result.ExperiencePeriodStart);
            Assert.Equal(new DateOnly(2026, 2, 28), result.ExperiencePeriodEnd);
            Assert.Equal(2, result.ReportMonthCount);
            Assert.Equal("0155052", result.ClientNumber);
            Assert.Equal("Selected Client", result.ClientName);
            Assert.Equal(300m, result.Summaries.Single(row => row.BenefitLine == "EHC").PaidClaims);
            Assert.Equal(150m, result.Summaries.Single(row => row.BenefitLine == "Dental").PaidClaims);

            using var workbook = new XLWorkbook(result.WorkbookPath);
            var importRows = workbook.Table("tblImportLog").DataRange.Rows().ToList();
            Assert.Contains(importRows, row =>
                row.Field("Source Type").GetString() == "Excel experience report" &&
                row.Field("Source File").GetString() == Path.GetFileName(experiencePath));
            Assert.Equal(2, importRows.Count(row =>
                row.Field("Source File").GetString() == Path.GetFileName(matchingClaimsPath)));
            Assert.DoesNotContain(importRows, row =>
                row.Field("Source File").GetString() == Path.GetFileName(wrongPeriodPath));
            Assert.DoesNotContain(importRows, row =>
                row.Field("Source File").GetString() == Path.GetFileName(unselectedPdfPath));

            var ehcRows = workbook.Table("tblEHCCategory").DataRange.Rows().ToList();
            Assert.Equal(300m, ehcRows.Sum(row => row.Field("Paid Claims").GetValue<decimal>()));
            Assert.All(ehcRows, row => Assert.Equal(Path.GetFileName(matchingClaimsPath), row.Field("Source File").GetString()));

            var dentalRows = workbook.Table("tblDentalCategory").DataRange.Rows().ToList();
            Assert.Equal(150m, dentalRows.Sum(row => row.Field("Paid Claims").GetValue<decimal>()));
            Assert.All(dentalRows, row => Assert.Equal(Path.GetFileName(matchingClaimsPath), row.Field("Source File").GetString()));
        }
        finally
        {
            if (Directory.Exists(testFolder))
            {
                Directory.Delete(testFolder, recursive: true);
            }
        }
    }

    private static void WriteExperienceWorkbook(string path)
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Experience Combined");
        sheet.Cell("A2").Value = "Experience Report";
        sheet.Cell("A5").Value = "Policy: 0155052";
        sheet.Cell("A6").Value = "Period: Jan 1, 2026 to Feb 28, 2026";
        sheet.Cell("A7").Value = "Benefit";
        sheet.Cell("B7").Value = "EHC";
        sheet.Cell("E7").Value = "Dental";
        sheet.Cell("A8").Value = "Month";
        sheet.Cell("B8").Value = "Premiums";
        sheet.Cell("C8").Value = "Claims";
        sheet.Cell("E8").Value = "Premiums";
        sheet.Cell("F8").Value = "Claims";
        WriteExperienceRow(sheet, 9, new DateTime(2026, 1, 1), 1_000m, 100m, 800m, 50m);
        WriteExperienceRow(sheet, 10, new DateTime(2026, 2, 1), 1_100m, 200m, 900m, 100m);
        sheet.Cell("A11").Value = "Total";
        sheet.Cell("B11").Value = 2_100m;
        sheet.Cell("C11").Value = 300m;
        sheet.Cell("E11").Value = 1_700m;
        sheet.Cell("F11").Value = 150m;
        workbook.SaveAs(path);
    }

    private static void WriteExperienceRow(
        IXLWorksheet sheet,
        int row,
        DateTime month,
        decimal ehcPremium,
        decimal ehcClaims,
        decimal dentalPremium,
        decimal dentalClaims)
    {
        sheet.Cell(row, 1).Value = month;
        sheet.Cell(row, 1).Style.DateFormat.Format = "MMM-yyyy";
        sheet.Cell(row, 2).Value = ehcPremium;
        sheet.Cell(row, 3).Value = ehcClaims;
        sheet.Cell(row, 5).Value = dentalPremium;
        sheet.Cell(row, 6).Value = dentalClaims;
    }

    private static void WriteClaimsSummaryWorkbook(
        string path,
        string reportingPeriod,
        decimal dentalPaid,
        decimal ehcPaid)
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Summary");
        sheet.Cell("A3").Value = "Claim Service Detail by Relationship Report";
        sheet.Cell("A4").Value = "Selected Client";
        sheet.Cell("A5").Value = reportingPeriod;
        sheet.Cell("A7").Value = "Benefit";
        sheet.Cell("B7").Value = "Benefit category";
        sheet.Cell("C7").Value = "Service";
        sheet.Cell("D7").Value = "Number of occurrences";
        sheet.Cell("E7").Value = "Total amount submitted ($)";
        sheet.Cell("F7").Value = "Total amount eligible ($)";
        sheet.Cell("G7").Value = "Total amount paid ($)";
        WriteClaimsRow(sheet, 8, "Dental", "Basic Dental", "Preventative", 10m, dentalPaid);
        WriteClaimsRow(sheet, 9, "", "", "Subtotal", 10m, dentalPaid);
        WriteClaimsRow(sheet, 10, "Health", "Drugs", "Prescription Drugs", 20m, ehcPaid);
        WriteClaimsRow(sheet, 11, "", "", "Subtotal", 20m, ehcPaid);
        workbook.SaveAs(path);
    }

    private static void WriteClaimsRow(
        IXLWorksheet sheet,
        int row,
        string benefit,
        string category,
        string service,
        decimal occurrences,
        decimal paid)
    {
        sheet.Cell(row, 1).Value = benefit;
        sheet.Cell(row, 2).Value = category;
        sheet.Cell(row, 3).Value = service;
        sheet.Cell(row, 4).Value = occurrences;
        sheet.Cell(row, 5).Value = paid;
        sheet.Cell(row, 6).Value = paid;
        sheet.Cell(row, 7).Value = paid;
    }
}

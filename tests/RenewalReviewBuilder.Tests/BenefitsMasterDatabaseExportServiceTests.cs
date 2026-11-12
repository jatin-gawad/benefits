using ClosedXML.Excel;
using Microsoft.Data.Sqlite;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class BenefitsMasterDatabaseExportServiceTests
{
    [Fact]
    public async Task RebuildAsync_WithoutClientFilterExportsOnlyMostRecentlySavedClient()
    {
        var testFolder = Path.Combine(Path.GetTempPath(), $"sqlite-benefits-master-latest-{Guid.NewGuid():N}");
        var databasePath = Path.Combine(testFolder, "Benefits_Intelligence.db");
        try
        {
            Directory.CreateDirectory(testFolder);
            var database = new DatabaseService(databasePath);
            var older = await database.SavePlanDesignAsync(
                new ClientRecord { ClientName = "Older Client Ltd.", Carrier = "Test Carrier" },
                [],
                [PlanRow("Older Class", "$10,000")]);
            Assert.True(older.Succeeded, older.Message);

            await Task.Delay(50);
            var latest = await database.SavePlanDesignAsync(
                new ClientRecord { ClientName = "Latest Client Ltd.", ClientNumber = "LATEST-123", Carrier = "Test Carrier" },
                [],
                [PlanRow("Latest Class", "$20,000")]);
            Assert.True(latest.Succeeded, latest.Message);

            var result = await new BenefitsMasterDatabaseExportService(databasePath).RebuildAsync(testFolder);

            Assert.Equal(1, result.ClientCount);
            using var workbook = new XLWorkbook(result.MasterWorkbookPath);
            var summary = Assert.Single(
                workbook.Worksheets,
                sheet => sheet.Name.StartsWith("Summary_", StringComparison.OrdinalIgnoreCase));
            Assert.Equal("Latest Client Ltd.", summary.Cell(6, 1).GetString());
            var planDesign = workbook.Worksheet("PlanDesign");
            Assert.Contains("Latest Client Ltd.", planDesign.CellsUsed().Select(cell => cell.GetString()));
            Assert.DoesNotContain("Older Client Ltd.", planDesign.CellsUsed().Select(cell => cell.GetString()));
            Assert.Equal("LATEST-123", planDesign.Cell(4, 4).GetString());
        }
        finally
        {
            SqliteConnection.ClearAllPools();
            if (Directory.Exists(testFolder))
            {
                Directory.Delete(testFolder, recursive: true);
            }
        }
    }

    [Fact]
    public async Task RebuildAsync_RegeneratesDeletedWorkbookFromSQLite()
    {
        var testFolder = Path.Combine(Path.GetTempPath(), $"sqlite-benefits-master-{Guid.NewGuid():N}");
        var databasePath = Path.Combine(testFolder, "Benefits_Intelligence.db");
        try
        {
            Directory.CreateDirectory(testFolder);
            var database = new DatabaseService(databasePath);
            var save = await database.SavePlanDesignAsync(
                new ClientRecord
                {
                    ClientName = "SQLite Rebuild Test Ltd.",
                    Carrier = "Test Carrier",
                    EffectiveDate = new DateOnly(2026, 1, 1)
                },
                [],
                [
                    new PlanDesignRecord
                    {
                        ClassName = "Class 01 - All Eligible Employees",
                        BenefitCategory = "Life / AD&D",
                        BenefitName = "Life / AD&D",
                        CoverageDescription = "$50,000",
                        MaximumAmount = "$50,000",
                        TerminationAge = "Age 70; retirement",
                        SourceFileName = "sqlite-source-booklet.pdf",
                        SourceSection = "Schedule of Benefits",
                        SourcePage = 3,
                        DateLoaded = new DateTime(2026, 8, 21, 10, 30, 0)
                    },
                    new PlanDesignRecord
                    {
                        ClassName = "Class 01 - All Eligible Employees",
                        BenefitCategory = "Dental Care",
                        BenefitName = "Dental Care",
                        CoverageDescription = "Basic Services: 80%; Major Restorative Services: 50%",
                        MaximumAmount = "$2,000 per calendar year",
                        SourceFileName = "sqlite-source-booklet.pdf",
                        SourceSection = "Schedule of Benefits",
                        SourcePage = 22,
                        DateLoaded = new DateTime(2026, 8, 21, 10, 30, 0)
                    }
                ]);

            Assert.True(save.Succeeded, save.Message);
            var exporter = new BenefitsMasterDatabaseExportService(databasePath);
            var first = await exporter.RebuildAsync(testFolder);

            Assert.True(File.Exists(first.MasterWorkbookPath));
            Assert.Equal(1, first.ClientCount);
            Assert.Equal(2, first.PlanDesignRowCount);
            Assert.True(first.SummarySheetCount >= 1);
            AssertWorkbook(first.MasterWorkbookPath);

            File.Delete(first.MasterWorkbookPath);
            Assert.False(File.Exists(first.MasterWorkbookPath));

            var rebuilt = await exporter.RebuildAsync(testFolder);

            Assert.True(File.Exists(rebuilt.MasterWorkbookPath));
            Assert.Equal(first.PlanDesignRowCount, rebuilt.PlanDesignRowCount);
            Assert.Equal(first.SummarySheetCount, rebuilt.SummarySheetCount);
            AssertWorkbook(rebuilt.MasterWorkbookPath);
        }
        finally
        {
            SqliteConnection.ClearAllPools();
            if (Directory.Exists(testFolder))
            {
                Directory.Delete(testFolder, recursive: true);
            }
        }
    }

    private static void AssertWorkbook(string path)
    {
        using var workbook = new XLWorkbook(path);
        var summary = Assert.Single(
            workbook.Worksheets,
            sheet => sheet.Name.StartsWith("Summary_", StringComparison.OrdinalIgnoreCase));
        Assert.Equal("SQLite Rebuild Test Ltd.", summary.Cell(6, 1).GetString());
        Assert.Contains("Life / AD&D", summary.CellsUsed().Select(cell => cell.GetString()));
        Assert.Contains(
            summary.CellsUsed(),
            cell => cell.GetString().Contains("$50,000", StringComparison.OrdinalIgnoreCase));

        var planDesign = workbook.Worksheet("PlanDesign");
        Assert.Contains("$50,000", planDesign.CellsUsed().Select(cell => cell.GetString()));
        Assert.NotNull(workbook.Worksheet("ProcessingLog"));
        Assert.NotNull(workbook.Worksheet("Booklet_Extraction_Review"));
        Assert.NotNull(workbook.Worksheet("Schedule_of_Benefits"));
        Assert.NotNull(workbook.Worksheet("Schedule_PPT_Export"));
        Assert.Contains(
            workbook.Worksheet("ProcessingLog").CellsUsed(),
            cell => cell.GetString().Equals("Benefits_Intelligence.db", StringComparison.OrdinalIgnoreCase));
    }

    private static PlanDesignRecord PlanRow(string className, string maximum) => new()
    {
        ClassName = className,
        BenefitCategory = "Life / AD&D",
        BenefitName = "Life / AD&D",
        CoverageDescription = maximum,
        MaximumAmount = maximum,
        SourceFileName = $"{className}.pdf",
        SourceSection = "Benefit Summary",
        SourcePage = 1,
        DateLoaded = DateTime.Now
    };
}

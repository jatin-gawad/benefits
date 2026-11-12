using ClosedXML.Excel;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class ClassComparisonExportServiceTests
{
    [Fact]
    public void ExportToExcel_Writes_Selected_Classes_Sections_Costs_Scores_And_History()
    {
        var testFolder = Path.Combine(Path.GetTempPath(), "ClassComparisonExportTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(testFolder);
        var outputPath = Path.Combine(testFolder, "comparison.xlsx");

        try
        {
            var model = ClassComparisonEngine.Build(new ClassComparisonBuildRequest
            {
                ClientId = 42,
                ClientName = "Export Test Client",
                EffectiveDateText = "Jan 1, 2026",
                BenefitRows =
                [
                    Benefit("Class A", "Eligibility", "Waiting Period", "3 months"),
                    Benefit("Class B", "Eligibility", "Waiting Period", "1 month"),
                    Benefit("Class A", "Extended Health Care", "Paramedical Maximum", "$500"),
                    Benefit("Class B", "Extended Health Care", "Paramedical Maximum", "$1,000")
                ],
                CostRows =
                [
                    new ClassCostSourceRow("Class A", 10, 1000m, 100m, 1100m, "billing.xlsx"),
                    new ClassCostSourceRow("Class B", 12, 1300m, 120m, 1420m, "billing.xlsx")
                ],
                Notes =
                [
                    new ClassComparisonNoteRow
                    {
                        NoteId = 1,
                        ClientId = 42,
                        NoteText = "Confirm the waiting period.",
                        ClassName = "Class A",
                        Author = "Test Analyst",
                        CreatedAt = new DateTime(2026, 7, 12, 9, 0, 0)
                    }
                ],
                History =
                [
                    new ClassComparisonHistoryRow
                    {
                        HistoryId = 1,
                        DateText = "Jul 12, 2026 9:00 AM",
                        Author = "Test Analyst",
                        ClassName = "Class B",
                        FieldName = "Recommended class",
                        PreviousValue = "Class A",
                        NewValue = "Class B",
                        Reason = "Comparison preference updated",
                        SourceDocument = "Class Comparison workspace"
                    }
                ],
                Preferences = new ClassComparisonPreferenceState
                {
                    BaselineClassName = "Class A",
                    SelectedClassNames = ["Class A", "Class B"],
                    RecommendedClassName = "Class B"
                }
            });

            var result = new ClassComparisonExportService().ExportToExcel(model, outputPath);

            Assert.True(File.Exists(outputPath));
            Assert.Equal(2, result.SelectedClassCount);
            Assert.True(result.ComparisonRowCount >= 2);

            using var workbook = new XLWorkbook(outputPath);
            Assert.Equal(
                ["Class Comparison", "Class Summary", "Cost Analysis", "Rate Comparison", "Notes & History"],
                workbook.Worksheets.Select(sheet => sheet.Name).ToArray());

            var comparisonSheet = workbook.Worksheet("Class Comparison");
            Assert.Equal("Export Test Client", comparisonSheet.Cell(2, 1).GetString());
            Assert.Equal("Class A", comparisonSheet.Cell(8, 2).GetString());
            Assert.Equal("Class B", comparisonSheet.Cell(8, 3).GetString());
            Assert.Contains(
                comparisonSheet.Column(1).CellsUsed().Select(cell => cell.GetString()),
                value => value.Contains("Eligibility", StringComparison.OrdinalIgnoreCase));
            Assert.Contains(
                comparisonSheet.Column(1).CellsUsed().Select(cell => cell.GetString()),
                value => value.Contains("Paramedical Maximum", StringComparison.Ordinal));

            var costSheet = workbook.Worksheet("Cost Analysis");
            Assert.Equal("Class A", costSheet.Cell(2, 1).GetString());
            Assert.Equal(13200m, costSheet.Cell(2, 8).GetValue<decimal>());

            var notesSheet = workbook.Worksheet("Notes & History");
            Assert.Contains(notesSheet.CellsUsed(), cell => cell.GetString().Contains("Confirm the waiting period", StringComparison.Ordinal));
            Assert.Contains(notesSheet.CellsUsed(), cell => cell.GetString() == "Recommended class");
        }
        finally
        {
            if (Directory.Exists(testFolder))
            {
                Directory.Delete(testFolder, recursive: true);
            }
        }
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

using ClosedXML.Excel;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed class PlanDesignComparisonExportService
{
    public PlanDesignComparisonExportResult ExportToExcel(PlanDesignComparisonWorkspaceViewModel model, string outputPath)
    {
        ArgumentNullException.ThrowIfNull(model);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);

        var rows = model.Groups
            .OrderBy(group => group.Order)
            .SelectMany(group => group.Rows)
            .Where(row => row.ScenarioKey.Equals(model.SelectedScenarioKey, StringComparison.OrdinalIgnoreCase))
            .ToList();
        if (rows.Count == 0)
        {
            throw new InvalidOperationException("No selected plan-design scenario is available to export.");
        }

        Directory.CreateDirectory(Path.GetDirectoryName(outputPath) ?? ".");
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("Plan Design Comparison");
        var headers = new[] { "Category", "Benefit", "Provision", "Current Plan", "Selected Plan / Scenario", "Difference", "Cost Impact", "Value / Recommendation", "Evidence" };
        for (var column = 0; column < headers.Length; column++)
        {
            sheet.Cell(1, column + 1).Value = headers[column];
        }

        var rowNumber = 2;
        foreach (var row in rows)
        {
            sheet.Cell(rowNumber, 1).Value = row.Category;
            sheet.Cell(rowNumber, 2).Value = row.BenefitName;
            sheet.Cell(rowNumber, 3).Value = row.ProvisionName;
            sheet.Cell(rowNumber, 4).Value = row.CurrentValue;
            sheet.Cell(rowNumber, 5).Value = row.SelectedValue;
            sheet.Cell(rowNumber, 6).Value = row.DifferenceStatus;
            sheet.Cell(rowNumber, 7).Value = row.CostImpactText;
            sheet.Cell(rowNumber, 8).Value = row.ValueRecommendation;
            sheet.Cell(rowNumber, 9).Value = row.EvidenceText;
            rowNumber++;
        }

        var headerRange = sheet.Range(1, 1, 1, headers.Length);
        headerRange.Style.Font.Bold = true;
        headerRange.Style.Fill.BackgroundColor = XLColor.FromHtml("F5F3FF");
        sheet.SheetView.FreezeRows(1);
        sheet.Columns().AdjustToContents(8, 58);
        sheet.Range(1, 1, rowNumber - 1, headers.Length).Style.Alignment.WrapText = true;
        workbook.SaveAs(outputPath);
        return new PlanDesignComparisonExportResult(outputPath, rows.Count);
    }
}

public sealed record PlanDesignComparisonExportResult(string OutputPath, int RowCount);

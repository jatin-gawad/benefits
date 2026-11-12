using ClosedXML.Excel;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed class ClassComparisonExportService
{
    public ClassComparisonExportResult ExportToExcel(ClassComparisonWorkspaceViewModel model, string outputPath)
    {
        ArgumentNullException.ThrowIfNull(model);
        if (!model.HasEnoughClasses || !model.HasComparableFeatures)
        {
            throw new InvalidOperationException("Load two classes with comparable plan-design features before exporting.");
        }

        if (string.IsNullOrWhiteSpace(outputPath))
        {
            throw new ArgumentException("Choose an Excel output path.", nameof(outputPath));
        }

        var fullPath = Path.GetFullPath(outputPath);
        var folder = Path.GetDirectoryName(fullPath);
        if (!string.IsNullOrWhiteSpace(folder))
        {
            Directory.CreateDirectory(folder);
        }

        using var workbook = new XLWorkbook();
        WriteComparisonSheet(workbook, model);
        WriteClassSummarySheet(workbook, model);
        WriteCostSheet(workbook, model);
        WriteRateSheet(workbook, model);
        WriteNotesAndHistorySheet(workbook, model);
        workbook.SaveAs(fullPath);

        return new ClassComparisonExportResult(fullPath, 2, model.FilteredComparisonRows.Count);
    }

    private static void WriteComparisonSheet(XLWorkbook workbook, ClassComparisonWorkspaceViewModel model)
    {
        var sheet = workbook.Worksheets.Add("Class Comparison");
        sheet.Cell(1, 1).Value = "Class Comparison";
        sheet.Cell(1, 1).Style.Font.Bold = true;
        sheet.Cell(1, 1).Style.Font.FontSize = 18;
        sheet.Cell(2, 1).Value = model.ClientName;
        sheet.Cell(3, 1).Value = "Effective Date";
        sheet.Cell(3, 2).Value = model.EffectiveDateText;
        sheet.Cell(4, 1).Value = "Baseline Class";
        sheet.Cell(4, 2).Value = model.BaselineClass?.DisplayName ?? model.BaselineClassName;
        sheet.Cell(5, 1).Value = "Comparison Class";
        sheet.Cell(5, 2).Value = model.SelectedComparisonClass?.DisplayName ?? model.ComparisonClassName;
        sheet.Cell(6, 1).Value = "Generated";
        sheet.Cell(6, 2).Value = DateTime.Now;
        sheet.Cell(6, 2).Style.DateFormat.Format = "mmm d, yyyy h:mm AM/PM";

        const int headerRow = 8;
        sheet.Cell(headerRow, 1).Value = "Benefit";
        sheet.Cell(headerRow, 2).Value = model.BaselineClass?.DisplayName ?? model.BaselineClassName;
        sheet.Cell(headerRow, 3).Value = model.SelectedComparisonClass?.DisplayName ?? model.ComparisonClassName;
        sheet.Cell(headerRow, 4).Value = "Difference";
        sheet.Cell(headerRow, 5).Value = "Evidence / Source";
        const int column = 5;
        StyleHeader(sheet.Range(headerRow, 1, headerRow, column));

        var rowNumber = headerRow + 1;
        foreach (var row in model.FilteredComparisonRows)
        {
            sheet.Cell(rowNumber, 1).Value = $"{row.Category} — {row.DisplayName}";
            sheet.Cell(rowNumber, 2).Value = row.BaselineDisplayValue;
            sheet.Cell(rowNumber, 3).Value = row.ComparisonDisplayValue;
            sheet.Cell(rowNumber, 4).Value = row.DifferenceLabel;
            sheet.Cell(rowNumber, 5).Value = string.Join(" | ", row.EvidenceReferences);
            if (model.HighlightDifferences && row.DifferenceKind != ClassComparisonDifferenceKind.Same)
            {
                sheet.Cell(rowNumber, 3).Style.Fill.BackgroundColor = XLColor.FromHtml("#F5F3FF");
            }
            rowNumber++;
        }

        var used = sheet.Range(headerRow, 1, Math.Max(rowNumber - 1, headerRow), column);
        used.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
        used.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        used.Style.Border.InsideBorderColor = XLColor.FromHtml("#D8E1EC");
        used.Style.Border.OutsideBorderColor = XLColor.FromHtml("#D8E1EC");
        used.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
        used.Style.Alignment.WrapText = true;
        sheet.SheetView.Freeze(headerRow, 1);
        sheet.Column(1).Width = 30;
        for (var index = 2; index <= column; index++)
        {
            sheet.Column(index).Width = index == column ? 20 : 34;
        }
        sheet.PageSetup.PageOrientation = XLPageOrientation.Landscape;
        sheet.PageSetup.FitToPages(1, 0);
        sheet.PageSetup.SetRowsToRepeatAtTop(headerRow, headerRow);
    }

    private static void WriteClassSummarySheet(XLWorkbook workbook, ClassComparisonWorkspaceViewModel model)
    {
        var sheet = workbook.Worksheets.Add("Class Summary");
        var headers = new[]
        {
            "Class",
            "Status",
            "Employees",
            "Annual Employer Cost",
            "Annual Employee Cost",
            "Annual Total Cost",
            "% of Client Cost",
            "Projected Change",
            "Value Score",
            "Rank"
        };
        WriteHeaders(sheet, headers);
        var scoreLookup = model.ValueScores.ToDictionary(item => item.ClassName, StringComparer.OrdinalIgnoreCase);
        var rowNumber = 2;
        foreach (var classCard in PairClasses(model))
        {
            scoreLookup.TryGetValue(classCard.ClassName, out var score);
            sheet.Cell(rowNumber, 1).Value = classCard.DisplayName;
            sheet.Cell(rowNumber, 2).Value = classCard.StatusLabel;
            SetNullableNumber(sheet.Cell(rowNumber, 3), classCard.EmployeeCount);
            SetNullableMoney(sheet.Cell(rowNumber, 4), classCard.AnnualEmployerCost);
            SetNullableMoney(sheet.Cell(rowNumber, 5), classCard.AnnualEmployeeCost);
            SetNullableMoney(sheet.Cell(rowNumber, 6), classCard.AnnualTotalCost);
            SetNullablePercent(sheet.Cell(rowNumber, 7), classCard.TotalPlanCostPercent);
            SetNullablePercent(sheet.Cell(rowNumber, 8), classCard.ProjectedChangePercent);
            SetNullableNumber(sheet.Cell(rowNumber, 9), score?.Score);
            SetNullableNumber(sheet.Cell(rowNumber, 10), score?.Rank);
            rowNumber++;
        }
        FormatTable(sheet, headers.Length, rowNumber - 1, "ClassSummaryTable");
    }

    private static void WriteCostSheet(XLWorkbook workbook, ClassComparisonWorkspaceViewModel model)
    {
        var sheet = workbook.Worksheets.Add("Cost Analysis");
        var headers = new[]
        {
            "Class",
            "Employees",
            "Monthly Employer Cost",
            "Monthly Employee Cost",
            "Monthly Total Cost",
            "Annual Employer Cost",
            "Annual Employee Cost",
            "Annual Total Cost",
            "Cost per Employee",
            "% of Client Cost",
            "Projected Change",
            "Source"
        };
        WriteHeaders(sheet, headers);
        var rowNumber = 2;
        var pairNames = PairClasses(model).Select(item => item.ClassName).ToHashSet(StringComparer.OrdinalIgnoreCase);
        foreach (var row in model.CostSummaries.Where(item => pairNames.Contains(item.ClassName)))
        {
            sheet.Cell(rowNumber, 1).Value = row.DisplayName;
            SetNullableNumber(sheet.Cell(rowNumber, 2), row.EmployeeCount);
            SetNullableMoney(sheet.Cell(rowNumber, 3), row.MonthlyEmployerCost);
            SetNullableMoney(sheet.Cell(rowNumber, 4), row.MonthlyEmployeeCost);
            SetNullableMoney(sheet.Cell(rowNumber, 5), row.MonthlyTotalCost);
            SetNullableMoney(sheet.Cell(rowNumber, 6), row.AnnualEmployerCost);
            SetNullableMoney(sheet.Cell(rowNumber, 7), row.AnnualEmployeeCost);
            SetNullableMoney(sheet.Cell(rowNumber, 8), row.AnnualTotalCost);
            SetNullableMoney(sheet.Cell(rowNumber, 9), row.CostPerEmployee);
            SetNullablePercent(sheet.Cell(rowNumber, 10), row.TotalClientCostPercent);
            SetNullablePercent(sheet.Cell(rowNumber, 11), row.ProjectedChangePercent);
            sheet.Cell(rowNumber, 12).Value = row.SourceDocument;
            rowNumber++;
        }
        FormatTable(sheet, headers.Length, rowNumber - 1, "ClassCostTable");
    }

    private static void WriteRateSheet(XLWorkbook workbook, ClassComparisonWorkspaceViewModel model)
    {
        var sheet = workbook.Worksheets.Add("Rate Comparison");
        var headers = new[]
        {
            "Benefit",
            "Class",
            "Current Rate",
            "Proposed Rate",
            "Negotiated Rate",
            "Variance",
            "Variance %",
            "Rate Basis",
            "Effective Date",
            "Source"
        };
        WriteHeaders(sheet, headers);
        var rowNumber = 2;
        var pairNames = PairClasses(model).Select(item => item.ClassName).ToHashSet(StringComparer.OrdinalIgnoreCase);
        foreach (var row in model.RateRows.Where(item => pairNames.Contains(item.ClassName)))
        {
            sheet.Cell(rowNumber, 1).Value = row.Benefit;
            sheet.Cell(rowNumber, 2).Value = row.ClassDisplayName;
            SetNullableNumber(sheet.Cell(rowNumber, 3), row.CurrentRate);
            SetNullableNumber(sheet.Cell(rowNumber, 4), row.ProposedRate);
            SetNullableNumber(sheet.Cell(rowNumber, 5), row.NegotiatedRate);
            SetNullableNumber(sheet.Cell(rowNumber, 6), row.Variance);
            SetNullablePercent(sheet.Cell(rowNumber, 7), row.VariancePercent);
            sheet.Cell(rowNumber, 8).Value = row.RateBasis;
            sheet.Cell(rowNumber, 9).Value = row.EffectiveDateText;
            sheet.Cell(rowNumber, 10).Value = row.SourceDocument;
            rowNumber++;
        }
        FormatTable(sheet, headers.Length, rowNumber - 1, "ClassRateTable");
    }

    private static void WriteNotesAndHistorySheet(XLWorkbook workbook, ClassComparisonWorkspaceViewModel model)
    {
        var sheet = workbook.Worksheets.Add("Notes & History");
        sheet.Cell(1, 1).Value = "Analyst Notes";
        sheet.Cell(1, 1).Style.Font.Bold = true;
        var noteHeaders = new[] { "Date", "Author", "Class", "Note" };
        for (var index = 0; index < noteHeaders.Length; index++)
        {
            sheet.Cell(2, index + 1).Value = noteHeaders[index];
        }
        StyleHeader(sheet.Range(2, 1, 2, noteHeaders.Length));
        var rowNumber = 3;
        foreach (var note in model.Notes)
        {
            sheet.Cell(rowNumber, 1).Value = note.CreatedAtText;
            sheet.Cell(rowNumber, 2).Value = note.Author;
            sheet.Cell(rowNumber, 3).Value = note.ClassName;
            sheet.Cell(rowNumber, 4).Value = note.NoteText;
            rowNumber++;
        }

        rowNumber += 2;
        sheet.Cell(rowNumber, 1).Value = "Audit History";
        sheet.Cell(rowNumber, 1).Style.Font.Bold = true;
        rowNumber++;
        var historyHeaders = new[] { "Date", "Author", "Class", "Field", "Previous", "New", "Reason", "Source" };
        for (var index = 0; index < historyHeaders.Length; index++)
        {
            sheet.Cell(rowNumber, index + 1).Value = historyHeaders[index];
        }
        StyleHeader(sheet.Range(rowNumber, 1, rowNumber, historyHeaders.Length));
        rowNumber++;
        foreach (var history in model.History)
        {
            sheet.Cell(rowNumber, 1).Value = history.DateText;
            sheet.Cell(rowNumber, 2).Value = history.Author;
            sheet.Cell(rowNumber, 3).Value = history.ClassName;
            sheet.Cell(rowNumber, 4).Value = history.FieldName;
            sheet.Cell(rowNumber, 5).Value = history.PreviousValue;
            sheet.Cell(rowNumber, 6).Value = history.NewValue;
            sheet.Cell(rowNumber, 7).Value = history.Reason;
            sheet.Cell(rowNumber, 8).Value = history.SourceDocument;
            rowNumber++;
        }
        sheet.Columns(1, 8).AdjustToContents(8, 48);
        sheet.Column(4).Width = 60;
        sheet.Column(7).Width = 48;
        sheet.RangeUsed()?.Style.Alignment.SetWrapText();
    }

    private static void WriteHeaders(IXLWorksheet sheet, IReadOnlyList<string> headers)
    {
        for (var index = 0; index < headers.Count; index++)
        {
            sheet.Cell(1, index + 1).Value = headers[index];
        }
        StyleHeader(sheet.Range(1, 1, 1, headers.Count));
    }

    private static void StyleHeader(IXLRange range)
    {
        range.Style.Fill.BackgroundColor = XLColor.FromHtml("#0B2341");
        range.Style.Font.FontColor = XLColor.White;
        range.Style.Font.Bold = true;
        range.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
    }

    private static void FormatTable(IXLWorksheet sheet, int columnCount, int lastRow, string tableName)
    {
        if (lastRow >= 2)
        {
            sheet.Range(1, 1, lastRow, columnCount).CreateTable(tableName);
        }
        sheet.SheetView.FreezeRows(1);
        sheet.Columns(1, columnCount).AdjustToContents(8, 42);
        sheet.RangeUsed()?.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Top);
        sheet.RangeUsed()?.Style.Alignment.SetWrapText();
    }

    private static XLColor StatusColor(ClassComparisonStatus status)
    {
        return status switch
        {
            ClassComparisonStatus.Better => XLColor.FromHtml("#DCFCE7"),
            ClassComparisonStatus.Worse => XLColor.FromHtml("#FEE2E2"),
            ClassComparisonStatus.Different => XLColor.FromHtml("#DBEAFE"),
            ClassComparisonStatus.Missing => XLColor.FromHtml("#F8FAFC"),
            _ => XLColor.White
        };
    }

    private static IReadOnlyList<ClassComparisonClassCard> PairClasses(ClassComparisonWorkspaceViewModel model)
    {
        return new[] { model.BaselineClass, model.SelectedComparisonClass }
            .OfType<ClassComparisonClassCard>()
            .DistinctBy(item => item.ClassName, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static void SetNullableMoney(IXLCell cell, decimal? value)
    {
        if (value.HasValue)
        {
            cell.Value = value.Value;
            cell.Style.NumberFormat.Format = "$#,##0";
        }
        else
        {
            cell.Value = "Not available";
        }
    }

    private static void SetNullablePercent(IXLCell cell, decimal? value)
    {
        if (value.HasValue)
        {
            cell.Value = value.Value;
            cell.Style.NumberFormat.Format = "0.0%";
        }
        else
        {
            cell.Value = "Not available";
        }
    }

    private static void SetNullableNumber(IXLCell cell, decimal? value)
    {
        if (value.HasValue)
        {
            cell.Value = value.Value;
            cell.Style.NumberFormat.Format = "0.####";
        }
        else
        {
            cell.Value = "Not available";
        }
    }

    private static void SetNullableNumber(IXLCell cell, int? value)
    {
        if (value.HasValue)
        {
            cell.Value = value.Value;
            cell.Style.NumberFormat.Format = "#,##0";
        }
        else
        {
            cell.Value = "Not available";
        }
    }
}

public sealed record ClassComparisonExportResult(string OutputPath, int SelectedClassCount, int ComparisonRowCount);

using System.Globalization;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using RenewalReviewBuilder.App.Models;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using S = DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace RenewalReviewBuilder.App.Services;

public sealed class ExcelReportWriter
{
    private readonly IReadOnlyList<UnderwritingReferenceRate> _underwritingReferenceRates;
    private static readonly XLColor WestlandTeal = XLColor.FromHtml("#0E7C7B");
    private static readonly XLColor WestlandGold = XLColor.FromHtml("#C69C3F");
    private static readonly XLColor SoftFill = XLColor.FromHtml("#F4F7F9");
    private static readonly XLColor WarningFill = XLColor.FromHtml("#FFF3CD");
    private static readonly XLColor CriticalFill = XLColor.FromHtml("#F8D7DA");
    private static readonly XLColor CheckGreen = XLColor.FromHtml("#00B050");
    private static readonly XLColor AccountFill = XLColor.FromHtml("#D9D9C3");
    private static readonly XLColor NegotiatedFill = XLColor.FromHtml("#92D050");
    private static readonly XLColor LineRatesDarkTeal = XLColor.FromHtml("#004650");
    private static readonly XLColor LineRatesLightTeal = XLColor.FromHtml("#EAF8F8");
    private static readonly XLColor LineRatesMediumTeal = XLColor.FromHtml("#CDECEF");
    private static readonly XLColor LineRatesGridTeal = XLColor.FromHtml("#8DD3C7");
    private static readonly XLColor LineRatesCurrentPlanFill = XLColor.FromHtml("#F2F2F2");
    private static readonly XLColor LineRatesCurrentPlanHeader = XLColor.FromHtml("#808080");

    public ExcelReportWriter(IReadOnlyList<UnderwritingReferenceRate>? underwritingReferenceRates = null)
    {
        _underwritingReferenceRates = underwritingReferenceRates is { Count: > 0 }
            ? underwritingReferenceRates
            : UnderwritingReferenceRateCatalog.Defaults;
    }

    public string Write(AnalysisResult analysis, string outputFolder)
    {
        Directory.CreateDirectory(outputFolder);
        var path = Path.Combine(outputFolder, "Renewal_Review.xlsx");

        using var workbook = new XLWorkbook();
        WriteDocumentDiscovery(workbook, analysis);
        WriteSummary(workbook, analysis);
        WritePreviousRenewals(workbook, analysis);
        WriteAssurePlusExperience(workbook, analysis);
        WriteAsoBillingActivity(workbook, analysis);
        WriteFindings(workbook, analysis);
        WriteRates(workbook, analysis);
        WriteNetAndGrossRates(workbook, analysis);
        WriteNems(workbook, analysis);
        WritePowerPointExportTables(workbook, analysis);
        WriteRateTableLineRatesFormatted(workbook, analysis);
        WriteHealthRenewalFormulaAnalysis(workbook, analysis);
        WriteCanadaLifeComponentAnalyses(workbook, analysis);
        WriteDentalRenewalFormulaAnalysis(workbook, analysis);
        WriteClaimsAnalysisEhc(workbook, analysis);
        WriteClaimsAnalysisDental(workbook, analysis);
        WriteClaimsAnalysisParamedical(workbook, analysis);
        WriteDemographicAnalysis(
            workbook,
            "Life Analysis",
            "Life Insurance and ASI Demographic Analysis",
            analysis.Pdf.LifeDemographics,
            GetLifeAgeBands(analysis.Pdf.LifeDemographics),
            "Prior",
            DemographicAgeBandNormalizer.NormalizeLifeAgeBand,
            "'Mortality Master'!L17");
        WriteDemographicAnalysis(
            workbook,
            "LTD Analysis",
            "Long-Term Disability Demographic Analysis",
            analysis.Pdf.LtdDemographics,
            GetLtdAgeBands(analysis.Pdf.LtdDemographics),
            "Previous",
            DemographicAgeBandNormalizer.NormalizeLtdAgeBand,
            "'Morbidity Master'!L16");
        WriteUnderwritingReferenceMaster(
            workbook,
            "Mortality Master",
            UnderwritingRateTypes.Mortality,
            "Life Analysis",
            GetLifeAgeBands(analysis.Pdf.LifeDemographics),
            1000m);
        WriteUnderwritingReferenceMaster(
            workbook,
            "Morbidity Master",
            UnderwritingRateTypes.Morbidity,
            "LTD Analysis",
            GetLtdAgeBands(analysis.Pdf.LtdDemographics),
            100m);
        WriteSunLifeExperience(workbook, analysis);
        WriteExperience(workbook, analysis);
        WriteClaimTrends(workbook, analysis, "EHC Trends", "Health", analysis.Workbook.EhcSpendingTrends);
        WriteClaimTrends(workbook, analysis, "Dental Trends", "Dental", analysis.Workbook.DentalSpendingTrends);
        WriteDrugs(workbook, analysis);
        WriteClaimsDistribution(workbook, analysis);
        WriteBenchmarks(workbook, analysis);
        WriteSourceChecks(workbook, analysis);

        foreach (var sheet in workbook.Worksheets)
        {
            var used = sheet.RangeUsed();
            if (used is not null)
            {
                used.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                used.Style.Border.InsideBorder = XLBorderStyleValues.Hair;
                used.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                sheet.SheetView.FreezeRows(1);
                sheet.Columns().AdjustToContents(8, 48);
                if ((sheet.Name.Equals("EHC Trends", StringComparison.OrdinalIgnoreCase) ||
                     sheet.Name.Equals("Dental Trends", StringComparison.OrdinalIgnoreCase)) &&
                    sheet.Cell(1, 3).GetString().Equals("Prior Period", StringComparison.OrdinalIgnoreCase))
                {
                    ApplyPairedClaimTrendLayout(sheet);
                }
            }
        }

        foreach (var sheetName in new[] { "CLhealthanalysis", "CLdruganalysis", "CLvisionanalysis" })
        {
            if (workbook.TryGetWorksheet(sheetName, out var componentSheet))
            {
                FormatCanadaLifeComponentSheet(componentSheet);
            }
        }

        if (workbook.TryGetWorksheet("ClaimsAnalysis_EHC", out var claimsAnalysisEhc))
        {
            claimsAnalysisEhc.Column(1).Width = 30;
            claimsAnalysisEhc.Columns(2, 4).Width = 18;
        }

        if (workbook.TryGetWorksheet("PowerPoint Export Tables", out var powerPointExportTables))
        {
            powerPointExportTables.Column(1).Width = 24;
            powerPointExportTables.Column(2).Width = 18;
            powerPointExportTables.Column(3).Width = 12;
            powerPointExportTables.Column(4).Width = 12;
            powerPointExportTables.Column(5).Width = 8;
        }

        ApplyPremiumIllustrationExportFinishing(workbook);

        if (workbook.TryGetWorksheet("NEMs", out var nems))
        {
            nems.Column(1).Width = 24;
            nems.Columns(2, 3).Width = 18;
            nems.Column(4).Width = 60;
        }

        if (workbook.TryGetWorksheet("AssurePLUS Experience", out var assurePlusExperience))
        {
            assurePlusExperience.Column(1).Width = 22;
            assurePlusExperience.Column(2).Width = 14;
            assurePlusExperience.Columns(3, 5).Width = 20;
            assurePlusExperience.Column(6).Width = 18;
            assurePlusExperience.Column(7).Width = 20;
            assurePlusExperience.Columns(8, 9).Width = 18;
            assurePlusExperience.Column(10).Width = 16;
            assurePlusExperience.Column(11).Width = 12;
            assurePlusExperience.Column(12).Width = 55;
        }

        if (workbook.TryGetWorksheet("Rate Reconciliation", out var rateReconciliation))
        {
            rateReconciliation.Columns(2, 17).Width = 10;
        }

        if (workbook.TryGetWorksheet("Net and Gross Rates", out var netAndGrossRates))
        {
            netAndGrossRates.Column(1).Width = 34;
            netAndGrossRates.Columns(2, 7).Width = 14;
            netAndGrossRates.Column(8).Width = 48;
        }

        if (workbook.TryGetWorksheet("ClaimsAnalysis_Dental", out var claimsAnalysisDental))
        {
            claimsAnalysisDental.Column(1).Width = 30;
            claimsAnalysisDental.Columns(2, 4).Width = 18;
        }

        if (workbook.TryGetWorksheet("ClaimsAnalysis_Paramedical", out var claimsAnalysisParamedical))
        {
            claimsAnalysisParamedical.Column(1).Width = 34;
            claimsAnalysisParamedical.Columns(2, 4).Width = 18;
        }

        if (workbook.TryGetWorksheet("HealthRenewal_FormulaAnalysis", out var healthFormula))
        {
            FormatHealthRenewalFormulaSheet(healthFormula);
        }

        if (workbook.TryGetWorksheet("DentalRenewal_FormulaAnalysis", out var dentalFormula))
        {
            FormatHealthRenewalFormulaSheet(dentalFormula);
        }

        if (workbook.TryGetWorksheet("RateTable_LineRates_Formatted", out var formattedLineRates))
        {
            FormatRateTableLineRatesFormattedSheet(formattedLineRates);
        }

        FormatDemographicAnalysisSheet(workbook, "Life Analysis");
        FormatDemographicAnalysisSheet(workbook, "LTD Analysis");
        FormatUnderwritingReferenceMasterSheet(workbook, "Mortality Master");
        FormatUnderwritingReferenceMasterSheet(workbook, "Morbidity Master");

        workbook.SaveAs(path);
        AddEhcSpendingTrendChart(path, analysis.Workbook.EhcSpendingTrends);
        AddDentalSpendingTrendChart(path, analysis.Workbook.DentalSpendingTrends);
        AddParamedicalServiceTrendChart(path, BuildParamedicalAnalysisTrends(analysis.Workbook.ParamedicalServiceTrends));
        ForceWorkbookRecalculation(path);
        return path;
    }

    private static void ForceWorkbookRecalculation(string workbookPath)
    {
        using var document = SpreadsheetDocument.Open(workbookPath, true);
        var workbook = document.WorkbookPart?.Workbook;
        if (workbook is null)
        {
            return;
        }

        workbook.CalculationProperties = new S.CalculationProperties
        {
            CalculationMode = S.CalculateModeValues.Auto,
            FullCalculationOnLoad = true,
            ForceFullCalculation = true
        };

        var calculationChain = document.WorkbookPart?.CalculationChainPart;
        if (calculationChain is not null)
        {
            document.WorkbookPart!.DeletePart(calculationChain);
        }

        workbook.Save();
    }

    private static void AddActivityAnalysisChart(string workbookPath, string sheetName, string chartName)
    {
        using var document = SpreadsheetDocument.Open(workbookPath, true);
        var workbookPart = document.WorkbookPart;
        var sheet = workbookPart?.Workbook.Sheets?.Elements<S.Sheet>()
            .FirstOrDefault(item => string.Equals(item.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
        if (workbookPart is null || sheet?.Id is null)
        {
            return;
        }

        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
        var drawingsPart = worksheetPart.DrawingsPart;
        if (drawingsPart is null)
        {
            drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
            drawingsPart.WorksheetDrawing = new Xdr.WorksheetDrawing();
            worksheetPart.Worksheet.Append(new S.Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
        }

        drawingsPart.WorksheetDrawing ??= new Xdr.WorksheetDrawing();
        var chartPart = drawingsPart.AddNewPart<ChartPart>();
        chartPart.ChartSpace = BuildActivityAnalysisChartSpace(sheetName);
        chartPart.ChartSpace.Save();

        var chartRelId = drawingsPart.GetIdOfPart(chartPart);
        var drawingId = drawingsPart.WorksheetDrawing
            .Descendants<Xdr.NonVisualDrawingProperties>()
            .Select(item => item.Id?.Value ?? 0u)
            .DefaultIfEmpty(1u)
            .Max() + 1u;

        drawingsPart.WorksheetDrawing.Append(BuildActivityChartAnchor(drawingId, chartRelId, chartName));
        drawingsPart.WorksheetDrawing.Save();
        worksheetPart.Worksheet.Save();
    }

    private static void AddEhcSpendingTrendChart(string workbookPath, IReadOnlyList<ServiceCategoryTrend> rows)
    {
        AddSpendingTrendChart(workbookPath, "ClaimsAnalysis_EHC", rows, "Extended Healthcare Spending Trends");
    }

    private static void AddDentalSpendingTrendChart(string workbookPath, IReadOnlyList<ServiceCategoryTrend> rows)
    {
        AddSpendingTrendChart(workbookPath, "ClaimsAnalysis_Dental", rows, "Dental Spending Trends");
    }

    private static void AddParamedicalServiceTrendChart(string workbookPath, IReadOnlyList<ServiceCategoryTrend> rows)
    {
        AddSpendingTrendChart(workbookPath, "ClaimsAnalysis_Paramedical", rows, "Paramedical Services Utilization Trends");
    }

    private static void AddSpendingTrendChart(
        string workbookPath,
        string sheetName,
        IReadOnlyList<ServiceCategoryTrend> rows,
        string chartName)
    {
        var chartRows = rows
            .Where(row => !row.Category.Equals("Total", StringComparison.OrdinalIgnoreCase))
            .ToList();
        if (chartRows.Count == 0)
        {
            return;
        }

        using var document = SpreadsheetDocument.Open(workbookPath, true);
        var workbookPart = document.WorkbookPart;
        var sheet = workbookPart?.Workbook.Sheets?.Elements<S.Sheet>()
            .FirstOrDefault(item => string.Equals(item.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
        if (workbookPart is null || sheet?.Id is null)
        {
            return;
        }

        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
        var drawingsPart = worksheetPart.DrawingsPart;
        if (drawingsPart is null)
        {
            drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
            drawingsPart.WorksheetDrawing = new Xdr.WorksheetDrawing();
            InsertWorksheetDrawing(worksheetPart, worksheetPart.GetIdOfPart(drawingsPart));
        }

        drawingsPart.WorksheetDrawing ??= new Xdr.WorksheetDrawing();
        var chartPart = drawingsPart.AddNewPart<ChartPart>();
        chartPart.ChartSpace = BuildEhcSpendingTrendChartSpace(sheetName, chartRows);
        chartPart.ChartSpace.Save();

        var chartRelId = drawingsPart.GetIdOfPart(chartPart);
        var drawingId = drawingsPart.WorksheetDrawing
            .Descendants<Xdr.NonVisualDrawingProperties>()
            .Select(item => item.Id?.Value ?? 0u)
            .DefaultIfEmpty(1u)
            .Max() + 1u;

        drawingsPart.WorksheetDrawing.Append(BuildSpendingChartAnchor(drawingId, chartRelId, chartName));
        drawingsPart.WorksheetDrawing.Save();
        worksheetPart.Worksheet.Save();
    }

    private static void InsertWorksheetDrawing(WorksheetPart worksheetPart, string relationshipId)
    {
        if (worksheetPart.Worksheet.Elements<S.Drawing>().Any())
        {
            return;
        }

        var drawing = new S.Drawing { Id = relationshipId };
        OpenXmlElement? insertBefore = worksheetPart.Worksheet.Elements<S.LegacyDrawing>().FirstOrDefault();
        insertBefore ??= worksheetPart.Worksheet.Elements<S.TableParts>().FirstOrDefault();
        insertBefore ??= worksheetPart.Worksheet.Elements<S.ExtensionList>().FirstOrDefault();

        if (insertBefore is null)
        {
            worksheetPart.Worksheet.Append(drawing);
        }
        else
        {
            worksheetPart.Worksheet.InsertBefore(drawing, insertBefore);
        }
    }

    private static C.ChartSpace BuildEhcSpendingTrendChartSpace(string sheetName, IReadOnlyList<ServiceCategoryTrend> rows)
    {
        const uint categoryAxisId = 73650112u;
        const uint valueAxisId = 73672768u;
        var categories = rows.Select(row => row.Category).ToList();
        var labels = GetServiceCategoryPeriodLabels(rows);
        var colors = new[] { "BDE6DC", "64C3AE", "9ABD5A" };
        var barChart = new C.BarChart(
            new C.BarDirection { Val = C.BarDirectionValues.Column },
            new C.BarGrouping { Val = C.BarGroupingValues.Clustered });
        for (var index = 0; index < labels.Count; index++)
        {
            var label = labels[index];
            barChart.Append(BuildEhcSpendingTrendSeries(
                (uint)index,
                sheetName,
                ColumnName(index + 2),
                label,
                categories,
                rows.Select(row => GetServiceCategoryPeriodAmount(row, label)).ToList(),
                colors[index % colors.Length]));
        }
        barChart.Append(new C.AxisId { Val = categoryAxisId });
        barChart.Append(new C.AxisId { Val = valueAxisId });

        return new C.ChartSpace(
            new C.EditingLanguage { Val = "en-US" },
            new C.Chart(
                new C.PlotArea(
                    new C.Layout(),
                    barChart,
                    new C.CategoryAxis(
                        new C.AxisId { Val = categoryAxisId },
                        new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
                        new C.AxisPosition { Val = C.AxisPositionValues.Bottom },
                        new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo },
                        new C.CrossingAxis { Val = valueAxisId },
                        new C.Crosses { Val = C.CrossesValues.AutoZero },
                        new C.AutoLabeled { Val = true },
                        new C.LabelAlignment { Val = C.LabelAlignmentValues.Center },
                        new C.LabelOffset { Val = 100 }),
                    new C.ValueAxis(
                        new C.AxisId { Val = valueAxisId },
                        new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
                        new C.AxisPosition { Val = C.AxisPositionValues.Left },
                        new C.MajorGridlines(),
                        new C.NumberingFormat { FormatCode = "$#,##0", SourceLinked = false },
                        new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo },
                        new C.CrossingAxis { Val = categoryAxisId },
                        new C.Crosses { Val = C.CrossesValues.AutoZero },
                        new C.CrossBetween { Val = C.CrossBetweenValues.Between }),
                    new C.DataTable(
                        new C.ShowHorizontalBorder { Val = true },
                        new C.ShowVerticalBorder { Val = true },
                        new C.ShowOutlineBorder { Val = true },
                        new C.ShowKeys { Val = true })),
                new C.PlotVisibleOnly { Val = true }));
    }

    private static C.BarChartSeries BuildEhcSpendingTrendSeries(
        uint index,
        string sheetName,
        string column,
        string name,
        IReadOnlyList<string> categories,
        IReadOnlyList<decimal> values,
        string color)
    {
        var escapedSheetName = $"'{sheetName.Replace("'", "''", StringComparison.Ordinal)}'";
        return new C.BarChartSeries(
            new C.Index { Val = index },
            new C.Order { Val = index },
            new C.SeriesText(new C.StringReference(
                new C.Formula($"{escapedSheetName}!${column}$1"),
                new C.StringCache(
                    new C.PointCount { Val = 1u },
                    new C.StringPoint { Index = 0u, NumericValue = new C.NumericValue(name) }))),
            new C.ChartShapeProperties(new A.SolidFill(new A.RgbColorModelHex { Val = color })),
            new C.CategoryAxisData(new C.StringReference(
                new C.Formula($"{escapedSheetName}!$A$2:$A${categories.Count + 1}"),
                BuildStringCache(categories))),
            new C.Values(new C.NumberReference(
                new C.Formula($"{escapedSheetName}!${column}$2:${column}${values.Count + 1}"),
                BuildNumberCache(values, "$#,##0"))));
    }

    private static C.StringCache BuildStringCache(IReadOnlyList<string> values)
    {
        var cache = new C.StringCache(new C.PointCount { Val = (uint)values.Count });
        for (var index = 0; index < values.Count; index++)
        {
            cache.Append(new C.StringPoint
            {
                Index = (uint)index,
                NumericValue = new C.NumericValue(values[index])
            });
        }

        return cache;
    }

    private static C.NumberingCache BuildNumberCache(IReadOnlyList<decimal> values, string formatCode)
    {
        var cache = new C.NumberingCache(new C.FormatCode(formatCode), new C.PointCount { Val = (uint)values.Count });
        for (var index = 0; index < values.Count; index++)
        {
            cache.Append(new C.NumericPoint
            {
                Index = (uint)index,
                NumericValue = new C.NumericValue(values[index].ToString("0.##", System.Globalization.CultureInfo.InvariantCulture))
            });
        }

        return cache;
    }

    private static Xdr.TwoCellAnchor BuildSpendingChartAnchor(uint drawingId, string chartRelId, string chartName)
    {
        return new Xdr.TwoCellAnchor(
            new Xdr.FromMarker(
                new Xdr.ColumnId("0"),
                new Xdr.ColumnOffset("0"),
                new Xdr.RowId("8"),
                new Xdr.RowOffset("0")),
            new Xdr.ToMarker(
                new Xdr.ColumnId("8"),
                new Xdr.ColumnOffset("0"),
                new Xdr.RowId("26"),
                new Xdr.RowOffset("0")),
            new Xdr.GraphicFrame(
                new Xdr.NonVisualGraphicFrameProperties(
                    new Xdr.NonVisualDrawingProperties { Id = drawingId, Name = chartName },
                    new Xdr.NonVisualGraphicFrameDrawingProperties()),
                new Xdr.Transform(
                    new A.Offset { X = 0, Y = 0 },
                    new A.Extents { Cx = 0, Cy = 0 }),
                new A.Graphic(new A.GraphicData(new C.ChartReference { Id = chartRelId })
                {
                    Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart"
                })),
            new Xdr.ClientData());
    }

    private static C.ChartSpace BuildActivityAnalysisChartSpace(string sheetName)
    {
        const uint categoryAxisId = 48650112u;
        const uint valueAxisId = 48672768u;

        return new C.ChartSpace(
            new C.EditingLanguage { Val = "en-US" },
            new C.Chart(
                new C.PlotArea(
                    new C.Layout(),
                    new C.BarChart(
                        new C.BarDirection { Val = C.BarDirectionValues.Column },
                        new C.BarGrouping { Val = C.BarGroupingValues.Clustered },
                        BuildActivityChartSeries(0u, sheetName, "Net Billed Premiums", "B", "BDE6DC", "$#,##0"),
                        BuildActivityChartSeries(1u, sheetName, "Paid Claims", "C", "64C3AE", "$#,##0"),
                        BuildActivityChartSeries(2u, sheetName, "Incurred Claims", "D", "C9EAF2", "$#,##0"),
                        new C.AxisId { Val = categoryAxisId },
                        new C.AxisId { Val = valueAxisId }),
                    new C.CategoryAxis(
                        new C.AxisId { Val = categoryAxisId },
                        new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
                        new C.AxisPosition { Val = C.AxisPositionValues.Bottom },
                        new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo },
                        new C.CrossingAxis { Val = valueAxisId },
                        new C.Crosses { Val = C.CrossesValues.AutoZero },
                        new C.AutoLabeled { Val = true },
                        new C.LabelAlignment { Val = C.LabelAlignmentValues.Center },
                        new C.LabelOffset { Val = 100 }),
                    new C.ValueAxis(
                        new C.AxisId { Val = valueAxisId },
                        new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
                        new C.AxisPosition { Val = C.AxisPositionValues.Left },
                        new C.MajorGridlines(),
                        new C.NumberingFormat { FormatCode = "$#,##0", SourceLinked = false },
                        new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo },
                        new C.CrossingAxis { Val = categoryAxisId },
                        new C.Crosses { Val = C.CrossesValues.AutoZero },
                        new C.CrossBetween { Val = C.CrossBetweenValues.Between }),
                    new C.DataTable(
                        new C.ShowHorizontalBorder { Val = true },
                        new C.ShowVerticalBorder { Val = true },
                        new C.ShowOutlineBorder { Val = true },
                        new C.ShowKeys { Val = true })),
                new C.Legend(new C.LegendPosition { Val = C.LegendPositionValues.Bottom }, new C.Layout()),
                new C.PlotVisibleOnly { Val = true }));
    }

    private static C.BarChartSeries BuildActivityChartSeries(uint index, string sheetName, string name, string column, string color, string formatCode)
    {
        var escapedSheetName = $"'{sheetName.Replace("'", "''", StringComparison.Ordinal)}'";
        return new C.BarChartSeries(
            new C.Index { Val = index },
            new C.Order { Val = index },
            new C.SeriesText(new C.StringReference(
                new C.Formula($"{escapedSheetName}!${column}$2"),
                new C.StringCache(
                    new C.PointCount { Val = 1u },
                    new C.StringPoint { Index = 0u, NumericValue = new C.NumericValue(name) }))),
            new C.ChartShapeProperties(new A.SolidFill(new A.RgbColorModelHex { Val = color })),
            new C.CategoryAxisData(new C.StringReference(
                new C.Formula($"{escapedSheetName}!$A$3:$A$5"),
                new C.StringCache(
                    new C.PointCount { Val = 3u },
                    BuildStringPoint(0u, "Prior 2 Periods"),
                    BuildStringPoint(1u, "Last Period"),
                    BuildStringPoint(2u, "Most Current")))),
            new C.Values(new C.NumberReference(
                new C.Formula($"{escapedSheetName}!${column}$3:${column}$5"),
                new C.NumberingCache(
                    new C.FormatCode(formatCode),
                    new C.PointCount { Val = 3u }))));
    }

    private static C.StringPoint BuildStringPoint(uint index, string value)
    {
        return new C.StringPoint { Index = index, NumericValue = new C.NumericValue(value) };
    }

    private static Xdr.TwoCellAnchor BuildActivityChartAnchor(uint drawingId, string chartRelId, string chartName)
    {
        return new Xdr.TwoCellAnchor(
            new Xdr.FromMarker(
                new Xdr.ColumnId("0"),
                new Xdr.ColumnOffset("0"),
                new Xdr.RowId("8"),
                new Xdr.RowOffset("0")),
            new Xdr.ToMarker(
                new Xdr.ColumnId("8"),
                new Xdr.ColumnOffset("0"),
                new Xdr.RowId("24"),
                new Xdr.RowOffset("0")),
            new Xdr.GraphicFrame(
                new Xdr.NonVisualGraphicFrameProperties(
                    new Xdr.NonVisualDrawingProperties { Id = drawingId, Name = chartName },
                    new Xdr.NonVisualGraphicFrameDrawingProperties()),
                new Xdr.Transform(
                    new A.Offset { X = 0, Y = 0 },
                    new A.Extents { Cx = 0, Cy = 0 }),
                new A.Graphic(new A.GraphicData(new C.ChartReference { Id = chartRelId })
                {
                    Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart"
                })),
            new Xdr.ClientData());
    }

    private static void WriteSummary(XLWorkbook workbook, AnalysisResult analysis)
    {
        var ws = workbook.AddWorksheet("Summary");
        Title(ws, "A1", "Renewal Review Summary");

        ws.Cell("A3").Value = "Client";
        ws.Cell("B3").Value = analysis.RenewalCase.ClientName;
        ws.Cell("A4").Value = "Policy";
        ws.Cell("B4").Value = analysis.RenewalCase.PolicyNumber;
        ws.Cell("A5").Value = "Renewal Date";
        ws.Cell("B5").Value = analysis.RenewalCase.RenewalDate?.ToString("dd-MMM-yyyy") ?? "";
        ws.Cell("A6").Value = "Generated";
        ws.Cell("B6").Value = analysis.CreatedAt;
        ws.Cell("B6").Style.DateFormat.Format = "dd-mmm-yyyy h:mm AM/PM";

        ws.Cell("A8").Value = "Current Monthly Premium";
        SetMoney(ws.Cell("B8"), analysis.CurrentMonthlyPremium);
        ws.Cell("A9").Value = "Renewal Monthly Premium";
        SetMoney(ws.Cell("B9"), analysis.RenewalMonthlyPremium);
        ws.Cell("A10").Value = "Monthly Change";
        SetMoney(ws.Cell("B10"), analysis.MonthlyPremiumChange);
        ws.Cell("A11").Value = "Monthly Change %";
        SetPercentNumber(ws.Cell("B11"), analysis.MonthlyPremiumChangePercent);

        var prior = analysis.MostRelevantPreviousRenewal;
        var findingsHeaderRow = 13;
        if (prior is not null)
        {
            ws.Cell("A13").Value = "Previous Renewal Source";
            ws.Cell("B13").Value = prior.SourceFileName;
            ws.Cell("A14").Value = "Previous Renewal Date";
            ws.Cell("B14").Value = prior.RenewalDate?.ToString("dd-MMM-yyyy") ?? "";
            ws.Cell("A15").Value = "Previous Carrier";
            ws.Cell("B15").Value = prior.Carrier;
            ws.Cell("A16").Value = "Previous Current Monthly Premium";
            SetMoney(ws.Cell("B16"), prior.CurrentMonthlyPremium);
            ws.Cell("A17").Value = "Previous Renewal Monthly Premium";
            SetMoney(ws.Cell("B17"), prior.RenewalMonthlyPremium);
            ws.Cell("A18").Value = "Previous Monthly Change";
            SetMoney(ws.Cell("B18"), prior.MonthlyPremiumChange);
            ws.Cell("A19").Value = "Previous Monthly Change %";
            SetPercentNumber(ws.Cell("B19"), prior.MonthlyPremiumChangePercent);
            ws.Range("A13:B19").Style.Fill.BackgroundColor = SoftFill;
            findingsHeaderRow = 22;
        }

        Header(ws.Range(findingsHeaderRow, 1, findingsHeaderRow, 5));
        ws.Cell(findingsHeaderRow, 1).Value = "Severity";
        ws.Cell(findingsHeaderRow, 2).Value = "Finding";
        ws.Cell(findingsHeaderRow, 3).Value = "Evidence";
        ws.Cell(findingsHeaderRow, 4).Value = "Recommended Action";
        ws.Cell(findingsHeaderRow, 5).Value = "Client Explanation";

        var row = findingsHeaderRow + 1;
        foreach (var finding in analysis.Findings.Take(10))
        {
            ws.Cell(row, 1).Value = finding.Severity.ToString();
            ws.Cell(row, 2).Value = finding.Title;
            ws.Cell(row, 3).Value = finding.Evidence;
            ws.Cell(row, 4).Value = finding.RecommendedAction;
            ws.Cell(row, 5).Value = finding.ClientExplanation;
            StyleSeverity(ws.Row(row), finding.Severity);
            row++;
        }

        ws.Range($"A3:A{Math.Max(11, findingsHeaderRow - 3)}").Style.Font.Bold = true;
        ws.Range("A3:B11").Style.Fill.BackgroundColor = SoftFill;
        ws.Column(3).Width = 42;
        ws.Column(4).Width = 44;
        ws.Column(5).Width = 44;
    }

    private static void WritePreviousRenewals(XLWorkbook workbook, AnalysisResult analysis)
    {
        if (analysis.PreviousRenewals.Count == 0)
        {
            return;
        }

        var ws = workbook.AddWorksheet("Previous Renewals");
        WriteRows(
            ws,
            [
                "Source",
                "Type",
                "Client",
                "Policy",
                "Carrier",
                "Renewal Date",
                "Current Monthly Premium",
                "Renewal Monthly Premium",
                "Monthly Change",
                "Monthly Change %",
                "Source Reference"
            ],
            analysis.PreviousRenewals.Select(snapshot => new object?[]
            {
                snapshot.SourceFileName,
                snapshot.SourceType,
                snapshot.ClientName,
                snapshot.PolicyNumber,
                snapshot.Carrier,
                snapshot.RenewalDate?.ToString("dd-MMM-yyyy") ?? "",
                snapshot.CurrentMonthlyPremium,
                snapshot.RenewalMonthlyPremium,
                snapshot.MonthlyPremiumChange,
                snapshot.MonthlyPremiumChangePercent.HasValue ? snapshot.MonthlyPremiumChangePercent.Value / 100m : null,
                snapshot.SourceReference
            }));

        FormatMoneyColumns(ws, 7, 8, 9);
        ws.Column(10).Style.NumberFormat.Format = "0.0%";
    }

    private static void WriteFindings(XLWorkbook workbook, AnalysisResult analysis)
    {
        var ws = workbook.AddWorksheet("Findings");
        WriteRows(
            ws,
            ["Severity", "Title", "Evidence", "Recommended Action", "Client Explanation", "Source"],
            analysis.Findings.Select(f => new object?[]
            {
                f.Severity.ToString(),
                f.Title,
                f.Evidence,
                f.RecommendedAction,
                f.ClientExplanation,
                f.SourceReference
            }));

        var lastFindingRow = ws.LastRowUsed()?.RowNumber() ?? 1;
        for (var row = 2; row <= lastFindingRow; row++)
        {
            if (Enum.TryParse<FindingSeverity>(ws.Cell(row, 1).GetString(), out var severity))
            {
                StyleSeverity(ws.Row(row), severity);
            }
        }
    }

    private static void WriteRates(XLWorkbook workbook, AnalysisResult analysis)
    {
        var ws = workbook.AddWorksheet("Rate Reconciliation");
        var rateLines = SelectedRateLines(analysis);

        var headers = new[]
        {
            "Benefit", "Tier/Class", "Volume", "Lives", "Current Rate", "Current Premium",
            "Current Variance", "Renewal Rate", "Renewal Premium",
            "Renewal Variance", "% Change", "Source",
            "Volume", "Lives", "Negotiated Rate", "Negotiated Premium", "% Change against Current"
        };

        for (var index = 0; index < headers.Length; index++)
        {
            ws.Cell(1, index + 1).Value = headers[index];
        }

        Header(ws.Range(1, 1, 1, headers.Length));
        ws.Range(1, 13, 1, 17).Style.Fill.BackgroundColor = NegotiatedFill;
        ws.Range(1, 13, 1, 17).Style.Font.FontColor = XLColor.White;
        ws.Range(1, 1, 1, headers.Length).SetAutoFilter();
        var rateTotalRows = WriteGroupedRateRows(ws, rateLines);
        WriteRateSourceTotalCheck(ws, rateTotalRows, analysis);
        FormatMoneyColumns(ws, 6, 7, 9, 10, 16);
        FormatNumberColumns(ws, 3, 11, 13);
        ws.Column(14).Style.NumberFormat.Format = "#,##0";
        ws.Column(15).Style.NumberFormat.Format = "0.000";
        ws.Column(17).Style.NumberFormat.Format = "0.0%";
        ws.Columns(2, headers.Length).Width = 10;
        ws.RangeUsed()?.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
    }

    private static void WriteNetAndGrossRates(XLWorkbook workbook, AnalysisResult analysis)
    {
        var version = analysis.SelectedRenewalVersion;
        var rows = version?.RateLines.Where(line => line.HasNetAndGrossRates).ToList() ?? [];
        rows = BenefitDisplayOrder.Sort(
            rows,
            line => line.Benefit,
            line => line.TierOrClass,
            line => line.AccountClass);
        if (rows.Count == 0)
        {
            return;
        }

        var ws = workbook.AddWorksheet("Net and Gross Rates");
        ws.Range(1, 1, 1, 8).Merge();
        ws.Cell(1, 1).Value = "Net and Gross Rates";
        ws.Range(1, 1, 1, 8).Style.Fill.BackgroundColor = WestlandTeal;
        ws.Range(1, 1, 1, 8).Style.Font.Bold = true;
        ws.Range(1, 1, 1, 8).Style.Font.FontColor = XLColor.White;
        ws.Range(1, 1, 1, 8).Style.Font.FontSize = 14;

        ws.Range(2, 1, 2, 8).Merge();
        ws.Cell(2, 1).Value =
            $"Selected version: {version!.VersionName}. Client presentation basis: Gross. PowerPoint exports use Gross rates only.";
        ws.Range(2, 1, 2, 8).Style.Fill.BackgroundColor = SoftFill;
        ws.Range(2, 1, 2, 8).Style.Font.Italic = true;

        ws.Cell(4, 1).Value = "Benefit / Tier";
        MergeAndLabel(ws.Range(4, 2, 4, 3), "Current Rates", LineRatesCurrentPlanHeader);
        MergeAndLabel(ws.Range(4, 4, 4, 5), "Carrier Renewal Rates", LineRatesDarkTeal);
        MergeAndLabel(ws.Range(4, 6, 4, 7), $"{version.VersionName} Rates", NegotiatedFill);
        ws.Cell(4, 8).Value = "Source";
        ws.Range(4, 1, 4, 8).Style.Font.Bold = true;
        ws.Range(4, 1, 4, 8).Style.Font.FontColor = XLColor.White;
        ws.Cell(4, 1).Style.Fill.BackgroundColor = LineRatesDarkTeal;
        ws.Cell(4, 8).Style.Fill.BackgroundColor = LineRatesDarkTeal;

        var headers = new[] { "Benefit / Tier", "Net", "Gross", "Net", "Gross", "Net", "Gross", "Source" };
        for (var column = 0; column < headers.Length; column++)
        {
            ws.Cell(5, column + 1).Value = headers[column];
        }

        Header(ws.Range(5, 1, 5, 8));
        ws.Range(5, 2, 5, 3).Style.Fill.BackgroundColor = LineRatesCurrentPlanHeader;
        ws.Range(5, 4, 5, 5).Style.Fill.BackgroundColor = LineRatesDarkTeal;
        ws.Range(5, 6, 5, 7).Style.Fill.BackgroundColor = NegotiatedFill;

        for (var index = 0; index < rows.Count; index++)
        {
            var line = rows[index];
            var row = index + 6;
            ws.Cell(row, 1).Value = string.IsNullOrWhiteSpace(line.TierOrClass) || line.TierOrClass.Equals("All", StringComparison.OrdinalIgnoreCase)
                ? line.Benefit
                : $"{line.Benefit} / {line.TierOrClass}";
            SetValue(ws.Cell(row, 2), line.CurrentNetRate);
            SetValue(ws.Cell(row, 3), line.CurrentGrossRate);
            SetValue(ws.Cell(row, 4), line.ProposedNetRate);
            SetValue(ws.Cell(row, 5), line.ProposedGrossRate);
            SetValue(ws.Cell(row, 6), line.VersionNetRate);
            SetValue(ws.Cell(row, 7), line.VersionGrossRate);
            ws.Cell(row, 8).Value = string.IsNullOrWhiteSpace(line.RateScheduleSourceFile)
                ? line.SourceReference
                : $"{line.RateScheduleSourceFile} | {line.RateScheduleSourceSheet} | row {line.RateScheduleSourceRow}";
        }

        var lastRow = rows.Count + 5;
        ws.Range(6, 2, lastRow, 7).Style.NumberFormat.Format = "0.000";
        ws.Range(6, 3, lastRow, 3).Style.Fill.BackgroundColor = LineRatesCurrentPlanFill;
        ws.Range(6, 5, lastRow, 5).Style.Fill.BackgroundColor = LineRatesLightTeal;
        ws.Range(6, 7, lastRow, 7).Style.Fill.BackgroundColor = LineRatesMediumTeal;
        ws.SheetView.FreezeRows(5);
        ws.Column(1).Width = 34;
        ws.Columns(2, 7).Width = 14;
        ws.Column(8).Width = 48;
        ws.Range(1, 1, lastRow, 8).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        ws.Range(2, 1, 2, 8).Style.Alignment.WrapText = true;
        ws.Range(6, 8, lastRow, 8).Style.Alignment.WrapText = true;
    }

    private static void WritePowerPointExportTables(XLWorkbook workbook, AnalysisResult analysis)
    {
        const int visibleLineRateColumnCount = 11;
        var ws = workbook.AddWorksheet("PowerPoint Export Tables");
        var row = 1;
        foreach (var table in PowerPointRateTableBuilder.BuildAllTables(analysis))
        {
            var isPremiumIllustrationTable = table.Title.Equals(
                PowerPointRateTableBuilder.PremiumIllustrationTableTitle,
                StringComparison.OrdinalIgnoreCase);
            Title(ws, row, 1, table.Title);
            if (isPremiumIllustrationTable)
            {
                ws.Cell(row, 1).Style.Font.FontName = "Arial";
                ws.Cell(row, 1).Style.Font.FontColor = XLColor.FromHtml("#004751");
            }
            row += 2;

            for (var column = 0; column < table.Headers.Length; column++)
            {
                ws.Cell(row, column + 1).Value = table.Headers[column];
            }

            Header(ws.Range(row, 1, row, table.Headers.Length));
            if (isPremiumIllustrationTable)
            {
                var illustrationHeader = ws.Range(row, 1, row, table.Headers.Length);
                illustrationHeader.Style.Fill.BackgroundColor = XLColor.FromHtml("#004751");
                illustrationHeader.Style.Font.FontName = "Arial";
                illustrationHeader.Style.Font.FontSize = 11;
                illustrationHeader.Style.Font.FontColor = XLColor.FromHtml("#E7F6F9");
                illustrationHeader.Style.Font.Bold = true;
                illustrationHeader.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                illustrationHeader.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                illustrationHeader.Style.Alignment.WrapText = true;
                illustrationHeader.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                illustrationHeader.Style.Border.InsideBorderColor = XLColor.FromHtml("#71C9B6");
                illustrationHeader.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                illustrationHeader.Style.Border.OutsideBorderColor = XLColor.FromHtml("#71C9B6");
                ws.Cell(row, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                ws.Row(row).Height = 60;
            }
            var isLineRateTable = table.Title.Equals("RateTable_LineRates", StringComparison.OrdinalIgnoreCase);
            if (isLineRateTable)
            {
                ws.Range(row, 10, row, Math.Min(11, table.Headers.Length)).Style.Fill.BackgroundColor = NegotiatedFill;
                ws.Range(row, 10, row, Math.Min(11, table.Headers.Length)).Style.Font.FontColor = XLColor.White;
            }

            row++;

            var lineRateDataStart = row;
            var lineRateTotalRow = 0;
            foreach (var tableRow in table.Rows)
            {
                for (var column = 0; column < table.Headers.Length; column++)
                {
                    var value = column < tableRow.Length ? tableRow[column] : "";
                    if (!isPremiumIllustrationTable || !string.IsNullOrEmpty(value))
                    {
                        ws.Cell(row, column + 1).Value = value;
                    }
                }

                ws.Range(row, 1, row, table.Headers.Length).Style.Alignment.WrapText = true;
                var isSummaryRow = tableRow.Any(cell =>
                    cell.Contains("Total", StringComparison.OrdinalIgnoreCase) ||
                    cell.Contains("Monthly Difference", StringComparison.OrdinalIgnoreCase));
                if (isPremiumIllustrationTable)
                {
                    var illustrationRow = ws.Range(row, 1, row, table.Headers.Length);
                    illustrationRow.Style.Fill.BackgroundColor = isSummaryRow
                        ? XLColor.FromHtml("#E7F6F9")
                        : XLColor.FromHtml("#CDECEF");
                    illustrationRow.Style.Font.FontName = "Arial";
                    illustrationRow.Style.Font.FontSize = 11;
                    illustrationRow.Style.Font.FontColor = XLColor.FromHtml("#004751");
                    illustrationRow.Style.Font.Bold = isSummaryRow;
                    illustrationRow.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    illustrationRow.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                    illustrationRow.Style.Border.InsideBorderColor = XLColor.FromHtml("#71C9B6");
                    illustrationRow.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    illustrationRow.Style.Border.OutsideBorderColor = XLColor.FromHtml("#71C9B6");
                    ws.Cell(row, 1).Style.Font.Bold = true;
                    ws.Cell(row, 2).Style.Font.Bold = true;
                    ws.Cell(row, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                    ws.Cell(row, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    ws.Range(row, 3, row, table.Headers.Length).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    var lineCount = tableRow.Take(table.Headers.Length)
                        .Select(value => value.Count(character => character == '\n') + 1)
                        .DefaultIfEmpty(1)
                        .Max();
                    ws.Row(row).Height = Math.Max(isSummaryRow ? 36 : 18, lineCount * 15);
                }
                if (isLineRateTable)
                {
                    if (tableRow.Length > visibleLineRateColumnCount)
                    {
                        // Keep calculation metadata out of the visible export table while making it
                        // available to the formula-backed formatted worksheet.
                        ws.Cell(row, 12).Value = tableRow[visibleLineRateColumnCount];
                    }

                    if (IsLineRateTotalRow(tableRow))
                    {
                        WriteLineRateTotalFormulas(ws, row, lineRateDataStart, row - 1);
                        lineRateTotalRow = row;
                    }
                    else if (IsLineRateDifferenceRow(tableRow))
                    {
                        WriteLineRateDifferenceFormulas(ws, row, lineRateTotalRow);
                    }
                    else
                    {
                        WriteLineRatePremiumFormulas(
                            ws,
                            row,
                            tableRow.Length > 1 ? tableRow[1] : "",
                            tableRow.Length > visibleLineRateColumnCount
                                ? tableRow[visibleLineRateColumnCount]
                                : "");
                    }
                }

                if (isSummaryRow)
                {
                    ws.Range(row, 1, row, table.Headers.Length).Style.Font.Bold = true;
                    ws.Range(row, 1, row, table.Headers.Length).Style.Fill.BackgroundColor = isPremiumIllustrationTable
                        ? XLColor.FromHtml("#E7F6F9")
                        : SoftFill;
                    if (isLineRateTable)
                    {
                        ws.Range(row, 10, row, Math.Min(11, table.Headers.Length)).Style.Fill.BackgroundColor = NegotiatedFill;
                    }
                }

                row++;
            }

            row += 2;
        }

        ws.SheetView.FreezeRows(1);
        ws.Column(12).Hide();
        ws.Columns().AdjustToContents(10, 42);
    }

    private static void ApplyPremiumIllustrationExportFinishing(XLWorkbook workbook)
    {
        if (!workbook.TryGetWorksheet("PowerPoint Export Tables", out var ws))
        {
            return;
        }

        var title = ws.CellsUsed().FirstOrDefault(cell =>
            !cell.HasFormula &&
            cell.GetString().Equals(PowerPointRateTableBuilder.PremiumIllustrationTableTitle, StringComparison.Ordinal));
        if (title is null)
        {
            return;
        }

        var headerRow = title.Address.RowNumber + 2;
        var lastRow = ws.RowsUsed().First(row =>
            row.RowNumber() > headerRow &&
            row.Cell(1).GetString().Equals("Monthly % Difference from Current", StringComparison.OrdinalIgnoreCase)).RowNumber();
        var tableRange = ws.Range(headerRow, 1, lastRow, 7);
        tableRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        tableRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
        tableRange.Style.Border.InsideBorderColor = XLColor.FromHtml("#71C9B6");
        tableRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        tableRange.Style.Border.OutsideBorderColor = XLColor.FromHtml("#71C9B6");

        ws.Column(1).Width = Math.Max(ws.Column(1).Width, 38);
        ws.Column(2).Width = Math.Max(ws.Column(2).Width, 14);
        ws.Column(3).Width = Math.Max(ws.Column(3).Width, 28);
        ws.Column(4).Width = Math.Max(ws.Column(4).Width, 20);
        ws.Column(5).Width = Math.Max(ws.Column(5).Width, 24);
        ws.Column(6).Width = Math.Max(ws.Column(6).Width, 26);
        ws.Column(7).Width = Math.Max(ws.Column(7).Width, 22);
    }

    private static void WriteNems(XLWorkbook workbook, AnalysisResult analysis)
    {
        var ws = workbook.AddWorksheet("NEMs");
        Title(ws, 1, 1, "Non-Evidence Maximums");
        var headers = new[] { "Benefit", "Current", "Available", "Source" };
        for (var column = 0; column < headers.Length; column++)
        {
            ws.Cell(3, column + 1).Value = headers[column];
        }

        Header(ws.Range(3, 1, 3, headers.Length));
        var displayRows = PowerPointRateTableBuilder.BuildNemMaximumsTable(analysis).Rows;
        for (var index = 0; index < displayRows.Count; index++)
        {
            var displayRow = displayRows[index];
            var sourceRow = FindNemSourceRow(analysis, displayRow[0]);
            var row = index + 4;
            ws.Cell(row, 1).Value = displayRow[0];
            ws.Cell(row, 2).Value = displayRow[1];
            ws.Cell(row, 3).Value = displayRow[2];
            ws.Cell(row, 4).Value = sourceRow is null ? "" : FirstNonBlank(sourceRow.SourceReference, sourceRow.SourceFileName);
            ws.Range(row, 1, row, headers.Length).Style.Alignment.WrapText = true;
        }

        ws.SheetView.FreezeRows(3);
        ws.Columns(2, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
        ws.RangeUsed()?.Style.Border.InsideBorderColor = XLColor.FromHtml("#D9E2E8");
        ws.RangeUsed()?.Style.Border.OutsideBorderColor = XLColor.FromHtml("#D9E2E8");
    }

    private static NemMaximumRow? FindNemSourceRow(AnalysisResult analysis, string benefit)
    {
        var normalized = NemMaximumExtractor.NormalizeBenefitName(benefit);
        return analysis.NemMaximums.FirstOrDefault(row =>
            NemMaximumExtractor.NormalizeBenefitName(row.Benefit).Equals(normalized, StringComparison.OrdinalIgnoreCase));
    }

    private static void WriteRateTableLineRatesFormatted(XLWorkbook workbook, AnalysisResult analysis)
    {
        if (workbook.TryGetWorksheet("RateTable_LineRates_Formatted", out var existing))
        {
            existing.Delete();
        }

        var ws = workbook.AddWorksheet("RateTable_LineRates_Formatted");
        WriteFormattedLineRatesHeaderRows(ws);

        var sourceRows = ReadLineRateSourceRows(workbook);
        if (sourceRows.Count > 0)
        {
            WriteFormattedLineRateFormulaRows(ws, sourceRows);
            return;
        }

        var table = PowerPointRateTableBuilder.BuildAllTables(analysis)
            .FirstOrDefault(item => item.Title.Equals("RateTable_LineRates_Formatted", StringComparison.OrdinalIgnoreCase));
        if (table is null || table.Rows.Count == 0)
        {
            ws.Cell(6, 1).Value = "RateTable_LineRates source rows were not found.";
            return;
        }

        var row = 6;
        foreach (var tableRow in table.Rows)
        {
            if (IsFormattedLineRateDirectColumnHeader(tableRow))
            {
                continue;
            }

            for (var column = 0; column < 10; column++)
            {
                var value = column < tableRow.Length ? tableRow[column] : "";
                WriteFormattedLineRateDirectValue(ws.Cell(row, column + 1), value);
            }

            if (IsFormattedLineRateDirectSectionRow(tableRow))
            {
                StyleFormattedLineRateSectionRow(ws, row);
            }
            else if (IsFormattedLineRateDirectSummaryRow(tableRow))
            {
                StyleFormattedLineRateSummaryRow(ws, row);
            }

            row++;
        }
    }

    private static void WriteFormattedLineRateFormulaRows(IXLWorksheet ws, List<FormattedLineRateSourceRow> sourceRows)
    {
        var groups = BuildFormattedLineRateGroups(sourceRows);
        var allDetailRows = new List<int>();
        var pooledRows = new List<int>();
        var ehcRows = new List<int>();
        var dentalRows = new List<int>();
        var pooledTotalWritten = false;
        var row = 6;

        for (var index = 0; index < groups.Count; index++)
        {
            var group = groups[index];
            if (!pooledTotalWritten && group.SortOrder >= 90 && pooledRows.Count > 0)
            {
                var pooledTotalRow = row++;
                WriteFormattedLineRateTotalRow(ws, pooledTotalRow, "Total Pooled Benefits", pooledRows);
                WriteFormattedLineRateDifferenceRow(ws, row++, "Pooled Difference", pooledTotalRow);
                pooledTotalWritten = true;
            }

            var detailRows = WriteFormattedLineRateGroup(ws, group, ref row);
            allDetailRows.AddRange(detailRows);

            if (group.SortOrder < 90)
            {
                pooledRows.AddRange(detailRows);
            }

            if (group.Kind == FormattedLineRateGroupKind.ExtendedHealth)
            {
                ehcRows.AddRange(detailRows);
            }
            else if (group.Kind == FormattedLineRateGroupKind.Dental)
            {
                dentalRows.AddRange(detailRows);
            }

            var nextKind = index + 1 < groups.Count ? groups[index + 1].Kind : (FormattedLineRateGroupKind?)null;
            if (group.Kind == FormattedLineRateGroupKind.ExtendedHealth &&
                nextKind != FormattedLineRateGroupKind.ExtendedHealth)
            {
                var ehcTotalRow = row++;
                WriteFormattedLineRateTotalRow(ws, ehcTotalRow, "Total Health Premiums", ehcRows);
                WriteFormattedLineRateAnnualRow(ws, row++, "Total Annual EHC premiums", ehcTotalRow);
            }

            if (group.Kind == FormattedLineRateGroupKind.Dental &&
                nextKind != FormattedLineRateGroupKind.Dental)
            {
                var dentalTotalRow = row++;
                WriteFormattedLineRateTotalRow(ws, dentalTotalRow, "Total Dental Premiums", dentalRows);
                WriteFormattedLineRateAnnualRow(ws, row++, "Total Annual Dental Premiums", dentalTotalRow);
                if (ehcRows.Count > 0)
                {
                    WriteFormattedLineRateTotalRow(ws, row++, "Total EHC & Dental", ehcRows.Concat(dentalRows));
                }
            }
        }

        if (!pooledTotalWritten && pooledRows.Count > 0)
        {
            var pooledTotalRow = row++;
            WriteFormattedLineRateTotalRow(ws, pooledTotalRow, "Total Pooled Benefits", pooledRows);
            WriteFormattedLineRateDifferenceRow(ws, row++, "Pooled Difference", pooledTotalRow);
        }

        var monthlyTotalRow = row++;
        WriteFormattedLineRateTotalRow(ws, monthlyTotalRow, "Total monthly premium", allDetailRows);
        WriteFormattedLineRateDifferenceRow(ws, row++, "Monthly premium difference", monthlyTotalRow);
        WriteFormattedLineRateAnnualRow(ws, row++, "Total annual premium", monthlyTotalRow);
        WriteFormattedLineRateAnnualDifferenceRow(ws, row++, "Annual premium difference", monthlyTotalRow);
    }

    private static bool IsFormattedLineRateDirectColumnHeader(string[] row)
    {
        return row.Length > 0 && row[0].Equals("Benefit / Class", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsFormattedLineRateDirectSectionRow(string[] row)
    {
        return row.Length > 0 &&
               !string.IsNullOrWhiteSpace(row[0]) &&
               row.Skip(1).All(string.IsNullOrWhiteSpace) &&
               !IsFormattedLineRateDirectSummaryRow(row);
    }

    private static bool IsFormattedLineRateDirectSummaryRow(string[] row)
    {
        if (row.Length == 0)
        {
            return false;
        }

        var label = row[0];
        return label.Contains("Total", StringComparison.OrdinalIgnoreCase) ||
               label.Contains("Difference", StringComparison.OrdinalIgnoreCase);
    }

    private static void StyleFormattedLineRateSectionRow(IXLWorksheet ws, int row)
    {
        var range = ws.Range(row, 1, row, 10);
        range.Style.Font.Bold = true;
        range.Style.Font.FontColor = LineRatesDarkTeal;
        range.Style.Border.TopBorder = XLBorderStyleValues.Thin;
        range.Style.Border.TopBorderColor = LineRatesGridTeal;
    }

    private static void WriteFormattedLineRateDirectValue(IXLCell cell, string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            cell.Value = "";
            return;
        }

        var trimmed = value.Trim();
        var numericText = trimmed.Replace("$", "", StringComparison.Ordinal).Replace(",", "", StringComparison.Ordinal);
        if (decimal.TryParse(numericText, out var numeric))
        {
            cell.Value = numeric;
            return;
        }

        cell.Value = value;
    }

    private static void WriteFormattedLineRatesHeaderRows(IXLWorksheet ws)
    {
        var title = ws.Range(1, 1, 1, 10);
        title.Merge();
        title.FirstCell().Value = "RateTable_LineRates_Formatted";
        title.Style.Fill.BackgroundColor = XLColor.FromHtml("#3399B3");
        title.Style.Font.Bold = true;
        title.Style.Font.FontSize = 14;
        title.Style.Font.FontColor = XLColor.White;
        title.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

        MergeAndLabel(ws.Range(3, 3, 3, 4), "Current Plan", LineRatesCurrentPlanHeader);
        MergeAndLabel(ws.Range(3, 5, 3, 6), "Calculated", LineRatesDarkTeal);
        MergeAndLabel(ws.Range(3, 7, 3, 8), "Proposed", LineRatesDarkTeal);
        MergeAndLabel(ws.Range(3, 9, 3, 10), "Negotiated", LineRatesDarkTeal);

        var headers = new[]
        {
            "Benefit / Class",
            "Volume",
            "Rate",
            "Premium",
            "Rate",
            "Premium",
            "Rate",
            "Premium",
            "Rate",
            "Premium"
        };

        for (var index = 0; index < headers.Length; index++)
        {
            ws.Cell(4, index + 1).Value = headers[index];
        }

        ws.Range(4, 1, 4, 10).Style.Font.Bold = true;
        ws.Range(4, 1, 4, 10).Style.Font.FontColor = LineRatesDarkTeal;
        ws.Range(4, 1, 4, 10).Style.Fill.BackgroundColor = LineRatesLightTeal;
    }

    private static void MergeAndLabel(IXLRange range, string label, XLColor fill)
    {
        range.Merge();
        range.FirstCell().Value = label;
        range.Style.Fill.BackgroundColor = fill;
        range.Style.Font.Bold = true;
        range.Style.Font.FontColor = XLColor.White;
        range.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
    }

    private static List<FormattedLineRateSourceRow> ReadLineRateSourceRows(XLWorkbook workbook)
    {
        if (!workbook.TryGetWorksheet("PowerPoint Export Tables", out var sourceSheet))
        {
            return [];
        }

        var used = sourceSheet.RangeUsed();
        if (used is null)
        {
            return [];
        }

        var titleCell = used.CellsUsed()
            .FirstOrDefault(cell => CellTextWithoutFormulaEvaluation(cell).Equals("RateTable_LineRates", StringComparison.OrdinalIgnoreCase));
        if (titleCell is null)
        {
            return [];
        }

        var headerRow = titleCell.Address.RowNumber + 2;
        var lastRow = sourceSheet.LastRowUsed()?.RowNumber() ?? headerRow;
        var rows = new List<FormattedLineRateSourceRow>();
        for (var row = headerRow + 1; row <= lastRow; row++)
        {
            var rowText = string.Concat(Enumerable.Range(1, 11).Select(column => CellTextWithoutFormulaEvaluation(sourceSheet.Cell(row, column))));
            if (string.IsNullOrWhiteSpace(rowText))
            {
                break;
            }

            var firstCell = CellTextWithoutFormulaEvaluation(sourceSheet.Cell(row, 1));
            var benefit = CellTextWithoutFormulaEvaluation(sourceSheet.Cell(row, 2));
            if (IsLineRateTotalRow([firstCell]) || IsLineRateDifferenceRow([firstCell]) || string.IsNullOrWhiteSpace(benefit))
            {
                continue;
            }

            rows.Add(new FormattedLineRateSourceRow(
                row,
                firstCell,
                benefit,
                CellTextWithoutFormulaEvaluation(sourceSheet.Cell(row, 3)),
                CellTextWithoutFormulaEvaluation(sourceSheet.Cell(row, 12))));
        }

        return rows;
    }

    private static string CellTextWithoutFormulaEvaluation(IXLCell cell)
    {
        if (!string.IsNullOrWhiteSpace(cell.FormulaA1))
        {
            return cell.FormulaA1;
        }

        return cell.GetString();
    }

    private static List<FormattedLineRateGroup> BuildFormattedLineRateGroups(List<FormattedLineRateSourceRow> sourceRows)
    {
        return sourceRows
            .Select((row, index) => new
            {
                Row = row,
                Index = index,
                Key = GetFormattedLineRateGroupKey(row.Benefit, row.RateBasis)
            })
            .GroupBy(item => item.Key.Key, StringComparer.OrdinalIgnoreCase)
            .Select(group =>
            {
                var first = group.OrderBy(item => item.Index).First();
                return new FormattedLineRateGroup(
                    first.Key.Key,
                    first.Key.DisplayName,
                    first.Key.SortOrder,
                    first.Key.Kind,
                    first.Index,
                    group.OrderBy(item => item.Index).Select(item => item.Row).ToList());
            })
            .OrderBy(group => group.SortOrder)
            .ThenBy(group => group.FirstIndex)
            .ThenBy(group => group.DisplayName, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static List<int> WriteFormattedLineRateGroup(IXLWorksheet ws, FormattedLineRateGroup group, ref int row)
    {
        var titleRow = row++;
        ws.Cell(titleRow, 1).Value = group.DisplayName;
        ws.Range(titleRow, 1, titleRow, 10).Style.Font.Bold = true;
        ws.Range(titleRow, 1, titleRow, 10).Style.Font.FontColor = LineRatesDarkTeal;
        ws.Range(titleRow, 1, titleRow, 10).Style.Border.TopBorder = XLBorderStyleValues.Thin;
        ws.Range(titleRow, 1, titleRow, 10).Style.Border.TopBorderColor = LineRatesGridTeal;

        var detailRows = new List<int>();
        foreach (var source in group.Rows)
        {
            var detailRow = row++;
            detailRows.Add(detailRow);
            ws.Cell(detailRow, 1).Value = FormatFormattedLineRateDetail(source);
            ws.Cell(detailRow, 2).FormulaA1 = SourceVolumeOrLivesFormula(source.RowNumber);
            ws.Cell(detailRow, 3).FormulaA1 = SourceValueFormula(source.RowNumber, "F");
            ws.Cell(detailRow, 4).FormulaA1 = FormattedLineRatePremiumFormula(detailRow, "C", source.Benefit, source.RateBasis);
            ws.Cell(detailRow, 5).FormulaA1 = FormattedLineRateCalculatedRateFormula(group.Kind, source, detailRow) ?? SourceValueFormula(source.RowNumber, "H");
            if (group.Kind is FormattedLineRateGroupKind.ExtendedHealth or FormattedLineRateGroupKind.Dental)
            {
                ws.Cell(detailRow, 5).Style.NumberFormat.Format = "$#,##0.00";
            }

            ws.Cell(detailRow, 6).FormulaA1 = FormattedLineRatePremiumFormula(detailRow, "E", source.Benefit, source.RateBasis);
            ws.Cell(detailRow, 7).FormulaA1 = SourceValueFormula(source.RowNumber, "H");
            ws.Cell(detailRow, 8).FormulaA1 = FormattedLineRatePremiumFormula(detailRow, "G", source.Benefit, source.RateBasis);
            ws.Cell(detailRow, 9).FormulaA1 = SourceValueFormula(source.RowNumber, "J");
            ws.Cell(detailRow, 10).FormulaA1 = FormattedLineRatePremiumFormula(detailRow, "I", source.Benefit, source.RateBasis);
        }

        return detailRows;
    }

    private static string? FormattedLineRateCalculatedRateFormula(
        FormattedLineRateGroupKind kind,
        FormattedLineRateSourceRow source,
        int detailRow)
    {
        if (IsEmployeeLifeCalculatedRate(source.Benefit))
        {
            return $"IFERROR(IF('Mortality Master'!$L$17=\"\",\"\",ROUND(C{detailRow}*(1+'Mortality Master'!$L$17),3)),\"\")";
        }

        if (IsLongTermDisabilityCalculatedRate(source.Benefit))
        {
            return $"IFERROR(IF('Morbidity Master'!$L$16=\"\",\"\",ROUND(C{detailRow}*(1+'Morbidity Master'!$L$16),3)),\"\")";
        }

        var sheetName = kind switch
        {
            FormattedLineRateGroupKind.ExtendedHealth => "HealthRenewal_FormulaAnalysis",
            FormattedLineRateGroupKind.Dental => "DentalRenewal_FormulaAnalysis",
            _ => ""
        };

        if (string.IsNullOrWhiteSpace(sheetName))
        {
            return null;
        }

        var row = HealthDentalCalculatedRateRow(source);
        return row.HasValue ? $"'{sheetName}'!F{row.Value}" : null;
    }

    private static bool IsEmployeeLifeCalculatedRate(string benefit)
    {
        return UnderwritingRiskCalculator.RateTypeForBenefit(benefit) == UnderwritingRateTypes.Mortality;
    }

    private static bool IsLongTermDisabilityCalculatedRate(string benefit)
    {
        return UnderwritingRiskCalculator.RateTypeForBenefit(benefit) == UnderwritingRateTypes.Morbidity;
    }

    private static int? HealthDentalCalculatedRateRow(FormattedLineRateSourceRow source)
    {
        if (NormalizeLineRateText(source.AccountClass).Contains("RWAM CLASS", StringComparison.Ordinal))
        {
            // RWAM divisions (for example 1A or 4C) are not the formula workbook's A/B/C classes.
            // Leaving the calculated rate blank preserves the source class instead of silently using Class A.
            return null;
        }

        var classIndex = FormulaClassIndex(source.AccountClass);
        var tierIndex = FormulaTierIndex(source.TierClass);
        if (!tierIndex.HasValue)
        {
            return null;
        }

        return 37 + classIndex * 17 + tierIndex.Value * 5;
    }

    private static int FormulaClassIndex(string accountClass)
    {
        var normalized = NormalizeLineRateText(accountClass);
        if (normalized.Contains("CLASS C", StringComparison.Ordinal) ||
            normalized.Equals("C", StringComparison.Ordinal))
        {
            return 2;
        }

        if (normalized.Contains("CLASS B", StringComparison.Ordinal) ||
            normalized.Equals("B", StringComparison.Ordinal))
        {
            return 1;
        }

        return 0;
    }

    private static int? FormulaTierIndex(string tierClass)
    {
        var normalized = NormalizeLineRateText(tierClass);
        if (normalized.Contains("SINGLE PARENT", StringComparison.Ordinal))
        {
            // The three-tier health formula sheet has no Single Parent row. Preserve the
            // carrier source rate rather than silently substituting the Single rate.
            return null;
        }

        if (normalized.Contains("SINGLE", StringComparison.Ordinal))
        {
            return 0;
        }

        if (normalized.Contains("COUPLE", StringComparison.Ordinal))
        {
            return 1;
        }

        if (normalized.Contains("FAMILY", StringComparison.Ordinal))
        {
            return 2;
        }

        return null;
    }

    private static string FormattedLineRatePremiumFormula(int row, string rateColumn, string benefit, string rateBasis)
    {
        if (UnderwritingEngine.IsPerEmployeeRateBasis(rateBasis))
        {
            return $"IF(OR(B{row}=\"\",{rateColumn}{row}=\"\"),\"\",IFERROR(ROUND(B{row}*{rateColumn}{row},2),\"\"))";
        }

        var divisor = RateVolumeDivisor(benefit, rateBasis);
        return divisor.HasValue
            ? $"IF(OR(B{row}=\"\",{rateColumn}{row}=\"\"),\"\",IFERROR(ROUND(B{row}*{rateColumn}{row}/{divisor.Value},2),\"\"))"
            : $"IF(OR(B{row}=\"\",{rateColumn}{row}=\"\"),\"\",IFERROR(ROUND(B{row}*{rateColumn}{row},2),\"\"))";
    }

    private static void WriteFormattedLineRateTotalRow(IXLWorksheet ws, int row, string label, IEnumerable<int> sourceRows)
    {
        ws.Cell(row, 1).Value = label;
        foreach (var column in new[] { "D", "F", "H", "J" })
        {
            ws.Cell(row, ColumnNumber(column)).FormulaA1 = SumFormula(sourceRows, column);
        }

        StyleFormattedLineRateSummaryRow(ws, row);
    }

    private static void WriteFormattedLineRateDifferenceRow(IXLWorksheet ws, int row, string label, int totalRow)
    {
        ws.Cell(row, 1).Value = label;
        foreach (var column in new[] { "F", "H", "J" })
        {
            ws.Cell(row, ColumnNumber(column)).FormulaA1 = $"IFERROR({column}{totalRow}-D{totalRow},\"\")";
        }

        StyleFormattedLineRateSummaryRow(ws, row);
    }

    private static void WriteFormattedLineRateAnnualRow(IXLWorksheet ws, int row, string label, int monthlyRow)
    {
        ws.Cell(row, 1).Value = label;
        foreach (var column in new[] { "D", "F", "H", "J" })
        {
            ws.Cell(row, ColumnNumber(column)).FormulaA1 = $"IFERROR({column}{monthlyRow}*12,\"\")";
        }

        StyleFormattedLineRateSummaryRow(ws, row);
    }

    private static void WriteFormattedLineRateAnnualDifferenceRow(IXLWorksheet ws, int row, string label, int monthlyTotalRow)
    {
        ws.Cell(row, 1).Value = label;
        foreach (var column in new[] { "F", "H", "J" })
        {
            ws.Cell(row, ColumnNumber(column)).FormulaA1 = $"IFERROR(({column}{monthlyTotalRow}-D{monthlyTotalRow})*12,\"\")";
        }

        StyleFormattedLineRateSummaryRow(ws, row);
    }

    private static void StyleFormattedLineRateSummaryRow(IXLWorksheet ws, int row)
    {
        var range = ws.Range(row, 1, row, 10);
        range.Style.Font.Bold = true;
        range.Style.Fill.BackgroundColor = LineRatesLightTeal;
        range.Style.Border.TopBorder = XLBorderStyleValues.Thin;
        range.Style.Border.TopBorderColor = LineRatesDarkTeal;
    }

    private static void FormatRateTableLineRatesFormattedSheet(IXLWorksheet ws)
    {
        var used = ws.RangeUsed();
        if (used is null)
        {
            return;
        }

        ws.SheetView.FreezeRows(4);
        ws.ShowGridLines = false;
        used.Style.Font.FontName = "Arial";
        used.Style.Font.FontSize = 10;
        used.Style.Font.FontColor = LineRatesDarkTeal;
        used.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        used.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
        used.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        used.Style.Border.InsideBorderColor = LineRatesGridTeal;
        used.Style.Border.OutsideBorderColor = LineRatesGridTeal;

        ws.Columns(1, 10).Style.Alignment.WrapText = true;
        ws.Columns(2, 10).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
        ws.Column(1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
        ws.Column(1).Width = 38;
        ws.Column(2).Width = 13;
        for (var column = 3; column <= 10; column++)
        {
            ws.Column(column).Width = 14;
        }

        ws.Columns(3, 3).Style.NumberFormat.Format = "$#,##0.00#";
        ws.Columns(5, 5).Style.NumberFormat.Format = "$#,##0.00#";
        ws.Columns(7, 7).Style.NumberFormat.Format = "$#,##0.00#";
        ws.Columns(9, 9).Style.NumberFormat.Format = "$#,##0.00#";
        ws.Columns(4, 4).Style.NumberFormat.Format = "$#,##0.00";
        ws.Columns(6, 6).Style.NumberFormat.Format = "$#,##0.00";
        ws.Columns(8, 8).Style.NumberFormat.Format = "$#,##0.00";
        ws.Columns(10, 10).Style.NumberFormat.Format = "$#,##0.00";
        ws.Column(2).Style.NumberFormat.Format = "#,##0";

        ws.RangeUsed()?.Style.Border.InsideBorderColor = LineRatesGridTeal;
        ws.RangeUsed()?.Style.Border.OutsideBorderColor = LineRatesGridTeal;
        ws.Range(3, 3, 3, 4).Style.Fill.BackgroundColor = LineRatesCurrentPlanHeader;
        ws.Range(3, 5, 3, 10).Style.Fill.BackgroundColor = LineRatesDarkTeal;
        ws.Range(4, 1, 4, 10).Style.Fill.BackgroundColor = LineRatesLightTeal;
        ws.Range(1, 1, 1, 10).Style.Fill.BackgroundColor = XLColor.FromHtml("#3399B3");
        ws.Range(1, 1, 1, 10).Style.Font.FontColor = XLColor.White;
        ws.Range(1, 1, 1, 10).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        ws.Range(3, 3, 3, 10).Style.Font.FontColor = XLColor.White;
        ws.Range(3, 3, 3, 10).Style.Font.Bold = true;

        var lastRow = used.LastRow().RowNumber();
        if (lastRow >= 5)
        {
            ws.Range(5, 3, lastRow, 4).Style.Fill.BackgroundColor = LineRatesCurrentPlanFill;
            ws.Range(5, 5, lastRow, 6).Style.Fill.BackgroundColor = LineRatesMediumTeal;
            ws.Range(5, 7, lastRow, 8).Style.Fill.BackgroundColor = LineRatesLightTeal;
            ws.Range(5, 9, lastRow, 10).Style.Fill.BackgroundColor = LineRatesMediumTeal;
        }

        ApplyHealthDentalCalculatedRateFormat(ws, lastRow);

        for (var row = 1; row <= lastRow; row++)
        {
            ws.Row(row).Height = 18;
        }

        ws.Row(1).Height = 24;
        ws.Row(3).Height = 22;
        ws.Row(4).Height = 20;
    }

    private static void ApplyHealthDentalCalculatedRateFormat(IXLWorksheet ws, int lastRow)
    {
        for (var row = 1; row <= lastRow; row++)
        {
            var calculatedRateCell = ws.Cell(row, 5);
            var formula = calculatedRateCell.FormulaA1;
            if (formula.Contains("HealthRenewal_FormulaAnalysis", StringComparison.OrdinalIgnoreCase) ||
                formula.Contains("DentalRenewal_FormulaAnalysis", StringComparison.OrdinalIgnoreCase))
            {
                calculatedRateCell.Style.NumberFormat.Format = "$#,##0.00";
            }
            else if (formula.Contains("Mortality Master", StringComparison.OrdinalIgnoreCase))
            {
                calculatedRateCell.Style.NumberFormat.Format = "0.000";
            }
        }
    }

    private static string SourceVolumeOrLivesFormula(int row)
    {
        var volume = SourceCell(row, "D");
        var lives = SourceCell(row, "E");
        return $"IFERROR(IF({volume}<>\"\",{volume},IF({lives}<>\"\",{lives},\"\")),\"\")";
    }

    private static string SourceValueFormula(int row, string column)
    {
        var source = SourceCell(row, column);
        return $"IFERROR(IF({source}=\"\",\"\",{source}),\"\")";
    }

    private static string SourceCell(int row, string column)
    {
        return $"'PowerPoint Export Tables'!{column}{row}";
    }

    private static string SumFormula(IEnumerable<int> sourceRows, string column)
    {
        var references = sourceRows.Select(row => $"{column}{row}").ToList();
        return references.Count == 0 ? "0" : $"SUM({string.Join(",", references)})";
    }

    private static int ColumnNumber(string column)
    {
        return column.Aggregate(0, (current, character) => current * 26 + character - 'A' + 1);
    }

    private static string FormatFormattedLineRateDetail(FormattedLineRateSourceRow source)
    {
        var tier = FormatFormattedTierClass(source.TierClass);
        if (!string.IsNullOrWhiteSpace(tier) && !tier.Equals("All", StringComparison.OrdinalIgnoreCase))
        {
            return $"  {tier}";
        }

        var account = CleanFormattedAccountLabel(source.AccountClass);
        if (!string.IsNullOrWhiteSpace(account) && !account.Equals("ALL", StringComparison.OrdinalIgnoreCase))
        {
            return account;
        }

        return "All";
    }

    private static string CleanFormattedAccountLabel(string accountClass)
    {
        return accountClass
            .Replace("ACCOUNT/CLASS:", "", StringComparison.OrdinalIgnoreCase)
            .Trim();
    }

    private static string FormatFormattedTierClass(string tierClass)
    {
        return string.IsNullOrWhiteSpace(tierClass) ? "" : tierClass.Trim();
    }

    private static FormattedLineRateGroupKey GetFormattedLineRateGroupKey(string benefit, string rateBasis)
    {
        var normalized = NormalizeLineRateText(benefit);

        if (normalized.Contains("OPTIONAL", StringComparison.Ordinal) && IsFormattedAdAndD(normalized))
        {
            return new("OPTIONAL AD&D", "Optional AD&D", 60, FormattedLineRateGroupKind.Pooled);
        }

        if (normalized.Contains("OPTIONAL LIFE", StringComparison.Ordinal))
        {
            return new("OPTIONAL LIFE", "Optional Life", 50, FormattedLineRateGroupKind.Pooled);
        }

        if (normalized.Contains("DEPENDENT LIFE", StringComparison.Ordinal))
        {
            return new("DEPENDENT LIFE", "Dependent Life", 40, FormattedLineRateGroupKind.Pooled);
        }

        if (normalized.Equals("LIFE", StringComparison.Ordinal) ||
            normalized.Contains("BASIC LIFE", StringComparison.Ordinal) ||
            normalized.Contains("EMPLOYEE LIFE", StringComparison.Ordinal))
        {
            return new("LIFE", "Employee Life Insurance", 10, FormattedLineRateGroupKind.Pooled);
        }

        if (IsFormattedAdAndD(normalized))
        {
            return new("AD&D", "Employee AD&D", 20, FormattedLineRateGroupKind.Pooled);
        }

        if (normalized.Contains("CRITICAL ILLNESS", StringComparison.Ordinal))
        {
            if (UnderwritingEngine.IsPerEmployeeRateBasis(rateBasis))
            {
                return new("STAND-ALONE CRITICAL ILLNESS", "Critical Illness", 150, FormattedLineRateGroupKind.Other);
            }

            return new("CRITICAL ILLNESS", "Critical Illness", 30, FormattedLineRateGroupKind.Pooled);
        }

        if (normalized.Equals("STD", StringComparison.Ordinal) ||
            normalized.Contains("SHORT TERM DISABILITY", StringComparison.Ordinal))
        {
            return new("STD", "Short-Term Disability", 70, FormattedLineRateGroupKind.Pooled);
        }

        if (normalized.Equals("LTD", StringComparison.Ordinal) ||
            normalized.Contains("LONG TERM DISABILITY", StringComparison.Ordinal))
        {
            return new("LTD", "Long-Term Disability", 80, FormattedLineRateGroupKind.Pooled);
        }

        if (normalized.StartsWith("EXTENDED HEALTH CARE", StringComparison.Ordinal) ||
            normalized.Equals("EHC", StringComparison.Ordinal))
        {
            return new("EHC", "Extended Health Care", 90, FormattedLineRateGroupKind.ExtendedHealth);
        }

        if (normalized.Contains("DENTAL", StringComparison.Ordinal))
        {
            return new("DENTAL", "Dental Care", 100, FormattedLineRateGroupKind.Dental);
        }

        if (normalized.Contains("EMPLOYEE ASSISTANCE", StringComparison.Ordinal) ||
            (normalized.Contains("EMPLOYEE", StringComparison.Ordinal) && normalized.Contains("FAMILY ASSISTANCE", StringComparison.Ordinal)) ||
            normalized.Equals("EAP", StringComparison.Ordinal) ||
            normalized.Equals("EFAP", StringComparison.Ordinal))
        {
            return new("EAP", "Employee Assistance Program", 110, FormattedLineRateGroupKind.Other);
        }

        if (normalized.Contains("MEDICAL SECOND OPINION", StringComparison.Ordinal) ||
            normalized.Equals("MSO", StringComparison.Ordinal) ||
            normalized.Contains("BEST DOCTORS", StringComparison.Ordinal))
        {
            return new("MSO", "Medical Second Opinion", 120, FormattedLineRateGroupKind.Other);
        }

        if (normalized.Contains("VIRTUAL", StringComparison.Ordinal) ||
            normalized.Contains("TELEMEDICINE", StringComparison.Ordinal) ||
            normalized.Contains("ONLINE DOCTOR", StringComparison.Ordinal) ||
            normalized.Contains("ONLINE ACCESS", StringComparison.Ordinal) ||
            normalized.Contains("HEALTHCARE ONLINE", StringComparison.Ordinal))
        {
            return new("VIRTUAL HEALTH CARE", "Virtual Health Care / Online Access", 130, FormattedLineRateGroupKind.Other);
        }

        if (normalized.Contains("HEALTH SPENDING", StringComparison.Ordinal) ||
            normalized.Equals("HSA", StringComparison.Ordinal) ||
            normalized.Equals("HCSA", StringComparison.Ordinal))
        {
            return new("HSA", "Health Spending Account", 140, FormattedLineRateGroupKind.Other);
        }

        if (normalized.Contains("PERSONAL SPENDING", StringComparison.Ordinal) ||
            normalized.Equals("PSA", StringComparison.Ordinal))
        {
            return new("PSA", "Personal Spending Account", 150, FormattedLineRateGroupKind.Other);
        }

        if (normalized.Contains("WELLNESS", StringComparison.Ordinal))
        {
            return new("WELLNESS", "Wellness", 160, FormattedLineRateGroupKind.Other);
        }

        var displayName = string.IsNullOrWhiteSpace(benefit) ? "Other Benefits" : benefit.Trim();
        return new($"OTHER::{NormalizeLineRateText(displayName)}", displayName, 900, FormattedLineRateGroupKind.Other);
    }

    private static bool IsFormattedAdAndD(string normalizedBenefit)
    {
        return normalizedBenefit.Equals("AD&D", StringComparison.Ordinal) ||
               normalizedBenefit.Equals("AD AND D", StringComparison.Ordinal) ||
               normalizedBenefit.Equals("ASI", StringComparison.Ordinal) ||
               normalizedBenefit.Contains("ACCIDENTAL DEATH", StringComparison.Ordinal) ||
               normalizedBenefit.Contains("DISMEMBERMENT", StringComparison.Ordinal);
    }

    private static string NormalizeLineRateText(string value)
    {
        var cleaned = ParsingHelpers.Clean(value)
            .Replace("&", " AND ", StringComparison.Ordinal)
            .Replace("-", " ", StringComparison.Ordinal)
            .Replace("/", " ", StringComparison.Ordinal)
            .ToUpperInvariant();

        return string.Join(" ", cleaned.Split(' ', StringSplitOptions.RemoveEmptyEntries));
    }

    private static bool IsLineRateTotalRow(string[] row)
    {
        return row.Length > 0 && row[0].Equals("Total Monthly Premiums", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsLineRateDifferenceRow(string[] row)
    {
        return row.Length > 0 && row[0].Equals("Monthly Difference from Current", StringComparison.OrdinalIgnoreCase);
    }

    private static void WriteLineRatePremiumFormulas(IXLWorksheet ws, int row, string benefit, string rateBasis)
    {
        ConvertNumberText(ws.Cell(row, 4), "#,##0.0");
        ConvertNumberText(ws.Cell(row, 5), "#,##0");
        ConvertNumberText(ws.Cell(row, 6), "0.000");
        ConvertNumberText(ws.Cell(row, 8), "0.000");
        ConvertNumberText(ws.Cell(row, 10), "0.000");

        ws.Cell(row, 7).FormulaA1 = LineRatePremiumFormula(row, "F", benefit, rateBasis);
        ws.Cell(row, 9).FormulaA1 = LineRatePremiumFormula(row, "H", benefit, rateBasis);
        ws.Cell(row, 11).FormulaA1 = LineRatePremiumFormula(row, "J", benefit, rateBasis);
        ws.Cell(row, 7).Style.NumberFormat.Format = "$#,##0.00";
        ws.Cell(row, 9).Style.NumberFormat.Format = "$#,##0.00";
        ws.Cell(row, 11).Style.NumberFormat.Format = "$#,##0.00";
    }

    private static void WriteLineRateTotalFormulas(IXLWorksheet ws, int row, int firstDataRow, int lastDataRow)
    {
        if (lastDataRow < firstDataRow)
        {
            return;
        }

        ws.Cell(row, 7).FormulaA1 = $"SUM(G{firstDataRow}:G{lastDataRow})";
        ws.Cell(row, 9).FormulaA1 = $"SUM(I{firstDataRow}:I{lastDataRow})";
        ws.Cell(row, 11).FormulaA1 = $"SUM(K{firstDataRow}:K{lastDataRow})";
        ws.Cell(row, 7).Style.NumberFormat.Format = "$#,##0.00";
        ws.Cell(row, 9).Style.NumberFormat.Format = "$#,##0.00";
        ws.Cell(row, 11).Style.NumberFormat.Format = "$#,##0.00";
    }

    private static void WriteLineRateDifferenceFormulas(IXLWorksheet ws, int row, int totalRow)
    {
        if (totalRow == 0)
        {
            return;
        }

        ws.Cell(row, 9).FormulaA1 = $"IFERROR(I{totalRow}-G{totalRow},\"\")";
        ws.Cell(row, 11).FormulaA1 = $"IFERROR(K{totalRow}-G{totalRow},\"\")";
        ws.Cell(row, 9).Style.NumberFormat.Format = "$#,##0.00";
        ws.Cell(row, 11).Style.NumberFormat.Format = "$#,##0.00";
    }

    private static string LineRatePremiumFormula(int row, string rateColumn, string benefit, string rateBasis)
    {
        if (UnderwritingEngine.IsPerEmployeeRateBasis(rateBasis))
        {
            return $"IF(OR(E{row}=\"\",{rateColumn}{row}=\"\"),\"\",IFERROR(ROUND(E{row}*{rateColumn}{row},2),\"\"))";
        }

        var divisor = RateVolumeDivisor(benefit, rateBasis);
        return divisor.HasValue
            ? $"IF(OR(D{row}=\"\",{rateColumn}{row}=\"\"),\"\",IFERROR(ROUND(D{row}*{rateColumn}{row}/{divisor.Value},2),\"\"))"
            : $"IF(OR(E{row}=\"\",{rateColumn}{row}=\"\"),\"\",IFERROR(ROUND(E{row}*{rateColumn}{row},2),\"\"))";
    }

    private static void ConvertNumberText(IXLCell cell, string numberFormat)
    {
        if (TryParseNumberText(cell.GetString(), out var value))
        {
            cell.Value = value;
            cell.Style.NumberFormat.Format = numberFormat;
        }
    }

    private static bool TryParseNumberText(string text, out decimal value)
    {
        value = 0m;
        var cleaned = text.Trim();
        if (string.IsNullOrWhiteSpace(cleaned) || cleaned.Equals("-", StringComparison.Ordinal))
        {
            return false;
        }

        var isNegative = cleaned.StartsWith("(", StringComparison.Ordinal) &&
                         cleaned.EndsWith(")", StringComparison.Ordinal);
        cleaned = cleaned
            .Trim('(', ')')
            .Replace("$", "", StringComparison.Ordinal)
            .Replace(",", "", StringComparison.Ordinal)
            .Replace("%", "", StringComparison.Ordinal)
            .Trim();

        if (!decimal.TryParse(
                cleaned,
                System.Globalization.NumberStyles.Number,
                System.Globalization.CultureInfo.InvariantCulture,
                out value))
        {
            return false;
        }

        if (isNegative)
        {
            value *= -1;
        }

        return true;
    }

    private static void WriteClaimsAnalysisEhc(XLWorkbook workbook, AnalysisResult analysis)
    {
        WriteClaimsAnalysisServiceCategory(
            workbook,
            "ClaimsAnalysis_EHC",
            analysis.Workbook.EhcSpendingTrends,
            "Warning: ClaimsAnalysis_EHC service category data was not found in the carrier workbook.",
            "Service Category",
            analysis.Workbook.ClaimsTrendValidations.FirstOrDefault(item => item.Area.Equals("ClaimsAnalysis_EHC", StringComparison.OrdinalIgnoreCase)));
    }

    private static void WriteClaimsAnalysisDental(XLWorkbook workbook, AnalysisResult analysis)
    {
        WriteClaimsAnalysisServiceCategory(
            workbook,
            "ClaimsAnalysis_Dental",
            analysis.Workbook.DentalSpendingTrends,
            "Warning: ClaimsAnalysis_Dental service category data was not found in the carrier workbook or claims Summary workbook.",
            "Service Category",
            analysis.Workbook.ClaimsTrendValidations.FirstOrDefault(item => item.Area.Equals("ClaimsAnalysis_Dental", StringComparison.OrdinalIgnoreCase)));
    }

    private static void WriteClaimsAnalysisParamedical(XLWorkbook workbook, AnalysisResult analysis)
    {
        var rows = BuildParamedicalAnalysisTrends(analysis.Workbook.ParamedicalServiceTrends);
        WriteClaimsAnalysisServiceCategory(
            workbook,
            "ClaimsAnalysis_Paramedical",
            rows,
            "Warning: ClaimsAnalysis_Paramedical service data was not found in a claims Summary workbook.",
            "Service",
            analysis.Workbook.ClaimsTrendValidations.FirstOrDefault(item => item.Area.Equals("ClaimsAnalysis_Paramedical", StringComparison.OrdinalIgnoreCase)));

        if (workbook.TryGetWorksheet("ClaimsAnalysis_Paramedical", out var ws))
        {
            WriteParamedicalShareColumn(ws, rows);
        }
    }

    private static List<ServiceCategoryTrend> BuildParamedicalAnalysisTrends(IReadOnlyCollection<ServiceCategoryTrend> trends)
    {
        var sourceRows = trends.ToList();
        var normalizedRows = sourceRows
            .Select(row => new ServiceCategoryTrend
            {
                Benefit = row.Benefit,
                Category = ParamedicalCategoryNormalizer.Normalize(row.Category),
                PriorPeriodLabel = row.PriorPeriodLabel,
                PriorAmount = row.PriorAmount,
                CurrentPeriodLabel = row.CurrentPeriodLabel,
                CurrentAmount = row.CurrentAmount,
                SourceReference = row.SourceReference
            })
            .ToList();

        for (var index = 0; index < sourceRows.Count; index++)
        {
            foreach (var period in ServiceCategoryHistory.GetPeriods(sourceRows[index], int.MaxValue))
            {
                ServiceCategoryHistory.Upsert(normalizedRows[index], period.PeriodLabel, period.Amount, period.SourceReference);
            }
        }
        ServiceCategoryHistory.AlignToSharedWindow(normalizedRows);

        var detailRows = normalizedRows
            .Where(row => !IsParamedicalRollupCategory(row.Category))
            .ToList();
        var selectedRows = detailRows.Count > 0 ? detailRows : normalizedRows;

        var rows = selectedRows
            .GroupBy(row => row.Category, StringComparer.OrdinalIgnoreCase)
            .Select(group => new ServiceCategoryTrend
            {
                Benefit = group.FirstOrDefault(row => !string.IsNullOrWhiteSpace(row.Benefit))?.Benefit ?? "Health",
                Category = group.First().Category,
                PriorPeriodLabel = group.FirstOrDefault(row => !string.IsNullOrWhiteSpace(row.PriorPeriodLabel))?.PriorPeriodLabel ?? "",
                PriorAmount = group.Sum(row => row.PriorAmount),
                CurrentPeriodLabel = group.FirstOrDefault(row => !string.IsNullOrWhiteSpace(row.CurrentPeriodLabel))?.CurrentPeriodLabel ?? "",
                CurrentAmount = group.Sum(row => row.CurrentAmount),
                SourceReference = string.Join("; ", group.Select(row => row.SourceReference).Where(value => !string.IsNullOrWhiteSpace(value)).Distinct(StringComparer.OrdinalIgnoreCase))
            })
            .OrderBy(row => ParamedicalCategorySort(row.Category))
            .ThenBy(row => row.Category, StringComparer.OrdinalIgnoreCase)
            .ToList();

        var rollupRows = normalizedRows
            .Where(row => IsParamedicalRollupCategory(row.Category))
            .ToList();
        if (detailRows.Count > 0 && rollupRows.Count > 0)
        {
            rows.Add(new ServiceCategoryTrend
            {
                Benefit = rollupRows.FirstOrDefault(row => !string.IsNullOrWhiteSpace(row.Benefit))?.Benefit ?? "Health",
                Category = "Paramedical Practitioners",
                PriorPeriodLabel = rollupRows.FirstOrDefault(row => !string.IsNullOrWhiteSpace(row.PriorPeriodLabel))?.PriorPeriodLabel ?? "",
                PriorAmount = rollupRows.Sum(row => row.PriorAmount),
                CurrentPeriodLabel = rollupRows.FirstOrDefault(row => !string.IsNullOrWhiteSpace(row.CurrentPeriodLabel))?.CurrentPeriodLabel ?? "",
                CurrentAmount = rollupRows.Sum(row => row.CurrentAmount),
                SourceReference = string.Join("; ", rollupRows.Select(row => row.SourceReference).Where(value => !string.IsNullOrWhiteSpace(value)).Distinct(StringComparer.OrdinalIgnoreCase))
            });
        }

        foreach (var row in rows)
        {
            var contributors = IsParamedicalRollupCategory(row.Category)
                ? rollupRows
                : selectedRows.Where(source => source.Category.Equals(row.Category, StringComparison.OrdinalIgnoreCase)).ToList();
            foreach (var period in contributors
                         .SelectMany(source => ServiceCategoryHistory.GetPeriods(source, int.MaxValue))
                         .GroupBy(period => period.PeriodLabel, StringComparer.OrdinalIgnoreCase)
                         .Select(group => new
                         {
                             PeriodLabel = group.Key,
                             Amount = group.Sum(period => period.Amount),
                             SourceReference = string.Join("; ", group.Select(period => period.SourceReference).Where(value => !string.IsNullOrWhiteSpace(value)).Distinct(StringComparer.OrdinalIgnoreCase))
                         }))
            {
                ServiceCategoryHistory.Upsert(row, period.PeriodLabel, period.Amount, period.SourceReference);
            }
        }
        ServiceCategoryHistory.AlignToSharedWindow(rows);

        return rows;
    }

    private static int ParamedicalCategorySort(string category)
    {
        return category switch
        {
            "Physiotherapy" => 10,
            "Massage" => 20,
            "Chiropractic" => 30,
            "Naturopath" => 40,
            "Psychologist" => 50,
            "Acupuncture" => 60,
            "Clinical Counsellor" => 70,
            "Social Worker" => 80,
            "Osteopath" => 90,
            "Podiatrist" => 100,
            "Paramedical Practitioners" => 900,
            _ => 500
        };
    }

    private static bool IsParamedicalRollupCategory(string category)
    {
        return ParamedicalCategoryNormalizer.IsRollup(category) ||
               category.Equals("Paramedical Practitioners", StringComparison.OrdinalIgnoreCase);
    }

    private static void WriteParamedicalShareColumn(IXLWorksheet ws, IReadOnlyList<ServiceCategoryTrend> rows)
    {
        if (rows.Count == 0)
        {
            return;
        }

        var periodCount = Math.Max(2, GetServiceCategoryPeriodLabels(rows).Count);
        var latestAmountColumn = periodCount + 1;
        var shareColumn = latestAmountColumn + 1;
        var lastRow = rows.Count + 1;
        var totalRow = rows
            .Select((row, index) => new { row.Category, RowNumber = index + 2 })
            .FirstOrDefault(item => IsParamedicalRollupCategory(item.Category))?.RowNumber;
        if (!totalRow.HasValue)
        {
            totalRow = lastRow + 1;
            ws.Cell(totalRow.Value, 1).Value = "Paramedical Practitioners";
            for (var column = 2; column <= latestAmountColumn; column++)
            {
                var columnName = ColumnName(column);
                ws.Cell(totalRow.Value, column).FormulaA1 = "SUM(" + columnName + "2:" + columnName + lastRow + ")";
            }
            ws.Range(totalRow.Value, 1, totalRow.Value, shareColumn).Style.Font.Bold = true;
            ws.Range(totalRow.Value, 2, totalRow.Value, latestAmountColumn).Style.NumberFormat.Format = "$#,##0";
            lastRow = totalRow.Value;
        }

        ws.Cell(1, shareColumn).Value = "% of parameds";
        Header(ws.Range(1, shareColumn, 1, shareColumn));
        var latestAmountColumnName = ColumnName(latestAmountColumn);
        for (var rowNumber = 2; rowNumber <= lastRow; rowNumber++)
        {
            ws.Cell(rowNumber, shareColumn).FormulaA1 =
                "IFERROR(" + latestAmountColumnName + rowNumber + "/$" + latestAmountColumnName + "$" + totalRow.Value + ",0)";
        }

        ws.Range(2, shareColumn, lastRow, shareColumn).Style.NumberFormat.Format = "0.0%";
        ws.Range(2, shareColumn, lastRow, shareColumn).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
        ws.Range(2, 1, lastRow, shareColumn).Style.Font.FontColor = XLColor.FromHtml("#004650");
        ws.Range(1, 1, lastRow, shareColumn).Style.Border.InsideBorderColor = XLColor.FromHtml("#8DD3C7");
        ws.Range(1, 1, lastRow, shareColumn).Style.Border.OutsideBorderColor = XLColor.FromHtml("#8DD3C7");
        ws.Column(shareColumn).Width = 14;
    }
    private static void WriteHealthRenewalFormulaAnalysis(XLWorkbook workbook, AnalysisResult analysis)
    {
        var ws = workbook.TryGetWorksheet("HealthRenewal_FormulaAnalysis", out var existing)
            ? existing
            : workbook.AddWorksheet("HealthRenewal_FormulaAnalysis");
        ws.Clear();

        var experienceRows = GetHealthFormulaExperienceRows(analysis);
        var periodHeaders = new[] { "Prior Year -1", "Prior Year", "Current Year" };
        var dataColumns = new[] { 3, 4, 5 };
        var healthRateClasses = GetHealthFormulaRateClasses(analysis);
        var defaultTrendInput = GetDefaultFormulaTrendInput(experienceRows, GetDefaultHealthTrendInput(healthRateClasses));
        var defaultCredibility = GetDefaultFormulaCredibility(experienceRows, 0.836m);
        var defaultTargetLossRatio = GetDefaultFormulaTargetLossRatio(experienceRows, 0.71m);
        var defaultCarrierRateChange = GetDefaultFormulaCarrierRateChange(experienceRows, GetDefaultHealthCarrierRateChange(healthRateClasses));
        var defaultPoolingCharge = NormalizeFormulaPercent(analysis.Pdf.PoolingCharge) ?? 0m;
        var experienceWeights = GetFormulaExperienceWeights(experienceRows, dataColumns.Length);

        ws.Range("A1:G1").Merge();
        ws.Cell("A1").Value = "Health Renewal Formula Analysis";
        ws.Cell("A1").Style.Font.FontColor = XLColor.White;
        ws.Cell("A1").Style.Font.FontSize = 14;
        ws.Cell("A1").Style.Font.Bold = true;
        ws.Cell("A1").Style.Fill.BackgroundColor = XLColor.FromHtml("#3399B3");
        ws.Cell("A1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        ws.Cell("A1").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        ws.Row(1).Height = 24;

        for (var index = 0; index < periodHeaders.Length; index++)
        {
            ws.Cell(3, dataColumns[index]).Value = periodHeaders[index];
            ws.Cell(3, dataColumns[index]).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        }
        ApplySectionHeader(ws.Range("C3:E3"));

        var experienceLabels = new (int Row, string Label)[]
        {
            (5, "From:"),
            (6, "To:"),
            (8, "Paid Premiums:"),
            (9, "Paid Claims:"),
            (10, "Incurred but Not Realized:"),
            (11, "Incurred Claims:"),
            (12, "Incurred Loss Ratio:"),
            (14, "Trend Factor:"),
            (15, "Trended Loss Ratio:"),
            (17, "Target Loss Ratio:"),
            (18, "Projected Loss Ratio:"),
            (20, "Weighting:"),
            (22, "Experience Result:")
        };

        foreach (var (row, label) in experienceLabels)
        {
            ws.Cell(row, 1).Value = label;
            ws.Cell(row, 1).Style.Font.FontColor = XLColor.FromHtml("#404040");
        }

        ws.Range("C5:E6").Style.Fill.BackgroundColor = XLColor.FromHtml("#63C4B0");
        ws.Range("C5:E6").Style.Font.FontColor = XLColor.White;
        ws.Range("C5:E6").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        ws.Range("C5:E6").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

        for (var index = 0; index < dataColumns.Length; index++)
        {
            var column = dataColumns[index];
            var source = GetRightAlignedExperienceRow(experienceRows, index, dataColumns.Length);
            var (from, to) = SplitPeriod(source?.Period);
            SetOptionalText(ws.Cell(5, column), from);
            SetOptionalText(ws.Cell(6, column), to);
            ws.Cell(8, column).Value = source?.BilledPremium ?? 0m;
            ws.Cell(9, column).Value = source?.PaidClaims ?? 0m;
            ws.Cell(10, column).Value = (source?.IncurredClaims ?? 0m) - (source?.PaidClaims ?? 0m);
            ws.Cell(11, column).FormulaA1 = $"=IFERROR({ColumnName(column)}9+{ColumnName(column)}10,0)";
            ws.Cell(12, column).FormulaA1 = $"=IFERROR({ColumnName(column)}11/{ColumnName(column)}8,0)";
            var trendYears = dataColumns.Length - index;
            ws.Cell(14, column).FormulaA1 = trendYears <= 1
                ? "=IFERROR(1+$B$26,0)"
                : $"=IFERROR((1+$B$26)^{trendYears},0)";
            ws.Cell(15, column).FormulaA1 = $"=IFERROR({ColumnName(column)}12*{ColumnName(column)}14,0)";
            ws.Cell(17, column).FormulaA1 = "=$B$28";
            ws.Cell(18, column).FormulaA1 = $"=IFERROR({ColumnName(column)}15/{ColumnName(column)}17,0)";
            ws.Cell(20, column).Value = experienceWeights[index];
        }

        ws.Range("C8:E11").Style.NumberFormat.Format = "$#,##0";
        ws.Range("C12:E20").Style.NumberFormat.Format = "0.00%";
        ws.Range("C8:E20").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
        ws.Cell("C22").FormulaA1 = "=IFERROR(SUMPRODUCT(C18:E18,C20:E20)/SUM(C20:E20),0)";
        ws.Range("C22:E22").Merge();
        ws.Cell("C22").Style.NumberFormat.Format = "0.00%";
        ws.Cell("C22").Style.Font.Bold = true;
        ws.Cell("C22").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        ws.Cell("G22").Value = "Pooling Charge:";
        ws.Cell("H22").Value = defaultPoolingCharge;
        ws.Cell("H22").Style.NumberFormat.Format = "0.00%";
        ws.Cell("H22").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
        ws.Cell("G23").Value = "Pooling Threshold:";
        ws.Cell("H23").Value = analysis.Pdf.PoolingThreshold ?? 0m;
        ws.Cell("H23").Style.NumberFormat.Format = "$#,##0";
        ws.Cell("H23").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
        ws.Range("G22:H23").Style.Border.InsideBorder = XLBorderStyleValues.Thin;
        ws.Range("G22:H23").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        ws.Range("G22:H23").Style.Border.InsideBorderColor = XLColor.FromHtml("#A6A6A6");
        ws.Range("G22:H23").Style.Border.OutsideBorderColor = XLColor.FromHtml("#A6A6A6");

        ws.Range("A25:B25").Merge();
        ws.Cell("A25").Value = "Extended Health Renewal Factors";
        ApplySectionHeader(ws.Range("A25:B25"));
        ws.Cell("A26").Value = "Trend Factor:";
        ws.Cell("A27").Value = "Credibility:";
        ws.Cell("A28").Value = "Target Loss Ratio:";
        ws.Cell("A30").Value = "Class A";
        ws.Cell("A31").Value = "Class B";
        ws.Cell("A32").Value = "Class C";
        ws.Cell("B26").Value = defaultTrendInput;
        ws.Cell("B27").Value = defaultCredibility;
        ws.Cell("B28").Value = defaultTargetLossRatio;
        ws.Range("B26:B28").Style.NumberFormat.Format = "0.00%";
        ws.Range("B26:B28").Style.Fill.BackgroundColor = XLColor.FromHtml("#EAF2DE");
        ws.Range("A26:B32").Style.Font.FontColor = XLColor.FromHtml("#404040");

        WriteHealthRateProjectionBlocks(ws, healthRateClasses);
        WriteHealthFormulaSummaryBox(ws, defaultCarrierRateChange);
        WritePooledPremiumCrossCheck(ws, analysis);
        FormatHealthRenewalFormulaSheet(ws);
    }

    private static void WriteDentalRenewalFormulaAnalysis(XLWorkbook workbook, AnalysisResult analysis)
    {
        var ws = workbook.TryGetWorksheet("DentalRenewal_FormulaAnalysis", out var existing)
            ? existing
            : workbook.AddWorksheet("DentalRenewal_FormulaAnalysis");
        ws.Clear();

        var experienceRows = GetDentalFormulaExperienceRows(analysis);
        var periodHeaders = new[] { "Prior Year -1", "Prior Year", "Current Year" };
        var dataColumns = new[] { 3, 4, 5 };
        var dentalRateClasses = GetDentalFormulaRateClasses(analysis);
        var defaultTrendInput = GetDefaultFormulaTrendInput(experienceRows, GetDefaultHealthTrendInput(dentalRateClasses));
        var defaultCredibility = GetDefaultFormulaCredibility(experienceRows, 0.836m);
        var defaultCarrierRateChange = GetDefaultFormulaCarrierRateChange(experienceRows, GetDefaultHealthCarrierRateChange(dentalRateClasses));
        var defaultTargetLossRatio = GetDefaultFormulaTargetLossRatio(experienceRows, GetDefaultDentalTargetLossRatio(analysis));
        var experienceWeights = GetFormulaExperienceWeights(experienceRows, dataColumns.Length);

        ws.Range("A1:G1").Merge();
        ws.Cell("A1").Value = "Dental Renewal Formula Analysis";
        ws.Cell("A1").Style.Font.FontColor = XLColor.White;
        ws.Cell("A1").Style.Font.FontSize = 14;
        ws.Cell("A1").Style.Font.Bold = true;
        ws.Cell("A1").Style.Fill.BackgroundColor = XLColor.FromHtml("#3399B3");
        ws.Cell("A1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        ws.Cell("A1").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        ws.Row(1).Height = 24;

        for (var index = 0; index < periodHeaders.Length; index++)
        {
            ws.Cell(3, dataColumns[index]).Value = periodHeaders[index];
            ws.Cell(3, dataColumns[index]).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        }
        ApplySectionHeader(ws.Range("C3:E3"));

        var experienceLabels = new (int Row, string Label)[]
        {
            (5, "From:"),
            (6, "To:"),
            (8, "Paid Premiums:"),
            (9, "Paid Claims:"),
            (10, "Incurred but Not Realized:"),
            (11, "Incurred Claims:"),
            (12, "Incurred Loss Ratio:"),
            (14, "Trend Factor:"),
            (15, "Trended Loss Ratio:"),
            (17, "Target Loss Ratio:"),
            (18, "Projected Loss Ratio:"),
            (20, "Weighting:"),
            (22, "Experience Result:")
        };

        foreach (var (row, label) in experienceLabels)
        {
            ws.Cell(row, 1).Value = label;
            ws.Cell(row, 1).Style.Font.FontColor = XLColor.FromHtml("#404040");
        }

        ws.Range("C5:E6").Style.Fill.BackgroundColor = XLColor.FromHtml("#63C4B0");
        ws.Range("C5:E6").Style.Font.FontColor = XLColor.White;
        ws.Range("C5:E6").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        ws.Range("C5:E6").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

        for (var index = 0; index < dataColumns.Length; index++)
        {
            var column = dataColumns[index];
            var source = GetRightAlignedExperienceRow(experienceRows, index, dataColumns.Length);
            var (from, to) = SplitPeriod(source?.Period);
            SetOptionalText(ws.Cell(5, column), from);
            SetOptionalText(ws.Cell(6, column), to);
            ws.Cell(8, column).Value = source?.BilledPremium ?? 0m;
            ws.Cell(9, column).Value = source?.PaidClaims ?? 0m;
            ws.Cell(10, column).Value = (source?.IncurredClaims ?? 0m) - (source?.PaidClaims ?? 0m);
            ws.Cell(11, column).FormulaA1 = $"=IFERROR({ColumnName(column)}9+{ColumnName(column)}10,0)";
            ws.Cell(12, column).FormulaA1 = $"=IFERROR({ColumnName(column)}11/{ColumnName(column)}8,0)";
            var trendYears = dataColumns.Length - index;
            ws.Cell(14, column).FormulaA1 = trendYears <= 1
                ? "=IFERROR(1+$B$26,0)"
                : $"=IFERROR((1+$B$26)^{trendYears},0)";
            ws.Cell(15, column).FormulaA1 = $"=IFERROR({ColumnName(column)}12*{ColumnName(column)}14,0)";
            ws.Cell(17, column).FormulaA1 = "=$B$28";
            ws.Cell(18, column).FormulaA1 = $"=IFERROR({ColumnName(column)}15/{ColumnName(column)}17,0)";
            ws.Cell(20, column).Value = experienceWeights[index];
        }

        ws.Range("C8:E11").Style.NumberFormat.Format = "$#,##0";
        ws.Range("C12:E20").Style.NumberFormat.Format = "0.00%";
        ws.Range("C8:E20").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
        ws.Cell("C22").FormulaA1 = "=IFERROR(SUMPRODUCT(C18:E18,C20:E20)/SUM(C20:E20),0)";
        ws.Range("C22:E22").Merge();
        ws.Cell("C22").Style.NumberFormat.Format = "0.00%";
        ws.Cell("C22").Style.Font.Bold = true;
        ws.Cell("C22").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

        ws.Range("A25:B25").Merge();
        ws.Cell("A25").Value = "Dental Renewal Factors";
        ApplySectionHeader(ws.Range("A25:B25"));
        ws.Cell("A26").Value = "Trend Factor:";
        ws.Cell("A27").Value = "Credibility:";
        ws.Cell("A28").Value = "Target Loss Ratio:";
        ws.Cell("A30").Value = "Class A";
        ws.Cell("A31").Value = "Class B";
        ws.Cell("A32").Value = "Class C";
        ws.Cell("B26").Value = defaultTrendInput;
        ws.Cell("B27").Value = defaultCredibility;
        ws.Cell("B28").Value = defaultTargetLossRatio;
        ws.Range("B26:B28").Style.NumberFormat.Format = "0.00%";
        ws.Range("B26:B28").Style.Fill.BackgroundColor = XLColor.FromHtml("#EAF2DE");
        ws.Range("A26:B32").Style.Font.FontColor = XLColor.FromHtml("#404040");

        WriteHealthRateProjectionBlocks(ws, dentalRateClasses, "Dental");
        WriteHealthFormulaSummaryBox(ws, defaultCarrierRateChange);
        FormatHealthRenewalFormulaSheet(ws);
    }

    private static void WriteHealthRateProjectionBlocks(IXLWorksheet ws, IReadOnlyList<HealthFormulaRateClass> healthRateClasses, string projectionBenefitLabel = "Extended Health")
    {
        var row = 35;
        for (var classIndex = 0; classIndex < 3; classIndex++)
        {
            var className = $"Class {(char)('A' + classIndex)}";
            var classRates = healthRateClasses.Count > classIndex
                ? healthRateClasses[classIndex]
                : healthRateClasses.FirstOrDefault() ?? HealthFormulaRateClass.Empty;

            ws.Cell(row, 1).Value = className;
            ws.Cell(row, 1).Style.Font.FontColor = XLColor.FromHtml("#666666");
            ws.Cell(row, 1).Style.Font.FontName = "Consolas";
            row++;

            foreach (var tier in new[] { "Single", "Couple", "Family" })
            {
                classRates.RatesByTier.TryGetValue(NormalizeHealthFormulaTier(tier), out var tierRate);
                WriteHealthRateProjectionBlock(ws, row, $"{tier} {projectionBenefitLabel} Rate Projection", tierRate);
                row += 5;
            }

            row++;
        }
    }

    private static void WriteHealthRateProjectionBlock(IXLWorksheet ws, int row, string title, HealthFormulaTierRate? tierRate)
    {
        ws.Range(row, 1, row, 6).Merge();
        ws.Cell(row, 1).Value = title;
        ApplyDarkSectionHeader(ws.Range(row, 1, row, 6));

        var experienceRow = row + 1;
        var pooledRow = row + 2;
        ws.Cell(experienceRow, 1).Value = "Experience";
        ws.Cell(pooledRow, 1).Value = "Pooled";
        SetMoney(ws.Cell(experienceRow, 2), tierRate?.CurrentRate);
        ws.Cell(experienceRow, 3).FormulaA1 = "=$C$22";
        ws.Cell(experienceRow, 4).FormulaA1 = "=$B$27";
        ws.Cell(experienceRow, 5).FormulaA1 = $"=IFERROR(B{experienceRow}*C{experienceRow}*D{experienceRow},0)";
        ws.Cell(pooledRow, 2).FormulaA1 = $"=IFERROR(B{experienceRow}*(1+$B$26),0)";
        ws.Cell(pooledRow, 3).Clear();
        ws.Cell(pooledRow, 4).FormulaA1 = "=IFERROR(1-$B$27,0)";
        ws.Cell(pooledRow, 5).FormulaA1 = $"=IFERROR(B{pooledRow}*D{pooledRow},0)";
        ws.Cell(experienceRow, 6).FormulaA1 = $"=IFERROR(E{experienceRow}+E{pooledRow},0)";
        ws.Cell(pooledRow, 6).FormulaA1 = $"=F{experienceRow}";

        ws.Range(experienceRow, 2, pooledRow, 2).Style.NumberFormat.Format = "$#,##0.00";
        ws.Range(experienceRow, 3, pooledRow, 4).Style.NumberFormat.Format = "0.00%";
        ws.Range(experienceRow, 5, pooledRow, 6).Style.NumberFormat.Format = "$#,##0.00";
        ws.Range(experienceRow, 1, pooledRow, 6).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
        ws.Range(experienceRow, 1, pooledRow, 6).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        ws.Range(experienceRow, 1, pooledRow, 6).Style.Border.InsideBorderColor = XLColor.FromHtml("#A6A6A6");
        ws.Range(experienceRow, 1, pooledRow, 6).Style.Border.OutsideBorderColor = XLColor.FromHtml("#A6A6A6");
        ws.Range(experienceRow, 6, pooledRow, 6).Style.Font.Bold = true;
        ws.Range(experienceRow, 6, pooledRow, 6).Style.Fill.BackgroundColor = SoftFill;
        ws.Range(experienceRow, 6, pooledRow, 6).Style.Font.FontColor = XLColor.FromHtml("#004650");
    }

    private static void WriteHealthFormulaSummaryBox(IXLWorksheet ws, decimal defaultCarrierRateChange)
    {
        ws.Cell("I6").Value = "Current Renewal";
        ws.Cell("J6").Value = "Prior Renewal";
        ApplySectionHeader(ws.Range("I6:J6"));
        ws.Range("I6:J6").Style.Font.Bold = true;
        ws.Range("I6:J6").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

        var labels = new[]
        {
            "Net Loss ratio",
            "Inflation/utilization factor",
            "Trended loss ratio",
            "Break-even point",
            "Calculated Rate Change",
            "Carrier Rate Change",
            "Negotiated Rate Change"
        };

        for (var index = 0; index < labels.Length; index++)
        {
            ws.Cell(7 + index, 8).Value = labels[index];
        }

        ws.Cell("I7").FormulaA1 = "=IFERROR(E12,0)";
        ws.Cell("J7").FormulaA1 = "=IFERROR(D12,0)";
        ws.Cell("I8").FormulaA1 = "=$B$26";
        ws.Cell("J8").FormulaA1 = "=$B$26";
        ws.Cell("I9").FormulaA1 = "=IFERROR(I7*(1+I8),0)";
        ws.Cell("J9").FormulaA1 = "=IFERROR(J7*(1+J8),0)";
        ws.Cell("I10").FormulaA1 = "=$B$28";
        ws.Cell("J10").FormulaA1 = "=$B$28";
        ws.Cell("I11").FormulaA1 = "=IFERROR(I9/I10-1,0)";
        ws.Cell("J11").FormulaA1 = "=IFERROR(J9/J10-1,0)";
        ws.Cell("I12").Value = defaultCarrierRateChange;
        ws.Cell("J12").Value = 0m;
        ws.Range("I13:J13").Value = 0m;
        ws.Range("I7:J13").Style.NumberFormat.Format = "0.00%";
        ws.Range("I12:J13").Style.Fill.BackgroundColor = XLColor.FromHtml("#F8C8D8");
        ws.Range("H6:J13").Style.Fill.BackgroundColor = XLColor.White;
        ws.Range("I6:J6").Style.Fill.BackgroundColor = XLColor.FromHtml("#3399B3");
        ws.Range("I6:J6").Style.Font.FontColor = XLColor.White;
        ws.Range("I12:J13").Style.Fill.BackgroundColor = XLColor.FromHtml("#F8C8D8");
        ws.Range("H6:J13").Style.Border.InsideBorder = XLBorderStyleValues.Thin;
        ws.Range("H6:J13").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        ws.Range("H6:J13").Style.Border.InsideBorderColor = XLColor.FromHtml("#A6A6A6");
        ws.Range("H6:J13").Style.Border.OutsideBorderColor = XLColor.FromHtml("#A6A6A6");
        ws.Range("H7:H13").Style.Font.FontColor = XLColor.FromHtml("#404040");
        ws.Range("I7:J13").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
    }

    private static void WritePooledPremiumCrossCheck(IXLWorksheet ws, AnalysisResult analysis)
    {
        var carrierAnnual = analysis.Pdf.PooledPremium;
        var estimatedAnnual = GetEstimatedAnnualPooledPremium(analysis);
        var difference = estimatedAnnual.HasValue && carrierAnnual.HasValue
            ? decimal.Round(estimatedAnnual.Value - carrierAnnual.Value, 2)
            : (decimal?)null;
        var differencePercent = difference.HasValue && carrierAnnual is > 0m
            ? difference.Value / carrierAnnual.Value
            : (decimal?)null;
        var status = carrierAnnual.HasValue && estimatedAnnual.HasValue
            ? Math.Abs(difference ?? 0m) <= 25m || Math.Abs(differencePercent ?? 0m) <= 0.01m
                ? "Matched"
                : "Review"
            : "Needs data";

        ws.Range("H15:K15").Merge();
        ws.Cell("H15").Value = "Pooled Premium Cross-Check";
        ApplySectionHeader(ws.Range("H15:K15"));

        ws.Cell("H16").Value = "Carrier annual pooled premium";
        ws.Cell("H17").Value = "Estimated annual pooled premium";
        ws.Cell("H18").Value = "Difference";
        ws.Cell("H19").Value = "Difference %";
        ws.Cell("H20").Value = "Status";
        ws.Cell("H21").Value = "Pooled claims";

        SetMoneyWhole(ws.Cell("I16"), carrierAnnual);
        SetMoneyWhole(ws.Cell("I17"), estimatedAnnual);
        SetMoneyWhole(ws.Cell("I18"), difference);
        SetPercentValue(ws.Cell("I19"), differencePercent);
        ws.Cell("I20").Value = status;
        ws.Range("I20:K20").Merge();
        ws.Cell("I21").Value = FormatPooledClaimsDetail(analysis);
        ws.Range("I21:K21").Merge();
        ws.Range("H16:H21").Style.Font.Bold = true;
        ws.Range("H16:K21").Style.Fill.BackgroundColor = XLColor.White;
        ws.Range("I21:K21").Style.Alignment.WrapText = true;

        if (status.Equals("Review", StringComparison.OrdinalIgnoreCase))
        {
            ws.Range("I20:K20").Style.Fill.BackgroundColor = WarningFill;
        }
        else if (status.Equals("Matched", StringComparison.OrdinalIgnoreCase))
        {
            ws.Range("I20:K20").Style.Fill.BackgroundColor = XLColor.FromHtml("#E2F0D9");
        }
    }

    private static decimal? GetEstimatedAnnualPooledPremium(AnalysisResult analysis)
    {
        var pooledLines = SelectedRateLines(analysis)
            .Where(line => line.Benefit.Equals("Extended Health Care - Pooled", StringComparison.OrdinalIgnoreCase))
            .ToList();

        if (pooledLines.Count == 0)
        {
            return null;
        }

        var monthly = pooledLines.Sum(line =>
            line.CalculatedRenewalPremium ??
            UnderwritingEngine.CalculatePremium(line.Benefit, line.Volume, line.Lives, line.RenewalRate, line.RateBasis) ??
            line.RenewalPremium ??
            0m);

        return monthly > 0m ? decimal.Round(monthly * 12m, 2) : null;
    }

    private static string FormatPooledClaimsDetail(AnalysisResult analysis)
    {
        var parts = new List<string>();
        if (analysis.Pdf.DrugStopLossClaims.HasValue)
        {
            parts.Add($"Drug stop-loss claims {ParsingHelpers.Money(analysis.Pdf.DrugStopLossClaims)}");
        }

        if (analysis.Pdf.OtherPooledClaims.HasValue)
        {
            parts.Add($"Other pooled claims {ParsingHelpers.Money(analysis.Pdf.OtherPooledClaims)}");
        }

        return parts.Count == 0 ? "" : string.Join("; ", parts);
    }

    private static void FormatHealthRenewalFormulaSheet(IXLWorksheet ws)
    {
        ws.Style.Fill.BackgroundColor = XLColor.White;
        ws.Style.Font.FontName = "Calibri";
        ws.Style.Font.FontColor = XLColor.FromHtml("#404040");
        ws.Columns("A:K").Width = 12;
        ws.Column(1).Width = 25;
        ws.Column(2).Width = 10;
        ws.Columns("C:E").Width = 14;
        ws.Column(6).Width = 12;
        ws.Column(7).Width = 12;
        ws.Column(8).Width = 27;
        ws.Columns("I:J").Width = 15;
        ws.Column(11).Width = 12;
        ws.Rows().Height = 17;
        ws.Row(1).Height = 24;
        ws.Row(3).Height = 18;

        ws.Range("A1:G1").Style.Fill.BackgroundColor = XLColor.FromHtml("#3399B3");
        ws.Range("A1:G1").Style.Font.FontColor = XLColor.White;
        ws.Range("A1:G1").Style.Font.Bold = true;
        ws.Range("A1:G1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        ws.Range("A1:G1").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

        ws.Range("C3:E3").Style.Fill.BackgroundColor = XLColor.FromHtml("#3399B3");
        ws.Range("C3:E3").Style.Font.FontColor = XLColor.White;
        ws.Range("C3:E3").Style.Font.Bold = true;
        ws.Range("C3:E3").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

        ws.Range("A3:E22").Style.Border.InsideBorder = XLBorderStyleValues.Thin;
        ws.Range("A3:E22").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        ws.Range("A3:E22").Style.Border.InsideBorderColor = XLColor.FromHtml("#A6A6A6");
        ws.Range("A3:E22").Style.Border.OutsideBorderColor = XLColor.FromHtml("#A6A6A6");
        ws.Range("A25:B32").Style.Border.InsideBorder = XLBorderStyleValues.Thin;
        ws.Range("A25:B32").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        ws.Range("A25:B32").Style.Border.InsideBorderColor = XLColor.FromHtml("#A6A6A6");
        ws.Range("A25:B32").Style.Border.OutsideBorderColor = XLColor.FromHtml("#A6A6A6");
        ws.Range("A25:B32").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        ws.Range("B26:B28").Style.Fill.BackgroundColor = XLColor.FromHtml("#EAF2DE");
        ws.Range("B26:B28").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
        ws.Range("C8:E11").Style.NumberFormat.Format = "$#,##0";
        ws.Range("C12:E20").Style.NumberFormat.Format = "0.00%";
        ws.Range("C8:E22").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
        ws.Range("C11:E11").Style.Fill.BackgroundColor = XLColor.FromHtml("#EAF2DE");
        ws.Range("C20:E20").Style.Fill.BackgroundColor = XLColor.FromHtml("#EAF2DE");
        ws.Range("C22:E22").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        ws.Range("C22:E22").Style.Fill.BackgroundColor = XLColor.White;
        ws.Range("C22:E22").Style.Font.FontColor = XLColor.FromHtml("#0070C0");
        ws.Range("A22:E22").Style.Border.TopBorder = XLBorderStyleValues.Thin;
        ws.Range("A22:E22").Style.Border.TopBorderColor = XLColor.FromHtml("#A6A6A6");

        ws.Range("H6:J13").Style.Border.InsideBorder = XLBorderStyleValues.Thin;
        ws.Range("H6:J13").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        ws.Range("H6:J13").Style.Border.InsideBorderColor = XLColor.FromHtml("#A6A6A6");
        ws.Range("H6:J13").Style.Border.OutsideBorderColor = XLColor.FromHtml("#A6A6A6");
        ws.Range("I6:J6").Style.Fill.BackgroundColor = XLColor.FromHtml("#3399B3");
        ws.Range("I6:J6").Style.Font.FontColor = XLColor.White;
        ws.Range("I6:J6").Style.Font.Bold = true;
        ws.Range("I7:J13").Style.NumberFormat.Format = "0.00%";
        ws.Range("I7:J13").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
        ws.Range("I12:J13").Style.Fill.BackgroundColor = XLColor.FromHtml("#F8C8D8");

        ws.Range("A35:F90").Style.Border.InsideBorder = XLBorderStyleValues.Thin;
        ws.Range("A35:F90").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        ws.Range("A35:F90").Style.Border.InsideBorderColor = XLColor.FromHtml("#A6A6A6");
        ws.Range("A35:F90").Style.Border.OutsideBorderColor = XLColor.FromHtml("#A6A6A6");
        ws.Range("B35:F90").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
        ws.Range("F35:F90").Style.Font.Bold = true;
        ws.Range("F35:F90").Style.Font.FontColor = XLColor.FromHtml("#004650");
        ws.Range("H16:K23").Style.Border.InsideBorder = XLBorderStyleValues.Thin;
        ws.Range("H16:K23").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        ws.Range("H16:K23").Style.Border.InsideBorderColor = XLColor.FromHtml("#A6A6A6");
        ws.Range("H16:K23").Style.Border.OutsideBorderColor = XLColor.FromHtml("#A6A6A6");
        ws.Range("I18:K21").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
        for (var row = 35; row <= 90; row++)
        {
            if (ws.Cell(row, 1).GetString().Contains("Rate Projection", StringComparison.OrdinalIgnoreCase))
            {
                ApplyDarkSectionHeader(ws.Range(row, 1, row, 6));
            }
        }
        ws.SheetView.FreezeRows(3);
    }

    private static IReadOnlyList<HealthFormulaRateClass> GetHealthFormulaRateClasses(AnalysisResult analysis)
    {
        var ehcLines = SelectedRateLines(analysis)
            .Where(line => IsHealthFormulaExtendedHealth(line.Benefit))
            .Where(IsHealthFormulaMeaningfulRateLine)
            .ToList();

        if (ehcLines.Count == 0)
        {
            return [HealthFormulaRateClass.Empty];
        }

        var hasExplicitTotals = ehcLines.Any(line => IsHealthFormulaExtendedHealthTotal(line.Benefit));
        var selectedLines = hasExplicitTotals
            ? ehcLines.Where(line => IsHealthFormulaExtendedHealthTotal(line.Benefit)).ToList()
            : ehcLines;

        var rateClasses = selectedLines
            .Select((line, index) => new { Line = line, Index = index })
            .GroupBy(item => HealthFormulaAccountKey(item.Line))
            .OrderBy(group => group.Min(item => item.Index))
            .Select(group =>
            {
                var accountLabel = HealthFormulaAccountLabel(group.First().Line);
                var tierRates = group
                    .GroupBy(item => NormalizeHealthFormulaTier(item.Line.TierOrClass))
                    .Where(tierGroup => !string.IsNullOrWhiteSpace(tierGroup.Key))
                    .ToDictionary(
                        tierGroup => tierGroup.Key,
                        tierGroup => BuildHealthFormulaTierRate(tierGroup.Select(item => item.Line).ToList(), hasExplicitTotals),
                        StringComparer.OrdinalIgnoreCase);

                return new HealthFormulaRateClass(accountLabel, tierRates);
            })
            .Where(rateClass => rateClass.RatesByTier.Count > 0)
            .ToList();

        return rateClasses.Count == 0 ? [HealthFormulaRateClass.Empty] : rateClasses;
    }

    private static IReadOnlyList<HealthFormulaRateClass> GetDentalFormulaRateClasses(AnalysisResult analysis)
    {
        var dentalLines = SelectedRateLines(analysis)
            .Where(line => IsDentalFormulaBenefit(line.Benefit))
            .Where(IsHealthFormulaMeaningfulRateLine)
            .ToList();

        if (dentalLines.Count == 0)
        {
            return [HealthFormulaRateClass.Empty];
        }

        var hasExplicitTotals = dentalLines.Any(line => IsDentalFormulaTotal(line.Benefit));
        var selectedLines = hasExplicitTotals
            ? dentalLines.Where(line => IsDentalFormulaTotal(line.Benefit)).ToList()
            : dentalLines;

        var rateClasses = selectedLines
            .Select((line, index) => new { Line = line, Index = index })
            .GroupBy(item => HealthFormulaAccountKey(item.Line))
            .OrderBy(group => group.Min(item => item.Index))
            .Select(group =>
            {
                var accountLabel = HealthFormulaAccountLabel(group.First().Line);
                var tierRates = group
                    .GroupBy(item => NormalizeHealthFormulaTier(item.Line.TierOrClass))
                    .Where(tierGroup => !string.IsNullOrWhiteSpace(tierGroup.Key))
                    .ToDictionary(
                        tierGroup => tierGroup.Key,
                        tierGroup => BuildHealthFormulaTierRate(tierGroup.Select(item => item.Line).ToList(), hasExplicitTotals),
                        StringComparer.OrdinalIgnoreCase);

                return new HealthFormulaRateClass(accountLabel, tierRates);
            })
            .Where(rateClass => rateClass.RatesByTier.Count > 0)
            .ToList();

        return rateClasses.Count == 0 ? [HealthFormulaRateClass.Empty] : rateClasses;
    }

    private static HealthFormulaTierRate BuildHealthFormulaTierRate(List<RateLine> lines, bool explicitTotalRows)
    {
        decimal? CombineRates(IEnumerable<decimal?> values)
        {
            var numbers = values.Where(value => value.HasValue).Select(value => value!.Value).ToList();
            if (numbers.Count == 0)
            {
                return null;
            }

            return explicitTotalRows
                ? numbers.First()
                : numbers.Sum();
        }

        return new HealthFormulaTierRate(
            CombineRates(lines.Select(line => line.CurrentRate)),
            CombineRates(lines.Select(line => line.RenewalRate)),
            lines.Sum(line => line.CurrentPremium ?? 0m),
            lines.Sum(line => line.RenewalPremium ?? 0m));
    }

    private static decimal GetDefaultHealthTrendInput(IReadOnlyList<HealthFormulaRateClass> rateClasses)
    {
        var currentPremium = rateClasses.SelectMany(rateClass => rateClass.RatesByTier.Values).Sum(rate => rate.CurrentPremium);
        var renewalPremium = rateClasses.SelectMany(rateClass => rateClass.RatesByTier.Values).Sum(rate => rate.RenewalPremium);
        if (currentPremium > 0m && renewalPremium > 0m)
        {
            return decimal.Round(renewalPremium / currentPremium - 1m, 4);
        }

        return 0.118m;
    }

    private static decimal GetDefaultFormulaTrendInput(IEnumerable<ExperienceLine> experienceRows, decimal fallback)
    {
        return OrderFormulaExperienceRowsLatestFirst(experienceRows)
            .Select(row => NormalizeFormulaPercent(row.TrendFactor))
            .FirstOrDefault(value => value.HasValue && value.Value > 0m) ?? fallback;
    }

    private static decimal GetDefaultFormulaCredibility(IEnumerable<ExperienceLine> experienceRows, decimal fallback)
    {
        return OrderFormulaExperienceRowsLatestFirst(experienceRows)
            .Select(row => NormalizeFormulaPercent(row.Credibility))
            .FirstOrDefault(value => value.HasValue && value.Value > 0m) ?? fallback;
    }

    private static decimal GetDefaultFormulaTargetLossRatio(IEnumerable<ExperienceLine> experienceRows, decimal fallback)
    {
        return OrderFormulaExperienceRowsLatestFirst(experienceRows)
            .Select(row => NormalizeFormulaPercent(row.TargetLossRatio ?? row.Weighting))
            .FirstOrDefault(value => value.HasValue && value.Value > 0m) ?? fallback;
    }

    private static decimal GetDefaultFormulaCarrierRateChange(IEnumerable<ExperienceLine> experienceRows, decimal fallback)
    {
        return OrderFormulaExperienceRowsLatestFirst(experienceRows)
            .Select(row => NormalizeFormulaPercent(row.ProposedAdjustment))
            .FirstOrDefault(value => value.HasValue) ?? fallback;
    }

    private static IEnumerable<ExperienceLine> OrderFormulaExperienceRowsLatestFirst(IEnumerable<ExperienceLine> experienceRows)
    {
        return experienceRows
            .OrderByDescending(row => GetFormulaPeriodStart(row) ?? DateOnly.MinValue)
            .ThenBy(row => row.SourceReference.StartsWith("Previous renewal PDF", StringComparison.OrdinalIgnoreCase) ? 1 : 0);
    }

    private static decimal GetDefaultHealthCarrierRateChange(IReadOnlyList<HealthFormulaRateClass> rateClasses)
    {
        return GetDefaultHealthTrendInput(rateClasses);
    }

    private static bool IsHealthFormulaMeaningfulRateLine(RateLine line)
    {
        return line.CurrentRate.HasValue ||
               line.RenewalRate.HasValue ||
               line.CurrentPremium.HasValue ||
               line.RenewalPremium.HasValue;
    }

    private static bool IsHealthFormulaExtendedHealth(string benefit)
    {
        return NormalizeHealthFormulaText(benefit).StartsWith("EXTENDED HEALTH CARE", StringComparison.Ordinal);
    }

    private static bool IsHealthFormulaExtendedHealthTotal(string benefit)
    {
        var normalized = NormalizeHealthFormulaText(benefit);
        return normalized.StartsWith("EXTENDED HEALTH CARE", StringComparison.Ordinal) &&
               normalized.Contains("TOTAL", StringComparison.Ordinal);
    }

    private static bool IsDentalFormulaBenefit(string benefit)
    {
        var normalized = NormalizeHealthFormulaText(benefit);
        return normalized.StartsWith("DENTAL", StringComparison.Ordinal) ||
               normalized.Contains(" DENTAL", StringComparison.Ordinal);
    }

    private static bool IsDentalFormulaTotal(string benefit)
    {
        return IsDentalFormulaBenefit(benefit) &&
               NormalizeHealthFormulaText(benefit).Contains("TOTAL", StringComparison.Ordinal);
    }

    private static string HealthFormulaAccountKey(RateLine line)
    {
        var label = HealthFormulaAccountLabel(line);
        return string.IsNullOrWhiteSpace(label) ? NormalizeHealthFormulaText(line.SourceReference) : NormalizeHealthFormulaText(label);
    }

    private static string HealthFormulaAccountLabel(RateLine line)
    {
        var value = !string.IsNullOrWhiteSpace(line.AccountClass) ? line.AccountClass : line.SourceReference;
        return value.Replace("ACCOUNT/CLASS:", "", StringComparison.OrdinalIgnoreCase).Trim();
    }

    private static string NormalizeHealthFormulaTier(string tierOrClass)
    {
        var normalized = NormalizeHealthFormulaText(tierOrClass);
        if (normalized.Contains("SINGLE PARENT", StringComparison.Ordinal))
        {
            return "SINGLE PARENT";
        }

        if (normalized.Contains("SINGLE", StringComparison.Ordinal))
        {
            return "SINGLE";
        }

        if (normalized.Contains("COUPLE", StringComparison.Ordinal))
        {
            return "COUPLE";
        }

        if (normalized.Contains("FAMILY", StringComparison.Ordinal))
        {
            return "FAMILY";
        }

        return string.IsNullOrWhiteSpace(normalized) ? "ALL" : normalized;
    }

    private static string NormalizeHealthFormulaText(string value)
    {
        return ParsingHelpers.Clean(value).ToUpperInvariant();
    }

    private static void ApplySectionHeader(IXLRange range)
    {
        range.Style.Font.Bold = true;
        range.Style.Font.FontColor = XLColor.White;
        range.Style.Fill.BackgroundColor = XLColor.FromHtml("#3399B3");
        range.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
        range.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
    }

    private static void ApplyDarkSectionHeader(IXLRange range)
    {
        range.Style.Font.Bold = true;
        range.Style.Font.FontColor = XLColor.White;
        range.Style.Fill.BackgroundColor = XLColor.FromHtml("#004650");
        range.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
        range.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
    }

    private static List<ExperienceLine> GetHealthFormulaExperienceRows(AnalysisResult analysis)
    {
        if (IsCanadaLifeAnalysis(analysis))
        {
            var reconciliation = GetCanadaLifeExperienceReconciliation(analysis);
            var selectedRows = reconciliation.DrugRows.Count > 0
                ? reconciliation.CombinedRows
                : reconciliation.HealthRows;
            return selectedRows
                .OrderBy(row => GetFormulaPeriodStart(row) ?? DateOnly.MinValue)
                .TakeLast(3)
                .ToList();
        }

        var rows = analysis.Pdf.ExperienceLines
            .Where(row => row.Benefit.Equals("Non Pooled Extended Health Care - Subtotal", StringComparison.OrdinalIgnoreCase))
            .Where(IsFormulaExperiencePeriod)
            .OrderBy(row => GetFormulaPeriodStart(row) ?? DateOnly.MinValue)
            .TakeLast(3)
            .ToList();

        if (rows.Count > 0)
        {
            return rows;
        }

        return analysis.Pdf.ExperienceLines
            .Where(row => row.Benefit.Contains("Extended Health", StringComparison.OrdinalIgnoreCase) ||
                           row.Benefit.Contains("Health Care", StringComparison.OrdinalIgnoreCase))
            .Where(IsFormulaExperiencePeriod)
            .OrderBy(row => GetFormulaPeriodStart(row) ?? DateOnly.MinValue)
            .TakeLast(3)
            .ToList();
    }

    private static List<ExperienceLine> GetDentalFormulaExperienceRows(AnalysisResult analysis)
    {
        return analysis.Pdf.ExperienceLines
            .Where(row => IsDentalFormulaBenefit(row.Benefit))
            .Where(IsFormulaExperiencePeriod)
            .OrderBy(row => GetFormulaPeriodStart(row) ?? DateOnly.MinValue)
            .TakeLast(3)
            .ToList();
    }

    private static void WriteCanadaLifeComponentAnalyses(XLWorkbook workbook, AnalysisResult analysis)
    {
        if (!IsCanadaLifeAnalysis(analysis) ||
            !workbook.TryGetWorksheet("HealthRenewal_FormulaAnalysis", out var healthAnalysis))
        {
            return;
        }

        var components = new[]
        {
            new CanadaLifeComponentSheet("CLhealthanalysis", "Canada Life Health Experience Analysis", GetCanadaLifeComponentRows(analysis, "Extended Health Care")),
            new CanadaLifeComponentSheet("CLdruganalysis", "Canada Life Drug Experience Analysis", GetCanadaLifeComponentRows(analysis, "Drugs")),
            new CanadaLifeComponentSheet("CLvisionanalysis", "Canada Life Vision Experience Analysis", GetCanadaLifeComponentRows(analysis, "Vision"))
        };

        foreach (var component in components.Where(item => item.Rows.Count > 0))
        {
            if (workbook.TryGetWorksheet(component.SheetName, out var existing))
            {
                existing.Delete();
            }

            var ws = healthAnalysis.CopyTo(component.SheetName);
            ws.Cell("A1").Value = component.Title;
            WriteCanadaLifeComponentExperience(ws, component.Rows);
        }
    }

    private static void WriteCanadaLifeComponentExperience(IXLWorksheet ws, IReadOnlyList<ExperienceLine> rows)
    {
        var dataColumns = new[] { 3, 4, 5 };
        var selectedRows = rows
            .Where(IsFormulaExperiencePeriod)
            .OrderBy(row => GetFormulaPeriodStart(row) ?? DateOnly.MinValue)
            .TakeLast(dataColumns.Length)
            .ToList();

        for (var index = 0; index < dataColumns.Length; index++)
        {
            var column = dataColumns[index];
            var source = GetRightAlignedExperienceRow(selectedRows, index, dataColumns.Length);
            var (from, to) = SplitPeriod(source?.Period);
            SetOptionalText(ws.Cell(5, column), from);
            SetOptionalText(ws.Cell(6, column), to);
            ws.Cell(8, column).Value = source?.BilledPremium ?? 0m;
            ws.Cell(9, column).Value = source?.PaidClaims ?? 0m;
            ws.Cell(10, column).Value = (source?.IncurredClaims ?? 0m) - (source?.PaidClaims ?? 0m);
            ws.Cell(11, column).FormulaA1 = $"=IFERROR({ColumnName(column)}9+{ColumnName(column)}10,0)";
            ws.Cell(12, column).FormulaA1 = $"=IFERROR({ColumnName(column)}11/{ColumnName(column)}8,0)";
            ws.Cell(20, column).Value = NormalizeFormulaPercent(source?.Weighting) ?? 0m;
        }

        ws.Cell("B26").Value = GetDefaultFormulaTrendInput(selectedRows, 0m);
        ws.Cell("B27").Value = GetDefaultFormulaCredibility(selectedRows, 0m);
        ws.Cell("B28").Value = GetDefaultFormulaTargetLossRatio(selectedRows, 0m);
        ws.Cell("I12").Value = GetDefaultFormulaCarrierRateChange(selectedRows, 0m);

        const int auditHeaderRow = 93;
        ws.Cell(auditHeaderRow, 1).Value = "Canada Life source-extracted component rows";
        ws.Range(auditHeaderRow, 1, auditHeaderRow, 6).Merge();
        ApplySectionHeader(ws.Range(auditHeaderRow, 1, auditHeaderRow, 6));
        var headers = new[] { "Period", "Premium", "Paid Claims", "Incurred Claims", "Incurred Loss Ratio", "Source" };
        for (var index = 0; index < headers.Length; index++)
        {
            ws.Cell(auditHeaderRow + 1, index + 1).Value = headers[index];
        }
        ApplyDarkSectionHeader(ws.Range(auditHeaderRow + 1, 1, auditHeaderRow + 1, headers.Length));

        var rowNumber = auditHeaderRow + 2;
        foreach (var source in rows.OrderBy(row => GetFormulaPeriodStart(row) ?? DateOnly.MinValue))
        {
            ws.Cell(rowNumber, 1).Value = source.Period;
            ws.Cell(rowNumber, 2).Value = source.BilledPremium ?? 0m;
            ws.Cell(rowNumber, 3).Value = source.PaidClaims ?? 0m;
            ws.Cell(rowNumber, 4).Value = source.IncurredClaims ?? 0m;
            ws.Cell(rowNumber, 5).FormulaA1 = $"=IFERROR(D{rowNumber}/B{rowNumber},0)";
            ws.Cell(rowNumber, 6).Value = source.SourceReference;
            rowNumber++;
        }

        if (rowNumber > auditHeaderRow + 2)
        {
            ws.Range(auditHeaderRow + 2, 2, rowNumber - 1, 4).Style.NumberFormat.Format = "$#,##0";
            ws.Range(auditHeaderRow + 2, 5, rowNumber - 1, 5).Style.NumberFormat.Format = "0.00%";
        }
    }

    private static void FormatCanadaLifeComponentSheet(IXLWorksheet ws)
    {
        ws.Column(1).Width = 32;
        ws.Column(2).Width = 18;
        ws.Columns(3, 5).Width = 18;
        ws.Column(6).Width = 30;
        ws.Column(7).Width = 12;
        ws.Column(8).Width = 27;
        ws.Columns(9, 10).Width = 15;
        ws.Column(11).Width = 12;
        ws.Range("A93:F94").Style.Alignment.WrapText = true;
    }

    private static bool IsCanadaLifeAnalysis(AnalysisResult analysis)
    {
        return analysis.Pdf.Carrier.Equals("Canada Life", StringComparison.OrdinalIgnoreCase);
    }

    private static List<ExperienceLine> GetCanadaLifeComponentRows(AnalysisResult analysis, string benefit)
    {
        return analysis.Pdf.ExperienceLines
            .Where(IsFormulaExperiencePeriod)
            .Where(row => row.Benefit.Equals(benefit, StringComparison.OrdinalIgnoreCase))
            .ToList();
    }

    private static CanadaLifeExperienceReconciliation GetCanadaLifeExperienceReconciliation(AnalysisResult analysis)
    {
        var healthRows = GetCanadaLifeComponentRows(analysis, "Extended Health Care");
        var drugRows = GetCanadaLifeComponentRows(analysis, "Drugs");
        var visionRows = GetCanadaLifeComponentRows(analysis, "Vision");
        var healthByPeriod = GroupCanadaLifeRowsByCompletePeriod(healthRows);
        var drugByPeriod = GroupCanadaLifeRowsByCompletePeriod(drugRows);
        var visionByPeriod = GroupCanadaLifeRowsByCompletePeriod(visionRows);
        var combinedRows = new List<ExperienceLine>();

        foreach (var period in healthByPeriod.Keys.Intersect(drugByPeriod.Keys).OrderBy(key => key.Start))
        {
            var health = healthByPeriod[period];
            var drugs = drugByPeriod[period];
            visionByPeriod.TryGetValue(period, out var vision);
            var included = health.Concat(drugs).Concat(vision ?? []).ToList();
            var premium = included.Sum(row => row.BilledPremium ?? 0m);
            var paidClaims = included.Sum(row => row.PaidClaims ?? 0m);
            var incurredClaims = included.Sum(row => row.IncurredClaims ?? 0m);
            var factorSource = health[0];

            combinedRows.Add(new ExperienceLine
            {
                Benefit = "Extended Health Care",
                Period = factorSource.Period,
                BilledPremium = premium,
                PaidClaims = paidClaims,
                IncurredClaims = incurredClaims,
                PaidLossRatio = premium == 0m ? null : decimal.Round(paidClaims / premium * 100m, 4),
                IncurredLossRatio = premium == 0m ? null : decimal.Round(incurredClaims / premium * 100m, 4),
                Weighting = factorSource.Weighting,
                ProposedAdjustment = factorSource.ProposedAdjustment,
                TrendFactor = factorSource.TrendFactor,
                Credibility = factorSource.Credibility,
                TargetLossRatio = factorSource.TargetLossRatio,
                SourceReference = $"Canada Life exact-period combined: {string.Join("; ", included.Select(row => row.SourceReference).Distinct(StringComparer.OrdinalIgnoreCase))}"
            });
        }

        var combinedPeriods = combinedRows
            .Select(GetFormulaPeriodKey)
            .Where(key => key.HasValue)
            .Select(key => key!.Value)
            .ToHashSet();
        var unmatchedHealth = healthByPeriod.Keys.Except(combinedPeriods).OrderBy(key => key.Start).ToList();
        var unmatchedDrugs = drugByPeriod.Keys.Except(combinedPeriods).OrderBy(key => key.Start).ToList();
        var unmatchedVision = visionByPeriod.Keys.Except(combinedPeriods).OrderBy(key => key.Start).ToList();

        return new CanadaLifeExperienceReconciliation(
            healthRows,
            drugRows,
            visionRows,
            combinedRows,
            unmatchedHealth,
            unmatchedDrugs,
            unmatchedVision);
    }

    private static Dictionary<FormulaPeriodKey, List<ExperienceLine>> GroupCanadaLifeRowsByCompletePeriod(IEnumerable<ExperienceLine> rows)
    {
        return rows
            .Select(row => new { Row = row, Key = GetFormulaPeriodKey(row) })
            .Where(item => item.Key.HasValue)
            .GroupBy(item => item.Key!.Value)
            .ToDictionary(group => group.Key, group => group.Select(item => item.Row).ToList());
    }

    private static FormulaPeriodKey? GetFormulaPeriodKey(ExperienceLine row)
    {
        var (from, to) = SplitPeriod(row.Period);
        var start = ParsingHelpers.ParseDateOnly(from);
        var end = ParsingHelpers.ParseDateOnly(to);
        if (!start.HasValue || !end.HasValue)
        {
            return null;
        }

        return new FormulaPeriodKey(start.Value, end.Value);
    }

    private static bool IsFormulaExperiencePeriod(ExperienceLine row)
    {
        return !string.IsNullOrWhiteSpace(row.Period) &&
               !row.Period.Contains("Weighted", StringComparison.OrdinalIgnoreCase) &&
               !row.SourceReference.Contains("Weighted", StringComparison.OrdinalIgnoreCase);
    }

    private static DateOnly? GetFormulaPeriodStart(ExperienceLine row)
    {
        var (from, _) = SplitPeriod(row.Period);
        var parsed = ParsingHelpers.ParseDateOnly(from);
        if (parsed.HasValue)
        {
            return parsed;
        }

        return DateOnly.TryParse(from, CultureInfo.InvariantCulture, DateTimeStyles.None, out var date)
            ? date
            : null;
    }

    private static decimal GetDefaultDentalTargetLossRatio(AnalysisResult analysis)
    {
        var extracted = analysis.Pdf.ExperienceLines
            .Where(row => IsDentalFormulaBenefit(row.Benefit))
            .Select(row => NormalizeFormulaPercent(row.Weighting))
            .FirstOrDefault(value => value.HasValue && value.Value > 0m);

        return extracted ?? 0.755m;
    }

    private static decimal? NormalizeFormulaPercent(decimal? value)
    {
        if (!value.HasValue)
        {
            return null;
        }

        return Math.Abs(value.Value) > 1m ? value.Value / 100m : value.Value;
    }

    private static ExperienceLine? GetRightAlignedExperienceRow(IReadOnlyList<ExperienceLine> rows, int columnIndex, int columnCount)
    {
        var rowCount = Math.Min(rows.Count, columnCount);
        var firstDataColumnIndex = columnCount - rowCount;
        if (columnIndex < firstDataColumnIndex)
        {
            return null;
        }

        var rowIndex = columnIndex - firstDataColumnIndex;
        return rowIndex >= 0 && rowIndex < rows.Count ? rows[rowIndex] : null;
    }

    private static IReadOnlyList<decimal> GetFormulaExperienceWeights(IReadOnlyList<ExperienceLine> rows, int columnCount)
    {
        var fallback = new decimal[columnCount];
        if (columnCount >= 2)
        {
            fallback[^2] = 0.33m;
            fallback[^1] = 0.67m;
        }

        var extracted = new decimal[columnCount];
        var total = 0m;
        for (var index = columnCount - 1; index >= 0; index--)
        {
            var source = GetRightAlignedExperienceRow(rows, index, columnCount);
            var weight = NormalizeFormulaPercent(source?.Weighting);
            if (!weight.HasValue || weight.Value <= 0m || weight.Value > 1m)
            {
                return fallback;
            }

            total += weight.Value;
            if (total > 1.005m)
            {
                return fallback;
            }

            extracted[index] = weight.Value;
            if (Math.Abs(total - 1m) <= 0.005m)
            {
                return extracted;
            }
        }

        return fallback;
    }

    private static void SetOptionalText(IXLCell cell, string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            cell.Clear(XLClearOptions.Contents);
            return;
        }

        cell.Value = value;
    }

    private static (string From, string To) SplitPeriod(string? period)
    {
        if (string.IsNullOrWhiteSpace(period))
        {
            return ("", "");
        }

        var normalized = ParsingHelpers.Clean(period);
        var dashParts = normalized.Split(" - ", StringSplitOptions.None);
        if (dashParts.Length >= 2)
        {
            return (dashParts[0], dashParts[1]);
        }

        var toIndex = normalized.IndexOf(" to ", StringComparison.OrdinalIgnoreCase);
        if (toIndex > 0)
        {
            return (
                normalized[..toIndex],
                normalized[(toIndex + " to ".Length)..]);
        }

        return (normalized, "");
    }

    private static string ColumnName(int column)
    {
        var dividend = column;
        var columnName = "";
        while (dividend > 0)
        {
            var modulo = (dividend - 1) % 26;
            columnName = Convert.ToChar('A' + modulo) + columnName;
            dividend = (dividend - modulo) / 26;
        }

        return columnName;
    }

    private static void WriteClaimsAnalysisServiceCategory(
        XLWorkbook workbook,
        string sheetName,
        IReadOnlyCollection<ServiceCategoryTrend> trends,
        string warningMessage,
        string firstColumnHeader = "Service Category",
        ClaimsTrendValidation? validation = null)
    {
        var ws = workbook.AddWorksheet(sheetName);
        var rows = trends.ToList();
        var periodLabels = GetServiceCategoryPeriodLabels(rows);
        if (periodLabels.Count == 0)
        {
            periodLabels = ["Prior Period", "Current Period"];
        }
        var lastColumn = periodLabels.Count + 1;

        ws.Cell(1, 1).Value = firstColumnHeader;
        for (var index = 0; index < periodLabels.Count; index++)
        {
            ws.Cell(1, index + 2).Value = periodLabels[index];
        }
        Header(ws.Range(1, 1, 1, lastColumn));

        if (rows.Count == 0)
        {
            ws.Cell("A3").Value = warningMessage;
            ws.Range(3, 1, 3, lastColumn).Merge();
            ws.Cell("A3").Style.Fill.BackgroundColor = WarningFill;
            ws.Columns(1, lastColumn).AdjustToContents(12, 36);
            return;
        }

        for (var index = 0; index < rows.Count; index++)
        {
            var rowNumber = index + 2;
            ws.Cell(rowNumber, 1).Value = rows[index].Category;
            for (var periodIndex = 0; periodIndex < periodLabels.Count; periodIndex++)
            {
                ws.Cell(rowNumber, periodIndex + 2).Value = GetServiceCategoryPeriodAmount(rows[index], periodLabels[periodIndex]);
            }
        }

        var lastRow = rows.Count + 1;
        ws.Range(2, 2, lastRow, lastColumn).Style.NumberFormat.Format = "$#,##0";
        ws.Range(2, 2, lastRow, lastColumn).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
        ws.Range(2, 1, lastRow, lastColumn).Style.Font.FontColor = XLColor.FromHtml("#004650");
        ws.Range(1, 1, lastRow, lastColumn).Style.Border.InsideBorderColor = XLColor.FromHtml("#8DD3C7");
        ws.Range(1, 1, lastRow, lastColumn).Style.Border.OutsideBorderColor = XLColor.FromHtml("#8DD3C7");
        ws.Column(1).Width = 30;
        ws.Columns(2, lastColumn).Width = 18;

        if (validation is not null)
        {
            var validationStart = lastRow + 3;
            ws.Cell(validationStart, 1).Value = "Validation";
            Header(ws.Range(validationStart, 1, validationStart, lastColumn));

            var validationRows = new (string Label, string Value)[]
            {
                ("Client Name", validation.ClientName),
                ("Policy Number", validation.PolicyNumber),
                ("Prior Period", validation.PriorPeriod),
                ("Current Period", validation.CurrentPeriod),
                ("Source sheet used", validation.SourceSheetUsed),
                ("Status", validation.Status),
                ("Warning", validation.Warning)
            };

            for (var index = 0; index < validationRows.Length; index++)
            {
                var rowNumber = validationStart + index + 1;
                ws.Cell(rowNumber, 1).Value = validationRows[index].Label;
                ws.Cell(rowNumber, 2).Value = validationRows[index].Value;
                ws.Range(rowNumber, 2, rowNumber, lastColumn).Merge();
            }

            ws.Range(validationStart + 1, 1, validationStart + validationRows.Length, 1).Style.Font.Bold = true;
            ws.Range(validationStart, 1, validationStart + validationRows.Length, lastColumn).Style.Border.InsideBorderColor = XLColor.FromHtml("#8DD3C7");
            ws.Range(validationStart, 1, validationStart + validationRows.Length, lastColumn).Style.Border.OutsideBorderColor = XLColor.FromHtml("#8DD3C7");
            ws.Range(validationStart + 1, 2, validationStart + validationRows.Length, lastColumn).Style.Alignment.WrapText = true;
        }
    }

    private static List<string> GetServiceCategoryPeriodLabels(IEnumerable<ServiceCategoryTrend> rows, int maximum = 3)
    {
        return rows
            .SelectMany(row => ServiceCategoryHistory.GetPeriods(row, int.MaxValue))
            .Where(period => !string.IsNullOrWhiteSpace(period.PeriodLabel))
            .GroupBy(period => period.PeriodLabel, StringComparer.OrdinalIgnoreCase)
            .Select(group => group.First())
            .OrderBy(period => ServiceCategoryHistory.ParsePeriodStart(period.PeriodLabel) ?? DateOnly.MinValue)
            .ThenBy(period => period.PeriodLabel, StringComparer.OrdinalIgnoreCase)
            .TakeLast(Math.Max(1, maximum))
            .Select(period => period.PeriodLabel)
            .ToList();
    }

    private static decimal GetServiceCategoryPeriodAmount(ServiceCategoryTrend row, string periodLabel)
    {
        return ServiceCategoryHistory.GetPeriods(row, int.MaxValue)
            .FirstOrDefault(period => period.PeriodLabel.Equals(periodLabel, StringComparison.OrdinalIgnoreCase))?.Amount ?? 0m;
    }

    private static List<int> WriteGroupedRateRows(IXLWorksheet ws, List<RateLine> rateLines)
    {
        var totalRows = new List<int>();
        var row = 2;
        foreach (var sourceGroup in rateLines.GroupBy(line => line.SourceReference))
        {
            var groupStart = row;
            var additiveRows = new List<int>();
            var previousAccountClass = "";
            foreach (var line in sourceGroup)
            {
                if (!string.IsNullOrWhiteSpace(line.AccountClass) &&
                    !line.AccountClass.Equals(previousAccountClass, StringComparison.OrdinalIgnoreCase))
                {
                    WriteRateAccountClassHeading(ws, row, line.AccountClass);
                    previousAccountClass = line.AccountClass;
                    row++;
                }

                WriteRateLineRow(ws, row, line);
                if (line.IncludeInPremiumTotal)
                {
                    additiveRows.Add(row);
                }
                if (line.Benefit.StartsWith("Extended Health Care", StringComparison.OrdinalIgnoreCase))
                {
                    ws.Row(row).Style.Font.FontColor = CheckGreen;
                }

                row++;
            }

            if (row > groupStart)
            {
                ws.Cell(row, 1).Value = "Totals";
                ws.Cell(row, 6).FormulaA1 = RateTotalFormula(additiveRows, "F");
                ws.Cell(row, 9).FormulaA1 = RateTotalFormula(additiveRows, "I");
                ws.Cell(row, 16).FormulaA1 = RateTotalFormula(additiveRows, "P");
                ws.Cell(row, 17).FormulaA1 = $"IFERROR(IF(F{row}=0,\"\",P{row}/F{row}-1),\"\")";
                ws.Row(row).Style.Font.Bold = true;
                ws.Row(row).Style.Font.FontColor = CheckGreen;
                totalRows.Add(row);
                row += 2;
            }
        }

        return totalRows;
    }

    private static void WriteRateAccountClassHeading(IXLWorksheet ws, int row, string accountClass)
    {
        ws.Cell(row, 1).Value = accountClass;
        ws.Range(row, 1, row, 17).Merge();
        ws.Range(row, 1, row, 17).Style.Font.Bold = true;
        ws.Range(row, 1, row, 17).Style.Fill.BackgroundColor = AccountFill;
        ws.Range(row, 1, row, 17).Style.Border.TopBorder = XLBorderStyleValues.Thin;
        ws.Range(row, 1, row, 17).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
    }

    private static void WriteRateSourceTotalCheck(
        IXLWorksheet ws,
        IReadOnlyCollection<int> rateTotalRows,
        AnalysisResult analysis)
    {
        if (rateTotalRows.Count == 0)
        {
            return;
        }

        var titleRow = (ws.LastRowUsed()?.RowNumber() ?? 1) + 2;
        var headerRow = titleRow + 1;
        var reconciliationTotalRow = titleRow + 2;
        var sourceTotalRow = titleRow + 3;
        var varianceRow = titleRow + 4;
        var statusRow = titleRow + 5;

        ws.Cell(titleRow, 1).Value = "Source Total Check";
        ws.Range(titleRow, 1, titleRow, 17).Merge();
        ws.Range(titleRow, 1, titleRow, 17).Style.Fill.BackgroundColor = WestlandTeal;
        ws.Range(titleRow, 1, titleRow, 17).Style.Font.FontColor = XLColor.White;
        ws.Range(titleRow, 1, titleRow, 17).Style.Font.Bold = true;

        ws.Cell(headerRow, 1).Value = "Check";
        ws.Cell(headerRow, 6).Value = "Current Premium";
        ws.Cell(headerRow, 9).Value = "Renewal Premium";
        ws.Cell(headerRow, 11).Value = "Tolerance";
        ws.Cell(headerRow, 12).Value = "$1.00";
        Header(ws.Range(headerRow, 1, headerRow, 12));

        ws.Cell(reconciliationTotalRow, 1).Value = "Rate Reconciliation total";
        ws.Cell(reconciliationTotalRow, 6).FormulaA1 = RateTotalFormula(rateTotalRows, "F");
        ws.Cell(reconciliationTotalRow, 9).FormulaA1 = RateTotalFormula(rateTotalRows, "I");

        ws.Cell(sourceTotalRow, 1).Value = "Source total (PDF / workbook)";
        SetMoney(ws.Cell(sourceTotalRow, 6), analysis.Pdf.CurrentMonthlyPremium);
        SetMoney(ws.Cell(sourceTotalRow, 9), analysis.Pdf.RenewalMonthlyPremium);

        ws.Cell(varianceRow, 1).Value = "Variance to source total";
        ws.Cell(varianceRow, 6).FormulaA1 = $"IF(F{sourceTotalRow}=\"\",\"\",F{reconciliationTotalRow}-F{sourceTotalRow})";
        ws.Cell(varianceRow, 9).FormulaA1 = $"IF(I{sourceTotalRow}=\"\",\"\",I{reconciliationTotalRow}-I{sourceTotalRow})";

        ws.Cell(statusRow, 1).Value = "Status";
        ws.Cell(statusRow, 6).FormulaA1 = $"IF(F{sourceTotalRow}=\"\",\"Source total missing\",IF(ABS(F{varianceRow})<=1,\"Match\",\"Review\"))";
        ws.Cell(statusRow, 9).FormulaA1 = $"IF(I{sourceTotalRow}=\"\",\"Source total missing\",IF(ABS(I{varianceRow})<=1,\"Match\",\"Review\"))";

        ws.Range(reconciliationTotalRow, 1, statusRow, 17).Style.Fill.BackgroundColor = SoftFill;
        ws.Range(reconciliationTotalRow, 1, statusRow, 17).Style.Font.Bold = true;
        ws.Range(reconciliationTotalRow, 6, varianceRow, 9).Style.NumberFormat.Format = "$#,##0.00";
        ws.Range(headerRow, 1, statusRow, 17).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        ws.Range(headerRow, 1, statusRow, 17).Style.Border.InsideBorder = XLBorderStyleValues.Hair;
    }

    private static string RateTotalFormula(IEnumerable<int> rateTotalRows, string columnLetter)
    {
        return $"SUM({string.Join(",", rateTotalRows.Select(row => $"{columnLetter}{row}"))})";
    }

    private static void WriteRateLineRow(IXLWorksheet ws, int row, RateLine line)
    {
        SetValue(ws.Cell(row, 1), line.Benefit);
        SetValue(ws.Cell(row, 2), line.TierOrClass);
        SetValue(ws.Cell(row, 3), line.Volume);
        SetValue(ws.Cell(row, 4), line.Lives);
        SetValue(ws.Cell(row, 5), line.CurrentRate);
        WriteRatePremium(ws.Cell(row, 6), row, "E", "C", "D", line, line.CurrentRate, line.CurrentPremium);
        SetValue(ws.Cell(row, 7), line.CurrentVariance);
        SetValue(ws.Cell(row, 8), line.RenewalRate);
        var fallbackRateFormat = UsesThreeDecimalRateDisplay(line.Benefit) ? "0.000" : "0.00";
        ws.Cell(row, 5).Style.NumberFormat.Format = RateFormat(line.CurrentRatePrecision, fallbackRateFormat);
        ws.Cell(row, 8).Style.NumberFormat.Format = RateFormat(line.RenewalRatePrecision, fallbackRateFormat);
        WriteRatePremium(ws.Cell(row, 9), row, "H", "C", "D", line, line.RenewalRate, line.RenewalPremium);
        SetValue(ws.Cell(row, 10), line.RenewalVariance);
        ws.Cell(row, 11).FormulaA1 = $"IFERROR(IF(F{row}=0,\"\",I{row}/F{row}-1),\"\")";
        SetValue(ws.Cell(row, 12), RateSourceDisplay(line));
        ws.Cell(row, 13).FormulaA1 = $"IFERROR(IF(C{row}=\"\",\"\",N(C{row})),\"\")";
        ws.Cell(row, 14).FormulaA1 = $"IFERROR(IF(D{row}=\"\",\"\",N(D{row})),\"\")";
        if (line.NegotiatedRate.HasValue)
        {
            SetValue(ws.Cell(row, 15), line.NegotiatedRate);
        }
        else
        {
            ws.Cell(row, 15).FormulaA1 = $"IFERROR(IF(H{row}=\"\",\"\",N(H{row})),\"\")";
        }
        if (CanCalculateRatePremium(line, line.NegotiatedRate ?? line.RenewalRate))
        {
            ws.Cell(row, 16).FormulaA1 = RatePremiumFormula(row, "O", "M", "N", line.Benefit, line.RateBasis);
        }
        else
        {
            SetValue(ws.Cell(row, 16), line.NegotiatedPremium ?? line.RenewalPremium);
        }
        ws.Cell(row, 17).FormulaA1 = $"IFERROR(IF(F{row}=0,\"\",P{row}/F{row}-1),\"\")";
    }

    private static void WriteRatePremium(
        IXLCell cell,
        int row,
        string rateColumn,
        string volumeColumn,
        string livesColumn,
        RateLine line,
        decimal? rate,
        decimal? fallbackPremium)
    {
        if (CanCalculateRatePremium(line, rate))
        {
            cell.FormulaA1 = RatePremiumFormula(row, rateColumn, volumeColumn, livesColumn, line.Benefit, line.RateBasis);
            return;
        }

        SetValue(cell, fallbackPremium);
    }

    private static bool CanCalculateRatePremium(RateLine line, decimal? rate)
    {
        if (!rate.HasValue || line.ValidationStatus.Equals(RateLineValidationStatuses.NeedsReview, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        if (UnderwritingEngine.IsPerEmployeeRateBasis(line.RateBasis))
        {
            return line.Lives.HasValue;
        }

        return RateVolumeDivisor(line.Benefit, line.RateBasis).HasValue
            ? line.Volume.HasValue
            : line.Lives.HasValue || line.Volume.HasValue;
    }

    private static string RateFormat(int? printedPrecision, string fallback)
    {
        return printedPrecision.HasValue
            ? printedPrecision.Value == 0 ? "0" : $"0.{new string('0', printedPrecision.Value)}"
            : fallback;
    }

    private static string RateSourceDisplay(RateLine line)
    {
        if (!line.ValidationStatus.Equals(RateLineValidationStatuses.NeedsReview, StringComparison.OrdinalIgnoreCase))
        {
            return line.SourceReference;
        }

        return string.IsNullOrWhiteSpace(line.ValidationMessage)
            ? $"{line.SourceReference} | Needs Review"
            : $"{line.SourceReference} | Needs Review: {line.ValidationMessage}";
    }

    private static bool UsesThreeDecimalRateDisplay(string benefit)
    {
        var normalized = NormalizeLineRateText(benefit);
        return (normalized.Contains("LIFE", StringComparison.Ordinal) &&
                !normalized.Contains("CRITICAL ILLNESS", StringComparison.Ordinal)) ||
               IsFormattedAdAndD(normalized) ||
               normalized.Equals("LTD", StringComparison.Ordinal) ||
               normalized.Contains("LONG TERM DISABILITY", StringComparison.Ordinal);
    }

    private static string RatePremiumFormula(
        int row,
        string rateColumn,
        string volumeColumn,
        string livesColumn,
        string benefit,
        string rateBasis)
    {
        var volumeReference = NumericQuantityReference(volumeColumn, row);
        var livesReference = NumericQuantityReference(livesColumn, row);
        if (UnderwritingEngine.IsPerEmployeeRateBasis(rateBasis))
        {
            return $"IFERROR(IF({rateColumn}{row}=\"\",\"\",IF({livesColumn}{row}=\"\",\"\",ROUND({livesReference}*{rateColumn}{row},2))),\"\")";
        }

        var basis = RateVolumeDivisor(benefit, rateBasis);
        return basis.HasValue
            ? $"IFERROR(IF({rateColumn}{row}=\"\",\"\",IF({volumeColumn}{row}=\"\",\"\",ROUND({volumeReference}*{rateColumn}{row}/{basis.Value},2))),\"\")"
            : $"IFERROR(IF({rateColumn}{row}=\"\",\"\",IF({livesColumn}{row}<>\"\",ROUND({livesReference}*{rateColumn}{row},2),IF({volumeColumn}{row}<>\"\",ROUND({volumeReference}*{rateColumn}{row},2),\"\"))),\"\")";
    }

    private static string NumericQuantityReference(string column, int row)
    {
        var reference = $"{column}{row}";
        if (column.Equals("M", StringComparison.OrdinalIgnoreCase))
        {
            return $"IF(N({reference})=0,C{row},N({reference}))";
        }

        if (column.Equals("N", StringComparison.OrdinalIgnoreCase))
        {
            return $"IF(N({reference})=0,D{row},N({reference}))";
        }

        return reference;
    }

    private static void WriteRateTotals(IXLWorksheet ws, List<RateLine> rateLines, AnalysisResult analysis)
    {
        if (rateLines.Count == 0)
        {
            return;
        }

        var lastDataRow = rateLines.Count + 1;
        var totalRow = lastDataRow + 3;
        var pdfRow = totalRow + 1;
        var varianceRow = totalRow + 2;

        ws.Cell(totalRow, 1).Value = "Extracted benefit total";
        ws.Cell(totalRow, 6).FormulaA1 = $"SUM(F2:F{lastDataRow})";
        ws.Cell(totalRow, 10).FormulaA1 = $"SUM(J2:J{lastDataRow})";

        ws.Cell(pdfRow, 1).Value = "PDF total monthly premium";
        SetMoney(ws.Cell(pdfRow, 6), analysis.Pdf.CurrentMonthlyPremium);
        SetMoney(ws.Cell(pdfRow, 10), analysis.Pdf.RenewalMonthlyPremium);

        ws.Cell(varianceRow, 1).Value = "Variance to PDF";
        ws.Cell(varianceRow, 6).FormulaA1 = $"F{totalRow}-F{pdfRow}";
        ws.Cell(varianceRow, 10).FormulaA1 = $"J{totalRow}-J{pdfRow}";

        var ehcStart = varianceRow + 2;
        WritePremiumCheckRows(
            ws,
            ehcStart,
            "Extracted EHC detail total",
            rateLines.Where(line => line.Benefit.StartsWith("Extended Health Care", StringComparison.OrdinalIgnoreCase) &&
                                    !line.Benefit.Equals("Extended Health Care - Total", StringComparison.OrdinalIgnoreCase)),
            analysis.Pdf.CurrentEhcPremium,
            analysis.Pdf.RenewalEhcPremium,
            "PDF total extended health care");

        var dentalStart = ehcStart + 4;
        WritePremiumCheckRows(
            ws,
            dentalStart,
            "Extracted dental detail total",
            rateLines.Where(line => line.Benefit.Equals("Dental", StringComparison.OrdinalIgnoreCase)),
            analysis.Pdf.CurrentDentalPremium,
            analysis.Pdf.RenewalDentalPremium,
            "PDF total dental");

        foreach (var row in new[] { totalRow, pdfRow, varianceRow, ehcStart, ehcStart + 1, ehcStart + 2, dentalStart, dentalStart + 1, dentalStart + 2 })
        {
            ws.Row(row).Style.Font.Bold = true;
        }

        ws.Range(totalRow, 1, varianceRow, 14).Style.Fill.BackgroundColor = SoftFill;
        ws.Range(ehcStart, 1, ehcStart + 2, 14).Style.Fill.BackgroundColor = SoftFill;
        ws.Range(dentalStart, 1, dentalStart + 2, 14).Style.Fill.BackgroundColor = SoftFill;
    }

    private static void WriteDemographicAnalysis(
        XLWorkbook workbook,
        string sheetName,
        string title,
        IReadOnlyCollection<DemographicLine> demographics,
        string[] ageBands,
        string previousLabel,
        Func<string?, string>? ageBandNormalizer = null,
        string? underwritingRiskFormula = null)
    {
        var ws = workbook.AddWorksheet(sheetName);
        Title(ws, "A1", title);
        ws.Range("A1:M1").Merge();

        ws.Cell("B3").Value = "Current";
        ws.Range("B3:E3").Merge();
        ws.Cell("F3").Value = previousLabel;
        ws.Range("F3:I3").Merge();
        ws.Cell("J3").Value = "Total";
        ws.Range("J3:M3").Merge();
        Header(ws.Range("B3:M3"));
        ws.Range("B3:M3").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

        var headers = new[]
        {
            "Age",
            "# Female Lives",
            "Current Female Volume",
            "# Male Lives",
            "Current Male Volume",
            "# Female Lives",
            $"{previousLabel} Female Volume",
            "# Male Lives",
            $"{previousLabel} Male Volume",
            "Current Lives Combined",
            $"{previousLabel} Lives Combined",
            "Last Renewal Total Volume",
            "Current Renewal Total Volume"
        };

        for (var index = 0; index < headers.Length; index++)
        {
            ws.Cell(4, index + 1).Value = headers[index];
        }

        Header(ws.Range("A4:M4"));

        var rowsByAge = BuildDemographicsByAge(demographics, ageBandNormalizer);
        var firstDataRow = 5;
        for (var index = 0; index < ageBands.Length; index++)
        {
            var rowNumber = firstDataRow + index;
            var ageBand = ageBands[index];
            ws.Cell(rowNumber, 1).Value = ageBand;
            ws.Cell(rowNumber, 14).Value = GetAgeMidpoint(ageBand);

            var lookupAgeBand = ageBandNormalizer?.Invoke(ageBand) ?? ageBand;
            if (rowsByAge.TryGetValue(lookupAgeBand, out var demographic))
            {
                SetValue(ws.Cell(rowNumber, 2), demographic.CurrentFemaleLives);
                SetValue(ws.Cell(rowNumber, 3), demographic.CurrentFemaleVolume);
                SetValue(ws.Cell(rowNumber, 4), demographic.CurrentMaleLives);
                SetValue(ws.Cell(rowNumber, 5), demographic.CurrentMaleVolume);
                SetValue(ws.Cell(rowNumber, 6), demographic.PriorFemaleLives);
                SetValue(ws.Cell(rowNumber, 7), demographic.PriorFemaleVolume);
                SetValue(ws.Cell(rowNumber, 8), demographic.PriorMaleLives);
                SetValue(ws.Cell(rowNumber, 9), demographic.PriorMaleVolume);
            }

            ws.Cell(rowNumber, 10).FormulaA1 = $"IFERROR(N(B{rowNumber})+N(D{rowNumber}),0)";
            ws.Cell(rowNumber, 11).FormulaA1 = $"IFERROR(N(F{rowNumber})+N(H{rowNumber}),0)";
            ws.Cell(rowNumber, 12).FormulaA1 = $"IFERROR(N(G{rowNumber})+N(I{rowNumber}),0)";
            ws.Cell(rowNumber, 13).FormulaA1 = $"IFERROR(N(C{rowNumber})+N(E{rowNumber}),0)";
        }

        var lastDataRow = firstDataRow + ageBands.Length - 1;
        ws.Range(firstDataRow, 2, lastDataRow, 5).Style.Fill.BackgroundColor = XLColor.FromHtml("#E8EFCF");
        ws.Range(firstDataRow, 6, lastDataRow, 9).Style.Fill.BackgroundColor = XLColor.FromHtml("#D9EEF7");
        var totalRow = lastDataRow + 3;
        ws.Cell(totalRow, 1).Value = "Total";
        foreach (var column in Enumerable.Range(2, 12))
        {
            ws.Cell(totalRow, column).FormulaA1 = $"IFERROR(SUM({ws.Cell(firstDataRow, column).Address}:{ws.Cell(lastDataRow, column).Address}),0)";
        }

        ws.Range(totalRow, 1, totalRow, 13).Style.Font.Bold = true;
        ws.Range(totalRow, 1, totalRow, 13).Style.Fill.BackgroundColor = SoftFill;

        if (demographics.Count == 0)
        {
            ws.Cell(totalRow + 1, 1).Value = "No demographic age/volume distribution was extracted from the PDF for this benefit.";
            ws.Range(totalRow + 1, 1, totalRow + 1, 13).Merge();
            ws.Cell(totalRow + 1, 1).Style.Fill.BackgroundColor = WarningFill;
        }

        var metricStart = totalRow + 3;
        ws.Cell(metricStart, 1).Value = "Data";
        Header(ws.Range(metricStart, 1, metricStart, 2));
        WriteDemographicMetricRows(ws, metricStart + 1, totalRow, firstDataRow, lastDataRow, underwritingRiskFormula);

        var weightedStart = metricStart + 13;
        WriteWeightedAgeBlock(ws, weightedStart, firstDataRow, lastDataRow);

        foreach (var column in new[] { 2, 4, 6, 8, 10, 11 })
        {
            ws.Range(firstDataRow, column, totalRow, column).Style.NumberFormat.Format = "#,##0";
        }

        foreach (var column in new[] { 3, 5, 7, 9, 12, 13 })
        {
            ws.Range(firstDataRow, column, totalRow, column).Style.NumberFormat.Format = "$#,##0";
        }

        ws.Column(14).Hide();

        ws.Column(1).Width = 18;
        ws.Columns(2, 13).Width = 15;
        ws.Range("A4:M4").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        ws.Range(firstDataRow, 1, totalRow, 13).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
        ws.Range(firstDataRow, 1, totalRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
    }

    private static void FormatDemographicAnalysisSheet(XLWorkbook workbook, string sheetName)
    {
        if (!workbook.TryGetWorksheet(sheetName, out var ws))
        {
            return;
        }

        ws.Columns(1, 13).Width = 10;
        ws.Columns(1, 13).Style.Alignment.WrapText = true;
        ws.Row(4).Height = 60;
        ws.Row(4).Style.Alignment.WrapText = true;
        ws.Row(4).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
    }

    private static void FormatUnderwritingReferenceMasterSheet(XLWorkbook workbook, string sheetName)
    {
        if (!workbook.TryGetWorksheet(sheetName, out var ws))
        {
            return;
        }

        ws.Column(1).Width = 30;
        foreach (var column in new[] { 2, 4, 6, 8, 12 })
        {
            ws.Column(column).Width = 12;
        }
        foreach (var column in new[] { 3, 5, 7, 9, 10, 11 })
        {
            ws.Column(column).Width = 16;
        }
        ws.Rows(1, 4).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        ws.Rows(3, 4).Style.Alignment.WrapText = true;
        ws.Row(1).Height = 24;
        ws.Row(3).Height = 24;
        ws.Row(4).Height = 36;
        ws.SheetView.FreezeRows(4);
    }

    private static Dictionary<string, DemographicLine> BuildDemographicsByAge(
        IReadOnlyCollection<DemographicLine> demographics,
        Func<string?, string>? ageBandNormalizer)
    {
        var rowsByAge = new Dictionary<string, DemographicLine>(StringComparer.OrdinalIgnoreCase);
        foreach (var row in demographics)
        {
            var ageBand = ageBandNormalizer?.Invoke(row.AgeBand) ?? row.AgeBand;
            if (string.IsNullOrWhiteSpace(ageBand))
            {
                continue;
            }

            if (!rowsByAge.TryGetValue(ageBand, out var existing))
            {
                rowsByAge[ageBand] = new DemographicLine
                {
                    Benefit = row.Benefit,
                    AgeBand = ageBand,
                    CurrentFemaleLives = row.CurrentFemaleLives,
                    CurrentFemaleVolume = row.CurrentFemaleVolume,
                    CurrentMaleLives = row.CurrentMaleLives,
                    CurrentMaleVolume = row.CurrentMaleVolume,
                    PriorFemaleLives = row.PriorFemaleLives,
                    PriorFemaleVolume = row.PriorFemaleVolume,
                    PriorMaleLives = row.PriorMaleLives,
                    PriorMaleVolume = row.PriorMaleVolume,
                    SourceReference = row.SourceReference
                };
                continue;
            }

            existing.CurrentFemaleLives = AddNullable(existing.CurrentFemaleLives, row.CurrentFemaleLives);
            existing.CurrentFemaleVolume = AddNullable(existing.CurrentFemaleVolume, row.CurrentFemaleVolume);
            existing.CurrentMaleLives = AddNullable(existing.CurrentMaleLives, row.CurrentMaleLives);
            existing.CurrentMaleVolume = AddNullable(existing.CurrentMaleVolume, row.CurrentMaleVolume);
            existing.PriorFemaleLives = AddNullable(existing.PriorFemaleLives, row.PriorFemaleLives);
            existing.PriorFemaleVolume = AddNullable(existing.PriorFemaleVolume, row.PriorFemaleVolume);
            existing.PriorMaleLives = AddNullable(existing.PriorMaleLives, row.PriorMaleLives);
            existing.PriorMaleVolume = AddNullable(existing.PriorMaleVolume, row.PriorMaleVolume);
            if (!string.IsNullOrWhiteSpace(row.SourceReference) &&
                !existing.SourceReference.Contains(row.SourceReference, StringComparison.OrdinalIgnoreCase))
            {
                existing.SourceReference = string.IsNullOrWhiteSpace(existing.SourceReference)
                    ? row.SourceReference
                    : $"{existing.SourceReference}; {row.SourceReference}";
            }
        }

        return rowsByAge;
    }

    private static int? AddNullable(int? current, int? next)
    {
        return current.HasValue || next.HasValue ? (current ?? 0) + (next ?? 0) : null;
    }

    private static decimal? AddNullable(decimal? current, decimal? next)
    {
        return current.HasValue || next.HasValue ? (current ?? 0m) + (next ?? 0m) : null;
    }

    private static string[] GetLifeAgeBands(IReadOnlyCollection<DemographicLine> demographics)
    {
        return DemographicAgeBandNormalizer.GetLifeAgeBands(demographics);
    }

    private static string[] GetLtdAgeBands(IReadOnlyCollection<DemographicLine> demographics)
    {
        return DemographicAgeBandNormalizer.GetLtdAgeBands(demographics);
    }

    private static void WriteDemographicMetricRows(
        IXLWorksheet ws,
        int startRow,
        int totalRow,
        int firstDataRow,
        int lastDataRow,
        string? underwritingRiskFormula)
    {
        var metrics = new (string Label, string Formula, string Format)[]
        {
            ("Change in EE", $"IFERROR(J{totalRow}-K{totalRow},0)", "#,##0"),
            ("% Change in EE", $"IFERROR(IF(K{totalRow}=0,\"\",(J{totalRow}-K{totalRow})/K{totalRow}),\"\")", "0.0%"),
            ("Current Males %", $"IFERROR(IF(J{totalRow}=0,\"\",D{totalRow}/J{totalRow}),\"\")", "0.0%"),
            ("Current Females %", $"IFERROR(IF(J{totalRow}=0,\"\",B{totalRow}/J{totalRow}),\"\")", "0.0%"),
            ("Current Avg Age", $"IFERROR(IF(J{totalRow}=0,\"\",SUMPRODUCT(J{firstDataRow}:J{lastDataRow},N{firstDataRow}:N{lastDataRow})/J{totalRow}),\"\")", "0.0"),
            ("Previous Males %", $"IFERROR(IF(K{totalRow}=0,\"\",H{totalRow}/K{totalRow}),\"\")", "0.0%"),
            ("Previous Females %", $"IFERROR(IF(K{totalRow}=0,\"\",F{totalRow}/K{totalRow}),\"\")", "0.0%"),
            ("Previous Avg Age", $"IFERROR(IF(K{totalRow}=0,\"\",SUMPRODUCT(K{firstDataRow}:K{lastDataRow},N{firstDataRow}:N{lastDataRow})/K{totalRow}),\"\")", "0.0"),
            ("Change in risk",
                string.IsNullOrWhiteSpace(underwritingRiskFormula)
                    ? $"IFERROR(IF(OR(B{startRow + 4}=\"\",B{startRow + 7}=\"\"),\"\",B{startRow + 4}-B{startRow + 7}),\"\")"
                    : $"IFERROR({underwritingRiskFormula},\"\")",
                string.IsNullOrWhiteSpace(underwritingRiskFormula) ? "0.0" : "0.0%")
        };

        for (var index = 0; index < metrics.Length; index++)
        {
            var row = startRow + index;
            ws.Cell(row, 1).Value = metrics[index].Label;
            ws.Cell(row, 2).FormulaA1 = metrics[index].Formula;
            ws.Cell(row, 2).Style.NumberFormat.Format = metrics[index].Format;
        }

        ws.Range(startRow, 1, startRow + metrics.Length - 1, 1).Style.Font.Bold = true;
        ws.Range(startRow, 1, startRow + metrics.Length - 1, 2).Style.Fill.BackgroundColor = SoftFill;
    }

    private static void WriteWeightedAgeBlock(IXLWorksheet ws, int startRow, int firstDataRow, int lastDataRow)
    {
        ws.Cell(startRow, 1).Value = "Weighted Age Analysis";
        ws.Range(startRow, 1, startRow, 5).Merge();
        Header(ws.Range(startRow, 1, startRow, 5));

        var headers = new[] { "Age", "Current Lives", "Current Weighted Age", "Previous Lives", "Previous Weighted Age" };
        for (var index = 0; index < headers.Length; index++)
        {
            ws.Cell(startRow + 2, index + 1).Value = headers[index];
        }

        Header(ws.Range(startRow + 2, 1, startRow + 2, headers.Length));

        for (var row = firstDataRow; row <= lastDataRow; row++)
        {
            var targetRow = startRow + 3 + row - firstDataRow;
            ws.Cell(targetRow, 1).FormulaA1 = $"A{row}";
            ws.Cell(targetRow, 2).FormulaA1 = $"IFERROR(N(J{row}),0)";
            ws.Cell(targetRow, 3).FormulaA1 = $"IFERROR(N(J{row})*N{row},0)";
            ws.Cell(targetRow, 4).FormulaA1 = $"IFERROR(N(K{row}),0)";
            ws.Cell(targetRow, 5).FormulaA1 = $"IFERROR(N(K{row})*N{row},0)";
        }

        var firstWeightedDataRow = startRow + 3;
        var lastWeightedDataRow = firstWeightedDataRow + lastDataRow - firstDataRow;
        var totalRow = lastWeightedDataRow + 1;
        ws.Cell(totalRow, 1).Value = "Total";
        foreach (var column in Enumerable.Range(2, 4))
        {
            ws.Cell(totalRow, column).FormulaA1 = $"IFERROR(SUM({ws.Cell(firstWeightedDataRow, column).Address}:{ws.Cell(lastWeightedDataRow, column).Address}),0)";
        }

        ws.Range(totalRow, 1, totalRow, 5).Style.Font.Bold = true;
        ws.Range(totalRow, 1, totalRow, 5).Style.Fill.BackgroundColor = SoftFill;
        ws.Range(firstWeightedDataRow, 2, totalRow, 5).Style.NumberFormat.Format = "#,##0.0";
    }

    private static decimal GetAgeMidpoint(string ageBand)
    {
        return ageBand switch
        {
            "Under 25" => 22m,
            "25-29" => 27m,
            "30-34" => 32m,
            "35-39" => 37m,
            "40-44" => 42m,
            "45-49" => 47m,
            "50-54" => 52m,
            "55-59" => 57m,
            "60-64" => 62m,
            "65-69" => 67m,
            "65+" => 67m,
            "65-69+" => 67m,
            "60 & over" => 62m,
            "70 & over" => 72m,
            _ => 0m
        };
    }

    private static void WritePremiumCheckRows(
        IXLWorksheet ws,
        int startRow,
        string extractedLabel,
        IEnumerable<RateLine> lines,
        decimal? pdfCurrent,
        decimal? pdfRenewal,
        string pdfLabel)
    {
        var lineList = lines.ToList();
        var current = lineList.Sum(line => line.CurrentPremium ?? 0m);
        var renewal = lineList.Sum(line => line.RenewalPremium ?? 0m);

        ws.Cell(startRow, 1).Value = extractedLabel;
        SetMoney(ws.Cell(startRow, 6), current);
        SetMoney(ws.Cell(startRow, 10), renewal);

        ws.Cell(startRow + 1, 1).Value = pdfLabel;
        SetMoney(ws.Cell(startRow + 1, 6), pdfCurrent);
        SetMoney(ws.Cell(startRow + 1, 10), pdfRenewal);

        ws.Cell(startRow + 2, 1).Value = "Variance to PDF";
        ws.Cell(startRow + 2, 6).FormulaA1 = $"F{startRow}-F{startRow + 1}";
        ws.Cell(startRow + 2, 10).FormulaA1 = $"J{startRow}-J{startRow + 1}";
    }

    private static void WriteExperience(XLWorkbook workbook, AnalysisResult analysis)
    {
        var ws = workbook.AddWorksheet("Experience");
        WriteRows(
            ws,
            ["Benefit", "Period", "Average Lives", "Billed Premium", "Paid Claims", "Incurred Claims", "Paid Loss Ratio", "Incurred Loss Ratio", "Weighting", "Carrier Rate Adjustment", "Trend Factor", "Credibility", "Target Loss Ratio", "Source"],
            analysis.Pdf.ExperienceLines.Select(line => new object?[]
            {
                line.Benefit,
                line.Period,
                line.AverageLives,
                line.BilledPremium,
                line.PaidClaims,
                line.IncurredClaims,
                line.PaidLossRatio,
                line.IncurredLossRatio,
                line.Weighting,
                line.ProposedAdjustment,
                line.TrendFactor,
                line.Credibility,
                line.TargetLossRatio,
                line.SourceReference
            }));

        FormatMoneyColumns(ws, 4, 5, 6);
        FormatNumberColumns(ws, 7, 8, 9, 10, 11, 12, 13);
    }

    private static void WriteEhcAnalysis(XLWorkbook workbook, AnalysisResult analysis)
    {
        var ws = workbook.AddWorksheet("EHC Analysis");
        var rows = GetActivityRows(analysis, "Non Pooled Extended Health Care - Subtotal");
        WriteActivityAnalysis(
            ws,
            "Non Pooled Extended Health Care - Subtotal",
            rows,
            "EHC activity rows");
    }

    private static void WriteDentalAnalysis(XLWorkbook workbook, AnalysisResult analysis)
    {
        var ws = workbook.AddWorksheet("Dental Analysis");
        var rows = GetActivityRows(analysis, "Dental");
        WriteActivityAnalysis(
            ws,
            "Dental Analysis",
            rows,
            "Dental activity rows");
    }

    private static void WriteActivityAnalysis(IXLWorksheet ws, string title, IReadOnlyList<ExperienceLine> activityRows, string warningLabel)
    {
        Title(ws, "A1", title);

        var headers = new[] { "Net Billed Premiums", "Paid Claims", "Incurred Claims", "Paid Loss Ratios" };
        for (var index = 0; index < headers.Length; index++)
        {
            ws.Cell(2, index + 2).Value = headers[index];
        }

        ws.Range("B2:E2").Style.Font.Bold = true;

        var labels = new[] { "Prior 2 Periods", "Last Period", "Most Current" };
        var rows = activityRows.TakeLast(3).ToList();
        for (var index = 0; index < labels.Length; index++)
        {
            var rowNumber = index + 3;
            ws.Cell(rowNumber, 1).Value = labels[index];
        }

        var startIndex = Math.Max(0, labels.Length - rows.Count);
        for (var index = 0; index < rows.Count; index++)
        {
            var rowNumber = startIndex + index + 3;

            var row = rows[index];
            SetMoneyWhole(ws.Cell(rowNumber, 2), row.BilledPremium);
            SetMoneyWhole(ws.Cell(rowNumber, 3), row.PaidClaims);
            SetMoneyWhole(ws.Cell(rowNumber, 4), row.IncurredClaims);
            SetPercentValue(ws.Cell(rowNumber, 5), row.PaidLossRatio.HasValue ? row.PaidLossRatio.Value / 100m : null);
        }

        ws.Range("A1:E5").Style.Font.FontColor = XLColor.FromHtml("#004650");
        ws.Range("B3:E5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
        ws.Range("A3:A5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;

        if (rows.Count < 3)
        {
            ws.Cell("A7").Value = $"Warning: expected 3 {warningLabel}, but extracted {rows.Count}.";
            ws.Range("A7:E7").Merge();
            ws.Cell("A7").Style.Fill.BackgroundColor = WarningFill;
        }

        ws.Column(1).Width = 26;
        ws.Columns(2, 5).Width = 18;
    }

    private static List<ExperienceLine> GetActivityRows(AnalysisResult analysis, string benefit)
    {
        return analysis.Pdf.ExperienceLines
            .Where(row => ActivityBenefitMatches(row.Benefit, benefit))
            .OrderBy(row => ParsingHelpers.ParseDateOnly(row.Period.Split(" - ", StringSplitOptions.None).FirstOrDefault()))
            .GroupBy(row => row.Period, StringComparer.OrdinalIgnoreCase)
            .Select(group => group.Last())
            .ToList();
    }

    private static bool ActivityBenefitMatches(string sourceBenefit, string targetBenefit)
    {
        var source = NormalizeAnalysisBenefitKey(sourceBenefit);
        var target = NormalizeAnalysisBenefitKey(targetBenefit);
        if (source.Equals(target, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        if (target.Contains("extendedhealth", StringComparison.OrdinalIgnoreCase) ||
            target.Contains("healthcare", StringComparison.OrdinalIgnoreCase) ||
            target.Equals("health", StringComparison.OrdinalIgnoreCase))
        {
            return source.Contains("extendedhealth", StringComparison.OrdinalIgnoreCase) ||
                   source.Contains("healthcare", StringComparison.OrdinalIgnoreCase) ||
                   source.Equals("health", StringComparison.OrdinalIgnoreCase);
        }

        return false;
    }

    private static string NormalizeAnalysisBenefitKey(string value)
    {
        return new string(ParsingHelpers.Clean(value)
            .Where(char.IsLetterOrDigit)
            .Select(char.ToLowerInvariant)
            .ToArray());
    }

    private static void WriteClaimTrends(
        XLWorkbook workbook,
        AnalysisResult analysis,
        string sheetName,
        string benefitFilter,
        IReadOnlyCollection<ServiceCategoryTrend> spendingTrends)
    {
        var ws = workbook.AddWorksheet(sheetName);
        if (spendingTrends.Count == 0)
        {
            WriteLegacyClaimTrends(ws, analysis, benefitFilter);
            return;
        }

        var trendRows = spendingTrends.ToList();
        if (trendRows.Any(row => ServiceCategoryHistory.GetPeriods(row).Count >= 3))
        {
            WriteThreePeriodClaimTrends(ws, analysis, benefitFilter, trendRows);
            return;
        }

        var totalRow = trendRows.FirstOrDefault(row =>
            NormalizeClaimTrendCategoryKey(row.Category).Equals("total", StringComparison.OrdinalIgnoreCase));
        var detailRows = trendRows.Where(row =>
            !NormalizeClaimTrendCategoryKey(row.Category).Equals("total", StringComparison.OrdinalIgnoreCase));
        var priorTotal = totalRow?.PriorAmount ?? detailRows.Sum(row => row.PriorAmount);
        var currentTotal = totalRow?.CurrentAmount ?? detailRows.Sum(row => row.CurrentAmount);

        var benchmarks = analysis.Workbook.ClaimTrends
            .Where(trend =>
                trend.Benefit.Contains(benefitFilter, StringComparison.OrdinalIgnoreCase) &&
                trend.Period.Equals("Benchmark", StringComparison.OrdinalIgnoreCase))
            .GroupBy(trend => NormalizeClaimTrendCategoryKey(trend.Category), StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.Last(), StringComparer.OrdinalIgnoreCase);

        var rows = trendRows.Select(trend =>
        {
            benchmarks.TryGetValue(NormalizeClaimTrendCategoryKey(trend.Category), out var benchmark);
            return new object?[]
            {
                trend.Benefit,
                trend.Category,
                trend.PriorPeriodLabel,
                trend.PriorAmount,
                CalculateClaimShare(trend.PriorAmount, priorTotal),
                trend.CurrentPeriodLabel,
                trend.CurrentAmount,
                CalculateClaimShare(trend.CurrentAmount, currentTotal),
                ToExcelPercent(benchmark?.BenchmarkPercent),
                ToExcelPercent(benchmark?.RegionalIndustryBenchmarkPercent),
                trend.SourceReference
            };
        });

        WriteRows(
            ws,
            [
                "Benefit",
                "Category",
                "Prior Period",
                "Prior Amount",
                "Prior %",
                "Current Period",
                "Current Amount",
                "Current %",
                "National Benchmark %",
                "Regional Industry Benchmark %",
                "Source"
            ],
            rows);

        FormatMoneyColumns(ws, 4, 7);
        var lastRow = trendRows.Count + 1;
        ws.Range(2, 5, lastRow, 5).Style.NumberFormat.Format = "0.0%";
        ws.Range(2, 8, lastRow, 10).Style.NumberFormat.Format = "0.0%";
        ws.Column(11).Style.Alignment.WrapText = true;
        ApplyPairedClaimTrendLayout(ws);

        for (var index = 0; index < trendRows.Count; index++)
        {
            if (!NormalizeClaimTrendCategoryKey(trendRows[index].Category)
                    .Equals("total", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            ws.Range(index + 2, 1, index + 2, 11).Style.Font.Bold = true;
            ws.Range(index + 2, 1, index + 2, 11).Style.Fill.BackgroundColor = SoftFill;
        }
    }

    private static void WriteThreePeriodClaimTrends(
        IXLWorksheet ws,
        AnalysisResult analysis,
        string benefitFilter,
        IReadOnlyCollection<ServiceCategoryTrend> trendRows)
    {
        var labels = trendRows.SelectMany(row => ServiceCategoryHistory.GetPeriods(row))
            .GroupBy(period => period.PeriodLabel, StringComparer.OrdinalIgnoreCase)
            .Select(group => new { Label = group.Key, Start = ServiceCategoryHistory.ParsePeriodStart(group.Key) ?? DateOnly.MinValue })
            .OrderBy(item => item.Start)
            .TakeLast(3)
            .Select(item => item.Label)
            .ToList();
        var detailRows = trendRows.Where(row => !NormalizeClaimTrendCategoryKey(row.Category).Equals("total", StringComparison.OrdinalIgnoreCase)).ToList();
        var totals = labels.ToDictionary(
            label => label,
            label => trendRows.FirstOrDefault(row => NormalizeClaimTrendCategoryKey(row.Category).Equals("total", StringComparison.OrdinalIgnoreCase)) is { } total
                ? ServiceCategoryHistory.GetPeriods(total).FirstOrDefault(period => period.PeriodLabel.Equals(label, StringComparison.OrdinalIgnoreCase))?.Amount ?? 0m
                : detailRows.Sum(row => ServiceCategoryHistory.GetPeriods(row).FirstOrDefault(period => period.PeriodLabel.Equals(label, StringComparison.OrdinalIgnoreCase))?.Amount ?? 0m),
            StringComparer.OrdinalIgnoreCase);

        var benchmarks = analysis.Workbook.ClaimTrends
            .Where(trend => trend.Benefit.Contains(benefitFilter, StringComparison.OrdinalIgnoreCase) && trend.Period.Equals("Benchmark", StringComparison.OrdinalIgnoreCase))
            .GroupBy(trend => NormalizeClaimTrendCategoryKey(trend.Category), StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.Last(), StringComparer.OrdinalIgnoreCase);
        var headers = new List<string> { "Benefit", "Category" };
        foreach (var label in labels) headers.AddRange([label, $"{label} Amount", $"{label} %"]);
        headers.AddRange(["National Benchmark %", "Regional Industry Benchmark %", "Source"]);
        var rows = trendRows.Select(trend =>
        {
            var values = new List<object?> { trend.Benefit, trend.Category };
            var periods = ServiceCategoryHistory.GetPeriods(trend);
            foreach (var label in labels)
            {
                var amount = periods.FirstOrDefault(period => period.PeriodLabel.Equals(label, StringComparison.OrdinalIgnoreCase))?.Amount ?? 0m;
                values.Add(label); values.Add(amount); values.Add(CalculateClaimShare(amount, totals[label]));
            }
            benchmarks.TryGetValue(NormalizeClaimTrendCategoryKey(trend.Category), out var benchmark);
            values.Add(ToExcelPercent(benchmark?.BenchmarkPercent));
            values.Add(ToExcelPercent(benchmark?.RegionalIndustryBenchmarkPercent));
            values.Add(trend.SourceReference);
            return values.ToArray();
        });
        WriteRows(ws, headers.ToArray(), rows);
        var lastRow = trendRows.Count + 1;
        for (var periodIndex = 0; periodIndex < labels.Count; periodIndex++)
        {
            var amountColumn = 4 + periodIndex * 3;
            ws.Column(amountColumn).Style.NumberFormat.Format = "$#,##0";
            ws.Range(2, amountColumn + 1, lastRow, amountColumn + 1).Style.NumberFormat.Format = "0.0%";
        }
        var benchmarkStart = 3 + labels.Count * 3;
        ws.Range(2, benchmarkStart, lastRow, benchmarkStart + 1).Style.NumberFormat.Format = "0.0%";
        ws.Column(benchmarkStart + 2).Style.Alignment.WrapText = true;
        ApplyPairedClaimTrendLayout(ws);
        for (var index = 0; index < trendRows.Count; index++)
        {
            if (!NormalizeClaimTrendCategoryKey(trendRows.ElementAt(index).Category).Equals("total", StringComparison.OrdinalIgnoreCase)) continue;
            ws.Range(index + 2, 1, index + 2, headers.Count).Style.Font.Bold = true;
            ws.Range(index + 2, 1, index + 2, headers.Count).Style.Fill.BackgroundColor = SoftFill;
        }
    }

    private static void ApplyPairedClaimTrendLayout(IXLWorksheet ws)
    {
        ws.Column(1).Width = 24;
        ws.Column(2).Width = 24;
        ws.Column(3).Width = 22;
        ws.Column(4).Width = 16;
        ws.Column(5).Width = 12;
        ws.Column(6).Width = 22;
        ws.Column(7).Width = 16;
        ws.Column(8).Width = 12;
        ws.Column(9).Width = 19;
        ws.Column(10).Width = 24;
        ws.Column(11).Width = 48;
        ws.Row(1).Height = 42;
    }

    private static void WriteLegacyClaimTrends(
        IXLWorksheet ws,
        AnalysisResult analysis,
        string benefitFilter)
    {
        var rows = analysis.Workbook.ClaimTrends
            .Where(trend => trend.Benefit.Contains(benefitFilter, StringComparison.OrdinalIgnoreCase))
            .Select(trend => new object?[]
            {
                trend.Benefit,
                trend.Category,
                trend.Period,
                trend.Amount,
                trend.PercentOfTotal,
                trend.PriorPeriodPercent,
                trend.BenchmarkPercent,
                trend.RegionalIndustryBenchmarkPercent,
                trend.SourceReference
            });

        WriteRows(
            ws,
            ["Benefit", "Category", "Period", "Amount", "% of Total", "Prior %", "National Benchmark %", "Regional Industry Benchmark %", "Source"],
            rows);

        FormatMoneyColumns(ws, 4);
        FormatNumberColumns(ws, 5, 6, 7, 8);
    }

    private static decimal? CalculateClaimShare(decimal amount, decimal total)
    {
        return total == 0m ? null : amount / total;
    }

    private static string NormalizeClaimTrendCategoryKey(string value)
    {
        var key = new string(ParsingHelpers.Clean(value)
            .Where(char.IsLetterOrDigit)
            .Select(char.ToLowerInvariant)
            .ToArray());

        return key switch
        {
            "basicdental" => "basic",
            "majordental" => "major",
            "endodontics" => "endodontic",
            "orthodontics" => "orthodontic",
            "periodontics" => "periodontic",
            "drugs" => "drug",
            _ => key
        };
    }

    private static void WriteDrugs(XLWorkbook workbook, AnalysisResult analysis)
    {
        var ws = workbook.AddWorksheet("Drugs Indications");
        Title(ws, "A1", "Top Drugs");
        WriteRows(
            ws,
            ["DIN", "Drug Name", "Treatment", "Rx Count", "Net Amount Paid", "Prior Paid", "Source"],
            analysis.Workbook.DrugTrends
                .OrderByDescending(drug => drug.NetAmountPaid ?? drug.SubmittedTotal ?? 0m)
                .Take(10)
                .Select(drug => new object?[]
                {
                    drug.Din,
                    drug.DrugName,
                    drug.Treats,
                    drug.PrescriptionCount,
                    drug.NetAmountPaid ?? drug.SubmittedTotal,
                    drug.PriorPaid,
                    drug.SourceReference
                }),
            3);

        if (analysis.Workbook.PriorDrugTrends.Count > 0)
        {
            var priorStart = ws.LastRowUsed()?.RowNumber() + 3 ?? 32;
            Title(ws, $"A{priorStart}", "Prior Top Drugs");
            WriteRows(
                ws,
                ["DIN", "Drug Name", "Treatment", "Rx Count", "Net Amount Paid", "Prior Paid", "Source"],
                analysis.Workbook.PriorDrugTrends
                    .OrderByDescending(drug => drug.NetAmountPaid ?? drug.SubmittedTotal ?? 0m)
                    .Take(10)
                    .Select(drug => new object?[]
                    {
                        drug.Din,
                        drug.DrugName,
                        drug.Treats,
                        drug.PrescriptionCount,
                        drug.NetAmountPaid ?? drug.SubmittedTotal,
                        drug.PriorPaid,
                        drug.SourceReference
                    }),
                priorStart + 2);
        }

        var indicationStart = ws.LastRowUsed()?.RowNumber() + 3 ?? 32;
        Title(ws, $"A{indicationStart}", "Top Indications");
        WriteRows(
            ws,
            ["Indication", "Rx Count", "Submitted Total", "Allowable Total", "Prior Paid", "Source"],
            analysis.Workbook.IndicationTrends
                .OrderByDescending(indication => indication.AllowableTotal ?? indication.SubmittedTotal ?? 0m)
                .Take(25)
                .Select(indication => new object?[]
                {
                    indication.Indication,
                    indication.PrescriptionCount,
                    indication.SubmittedTotal,
                    indication.AllowableTotal,
                    indication.PriorPaid,
                    indication.SourceReference
                }),
            indicationStart + 2);

        FormatMoneyColumns(ws, 5, 6);
    }

    private static void WriteBenchmarks(XLWorkbook workbook, AnalysisResult analysis)
    {
        var ws = workbook.AddWorksheet("Benchmarks");
        WriteRows(
            ws,
            ["Benefit", "Quartile", "Category", "% of Total", "Average Claim Cost", "Source"],
            analysis.Workbook.QuartileMetrics.Select(metric => new object?[]
            {
                metric.Benefit,
                metric.Quartile,
                metric.Category,
                metric.PercentOfTotal,
                metric.AverageClaimCost,
                metric.SourceReference
            }));

        FormatNumberColumns(ws, 4);
        FormatMoneyColumns(ws, 5);

        var start = ws.LastRowUsed()?.RowNumber() + 3 ?? 4;
        Title(ws, $"A{start}", "Certificate Concentration");
        WriteRows(
            ws,
            ["Certificate", "EHC", "Dental", "Total", "Quartile", "Source"],
            analysis.Workbook.CertificateClaims
                .OrderByDescending(certificate => certificate.Total ?? 0m)
                .Take(25)
                .Select(certificate => new object?[]
                {
                    certificate.Certificate,
                    certificate.ExtendedHealthCare,
                    certificate.Dental,
                    certificate.Total,
                    certificate.Quartile,
                    certificate.SourceReference
                }),
            start + 2);

        FormatMoneyColumns(ws, 2, 3, 4);

        WriteClaimsDistributionBenchmarks(ws, analysis);
    }

    private static void WriteClaimsDistribution(XLWorkbook workbook, AnalysisResult analysis)
    {
        if (workbook.TryGetWorksheet("Claims Distribution", out var existing))
        {
            existing.Delete();
        }

        var ws = workbook.AddWorksheet("Claims Distribution");
        var headers = new[]
        {
            "Benefit",
            "Service Category",
            "Service Description Original",
            "Amount Paid",
            "% Paid",
            "Norm % Paid",
            "Occurrences",
            "% Occurrences",
            "Norm % Occurrences",
            "Avg Claim",
            "Source",
            "Carrier",
            "Policy Number",
            "Report Period",
            "Division",
            "Class"
        };

        for (var column = 0; column < headers.Length; column++)
        {
            ws.Cell(1, column + 1).Value = headers[column];
        }

        Header(ws.Range(1, 1, 1, headers.Length));

        var rowIndex = 2;
        foreach (var row in analysis.Workbook.ClaimsDistributionRows)
        {
            ws.Cell(rowIndex, 1).Value = row.Benefit;
            ws.Cell(rowIndex, 2).Value = row.ServiceCategory;
            ws.Cell(rowIndex, 3).Value = row.ServiceDescriptionOriginal;
            SetMoney(ws.Cell(rowIndex, 4), row.AmountPaid);
            SetPercentNumber(ws.Cell(rowIndex, 5), row.PercentPaid);
            SetPercentNumber(ws.Cell(rowIndex, 6), row.NormPercentPaid);
            if (row.Occurrences.HasValue)
            {
                ws.Cell(rowIndex, 7).Value = row.Occurrences.Value;
            }

            SetPercentNumber(ws.Cell(rowIndex, 8), row.PercentOccurrences);
            SetPercentNumber(ws.Cell(rowIndex, 9), row.NormPercentOccurrences);
            SetMoney(ws.Cell(rowIndex, 10), row.AverageClaim);
            ws.Cell(rowIndex, 11).Value = FirstNonBlank(row.SourceReference, row.SourceFile, row.CarrierName);
            ws.Cell(rowIndex, 12).Value = row.CarrierName;
            ws.Cell(rowIndex, 13).Value = row.PolicyNumber;
            ws.Cell(rowIndex, 14).Value = row.ReportPeriod;
            ws.Cell(rowIndex, 15).Value = row.Division;
            ws.Cell(rowIndex, 16).Value = row.Class;
            rowIndex++;
        }

        ws.Range(1, 1, Math.Max(rowIndex - 1, 1), headers.Length).SetAutoFilter();
        WriteParamedicalDetailReconciliation(
            ws,
            analysis.Workbook.ClaimsDistributionValidation,
            rowIndex + 2);
        ws.SheetView.FreezeRows(1);
        ws.Column(1).Width = 18;
        ws.Column(2).Width = 22;
        ws.Column(3).Width = 28;
        ws.Columns(4, 10).Width = 15;
        ws.Column(11).Width = 24;
        ws.Column(14).Width = 22;
    }

    private static void WriteParamedicalDetailReconciliation(
        IXLWorksheet ws,
        ClaimsDistributionValidation validation,
        int startRow)
    {
        if (!validation.ParamedicalSubtotalAmountPaid.HasValue &&
            !validation.PriorParamedicalSubtotalAmountPaid.HasValue)
        {
            return;
        }

        Title(ws, startRow, 1, "Paramedical Detail Reconciliation");
        var headerRow = startRow + 2;
        var headers = new[]
        {
            "Period",
            "Practitioner Detail Sum",
            "Carrier Paramedical Total",
            "Variance",
            "Status"
        };
        for (var column = 0; column < headers.Length; column++)
        {
            ws.Cell(headerRow, column + 1).Value = headers[column];
        }

        Header(ws.Range(headerRow, 1, headerRow, headers.Length));
        var outputRow = headerRow + 1;
        outputRow = WriteParamedicalReconciliationPeriod(
            ws,
            outputRow,
            "Current",
            validation.SumOfParamedicalDetailAmounts,
            validation.ParamedicalSubtotalAmountPaid,
            validation.ParamedicalDetailVariance,
            validation.ParamedicalDetailsBalanceToTotal,
            validation.UnclassifiedParamedicalDetailFound);
        WriteParamedicalReconciliationPeriod(
            ws,
            outputRow,
            "Prior",
            validation.SumOfPriorParamedicalDetailAmounts,
            validation.PriorParamedicalSubtotalAmountPaid,
            validation.PriorParamedicalDetailVariance,
            validation.PriorParamedicalDetailsBalanceToTotal,
            validation.UnclassifiedParamedicalDetailFound);

        ws.Columns(1, 5).AdjustToContents();
        ws.Columns(2, 4).Style.NumberFormat.Format = "$#,##0.00";
    }

    private static int WriteParamedicalReconciliationPeriod(
        IXLWorksheet ws,
        int row,
        string period,
        decimal? detailSum,
        decimal? subtotal,
        decimal? variance,
        bool? balances,
        bool hasUnclassifiedAmount)
    {
        if (!subtotal.HasValue)
        {
            return row;
        }

        var needsReview = balances != true || hasUnclassifiedAmount;
        ws.Cell(row, 1).Value = period;
        SetMoney(ws.Cell(row, 2), detailSum);
        SetMoney(ws.Cell(row, 3), subtotal);
        SetMoney(ws.Cell(row, 4), variance);
        ws.Cell(row, 5).Value = needsReview ? "Review" : "Match";
        ws.Cell(row, 5).Style.Font.Bold = true;
        ws.Cell(row, 5).Style.Fill.BackgroundColor = needsReview
            ? XLColor.FromHtml("#FFF2CC")
            : XLColor.FromHtml("#E2F0D9");
        return row + 1;
    }

    private static void WriteClaimsDistributionBenchmarks(IXLWorksheet ws, AnalysisResult analysis)
    {
        var rows = analysis.Workbook.ClaimsDistributionRows
            .Where(row => !row.ServiceCategory.Equals("Total", StringComparison.OrdinalIgnoreCase))
            .Where(row => IsBenchmarkReadyClaimsDistributionCategory(row.ServiceCategory))
            .Select(row => new object?[]
            {
                row.Benefit,
                "",
                row.ServiceCategory,
                ToExcelPercent(row.PercentPaid),
                ToExcelPercent(row.NormPercentPaid),
                ToExcelPercent(row.PercentPaid.HasValue && row.NormPercentPaid.HasValue
                    ? row.PercentPaid.Value - row.NormPercentPaid.Value
                    : null),
                row.AverageClaim,
                FirstNonBlank(row.SourceReference, row.CarrierName)
            })
            .ToList();

        var start = ws.LastRowUsed()?.RowNumber() + 3 ?? 1;
        Title(ws, start, 1, "Claims Distribution Benchmarks");
        WriteRows(
            ws,
            ["Benefit", "Quartile", "Category", "% of Total", "Norm % of Total", "Variance to Norm", "Average Claim Cost", "Source"],
            rows,
            start + 2);

        ws.Columns(4, 6).Style.NumberFormat.Format = "0.0%";
        FormatMoneyColumns(ws, 7);
    }

    private static bool IsBenchmarkReadyClaimsDistributionCategory(string category)
    {
        return category.Equals("Drugs", StringComparison.OrdinalIgnoreCase) ||
               category.Equals("Paramedical", StringComparison.OrdinalIgnoreCase) ||
               category.Equals("Hospital", StringComparison.OrdinalIgnoreCase) ||
               category.Equals("Ambulance", StringComparison.OrdinalIgnoreCase) ||
               category.Equals("Medical Supplies", StringComparison.OrdinalIgnoreCase) ||
               category.Equals("Out of Province", StringComparison.OrdinalIgnoreCase) ||
               category.Equals("Out of Country", StringComparison.OrdinalIgnoreCase) ||
               category.Equals("Vision", StringComparison.OrdinalIgnoreCase) ||
               category.Equals("Dental", StringComparison.OrdinalIgnoreCase) ||
               category.Equals("Other", StringComparison.OrdinalIgnoreCase);
    }

    private static void WriteDocumentDiscovery(XLWorkbook workbook, AnalysisResult analysis)
    {
        if (analysis.DocumentDiscovery is null)
        {
            return;
        }

        var ws = workbook.AddWorksheet("Document Discovery");
        WriteRows(
            ws,
            [
                "Filename",
                "Carrier",
                "Document Type",
                "Confidence",
                "Contract Number",
                "Contract Holder",
                "Rate Effective Date",
                "Status",
                "Missing Companion Document",
                "Validation Notes",
                "Matched Signals",
                "Pages Detected"
            ],
            analysis.DocumentDiscovery.Documents.Select(document => new object?[]
            {
                document.FileName,
                document.Carrier,
                document.DocumentType,
                document.DocumentConfidenceScore,
                document.ContractNumber,
                document.ContractHolder,
                document.RateEffectiveDate?.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
                document.Status,
                MissingCompanionDocument(analysis.DocumentDiscovery, document),
                DocumentValidationNotes(analysis.DocumentDiscovery, document),
                string.Join("; ", document.MatchedSignals),
                string.Join(", ", document.PagesDetected)
            }));

        ws.SheetView.FreezeRows(1);
        ws.Column(1).Width = 42;
        ws.Columns(2, 8).Width = 20;
        ws.Column(9).Width = 28;
        ws.Column(10).Width = 34;
        ws.Column(11).Width = 52;
        ws.Column(12).Width = 18;
        ws.Column(4).Style.NumberFormat.Format = "0%";
        ws.Columns().Style.Alignment.WrapText = true;
    }

    private static void WriteSunLifeExperience(XLWorkbook workbook, AnalysisResult analysis)
    {
        if (analysis.Pdf.SunLifeExperienceRows.Count == 0)
        {
            return;
        }

        var ws = workbook.AddWorksheet("SunLife_Experience");
        WriteRows(
            ws,
            [
                "Benefit",
                "Period",
                "Average Covered Employees",
                "Billed Premium",
                "Adjusted Premium",
                "Paid Claims",
                "Incurred Claims",
                "Adjusted Loss Ratio",
                "Weighting",
                "Large Amount Pooling Premium",
                "Quebec Industry Pooling Premium",
                "Pooled Claims",
                "Quebec Industry Pooled Claims",
                "Weighted Adjusted Loss Ratio",
                "Credibility",
                "IBNR",
                "Target Loss Ratio",
                "Annual Trend Factor",
                "Rate Adjustment",
                "Source"
            ],
            analysis.Pdf.SunLifeExperienceRows.Select(row => new object?[]
            {
                row.Benefit,
                row.Period,
                row.AverageCoveredEmployees,
                row.BilledPremium,
                row.AdjustedPremium,
                row.PaidClaims,
                row.IncurredClaims,
                ToExcelPercent(row.AdjustedLossRatio),
                ToExcelPercent(row.Weighting),
                row.LargeAmountPoolingPremium,
                row.QuebecIndustryPoolingPremium,
                row.PooledClaims,
                row.QuebecIndustryPooledClaims,
                ToExcelPercent(row.WeightedAdjustedLossRatio),
                ToExcelPercent(row.Credibility),
                ToExcelPercent(row.Ibnr),
                ToExcelPercent(row.TargetLossRatio),
                ToExcelPercent(row.AnnualTrendFactor),
                ToExcelPercent(row.RateAdjustment),
                row.SourceReference
            }));

        ws.SheetView.FreezeRows(1);
        ws.Column(1).Width = 22;
        ws.Column(2).Width = 28;
        ws.Columns(3, 19).Width = 16;
        ws.Column(20).Width = 28;
        FormatNumberColumns(ws, 3);
        FormatMoneyColumns(ws, 4, 5, 6, 7, 10, 11, 12, 13);
        ws.Columns(8, 9).Style.NumberFormat.Format = "0.0%";
        ws.Columns(14, 19).Style.NumberFormat.Format = "0.0%";
    }

    private static string MissingCompanionDocument(RenewalDocumentDiscoveryResult discovery, DiscoveredDocument document)
    {
        if (!document.Carrier.Equals("Sun Life", StringComparison.OrdinalIgnoreCase))
        {
            return "";
        }

        var hasRenewal = discovery.Documents.Any(item =>
            item.DocumentType.Equals(DiscoveryDocumentTypes.SunLifeRenewalReport, StringComparison.OrdinalIgnoreCase));
        var hasExperience = discovery.Documents.Any(item =>
            item.DocumentType.Equals(DiscoveryDocumentTypes.SunLifeExperienceAnalysis, StringComparison.OrdinalIgnoreCase));

        if (!hasRenewal)
        {
            return "SUN_LIFE_RENEWAL_REPORT";
        }

        return hasExperience ? "" : "SUN_LIFE_EXPERIENCE_ANALYSIS";
    }

    private static string DocumentValidationNotes(RenewalDocumentDiscoveryResult discovery, DiscoveredDocument document)
    {
        if (!document.Carrier.Equals("Sun Life", StringComparison.OrdinalIgnoreCase))
        {
            return "";
        }

        var companionDocuments = discovery.Documents
            .Where(item =>
                item.Carrier.Equals("Sun Life", StringComparison.OrdinalIgnoreCase) &&
                !ReferenceEquals(item, document))
            .ToList();
        var notes = new List<string>();

        foreach (var companion in companionDocuments)
        {
            if (!string.IsNullOrWhiteSpace(document.ContractNumber) &&
                !string.IsNullOrWhiteSpace(companion.ContractNumber) &&
                !document.ContractNumber.Equals(companion.ContractNumber, StringComparison.OrdinalIgnoreCase))
            {
                notes.Add($"Contract number differs from {companion.FileName}.");
            }

            if (!string.IsNullOrWhiteSpace(document.ContractHolder) &&
                !string.IsNullOrWhiteSpace(companion.ContractHolder) &&
                !document.ContractHolder.Equals(companion.ContractHolder, StringComparison.OrdinalIgnoreCase))
            {
                notes.Add($"Contract holder differs from {companion.FileName}.");
            }

            if (document.RateEffectiveDate.HasValue &&
                companion.RateEffectiveDate.HasValue &&
                document.RateEffectiveDate.Value != companion.RateEffectiveDate.Value)
            {
                notes.Add($"Rate effective date differs from {companion.FileName}.");
            }
        }

        return notes.Count == 0 ? "" : string.Join(" ", notes.Distinct(StringComparer.OrdinalIgnoreCase));
    }

    private static void WriteSourceChecks(XLWorkbook workbook, AnalysisResult analysis)
    {
        var ws = workbook.AddWorksheet("Source Checks");
        var checks = analysis.SourceChecks
            .Select(check => new object?[] { check.Area, check.Status, check.Detail })
            .ToList();
        if (IsCanadaLifeAnalysis(analysis))
        {
            var reconciliation = GetCanadaLifeExperienceReconciliation(analysis);
            checks.Add(CanadaLifeComponentSourceCheck("Health", "CLhealthanalysis", reconciliation.HealthRows));
            checks.Add(CanadaLifeComponentSourceCheck("Drug", "CLdruganalysis", reconciliation.DrugRows));
            if (reconciliation.VisionRows.Count > 0)
            {
                checks.Add(CanadaLifeComponentSourceCheck("Vision", "CLvisionanalysis", reconciliation.VisionRows));
            }

            var unmatched = new List<string>();
            AddUnmatchedCanadaLifePeriods(unmatched, "Health", reconciliation.UnmatchedHealthPeriods);
            AddUnmatchedCanadaLifePeriods(unmatched, "Drug", reconciliation.UnmatchedDrugPeriods);
            AddUnmatchedCanadaLifePeriods(unmatched, "Vision", reconciliation.UnmatchedVisionPeriods);
            checks.Add(new object?[]
            {
                "Canada Life combined EHC experience",
                unmatched.Count == 0 && reconciliation.CombinedRows.Count > 0 ? "Pass" : "Needs Review",
                reconciliation.DrugRows.Count == 0 && reconciliation.HealthRows.Count > 0
                    ? "No separate Drug experience rows were available. Health experience remains visible without aggregation; source review is required."
                    : unmatched.Count == 0 && reconciliation.CombinedRows.Count > 0
                    ? $"{reconciliation.CombinedRows.Count} exact Health + Drug period(s) combined; Vision included only on the same complete dates. Premium, paid claims, and incurred claims are summed; incurred loss ratio is recalculated from totals."
                    : reconciliation.CombinedRows.Count == 0
                        ? "No exact complete-date Health + Drug period match was found; no components were silently combined."
                        : $"Only exact complete-date matches were combined. Source check required for: {string.Join("; ", unmatched)}."
            });
        }

        WriteRows(
            ws,
            ["Area", "Status", "Detail"],
            checks);
    }

    private static object?[] CanadaLifeComponentSourceCheck(string component, string sheetName, IReadOnlyCollection<ExperienceLine> rows)
    {
        return
        [
            $"Canada Life {component} component experience",
            rows.Count > 0 ? "Pass" : "Needs Review",
            rows.Count > 0
                ? $"{rows.Count} source-extracted row(s) written without aggregation to {sheetName}; the source audit table preserves period, premium, paid claims, incurred claims, recalculated loss ratio, and PDF reference."
                : $"No source-extracted {component} experience rows were available; {sheetName} was not created."
        ];
    }

    private static void AddUnmatchedCanadaLifePeriods(List<string> details, string component, IEnumerable<FormulaPeriodKey> periods)
    {
        details.AddRange(periods.Select(period => $"{component} {period.Start:MMM dd, yyyy} - {period.End:MMM dd, yyyy}"));
    }

    private void WriteUnderwritingReferenceMaster(
        XLWorkbook workbook,
        string sheetName,
        string rateType,
        string demographicSheetName,
        IReadOnlyCollection<string> demographicAgeBands,
        decimal volumeDivisor)
    {
        if (workbook.TryGetWorksheet(sheetName, out var existing))
        {
            existing.Delete();
        }

        var candidates = _underwritingReferenceRates
            .Where(rate => rate.RateType.Equals(rateType, StringComparison.OrdinalIgnoreCase))
            .ToList();
        var versionName = candidates.FirstOrDefault()?.VersionName ?? UnderwritingReferenceRateCatalog.VersionName;
        var rates = candidates
            .Where(rate => rate.VersionName.Equals(versionName, StringComparison.OrdinalIgnoreCase))
            .GroupBy(rate => rate.AgeBand, StringComparer.OrdinalIgnoreCase)
            .Select(group => new
            {
                AgeBand = group.Key,
                DisplayOrder = group.Min(rate => rate.DisplayOrder),
                FemaleRate = group.FirstOrDefault(rate => rate.Gender.Equals(UnderwritingRateGenders.Female, StringComparison.OrdinalIgnoreCase))?.Rate,
                MaleRate = group.FirstOrDefault(rate => rate.Gender.Equals(UnderwritingRateGenders.Male, StringComparison.OrdinalIgnoreCase))?.Rate,
                UnitBasis = group.First().UnitBasis,
                SourceName = group.First().SourceName,
                SourceSheet = group.First().SourceSheet
            })
            .Where(row => row.FemaleRate.HasValue && row.MaleRate.HasValue)
            .OrderBy(row => row.DisplayOrder)
            .ThenBy(row => row.AgeBand, StringComparer.OrdinalIgnoreCase)
            .ToList();

        var ws = workbook.AddWorksheet(sheetName);
        ws.Cell("A1").Value = $"{rateType} Master — Underwriting Reference Rates";
        ws.Range("A1:L1").Merge();
        ws.Range("A1:L1").Style.Fill.BackgroundColor = LineRatesDarkTeal;
        ws.Range("A1:L1").Style.Font.FontColor = XLColor.White;
        ws.Range("A1:L1").Style.Font.Bold = true;
        ws.Range("A1:L1").Style.Font.FontSize = 14;

        ws.Cell("B2").Value = "Current";
        ws.Range("B2:E2").Merge();
        ws.Cell("F2").Value = "Previous";
        ws.Range("F2:I2").Merge();
        ws.Cell("J2").Value = "Total";
        ws.Range("J2:L2").Merge();
        Header(ws.Range("B2:L2"));
        ws.Range("B2:L2").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

        var headerRow3 = new[]
        {
            "Age", "Female", "Female", "Male", "Male", "Female", "Female", "Male", "Male", "Current", "Previous", "Change"
        };
        var headerRow4 = new[]
        {
            "Group", "Rate", "Weighted premium", "Rate", "Weighted premium", "Rate", "Weighted premium", "Rate", "Weighted premium", "Weighted premium", "Weighted premium", "Risk"
        };
        for (var column = 1; column <= 12; column++)
        {
            ws.Cell(3, column).Value = headerRow3[column - 1];
            ws.Cell(4, column).Value = headerRow4[column - 1];
        }
        Header(ws.Range("A3:L4"));
        ws.Range("A3:L4").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        ws.Range("A3:L4").Style.Alignment.WrapText = true;

        const int firstRateRow = 6;
        var sourceFirstRow = 5;
        var sourceLastRow = sourceFirstRow + demographicAgeBands.Count - 1;
        for (var index = 0; index < rates.Count; index++)
        {
            var row = firstRateRow + index;
            var rate = rates[index];
            ws.Cell(row, 1).Value = rate.AgeBand;
            ws.Cell(row, 2).Value = rate.FemaleRate!.Value;
            ws.Cell(row, 4).Value = rate.MaleRate!.Value;
            ws.Cell(row, 6).Value = rate.FemaleRate.Value;
            ws.Cell(row, 8).Value = rate.MaleRate.Value;

            ws.Cell(row, 3).FormulaA1 = $"(({DemographicBandSumFormula(demographicSheetName, sourceFirstRow, sourceLastRow, 3, rate.AgeBand)})*B{row})/{volumeDivisor.ToString(CultureInfo.InvariantCulture)}";
            ws.Cell(row, 5).FormulaA1 = $"(({DemographicBandSumFormula(demographicSheetName, sourceFirstRow, sourceLastRow, 5, rate.AgeBand)})*D{row})/{volumeDivisor.ToString(CultureInfo.InvariantCulture)}";
            ws.Cell(row, 7).FormulaA1 = $"(({DemographicBandSumFormula(demographicSheetName, sourceFirstRow, sourceLastRow, 7, rate.AgeBand)})*F{row})/{volumeDivisor.ToString(CultureInfo.InvariantCulture)}";
            ws.Cell(row, 9).FormulaA1 = $"(({DemographicBandSumFormula(demographicSheetName, sourceFirstRow, sourceLastRow, 9, rate.AgeBand)})*H{row})/{volumeDivisor.ToString(CultureInfo.InvariantCulture)}";
            ws.Cell(row, 10).FormulaA1 = $"C{row}+E{row}";
            ws.Cell(row, 11).FormulaA1 = $"G{row}+I{row}";
        }

        var lastRateRow = firstRateRow + rates.Count - 1;
        var totalRow = lastRateRow + 2;
        ws.Cell(totalRow, 1).Value = "TOTALS";
        foreach (var column in new[] { 3, 5, 7, 9, 10, 11 })
        {
            ws.Cell(totalRow, column).FormulaA1 = $"SUM({ws.Cell(firstRateRow, column).Address}:{ws.Cell(lastRateRow, column).Address})";
        }

        const int dataHeaderRow = 19;
        ws.Cell(dataHeaderRow, 1).Value = "Data";
        ws.Range(dataHeaderRow, 1, dataHeaderRow, 3).Merge();
        Header(ws.Range(dataHeaderRow, 1, dataHeaderRow, 3));
        ws.Cell(20, 1).Value = "Current Volume";
        ws.Cell(21, 1).Value = "Current EE";
        ws.Cell(22, 1).Value = "Previous Volume";
        ws.Cell(23, 1).Value = "Previous EE";
        ws.Cell(20, 3).FormulaA1 = $"SUM('{demographicSheetName}'!C{sourceFirstRow}:C{sourceLastRow},'{demographicSheetName}'!E{sourceFirstRow}:E{sourceLastRow})";
        ws.Cell(21, 3).FormulaA1 = $"SUM('{demographicSheetName}'!B{sourceFirstRow}:B{sourceLastRow},'{demographicSheetName}'!D{sourceFirstRow}:D{sourceLastRow})";
        ws.Cell(22, 3).FormulaA1 = $"SUM('{demographicSheetName}'!G{sourceFirstRow}:G{sourceLastRow},'{demographicSheetName}'!I{sourceFirstRow}:I{sourceLastRow})";
        ws.Cell(23, 3).FormulaA1 = $"SUM('{demographicSheetName}'!F{sourceFirstRow}:F{sourceLastRow},'{demographicSheetName}'!H{sourceFirstRow}:H{sourceLastRow})";
        ws.Cell(totalRow, 12).FormulaA1 = $"IF(OR(C20<=0,C22<=0,K{totalRow}<=0),\"\",((J{totalRow}/C20)/(K{totalRow}/C22))-1)";

        var metadataRow = 25;
        ws.Cell(metadataRow, 1).Value = "Version";
        ws.Cell(metadataRow, 2).Value = versionName;
        ws.Cell(metadataRow + 1, 1).Value = "Unit basis";
        ws.Cell(metadataRow + 1, 2).Value = rates.FirstOrDefault()?.UnitBasis ?? "";
        ws.Cell(metadataRow + 2, 1).Value = "Source";
        ws.Cell(metadataRow + 2, 2).Value = rates.Count == 0 ? "" : $"{rates[0].SourceName} — {rates[0].SourceSheet}";
        ws.Cell(metadataRow + 3, 1).Value = "Control note";
        ws.Cell(metadataRow + 3, 2).Value = rateType.Equals(UnderwritingRateTypes.Mortality, StringComparison.OrdinalIgnoreCase)
            ? "Calculated Life rate remains blank until both current and previous Life volumes are greater than zero."
            : "Morbidity change remains blank until both current and previous LTD volumes are greater than zero.";
        for (var row = metadataRow; row <= metadataRow + 3; row++)
        {
            ws.Range(row, 2, row, 12).Merge();
        }

        WriteSoaBenchmarkBlock(
            ws,
            rateType,
            demographicSheetName,
            sourceFirstRow,
            sourceLastRow,
            totalRow,
            metadataRow + 6);

        foreach (var column in new[] { 2, 4, 6, 8 })
        {
            ws.Range(firstRateRow, column, lastRateRow, column).Style.NumberFormat.Format = "0.0000";
        }
        foreach (var column in new[] { 3, 5, 7, 9, 10, 11 })
        {
            ws.Range(firstRateRow, column, lastRateRow, column).Style.NumberFormat.Format = "$#,##0.00";
        }
        ws.Range(totalRow, 3, totalRow, 11).Style.NumberFormat.Format = "$#,##0.00";
        ws.Cell(totalRow, 12).Style.NumberFormat.Format = "0.0%";
        ws.Range(firstRateRow, 1, lastRateRow, 5).Style.Fill.BackgroundColor = XLColor.FromHtml("#E8EFCF");
        ws.Range(firstRateRow, 6, lastRateRow, 9).Style.Fill.BackgroundColor = XLColor.FromHtml("#D9EEF7");
        ws.Range(firstRateRow, 10, lastRateRow, 12).Style.Fill.BackgroundColor = XLColor.FromHtml("#FFF2CC");
        ws.Range(totalRow, 1, totalRow, 12).Style.Font.Bold = true;
        ws.Range(totalRow, 1, totalRow, 12).Style.Fill.BackgroundColor = SoftFill;
        ws.Column(1).Width = 30;
        ws.Columns(2, 12).Width = 15;
        ws.Row(3).Height = 24;
        ws.Row(4).Height = 36;
        ws.SheetView.FreezeRows(4);
    }

    private void WriteSoaBenchmarkBlock(
        IXLWorksheet ws,
        string westlandRateType,
        string demographicSheetName,
        int sourceFirstRow,
        int sourceLastRow,
        int westlandTotalRow,
        int titleRow)
    {
        var benchmarkRateType = westlandRateType.Equals(UnderwritingRateTypes.Mortality, StringComparison.OrdinalIgnoreCase)
            ? UnderwritingRateTypes.SoaBasicLifeMortality
            : UnderwritingRateTypes.SoaGroupLifeWaiver;
        var benchmarkRates = _underwritingReferenceRates
            .Where(rate => rate.RateType.Equals(benchmarkRateType, StringComparison.OrdinalIgnoreCase))
            .GroupBy(rate => rate.AgeBand, StringComparer.OrdinalIgnoreCase)
            .Select(group => new
            {
                AgeBand = group.Key,
                DisplayOrder = group.Min(rate => rate.DisplayOrder),
                FemaleRate = group.FirstOrDefault(rate => rate.Gender.Equals(UnderwritingRateGenders.Female, StringComparison.OrdinalIgnoreCase))?.Rate,
                MaleRate = group.FirstOrDefault(rate => rate.Gender.Equals(UnderwritingRateGenders.Male, StringComparison.OrdinalIgnoreCase))?.Rate,
                VersionName = group.First().VersionName,
                UnitBasis = group.First().UnitBasis,
                SourceName = group.First().SourceName,
                SourceSheet = group.First().SourceSheet
            })
            .Where(row => row.FemaleRate.HasValue && row.MaleRate.HasValue)
            .OrderBy(row => row.DisplayOrder)
            .ThenBy(row => row.AgeBand, StringComparer.OrdinalIgnoreCase)
            .ToList();
        if (benchmarkRates.Count == 0)
        {
            return;
        }

        var isMortality = benchmarkRateType.Equals(UnderwritingRateTypes.SoaBasicLifeMortality, StringComparison.OrdinalIgnoreCase);
        ws.Cell(titleRow, 1).Value = isMortality
            ? "SOA Benchmark — Basic Group Life Expected Mortality"
            : "SOA Benchmark — Group Life Waiver of Premium (Informational Only)";
        ws.Range(titleRow, 1, titleRow, 12).Merge();
        ws.Range(titleRow, 1, titleRow, 12).Style.Fill.BackgroundColor = LineRatesDarkTeal;
        ws.Range(titleRow, 1, titleRow, 12).Style.Font.FontColor = XLColor.White;
        ws.Range(titleRow, 1, titleRow, 12).Style.Font.Bold = true;
        ws.Range(titleRow, 1, titleRow, 12).Style.Font.FontSize = 12;

        var groupHeaderRow = titleRow + 1;
        ws.Cell(groupHeaderRow, 2).Value = "Current";
        ws.Range(groupHeaderRow, 2, groupHeaderRow, 5).Merge();
        ws.Cell(groupHeaderRow, 6).Value = "Previous";
        ws.Range(groupHeaderRow, 6, groupHeaderRow, 7).Merge();
        ws.Cell(groupHeaderRow, 8).Value = "SOA Expected-Risk Comparison";
        ws.Range(groupHeaderRow, 8, groupHeaderRow, 12).Merge();
        Header(ws.Range(groupHeaderRow, 1, groupHeaderRow, 12));
        ws.Range(groupHeaderRow, 2, groupHeaderRow, 5).Style.Fill.BackgroundColor = XLColor.FromHtml("#5D6B1D");
        ws.Range(groupHeaderRow, 6, groupHeaderRow, 7).Style.Fill.BackgroundColor = XLColor.FromHtml("#006878");
        ws.Range(groupHeaderRow, 8, groupHeaderRow, 12).Style.Fill.BackgroundColor = XLColor.FromHtml("#D6A900");

        var headerRow = titleRow + 2;
        var headers = new[]
        {
            "Age Group", "Female SOA Rate", "Current Female Expected", "Male SOA Rate", "Current Male Expected",
            "Previous Female Expected", "Previous Male Expected", "Current Expected", "Previous Expected",
            "Current Expected / Volume", "Previous Expected / Volume", "SOA Mix Change"
        };
        for (var column = 1; column <= headers.Length; column++)
        {
            ws.Cell(headerRow, column).Value = headers[column - 1];
        }
        Header(ws.Range(headerRow, 1, headerRow, 12));
        ws.Range(groupHeaderRow, 1, headerRow, 12).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        ws.Range(groupHeaderRow, 1, headerRow, 12).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        ws.Range(groupHeaderRow, 1, headerRow, 12).Style.Alignment.WrapText = true;

        const decimal volumeDivisor = 1_000m;
        var firstDataRow = titleRow + 3;
        for (var index = 0; index < benchmarkRates.Count; index++)
        {
            var row = firstDataRow + index;
            var rate = benchmarkRates[index];
            var currentFemaleVolume = DemographicBandSumFormula(demographicSheetName, sourceFirstRow, sourceLastRow, 3, rate.AgeBand);
            var currentMaleVolume = DemographicBandSumFormula(demographicSheetName, sourceFirstRow, sourceLastRow, 5, rate.AgeBand);
            var previousFemaleVolume = DemographicBandSumFormula(demographicSheetName, sourceFirstRow, sourceLastRow, 7, rate.AgeBand);
            var previousMaleVolume = DemographicBandSumFormula(demographicSheetName, sourceFirstRow, sourceLastRow, 9, rate.AgeBand);

            ws.Cell(row, 1).Value = rate.AgeBand;
            ws.Cell(row, 2).Value = rate.FemaleRate!.Value;
            ws.Cell(row, 3).FormulaA1 = $"(({currentFemaleVolume})*B{row})/{volumeDivisor.ToString(CultureInfo.InvariantCulture)}";
            ws.Cell(row, 4).Value = rate.MaleRate!.Value;
            ws.Cell(row, 5).FormulaA1 = $"(({currentMaleVolume})*D{row})/{volumeDivisor.ToString(CultureInfo.InvariantCulture)}";
            ws.Cell(row, 6).FormulaA1 = $"(({previousFemaleVolume})*B{row})/{volumeDivisor.ToString(CultureInfo.InvariantCulture)}";
            ws.Cell(row, 7).FormulaA1 = $"(({previousMaleVolume})*D{row})/{volumeDivisor.ToString(CultureInfo.InvariantCulture)}";
            ws.Cell(row, 8).FormulaA1 = $"C{row}+E{row}";
            ws.Cell(row, 9).FormulaA1 = $"F{row}+G{row}";
            ws.Cell(row, 10).FormulaA1 = $"IF(({currentFemaleVolume})+({currentMaleVolume})<=0,\"\",H{row}/(({currentFemaleVolume})+({currentMaleVolume})))";
            ws.Cell(row, 11).FormulaA1 = $"IF(({previousFemaleVolume})+({previousMaleVolume})<=0,\"\",I{row}/(({previousFemaleVolume})+({previousMaleVolume})))";
        }

        var lastDataRow = firstDataRow + benchmarkRates.Count - 1;
        var totalRow = lastDataRow + 2;
        ws.Cell(totalRow, 1).Value = "TOTALS";
        foreach (var column in new[] { 3, 5, 6, 7, 8, 9 })
        {
            ws.Cell(totalRow, column).FormulaA1 = $"SUM({ws.Cell(firstDataRow, column).Address}:{ws.Cell(lastDataRow, column).Address})";
        }
        ws.Cell(totalRow, 10).FormulaA1 = $"IF(C20<=0,\"\",H{totalRow}/C20)";
        ws.Cell(totalRow, 11).FormulaA1 = $"IF(C22<=0,\"\",I{totalRow}/C22)";
        ws.Cell(totalRow, 12).FormulaA1 = $"IF(OR(J{totalRow}<=0,K{totalRow}<=0),\"\",J{totalRow}/K{totalRow}-1)";

        var metricsRow = totalRow + 2;
        ws.Cell(metricsRow, 1).Value = "SOA benchmark change in risk";
        ws.Cell(metricsRow, 2).FormulaA1 = $"L{totalRow}";
        ws.Cell(metricsRow + 1, 1).Value = "Westland master change in risk";
        ws.Cell(metricsRow + 1, 2).FormulaA1 = $"L{westlandTotalRow}";
        ws.Cell(metricsRow + 2, 1).Value = "Difference (SOA - Westland)";
        ws.Cell(metricsRow + 2, 2).FormulaA1 = $"IF(OR(B{metricsRow}=\"\",B{metricsRow + 1}=\"\"),\"\",B{metricsRow}-B{metricsRow + 1})";
        ws.Range(metricsRow, 1, metricsRow + 2, 2).Style.Font.Bold = true;
        ws.Range(metricsRow, 1, metricsRow + 2, 2).Style.Fill.BackgroundColor = XLColor.FromHtml("#FFF2CC");
        ws.Range(metricsRow, 2, metricsRow + 2, 2).Style.NumberFormat.Format = "0.0%";

        var noteRow = metricsRow + 4;
        var notes = new[]
        {
            ("Version", benchmarkRates[0].VersionName),
            ("Unit basis", benchmarkRates[0].UnitBasis),
            ("Source", $"{benchmarkRates[0].SourceName} — {benchmarkRates[0].SourceSheet}"),
            ("A/E status", "Not calculated — observed deaths or waiver claims are not present in renewal demographics."),
            ("Use in calculated rate", "No — informational benchmark only; the Westland master change above remains the renewal-rate driver."),
            ("Scope note", isMortality
                ? "U.S. Basic Group Life experience benchmark. Use as a reasonableness check, not as a Canadian pricing basis."
                : "This SOA measure is Group Life disability waiver of premium. It is not GLTD incidence or an LTD pricing factor.")
        };
        for (var index = 0; index < notes.Length; index++)
        {
            var row = noteRow + index;
            ws.Cell(row, 1).Value = notes[index].Item1;
            ws.Cell(row, 2).Value = notes[index].Item2;
            ws.Range(row, 2, row, 12).Merge();
        }

        ws.Range(firstDataRow, 1, lastDataRow, 5).Style.Fill.BackgroundColor = XLColor.FromHtml("#E8EFCF");
        ws.Range(firstDataRow, 6, lastDataRow, 7).Style.Fill.BackgroundColor = XLColor.FromHtml("#D9EEF7");
        ws.Range(firstDataRow, 8, lastDataRow, 12).Style.Fill.BackgroundColor = XLColor.FromHtml("#FFF2CC");
        ws.Range(totalRow, 1, totalRow, 12).Style.Font.Bold = true;
        ws.Range(totalRow, 1, totalRow, 12).Style.Fill.BackgroundColor = SoftFill;
        ws.Range(firstDataRow, 2, lastDataRow, 2).Style.NumberFormat.Format = "0.0000";
        ws.Range(firstDataRow, 4, lastDataRow, 4).Style.NumberFormat.Format = "0.0000";
        foreach (var column in new[] { 3, 5, 6, 7, 8, 9 })
        {
            ws.Range(firstDataRow, column, totalRow, column).Style.NumberFormat.Format = "#,##0.00";
        }
        ws.Range(firstDataRow, 10, totalRow, 11).Style.NumberFormat.Format = "0.000000";
        ws.Cell(totalRow, 12).Style.NumberFormat.Format = "0.0%";
        ws.Row(titleRow).Height = 24;
        ws.Row(headerRow).Height = 42;
    }

    private static string DemographicBandSumFormula(
        string sheetName,
        int firstRow,
        int lastRow,
        int valueColumn,
        string ageBand)
    {
        var criteria = ageBand.Equals("65-69", StringComparison.OrdinalIgnoreCase)
            ? new[] { "65-69", "65+" }
            : new[] { ageBand };
        var valueColumnLetter = XLHelper.GetColumnLetterFromNumber(valueColumn);
        return string.Join(
            "+",
            criteria.Select(item =>
                $"SUMIF('{sheetName}'!$A${firstRow}:$A${lastRow},\"{item}\",'{sheetName}'!${valueColumnLetter}${firstRow}:${valueColumnLetter}${lastRow})"));
    }

    private static void WriteAsoBillingActivity(XLWorkbook workbook, AnalysisResult analysis)
    {
        if (analysis.AsoBillingActivity.Count == 0)
        {
            return;
        }

        var ws = workbook.AddWorksheet("ASO Billing Activity");
        WriteRows(
            ws,
            ["Service Month", "Paid Dental Claims", "Billed Fees", "GST", "Interest / Credit", "Actual Total Charges", "Fee % of Claims", "Source File"],
            analysis.AsoBillingActivity
                .OrderBy(row => row.ServiceMonth)
                .Select(row => new object?[]
                {
                    row.ServiceMonth.ToString("MMM yyyy", CultureInfo.InvariantCulture),
                    row.PaidDentalClaims,
                    row.AdministrationFee,
                    row.Gst,
                    row.InterestCreditCharge,
                    row.TotalAsoCost,
                    row.AdministrationFee.HasValue && row.PaidDentalClaims > 0m
                        ? row.AdministrationFee.Value / row.PaidDentalClaims
                        : null,
                    Path.GetFileName(row.SourcePath)
                }));

        var totalRow = analysis.AsoBillingActivity.Count + 2;
        ws.Cell(totalRow, 1).Value = "Total";
        ws.Cell(totalRow, 2).FormulaA1 = $"=SUM(B2:B{totalRow - 1})";
        ws.Cell(totalRow, 3).FormulaA1 = $"=SUM(C2:C{totalRow - 1})";
        ws.Cell(totalRow, 4).FormulaA1 = $"=SUM(D2:D{totalRow - 1})";
        ws.Cell(totalRow, 5).FormulaA1 = $"=SUM(E2:E{totalRow - 1})";
        ws.Cell(totalRow, 6).FormulaA1 = $"=SUM(F2:F{totalRow - 1})";
        ws.Range(2, 2, totalRow, 6).Style.NumberFormat.Format = "$#,##0.00;($#,##0.00)";
        ws.Range(2, 7, totalRow, 7).Style.NumberFormat.Format = "0.00%";
        ws.Row(totalRow).Style.Font.Bold = true;
        ws.Row(totalRow).Style.Fill.BackgroundColor = SoftFill;
    }

    private static void WriteAssurePlusExperience(XLWorkbook workbook, AnalysisResult analysis)
    {
        var ws = workbook.AddWorksheet("AssurePLUS Experience");
        Title(ws, "A1", "AssurePLUS Experience");
        var summary = analysis.AssurePlusExperience;
        if (summary is null)
        {
            ws.Cell("A3").Value = "No AssurePLUS Health and Dental pricing page was extracted from the selected renewal PDFs.";
            ws.Range("A3:F3").Merge();
            ws.Range("A3:F3").Style.Fill.BackgroundColor = WarningFill;
            return;
        }

        ws.Cell("A3").Value = "Benefit";
        var years = summary.AnnualRatings
            .Select(row => row.Year)
            .Distinct()
            .OrderBy(year => year)
            .ToList();
        for (var index = 0; index < years.Count; index++)
        {
            ws.Cell(3, index + 2).Value = years[index];
        }
        Header(ws.Range(3, 1, 3, Math.Max(2, years.Count + 1)));

        var historyRow = 4;
        foreach (var benefit in new[] { "Health", "Dental" })
        {
            ws.Cell(historyRow, 1).Value = benefit;
            for (var index = 0; index < years.Count; index++)
            {
                var rating = summary.AnnualRatings.FirstOrDefault(row =>
                    row.Year == years[index] && row.Benefit.Equals(benefit, StringComparison.OrdinalIgnoreCase));
                ws.Cell(historyRow, index + 2).Value = rating?.Rating ?? "Missing";
                if (rating?.Conflict == true)
                {
                    ws.Cell(historyRow, index + 2).Style.Fill.BackgroundColor = WarningFill;
                }
            }
            historyRow++;
        }

        ws.Cell("A7").Value = "Health commentary";
        ws.Cell("B7").Value = summary.HealthCommentary;
        ws.Range("B7:G7").Merge();
        ws.Cell("A8").Value = "Dental commentary";
        ws.Cell("B8").Value = summary.DentalCommentary;
        ws.Range("B8:G8").Merge();
        ws.Cell("A9").Value = "Client takeaway";
        ws.Cell("B9").Value = summary.Takeaway;
        ws.Range("B9:G9").Merge();
        ws.Range("B7:G9").Style.Alignment.WrapText = true;

        Title(ws, "A11", "AssurePLUS pricing history");
        var pricingRows = analysis.Pdf.AssurePlusPricingReports
            .OrderBy(report => report.EffectiveDate)
            .SelectMany(report => report.Benefits.Select(row => new
            {
                report.EffectiveDate,
                Row = row
            }))
            .ToList();
        WriteRows(
            ws,
            ["Effective Date", "Renewal Year", "Benefit", "Previous Status", "Recent Status", "Standard Change", "Experience Adjustment", "Reported Total", "Calculated Total", "Variance", "Check", "Source"],
            pricingRows.Select(item => new object?[]
            {
                item.EffectiveDate?.ToString("MMMM d, yyyy", CultureInfo.InvariantCulture),
                item.EffectiveDate?.Year,
                item.Row.Benefit,
                item.Row.PreviousYearRating,
                item.Row.RecentYearRating,
                ToExcelPercent(item.Row.StandardChangePercent),
                ToExcelPercent(item.Row.ExperienceAdjustmentPercent),
                ToExcelPercent(item.Row.TotalChangePercent),
                null,
                null,
                null,
                item.Row.SourceReference
            }),
            12);
        for (var index = 0; index < pricingRows.Count; index++)
        {
            var rowNumber = 13 + index;
            ws.Cell(rowNumber, 9).FormulaA1 = $"ROUND(F{rowNumber}+G{rowNumber},4)";
            ws.Cell(rowNumber, 10).FormulaA1 = $"ROUND(H{rowNumber}-I{rowNumber},4)";
            ws.Cell(rowNumber, 11).FormulaA1 = $"IF(ABS(J{rowNumber})<=0.0001,\"OK\",\"Review\")";
        }
        if (pricingRows.Count > 0)
        {
            ws.Range(13, 6, 12 + pricingRows.Count, 10).Style.NumberFormat.Format = "+0.00%;-0.00%;0.00%";
        }

        var distributionTitleRow = 15 + pricingRows.Count;
        Title(ws, distributionTitleRow, 1, "Claims distribution detail");
        var distributionHeaderRow = distributionTitleRow + 1;
        var distributionRows = analysis.Pdf.AssurePlusPricingReports
            .OrderBy(report => report.EffectiveDate)
            .SelectMany(report => report.DistributionRows.Select(row => new
            {
                report.EffectiveDate,
                Row = row
            }))
            .ToList();
        WriteRows(
            ws,
            ["Renewal Year", "Benefit", "Category", "Your Claims", "Benchmark Claims", "Source"],
            distributionRows.Select(item => new object?[]
            {
                item.EffectiveDate?.Year,
                item.Row.Benefit,
                item.Row.Category,
                ToExcelPercent(item.Row.YourClaimsPercent),
                ToExcelPercent(item.Row.BenchmarkClaimsPercent),
                item.Row.SourceReference
            }),
            distributionHeaderRow);
        if (distributionRows.Count > 0)
        {
            ws.Range(distributionHeaderRow + 1, 4, distributionHeaderRow + distributionRows.Count, 5)
                .Style.NumberFormat.Format = "0.00%";
        }

        var warningTitleRow = distributionHeaderRow + distributionRows.Count + 3;
        Title(ws, warningTitleRow, 1, "Validation warnings");
        if (summary.Warnings.Count == 0)
        {
            ws.Cell(warningTitleRow + 1, 1).Value = "None - duplicate-year statuses and rate arithmetic reconciled.";
        }
        else
        {
            WriteRows(
                ws,
                ["Warning"],
                summary.Warnings.Select(warning => new object?[] { warning }),
                warningTitleRow + 1);
            ws.Range(warningTitleRow + 1, 1, warningTitleRow + summary.Warnings.Count + 1, 1)
                .Style.Fill.BackgroundColor = WarningFill;
        }

    }

    private static void WriteRows(IXLWorksheet ws, string[] headers, IEnumerable<object?[]> rows, int startRow = 1)
    {
        for (var col = 0; col < headers.Length; col++)
        {
            ws.Cell(startRow, col + 1).Value = headers[col];
        }

        Header(ws.Range(startRow, 1, startRow, headers.Length));

        var rowIndex = startRow + 1;
        foreach (var row in rows)
        {
            for (var col = 0; col < row.Length; col++)
            {
                SetValue(ws.Cell(rowIndex, col + 1), row[col]);
            }

            rowIndex++;
        }

        if (rowIndex > startRow + 1)
        {
            ws.Range(startRow, 1, rowIndex - 1, headers.Length).SetAutoFilter();
        }
    }

    private static void SetValue(IXLCell cell, object? value)
    {
        switch (value)
        {
            case null:
                cell.Value = "";
                break;
            case decimal decimalValue:
                cell.Value = decimalValue;
                break;
            case int intValue:
                cell.Value = intValue;
                break;
            case DateTime dateTime:
                cell.Value = dateTime;
                break;
            default:
                cell.Value = value.ToString();
                break;
        }
    }

    private static void SetMoney(IXLCell cell, decimal? value)
    {
        if (value.HasValue)
        {
            cell.Value = value.Value;
            cell.Style.NumberFormat.Format = "$#,##0.00";
        }
    }

    private static void SetMoneyWhole(IXLCell cell, decimal? value)
    {
        if (value.HasValue)
        {
            cell.Value = value.Value;
            cell.Style.NumberFormat.Format = "$#,##0";
        }
    }

    private static void SetPercentValue(IXLCell cell, decimal? value)
    {
        if (value.HasValue)
        {
            cell.Value = value.Value;
            cell.Style.NumberFormat.Format = "0.00%";
        }
    }

    private static void SetPercentNumber(IXLCell cell, decimal? value)
    {
        if (value.HasValue)
        {
            cell.Value = value.Value / 100m;
            cell.Style.NumberFormat.Format = "0.0%";
        }
    }

    private static decimal? ToExcelPercent(decimal? value)
    {
        return value.HasValue ? value.Value / 100m : null;
    }

    private static string FirstNonBlank(params string[] values)
    {
        return values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)) ?? "";
    }

    private static void FormatMoneyColumns(IXLWorksheet ws, params int[] columns)
    {
        foreach (var column in columns)
        {
            ws.Column(column).Style.NumberFormat.Format = "$#,##0.00";
        }
    }

    private static void FormatNumberColumns(IXLWorksheet ws, params int[] columns)
    {
        foreach (var column in columns)
        {
            ws.Column(column).Style.NumberFormat.Format = "#,##0.0";
        }
    }

    private static void Title(IXLWorksheet ws, string address, string title)
    {
        var cell = ws.Cell(address);
        StyleTitleCell(cell, title);
    }

    private static void Title(IXLWorksheet ws, int row, int column, string title)
    {
        StyleTitleCell(ws.Cell(row, column), title);
    }

    private static void StyleTitleCell(IXLCell cell, string title)
    {
        cell.Value = title;
        cell.Style.Font.Bold = true;
        cell.Style.Font.FontSize = 16;
        cell.Style.Font.FontColor = WestlandTeal;
    }

    private static void Header(IXLRange range)
    {
        range.Style.Font.Bold = true;
        range.Style.Font.FontColor = XLColor.White;
        range.Style.Fill.BackgroundColor = WestlandTeal;
        range.Style.Alignment.WrapText = true;
    }

    private sealed record HealthFormulaRateClass(
        string AccountLabel,
        Dictionary<string, HealthFormulaTierRate> RatesByTier)
    {
        public static HealthFormulaRateClass Empty { get; } = new("", new Dictionary<string, HealthFormulaTierRate>(StringComparer.OrdinalIgnoreCase));
    }

    private sealed record FormattedLineRateSourceRow(
        int RowNumber,
        string AccountClass,
        string Benefit,
        string TierClass,
        string RateBasis);

    private sealed record FormattedLineRateGroupKey(
        string Key,
        string DisplayName,
        int SortOrder,
        FormattedLineRateGroupKind Kind);

    private sealed record FormattedLineRateGroup(
        string Key,
        string DisplayName,
        int SortOrder,
        FormattedLineRateGroupKind Kind,
        int FirstIndex,
        List<FormattedLineRateSourceRow> Rows);

    private readonly record struct FormulaPeriodKey(DateOnly Start, DateOnly End);

    private sealed record CanadaLifeComponentSheet(
        string SheetName,
        string Title,
        IReadOnlyList<ExperienceLine> Rows);

    private sealed record CanadaLifeExperienceReconciliation(
        IReadOnlyList<ExperienceLine> HealthRows,
        IReadOnlyList<ExperienceLine> DrugRows,
        IReadOnlyList<ExperienceLine> VisionRows,
        List<ExperienceLine> CombinedRows,
        IReadOnlyList<FormulaPeriodKey> UnmatchedHealthPeriods,
        IReadOnlyList<FormulaPeriodKey> UnmatchedDrugPeriods,
        IReadOnlyList<FormulaPeriodKey> UnmatchedVisionPeriods);

    private enum FormattedLineRateGroupKind
    {
        Pooled,
        ExtendedHealth,
        Dental,
        Other
    }

    private sealed record HealthFormulaTierRate(
        decimal? CurrentRate,
        decimal? RenewalRate,
        decimal CurrentPremium,
        decimal RenewalPremium);

    private static int? RateVolumeDivisor(string benefit, string? rateBasis = null)
    {
        return UnderwritingEngine.RateVolumeDivisor(benefit, rateBasis);
    }

    private static List<RateLine> SelectedRateLines(AnalysisResult analysis)
    {
        var version = analysis.SelectedRenewalVersion;
        if (version?.RateLines.Count > 0)
        {
            return BenefitDisplayOrder.Sort(
                version.RateLines.Select(ToRateLine),
                line => line.Benefit,
                line => line.TierOrClass,
                line => line.AccountClass);
        }

        return BenefitDisplayOrder.Sort(
            analysis.Pdf.RateLines,
            line => line.Benefit,
            line => line.TierOrClass,
            line => line.AccountClass);
    }

    private static RateLine ToRateLine(RenewalVersionRateLine line)
    {
        var currentPremium = line.CurrentPremium;
        var proposedPremium = line.ProposedPremium;
        var negotiatedPremium = line.VersionPremium ?? proposedPremium;
        var currentRate = ClientPresentationRateSelector.CurrentRate(line);
        var proposedRate = ClientPresentationRateSelector.ProposedRate(line);
        var versionRate = ClientPresentationRateSelector.VersionRate(line);
        var canCalculate = !line.ValidationStatus.Equals(RateLineValidationStatuses.NeedsReview, StringComparison.OrdinalIgnoreCase);
        var currentCalculated = canCalculate ? UnderwritingEngine.CalculatePremium(line.Benefit, line.Volume, line.Lives, currentRate, line.RateBasis) : null;
        var proposedCalculated = canCalculate ? UnderwritingEngine.CalculatePremium(line.Benefit, line.Volume, line.Lives, proposedRate, line.RateBasis) : null;

        return new RateLine
        {
            Benefit = line.Benefit,
            TierOrClass = line.TierOrClass,
            AccountClass = line.AccountClass,
            Volume = line.Volume,
            Lives = line.Lives,
            CurrentRate = currentRate,
            CurrentPremium = currentPremium,
            RenewalRate = proposedRate,
            RenewalPremium = proposedPremium,
            CarrierProposedRate = proposedRate,
            CarrierProposedPremium = proposedPremium,
            NegotiatedRate = versionRate,
            NegotiatedPremium = negotiatedPremium,
            PercentChange = UnderwritingEngine.CalculatePercentChange(currentPremium, proposedPremium),
            CalculatedCurrentPremium = currentCalculated,
            CalculatedRenewalPremium = proposedCalculated,
            CurrentVariance = currentPremium.HasValue && currentCalculated.HasValue
                ? decimal.Round(currentPremium.Value - currentCalculated.Value, 2)
                : null,
            RenewalVariance = proposedPremium.HasValue && proposedCalculated.HasValue
                ? decimal.Round(proposedPremium.Value - proposedCalculated.Value, 2)
                : null,
            SourceReference = line.SourceReference,
            RateBasis = line.RateBasis,
            IncludeInPremiumTotal = line.IncludeInPremiumTotal,
            SourceRowRole = line.SourceRowRole,
            ValidationStatus = line.ValidationStatus,
            ValidationMessage = line.ValidationMessage,
            SourceCarrier = line.SourceCarrier,
            SourceFileName = line.SourceFileName,
            SourcePage = line.SourcePage,
            SourceText = line.SourceText,
            CurrentRatePrecision = line.CurrentRatePrecision,
            RenewalRatePrecision = line.RenewalRatePrecision,
            CurrentPeriodLabel = line.CurrentPeriodLabel,
            RenewalPeriodLabel = line.RenewalPeriodLabel,
            RenewalEffectiveDate = line.RenewalEffectiveDate
        };
    }

    private static void StyleSeverity(IXLRow row, FindingSeverity severity)
    {
        row.Style.Fill.BackgroundColor = severity switch
        {
            FindingSeverity.Critical => CriticalFill,
            FindingSeverity.Warning => WarningFill,
            _ => XLColor.White
        };

        if (severity == FindingSeverity.Info)
        {
            row.Cell(1).Style.Font.FontColor = WestlandTeal;
        }
        else if (severity == FindingSeverity.Warning)
        {
            row.Cell(1).Style.Font.FontColor = WestlandGold;
        }
        else
        {
            row.Cell(1).Style.Font.FontColor = XLColor.DarkRed;
        }
    }
}

public sealed record PowerPointExportTable(string Title, string[] Headers, List<string[]> Rows);

public static class PowerPointRateTableBuilder
{
    public const string PremiumIllustrationTableTitle = "Prem Illusn";

    private static readonly string[] RateTableMasterHeaders =
    [
        "Benefit",
        "Current Premium",
        "Proposed Premium",
        "Change (%)"
    ];

    private static readonly string[] TopDrugsHeaders =
    [
        "",
        "Drug Name",
        "Treatment",
        "Total Amount Paid"
    ];

    private static readonly string[] RateLineCalculationHeaders =
    [
        "Account/Class",
        "Benefit",
        "Tier/Class",
        "Volume",
        "Lives",
        "Current Rate",
        "Current Premium",
        "Renewal Rate",
        "Renewal Premium",
        "Negotiated Rate",
        "Negotiated Renewal Premium"
    ];

    private static readonly string[] FormattedRateLineGroupHeaders =
    [
        "",
        "",
        "Current Plan",
        "",
        "Calculated",
        "",
        "Proposed",
        "",
        "Negotiated",
        ""
    ];

    private static readonly string[] FormattedRateLineColumnHeaders =
    [
        "Benefit / Class",
        "Volume",
        "Rate",
        "Premium",
        "Rate",
        "Premium",
        "Rate",
        "Premium",
        "Rate",
        "Premium"
    ];

    private static readonly string[] PremiumOverview3Headers =
    [
        "Benefit",
        "Current Premium",
        "Proposed Premium",
        "Change (%)",
        "Negotiated Premium",
        "Change (%)"
    ];

    private static readonly string[] NemMaximumHeaders =
    [
        "Benefit",
        "Current",
        "Available"
    ];

    public static List<PowerPointExportTable> BuildAllTables(AnalysisResult analysis)
    {
        var tables = new List<PowerPointExportTable>
        {
            BuildRateTableMaster(analysis),
            BuildPremiumOverviewTable(analysis),
            BuildPremiumOverview2Table(analysis),
            BuildRatesOverviewTable(analysis),
            BuildPremiumIllustrationTable(analysis),
            BuildPremiumOverview3Table(analysis),
            BuildHealthDentalBreakdownTable(analysis),
            BuildDivisionsTable(analysis),
            BuildRateLineCalculationTable(analysis),
            BuildFormattedRateLineCalculationTable(analysis, "RateTable_LineRates_Formatted"),
            BuildNemMaximumsTable(analysis),
            BuildTopDrugsTable(analysis),
            BuildPriorTopDrugsTable(analysis)
        };

        tables.AddRange(BuildSingleDivisionTables(analysis));
        return tables;
    }

    public static Dictionary<string, PowerPointExportTable> BuildPlaceholderTables(AnalysisResult analysis)
    {
        var rateTableMaster = BuildRateTableMaster(analysis, omitZeroPremiumRows: true);
        var premiumOverviewTable = BuildPremiumOverviewTable(analysis, omitZeroPremiumRows: true);
        var premiumOverview2Table = BuildPremiumOverview2Table(analysis, omitZeroPremiumRows: true);
        var ratesOverviewTable = BuildRatesOverviewTable(analysis, omitZeroPremiumRows: true);
        var premiumIllustrationTable = BuildPremiumIllustrationTable(analysis, omitZeroPremiumRows: true);
        var premiumOverview3Table = BuildPremiumOverview3Table(analysis, omitZeroPremiumRows: true);
        var healthDentalBreakdownTable = BuildHealthDentalBreakdownTable(analysis, omitZeroPremiumRows: true);
        var divisionsTable = BuildDivisionsTable(analysis, omitZeroPremiumRows: true);
        var rateLineCalculationTable = BuildRateLineCalculationTable(analysis, omitZeroPremiumRows: true);
        var formattedRateLineCalculationTable = BuildFormattedRateLineCalculationTable(
            analysis,
            "RateTable_LineRates_Formatted",
            omitZeroPremiumRows: true);
        var nemMaximumsTable = BuildNemMaximumsTable(analysis);
        var topDrugsTable = BuildTopDrugsTable(analysis);
        var priorTopDrugsTable = BuildPriorTopDrugsTable(analysis);
        var tables = new Dictionary<string, PowerPointExportTable>(StringComparer.Ordinal)
        {
            ["<<RateTable_Master>>"] = rateTableMaster,
            ["<<RateTable_Divisions>>"] = divisionsTable,
            ["<<PremiumOverviewTable>>"] = premiumOverviewTable,
            ["<<PremiumOverview_Master>>"] = premiumOverviewTable,
            ["<<PremiumOverview2>>"] = premiumOverview2Table,
            ["<<PremiumOverview2Table>>"] = premiumOverview2Table,
            ["<<PremiumOverview_2>>"] = premiumOverview2Table,
            ["<<PremiumOverview_2_Table>>"] = premiumOverview2Table,
            ["<<PremiumOververview2>>"] = premiumOverview2Table,
            ["<<PremiumOververview2Table>>"] = premiumOverview2Table,
            ["<<RatesOverview>>"] = ratesOverviewTable,
            ["<<RatesOverviewTable>>"] = ratesOverviewTable,
            ["<<RateOverview>>"] = ratesOverviewTable,
            ["<<RateOverviewTable>>"] = ratesOverviewTable,
            ["<<PremIllusn>>"] = premiumIllustrationTable,
            ["<<Prem Illusn>>"] = premiumIllustrationTable,
            ["<<Prem_Illusn>>"] = premiumIllustrationTable,
            ["<<PremiumIllustration>>"] = premiumIllustrationTable,
            ["<<PremiumIllustrationTable>>"] = premiumIllustrationTable,
            ["<<PremiumOverview3>>"] = premiumOverview3Table,
            ["<<PremiumOverview3Table>>"] = premiumOverview3Table,
            ["<<PremiumOverview_3>>"] = premiumOverview3Table,
            ["<<PremiumOverview_3_Table>>"] = premiumOverview3Table,
            ["<<RateTable_MasterBreakdown>>"] = premiumOverview2Table,
            ["<<RateTable_Master_Breakdown>>"] = premiumOverview2Table,
            ["<<HealthDentalBreakdown>>"] = healthDentalBreakdownTable,
            ["<<PremiumOverview2_Breakdown>>"] = healthDentalBreakdownTable,
            ["<<RateTable_HealthDentalBreakdown>>"] = healthDentalBreakdownTable,
            ["<<PremiumOverview_Divisions>>"] = divisionsTable,
            // Use <<RateTable_LineRates_Formatted>> for the polished client-facing
            // table. The original placeholder is kept as the raw export shape so
            // templates can choose explicitly instead of guessing which layout to use.
            ["<<RateTable_LineRates>>"] = rateLineCalculationTable,
            ["<<RateTable_LineRates_Formatted>>"] = formattedRateLineCalculationTable,
            ["<<RateTable_RateLines>>"] = formattedRateLineCalculationTable,
            ["<<RateTable_CalculationLines>>"] = formattedRateLineCalculationTable,
            ["<<RateTable_ByRate>>"] = formattedRateLineCalculationTable,
            ["<<RateTable_LineRates_Raw>>"] = rateLineCalculationTable,
            ["<<NEMs>>"] = nemMaximumsTable,
            ["<<NEMS>>"] = nemMaximumsTable,
            ["<<NEMTable>>"] = nemMaximumsTable,
            ["<<NonEvidenceMaximums>>"] = nemMaximumsTable,
            ["<<NonEvidenceMaximumsTable>>"] = nemMaximumsTable,
            ["<<PlanMaximums>>"] = nemMaximumsTable,
            ["<<PlanMaximumsTable>>"] = nemMaximumsTable,
            ["<<CurrentTopDrugsTable>>"] = topDrugsTable,
            ["<<CurrentTopDrugs>>"] = topDrugsTable,
            ["<<CurrentTop10Drugs>>"] = topDrugsTable,
            ["<<CurrentTop_10_DINs>>"] = topDrugsTable,
            ["<<TopDrugsTable>>"] = topDrugsTable,
            ["<<TopDrugs>>"] = topDrugsTable,
            ["<<Top10Drugs>>"] = topDrugsTable,
            ["<<Top_10_Drugs>>"] = topDrugsTable,
            ["<<Top_10_DINs>>"] = topDrugsTable,
            ["<<Top_10_DINS>>"] = topDrugsTable,
            ["<<PriorTopDrugsTable>>"] = priorTopDrugsTable,
            ["<<PriorTopDrugs>>"] = priorTopDrugsTable,
            ["<<PriorTop10Drugs>>"] = priorTopDrugsTable,
            ["<<PriorTop_10_DINs>>"] = priorTopDrugsTable
        };

        var divisionTables = BuildSingleDivisionTables(analysis, omitZeroPremiumRows: true).ToList();
        for (var index = 0; index < divisionTables.Count; index++)
        {
            tables[$"<<RateTable_Division{index + 1}>>"] = divisionTables[index];
            tables[$"<<RateTable_Division_{index + 1}>>"] = divisionTables[index];
        }

        return tables;
    }

    public static PowerPointExportTable BuildTopDrugsTable(AnalysisResult analysis)
    {
        return BuildTopDrugsTable("CurrentTopDrugs", analysis.Workbook.DrugTrends);
    }

    public static PowerPointExportTable BuildPriorTopDrugsTable(AnalysisResult analysis)
    {
        return BuildTopDrugsTable("PriorTopDrugs", analysis.Workbook.PriorDrugTrends);
    }

    public static PowerPointExportTable BuildNemMaximumsTable(AnalysisResult analysis)
    {
        return new PowerPointExportTable(
            "NEMs",
            NemMaximumHeaders,
            [
                BuildNemMaximumRow(analysis, "Life Insurance"),
                BuildNemMaximumRow(analysis, "Long-Term Disability")
            ]);
    }

    private static string[] BuildNemMaximumRow(AnalysisResult analysis, string benefit)
    {
        var row = FindNemMaximum(analysis, benefit);
        return
        [
            benefit,
            row?.CurrentNem ?? "",
            row?.AvailableNem ?? ""
        ];
    }

    private static NemMaximumRow? FindNemMaximum(AnalysisResult analysis, string benefit)
    {
        var normalized = NemMaximumExtractor.NormalizeBenefitName(benefit);
        return analysis.NemMaximums.FirstOrDefault(row =>
            NemMaximumExtractor.NormalizeBenefitName(row.Benefit).Equals(normalized, StringComparison.OrdinalIgnoreCase));
    }

    private static PowerPointExportTable BuildTopDrugsTable(string title, IEnumerable<DrugTrend> drugs)
    {
        var rows = SelectDisplayDrugPeriod(drugs)
            .Where(IsDrugDetailRow)
            .OrderByDescending(drug => drug.NetAmountPaid ?? drug.SubmittedTotal ?? 0m)
            .Take(10)
            .Select((drug, index) => new[]
            {
                (index + 1).ToString(),
                drug.DrugName,
                drug.Treats,
                ParsingHelpers.Money(drug.NetAmountPaid ?? drug.SubmittedTotal)
            })
            .ToList();

        while (rows.Count < 10)
        {
            rows.Add([(rows.Count + 1).ToString(), "", "", ""]);
        }

        return new PowerPointExportTable(title, TopDrugsHeaders, rows);
    }

    private static IEnumerable<DrugTrend> SelectDisplayDrugPeriod(IEnumerable<DrugTrend> drugs)
    {
        var rows = drugs
            .Where(IsDrugDetailRow)
            .ToList();
        var periodGroups = rows
            .Where(drug => !string.IsNullOrWhiteSpace(drug.Period))
            .GroupBy(drug => drug.Period, StringComparer.OrdinalIgnoreCase)
            .Select(group => new
            {
                PeriodEnd = ParseDrugPeriodEnd(group.Key),
                Rows = group.ToList()
            })
            .ToList();

        if (periodGroups.Count == 0)
        {
            return rows;
        }

        return periodGroups
            .OrderByDescending(group => group.PeriodEnd ?? DateOnly.MinValue)
            .ThenByDescending(group => group.Rows.Count)
            .First()
            .Rows;
    }

    private static DateOnly? ParseDrugPeriodEnd(string period)
    {
        var separator = period.Contains(" to ", StringComparison.OrdinalIgnoreCase) ? " to " : " - ";
        var parts = period.Split(separator, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parts.Length != 2)
        {
            return null;
        }

        if (DateOnly.TryParse(parts[1], CultureInfo.InvariantCulture, DateTimeStyles.None, out var endDate))
        {
            return endDate;
        }

        foreach (var format in new[] { "MMM yyyy", "MMMM yyyy" })
        {
            if (DateTime.TryParseExact(parts[1], format, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out var monthEnd))
            {
                return new DateOnly(monthEnd.Year, monthEnd.Month, DateTime.DaysInMonth(monthEnd.Year, monthEnd.Month));
            }
        }

        return null;
    }

    private static bool IsDrugDetailRow(DrugTrend drug)
    {
        if (string.IsNullOrWhiteSpace(drug.DrugName))
        {
            return false;
        }

        if (drug.DrugName.Equals("Grand Total", StringComparison.OrdinalIgnoreCase) ||
            drug.DrugName.Contains("Top 50", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        return (drug.NetAmountPaid ?? drug.SubmittedTotal ?? 0m) > 0m;
    }

    private static PowerPointExportTable BuildRateTableMaster(AnalysisResult analysis, bool omitZeroPremiumRows = false)
    {
        var powerPointLines = BuildPowerPointRateLines(analysis, omitZeroPremiumRows);
        var rows = BuildMasterAggregates(powerPointLines)
            .Select(ToSimpleRateTableRow)
            .ToList();

        var currentTotal = powerPointLines.Sum(line => line.CurrentPremium ?? 0m);
        var renewalTotal = powerPointLines.Sum(line => line.RenewalPremium ?? 0m);
        rows.Add(["Monthly Total", ParsingHelpers.Money(currentTotal), ParsingHelpers.Money(renewalTotal), ParsingHelpers.Percent(PercentChange(currentTotal, renewalTotal))]);

        return new PowerPointExportTable("RateTable_Master", RateTableMasterHeaders, rows);
    }

    private static PowerPointExportTable BuildPremiumOverviewTable(AnalysisResult analysis, bool omitZeroPremiumRows = false)
    {
        var powerPointLines = BuildPowerPointRateLines(analysis, omitZeroPremiumRows);
        var rows = BuildMasterAggregates(powerPointLines)
            .Select(ToPremiumOverviewRow)
            .ToList();

        var currentTotal = powerPointLines.Sum(line => line.CurrentPremium ?? 0m);
        var renewalTotal = powerPointLines.Sum(line => line.RenewalPremium ?? 0m);
        AddPremiumTotals(rows, currentTotal, renewalTotal);

        return new PowerPointExportTable("PremiumOverviewTable", MasterHeaders(analysis), rows);
    }

    private static PowerPointExportTable BuildPremiumOverview2Table(AnalysisResult analysis, bool omitZeroPremiumRows = false)
    {
        var powerPointLines = BuildPowerPointRateLines(analysis, omitZeroPremiumRows);
        var rows = new List<string[]>();
        foreach (var aggregate in BuildMasterAggregates(powerPointLines))
        {
            if (IsExtendedHealthBenefit(aggregate.Benefit))
            {
                rows.AddRange(BuildHealthDentalBreakdownRows(analysis, HealthDentalBenefit.ExtendedHealthCare, includeEmptyFallback: true, omitZeroPremiumRows: omitZeroPremiumRows));
                continue;
            }

            if (IsDentalBenefit(aggregate.Benefit))
            {
                rows.AddRange(BuildHealthDentalBreakdownRows(analysis, HealthDentalBenefit.Dental, includeEmptyFallback: true, omitZeroPremiumRows: omitZeroPremiumRows));
                continue;
            }

            rows.Add(ToPremiumOverviewRow(aggregate));
        }

        var currentTotal = powerPointLines.Sum(line => line.CurrentPremium ?? 0m);
        var renewalTotal = powerPointLines.Sum(line => line.RenewalPremium ?? 0m);
        AddPremiumTotals(rows, currentTotal, renewalTotal);

        return new PowerPointExportTable("PremiumOverview2", MasterHeaders(analysis), rows);
    }

    private static PowerPointExportTable BuildRatesOverviewTable(AnalysisResult analysis, bool omitZeroPremiumRows = false)
    {
        var rows = new List<string[]>();
        var overviewGroups = BuildOverviewRateGroups(BuildPowerPointRateLines(analysis, omitZeroPremiumRows));
        foreach (var group in overviewGroups)
        {
            var lines = group.Lines;
            var benefit = AddOverviewAccountSuffix(FriendlyRatesOverviewBenefitName(lines[0].Benefit), group);

            if (lines.Any(line => IsExtendedHealthBenefit(line.Benefit) || IsDentalBenefit(line.Benefit)))
            {
                rows.Add(BuildTieredRatesOverviewRow(benefit, lines));
                continue;
            }

            rows.Add(BuildSingleRatesOverviewRow(benefit, lines));
        }

        var overviewLines = overviewGroups.SelectMany(group => group.Lines).ToList();
        var currentTotal = overviewLines.Sum(line => line.CurrentPremium ?? 0m);
        var proposedTotal = overviewLines.Sum(line => CarrierProposedPremium(line) ?? 0m);
        var negotiatedTotal = overviewLines.Sum(line => NegotiatedPremium(line) ?? 0m);
        rows.Add(
        [
            "Total Monthly Premiums",
            "",
            ParsingHelpers.Money(currentTotal),
            ParsingHelpers.Money(proposedTotal),
            ParsingHelpers.Money(negotiatedTotal),
            ""
        ]);
        rows.Add(
        [
            "Monthly Difference from Current",
            "",
            "",
            ParsingHelpers.Money(proposedTotal - currentTotal),
            ParsingHelpers.Money(negotiatedTotal - currentTotal),
            ""
        ]);
        rows.Add(
        [
            "Monthly % Difference from Current",
            "",
            "",
            ParsingHelpers.Percent(PercentChange(currentTotal, proposedTotal)),
            ParsingHelpers.Percent(PercentChange(currentTotal, negotiatedTotal)),
            ""
        ]);

        return new PowerPointExportTable("RatesOverview", RatesOverviewHeaders(analysis), rows);
    }

    private static PowerPointExportTable BuildPremiumIllustrationTable(AnalysisResult analysis, bool omitZeroPremiumRows = false)
    {
        var rows = new List<string[]>();
        var overviewGroups = BuildOverviewRateGroups(BuildPowerPointRateLines(analysis, omitZeroPremiumRows));
        foreach (var group in overviewGroups)
        {
            var lines = group.Lines;
            var benefit = PremiumIllustrationBenefitName(lines[0].Benefit);
            var accountClass = PremiumIllustrationClassName(group.AccountLabel);

            rows.Add(lines.Any(line => IsExtendedHealthBenefit(line.Benefit) || IsDentalBenefit(line.Benefit))
                ? BuildTieredPremiumIllustrationRow(benefit, accountClass, lines)
                : BuildSinglePremiumIllustrationRow(benefit, accountClass, lines));
        }

        var overviewLines = overviewGroups.SelectMany(group => group.Lines).ToList();
        var currentTotal = overviewLines.Sum(line => PremiumIllustrationCurrentPremium(line) ?? 0m);
        var proposedTotal = overviewLines.Sum(line => PremiumIllustrationProposedPremium(line) ?? 0m);
        var negotiatedTotal = overviewLines.Sum(line => PremiumIllustrationNegotiatedPremium(line) ?? 0m);
        rows.Add(
        [
            "Total Monthly Premiums",
            "",
            "",
            ParsingHelpers.Money(currentTotal),
            ParsingHelpers.Money(proposedTotal),
            ParsingHelpers.Money(negotiatedTotal),
            ""
        ]);
        rows.Add(
        [
            "Monthly Difference from Current",
            "",
            "",
            "",
            ParsingHelpers.Money(proposedTotal - currentTotal),
            ParsingHelpers.Money(negotiatedTotal - currentTotal),
            ""
        ]);
        rows.Add(
        [
            "Monthly % Difference from Current",
            "",
            "",
            "",
            ParsingHelpers.Percent(PercentChange(currentTotal, proposedTotal)),
            ParsingHelpers.Percent(PercentChange(currentTotal, negotiatedTotal)),
            ""
        ]);

        return new PowerPointExportTable(PremiumIllustrationTableTitle, PremiumIllustrationHeaders(analysis), rows);
    }

    private static string[] PremiumIllustrationHeaders(AnalysisResult analysis)
    {
        var renewalDate = FormatRenewalDate(analysis);
        return
        [
            "Benefit",
            "Class",
            "Insured Volume",
            "Current",
            $"{renewalDate}\nRenewal",
            $"{renewalDate}\nRenewal Negotiated",
            "% Change from Current"
        ];
    }

    private static string[] BuildSinglePremiumIllustrationRow(string benefit, string accountClass, List<RateLine> lines)
    {
        var currentPremium = SumNullable(lines.Select(PremiumIllustrationCurrentPremium));
        var proposedPremium = SumNullable(lines.Select(PremiumIllustrationProposedPremium));
        var negotiatedPremium = SumNullable(lines.Select(PremiumIllustrationNegotiatedPremium));
        var volume = lines.Any(line => line.Volume.HasValue) ? lines.Sum(line => line.Volume ?? 0m) : (decimal?)null;
        var lives = lines.Sum(line => line.Lives ?? 0);

        return
        [
            benefit,
            accountClass,
            FormatInsuredVolume(volume, lives == 0 ? null : lives),
            FormatPremiumCell(SummarizeRateCurrency(lines.Select(line => line.CurrentRate)), currentPremium),
            FormatPremiumCell(SummarizeRateCurrency(lines.Select(CarrierProposedRate)), proposedPremium),
            FormatPremiumCell(SummarizeRateCurrency(lines.Select(NegotiatedRate)), negotiatedPremium),
            ParsingHelpers.Percent(PercentChange(currentPremium, negotiatedPremium))
        ];
    }

    private static string[] BuildTieredPremiumIllustrationRow(string benefit, string accountClass, List<RateLine> lines)
    {
        var tiers = BuildPremiumIllustrationTiers(lines);
        if (tiers.Count == 0)
        {
            return BuildSinglePremiumIllustrationRow(benefit, accountClass, lines);
        }

        var currentPremium = SumNullable(tiers.Select(tier => tier.CurrentPremium));
        var proposedPremium = SumNullable(tiers.Select(tier => tier.ProposedPremium));
        var negotiatedPremium = SumNullable(tiers.Select(tier => tier.NegotiatedPremium));

        return
        [
            benefit,
            accountClass,
            string.Join("\n", tiers.Select(tier => $"{tier.Name}    {FormatTierVolume(tier.Volume, tier.Lives)}".TrimEnd())),
            FormatTieredPremiumIllustrationCell(tiers.Select(tier => tier.CurrentRate), currentPremium),
            FormatTieredPremiumIllustrationCell(tiers.Select(tier => tier.ProposedRate), proposedPremium),
            FormatTieredPremiumIllustrationCell(tiers.Select(tier => tier.NegotiatedRate), negotiatedPremium),
            ParsingHelpers.Percent(PercentChange(currentPremium, negotiatedPremium))
        ];
    }

    private static List<PremiumIllustrationTier> BuildPremiumIllustrationTiers(List<RateLine> lines)
    {
        return lines
            .GroupBy(line => NormalizeTier(line.TierOrClass))
            .Select(group =>
            {
                var tierLines = group.ToList();
                var currentRate = SumNullable(tierLines.Select(line => line.CurrentRate));
                var proposedRate = SumNullable(tierLines.Select(CarrierProposedRate));
                var negotiatedRate = SumNullable(tierLines.Select(NegotiatedRate));
                return new PremiumIllustrationTier(
                    TierSort(group.Key),
                    FriendlyTierName(group.Key),
                    tierLines.Sum(line => line.Lives ?? 0),
                    tierLines.Any(line => line.Volume.HasValue) ? tierLines.Sum(line => line.Volume ?? 0m) : null,
                    currentRate,
                    proposedRate,
                    negotiatedRate,
                    SumNullable(tierLines.Select(PremiumIllustrationCurrentPremium)),
                    SumNullable(tierLines.Select(PremiumIllustrationProposedPremium)),
                    SumNullable(tierLines.Select(PremiumIllustrationNegotiatedPremium)));
            })
            .Where(tier => tier.Lives != 0 ||
                           (tier.Volume ?? 0m) != 0m ||
                           tier.CurrentRate.HasValue ||
                           tier.ProposedRate.HasValue ||
                           tier.NegotiatedRate.HasValue ||
                           tier.CurrentPremium.HasValue ||
                           tier.ProposedPremium.HasValue ||
                           tier.NegotiatedPremium.HasValue)
            .OrderBy(tier => tier.SortOrder)
            .ThenBy(tier => tier.Name)
            .ToList();
    }

    private static string FormatTieredPremiumIllustrationCell(IEnumerable<decimal?> rates, decimal? premium)
    {
        var lines = rates.Select(FormatRateCurrency).ToList();
        var premiumText = ParsingHelpers.Money(premium);
        if (!string.IsNullOrWhiteSpace(premiumText))
        {
            lines.Add(premiumText);
        }

        return string.Join("\n", lines);
    }

    private static decimal? PremiumIllustrationCurrentPremium(RateLine line)
    {
        return line.CurrentPremium ?? CalculatePremium(line.Benefit, line.Volume, line.Lives, line.CurrentRate, line.RateBasis);
    }

    private static decimal? PremiumIllustrationProposedPremium(RateLine line)
    {
        return CarrierProposedPremium(line) ?? CalculatePremium(line.Benefit, line.Volume, line.Lives, CarrierProposedRate(line), line.RateBasis);
    }

    private static decimal? PremiumIllustrationNegotiatedPremium(RateLine line)
    {
        return NegotiatedPremium(line) ?? CalculatePremium(line.Benefit, line.Volume, line.Lives, NegotiatedRate(line), line.RateBasis);
    }

    private static PowerPointExportTable BuildPremiumOverview3Table(AnalysisResult analysis, bool omitZeroPremiumRows = false)
    {
        var rows = new List<string[]>();
        var currentTotal = 0m;
        var proposedTotal = 0m;
        var negotiatedTotal = 0m;

        foreach (var group in BuildOverviewRateGroups(BuildPowerPointRateLines(analysis, omitZeroPremiumRows)))
        {
            var lines = group.Lines;
            var current = lines.Sum(line => line.CurrentPremium ?? 0m);
            var proposed = lines.Sum(line => CarrierProposedPremium(line) ?? 0m);
            var negotiated = lines.Sum(line =>
                NegotiatedPremium(line) ??
                CalculatePremium(line.Benefit, line.Volume, line.Lives, NegotiatedRate(line), line.RateBasis) ?? 0m);

            currentTotal += current;
            proposedTotal += proposed;
            negotiatedTotal += negotiated;

            rows.Add(
            [
                AddOverviewAccountSuffix(FriendlyPremiumOverviewBenefitName(lines[0].Benefit), group),
                ParsingHelpers.Money(current * 12m),
                ParsingHelpers.Money(proposed * 12m),
                ParsingHelpers.Percent(PercentChange(current, proposed)),
                ParsingHelpers.Money(negotiated * 12m),
                ParsingHelpers.Percent(PercentChange(current, negotiated))
            ]);
        }

        rows.Add(
        [
            "Annual Total",
            ParsingHelpers.Money(currentTotal * 12m),
            ParsingHelpers.Money(proposedTotal * 12m),
            ParsingHelpers.Percent(PercentChange(currentTotal, proposedTotal)),
            ParsingHelpers.Money(negotiatedTotal * 12m),
            ParsingHelpers.Percent(PercentChange(currentTotal, negotiatedTotal))
        ]);

        return new PowerPointExportTable("PremiumOverview3", PremiumOverview3Headers, rows);
    }

    private static PowerPointExportTable BuildHealthDentalBreakdownTable(AnalysisResult analysis, bool omitZeroPremiumRows = false)
    {
        var rows = new List<string[]>();
        rows.AddRange(BuildHealthDentalBreakdownRows(analysis, HealthDentalBenefit.ExtendedHealthCare, includeEmptyFallback: false, omitZeroPremiumRows: omitZeroPremiumRows));
        rows.AddRange(BuildHealthDentalBreakdownRows(analysis, HealthDentalBenefit.Dental, includeEmptyFallback: false, omitZeroPremiumRows: omitZeroPremiumRows));

        return new PowerPointExportTable("HealthDentalBreakdown", MasterHeaders(analysis), rows);
    }

    private static string[] RatesOverviewHeaders(AnalysisResult analysis)
    {
        var renewalDate = FormatShortRenewalDate(analysis);
        return
        [
            "Benefit",
            "Insured Volumes",
            "Current",
            $"{renewalDate}\nRenewal Proposed",
            $"{renewalDate}\nRenewal Negotiated",
            "% Change from Current"
        ];
    }

    private static string[] BuildSingleRatesOverviewRow(string benefit, List<RateLine> lines)
    {
        var current = lines.Sum(line => line.CurrentPremium ?? 0m);
        var proposed = lines.Sum(line => CarrierProposedPremium(line) ?? 0m);
        var negotiated = lines.Sum(line => NegotiatedPremium(line) ?? 0m);
        var volume = lines.Any(line => line.Volume.HasValue) ? lines.Sum(line => line.Volume ?? 0m) : (decimal?)null;
        var lives = lines.Sum(line => line.Lives ?? 0);
        var proposedRate = SummarizeRateCurrency(lines.Select(CarrierProposedRate));
        var negotiatedRate = SummarizeRateCurrency(lines.Select(NegotiatedRate));

        return
        [
            benefit,
            FormatInsuredVolume(volume, lives == 0 ? null : lives),
            SummarizeRateCurrency(lines.Select(line => line.CurrentRate)),
            proposedRate,
            negotiatedRate,
            ParsingHelpers.Percent(PercentChange(current, negotiated))
        ];
    }

    private static string[] BuildTieredRatesOverviewRow(string benefit, List<RateLine> lines)
    {
        var tiers = lines
            .GroupBy(line => new
            {
                Account = string.IsNullOrWhiteSpace(line.AccountClass) ? line.SourceReference : line.AccountClass,
                Tier = NormalizeTier(line.TierOrClass)
            })
            .Select(group =>
            {
                var tierLines = group.ToList();
                var countLivesOnce = tierLines.Count > 1 && tierLines.Any(line => IsExtendedHealthBenefit(line.Benefit));
                var currentRate = tierLines.Sum(line => line.CurrentRate ?? 0m);
                var proposedRate = tierLines.Sum(line => CarrierProposedRate(line) ?? 0m);
                var negotiatedRate = tierLines.Sum(line => NegotiatedRate(line) ?? 0m);
                return new RatesOverviewTier(
                    TierSort(group.Key.Tier),
                    FriendlyTierName(group.Key.Tier),
                    countLivesOnce ? tierLines.Select(line => line.Lives ?? 0).DefaultIfEmpty(0).Max() : tierLines.Sum(line => line.Lives ?? 0),
                    countLivesOnce && tierLines.Any(line => line.Volume.HasValue)
                        ? tierLines.Select(line => line.Volume ?? 0m).DefaultIfEmpty(0m).Max()
                        : tierLines.Any(line => line.Volume.HasValue) ? tierLines.Sum(line => line.Volume ?? 0m) : (decimal?)null,
                    currentRate == 0m && !tierLines.Any(line => line.CurrentRate.HasValue) ? null : currentRate,
                    proposedRate == 0m && !tierLines.Any(line => CarrierProposedRate(line).HasValue) ? null : proposedRate,
                    negotiatedRate == 0m && !tierLines.Any(line => NegotiatedRate(line).HasValue) ? null : negotiatedRate,
                    tierLines.Sum(line => line.CurrentPremium ?? 0m),
                    tierLines.Sum(line => CarrierProposedPremium(line) ?? 0m),
                    tierLines.Sum(line => NegotiatedPremium(line) ?? 0m));
            })
            .GroupBy(tier => tier.Name)
            .Select(group =>
            {
                var tierItems = group.ToList();
                return new RatesOverviewTier(
                    tierItems.Min(tier => tier.SortOrder),
                    group.Key,
                    tierItems.Sum(tier => tier.Lives),
                    tierItems.Any(tier => tier.Volume.HasValue) ? tierItems.Sum(tier => tier.Volume ?? 0m) : (decimal?)null,
                    CollapseRate(tierItems.Select(tier => tier.CurrentRate)),
                    CollapseRate(tierItems.Select(tier => tier.ProposedRate)),
                    CollapseRate(tierItems.Select(tier => tier.NegotiatedRate)),
                    tierItems.Sum(tier => tier.CurrentPremium),
                    tierItems.Sum(tier => tier.ProposedPremium),
                    tierItems.Sum(tier => tier.NegotiatedPremium));
            })
            .Where(tier => tier.Lives != 0 || (tier.Volume ?? 0m) != 0m || tier.CurrentRate.HasValue || tier.ProposedRate.HasValue || tier.NegotiatedRate.HasValue)
            .OrderBy(tier => tier.SortOrder)
            .ThenBy(tier => tier.Name)
            .ToList();

        if (tiers.Count == 0)
        {
            return BuildSingleRatesOverviewRow(benefit, lines);
        }

        var volumeText = string.Join("\n", tiers.Select(tier => $"{tier.Name}    {FormatTierVolume(tier.Volume, tier.Lives)}".TrimEnd()));
        var currentRateText = string.Join("\n", tiers.Select(tier => FormatRateCurrency(tier.CurrentRate)));
        var proposedRateText = string.Join("\n", tiers.Select(tier => FormatRateCurrency(tier.ProposedRate)));
        var negotiatedRateText = string.Join("\n", tiers.Select(tier => FormatRateCurrency(tier.NegotiatedRate)));
        var changeText = string.Join("\n", tiers.Select(tier => ParsingHelpers.Percent(PercentChange(tier.CurrentPremium, tier.NegotiatedPremium))));

        return
        [
            benefit,
            volumeText,
            currentRateText,
            proposedRateText,
            negotiatedRateText,
            changeText
        ];
    }

    private static List<OverviewRateGroup> BuildOverviewRateGroups(IEnumerable<RateLine> rateLines)
    {
        var result = new List<OverviewRateGroup>();
        foreach (var benefitGroup in OrderedBenefitGroups(rateLines))
        {
            var selectedLines = SelectOverviewRateLines(benefitGroup.Select(item => item.Line).ToList());
            var selectedItems = selectedLines
                .Select(line => new OrderedRateLine(
                    benefitGroup.First(item => ReferenceEquals(item.Line, line)).Index,
                    line))
                .ToList();

            if (selectedItems.Count == 0)
            {
                continue;
            }

            var accountGroups = selectedItems
                .GroupBy(item => OverviewAccountKey(item.Line))
                .OrderBy(group => group.Min(item => item.Index))
                .ToList();
            var splitByAccount = accountGroups.Count > 1;
            var fallbackClassNumber = 1;

            foreach (var accountGroup in accountGroups)
            {
                var accountLabel = OverviewAccountLabel(accountGroup.First().Line);
                if (string.IsNullOrWhiteSpace(accountLabel) && splitByAccount)
                {
                    accountLabel = $"Class {fallbackClassNumber}";
                }

                result.Add(new OverviewRateGroup(
                    accountGroup.Min(item => item.Index),
                    accountLabel,
                    splitByAccount,
                    accountGroup.Select(item => item.Line).ToList()));
                fallbackClassNumber++;
            }
        }

        return result.OrderBy(group => group.Index).ToList();
    }

    private static List<RateLine> SelectOverviewRateLines(List<RateLine> lines)
    {
        if (lines.Any(line => IsExtendedHealthBenefit(line.Benefit)))
        {
            var totalLines = lines
                .Where(line => IsExtendedHealthTotalBenefit(line.Benefit))
                .Where(IsMeaningfulRateLine)
                .ToList();
            if (totalLines.Count > 0)
            {
                return totalLines;
            }
        }

        return lines;
    }

    private static List<RateLine> BuildPowerPointRateLines(AnalysisResult analysis, bool omitZeroPremiumRows = false)
    {
        var lines = BuildPowerPointRateLines(SelectedRateLines(analysis));
        return omitZeroPremiumRows
            ? lines.Where(ShouldIncludeRenewalPowerPointBenefitRow).ToList()
            : lines;
    }

    private static bool ShouldIncludeRenewalPowerPointBenefitRow(RateLine line)
    {
        // This is a client-facing presentation filter only. Explicit $0 / $0
        // benefit rows add no renewal value, but missing amounts must remain
        // visible for review and prior-only or renewal-only benefits are retained.
        return line.CurrentPremium != 0m || line.RenewalPremium != 0m;
    }

    private static List<RateLine> SelectedRateLines(AnalysisResult analysis)
    {
        var version = analysis.SelectedRenewalVersion;
        if (version?.RateLines.Count > 0)
        {
            return BenefitDisplayOrder.Sort(
                version.RateLines.Select(ToRateLine),
                line => line.Benefit,
                line => line.TierOrClass,
                line => line.AccountClass);
        }

        return BenefitDisplayOrder.Sort(
            analysis.Pdf.RateLines,
            line => line.Benefit,
            line => line.TierOrClass,
            line => line.AccountClass);
    }

    private static RateLine ToRateLine(RenewalVersionRateLine line)
    {
        var currentPremium = line.CurrentPremium;
        var renewalPremium = line.VersionPremium ?? line.ProposedPremium;
        var currentRate = ClientPresentationRateSelector.CurrentRate(line);
        var proposedRate = ClientPresentationRateSelector.ProposedRate(line);
        var versionRate = ClientPresentationRateSelector.VersionRate(line);
        var canCalculate = !line.ValidationStatus.Equals(RateLineValidationStatuses.NeedsReview, StringComparison.OrdinalIgnoreCase);
        var currentCalculated = canCalculate ? UnderwritingEngine.CalculatePremium(line.Benefit, line.Volume, line.Lives, currentRate, line.RateBasis) : null;
        var renewalCalculated = canCalculate ? UnderwritingEngine.CalculatePremium(line.Benefit, line.Volume, line.Lives, versionRate, line.RateBasis) : null;

        return new RateLine
        {
            Benefit = line.Benefit,
            TierOrClass = line.TierOrClass,
            AccountClass = line.AccountClass,
            Volume = line.Volume,
            Lives = line.Lives,
            CurrentRate = currentRate,
            CurrentPremium = currentPremium,
            RenewalRate = versionRate,
            RenewalPremium = renewalPremium,
            CarrierProposedRate = proposedRate,
            CarrierProposedPremium = line.ProposedPremium,
            NegotiatedRate = versionRate,
            NegotiatedPremium = renewalPremium,
            PercentChange = line.ChangePercent ?? UnderwritingEngine.CalculatePercentChange(currentPremium, renewalPremium),
            CalculatedCurrentPremium = currentCalculated,
            CalculatedRenewalPremium = renewalCalculated,
            CurrentVariance = currentPremium.HasValue && currentCalculated.HasValue
                ? decimal.Round(currentPremium.Value - currentCalculated.Value, 2)
                : null,
            RenewalVariance = renewalPremium.HasValue && renewalCalculated.HasValue
                ? decimal.Round(renewalPremium.Value - renewalCalculated.Value, 2)
                : null,
            SourceReference = line.SourceReference,
            RateBasis = line.RateBasis,
            IncludeInPremiumTotal = line.IncludeInPremiumTotal,
            SourceRowRole = line.SourceRowRole,
            ValidationStatus = line.ValidationStatus,
            ValidationMessage = line.ValidationMessage,
            SourceCarrier = line.SourceCarrier,
            SourceFileName = line.SourceFileName,
            SourcePage = line.SourcePage,
            SourceText = line.SourceText,
            CurrentRatePrecision = line.CurrentRatePrecision,
            RenewalRatePrecision = line.RenewalRatePrecision,
            CurrentPeriodLabel = line.CurrentPeriodLabel,
            RenewalPeriodLabel = line.RenewalPeriodLabel,
            RenewalEffectiveDate = line.RenewalEffectiveDate
        };
    }

    private static List<RateLine> BuildPowerPointRateLines(IEnumerable<RateLine> rateLines)
    {
        var result = new List<RateLine>();
        foreach (var benefitGroup in OrderedBenefitGroups(rateLines))
        {
            var lines = benefitGroup
                .OrderBy(item => item.Index)
                .Select(item => item.Line)
                .ToList();

            result.AddRange(lines.Any(line => IsExtendedHealthBenefit(line.Benefit))
                ? BuildExtendedHealthPowerPointLines(lines)
                : lines);
        }

        return result;
    }

    private static List<RateLine> BuildExtendedHealthPowerPointLines(List<RateLine> lines)
    {
        var explicitTotals = lines
            .Where(line => IsExtendedHealthTotalBenefit(line.Benefit))
            .Where(IsMeaningfulRateLine)
            .ToList();
        if (explicitTotals.Count > 0)
        {
            return explicitTotals;
        }

        var componentLines = lines.Where(IsMeaningfulRateLine).ToList();
        if (componentLines.Count == 0)
        {
            return [];
        }

        return componentLines
            .Select((line, index) => new OrderedRateLine(index, line))
            .GroupBy(item => new
            {
                Account = OverviewAccountKey(item.Line),
                Source = Normalize(item.Line.SourceReference),
                Tier = NormalizeTier(item.Line.TierOrClass)
            })
            .OrderBy(group => group.Min(item => item.Index))
            .Select(group => BuildExtendedHealthTotalLine(group.Select(item => item.Line).ToList(), group.Key.Tier))
            .Where(IsMeaningfulRateLine)
            .ToList();
    }

    private static RateLine BuildExtendedHealthTotalLine(List<RateLine> lines, string normalizedTier)
    {
        var first = lines[0];
        var currentPremium = SumNullable(lines.Select(line => line.CurrentPremium));
        var proposedPremium = SumNullable(lines.Select(CarrierProposedPremium));
        var negotiatedPremium = SumNullable(lines.Select(NegotiatedPremium));
        var currentRate = SumNullable(lines.Select(line => line.CurrentRate));
        var proposedRate = SumNullable(lines.Select(CarrierProposedRate));
        var negotiatedRate = SumNullable(lines.Select(NegotiatedRate));
        var volume = lines.Any(line => line.Volume.HasValue)
            ? lines.Max(line => line.Volume ?? 0m)
            : (decimal?)null;
        var lives = lines.Any(line => line.Lives.HasValue)
            ? lines.Max(line => line.Lives ?? 0)
            : (int?)null;
        var calculatedCurrentPremium = CalculatePremium("Extended Health Care - Total", volume, lives, currentRate);
        var calculatedRenewalPremium = CalculatePremium("Extended Health Care - Total", volume, lives, negotiatedRate);

        return new RateLine
        {
            Benefit = "Extended Health Care - Total",
            TierOrClass = FriendlyTierName(normalizedTier),
            Volume = volume,
            Lives = lives,
            CurrentRate = currentRate,
            CurrentPremium = currentPremium,
            RenewalRate = negotiatedRate,
            RenewalPremium = negotiatedPremium,
            CarrierProposedRate = proposedRate,
            CarrierProposedPremium = proposedPremium,
            NegotiatedRate = negotiatedRate,
            NegotiatedPremium = negotiatedPremium,
            PercentChange = PercentChange(currentPremium, negotiatedPremium),
            CalculatedCurrentPremium = calculatedCurrentPremium,
            CalculatedRenewalPremium = calculatedRenewalPremium,
            CurrentVariance = currentPremium.HasValue && calculatedCurrentPremium.HasValue
                ? currentPremium.Value - calculatedCurrentPremium.Value
                : null,
            RenewalVariance = negotiatedPremium.HasValue && calculatedRenewalPremium.HasValue
                ? negotiatedPremium.Value - calculatedRenewalPremium.Value
                : null,
            SourceReference = first.SourceReference,
            AccountClass = first.AccountClass
        };
    }

    private static bool IsMeaningfulRateLine(RateLine line)
    {
        return line.Volume.HasValue ||
               line.Lives.HasValue ||
               line.CurrentRate.HasValue ||
               line.CurrentPremium.HasValue ||
               line.RenewalRate.HasValue ||
               line.RenewalPremium.HasValue;
    }

    private static decimal? SumNullable(IEnumerable<decimal?> values)
    {
        var numbers = values
            .Where(value => value.HasValue)
            .Select(value => value!.Value)
            .ToList();

        return numbers.Count == 0 ? null : numbers.Sum();
    }

    private static string AddOverviewAccountSuffix(string benefit, OverviewRateGroup group)
    {
        return group.ShowAccountLabel && !string.IsNullOrWhiteSpace(group.AccountLabel)
            ? $"{benefit} ({group.AccountLabel})"
            : benefit;
    }

    private static string OverviewAccountKey(RateLine line)
    {
        return Normalize(OverviewAccountLabel(line));
    }

    private static string OverviewAccountLabel(RateLine line)
    {
        var value = !string.IsNullOrWhiteSpace(line.AccountClass) ? line.AccountClass : "";
        value = value.Replace("ACCOUNT/CLASS:", "", StringComparison.OrdinalIgnoreCase).Trim();
        return value;
    }

    private static decimal? CollapseRate(IEnumerable<decimal?> rates)
    {
        var distinctRates = rates
            .Where(rate => rate.HasValue)
            .Select(rate => decimal.Round(rate!.Value, 4))
            .Distinct()
            .ToList();

        return distinctRates.Count == 1 ? distinctRates[0] : null;
    }

    private static IEnumerable<IGrouping<string, OrderedRateLine>> OrderedBenefitGroups(IEnumerable<RateLine> rateLines)
    {
        return rateLines
            .Select((line, index) => new OrderedRateLine(index, line))
            .GroupBy(item => MasterBenefitKey(item.Line.Benefit))
            .OrderBy(group => group.Min(item => item.Index));
    }

    private static List<RateAggregate> BuildMasterAggregates(IEnumerable<RateLine> rateLines)
    {
        return OrderedBenefitGroups(rateLines)
            .Select(group =>
            {
                var first = group.OrderBy(item => item.Index).First().Line;
                var lines = group.Select(item => item.Line).ToList();
                var currentPremium = lines.Sum(line => line.CurrentPremium ?? 0m);
                var renewalPremium = lines.Sum(line => line.RenewalPremium ?? 0m);
                return new RateAggregate(
                    group.Min(item => item.Index),
                    FriendlyBenefitName(first.Benefit),
                    SummarizeClass(lines.Select(line => line.TierOrClass)),
                    lines.Sum(line => line.Lives ?? 0),
                    lines.Any(line => line.Volume.HasValue) ? lines.Sum(line => line.Volume ?? 0m) : null,
                    SummarizeRate(lines.Select(line => line.CurrentRate)),
                    currentPremium,
                    SummarizeRate(lines.Select(line => line.RenewalRate)),
                    renewalPremium,
                    PercentChange(currentPremium, renewalPremium));
            })
            .Where(row => row.Lives != 0 || (row.Volume ?? 0m) != 0m || row.CurrentPremium != 0m || row.RenewalPremium != 0m)
            .OrderBy(row => row.Index)
            .ToList();
    }

    private static PowerPointExportTable BuildDivisionsTable(AnalysisResult analysis, bool omitZeroPremiumRows = false)
    {
        var rows = new List<string[]>();
        foreach (var group in BuildPowerPointRateLines(analysis, omitZeroPremiumRows).GroupBy(line => line.SourceReference))
        {
            var accountClass = group.FirstOrDefault(line => !string.IsNullOrWhiteSpace(line.AccountClass))?.AccountClass ?? group.Key;
            var currentTotal = 0m;
            var renewalTotal = 0m;
            foreach (var line in group)
            {
                currentTotal += line.CurrentPremium ?? 0m;
                renewalTotal += line.RenewalPremium ?? 0m;
                rows.Add(ToDivisionRow(accountClass, line));
            }

            AddDivisionTotals(rows, accountClass, currentTotal, renewalTotal);
        }

        return new PowerPointExportTable("RateTable_Divisions", DivisionHeaders(analysis), rows);
    }

    private static PowerPointExportTable BuildRateLineCalculationTable(AnalysisResult analysis, bool omitZeroPremiumRows = false)
    {
        var powerPointLines = BuildPowerPointRateLines(analysis, omitZeroPremiumRows);
        var rows = powerPointLines
            .Select(line => new[]
            {
                string.IsNullOrWhiteSpace(line.AccountClass) ? line.SourceReference : line.AccountClass,
                line.Benefit,
                FormatClass(line.TierOrClass),
                line.Volume.HasValue ? line.Volume.Value.ToString("#,##0.0") : "",
                line.Lives.HasValue ? line.Lives.Value.ToString("#,##0") : "",
                FormatRate(line.CurrentRate),
                ParsingHelpers.Money(line.CurrentPremium),
                FormatRate(CarrierProposedRate(line)),
                ParsingHelpers.Money(CarrierProposedPremium(line)),
                FormatRate(NegotiatedRate(line)),
                ParsingHelpers.Money(
                    NegotiatedPremium(line) ??
                    CalculatePremium(line.Benefit, line.Volume, line.Lives, NegotiatedRate(line), line.RateBasis)),
                // Hidden worksheet metadata. The visible/PPT table remains eleven columns.
                line.RateBasis
            })
            .ToList();

        var currentTotal = powerPointLines.Sum(line => line.CurrentPremium ?? 0m);
        var renewalTotal = powerPointLines.Sum(line => CarrierProposedPremium(line) ?? 0m);
        var negotiatedTotal = powerPointLines.Sum(line =>
            NegotiatedPremium(line) ??
            CalculatePremium(line.Benefit, line.Volume, line.Lives, NegotiatedRate(line), line.RateBasis) ?? 0m);
        rows.Add(
        [
            "Total Monthly Premiums",
            "",
            "",
            "",
            "",
            "",
            ParsingHelpers.Money(currentTotal),
            "",
            ParsingHelpers.Money(renewalTotal),
            "",
            ParsingHelpers.Money(negotiatedTotal)
        ]);
        rows.Add(
        [
            "Monthly Difference from Current",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            ParsingHelpers.Money(renewalTotal - currentTotal),
            "",
            ParsingHelpers.Money(negotiatedTotal - currentTotal)
        ]);

        return new PowerPointExportTable("RateTable_LineRates", RateLineCalculationHeaders, rows);
    }

    private static PowerPointExportTable BuildFormattedRateLineCalculationTable(
        AnalysisResult analysis,
        string title = "RateTable_LineRates",
        bool omitZeroPremiumRows = false)
    {
        var lines = BuildPowerPointRateLines(analysis, omitZeroPremiumRows);
        var groups = BuildPowerPointFormattedLineRateGroups(lines);
        var rows = new List<string[]> { FormattedRateLineColumnHeaders };
        var allDetailLines = new List<RateLine>();
        var pooledLines = new List<RateLine>();
        var ehcLines = new List<RateLine>();
        var dentalLines = new List<RateLine>();
        var pooledTotalWritten = false;

        for (var index = 0; index < groups.Count; index++)
        {
            var group = groups[index];
            if (!pooledTotalWritten && group.SortOrder >= 90 && pooledLines.Count > 0)
            {
                AddFormattedLineRateTotalRows(rows, "Total Pooled Benefits", pooledLines);
                AddFormattedLineRateDifferenceRow(rows, "Pooled Difference", pooledLines);
                pooledTotalWritten = true;
            }

            rows.Add(FormattedLineRateSectionRow(group.DisplayName));
            foreach (var line in group.Lines)
            {
                rows.Add(ToFormattedLineRateRow(line));
                allDetailLines.Add(line);

                if (group.SortOrder < 90)
                {
                    pooledLines.Add(line);
                }

                if (group.Kind == PowerPointFormattedLineRateGroupKind.ExtendedHealth)
                {
                    ehcLines.Add(line);
                }
                else if (group.Kind == PowerPointFormattedLineRateGroupKind.Dental)
                {
                    dentalLines.Add(line);
                }
            }

            var nextKind = index + 1 < groups.Count ? groups[index + 1].Kind : (PowerPointFormattedLineRateGroupKind?)null;
            if (group.Kind == PowerPointFormattedLineRateGroupKind.ExtendedHealth &&
                nextKind != PowerPointFormattedLineRateGroupKind.ExtendedHealth)
            {
                AddFormattedLineRateTotalRows(rows, "Total Health Premiums", ehcLines);
                AddFormattedLineRateAnnualRow(rows, "Total Annual EHC premiums", ehcLines);
            }

            if (group.Kind == PowerPointFormattedLineRateGroupKind.Dental &&
                nextKind != PowerPointFormattedLineRateGroupKind.Dental)
            {
                AddFormattedLineRateTotalRows(rows, "Total Dental Premiums", dentalLines);
                AddFormattedLineRateAnnualRow(rows, "Total Annual Dental Premiums", dentalLines);
                if (ehcLines.Count > 0)
                {
                    AddFormattedLineRateTotalRows(rows, "Total EHC & Dental", ehcLines.Concat(dentalLines));
                }
            }
        }

        if (!pooledTotalWritten && pooledLines.Count > 0)
        {
            AddFormattedLineRateTotalRows(rows, "Total Pooled Benefits", pooledLines);
            AddFormattedLineRateDifferenceRow(rows, "Pooled Difference", pooledLines);
        }

        AddFormattedLineRateTotalRows(rows, "Total monthly premium", allDetailLines);
        AddFormattedLineRateDifferenceRow(rows, "Monthly premium difference", allDetailLines);
        AddFormattedLineRateAnnualRow(rows, "Total annual premium", allDetailLines);
        AddFormattedLineRateAnnualDifferenceRow(rows, "Annual premium difference", allDetailLines);

        return new PowerPointExportTable(title, FormattedRateLineGroupHeaders, rows);
    }

    private static List<PowerPointFormattedLineRateGroup> BuildPowerPointFormattedLineRateGroups(List<RateLine> lines)
    {
        return lines
            .Select((line, index) => new
            {
                Line = line,
                Index = index,
                Key = GetPowerPointFormattedLineRateGroupKey(line.Benefit, line.RateBasis)
            })
            .GroupBy(item => item.Key.Key, StringComparer.OrdinalIgnoreCase)
            .Select(group =>
            {
                var first = group.OrderBy(item => item.Index).First();
                return new PowerPointFormattedLineRateGroup(
                    first.Key.Key,
                    first.Key.DisplayName,
                    first.Key.SortOrder,
                    first.Key.Kind,
                    first.Index,
                    group.OrderBy(item => item.Index).Select(item => item.Line).ToList());
            })
            .OrderBy(group => group.SortOrder)
            .ThenBy(group => group.FirstIndex)
            .ThenBy(group => group.DisplayName, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static string[] FormattedLineRateSectionRow(string title)
    {
        return [title, "", "", "", "", "", "", "", "", ""];
    }

    private static string[] ToFormattedLineRateRow(RateLine line)
    {
        var calculatedPremium = CalculatePremium(line.Benefit, line.Volume, line.Lives, NegotiatedRate(line), line.RateBasis);
        var proposedPremium = CarrierProposedPremium(line);
        var negotiatedPremium = NegotiatedPremium(line) ?? calculatedPremium;
        return
        [
            FormatPowerPointFormattedLineRateDetail(line),
            FormatLineRateVolume(line),
            FormatRateCurrency(line.CurrentRate),
            ParsingHelpers.Money(line.CurrentPremium),
            FormatRateCurrency(NegotiatedRate(line)),
            ParsingHelpers.Money(calculatedPremium),
            FormatRateCurrency(CarrierProposedRate(line)),
            ParsingHelpers.Money(proposedPremium),
            FormatRateCurrency(NegotiatedRate(line)),
            ParsingHelpers.Money(negotiatedPremium)
        ];
    }

    private static void AddFormattedLineRateTotalRows(List<string[]> rows, string label, IEnumerable<RateLine> sourceLines)
    {
        var lines = sourceLines.ToList();
        rows.Add(
        [
            label,
            "",
            "",
            ParsingHelpers.Money(lines.Sum(line => line.CurrentPremium ?? 0m)),
            "",
            ParsingHelpers.Money(lines.Sum(line => CalculatePremium(line.Benefit, line.Volume, line.Lives, NegotiatedRate(line), line.RateBasis) ?? 0m)),
            "",
            ParsingHelpers.Money(lines.Sum(line => CarrierProposedPremium(line) ?? 0m)),
            "",
            ParsingHelpers.Money(lines.Sum(line => NegotiatedPremium(line) ?? CalculatePremium(line.Benefit, line.Volume, line.Lives, NegotiatedRate(line), line.RateBasis) ?? 0m))
        ]);
    }

    private static void AddFormattedLineRateDifferenceRow(List<string[]> rows, string label, IEnumerable<RateLine> sourceLines)
    {
        var lines = sourceLines.ToList();
        var current = lines.Sum(line => line.CurrentPremium ?? 0m);
        var calculated = lines.Sum(line => CalculatePremium(line.Benefit, line.Volume, line.Lives, NegotiatedRate(line), line.RateBasis) ?? 0m);
        var proposed = lines.Sum(line => CarrierProposedPremium(line) ?? 0m);
        var negotiated = lines.Sum(line => NegotiatedPremium(line) ?? CalculatePremium(line.Benefit, line.Volume, line.Lives, NegotiatedRate(line), line.RateBasis) ?? 0m);
        rows.Add(
        [
            label,
            "",
            "",
            "",
            "",
            ParsingHelpers.Money(calculated - current),
            "",
            ParsingHelpers.Money(proposed - current),
            "",
            ParsingHelpers.Money(negotiated - current)
        ]);
    }

    private static void AddFormattedLineRateAnnualRow(List<string[]> rows, string label, IEnumerable<RateLine> sourceLines)
    {
        var lines = sourceLines.ToList();
        rows.Add(
        [
            label,
            "",
            "",
            ParsingHelpers.Money(lines.Sum(line => line.CurrentPremium ?? 0m) * 12m),
            "",
            ParsingHelpers.Money(lines.Sum(line => CalculatePremium(line.Benefit, line.Volume, line.Lives, NegotiatedRate(line), line.RateBasis) ?? 0m) * 12m),
            "",
            ParsingHelpers.Money(lines.Sum(line => CarrierProposedPremium(line) ?? 0m) * 12m),
            "",
            ParsingHelpers.Money(lines.Sum(line => NegotiatedPremium(line) ?? CalculatePremium(line.Benefit, line.Volume, line.Lives, NegotiatedRate(line), line.RateBasis) ?? 0m) * 12m)
        ]);
    }

    private static void AddFormattedLineRateAnnualDifferenceRow(List<string[]> rows, string label, IEnumerable<RateLine> sourceLines)
    {
        var lines = sourceLines.ToList();
        var current = lines.Sum(line => line.CurrentPremium ?? 0m);
        var calculated = lines.Sum(line => CalculatePremium(line.Benefit, line.Volume, line.Lives, NegotiatedRate(line), line.RateBasis) ?? 0m);
        var proposed = lines.Sum(line => CarrierProposedPremium(line) ?? 0m);
        var negotiated = lines.Sum(line => NegotiatedPremium(line) ?? CalculatePremium(line.Benefit, line.Volume, line.Lives, NegotiatedRate(line), line.RateBasis) ?? 0m);
        rows.Add(
        [
            label,
            "",
            "",
            "",
            "",
            ParsingHelpers.Money((calculated - current) * 12m),
            "",
            ParsingHelpers.Money((proposed - current) * 12m),
            "",
            ParsingHelpers.Money((negotiated - current) * 12m)
        ]);
    }

    private static string FormatPowerPointFormattedLineRateDetail(RateLine line)
    {
        var tier = FormatClass(line.TierOrClass);
        if (!string.IsNullOrWhiteSpace(tier) && !tier.Equals("All", StringComparison.OrdinalIgnoreCase))
        {
            return $"  {tier}";
        }

        var account = OverviewAccountLabel(line);
        if (!string.IsNullOrWhiteSpace(account) && !account.Equals("ALL", StringComparison.OrdinalIgnoreCase))
        {
            return account;
        }

        return "All";
    }

    private static string FormatLineRateVolume(RateLine line)
    {
        if (line.Volume.HasValue && line.Volume.Value != 0m)
        {
            return line.Volume.Value.ToString("#,##0");
        }

        return line.Lives.HasValue && line.Lives.Value != 0 ? line.Lives.Value.ToString("#,##0") : "";
    }

    private static PowerPointFormattedLineRateGroupKey GetPowerPointFormattedLineRateGroupKey(string benefit, string rateBasis)
    {
        var normalized = NormalizeForFormattedLineRate(benefit);

        if (normalized.Contains("OPTIONAL", StringComparison.Ordinal) && IsFormattedLineRateAdAndD(normalized))
        {
            return new("OPTIONAL AD&D", "Optional AD&D", 60, PowerPointFormattedLineRateGroupKind.Other);
        }

        if (normalized.Contains("OPTIONAL LIFE", StringComparison.Ordinal))
        {
            return new("OPTIONAL LIFE", "Optional Life", 50, PowerPointFormattedLineRateGroupKind.Other);
        }

        if (normalized.Contains("DEPENDENT LIFE", StringComparison.Ordinal))
        {
            return new("DEPENDENT LIFE", "Dependent Life", 40, PowerPointFormattedLineRateGroupKind.Pooled);
        }

        if (normalized.Equals("LIFE", StringComparison.Ordinal) ||
            normalized.Contains("BASIC LIFE", StringComparison.Ordinal) ||
            normalized.Contains("EMPLOYEE LIFE", StringComparison.Ordinal))
        {
            return new("LIFE", "Employee Life Insurance", 10, PowerPointFormattedLineRateGroupKind.Pooled);
        }

        if (IsFormattedLineRateAdAndD(normalized))
        {
            return new("AD&D", "Employee AD&D", 20, PowerPointFormattedLineRateGroupKind.Pooled);
        }

        if (normalized.Contains("CRITICAL ILLNESS", StringComparison.Ordinal))
        {
            if (UnderwritingEngine.IsPerEmployeeRateBasis(rateBasis))
            {
                return new("STAND-ALONE CRITICAL ILLNESS", "Critical Illness", 150, PowerPointFormattedLineRateGroupKind.Other);
            }

            return new("CRITICAL ILLNESS", "Critical Illness", 30, PowerPointFormattedLineRateGroupKind.Pooled);
        }

        if (normalized.Equals("STD", StringComparison.Ordinal) ||
            normalized.Contains("SHORT TERM DISABILITY", StringComparison.Ordinal))
        {
            return new("STD", "Short-Term Disability", 70, PowerPointFormattedLineRateGroupKind.Pooled);
        }

        if (normalized.Equals("LTD", StringComparison.Ordinal) ||
            normalized.Contains("LONG TERM DISABILITY", StringComparison.Ordinal))
        {
            return new("LTD", "Long-Term Disability", 80, PowerPointFormattedLineRateGroupKind.Pooled);
        }

        if (normalized.StartsWith("EXTENDED HEALTH CARE", StringComparison.Ordinal) ||
            normalized.Equals("EHC", StringComparison.Ordinal))
        {
            return new("EHC", "Extended Health Care", 90, PowerPointFormattedLineRateGroupKind.ExtendedHealth);
        }

        if (normalized.Contains("DENTAL", StringComparison.Ordinal))
        {
            return new("DENTAL", "Dental Care", 100, PowerPointFormattedLineRateGroupKind.Dental);
        }

        if (normalized.Contains("EMPLOYEE ASSISTANCE", StringComparison.Ordinal) ||
            (normalized.Contains("EMPLOYEE", StringComparison.Ordinal) && normalized.Contains("FAMILY ASSISTANCE", StringComparison.Ordinal)) ||
            normalized.Equals("EAP", StringComparison.Ordinal) ||
            normalized.Equals("EFAP", StringComparison.Ordinal))
        {
            return new("EAP", "Employee Assistance Program", 110, PowerPointFormattedLineRateGroupKind.Other);
        }

        if (normalized.Contains("MEDICAL SECOND OPINION", StringComparison.Ordinal) ||
            normalized.Equals("MSO", StringComparison.Ordinal) ||
            normalized.Contains("BEST DOCTORS", StringComparison.Ordinal))
        {
            return new("MSO", "Medical Second Opinion", 120, PowerPointFormattedLineRateGroupKind.Other);
        }

        if (normalized.Contains("VIRTUAL", StringComparison.Ordinal) ||
            normalized.Contains("TELEMEDICINE", StringComparison.Ordinal) ||
            normalized.Contains("ONLINE DOCTOR", StringComparison.Ordinal) ||
            normalized.Contains("ONLINE ACCESS", StringComparison.Ordinal) ||
            normalized.Contains("HEALTHCARE ONLINE", StringComparison.Ordinal))
        {
            return new("VIRTUAL HEALTH CARE", "Virtual Health Care / Online Access", 130, PowerPointFormattedLineRateGroupKind.Other);
        }

        if (normalized.Contains("HEALTH SPENDING", StringComparison.Ordinal) ||
            normalized.Equals("HSA", StringComparison.Ordinal) ||
            normalized.Equals("HCSA", StringComparison.Ordinal))
        {
            return new("HSA", "Health Spending Account", 140, PowerPointFormattedLineRateGroupKind.Other);
        }

        if (normalized.Contains("PERSONAL SPENDING", StringComparison.Ordinal) ||
            normalized.Equals("PSA", StringComparison.Ordinal))
        {
            return new("PSA", "Personal Spending Account", 150, PowerPointFormattedLineRateGroupKind.Other);
        }

        if (normalized.Contains("WELLNESS", StringComparison.Ordinal))
        {
            return new("WELLNESS", "Wellness", 160, PowerPointFormattedLineRateGroupKind.Other);
        }

        var displayName = string.IsNullOrWhiteSpace(benefit) ? "Other Benefits" : benefit.Trim();
        return new($"OTHER::{NormalizeForFormattedLineRate(displayName)}", displayName, 900, PowerPointFormattedLineRateGroupKind.Other);
    }

    private static bool IsFormattedLineRateAdAndD(string normalizedBenefit)
    {
        return normalizedBenefit.Equals("AD&D", StringComparison.Ordinal) ||
               normalizedBenefit.Equals("AD AND D", StringComparison.Ordinal) ||
               normalizedBenefit.Equals("ASI", StringComparison.Ordinal) ||
               normalizedBenefit.Contains("ACCIDENTAL DEATH", StringComparison.Ordinal) ||
               normalizedBenefit.Contains("DISMEMBERMENT", StringComparison.Ordinal);
    }

    private static string NormalizeForFormattedLineRate(string value)
    {
        var cleaned = ParsingHelpers.Clean(value)
            .Replace("&", " AND ", StringComparison.Ordinal)
            .Replace("-", " ", StringComparison.Ordinal)
            .Replace("/", " ", StringComparison.Ordinal)
            .ToUpperInvariant();

        return string.Join(" ", cleaned.Split(' ', StringSplitOptions.RemoveEmptyEntries));
    }

    private static decimal? CalculatePremium(string benefit, decimal? volume, int? lives, decimal? rate, string rateBasis = "")
    {
        return UnderwritingEngine.CalculatePremium(benefit, volume, lives, rate, rateBasis);
    }

    private static decimal? CarrierProposedRate(RateLine line) => line.CarrierProposedRate ?? line.RenewalRate;

    private static decimal? NegotiatedRate(RateLine line) => line.NegotiatedRate ?? line.RenewalRate;

    private static decimal? CarrierProposedPremium(RateLine line) => line.CarrierProposedPremium ?? line.RenewalPremium;

    private static decimal? NegotiatedPremium(RateLine line) => line.NegotiatedPremium ?? line.RenewalPremium;

    private static int? RateVolumeDivisor(string benefit, string? rateBasis = null)
    {
        return UnderwritingEngine.RateVolumeDivisor(benefit, rateBasis);
    }

    private static IEnumerable<PowerPointExportTable> BuildSingleDivisionTables(AnalysisResult analysis, bool omitZeroPremiumRows = false)
    {
        var index = 1;
        foreach (var group in BuildPowerPointRateLines(analysis, omitZeroPremiumRows).GroupBy(line => line.SourceReference))
        {
            var accountClass = group.FirstOrDefault(line => !string.IsNullOrWhiteSpace(line.AccountClass))?.AccountClass ?? group.Key;
            var rows = group.Select(line => ToDivisionRow(accountClass, line)).ToList();
            var currentTotal = group.Sum(line => line.CurrentPremium ?? 0m);
            var renewalTotal = group.Sum(line => line.RenewalPremium ?? 0m);
            AddDivisionTotals(rows, accountClass, currentTotal, renewalTotal);
            yield return new PowerPointExportTable($"RateTable_Division{index}: {accountClass}", DivisionHeaders(analysis), rows);
            index++;
        }
    }

    private static string[] ToDivisionRow(string accountClass, RateLine line)
    {
        return
        [
            accountClass,
            line.Benefit,
            FormatClass(line.TierOrClass),
            FormatInsuredVolume(line.Volume, line.Lives),
            FormatPremiumCell(FormatRate(line.CurrentRate), line.CurrentPremium),
            FormatPremiumCell(FormatRate(line.RenewalRate), line.RenewalPremium),
            line.PercentChange.HasValue ? ParsingHelpers.Percent(line.PercentChange) : ParsingHelpers.Percent(PercentChange(line.CurrentPremium ?? 0m, line.RenewalPremium ?? 0m))
        ];
    }

    private static string[] ToPremiumOverviewRow(RateAggregate row)
    {
        return
        [
            row.Benefit,
            FormatClass(row.TierOrClass),
            FormatInsuredVolume(row.Volume, row.Lives),
            FormatPremiumCell(row.CurrentRate, row.CurrentPremium),
            FormatPremiumCell(row.RenewalRate, row.RenewalPremium),
            ParsingHelpers.Percent(row.ChangePercent)
        ];
    }

    private static string[] ToSimpleRateTableRow(RateAggregate row)
    {
        return
        [
            SimpleBenefitName(row.Benefit),
            ParsingHelpers.Money(row.CurrentPremium),
            ParsingHelpers.Money(row.RenewalPremium),
            ParsingHelpers.Percent(row.ChangePercent)
        ];
    }

    private static List<string[]> BuildHealthDentalBreakdownRows(
        AnalysisResult analysis,
        HealthDentalBenefit benefitType,
        bool includeEmptyFallback,
        bool omitZeroPremiumRows = false)
    {
        var lines = BuildPowerPointRateLines(analysis, omitZeroPremiumRows)
            .Where(line => benefitType == HealthDentalBenefit.ExtendedHealthCare
                ? IsExtendedHealthBenefit(line.Benefit)
                : IsDentalBenefit(line.Benefit))
            .ToList();

        if (lines.Count == 0)
        {
            return [];
        }

        var tiers = lines
            .GroupBy(line => NormalizeTier(line.TierOrClass))
            .Select(group =>
            {
                var tierLines = group.ToList();
                return new TierAggregate(
                    TierSort(group.Key),
                    FriendlyTierName(group.Key),
                    tierLines.Sum(line => line.Lives ?? 0),
                    tierLines.Any(line => line.Volume.HasValue) ? tierLines.Sum(line => line.Volume ?? 0m) : null,
                    tierLines.Sum(line => line.CurrentPremium ?? 0m),
                    tierLines.Sum(line => line.RenewalPremium ?? 0m));
            })
            .Where(tier => tier.Lives != 0 || (tier.Volume ?? 0m) != 0m || tier.CurrentPremium != 0m || tier.RenewalPremium != 0m)
            .OrderBy(tier => tier.SortOrder)
            .ThenBy(tier => tier.Name)
            .ToList();

        if (tiers.Count == 0 && !includeEmptyFallback)
        {
            return [];
        }

        if (tiers.Count == 0)
        {
            var currentTotal = lines.Sum(line => line.CurrentPremium ?? 0m);
            var renewalTotal = lines.Sum(line => line.RenewalPremium ?? 0m);
            return
            [
                [
                    benefitType == HealthDentalBenefit.ExtendedHealthCare ? "Extended Health Care" : "Dental Care",
                    "All",
                    "",
                    ParsingHelpers.Money(currentTotal),
                    ParsingHelpers.Money(renewalTotal),
                    ParsingHelpers.Percent(PercentChange(currentTotal, renewalTotal))
                ]
            ];
        }

        var currentBenefitTotal = tiers.Sum(tier => tier.CurrentPremium);
        var renewalBenefitTotal = tiers.Sum(tier => tier.RenewalPremium);
        var classText = string.Join("\n", tiers.Select(tier => tier.Name));
        var volumeText = string.Join("\n", tiers.Select(tier => FormatTierVolume(tier.Volume, tier.Lives)));
        var currentText = string.Join("\n", tiers.Select(tier => ParsingHelpers.Money(tier.CurrentPremium)).Append(ParsingHelpers.Money(currentBenefitTotal)));
        var renewalText = string.Join("\n", tiers.Select(tier => ParsingHelpers.Money(tier.RenewalPremium)).Append(ParsingHelpers.Money(renewalBenefitTotal)));

        return
        [
            [
                benefitType == HealthDentalBenefit.ExtendedHealthCare ? "Extended Health Care" : "Dental Care",
                classText,
                volumeText,
                currentText,
                renewalText,
                ParsingHelpers.Percent(PercentChange(currentBenefitTotal, renewalBenefitTotal))
            ]
        ];
    }

    private static string[] MasterHeaders(AnalysisResult analysis)
    {
        return
        [
            "Benefit",
            "Class",
            "Insured Volume",
            "Current",
            $"{FormatRenewalDate(analysis)} Renewal",
            "% Change from Current"
        ];
    }

    private static string[] DivisionHeaders(AnalysisResult analysis)
    {
        return
        [
            "Account/Class",
            "Benefit",
            "Class",
            "Insured Volume",
            "Current",
            $"{FormatRenewalDate(analysis)} Renewal",
            "% Change from Current"
        ];
    }

    private static void AddPremiumTotals(List<string[]> rows, decimal currentTotal, decimal renewalTotal)
    {
        rows.Add(["Total Monthly Premiums", "", "", ParsingHelpers.Money(currentTotal), ParsingHelpers.Money(renewalTotal), ParsingHelpers.Percent(PercentChange(currentTotal, renewalTotal))]);
        rows.Add(["Monthly Difference from Current", "", "", "", ParsingHelpers.Money(renewalTotal - currentTotal), ParsingHelpers.Percent(PercentChange(currentTotal, renewalTotal))]);
    }

    private static void AddDivisionTotals(List<string[]> rows, string accountClass, decimal currentTotal, decimal renewalTotal)
    {
        rows.Add([accountClass, "Total Monthly Premiums", "", "", ParsingHelpers.Money(currentTotal), ParsingHelpers.Money(renewalTotal), ParsingHelpers.Percent(PercentChange(currentTotal, renewalTotal))]);
        rows.Add([accountClass, "Monthly Difference from Current", "", "", "", ParsingHelpers.Money(renewalTotal - currentTotal), ParsingHelpers.Percent(PercentChange(currentTotal, renewalTotal))]);
    }

    private static string FormatClass(string tierOrClass)
    {
        return string.IsNullOrWhiteSpace(tierOrClass) ? "All" : tierOrClass;
    }

    private static string FormatInsuredVolume(decimal? volume, int? lives)
    {
        if (volume.HasValue && volume.Value != 0m)
        {
            return volume.Value.ToString("$#,##0");
        }

        return lives.HasValue && lives.Value != 0 ? lives.Value.ToString("#,##0") : "";
    }

    private static string FormatTierVolume(decimal? volume, int lives)
    {
        if (volume.HasValue && volume.Value != 0m)
        {
            return volume.Value.ToString("$#,##0");
        }

        return lives == 0 ? "" : lives.ToString("#,##0");
    }

    private static string RateBasisLabel(string benefit)
    {
        if (benefit.Equals("Life", StringComparison.OrdinalIgnoreCase) ||
            benefit.Contains("Accidental Death", StringComparison.OrdinalIgnoreCase) ||
            benefit.Contains("Critical Illness", StringComparison.OrdinalIgnoreCase))
        {
            return "Rate per $1,000 of volume";
        }

        if (benefit.Contains("Long Term Disability", StringComparison.OrdinalIgnoreCase) ||
            benefit.Contains("Long-Term Disability", StringComparison.OrdinalIgnoreCase))
        {
            return "Rate per $100 of volume";
        }

        if (benefit.Contains("Short Term Disability", StringComparison.OrdinalIgnoreCase) ||
            benefit.Contains("Short-Term Disability", StringComparison.OrdinalIgnoreCase) ||
            benefit.Equals("STD", StringComparison.OrdinalIgnoreCase))
        {
            return "Rate per $10 of volume";
        }

        return "Rate per life/tier";
    }

    private static string FormatPremiumCell(string rate, decimal? premium)
    {
        var premiumText = ParsingHelpers.Money(premium);
        if (string.IsNullOrWhiteSpace(rate))
        {
            return premiumText;
        }

        return string.IsNullOrWhiteSpace(premiumText) ? rate : $"{rate}\n{premiumText}";
    }

    private static string SummarizeRate(IEnumerable<decimal?> rates)
    {
        var distinctRates = rates
            .Where(rate => rate.HasValue)
            .Select(rate => decimal.Round(rate!.Value, 4))
            .Distinct()
            .OrderBy(rate => rate)
            .ToList();

        return distinctRates.Count switch
        {
            0 => "",
            1 => FormatRate(distinctRates[0]),
            _ => "varies"
        };
    }

    private static string SummarizeRateCurrency(IEnumerable<decimal?> rates)
    {
        var distinctRates = rates
            .Where(rate => rate.HasValue)
            .Select(rate => decimal.Round(rate!.Value, 4))
            .Distinct()
            .OrderBy(rate => rate)
            .ToList();

        return distinctRates.Count switch
        {
            0 => "",
            1 => FormatRateCurrency(distinctRates[0]),
            _ => "varies"
        };
    }

    private static string SummarizeClass(IEnumerable<string> classes)
    {
        var distinctClasses = classes
            .Select(ParsingHelpers.Clean)
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(value => value)
            .ToList();

        return distinctClasses.Count switch
        {
            0 => "",
            1 => distinctClasses[0],
            _ => "All"
        };
    }

    private static string FormatRate(decimal? rate)
    {
        return rate.HasValue ? rate.Value.ToString("0.####") : "";
    }

    private static string FormatRateCurrency(decimal? rate)
    {
        if (!rate.HasValue)
        {
            return "";
        }

        var format = Math.Abs(rate.Value) < 1m ? "$0.####" : "$#,##0.00##";
        return rate.Value.ToString(format);
    }

    private static string FriendlyRatesOverviewBenefitName(string benefit)
    {
        if (Normalize(benefit).Equals("LIFE", StringComparison.Ordinal))
        {
            return "Employee Life Insurance (All Classes)";
        }

        if (IsAccidentalDeathAndDismemberment(benefit))
        {
            return "Employee AD&D (All Classes)";
        }

        if (IsExtendedHealthBenefit(benefit))
        {
            return "Extended Health Care - Total";
        }

        if (IsDentalBenefit(benefit))
        {
            return "Dental Care (All Classes)";
        }

        return FriendlyBenefitName(benefit);
    }

    private static string FriendlyPremiumOverviewBenefitName(string benefit)
    {
        if (Normalize(benefit).Equals("LIFE", StringComparison.Ordinal))
        {
            return "Employee Life";
        }

        if (IsAccidentalDeathAndDismemberment(benefit))
        {
            return "AD&D";
        }

        if (IsExtendedHealthBenefit(benefit))
        {
            return "Extended Health Care";
        }

        if (IsDentalBenefit(benefit))
        {
            return "Dental Care";
        }

        return FriendlyBenefitName(benefit);
    }

    private static string PremiumIllustrationBenefitName(string benefit)
    {
        return BenefitDisplayOrder.BenefitRank(benefit) switch
        {
            0 => "Employee Life Insurance",
            1 => "Employee AD&D",
            2 => "Dependent Life",
            3 => "Critical Illness",
            4 => "Short-Term Disability",
            5 => "Long-Term Disability",
            6 => "Extended Health Care",
            7 => "Dental Care",
            _ => FriendlyBenefitName(benefit)
        };
    }

    private static string PremiumIllustrationClassName(string accountClass)
    {
        var value = ParsingHelpers.Clean(accountClass);
        return string.IsNullOrWhiteSpace(value) ||
               value.Equals("All", StringComparison.OrdinalIgnoreCase) ||
               value.Equals("All Classes", StringComparison.OrdinalIgnoreCase)
            ? "All"
            : value;
    }

    private static string MasterBenefitKey(string benefit)
    {
        if (IsAccidentalDeathAndDismemberment(benefit))
        {
            return "ACCIDENTAL DEATH AND DISMEMBERMENT";
        }

        if (IsExtendedHealthBenefit(benefit))
        {
            return "EXTENDED HEALTH CARE";
        }

        if (IsDentalBenefit(benefit))
        {
            return "DENTAL";
        }

        return Normalize(benefit);
    }

    private static string FriendlyBenefitName(string benefit)
    {
        if (IsAccidentalDeathAndDismemberment(benefit))
        {
            return "Accidental Death and Dismemberment";
        }

        if (IsExtendedHealthBenefit(benefit))
        {
            return "Extended Health Care";
        }

        if (IsDentalBenefit(benefit))
        {
            return "Dental Care";
        }

        return benefit;
    }

    private static string SimpleBenefitName(string benefit)
    {
        if (Normalize(benefit).Equals("LIFE", StringComparison.Ordinal))
        {
            return "Basic Life";
        }

        if (IsAccidentalDeathAndDismemberment(benefit))
        {
            return "AD&D";
        }

        return benefit;
    }

    private static bool IsAccidentalDeathAndDismemberment(string benefit)
    {
        var normalized = Normalize(benefit);
        return normalized.Equals("AD&D", StringComparison.Ordinal) ||
               normalized.Equals("AD AND D", StringComparison.Ordinal) ||
               normalized.Equals("ASI", StringComparison.Ordinal) ||
               normalized.Contains("ACCIDENTAL DEATH", StringComparison.Ordinal) ||
               normalized.Contains("DISMEMBERMENT", StringComparison.Ordinal);
    }

    private static bool IsExtendedHealthBenefit(string benefit)
    {
        return Normalize(benefit).StartsWith("EXTENDED HEALTH CARE", StringComparison.Ordinal);
    }

    private static bool IsExtendedHealthTotalBenefit(string benefit)
    {
        var normalized = Normalize(benefit);
        return normalized.StartsWith("EXTENDED HEALTH CARE", StringComparison.Ordinal) &&
               normalized.Contains("TOTAL", StringComparison.Ordinal);
    }

    private static bool IsDentalBenefit(string benefit)
    {
        var normalized = Normalize(benefit);
        return normalized.Equals("DENTAL", StringComparison.Ordinal) ||
               normalized.StartsWith("DENTAL ", StringComparison.Ordinal) ||
               normalized.EndsWith(" DENTAL", StringComparison.Ordinal) ||
               normalized.Contains(" DENTAL ", StringComparison.Ordinal);
    }

    private static string NormalizeTier(string tierOrClass)
    {
        var normalized = Normalize(tierOrClass);
        if (normalized.Contains("SINGLE PARENT", StringComparison.Ordinal))
        {
            return "SINGLE PARENT";
        }

        if (normalized.Contains("SINGLE", StringComparison.Ordinal))
        {
            return "SINGLE";
        }

        if (normalized.Contains("COUPLE", StringComparison.Ordinal))
        {
            return "COUPLE";
        }

        if (normalized.Contains("FAMILY", StringComparison.Ordinal))
        {
            return "FAMILY";
        }

        return string.IsNullOrWhiteSpace(normalized) ? "ALL" : normalized;
    }

    private static string FriendlyTierName(string normalizedTier)
    {
        return normalizedTier switch
        {
            "SINGLE" => "Single",
            "SINGLE PARENT" => "Single Parent",
            "COUPLE" => "Couple",
            "FAMILY" => "Family",
            "ALL" => "All",
            _ => normalizedTier
        };
    }

    private static int TierSort(string normalizedTier)
    {
        return BenefitDisplayOrder.TierRank(normalizedTier);
    }

    private static string FormatRenewalDate(AnalysisResult analysis)
    {
        var renewalDate = analysis.RenewalCase.RenewalDate ?? analysis.Pdf.RenewalDate;
        return renewalDate.HasValue ? renewalDate.Value.ToString("MMMM d, yyyy") : "Proposed";
    }

    private static string FormatShortRenewalDate(AnalysisResult analysis)
    {
        var renewalDate = analysis.RenewalCase.RenewalDate ?? analysis.Pdf.RenewalDate;
        return renewalDate.HasValue
            ? renewalDate.Value.ToString("d-MMM-yy", System.Globalization.CultureInfo.InvariantCulture)
            : "Proposed";
    }

    private static string Normalize(string value)
    {
        return ParsingHelpers.Clean(value).ToUpperInvariant();
    }

    private static string FormatWholeNumber(decimal? value)
    {
        return value.HasValue && value.Value != 0m ? value.Value.ToString("#,##0") : "";
    }

    private static decimal? PercentChange(decimal current, decimal renewal)
    {
        return UnderwritingEngine.CalculatePercentChange(current, renewal);
    }

    private static decimal? PercentChange(decimal? current, decimal? renewal)
    {
        return current.HasValue && renewal.HasValue
            ? PercentChange(current.Value, renewal.Value)
            : null;
    }

    private sealed record RateAggregate(
        int Index,
        string Benefit,
        string TierOrClass,
        int Lives,
        decimal? Volume,
        string CurrentRate,
        decimal CurrentPremium,
        string RenewalRate,
        decimal RenewalPremium,
        decimal? ChangePercent);

    private sealed record OrderedRateLine(int Index, RateLine Line);

    private sealed record OverviewRateGroup(
        int Index,
        string AccountLabel,
        bool ShowAccountLabel,
        List<RateLine> Lines);

    private sealed record PowerPointFormattedLineRateGroupKey(
        string Key,
        string DisplayName,
        int SortOrder,
        PowerPointFormattedLineRateGroupKind Kind);

    private sealed record PowerPointFormattedLineRateGroup(
        string Key,
        string DisplayName,
        int SortOrder,
        PowerPointFormattedLineRateGroupKind Kind,
        int FirstIndex,
        List<RateLine> Lines);

    private enum PowerPointFormattedLineRateGroupKind
    {
        Pooled,
        ExtendedHealth,
        Dental,
        Other
    }

    private sealed record RatesOverviewTier(
        int SortOrder,
        string Name,
        int Lives,
        decimal? Volume,
        decimal? CurrentRate,
        decimal? ProposedRate,
        decimal? NegotiatedRate,
        decimal CurrentPremium,
        decimal ProposedPremium,
        decimal NegotiatedPremium);

    private sealed record PremiumIllustrationTier(
        int SortOrder,
        string Name,
        int Lives,
        decimal? Volume,
        decimal? CurrentRate,
        decimal? ProposedRate,
        decimal? NegotiatedRate,
        decimal? CurrentPremium,
        decimal? ProposedPremium,
        decimal? NegotiatedPremium);

    private sealed record TierAggregate(
        int SortOrder,
        string Name,
        int Lives,
        decimal? Volume,
        decimal CurrentPremium,
        decimal RenewalPremium);

    private enum HealthDentalBenefit
    {
        ExtendedHealthCare,
        Dental
    }
}

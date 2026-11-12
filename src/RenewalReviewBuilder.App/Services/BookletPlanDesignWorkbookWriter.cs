using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed class BookletPlanDesignWorkbookRequest
{
    public string OutputFolder { get; set; } = "";
    public string ClientName { get; set; } = "";
    public string Carrier { get; set; } = "";
    public string PolicyNumber { get; set; } = "";
    public string SourceFilePath { get; set; } = "";
    public string SourceFileName { get; set; } = "";
    public string PlanDesignSource { get; set; } = PlanDesignSources.None;
    public string EffectiveDate { get; set; } = "";
    public DateTime? UpdatedDate { get; set; }
    public IReadOnlyList<PlanDesignFeatureExtraction> Features { get; set; } = [];
}

public sealed class BookletPlanDesignWorkbookResult
{
    public string MasterWorkbookPath { get; set; } = "";
    public string ReviewWorkbookPath { get; set; } = "";
    public int RowCount { get; set; }
    public int ScheduleRowCount { get; set; }
}

public sealed class BenefitsMasterDatabaseWorkbookRequest
{
    public string OutputFolder { get; init; } = "";
    public DateTime GeneratedAt { get; init; } = DateTime.Now;
    public IReadOnlyList<BenefitsMasterDatabaseClientSnapshot> Clients { get; init; } = [];
}

public sealed class BenefitsMasterDatabaseClientSnapshot
{
    public ClientRecord Client { get; init; } = new();
    public IReadOnlyList<PlanDesignRecord> PlanDesignRows { get; init; } = [];
    public IReadOnlyList<BenefitSummaryViewModel> BenefitSummaries { get; init; } = [];
}

public sealed class BenefitsMasterDatabaseWorkbookResult
{
    public string MasterWorkbookPath { get; init; } = "";
    public int ClientCount { get; init; }
    public int PlanDesignRowCount { get; init; }
    public int SummarySheetCount { get; init; }
}

public static class BookletPlanDesignWorkbookWriter
{
    private const string SummaryLogoFileName = "WestlandBenefitsHorizontalColour.png";
    private const int SummaryLogoTopRow = 1;
    private const int SummaryLogoBottomRow = 4;
    private const int SummaryClientNameRow = 6;
    private const int SummarySubtitleRow = 7;
    private const int SummaryEffectiveDateRow = 8;
    private const int SummaryBorderRow = 9;
    private const int SummaryFirstDataRow = 11;

    private static readonly string[] PlanDesignHeaders =
    [
        "ClientID",
        "ClientName",
        "Carrier",
        "PolicyNumber",
        "ClassName",
        "BenefitCategory",
        "Benefit",
        "CoverageDescription",
        "Coinsurance",
        "Maximum",
        "Frequency",
        "WaitingPeriod",
        "TerminationAge",
        "Notes",
        "DentalNormalizationField",
        "RawDentalValue",
        "NormalizedDentalValue",
        "DentalNormalizationNotes",
        "SourceFile",
        "SourceLocation",
        "Confidence",
        "Status",
        "SourcePage"
    ];

    private static readonly string[] BookletReviewHeaders =
    [
        "ClientID",
        "ClientName",
        "Industry",
        "Province",
        "EmployeeCount",
        "Carrier",
        "PolicyNumber",
        "SourceFile",
        "ClassName",
        "ClassDescription",
        "BenefitCategory",
        "BenefitName",
        "BenefitIncluded",
        "BenefitAmountText",
        "BenefitAmountType",
        "BenefitAmountNumeric",
        "EarningsMultiplier",
        "SalaryPercent",
        "CoinsuranceText",
        "CoinsurancePercent",
        "DeductibleText",
        "SingleDeductible",
        "FamilyDeductible",
        "MaximumText",
        "MaximumAmount",
        "MaximumFrequency",
        "LifetimeMaximum",
        "NonEvidenceMaximum",
        "WaitingPeriodText",
        "WaitingPeriodMonths",
        "EliminationPeriodText",
        "EliminationPeriodDays",
        "BenefitPeriod",
        "TerminationAgeText",
        "TerminationAgeNumber",
        "TerminationTrigger",
        "TaxStatus",
        "WaiverOfPremium",
        "DefinitionOfDisability",
        "ReviewFlag",
        "Confidence",
        "LastUpdated",
        "Notes",
        "Benefit",
        "CoverageDescription",
        "Coinsurance",
        "Maximum",
        "Frequency",
        "WaitingPeriod",
        "TerminationAge",
        "LifeBenefitAmount",
        "LifeEarningsMultiplier",
        "ADDBenefitAmount",
        "DependentLifeSpouseAmount",
        "DependentLifeChildAmount",
        "CriticalIllnessAmount",
        "CriticalIllnessPreEx",
        "CriticalIllnessTerminationAge",
        "STDBenefitPercent",
        "STDMaximum",
        "LTDBenefitPercent",
        "LTDMaximum",
        "LTDEliminationPeriodDays",
        "PrescriptionDrugCoinsurance",
        "DrugPlanType",
        "DrugMax",
        "DrugType",
        "HospitalCoverage",
        "ParamedicalStandardMaximum",
        "ParamedicalPsychologyMaximum",
        "VisionMaximum",
        "VisionFrequencyMonths",
        "EyeExamFrequencyMonths",
        "DentalBasicCoinsurance",
        "DentalMajorCoinsurance",
        "DentalOrthoCoinsurance",
        "DentalBasicMaximum",
        "DentalMajorMaximum",
        "DentalOrthoMaximum",
        "DentalRecallFrequency",
        "DentalBasicRecallRaw",
        "DentalBasicRecallNormalized",
        "DentalMajorAnnualMaximumRaw",
        "DentalMajorAnnualMaximumNormalized",
        "DentalOrthoLifetimeMaximumRaw",
        "DentalOrthoLifetimeMaximumNormalized",
        "DentalNormalizationNotes",
        "TravelMaximum",
        "TravelPerTripMaximum",
        "TravelTripDurationDays",
        "EAPIncluded",
        "VirtualCareIncluded",
        "SecondOpinionIncluded"
    ];

    private static readonly string[] ScheduleHeaders =
    [
        "Slide Page",
        "Page Total",
        "Section Order",
        "Client Name",
        "Carrier",
        "Class Name",
        "Section Name",
        "Detail Text",
        "Source File",
        "Notes"
    ];

    public static BookletPlanDesignWorkbookResult Write(BookletPlanDesignWorkbookRequest request)
    {
        Directory.CreateDirectory(request.OutputFolder);
        var masterPath = Path.Combine(request.OutputFolder, "Benefits_Master_Database.xlsx");
        var reviewPath = Path.Combine(request.OutputFolder, "Booklet_Extraction_Review.xlsx");
        var rows = BuildRows(request).ToList();
        var scheduleRows = BuildScheduleRows(request, rows).ToList();

        using (var workbook = new XLWorkbook())
        {
            WriteSummarySheets(workbook, request, rows);
            WritePlanDesignSheet(workbook.AddWorksheet("PlanDesign"), request, rows);
            workbook.SaveAs(masterPath);
        }

        using (var workbook = new XLWorkbook())
        {
            WriteBookletExtractionReviewSheet(workbook.AddWorksheet("Booklet_Extraction_Review"), request, rows);
            WriteScheduleOfBenefitsSheet(workbook.AddWorksheet("Schedule_of_Benefits"), request, scheduleRows);
            WriteSchedulePptExportSheet(workbook.AddWorksheet("Schedule_PPT_Export"), scheduleRows);
            workbook.SaveAs(reviewPath);
        }

        return new BookletPlanDesignWorkbookResult
        {
            MasterWorkbookPath = masterPath,
            ReviewWorkbookPath = reviewPath,
            RowCount = rows.Count,
            ScheduleRowCount = scheduleRows.Count
        };
    }

    public static BenefitsMasterDatabaseWorkbookResult WriteDatabaseSnapshot(
        BenefitsMasterDatabaseWorkbookRequest request)
    {
        Directory.CreateDirectory(request.OutputFolder);
        var masterPath = Path.Combine(request.OutputFolder, "Benefits_Master_Database.xlsx");
        var temporaryPath = Path.Combine(
            request.OutputFolder,
            $".Benefits_Master_Database.{Guid.NewGuid():N}.tmp.xlsx");
        var clients = request.Clients
            .Where(snapshot => snapshot.PlanDesignRows.Count > 0 || snapshot.BenefitSummaries.Count > 0)
            .OrderBy(snapshot => snapshot.Client.ClientName, StringComparer.OrdinalIgnoreCase)
            .ToList();
        var rows = BuildDatabaseRows(clients).ToList();
        var scheduleRows = BuildDatabaseScheduleRows(rows).ToList();
        var summarySheetCount = clients.Sum(snapshot =>
            snapshot.BenefitSummaries.Count(summary => summary.Sections.Count > 0));

        try
        {
            using (var workbook = new XLWorkbook())
            {
                WriteDatabaseSummarySheets(workbook, clients);
                WritePlanDesignSheet(
                    workbook.AddWorksheet("PlanDesign"),
                    new BookletPlanDesignWorkbookRequest { ClientName = "SQLite Benefits Master Database" },
                    rows);
                WriteDatabaseProcessingLogSheet(workbook.AddWorksheet("ProcessingLog"), clients, request.GeneratedAt);
                WriteBookletExtractionReviewSheet(
                    workbook.AddWorksheet("Booklet_Extraction_Review"),
                    new BookletPlanDesignWorkbookRequest { ClientName = "SQLite Benefits Master Database" },
                    rows);
                WriteDatabaseScheduleOfBenefitsSheet(workbook.AddWorksheet("Schedule_of_Benefits"), scheduleRows);
                WriteSchedulePptExportSheet(workbook.AddWorksheet("Schedule_PPT_Export"), scheduleRows);
                workbook.SaveAs(temporaryPath);
            }

            File.Move(temporaryPath, masterPath, overwrite: true);
        }
        finally
        {
            if (File.Exists(temporaryPath))
            {
                File.Delete(temporaryPath);
            }
        }

        return new BenefitsMasterDatabaseWorkbookResult
        {
            MasterWorkbookPath = masterPath,
            ClientCount = clients.Count,
            PlanDesignRowCount = rows.Count,
            SummarySheetCount = summarySheetCount
        };
    }

    private static IEnumerable<PlanDesignWorkbookRow> BuildDatabaseRows(
        IEnumerable<BenefitsMasterDatabaseClientSnapshot> clients)
    {
        foreach (var snapshot in clients)
        {
            foreach (var row in snapshot.PlanDesignRows)
            {
                yield return new PlanDesignWorkbookRow
                {
                    ClientId = row.ClientId.ToString(CultureInfo.InvariantCulture),
                    ClientName = snapshot.Client.ClientName,
                    Carrier = FirstNonBlank(row.Carrier, snapshot.Client.Carrier),
                    PolicyNumber = snapshot.Client.ClientNumber,
                    ClassName = row.ClassName,
                    BenefitCategory = row.BenefitCategory,
                    BenefitName = row.BenefitName,
                    ProvisionName = FirstNonBlank(row.ProvisionLabel, row.SourceSection, row.BenefitName),
                    CoverageDescription = row.CoverageDescription,
                    Coinsurance = FirstNonBlank(
                        row.ReimbursementText,
                        row.ReimbursementPercent.HasValue
                            ? $"{row.ReimbursementPercent.Value:0.##}%"
                            : ""),
                    Maximum = row.MaximumAmount,
                    Frequency = row.Frequency,
                    WaitingPeriod = row.WaitingPeriod,
                    TerminationAge = row.TerminationAge,
                    Notes = row.Notes,
                    SourceText = row.CoverageDescription,
                    SourceFileName = row.SourceFileName,
                    SourcePage = row.SourcePage,
                    Confidence = 0.80m,
                    Status = "FOUND"
                };
            }
        }
    }

    private static IEnumerable<ScheduleWorkbookRow> BuildDatabaseScheduleRows(
        IReadOnlyCollection<PlanDesignWorkbookRow> rows)
    {
        var groups = rows
            .GroupBy(row => new { row.ClientName, row.Carrier, row.ClassName, row.BenefitCategory })
            .OrderBy(group => group.Key.ClientName, StringComparer.OrdinalIgnoreCase)
            .ThenBy(group => group.Key.ClassName, StringComparer.OrdinalIgnoreCase)
            .ThenBy(group => BenefitSortOrder(group.Key.BenefitCategory))
            .ThenBy(group => group.Key.BenefitCategory, StringComparer.OrdinalIgnoreCase)
            .ToList();
        var pageTotal = Math.Max(1, (int)Math.Ceiling(groups.Count / 8m));
        var sectionOrder = 1;

        foreach (var group in groups)
        {
            var details = string.Join(Environment.NewLine, group
                .OrderBy(row => row.ProvisionName, StringComparer.OrdinalIgnoreCase)
                .Select(row => $"{row.ProvisionName}: {row.CoverageDescription}")
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Distinct(StringComparer.OrdinalIgnoreCase));

            yield return new ScheduleWorkbookRow
            {
                SlidePage = Math.Max(1, (int)Math.Ceiling(sectionOrder / 8m)),
                PageTotal = pageTotal,
                SectionOrder = sectionOrder,
                ClientName = group.Key.ClientName,
                Carrier = group.Key.Carrier,
                ClassName = group.Key.ClassName,
                SectionName = group.Key.BenefitCategory,
                DetailText = details,
                SourceFile = group.Select(row => row.SourceFileName)
                    .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)) ?? "",
                Notes = BuildScheduleNotes(group)
            };
            sectionOrder++;
        }
    }

    private static void WriteDatabaseSummarySheets(
        XLWorkbook workbook,
        IReadOnlyCollection<BenefitsMasterDatabaseClientSnapshot> clients)
    {
        var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var snapshot in clients)
        {
            foreach (var summary in snapshot.BenefitSummaries.Where(summary => summary.Sections.Count > 0))
            {
                var sheetName = UniqueSheetName(
                    BuildSummarySheetName(snapshot.Client.ClientName, summary.SelectionLabel),
                    usedNames);
                var sheet = workbook.AddWorksheet(sheetName);
                WriteSummaryHeader(sheet, summary);
                var rowNumber = SummaryFirstDataRow;
                foreach (var section in summary.Sections)
                {
                    sheet.Cell(rowNumber, 1).Value = section.SectionTitle;
                    sheet.Cell(rowNumber, 2).Value = section.PlainText;
                    FormatSummarySectionRow(sheet, rowNumber, section);
                    rowNumber++;
                }

                var lastRow = WriteSummaryFooter(sheet, rowNumber + 1, summary);
                ApplySummarySizing(sheet, lastRow);
            }
        }

        if (usedNames.Count == 0)
        {
            var sheet = workbook.AddWorksheet("Summary");
            sheet.Cell(1, 1).Value = "No saved plan design data is available in SQLite.";
        }
    }

    private static void WriteDatabaseProcessingLogSheet(
        IXLWorksheet sheet,
        IReadOnlyCollection<BenefitsMasterDatabaseClientSnapshot> clients,
        DateTime generatedAt)
    {
        var headers = new[] { "GeneratedAt", "Source", "ClientID", "ClientName", "PlanDesignRows", "SummarySheets" };
        FormatTitle(sheet, "PROCESSING_LOG", headers.Length);
        WriteHeaders(sheet, 3, headers);
        var rowNumber = 4;
        foreach (var snapshot in clients)
        {
            WriteRow(sheet, rowNumber, [
                generatedAt.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture),
                "Benefits_Intelligence.db",
                snapshot.Client.ClientId,
                snapshot.Client.ClientName,
                snapshot.PlanDesignRows.Count,
                snapshot.BenefitSummaries.Count(summary => summary.Sections.Count > 0)
            ]);
            rowNumber++;
        }

        FormatTableArea(sheet, Math.Max(rowNumber - 1, 3), headers.Length);
        sheet.SheetView.FreezeRows(3);
        SetWidths(sheet, [22, 28, 14, 38, 18, 18]);
    }

    private static void WriteDatabaseScheduleOfBenefitsSheet(
        IXLWorksheet sheet,
        IReadOnlyCollection<ScheduleWorkbookRow> scheduleRows)
    {
        sheet.Style.Font.FontName = "Arial";
        sheet.Cell(1, 1).Value = "Benefits Master Database";
        sheet.Cell(1, 1).Style.Font.FontSize = 20;
        sheet.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml("#004C55");
        sheet.Range(1, 1, 1, 7).Merge();
        sheet.Cell(2, 1).Value = "Schedule of Benefits | Generated from SQLite";
        sheet.Cell(2, 1).Style.Font.FontSize = 12;
        sheet.Cell(2, 1).Style.Font.FontColor = XLColor.FromHtml("#39B8AD");
        sheet.Range(2, 1, 2, 7).Merge();

        var rowNumber = 4;
        foreach (var clientGroup in scheduleRows
                     .GroupBy(row => row.ClientName)
                     .OrderBy(group => group.Key, StringComparer.OrdinalIgnoreCase))
        {
            sheet.Range(rowNumber, 1, rowNumber, 7).Merge();
            sheet.Cell(rowNumber, 1).Value = clientGroup.Key;
            sheet.Cell(rowNumber, 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#004C55");
            sheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.White;
            sheet.Cell(rowNumber, 1).Style.Font.Bold = true;
            rowNumber++;

            foreach (var classGroup in clientGroup
                         .GroupBy(row => row.ClassName)
                         .OrderBy(group => group.Key, StringComparer.OrdinalIgnoreCase))
            {
                sheet.Cell(rowNumber, 1).Value = $"Class: {classGroup.Key}";
                sheet.Cell(rowNumber, 3).Value = classGroup.Select(row => row.Carrier)
                    .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)) ?? "";
                sheet.Range(rowNumber, 1, rowNumber, 7).Style.Fill.BackgroundColor = XLColor.FromHtml("#D7F0EE");
                sheet.Range(rowNumber, 1, rowNumber, 7).Style.Alignment.WrapText = true;
                rowNumber++;

                foreach (var scheduleRow in classGroup.OrderBy(row => row.SectionOrder))
                {
                    WriteSectionHeader(sheet, rowNumber, scheduleRow.SectionName, 7);
                    rowNumber++;
                    foreach (var detail in SplitDetails(scheduleRow.DetailText))
                    {
                        var parts = detail.Split(':', 2, StringSplitOptions.TrimEntries);
                        sheet.Cell(rowNumber, 1).Value = parts.Length == 2 ? $"{parts[0]}:" : scheduleRow.SectionName;
                        sheet.Cell(rowNumber, 3).Value = parts.Length == 2 ? parts[1] : detail;
                        sheet.Range(rowNumber, 1, rowNumber, 7).Style.Alignment.WrapText = true;
                        sheet.Range(rowNumber, 1, rowNumber, 7).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        sheet.Range(rowNumber, 1, rowNumber, 7).Style.Border.BottomBorderColor = XLColor.FromHtml("#B9E2DF");
                        rowNumber++;
                    }
                }
            }
        }

        if (scheduleRows.Count == 0)
        {
            sheet.Cell(rowNumber, 1).Value = "No saved Schedule of Benefits rows are available in SQLite.";
        }

        sheet.Columns(1, 7).AdjustToContents();
        sheet.Column(1).Width = 28;
        sheet.Column(3).Width = 82;
        sheet.Column(3).Style.Alignment.WrapText = true;
    }

    private static IEnumerable<PlanDesignWorkbookRow> BuildRows(BookletPlanDesignWorkbookRequest request)
    {
        var clientId = BuildClientId(request.ClientName);
        var sourceFileName = FirstNonBlank(request.SourceFileName, Path.GetFileName(request.SourceFilePath));
        foreach (var feature in request.Features)
        {
            var benefit = NormalizeBenefit(feature.Benefit, feature.FeatureName, feature.FeatureValue);
            var description = FirstNonBlank(feature.FeatureValue, feature.SourceText);
            if (string.IsNullOrWhiteSpace(benefit) || string.IsNullOrWhiteSpace(description))
            {
                continue;
            }

            yield return new PlanDesignWorkbookRow
            {
                ClientId = clientId,
                ClientName = request.ClientName,
                Carrier = request.Carrier,
                PolicyNumber = request.PolicyNumber,
                ClassName = BuildClassLabel(feature),
                BenefitCategory = BenefitCategory(benefit),
                BenefitName = benefit,
                ProvisionName = FirstNonBlank(feature.FeatureName, benefit),
                CoverageDescription = description,
                Coinsurance = ResolveCoinsurance(feature, benefit, description),
                Maximum = FirstNonBlank(feature.Maximum, IsMaximumFeature(feature) ? description : ""),
                Frequency = feature.Frequency,
                WaitingPeriod = IsWaitingPeriodFeature(feature) ? description : "",
                TerminationAge = FirstNonBlank(feature.TerminationAge, IsTerminationFeature(feature) ? description : ""),
                Notes = BuildNotes(request.PlanDesignSource, feature),
                SourceText = feature.SourceText,
                SourceFileName = sourceFileName,
                SourcePage = feature.SourcePage,
                Confidence = feature.ConfidenceScore > 0m ? feature.ConfidenceScore : 0.70m,
                Status = feature.ConfidenceScore >= 0.85m ? "FOUND" : "NEEDS REVIEW"
            };
        }
    }

    private static IEnumerable<ScheduleWorkbookRow> BuildScheduleRows(
        BookletPlanDesignWorkbookRequest request,
        IReadOnlyCollection<PlanDesignWorkbookRow> rows)
    {
        var groups = rows
            .GroupBy(row => new { row.ClassName, row.BenefitCategory })
            .OrderBy(group => group.Key.ClassName, StringComparer.OrdinalIgnoreCase)
            .ThenBy(group => BenefitSortOrder(group.Key.BenefitCategory))
            .ThenBy(group => group.Key.BenefitCategory, StringComparer.OrdinalIgnoreCase)
            .ToList();
        var pageTotal = Math.Max(1, (int)Math.Ceiling(groups.Count / 8m));
        var sectionOrder = 1;

        foreach (var group in groups)
        {
            var details = string.Join(Environment.NewLine, group
                .OrderBy(row => row.ProvisionName, StringComparer.OrdinalIgnoreCase)
                .Select(row => $"{row.ProvisionName}: {row.CoverageDescription}")
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Distinct(StringComparer.OrdinalIgnoreCase));

            yield return new ScheduleWorkbookRow
            {
                SlidePage = Math.Max(1, (int)Math.Ceiling(sectionOrder / 8m)),
                PageTotal = pageTotal,
                SectionOrder = sectionOrder,
                ClientName = request.ClientName,
                Carrier = request.Carrier,
                ClassName = group.Key.ClassName,
                SectionName = group.Key.BenefitCategory,
                DetailText = details,
                SourceFile = group.Select(row => row.SourceFileName).FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)) ?? "",
                Notes = BuildScheduleNotes(group)
            };
            sectionOrder++;
        }
    }

    private static void WriteSummarySheets(
        XLWorkbook workbook,
        BookletPlanDesignWorkbookRequest request,
        IReadOnlyCollection<PlanDesignWorkbookRow> rows)
    {
        if (rows.Count == 0)
        {
            var sheet = workbook.AddWorksheet("Summary");
            var emptySummary = new BenefitSummaryViewModel
            {
                ClientName = request.ClientName,
                SummaryTitle = "Employee Benefits Summary",
                EffectiveDate = request.EffectiveDate,
                UpdatedDate = ResolveSummaryUpdatedDate(request),
                UpdatedDateSource = ResolveSummaryUpdatedDateSource(request),
                DisclaimerText = BenefitSummaryDisplayFormatter.DefaultDisclaimer
            };
            WriteSummaryHeader(sheet, emptySummary);
            sheet.Cell(SummaryFirstDataRow, 1).Value = "No plan design rows were extracted.";
            sheet.Range(SummaryFirstDataRow, 1, SummaryFirstDataRow, 2).Merge();
            sheet.Cell(SummaryFirstDataRow, 1).Style.Font.FontColor = XLColor.FromHtml("#4B5563");
            var lastRow = WriteSummaryFooter(sheet, SummaryFirstDataRow + 2, emptySummary);
            ApplySummarySizing(sheet, lastRow);
            return;
        }

        var sourceRows = rows.Select(row => new BenefitSummarySourceRow
        {
            ClassName = row.ClassName,
            Benefit = row.BenefitName,
            Category = row.BenefitCategory,
            ProvisionName = row.ProvisionName,
            RawValue = row.CoverageDescription,
            Coinsurance = row.Coinsurance,
            Maximum = row.Maximum,
            Frequency = row.Frequency,
            WaitingPeriod = row.WaitingPeriod,
            TerminationAge = row.TerminationAge,
            Carrier = row.Carrier,
            PolicyNumber = row.PolicyNumber,
            SourceText = row.SourceText,
            SourceDocument = row.SourceFileName,
            SourcePage = row.SourcePage
        }).ToList();
        foreach (var mapping in TieredReimbursementRuleNormalizer.MapBenefitSummaryRows(sourceRows))
        {
            sourceRows[mapping.SourceIndex].Rules = [mapping.Rule];
        }

        var summaries = BenefitSummaryDisplayFormatter.Build(new BenefitSummaryBuildRequest
        {
            ClientName = request.ClientName,
            Carrier = request.Carrier,
            PolicyNumber = request.PolicyNumber,
            EffectiveDate = request.EffectiveDate,
            UpdatedDate = ResolveSummaryUpdatedDate(request),
            UpdatedDateSource = ResolveSummaryUpdatedDateSource(request),
            Rows = sourceRows
        });

        var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var summary in summaries)
        {
            var sheetName = UniqueSheetName(BuildSummarySheetName(request.ClientName, summary.SelectionLabel), usedNames);
            var sheet = workbook.AddWorksheet(sheetName);
            WriteSummaryHeader(sheet, summary);
            var rowNumber = SummaryFirstDataRow;

            foreach (var section in summary.Sections)
            {
                sheet.Cell(rowNumber, 1).Value = section.SectionTitle;
                sheet.Cell(rowNumber, 2).Value = section.PlainText;
                FormatSummarySectionRow(sheet, rowNumber, section);
                rowNumber++;
            }

            var lastRow = WriteSummaryFooter(sheet, rowNumber + 1, summary);
            ApplySummarySizing(sheet, lastRow);
        }
    }

    private static void WritePlanDesignSheet(
        IXLWorksheet sheet,
        BookletPlanDesignWorkbookRequest request,
        IReadOnlyCollection<PlanDesignWorkbookRow> rows)
    {
        FormatTitle(sheet, "PLANDESIGN", PlanDesignHeaders.Length);
        WriteHeaders(sheet, 3, PlanDesignHeaders);
        var rowNumber = 4;
        foreach (var row in rows
                     .OrderBy(row => row.ClassName, StringComparer.OrdinalIgnoreCase)
                     .ThenBy(row => BenefitSortOrder(row.BenefitCategory))
                     .ThenBy(row => row.BenefitName, StringComparer.OrdinalIgnoreCase)
                     .ThenBy(row => row.ProvisionName, StringComparer.OrdinalIgnoreCase))
        {
            var values = new object?[]
            {
                row.ClientId,
                row.ClientName,
                row.Carrier,
                row.PolicyNumber,
                row.ClassName,
                row.BenefitCategory,
                row.BenefitName,
                row.CoverageDescription,
                row.Coinsurance,
                row.Maximum,
                row.Frequency,
                row.WaitingPeriod,
                row.TerminationAge,
                row.Notes,
                "",
                "",
                "",
                "",
                row.SourceFileName,
                SourceLocation(row),
                ConfidenceText(row.Confidence),
                row.Status,
                row.SourcePage
            };
            WriteRow(sheet, rowNumber, values);
            FormatStatusCell(sheet.Cell(rowNumber, 22), row.Status);
            rowNumber++;
        }

        if (rows.Count == 0)
        {
            sheet.Cell(rowNumber, 1).Value = BuildClientId(request.ClientName);
            sheet.Cell(rowNumber, 2).Value = request.ClientName;
            sheet.Cell(rowNumber, 3).Value = request.Carrier;
            sheet.Cell(rowNumber, 22).Value = "NOT AVAILABLE";
            FormatStatusCell(sheet.Cell(rowNumber, 22), "NOT AVAILABLE");
            rowNumber++;
        }

        FormatTableArea(sheet, Math.Max(rowNumber - 1, 3), PlanDesignHeaders.Length);
        sheet.SheetView.FreezeRows(3);
        SetWidths(sheet, [24, 34, 24, 18, 24, 24, 26, 58, 18, 24, 20, 24, 24, 34, 24, 24, 26, 34, 34, 18, 14, 16, 12]);
    }

    private static void WriteBookletExtractionReviewSheet(
        IXLWorksheet sheet,
        BookletPlanDesignWorkbookRequest request,
        IReadOnlyCollection<PlanDesignWorkbookRow> rows)
    {
        FormatTitle(sheet, "BOOKLET_EXTRACTION_REVIEW", BookletReviewHeaders.Length);
        WriteHeaders(sheet, 3, BookletReviewHeaders);
        var rowNumber = 4;
        foreach (var row in rows
                     .OrderBy(row => row.ClassName, StringComparer.OrdinalIgnoreCase)
                     .ThenBy(row => BenefitSortOrder(row.BenefitCategory))
                     .ThenBy(row => row.BenefitName, StringComparer.OrdinalIgnoreCase)
                     .ThenBy(row => row.ProvisionName, StringComparer.OrdinalIgnoreCase))
        {
            var values = BuildBookletReviewValues(row);
            WriteRow(sheet, rowNumber, BookletReviewHeaders.Select(header => values.TryGetValue(header, out var value) ? value : null));
            FormatStatusCell(sheet.Cell(rowNumber, 40), values.TryGetValue("ReviewFlag", out var reviewFlag) && string.Equals(reviewFlag?.ToString(), "Yes", StringComparison.OrdinalIgnoreCase)
                ? "NEEDS REVIEW"
                : "FOUND");
            rowNumber++;
        }

        if (rows.Count == 0)
        {
            var values = new Dictionary<string, object?>
            {
                ["ClientID"] = BuildClientId(request.ClientName),
                ["ClientName"] = request.ClientName,
                ["Carrier"] = request.Carrier,
                ["BenefitIncluded"] = "No",
                ["ReviewFlag"] = "Yes",
                ["Confidence"] = "LOW",
                ["LastUpdated"] = DateTime.Now.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
                ["Notes"] = "No Schedule of Benefits rows were extracted."
            };
            WriteRow(sheet, rowNumber, BookletReviewHeaders.Select(header => values.TryGetValue(header, out var value) ? value : null));
            FormatStatusCell(sheet.Cell(rowNumber, 40), "NEEDS REVIEW");
            rowNumber++;
        }

        FormatTableArea(sheet, Math.Max(rowNumber - 1, 3), BookletReviewHeaders.Length);
        sheet.SheetView.FreezeRows(3);
        SetWidths(sheet, BuildBookletReviewWidths());
    }

    private static void WriteScheduleOfBenefitsSheet(
        IXLWorksheet sheet,
        BookletPlanDesignWorkbookRequest request,
        IReadOnlyCollection<ScheduleWorkbookRow> scheduleRows)
    {
        sheet.Style.Font.FontName = "Arial";
        sheet.Cell(1, 1).Value = request.ClientName;
        sheet.Cell(1, 1).Style.Font.FontSize = 20;
        sheet.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml("#004C55");
        sheet.Range(1, 1, 1, 7).Merge();

        var pageTotal = scheduleRows.Count == 0 ? 1 : scheduleRows.Max(row => row.PageTotal);
        sheet.Cell(2, 1).Value = $"Schedule of Benefits | 1 of {pageTotal}";
        sheet.Cell(2, 1).Style.Font.FontSize = 12;
        sheet.Cell(2, 1).Style.Font.FontColor = XLColor.FromHtml("#39B8AD");
        sheet.Range(2, 1, 2, 7).Merge();

        var rowNumber = 4;
        foreach (var classGroup in scheduleRows
                     .GroupBy(row => row.ClassName)
                     .OrderBy(group => group.Key, StringComparer.OrdinalIgnoreCase))
        {
            sheet.Range(rowNumber, 1, rowNumber, 7).Merge();
            sheet.Cell(rowNumber, 1).Value = request.ClientName;
            sheet.Cell(rowNumber, 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#004C55");
            sheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.White;
            sheet.Cell(rowNumber, 1).Style.Font.Bold = true;
            rowNumber++;

            sheet.Cell(rowNumber, 1).Value = $"Class: {classGroup.Key}";
            sheet.Cell(rowNumber, 3).Value = request.Carrier;
            sheet.Range(rowNumber, 1, rowNumber, 7).Style.Fill.BackgroundColor = XLColor.FromHtml("#D7F0EE");
            sheet.Range(rowNumber, 1, rowNumber, 7).Style.Alignment.WrapText = true;
            rowNumber++;

            foreach (var scheduleRow in classGroup.OrderBy(row => row.SectionOrder))
            {
                WriteSectionHeader(sheet, rowNumber, scheduleRow.SectionName, 7);
                rowNumber++;
                foreach (var detail in SplitDetails(scheduleRow.DetailText))
                {
                    var parts = detail.Split(':', 2, StringSplitOptions.TrimEntries);
                    sheet.Cell(rowNumber, 1).Value = parts.Length == 2 ? $"{parts[0]}:" : scheduleRow.SectionName;
                    sheet.Cell(rowNumber, 3).Value = parts.Length == 2 ? parts[1] : detail;
                    sheet.Range(rowNumber, 1, rowNumber, 7).Style.Alignment.WrapText = true;
                    sheet.Range(rowNumber, 1, rowNumber, 7).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                    sheet.Range(rowNumber, 1, rowNumber, 7).Style.Border.BottomBorderColor = XLColor.FromHtml("#B9E2DF");
                    rowNumber++;
                }
            }
        }

        if (scheduleRows.Count == 0)
        {
            sheet.Cell(rowNumber, 1).Value = "No Schedule of Benefits rows were extracted.";
        }

        sheet.Columns(1, 7).AdjustToContents();
        sheet.Column(1).Width = 28;
        sheet.Column(3).Width = 82;
        sheet.Column(3).Style.Alignment.WrapText = true;
    }

    private static void WriteSchedulePptExportSheet(
        IXLWorksheet sheet,
        IReadOnlyCollection<ScheduleWorkbookRow> rows)
    {
        FormatTitle(sheet, "SCHEDULE_PPT_EXPORT", ScheduleHeaders.Length);
        WriteHeaders(sheet, 3, ScheduleHeaders);
        var rowNumber = 4;
        foreach (var row in rows.OrderBy(row => row.SlidePage).ThenBy(row => row.SectionOrder))
        {
            WriteRow(sheet, rowNumber, [
                row.SlidePage,
                row.PageTotal,
                row.SectionOrder,
                row.ClientName,
                row.Carrier,
                row.ClassName,
                row.SectionName,
                row.DetailText,
                row.SourceFile,
                row.Notes
            ]);
            rowNumber++;
        }

        if (rowNumber == 4)
        {
            WriteRow(sheet, rowNumber, [1, 1, 1, "", "", "", "Schedule of Benefits", "No Schedule of Benefits rows were extracted.", "", ""]);
            rowNumber++;
        }

        var range = sheet.Range(3, 1, Math.Max(4, rowNumber - 1), ScheduleHeaders.Length);
        var table = range.CreateTable("tblPPT_ScheduleOfBenefits");
        table.Theme = XLTableTheme.TableStyleMedium4;
        FormatTableArea(sheet, Math.Max(rowNumber - 1, 3), ScheduleHeaders.Length);
        sheet.Column(8).Width = 78;
        sheet.Column(8).Style.Alignment.WrapText = true;
        sheet.Column(10).Width = 38;
        sheet.Column(10).Style.Alignment.WrapText = true;
        sheet.SheetView.FreezeRows(3);
    }

    private static Dictionary<string, object?> BuildBookletReviewValues(PlanDesignWorkbookRow row)
    {
        var description = row.CoverageDescription;
        var values = new Dictionary<string, object?>
        {
            ["ClientID"] = row.ClientId,
            ["ClientName"] = row.ClientName,
            ["Carrier"] = row.Carrier,
            ["PolicyNumber"] = row.PolicyNumber,
            ["SourceFile"] = row.SourceFileName,
            ["ClassName"] = row.ClassName,
            ["BenefitCategory"] = row.BenefitCategory,
            ["BenefitName"] = row.BenefitName,
            ["BenefitIncluded"] = description.Contains("not eligible", StringComparison.OrdinalIgnoreCase) ? "No" : "Yes",
            ["ReviewFlag"] = string.Equals(row.Status, "FOUND", StringComparison.OrdinalIgnoreCase) ? "No" : "Yes",
            ["Confidence"] = ConfidenceText(row.Confidence),
            ["LastUpdated"] = DateTime.Now.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
            ["Notes"] = row.Notes,
            ["Benefit"] = row.BenefitName,
            ["CoverageDescription"] = description,
            ["Coinsurance"] = row.Coinsurance,
            ["Maximum"] = row.Maximum,
            ["Frequency"] = row.Frequency,
            ["WaitingPeriod"] = row.WaitingPeriod,
            ["TerminationAge"] = row.TerminationAge
        };

        if (row.ProvisionName.Contains("Benefit", StringComparison.OrdinalIgnoreCase) ||
            row.ProvisionName.Contains("Amount", StringComparison.OrdinalIgnoreCase))
        {
            values["BenefitAmountText"] = description;
            values["BenefitAmountNumeric"] = ExtractFirstMoney(description);
        }

        if (!string.IsNullOrWhiteSpace(row.Coinsurance))
        {
            values["CoinsuranceText"] = description;
            values["CoinsurancePercent"] = ExtractPercentNumber(row.Coinsurance);
        }

        if (row.ProvisionName.Contains("Deductible", StringComparison.OrdinalIgnoreCase) ||
            description.Contains("deductible", StringComparison.OrdinalIgnoreCase))
        {
            values["DeductibleText"] = description;
        }

        if (!string.IsNullOrWhiteSpace(row.Maximum))
        {
            values["MaximumText"] = description;
            values["MaximumAmount"] = ExtractFirstMoney(row.Maximum);
            values["MaximumFrequency"] = row.Frequency;
        }

        if (row.ProvisionName.Contains("Non-Evidence", StringComparison.OrdinalIgnoreCase) ||
            description.Contains("non-evidence", StringComparison.OrdinalIgnoreCase))
        {
            values["NonEvidenceMaximum"] = description;
        }

        if (!string.IsNullOrWhiteSpace(row.WaitingPeriod))
        {
            values["WaitingPeriodText"] = row.WaitingPeriod;
            values["WaitingPeriodMonths"] = ExtractMonthCount(row.WaitingPeriod);
        }

        if (description.Contains("elimination period", StringComparison.OrdinalIgnoreCase))
        {
            values["EliminationPeriodText"] = description;
            values["EliminationPeriodDays"] = ExtractDayCount(description);
        }

        if (!string.IsNullOrWhiteSpace(row.TerminationAge))
        {
            values["TerminationAgeText"] = row.TerminationAge;
            values["TerminationAgeNumber"] = ExtractAgeNumber(row.TerminationAge);
        }

        ApplyBenefitSpecificValues(values, row);
        return values;
    }

    private static void ApplyBenefitSpecificValues(Dictionary<string, object?> values, PlanDesignWorkbookRow row)
    {
        var benefit = row.BenefitName;
        var category = row.BenefitCategory;
        var description = row.CoverageDescription;

        if (benefit.Equals("Employee Life Insurance", StringComparison.OrdinalIgnoreCase) &&
            values.TryGetValue("BenefitAmountText", out var lifeText))
        {
            values["LifeBenefitAmount"] = lifeText;
        }

        if (benefit.Equals("Employee AD&D", StringComparison.OrdinalIgnoreCase) &&
            values.TryGetValue("BenefitAmountText", out var addText))
        {
            values["ADDBenefitAmount"] = addText;
        }

        if (category.Contains("Long-Term Disability", StringComparison.OrdinalIgnoreCase))
        {
            if (description.Contains('%'))
            {
                values["LTDBenefitPercent"] = ExtractPercentText(description);
            }

            if (!string.IsNullOrWhiteSpace(row.Maximum))
            {
                values["LTDMaximum"] = row.Maximum;
            }

            if (description.Contains("elimination", StringComparison.OrdinalIgnoreCase))
            {
                values["LTDEliminationPeriodDays"] = ExtractDayCount(description);
            }
        }
        else if (category.Contains("Prescription Drugs", StringComparison.OrdinalIgnoreCase))
        {
            values["PrescriptionDrugCoinsurance"] = row.Coinsurance;
            values["DrugMax"] = row.Maximum;
            if (description.Contains("generic", StringComparison.OrdinalIgnoreCase))
            {
                values["DrugType"] = "Mandatory Generic";
            }
        }
        else if (category.Contains("Paramedical", StringComparison.OrdinalIgnoreCase))
        {
            values["ParamedicalStandardMaximum"] = FirstNonBlank(row.Maximum, description);
        }
        else if (description.Contains("psycholog", StringComparison.OrdinalIgnoreCase) ||
                 description.Contains("counsel", StringComparison.OrdinalIgnoreCase))
        {
            values["ParamedicalPsychologyMaximum"] = FirstNonBlank(row.Maximum, description);
        }
        else if (category.Contains("Vision", StringComparison.OrdinalIgnoreCase))
        {
            values["VisionMaximum"] = row.Maximum;
            values["VisionFrequencyMonths"] = ExtractMonthCount(row.Frequency);
            if (description.Contains("eye exam", StringComparison.OrdinalIgnoreCase))
            {
                values["EyeExamFrequencyMonths"] = ExtractMonthCount(row.Frequency);
            }
        }
        else if (category.Contains("Dental", StringComparison.OrdinalIgnoreCase))
        {
            if (benefit.Contains("Basic", StringComparison.OrdinalIgnoreCase))
            {
                values["DentalBasicCoinsurance"] = row.Coinsurance;
                values["DentalBasicMaximum"] = row.Maximum;
            }
            else if (benefit.Contains("Major", StringComparison.OrdinalIgnoreCase))
            {
                values["DentalMajorCoinsurance"] = row.Coinsurance;
                values["DentalMajorMaximum"] = row.Maximum;
            }
            else if (benefit.Contains("Ortho", StringComparison.OrdinalIgnoreCase))
            {
                values["DentalOrthoCoinsurance"] = row.Coinsurance;
                values["DentalOrthoMaximum"] = row.Maximum;
            }

            if (description.Contains("recall", StringComparison.OrdinalIgnoreCase))
            {
                values["DentalRecallFrequency"] = FirstNonBlank(row.Frequency, description);
            }
        }

        if (category.Contains("Travel", StringComparison.OrdinalIgnoreCase))
        {
            if (row.ProvisionName.Contains("Per-Trip", StringComparison.OrdinalIgnoreCase) ||
                row.ProvisionName.Contains("Per Trip", StringComparison.OrdinalIgnoreCase))
            {
                values["TravelPerTripMaximum"] = FirstNonBlank(row.Maximum, description);
            }
            else if (row.ProvisionName.Contains("Maximum", StringComparison.OrdinalIgnoreCase))
            {
                values["TravelMaximum"] = FirstNonBlank(row.Maximum, description);
            }

            if (row.ProvisionName.Contains("Trip Duration", StringComparison.OrdinalIgnoreCase))
            {
                values["TravelTripDurationDays"] = ExtractDayCount(description);
            }
        }
    }

    private static string NormalizeBenefit(string benefit, string featureName, string featureValue)
    {
        var text = $"{benefit} {featureName} {featureValue}";
        if (text.Contains("survivor benefit", StringComparison.OrdinalIgnoreCase))
        {
            return "Survivor Benefit";
        }

        if (text.Contains("health care spending account", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("healthcare spending account", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("health spending account", StringComparison.OrdinalIgnoreCase) ||
            Regex.IsMatch(text, @"\b(?:HCSA|HSCA|HSA)\b", RegexOptions.IgnoreCase))
        {
            return "Health Care Spending Account";
        }

        if (text.Contains("lifestyle spending account", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("life style spending account", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("wellness spending account", StringComparison.OrdinalIgnoreCase) ||
            Regex.IsMatch(text, @"\b(?:LSA|WSA)\b", RegexOptions.IgnoreCase))
        {
            return "Lifestyle Spending Account";
        }

        if (text.Contains("employee assistance", StringComparison.OrdinalIgnoreCase) ||
            Regex.IsMatch(text, @"\bEAP\b|\bEFAP\b", RegexOptions.IgnoreCase))
        {
            return "Employee Assistance Program";
        }

        if (text.Contains("virtual health", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("virtual care", StringComparison.OrdinalIgnoreCase))
        {
            return "Virtual Health Care";
        }

        if (text.Contains("digital wellness", StringComparison.OrdinalIgnoreCase))
        {
            return "Digital Wellness Program";
        }

        if (text.Contains("critical illness", StringComparison.OrdinalIgnoreCase))
        {
            return "Critical Illness";
        }

        if (text.Contains("emergency travel", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("travel assist", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("out of province", StringComparison.OrdinalIgnoreCase))
        {
            return "Emergency Travel";
        }

        if (Regex.IsMatch(benefit, @"^\s*Life\s*/\s*AD\s*&\s*D\s*$", RegexOptions.IgnoreCase) &&
            !Regex.IsMatch(featureName, @"\b(?:employee\s+)?(?:basic\s+)?life\b|\bAD\s*&\s*D\b|accidental death", RegexOptions.IgnoreCase))
        {
            return "Life / AD&D";
        }

        if (Regex.IsMatch(featureName, @"\b(?:dep(?:endent|endant)?|dependant)\s+life\b", RegexOptions.IgnoreCase))
        {
            return "Dependent Life";
        }

        if (featureName.Contains("ad&d", StringComparison.OrdinalIgnoreCase) ||
            featureName.Contains("accidental death", StringComparison.OrdinalIgnoreCase))
        {
            return "Employee AD&D";
        }

        if (Regex.IsMatch(featureName, @"\blife\b", RegexOptions.IgnoreCase))
        {
            return "Employee Life Insurance";
        }

        if (text.Contains("ad&d", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("accidental death", StringComparison.OrdinalIgnoreCase))
        {
            return "Employee AD&D";
        }

        if (Regex.IsMatch(text, @"\b(?:dep(?:endent|endant)?|dependant)\s+life\b", RegexOptions.IgnoreCase))
        {
            return "Dependent Life";
        }

        if (text.Contains("long", StringComparison.OrdinalIgnoreCase) &&
            text.Contains("disability", StringComparison.OrdinalIgnoreCase))
        {
            return "Long Term Disability";
        }

        if (text.Contains("short", StringComparison.OrdinalIgnoreCase) &&
            text.Contains("disability", StringComparison.OrdinalIgnoreCase))
        {
            return "Short Term Disability";
        }

        if (text.Contains("drug", StringComparison.OrdinalIgnoreCase) ||
            benefit.Equals("Drugs", StringComparison.OrdinalIgnoreCase))
        {
            return "Prescription Drugs";
        }

        if (text.Contains("orthodont", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("dental", StringComparison.OrdinalIgnoreCase))
        {
            return "Dental Care";
        }

        if (text.Contains("vision", StringComparison.OrdinalIgnoreCase))
        {
            return "Vision Care";
        }

        if (Regex.IsMatch(text, @"\bpara\s*medical", RegexOptions.IgnoreCase))
        {
            return "Paramedical";
        }

        if (Regex.IsMatch(text, @"mental health|psychological services", RegexOptions.IgnoreCase))
        {
            return "Mental Health Practitioners";
        }

        if (benefit.Equals("EHC", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("health", StringComparison.OrdinalIgnoreCase))
        {
            return "Extended Health Care";
        }

        if (Regex.IsMatch(text, @"\blife\b", RegexOptions.IgnoreCase))
        {
            return "Employee Life Insurance";
        }

        return FirstNonBlank(benefit, "Schedule of Benefits");
    }

    private static string BenefitCategory(string benefit)
    {
        if (benefit.Equals("Employee Life Insurance", StringComparison.OrdinalIgnoreCase) ||
            benefit.Equals("Employee AD&D", StringComparison.OrdinalIgnoreCase) ||
            benefit.Equals("Dependent Life", StringComparison.OrdinalIgnoreCase))
        {
            return "Life / AD&D";
        }

        if (benefit.Equals("Long Term Disability", StringComparison.OrdinalIgnoreCase))
        {
            return "Long-Term Disability";
        }

        if (benefit.Equals("Short Term Disability", StringComparison.OrdinalIgnoreCase))
        {
            return "Short-Term Disability";
        }

        if (benefit.Equals("Prescription Drugs", StringComparison.OrdinalIgnoreCase))
        {
            return "Extended Health Care";
        }

        if (benefit.Equals("Vision Care", StringComparison.OrdinalIgnoreCase))
        {
            return "Vision Care";
        }

        if (benefit.Equals("Dental Care", StringComparison.OrdinalIgnoreCase))
        {
            return "Dental Care";
        }

        if (benefit.Equals("Paramedical", StringComparison.OrdinalIgnoreCase))
        {
            return "Paramedical";
        }

        if (benefit.Equals("Extended Health Care", StringComparison.OrdinalIgnoreCase))
        {
            return "Extended Health Care";
        }

        if (benefit.Equals("Survivor Benefit", StringComparison.OrdinalIgnoreCase) ||
            benefit.Equals("Health Care Spending Account", StringComparison.OrdinalIgnoreCase) ||
            benefit.Equals("Lifestyle Spending Account", StringComparison.OrdinalIgnoreCase) ||
            benefit.Equals("Employee Assistance Program", StringComparison.OrdinalIgnoreCase) ||
            benefit.Equals("Critical Illness", StringComparison.OrdinalIgnoreCase) ||
            benefit.Equals("Emergency Travel", StringComparison.OrdinalIgnoreCase))
        {
            return benefit;
        }

        return benefit;
    }

    private static int BenefitSortOrder(string benefit)
    {
        var text = benefit.ToLowerInvariant();
        if (text.Contains("survivor"))
        {
            return 90;
        }

        if (text.Contains("health care spending") || text.Contains("health spending"))
        {
            return 100;
        }

        if (text.Contains("lifestyle spending") || text.Contains("wellness spending"))
        {
            return 110;
        }

        if (text.Contains("employee assistance") || text.Equals("eap", StringComparison.Ordinal))
        {
            return 120;
        }

        if (text.Contains("virtual care"))
        {
            return 130;
        }

        if (text.Contains("life"))
        {
            return 10;
        }

        if (text.Contains("long-term") || text.Contains("long term"))
        {
            return 20;
        }

        if (text.Contains("short-term") || text.Contains("short term"))
        {
            return 30;
        }

        if (text.Contains("drug"))
        {
            return 40;
        }

        if (text.Contains("critical illness"))
        {
            return 35;
        }

        if (text.Contains("health"))
        {
            return 50;
        }

        if (text.Contains("paramedical"))
        {
            return 60;
        }

        if (text.Contains("vision"))
        {
            return 70;
        }

        if (text.Contains("travel"))
        {
            return 75;
        }

        if (text.Contains("dental"))
        {
            return 80;
        }

        return 999;
    }

    private static string BuildClassLabel(PlanDesignFeatureExtraction feature)
    {
        var classNumber = NormalizeClassNumber(feature.ClassNumber);
        if (!string.IsNullOrWhiteSpace(feature.ClassName) && !string.IsNullOrWhiteSpace(classNumber))
        {
            return $"Class {classNumber} - {feature.ClassName}";
        }

        return FirstNonBlank(feature.ClassName, string.IsNullOrWhiteSpace(classNumber) ? "" : $"Class {classNumber}", "Class 01");
    }

    private static string NormalizeClassNumber(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return "";
        }

        var digits = new string(value.Where(char.IsDigit).ToArray());
        return string.IsNullOrWhiteSpace(digits)
            ? value.Trim()
            : int.Parse(digits, CultureInfo.InvariantCulture).ToString("00", CultureInfo.InvariantCulture);
    }

    private static string BuildClientId(string clientName)
    {
        var code = new string(clientName.Where(char.IsLetterOrDigit).Take(12).ToArray()).ToUpperInvariant();
        return string.IsNullOrWhiteSpace(code) ? "CLIENT" : code;
    }

    private static string BuildSummarySheetName(string clientName, string className)
    {
        var clientCode = new string(clientName.Where(char.IsLetterOrDigit).Take(6).ToArray()).ToUpperInvariant();
        if (string.IsNullOrWhiteSpace(clientCode))
        {
            clientCode = "CLIENT";
        }

        var selectionCode = Regex.Replace(className, @"^Class\s+", "", RegexOptions.IgnoreCase);
        selectionCode = Regex.Replace(selectionCode, @"[^A-Za-z0-9]+", "_").Trim('_');
        return TrimSheetName($"Summary_{clientCode}_{FirstNonBlank(selectionCode, "01")}");
    }

    private static string UniqueSheetName(string sheetName, HashSet<string> usedNames)
    {
        var candidate = TrimSheetName(sheetName);
        if (usedNames.Add(candidate))
        {
            return candidate;
        }

        for (var index = 2; index < 100; index++)
        {
            var suffix = $"_{index}";
            var baseName = TrimSheetName(sheetName);
            candidate = $"{baseName[..Math.Min(baseName.Length, 31 - suffix.Length)]}{suffix}";
            if (usedNames.Add(candidate))
            {
                return candidate;
            }
        }

        throw new InvalidOperationException($"Unable to create a unique worksheet name from '{sheetName}'.");
    }

    private static string TrimSheetName(string sheetName)
    {
        var invalid = new[] { ':', '\\', '/', '?', '*', '[', ']' };
        var cleaned = new string(sheetName.Select(ch => invalid.Contains(ch) ? '-' : ch).ToArray()).Trim();
        if (string.IsNullOrWhiteSpace(cleaned))
        {
            cleaned = "Sheet";
        }

        return cleaned.Length <= 31 ? cleaned : cleaned[..31];
    }

    private static void WriteSummaryHeader(IXLWorksheet sheet, BenefitSummaryViewModel summary)
    {
        sheet.Style.Font.FontName = "Arial";
        sheet.Style.Fill.BackgroundColor = XLColor.White;
        sheet.ShowGridLines = false;
        AddSummaryLogo(sheet);
        SetSummaryLogoRowHeights(sheet);

        sheet.Cell(SummaryClientNameRow, 1).Value = summary.ClientName;
        sheet.Cell(SummaryClientNameRow, 1).Style.Font.FontSize = 14;
        sheet.Cell(SummaryClientNameRow, 1).Style.Font.Bold = true;
        sheet.Cell(SummaryClientNameRow, 1).Style.Font.FontColor = XLColor.FromHtml("#004C55");
        sheet.Range(SummaryClientNameRow, 1, SummaryClientNameRow, 2).Merge();

        sheet.Cell(SummarySubtitleRow, 1).Value = summary.SummaryTitle;
        sheet.Cell(SummarySubtitleRow, 1).Style.Font.FontSize = 11;
        sheet.Cell(SummarySubtitleRow, 1).Style.Font.FontColor = XLColor.FromHtml("#39B8AD");
        sheet.Range(SummarySubtitleRow, 1, SummarySubtitleRow, 2).Merge();

        if (!string.IsNullOrWhiteSpace(summary.EffectiveDateText))
        {
            sheet.Cell(SummaryEffectiveDateRow, 1).Value = summary.EffectiveDateText;
            sheet.Cell(SummaryEffectiveDateRow, 1).Style.Font.FontSize = 9;
            sheet.Cell(SummaryEffectiveDateRow, 1).Style.Font.FontColor = XLColor.FromHtml("#4B5563");
            sheet.Range(SummaryEffectiveDateRow, 1, SummaryEffectiveDateRow, 2).Merge();
        }

        sheet.Range(SummaryBorderRow, 1, SummaryBorderRow, 2).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
        sheet.Range(SummaryBorderRow, 1, SummaryBorderRow, 2).Style.Border.BottomBorderColor = XLColor.FromHtml("#6B7280");
    }

    private static void ApplySummarySizing(IXLWorksheet sheet, int lastRow)
    {
        sheet.Column(1).Width = 31;
        sheet.Column(1).Style.Alignment.WrapText = true;
        sheet.Column(2).Width = 88;
        sheet.Column(2).Style.Alignment.WrapText = true;
        sheet.Rows(SummaryClientNameRow, Math.Max(lastRow, SummaryClientNameRow)).Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
        sheet.Row(SummaryClientNameRow).Height = 22;
        sheet.Row(SummarySubtitleRow).Height = 19;
        sheet.Row(SummaryEffectiveDateRow).Height = 16;
        sheet.Row(SummaryBorderRow).Height = 8;
        sheet.Row(SummaryFirstDataRow - 1).Height = 6;
        sheet.PageSetup.PrintAreas.Clear();
        sheet.PageSetup.PrintAreas.Add($"A1:B{lastRow}");
        sheet.PageSetup.PageOrientation = XLPageOrientation.Portrait;
        sheet.PageSetup.FitToPages(1, 0);
        sheet.PageSetup.Margins.Top = 0.35;
        sheet.PageSetup.Margins.Bottom = 0.35;
        sheet.PageSetup.Margins.Left = 0.35;
        sheet.PageSetup.Margins.Right = 0.35;
        sheet.PageSetup.Margins.Header = 0.15;
        sheet.PageSetup.Margins.Footer = 0.15;
        sheet.PageSetup.SetRowsToRepeatAtTop(SummaryClientNameRow, SummaryBorderRow);
        SetSummaryLogoRowHeights(sheet);
    }

    private static void FormatSummarySectionRow(
        IXLWorksheet sheet,
        int rowNumber,
        BenefitSummarySectionViewModel section)
    {
        var range = sheet.Range(rowNumber, 1, rowNumber, 2);
        range.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
        range.Style.Border.BottomBorderColor = XLColor.FromHtml("#C7DAD8");
        range.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
        range.Style.Alignment.WrapText = true;
        range.Style.Fill.BackgroundColor = XLColor.White;
        sheet.Cell(rowNumber, 1).Style.Font.Bold = true;
        sheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.FromHtml("#004C55");
        sheet.Cell(rowNumber, 1).Style.Font.FontSize = 9.5;
        sheet.Cell(rowNumber, 2).Style.Font.FontColor = XLColor.FromHtml("#4B5563");
        sheet.Cell(rowNumber, 2).Style.Font.FontSize = 9;

        var detailLines = section.PlainText
            .Replace("\r\n", "\n", StringComparison.Ordinal)
            .Replace('\r', '\n')
            .Split('\n')
            .Sum(line => Math.Max(1, (int)Math.Ceiling(line.Length / 86m)));
        var sectionTitleLines = Math.Max(1, (int)Math.Ceiling(section.SectionTitle.Length / 28m));
        var visibleLines = Math.Max(sectionTitleLines, detailLines);
        sheet.Row(rowNumber).Height = Math.Min(320, Math.Max(20, visibleLines * 13.5 + 6));
    }

    private static int WriteSummaryFooter(
        IXLWorksheet sheet,
        int rowNumber,
        BenefitSummaryViewModel summary)
    {
        if (!string.IsNullOrWhiteSpace(summary.UpdatedDateText))
        {
            sheet.Range(rowNumber, 1, rowNumber, 2).Merge();
            sheet.Cell(rowNumber, 1).Value = summary.UpdatedDateText;
            sheet.Cell(rowNumber, 1).Style.Font.Bold = true;
            sheet.Cell(rowNumber, 1).Style.Font.FontSize = 8;
            sheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.FromHtml("#004C55");
            rowNumber++;
        }

        sheet.Range(rowNumber, 1, rowNumber, 2).Merge();
        sheet.Cell(rowNumber, 1).Value = summary.DisclaimerText;
        sheet.Cell(rowNumber, 1).Style.Font.Italic = true;
        sheet.Cell(rowNumber, 1).Style.Font.FontSize = 7.5;
        sheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.FromHtml("#64748B");
        sheet.Cell(rowNumber, 1).Style.Alignment.WrapText = true;
        sheet.Cell(rowNumber, 1).Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
        sheet.Row(rowNumber).Height = 26;
        return rowNumber;
    }

    private static DateTime ResolveSummaryUpdatedDate(BookletPlanDesignWorkbookRequest request)
    {
        if (request.UpdatedDate.HasValue)
        {
            return request.UpdatedDate.Value;
        }

        return !string.IsNullOrWhiteSpace(request.SourceFilePath) && File.Exists(request.SourceFilePath)
            ? File.GetLastWriteTime(request.SourceFilePath)
            : DateTime.Now;
    }

    private static string ResolveSummaryUpdatedDateSource(BookletPlanDesignWorkbookRequest request)
    {
        if (request.UpdatedDate.HasValue)
        {
            return "Requested plan-design update date";
        }

        return !string.IsNullOrWhiteSpace(request.SourceFilePath) && File.Exists(request.SourceFilePath)
            ? "Source document modified date"
            : "Workbook generated date";
    }

    private static void AddSummaryLogo(IXLWorksheet sheet)
    {
        var logoPath = ResolveSummaryLogoPath();
        if (string.IsNullOrWhiteSpace(logoPath))
        {
            return;
        }

        var picture = sheet.AddPicture(logoPath);
        picture.MoveTo(sheet.Cell(SummaryLogoTopRow, 1));
        picture.Scale(0.5, relativeToOriginal: true);
    }

    private static string ResolveSummaryLogoPath()
    {
        var candidates = new[]
        {
            Path.Combine(AppContext.BaseDirectory, "Assets", SummaryLogoFileName),
            Path.Combine(Directory.GetCurrentDirectory(), "Assets", SummaryLogoFileName),
            Path.Combine(Directory.GetCurrentDirectory(), "src", "RenewalReviewBuilder.App", "Assets", SummaryLogoFileName)
        };

        return candidates.FirstOrDefault(File.Exists) ?? "";
    }

    private static void SetSummaryLogoRowHeights(IXLWorksheet sheet)
    {
        sheet.Row(1).Height = 28;
        sheet.Row(2).Height = 28;
        sheet.Row(3).Height = 28;
        sheet.Row(4).Height = 28;
        sheet.Row(5).Height = 8;
    }

    private static void WriteSectionHeader(IXLWorksheet sheet, int rowNumber, string sectionName, int columnCount)
    {
        sheet.Range(rowNumber, 1, rowNumber, columnCount).Merge();
        sheet.Cell(rowNumber, 1).Value = sectionName;
        sheet.Cell(rowNumber, 1).Style.Font.Bold = true;
        sheet.Cell(rowNumber, 1).Style.Font.FontColor = XLColor.White;
        sheet.Cell(rowNumber, 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#66C5B4");
    }

    private static void FormatTitle(IXLWorksheet sheet, string title, int columnCount)
    {
        sheet.Style.Font.FontName = "Arial";
        sheet.Cell(1, 1).Value = title;
        sheet.Range(1, 1, 1, columnCount).Merge();
        sheet.Cell(1, 1).Style.Font.Bold = true;
        sheet.Cell(1, 1).Style.Font.FontSize = 16;
        sheet.Cell(1, 1).Style.Font.FontColor = XLColor.FromHtml("#004C55");
    }

    private static void WriteHeaders(IXLWorksheet sheet, int rowNumber, IReadOnlyList<string> headers)
    {
        for (var index = 0; index < headers.Count; index++)
        {
            var cell = sheet.Cell(rowNumber, index + 1);
            cell.Value = headers[index];
            cell.Style.Font.Bold = true;
            cell.Style.Font.FontColor = XLColor.White;
            cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#004C55");
            cell.Style.Alignment.WrapText = true;
        }
    }

    private static void WriteRow(IXLWorksheet sheet, int rowNumber, IEnumerable<object?> values)
    {
        var column = 1;
        foreach (var value in values)
        {
            sheet.Cell(rowNumber, column).Value = XLCellValue.FromObject(value);
            column++;
        }
    }

    private static void FormatTableArea(IXLWorksheet sheet, int lastRow, int lastColumn)
    {
        var range = sheet.Range(3, 1, Math.Max(lastRow, 3), lastColumn);
        range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        range.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
        range.Style.Border.OutsideBorderColor = XLColor.FromHtml("#B9E2DF");
        range.Style.Border.InsideBorderColor = XLColor.FromHtml("#B9E2DF");
        range.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
        sheet.Rows(4, Math.Max(lastRow, 4)).Style.Alignment.WrapText = true;
        sheet.Columns().AdjustToContents();
    }

    private static void SetWidths(IXLWorksheet sheet, IReadOnlyList<double> widths)
    {
        for (var index = 0; index < widths.Count; index++)
        {
            sheet.Column(index + 1).Width = widths[index];
        }
    }

    private static double[] BuildBookletReviewWidths()
    {
        var widths = Enumerable.Repeat(16d, BookletReviewHeaders.Length).ToArray();
        SetWidth("ClientID", 24);
        SetWidth("ClientName", 34);
        SetWidth("SourceFile", 34);
        SetWidth("ClassName", 24);
        SetWidth("ClassDescription", 28);
        SetWidth("BenefitCategory", 24);
        SetWidth("BenefitName", 24);
        SetWidth("BenefitAmountText", 38);
        SetWidth("CoinsuranceText", 34);
        SetWidth("DeductibleText", 30);
        SetWidth("MaximumText", 34);
        SetWidth("TerminationAgeText", 38);
        SetWidth("Notes", 34);
        SetWidth("Benefit", 24);
        SetWidth("CoverageDescription", 58);
        SetWidth("DentalNormalizationNotes", 42);
        return widths;

        void SetWidth(string header, double width)
        {
            var index = Array.IndexOf(BookletReviewHeaders, header);
            if (index >= 0)
            {
                widths[index] = width;
            }
        }
    }

    private static void FormatStatusCell(IXLCell cell, string? status)
    {
        if (string.Equals(status, "FOUND", StringComparison.OrdinalIgnoreCase))
        {
            cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#D9EAD3");
        }
        else if (string.Equals(status, "NEEDS REVIEW", StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(status, "NOT AVAILABLE", StringComparison.OrdinalIgnoreCase))
        {
            cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#FFF2CC");
        }
    }

    private static bool IsMaximumFeature(PlanDesignFeatureExtraction feature)
    {
        return feature.FeatureName.Contains("maximum", StringComparison.OrdinalIgnoreCase) ||
               feature.FeatureValue.Contains("maximum", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsWaitingPeriodFeature(PlanDesignFeatureExtraction feature)
    {
        return feature.FeatureName.Contains("waiting", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsTerminationFeature(PlanDesignFeatureExtraction feature)
    {
        return feature.FeatureName.Contains("termination", StringComparison.OrdinalIgnoreCase) ||
               feature.FeatureName.Contains("terminates", StringComparison.OrdinalIgnoreCase);
    }

    private static string BuildNotes(string source, PlanDesignFeatureExtraction feature)
    {
        var parts = new List<string>();
        if (!string.IsNullOrWhiteSpace(source))
        {
            parts.Add(source);
        }

        if (feature.SourcePage is not null)
        {
            parts.Add($"PDF page {feature.SourcePage}");
        }

        return string.Join("; ", parts);
    }

    private static string BuildScheduleNotes(IEnumerable<PlanDesignWorkbookRow> rows)
    {
        var pages = rows
            .Select(row => row.SourcePage)
            .Where(page => page.HasValue)
            .Select(page => page!.Value)
            .Distinct()
            .OrderBy(page => page)
            .Select(page => page.ToString(CultureInfo.InvariantCulture))
            .ToList();
        var source = rows.Select(row => row.Notes.Split(';').FirstOrDefault() ?? "").FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
        var sourceText = string.IsNullOrWhiteSpace(source) ? "" : source;
        return pages.Count == 0
            ? sourceText
            : $"{sourceText}; source page(s): {string.Join(", ", pages)}".Trim(' ', ';');
    }

    private static string SourceLocation(PlanDesignWorkbookRow row)
    {
        return row.SourcePage is null
            ? ""
            : $"PDF page {row.SourcePage.Value.ToString(CultureInfo.InvariantCulture)}";
    }

    private static string ConfidenceText(decimal confidence)
    {
        return confidence >= 0.85m
            ? "HIGH"
            : confidence >= 0.65m
                ? "MEDIUM"
                : "LOW";
    }

    private static string ResolveCoinsurance(
        PlanDesignFeatureExtraction feature,
        string benefit,
        string description)
    {
        return BenefitCategory(benefit).Equals("Life / AD&D", StringComparison.OrdinalIgnoreCase)
            ? ""
            : FirstNonBlank(feature.Coinsurance, ExtractPercentText(description));
    }

    private static string ExtractPercentText(string text)
    {
        return Regex.Match(text, @"\d+(?:\.\d+)?\s*%").Value;
    }

    private static decimal? ExtractPercentNumber(string text)
    {
        var match = Regex.Match(text, @"\d+(?:\.\d+)?");
        return match.Success && decimal.TryParse(match.Value, NumberStyles.Number, CultureInfo.InvariantCulture, out var value)
            ? value / 100m
            : null;
    }

    private static decimal? ExtractFirstMoney(string text)
    {
        var match = Regex.Match(text, @"\$?\s*(?<value>\d[\d,]*(?:\.\d+)?)");
        return match.Success && decimal.TryParse(match.Groups["value"].Value.Replace(",", "", StringComparison.Ordinal), NumberStyles.Number, CultureInfo.InvariantCulture, out var value)
            ? value
            : null;
    }

    private static int? ExtractMonthCount(string text)
    {
        var match = Regex.Match(text, @"(?<value>\d+)\s*month", RegexOptions.IgnoreCase);
        return match.Success && int.TryParse(match.Groups["value"].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value)
            ? value
            : null;
    }

    private static int? ExtractDayCount(string text)
    {
        var match = Regex.Match(text, @"(?<value>\d+)\s*day", RegexOptions.IgnoreCase);
        return match.Success && int.TryParse(match.Groups["value"].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value)
            ? value
            : null;
    }

    private static int? ExtractAgeNumber(string text)
    {
        var match = Regex.Match(text, @"(?:age|to age|at age)\s*(?<value>\d+)", RegexOptions.IgnoreCase);
        return match.Success && int.TryParse(match.Groups["value"].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value)
            ? value
            : null;
    }

    private static IEnumerable<string> SplitDetails(string details)
    {
        return details
            .Split([Environment.NewLine], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(value => !string.IsNullOrWhiteSpace(value));
    }

    private static string FirstNonBlank(params string[] values)
    {
        return values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)) ?? "";
    }

    private sealed class PlanDesignWorkbookRow
    {
        public string ClientId { get; init; } = "";
        public string ClientName { get; init; } = "";
        public string Carrier { get; init; } = "";
        public string PolicyNumber { get; init; } = "";
        public string ClassName { get; init; } = "";
        public string BenefitCategory { get; init; } = "";
        public string BenefitName { get; init; } = "";
        public string ProvisionName { get; init; } = "";
        public string CoverageDescription { get; init; } = "";
        public string Coinsurance { get; init; } = "";
        public string Maximum { get; init; } = "";
        public string Frequency { get; init; } = "";
        public string WaitingPeriod { get; init; } = "";
        public string TerminationAge { get; init; } = "";
        public string Notes { get; init; } = "";
        public string SourceText { get; init; } = "";
        public string SourceFileName { get; init; } = "";
        public int? SourcePage { get; init; }
        public decimal Confidence { get; init; }
        public string Status { get; init; } = "";
    }

    private sealed class ScheduleWorkbookRow
    {
        public int SlidePage { get; init; }
        public int PageTotal { get; init; }
        public int SectionOrder { get; init; }
        public string ClientName { get; init; } = "";
        public string Carrier { get; init; } = "";
        public string ClassName { get; init; } = "";
        public string SectionName { get; init; } = "";
        public string DetailText { get; init; } = "";
        public string SourceFile { get; init; } = "";
        public string Notes { get; init; } = "";
    }
}

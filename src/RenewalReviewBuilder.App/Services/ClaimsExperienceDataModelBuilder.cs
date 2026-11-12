using System.Globalization;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using RenewalReviewBuilder.App.Models;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using S = DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace RenewalReviewBuilder.App.Services;

public sealed partial class ClaimsExperienceDataModelBuilder
{
    private static readonly CultureInfo Culture = CultureInfo.InvariantCulture;
    private static readonly string[] MonthFormats = ["MMM-yyyy", "MMM yyyy", "MMMM, yyyy", "MMMM yyyy"];
    private const string Style1NetPremiumColor = "B7DEE8";
    private const string Style1PaidClaimsColor = "1F497D";
    private const string Style1GridlineColor = "D9E2E8";
    private const string Style1AxisLineColor = "D9D9D9";
    private const string Style1TextColor = "404040";
    private readonly PdfTextExtractor _pdfTextExtractor;

    public ClaimsExperienceDataModelBuilder(PdfTextExtractor? pdfTextExtractor = null)
    {
        _pdfTextExtractor = pdfTextExtractor ?? new PdfTextExtractor();
    }

    public Task<ClaimsExperienceDataModelResult> ExportAsync(ClaimsExperienceDataModelRequest request)
    {
        return Task.Run(() => Export(request));
    }

    public Task<ClaimsExperienceDataModelResult> ExtractAsync(ClaimsExperienceDataModelRequest request)
    {
        return Task.Run(() => Extract(request));
    }

    public ClaimsExperienceDataModelResult Export(ClaimsExperienceDataModelRequest request)
    {
        return BuildResult(request, writeWorkbook: true);
    }

    public ClaimsExperienceDataModelResult Extract(ClaimsExperienceDataModelRequest request)
    {
        return BuildResult(request, writeWorkbook: false);
    }

    private ClaimsExperienceDataModelResult BuildResult(
        ClaimsExperienceDataModelRequest request,
        bool writeWorkbook)
    {
        var requestedReportMonthCount = NormalizeRequestedReportMonthCount(request.ReportMonthCount);
        var sourceFiles = ResolveSourceFiles(request);
        if (sourceFiles.Count == 0)
        {
            throw new InvalidOperationException("Select the claims PDF and EHC/Dental Excel files before exporting the claims data model.");
        }

        var outputFolder = string.IsNullOrWhiteSpace(request.OutputFolder)
            ? Path.Combine(AppContext.BaseDirectory, "Exports")
            : request.OutputFolder;

        var warnings = new List<string>();
        var data = ReadSourceData(request.SourceFolder, sourceFiles, warnings, requestedReportMonthCount);
        data.RenewalTerms = request.RenewalTerms ?? new ClaimsExperienceRenewalTerms();
        var outputPath = "";
        var tableCount = 0;
        if (writeWorkbook)
        {
            Directory.CreateDirectory(outputFolder);
            ReadScheduleOfBenefits(request, sourceFiles, data, warnings);
            outputPath = UniqueDestinationPath(outputFolder, "Claims_Experience_Data_Model.xlsx");
            tableCount = WriteWorkbook(outputPath, data, warnings);
        }

        var result = new ClaimsExperienceDataModelResult
        {
            OutputFolder = outputFolder,
            WorkbookPath = outputPath,
            ClientName = data.ClientName,
            ClientNumber = data.PolicyNumber,
            Carrier = data.Carrier,
            ReportMonthCount = data.ReportMonthCount,
            ExperiencePeriodStart = data.PeriodStart,
            ExperiencePeriodEnd = data.PeriodEnd,
            TableCount = tableCount,
            SourceRowCount = data.Months.Count + data.PriorMonths.Count + data.PriorPriorMonths.Count + data.EhcCategories.Count + data.ParamedicalDetails.Count + data.DentalCategories.Count + data.BillingInsuredLivesDetails.Count + data.RenewalExperienceFinancials.Count + data.ScheduleOfBenefitsRows.Count,
            ScheduleOfBenefitsSourcePath = data.ScheduleOfBenefitsSourcePath
        };

        result.SourceDocumentPaths.AddRange(sourceFiles);
        result.SourceDocumentPaths.AddRange(data.ImportedFiles
            .Select(record => record.FullPath)
            .Where(path => !result.SourceDocumentPaths.Contains(path, StringComparer.OrdinalIgnoreCase)));
        result.Summaries.AddRange(BuildClaimsExperienceSummaries(data, data.ReportMonthCount));
        result.MonthlyClaims.AddRange(BuildMonthlyClaims(data, data.ReportMonthCount));
        result.AllPeriodMonthlyClaims.AddRange(BuildAllPeriodMonthlyClaims(data, data.ReportMonthCount));
        result.SourcePeriods.AddRange(data.SourcePeriods);
        result.ServiceAdjudicationFacts.AddRange(data.ServiceAdjudicationFacts);
        result.ServiceControlTotals.AddRange(data.ServiceControlTotals);
        result.Warnings.AddRange(warnings);
        return result;
    }

    private static List<ClaimsExperienceSummary> BuildClaimsExperienceSummaries(ClaimsExperienceDataModelData data, int reportMonthCount)
    {
        var selectedMonths = SelectExperienceMonths(data.Months, reportMonthCount);
        var useReportedTotals = data.Totals is not null &&
                                (data.DetectedReportMonthCount == 0 || reportMonthCount == data.DetectedReportMonthCount);

        var summaries = new List<ClaimsExperienceSummary>();
        AddSummary(
            CoverageNames.Ehc,
            useReportedTotals ? data.Totals!.EhcPremium : selectedMonths.Sum(month => month.EhcPremium),
            useReportedTotals ? data.Totals!.EhcClaims : selectedMonths.Sum(month => month.EhcClaims),
            AverageLives(data, selectedMonths, CoverageNames.Ehc));
        AddSummary(
            CoverageNames.Dental,
            useReportedTotals ? data.Totals!.DentalPremium : selectedMonths.Sum(month => month.DentalPremium),
            useReportedTotals ? data.Totals!.DentalClaims : selectedMonths.Sum(month => month.DentalClaims),
            AverageLives(data, selectedMonths, CoverageNames.Dental));

        return summaries;

        void AddSummary(string benefitLine, decimal premium, decimal paidClaims, int? lives)
        {
            if (premium == 0m && paidClaims == 0m && lives is null)
            {
                return;
            }

            summaries.Add(new ClaimsExperienceSummary
            {
                BenefitLine = benefitLine,
                Premium = premium == 0m ? null : premium,
                PaidClaims = paidClaims == 0m ? null : paidClaims,
                IncurredClaims = paidClaims == 0m ? null : paidClaims,
                LossRatio = premium > 0m && paidClaims > 0m ? decimal.Round(paidClaims / premium * 100m, 1) : null,
                Lives = lives
            });
        }
    }

    private static int? AverageLives(ClaimsExperienceDataModelData data, IReadOnlyList<ClaimsExperienceMonth> months, string coverage)
    {
        var lives = months
            .Select(month => data.BillingMonthlyLivesByMonth.TryGetValue(FirstDayOfMonth(month.Month), out var billingLives)
                ? coverage.Equals(CoverageNames.Ehc, StringComparison.OrdinalIgnoreCase)
                    ? billingLives.HealthInsuredLives
                    : billingLives.DentalInsuredLives
                : null)
            .Where(value => value.HasValue)
            .Select(value => value!.Value)
            .ToList();

        return lives.Count == 0
            ? null
            : (int)Math.Round(lives.Average(), MidpointRounding.AwayFromZero);
    }

    private static List<MonthlyClaimsRecord> BuildMonthlyClaims(ClaimsExperienceDataModelData data, int reportMonthCount)
    {
        return CreateMonthlyClaimsRecords(data, SelectExperienceMonths(data.Months, reportMonthCount));
    }

    private static List<MonthlyClaimsRecord> BuildAllPeriodMonthlyClaims(
        ClaimsExperienceDataModelData data,
        int reportMonthCount)
    {
        var selectedMonths = SelectExperienceMonths(data.PriorPriorMonths, 12)
            .Concat(SelectExperienceMonths(data.PriorMonths, reportMonthCount))
            .Concat(SelectExperienceMonths(data.Months, reportMonthCount));

        return CreateMonthlyClaimsRecords(data, selectedMonths);
    }

    private static List<MonthlyClaimsRecord> CreateMonthlyClaimsRecords(
        ClaimsExperienceDataModelData data,
        IEnumerable<ClaimsExperienceMonth> selectedMonths)
    {
        var rows = new List<MonthlyClaimsRecord>();

        foreach (var month in selectedMonths)
        {
            rows.Add(CreateMonthlyClaimsRecord(data, month, CoverageNames.Ehc, month.EhcPremium, month.EhcClaims));
            rows.Add(CreateMonthlyClaimsRecord(data, month, CoverageNames.Dental, month.DentalPremium, month.DentalClaims));
        }

        return rows
            .Where(row => row.Premium is > 0m || row.PaidClaims is > 0m || row.Lives is not null)
            .GroupBy(row => (row.Month, BenefitLine: row.BenefitLine.Trim().ToUpperInvariant()))
            .Select(group => group.Last())
            .OrderBy(row => row.Month)
            .ThenBy(row => row.BenefitLine, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static MonthlyClaimsRecord CreateMonthlyClaimsRecord(
        ClaimsExperienceDataModelData data,
        ClaimsExperienceMonth month,
        string coverage,
        decimal premium,
        decimal paidClaims)
    {
        var monthKey = FirstDayOfMonth(month.Month);
        data.BillingMonthlyLivesByMonth.TryGetValue(monthKey, out var billingLives);
        var lives = coverage.Equals(CoverageNames.Ehc, StringComparison.OrdinalIgnoreCase)
            ? billingLives?.HealthInsuredLives
            : billingLives?.DentalInsuredLives;
        var billingPremium = coverage.Equals(CoverageNames.Ehc, StringComparison.OrdinalIgnoreCase)
            ? billingLives?.HealthPremium
            : billingLives?.DentalPremium;
        var finalPremium = premium != 0m ? premium : billingPremium ?? premium;
        var claimsPlusFees = paidClaims;

        return new MonthlyClaimsRecord
        {
            Carrier = data.Carrier,
            BenefitLine = coverage,
            Month = monthKey,
            Lives = lives,
            Premium = finalPremium == 0m ? null : finalPremium,
            PaidClaims = paidClaims == 0m ? null : paidClaims,
            IncurredClaims = paidClaims == 0m ? null : paidClaims,
            ClaimsPlusFees = claimsPlusFees == 0m ? null : claimsPlusFees,
            LossRatio = finalPremium > 0m && claimsPlusFees > 0m
                ? decimal.Round(claimsPlusFees / finalPremium * 100m, 1)
                : null,
            DateLoaded = DateTime.Now
        };
    }

    private ClaimsExperienceDataModelData ReadSourceData(string sourceFolder, IReadOnlyList<string> sourceFiles, List<string> warnings, int reportMonthCount)
    {
        var data = new ClaimsExperienceDataModelData
        {
            Carrier = "",
            RequestedReportMonthCount = reportMonthCount,
            ReportMonthCount = reportMonthCount > 0 ? reportMonthCount : 12,
            CurrentSourcePeriodId = CurrentPeriodId(reportMonthCount > 0 ? reportMonthCount : 12),
            PriorSourcePeriodId = PriorPeriodId(reportMonthCount > 0 ? reportMonthCount : 12),
            BenchmarkSourcePeriodId = BenchmarkPeriodId(reportMonthCount > 0 ? reportMonthCount : 12)
        };

        var experienceReport = FindExperienceReport(sourceFiles);
        var experienceWorkbook = experienceReport is null
            ? FindExperienceWorkbook(sourceFiles)
            : null;

        if (experienceReport is null && experienceWorkbook is null)
        {
            if (!sourceFiles.Any(IsHybridAsoDentalClaimsExperienceFileName))
            {
                warnings.Add("A supported PDF or Excel experience report was not found in the selected files.");
            }
        }
        else if (experienceReport is not null)
        {
            data.ImportedFiles.Add(CreateImportRecord(experienceReport, "PDF experience report"));
            ReadExperienceReport(experienceReport, data, warnings);
            ReadAdditionalExperienceReports(sourceFiles, experienceReport, data, warnings);
        }
        else
        {
            data.ImportedFiles.Add(CreateImportRecord(experienceWorkbook!, "Excel experience report"));
            ReadExperienceWorkbook(experienceWorkbook!, data, warnings);
        }

        ReadHybridAsoDentalExperienceReports(sourceFiles, data, warnings);

        ResolveReportPeriodAndMonthCount(data, reportMonthCount, warnings);

        var ehcMonthlyWorkbook = FindMonthlyWorkbook(sourceFiles, "ehc");
        var ehcContractSummaryPaidTotal = ReadContractSummaryPaidGrandTotal(ehcMonthlyWorkbook, CoverageNames.Ehc, warnings);
        if (ehcMonthlyWorkbook is not null)
        {
            data.ImportedFiles.Add(CreateImportRecord(ehcMonthlyWorkbook, "EHC monthly contract summary validation workbook"));
            ReadMonthlyClaimsWorkbook(ehcMonthlyWorkbook, CoverageNames.Ehc, data, warnings);
        }

        var ehcWorkbook = FindClaimsWorkbook(sourceFiles, CoverageNames.Ehc, data.PeriodStart, data.PeriodEnd);
        var ehcServicePdf = FindClaimsServicePdf(sourceFiles, "ehc", experienceReport, data);
        if (ehcWorkbook is null && ehcServicePdf is null)
        {
            warnings.Add("EHC claims detail file was not found (PDF or Excel).");
        }
        else if (ehcWorkbook is not null)
        {
            data.ImportedFiles.Add(CreateImportRecord(ehcWorkbook, "EHC claims workbook"));
            ReadClaimsSummaryWorkbook(ehcWorkbook, CoverageNames.Ehc, data, warnings, data.CurrentSourcePeriodId, ehcContractSummaryPaidTotal, ehcMonthlyWorkbook);
            if (ehcServicePdf is not null)
            {
                ValidateClaimsServicePdfIdentity(ehcServicePdf, CoverageNames.Ehc, data, warnings);
            }
        }
        else
        {
            ReadClaimsServicePdf(ehcServicePdf!, CoverageNames.Ehc, data, warnings);
        }

        var dentalMonthlyWorkbook = FindMonthlyWorkbook(sourceFiles, "dental");
        var dentalContractSummaryPaidTotal = ReadContractSummaryPaidGrandTotal(dentalMonthlyWorkbook, CoverageNames.Dental, warnings);
        if (dentalMonthlyWorkbook is not null)
        {
            data.ImportedFiles.Add(CreateImportRecord(dentalMonthlyWorkbook, "Dental monthly contract summary validation workbook"));
            ReadMonthlyClaimsWorkbook(dentalMonthlyWorkbook, CoverageNames.Dental, data, warnings);
        }

        var dentalWorkbook = FindClaimsWorkbook(sourceFiles, CoverageNames.Dental, data.PeriodStart, data.PeriodEnd);
        var dentalServicePdf = FindClaimsServicePdf(sourceFiles, "dental", experienceReport, data);
        if (dentalWorkbook is null && dentalServicePdf is null)
        {
            warnings.Add("Dental claims detail file was not found (PDF or Excel).");
        }
        else if (dentalWorkbook is not null)
        {
            data.ImportedFiles.Add(CreateImportRecord(dentalWorkbook, "Dental claims workbook"));
            ReadClaimsSummaryWorkbook(dentalWorkbook, CoverageNames.Dental, data, warnings, data.CurrentSourcePeriodId, dentalContractSummaryPaidTotal, dentalMonthlyWorkbook);
            if (dentalServicePdf is not null)
            {
                ValidateClaimsServicePdfIdentity(dentalServicePdf, CoverageNames.Dental, data, warnings);
            }
        }
        else
        {
            ReadClaimsServicePdf(dentalServicePdf!, CoverageNames.Dental, data, warnings);
        }

        ReadSunLifeRenewalExperienceComparisons(sourceFolder, sourceFiles, data, warnings);
        ReadPriorSourceData(sourceFolder, data, warnings);
        ReadPriorPriorSourceData(sourceFolder, data, warnings);
        ReadBillingInsuredLivesForSourceFolders(sourceFolder, data, warnings);
        WarnForMissingBillingMonths(data, warnings);

        if (data.PeriodStart is null && data.Months.Count > 0)
        {
            data.PeriodStart = data.Months.Min(month => month.Month);
        }

        if (data.PeriodEnd is null && data.Months.Count > 0)
        {
            data.PeriodEnd = EndOfMonth(data.Months.Max(month => month.Month));
        }

        if (string.IsNullOrWhiteSpace(data.PeriodLabel) && data.PeriodStart is not null && data.PeriodEnd is not null)
        {
            data.PeriodLabel = $"{data.PeriodStart.Value.ToString("MMM yyyy", Culture)} - {data.PeriodEnd.Value.ToString("MMM yyyy", Culture)}";
        }

        return data;
    }

    private void ReadScheduleOfBenefits(
        ClaimsExperienceDataModelRequest request,
        IReadOnlyList<string> sourceFiles,
        ClaimsExperienceDataModelData data,
        List<string> warnings)
    {
        var sourcePath = ResolveScheduleOfBenefitsWorkbook(request, sourceFiles);
        if (string.IsNullOrWhiteSpace(sourcePath))
        {
            AddSchedulePlaceholder(data);
            return;
        }

        try
        {
            using var workbook = OpenWorkbookReadOnly(sourcePath);
            var rows = SelectScheduleRowsForClientAndClass(ReadSchedulePptExportRows(workbook), data);
            if (!HasUsableScheduleRows(rows))
            {
                var planDesignRows = SelectScheduleRowsForClientAndClass(ReadPlanDesignScheduleRows(workbook), data);
                if (HasUsableScheduleRows(planDesignRows))
                {
                    rows = planDesignRows;
                }
            }

            if (rows.Count == 0)
            {
                warnings.Add("Benefits master workbook was found, but no schedule of benefits rows could be read.");
                AddSchedulePlaceholder(data);
                return;
            }

            data.ScheduleOfBenefitsSourcePath = sourcePath;
            data.ScheduleOfBenefitsClientName = FirstNonBlank(rows.FirstOrDefault()?.ClientName, data.ClientName);
            data.ImportedFiles.Add(CreateImportRecord(sourcePath, "Benefits schedule master workbook"));
            data.ScheduleOfBenefitsRows.AddRange(ReorderScheduleRows(rows));
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            warnings.Add("Schedule of benefits could not be read from the benefits master workbook.");
            AddSchedulePlaceholder(data);
        }
    }

    private static string? ResolveScheduleOfBenefitsWorkbook(ClaimsExperienceDataModelRequest request, IReadOnlyList<string> sourceFiles)
    {
        if (!string.IsNullOrWhiteSpace(request.BookletMasterWorkbookPath) && File.Exists(request.BookletMasterWorkbookPath))
        {
            return request.BookletMasterWorkbookPath;
        }

        var sourceFile = sourceFiles.FirstOrDefault(path =>
            Path.GetFileName(path).Equals("Benefits_Master_Database.xlsx", StringComparison.OrdinalIgnoreCase));
        if (!string.IsNullOrWhiteSpace(sourceFile) && File.Exists(sourceFile))
        {
            return sourceFile;
        }

        if (!string.IsNullOrWhiteSpace(request.SourceFolder) && Directory.Exists(request.SourceFolder))
        {
            var candidates = new[]
            {
                Path.Combine(request.SourceFolder, "Benefits_Master_Database.xlsx"),
                Path.Combine(request.SourceFolder, "Current", "Benefits_Master_Database.xlsx")
            };
            var sourceFolderMatch = candidates.FirstOrDefault(File.Exists);
            if (!string.IsNullOrWhiteSpace(sourceFolderMatch))
            {
                return sourceFolderMatch;
            }
        }

        if (request.BenefitBookletPaths.Count > 0 && !string.IsNullOrWhiteSpace(request.OutputFolder))
        {
            var outputCandidate = Path.Combine(request.OutputFolder, "Benefits_Master_Database.xlsx");
            if (File.Exists(outputCandidate))
            {
                return outputCandidate;
            }
        }

        return null;
    }

    private static List<ScheduleOfBenefitsRow> ReadSchedulePptExportRows(XLWorkbook workbook)
    {
        var sheet = workbook.Worksheets.FirstOrDefault(worksheet =>
            worksheet.Name.Equals("Schedule_PPT_Export", StringComparison.OrdinalIgnoreCase));
        if (sheet is null)
        {
            return [];
        }

        var headerRow = sheet.RowsUsed().FirstOrDefault(row =>
            row.CellsUsed().Any(cell => NormalizeHeader(cell.GetString()).Equals("PowerPointTitle", StringComparison.OrdinalIgnoreCase)) &&
            row.CellsUsed().Any(cell => NormalizeHeader(cell.GetString()).Equals("PageNumber", StringComparison.OrdinalIgnoreCase)));
        if (headerRow is null)
        {
            return [];
        }

        var headers = BuildHeaderMap(headerRow);
        if (!TryGetColumn(headers, out var clientColumn, "ClientName", "Client") ||
            !TryGetColumn(headers, out var sectionColumn, "SectionName", "Benefit", "Category") ||
            !TryGetColumn(headers, out var detailColumn, "DetailText", "Details", "Description"))
        {
            return [];
        }

        TryGetColumn(headers, out var titleColumn, "PowerPointTitle", "Title");
        TryGetColumn(headers, out var sourceFileColumn, "SourceFile", "Source");
        TryGetColumn(headers, out var carrierColumn, "Carrier", "Insurer");
        TryGetColumn(headers, out var classColumn, "ClassName", "Class");
        TryGetColumn(headers, out var pageColumn, "PageNumber", "SlidePage", "Page");
        TryGetColumn(headers, out var pageTotalColumn, "PageTotal", "TotalPages");
        TryGetColumn(headers, out var orderColumn, "SectionOrder", "SortOrder", "Order");

        var rows = new List<ScheduleOfBenefitsRow>();
        foreach (var row in sheet.RowsUsed().Where(row => row.RowNumber() > headerRow.RowNumber()))
        {
            var detailText = CleanScheduleDetail(CellMultilineText(row.Cell(detailColumn)));
            var sectionName = FirstNonBlank(CellText(row.Cell(sectionColumn)), "Schedule of Benefits");
            if (string.IsNullOrWhiteSpace(detailText) && string.IsNullOrWhiteSpace(sectionName))
            {
                continue;
            }

            var slidePage = pageColumn == 0 ? SchedulePageForSection(sectionName) : ReadInt(row.Cell(pageColumn), SchedulePageForSection(sectionName));
            var pageTotal = pageTotalColumn == 0 ? 3 : ReadInt(row.Cell(pageTotalColumn), 3);
            var sectionOrder = orderColumn == 0 ? rows.Count + 1 : ReadInt(row.Cell(orderColumn), rows.Count + 1);
            rows.Add(new ScheduleOfBenefitsRow
            {
                SlidePage = NormalizeScheduleSlidePage(slidePage),
                PageTotal = pageTotal <= 0 ? 3 : pageTotal,
                SectionOrder = sectionOrder,
                ClientName = CellText(row.Cell(clientColumn)),
                Carrier = carrierColumn == 0 ? "" : CellText(row.Cell(carrierColumn)),
                ClassName = classColumn == 0 ? "Class 1" : FirstNonBlank(CellText(row.Cell(classColumn)), "Class 1"),
                SectionName = sectionName,
                DetailText = detailText,
                SourceFile = sourceFileColumn == 0 ? "" : CellText(row.Cell(sourceFileColumn)),
                Notes = titleColumn == 0 ? "" : CellText(row.Cell(titleColumn))
            });
        }

        return rows;
    }

    private static List<ScheduleOfBenefitsRow> ReadPlanDesignScheduleRows(XLWorkbook workbook)
    {
        var sheet = workbook.Worksheets.FirstOrDefault(worksheet =>
            worksheet.Name.Equals("PlanDesign", StringComparison.OrdinalIgnoreCase) ||
            worksheet.Name.Equals("Plan Design", StringComparison.OrdinalIgnoreCase));
        if (sheet is null)
        {
            return [];
        }

        var headerRow = sheet.RowsUsed().FirstOrDefault(row =>
            row.CellsUsed().Any(cell => NormalizeHeader(cell.GetString()).Equals("ClientName", StringComparison.OrdinalIgnoreCase)) &&
            row.CellsUsed().Any(cell => NormalizeHeader(cell.GetString()).Equals("Benefit", StringComparison.OrdinalIgnoreCase)));
        if (headerRow is null)
        {
            return [];
        }

        var headers = BuildHeaderMap(headerRow);
        if (!TryGetColumn(headers, out var clientColumn, "ClientName", "Client") ||
            !TryGetColumn(headers, out var benefitColumn, "Benefit", "BenefitName"))
        {
            return [];
        }

        TryGetColumn(headers, out var carrierColumn, "Carrier", "Insurer");
        TryGetColumn(headers, out var classColumn, "ClassName", "Class");
        TryGetColumn(headers, out var descriptionColumn, "CoverageDescription", "Description", "CoverageDetails");
        TryGetColumn(headers, out var coinsuranceColumn, "Coinsurance", "Reimbursement", "ReimbursementPercent");
        TryGetColumn(headers, out var maximumColumn, "Maximum", "MaximumAmount", "Max");
        TryGetColumn(headers, out var frequencyColumn, "Frequency");
        TryGetColumn(headers, out var waitingColumn, "WaitingPeriod", "Eligibility", "Waiting");
        TryGetColumn(headers, out var terminationColumn, "TerminationAge", "Termination");
        TryGetColumn(headers, out var sourceFileColumn, "SourceFile", "SourceDocument", "SourceFileName");
        TryGetColumn(headers, out var notesColumn, "Notes");

        var rows = new List<ScheduleOfBenefitsRow>();
        foreach (var row in sheet.RowsUsed().Where(row => row.RowNumber() > headerRow.RowNumber()))
        {
            var benefit = CellText(row.Cell(benefitColumn));
            if (string.IsNullOrWhiteSpace(benefit))
            {
                continue;
            }

            var sectionName = ScheduleSectionForBenefit(benefit);
            var detailParts = new List<string>();
            AddScheduleDetailPart(detailParts, benefit, descriptionColumn == 0 ? "" : CellMultilineText(row.Cell(descriptionColumn)));
            AddScheduleDetailPart(detailParts, "Coinsurance", coinsuranceColumn == 0 ? "" : CellText(row.Cell(coinsuranceColumn)));
            AddScheduleDetailPart(detailParts, "Maximum", maximumColumn == 0 ? "" : CellText(row.Cell(maximumColumn)));
            AddScheduleDetailPart(detailParts, "Frequency", frequencyColumn == 0 ? "" : CellText(row.Cell(frequencyColumn)));
            AddScheduleDetailPart(detailParts, "Waiting Period", waitingColumn == 0 ? "" : CellText(row.Cell(waitingColumn)));
            AddScheduleDetailPart(detailParts, "Termination Age", terminationColumn == 0 ? "" : CellText(row.Cell(terminationColumn)));

            rows.Add(new ScheduleOfBenefitsRow
            {
                SlidePage = SchedulePageForSection(sectionName),
                PageTotal = 3,
                SectionOrder = rows.Count + 1,
                ClientName = CellText(row.Cell(clientColumn)),
                Carrier = carrierColumn == 0 ? "" : CellText(row.Cell(carrierColumn)),
                ClassName = classColumn == 0 ? "Class 1" : FirstNonBlank(CellText(row.Cell(classColumn)), "Class 1"),
                SectionName = sectionName,
                DetailText = CleanScheduleDetail(string.Join("\n", detailParts)),
                SourceFile = sourceFileColumn == 0 ? "" : CellText(row.Cell(sourceFileColumn)),
                Notes = notesColumn == 0 ? "" : CellText(row.Cell(notesColumn))
            });
        }

        return rows;
    }

    private static bool HasUsableScheduleRows(IReadOnlyList<ScheduleOfBenefitsRow> rows)
    {
        if (rows.Count == 0)
        {
            return false;
        }

        var realPages = rows
            .Where(row => !IsPlaceholderScheduleDetail(row.DetailText))
            .Select(row => NormalizeScheduleSlidePage(row.SlidePage))
            .ToHashSet();
        return realPages.Contains(1) && realPages.Contains(2) && realPages.Contains(3);
    }

    private static bool IsPlaceholderScheduleDetail(string detailText)
    {
        return detailText.Contains("Attach a benefits booklet PDF", StringComparison.OrdinalIgnoreCase) ||
               detailText.Contains("supply Benefits_Master_Database.xlsx", StringComparison.OrdinalIgnoreCase) ||
               detailText.Contains("populate this schedule", StringComparison.OrdinalIgnoreCase);
    }

    private static List<ScheduleOfBenefitsRow> SelectScheduleRowsForClientAndClass(List<ScheduleOfBenefitsRow> rows, ClaimsExperienceDataModelData data)
    {
        if (rows.Count == 0)
        {
            return [];
        }

        var selectedRows = rows;
        var clientKey = NormalizeHeader(data.ClientName);
        if (!string.IsNullOrWhiteSpace(clientKey))
        {
            var matchingClientRows = rows
                .Where(row =>
                {
                    var rowClient = NormalizeHeader(row.ClientName);
                    return rowClient.Equals(clientKey, StringComparison.OrdinalIgnoreCase) ||
                           rowClient.Contains(clientKey, StringComparison.OrdinalIgnoreCase) ||
                           clientKey.Contains(rowClient, StringComparison.OrdinalIgnoreCase);
                })
                .ToList();
            if (matchingClientRows.Count > 0)
            {
                selectedRows = matchingClientRows;
            }
            else
            {
                var distinctClients = rows
                    .Select(row => row.ClientName)
                    .Where(value => !string.IsNullOrWhiteSpace(value))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();
                if (distinctClients.Count > 1)
                {
                    return [];
                }
            }
        }

        var selectedClass = selectedRows
            .Select(row => row.ClassName)
            .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
        if (!string.IsNullOrWhiteSpace(selectedClass))
        {
            selectedRows = selectedRows
                .Where(row => row.ClassName.Equals(selectedClass, StringComparison.OrdinalIgnoreCase))
                .ToList();
        }

        return selectedRows;
    }

    private static List<ScheduleOfBenefitsRow> ReorderScheduleRows(List<ScheduleOfBenefitsRow> rows)
    {
        var reordered = new List<ScheduleOfBenefitsRow>();
        foreach (var pageGroup in rows
                     .OrderBy(row => row.SlidePage)
                     .ThenBy(row => row.SectionOrder)
                     .GroupBy(row => row.SlidePage))
        {
            var order = 1;
            foreach (var row in pageGroup)
            {
                row.PageTotal = Math.Max(row.PageTotal, 3);
                row.SectionOrder = order++;
                reordered.Add(row);
            }
        }

        return reordered;
    }

    private static void AddSchedulePlaceholder(ClaimsExperienceDataModelData data)
    {
        if (data.ScheduleOfBenefitsRows.Count > 0)
        {
            return;
        }

        data.ScheduleOfBenefitsClientName = FirstNonBlank(data.ClientName, "Client Name");
        for (var page = 1; page <= 3; page++)
        {
            data.ScheduleOfBenefitsRows.Add(new ScheduleOfBenefitsRow
            {
                SlidePage = page,
                PageTotal = 3,
                SectionOrder = 1,
                ClientName = data.ScheduleOfBenefitsClientName,
                Carrier = data.Carrier,
                ClassName = "Class 1",
                SectionName = page == 1 ? "Schedule of Benefits" : $"Schedule of Benefits Continued {page}",
                DetailText = "Attach a benefits booklet PDF or supply Benefits_Master_Database.xlsx to populate this schedule.",
                SourceFile = "",
                Notes = "Placeholder row."
            });
        }
    }

    private static Dictionary<string, int> BuildHeaderMap(IXLRow row)
    {
        return row.CellsUsed()
            .GroupBy(cell => NormalizeHeader(cell.GetString()), StringComparer.OrdinalIgnoreCase)
            .Where(group => !string.IsNullOrWhiteSpace(group.Key))
            .ToDictionary(group => group.Key, group => group.First().Address.ColumnNumber, StringComparer.OrdinalIgnoreCase);
    }

    private static bool TryGetColumn(IReadOnlyDictionary<string, int> headers, out int column, params string[] names)
    {
        foreach (var name in names)
        {
            if (headers.TryGetValue(NormalizeHeader(name), out column))
            {
                return true;
            }
        }

        column = 0;
        return false;
    }

    private static int ReadInt(IXLCell cell, int fallback)
    {
        var number = ParsingHelpers.ParseDecimal(CellText(cell));
        return number.HasValue ? (int)Math.Round(number.Value, MidpointRounding.AwayFromZero) : fallback;
    }

    private static string CellMultilineText(IXLCell cell)
    {
        var text = cell.GetString();
        if (string.IsNullOrWhiteSpace(text))
        {
            text = cell.GetFormattedString();
        }

        return NormalizeMultiline(text);
    }

    private static string NormalizeMultiline(string value)
    {
        var lines = value
            .Replace("\r\n", "\n", StringComparison.Ordinal)
            .Replace('\r', '\n')
            .Split('\n')
            .Select(line => Regex.Replace(line.Trim(), @"\s+", " "))
            .Where(line => line.Length > 0)
            .ToList();

        return string.Join("\n", lines);
    }

    private static string CleanScheduleDetail(string value)
    {
        var cleaned = NormalizeMultiline(value)
            .Replace("\uf0a1", "-", StringComparison.Ordinal)
            .Trim();
        const int maxLength = 1100;
        return cleaned.Length <= maxLength ? cleaned : cleaned[..maxLength].TrimEnd() + "...";
    }

    private static void AddScheduleDetailPart(List<string> parts, string label, string value)
    {
        var cleaned = CleanScheduleDetail(value);
        if (string.IsNullOrWhiteSpace(cleaned))
        {
            return;
        }

        parts.Add(cleaned.StartsWith(label, StringComparison.OrdinalIgnoreCase)
            ? cleaned
            : $"{label}: {cleaned}");
    }

    private static int NormalizeScheduleSlidePage(int page)
    {
        return page switch
        {
            <= 1 => 1,
            2 => 2,
            _ => 3
        };
    }

    private static string ScheduleSectionForBenefit(string benefit)
    {
        var normalized = NormalizeHeader(benefit);
        if (normalized.Contains("Life", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("ADD", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("ADandD", StringComparison.OrdinalIgnoreCase))
        {
            return "Life / AD&D";
        }

        if (normalized.Contains("Drug", StringComparison.OrdinalIgnoreCase))
        {
            return "Prescription Drugs";
        }

        if (normalized.Contains("Paramedical", StringComparison.OrdinalIgnoreCase))
        {
            return "Paramedicals";
        }

        if (normalized.Contains("Vision", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("Health", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("Travel", StringComparison.OrdinalIgnoreCase))
        {
            return "Medical Services & Supplies";
        }

        if (normalized.Contains("Hospital", StringComparison.OrdinalIgnoreCase))
        {
            return "Hospital";
        }

        if (normalized.Contains("Dental", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("Ortho", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("Perio", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("Endo", StringComparison.OrdinalIgnoreCase))
        {
            return "Dental Care";
        }

        if (normalized.Contains("LTD", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("LongTermDisability", StringComparison.OrdinalIgnoreCase))
        {
            return "Long-Term Disability";
        }

        if (normalized.Contains("STD", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("ShortTermDisability", StringComparison.OrdinalIgnoreCase))
        {
            return "Short-Term Disability";
        }

        if (normalized.Contains("EAP", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("Assistance", StringComparison.OrdinalIgnoreCase))
        {
            return "Employee Assistance Program";
        }

        return FirstNonBlank(benefit, "Schedule of Benefits");
    }

    private static int SchedulePageForSection(string sectionName)
    {
        var normalized = NormalizeHeader(sectionName);
        if (normalized.Contains("Paramedical", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("Hospital", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("Termination", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("Assistance", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("Virtual", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }

        if (normalized.Contains("Disability", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("Dental", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("Travel", StringComparison.OrdinalIgnoreCase))
        {
            return 3;
        }

        return 1;
    }

    private void ReadPriorSourceData(string sourceFolder, ClaimsExperienceDataModelData data, List<string> warnings)
    {
        if (string.IsNullOrWhiteSpace(sourceFolder))
        {
            return;
        }

        var priorFolder = Path.Combine(sourceFolder, "Prior");
        if (!Directory.Exists(priorFolder))
        {
            if (!data.PriorMonths.Any(month => !month.IsAnnualSummary))
            {
                warnings.Add("No Prior folder found; prior period rows were left blank.");
            }

            return;
        }

        var priorFiles = ResolveSourceFilesFromFolder(priorFolder);
        if (priorFiles.Count == 0)
        {
            warnings.Add("Prior folder was found, but no prior claims PDF or Excel files were found.");
            return;
        }

        var priorData = new ClaimsExperienceDataModelData
        {
            Carrier = data.Carrier,
            ReportMonthCount = data.ReportMonthCount,
            CurrentSourcePeriodId = data.CurrentSourcePeriodId,
            PriorSourcePeriodId = data.PriorSourcePeriodId,
            BenchmarkSourcePeriodId = data.BenchmarkSourcePeriodId
        };

        var experienceReport = FindExperienceReport(priorFiles);

        if (experienceReport is null)
        {
            warnings.Add("Prior experience report PDF was not found.");
        }
        else
        {
            data.ImportedFiles.Add(CreateImportRecord(experienceReport, "Prior PDF experience report"));
            ReadExperienceReport(experienceReport, priorData, warnings);
        }

        var ehcMonthlyWorkbook = FindFile(priorFiles, [".xlsx", ".xlsm", ".xls"], "ehc monthly");
        var ehcContractSummaryPaidTotal = ReadContractSummaryPaidGrandTotal(ehcMonthlyWorkbook, CoverageNames.Ehc, warnings);
        if (ehcMonthlyWorkbook is not null)
        {
            data.ImportedFiles.Add(CreateImportRecord(ehcMonthlyWorkbook, "Prior EHC monthly claims workbook"));
            ReadMonthlyClaimsWorkbook(ehcMonthlyWorkbook, CoverageNames.Ehc, priorData, warnings);
        }

        var dentalMonthlyWorkbook = FindFile(priorFiles, [".xlsx", ".xlsm", ".xls"], "dental monthly");
        var dentalContractSummaryPaidTotal = ReadContractSummaryPaidGrandTotal(dentalMonthlyWorkbook, CoverageNames.Dental, warnings);
        if (dentalMonthlyWorkbook is not null)
        {
            data.ImportedFiles.Add(CreateImportRecord(dentalMonthlyWorkbook, "Prior Dental monthly claims workbook"));
            ReadMonthlyClaimsWorkbook(dentalMonthlyWorkbook, CoverageNames.Dental, priorData, warnings);
        }

        var ehcWorkbook = FindClaimsWorkbook(priorFiles, "ehc");
        var ehcServicePdf = FindClaimsServicePdf(priorFiles, "ehc", experienceReport, priorData);
        if (ehcWorkbook is null && ehcServicePdf is null)
        {
            warnings.Add("Prior EHC claims detail file was not found (PDF or Excel).");
        }
        else if (ehcWorkbook is not null)
        {
            data.ImportedFiles.Add(CreateImportRecord(ehcWorkbook, "Prior EHC claims workbook"));
            ReadClaimsSummaryWorkbook(ehcWorkbook, CoverageNames.Ehc, priorData, warnings, data.PriorSourcePeriodId, ehcContractSummaryPaidTotal, ehcMonthlyWorkbook);
        }
        else
        {
            ReadClaimsServicePdf(ehcServicePdf!, CoverageNames.Ehc, priorData, warnings);
        }

        var dentalWorkbook = FindClaimsWorkbook(priorFiles, "dental");
        var dentalServicePdf = FindClaimsServicePdf(priorFiles, "dental", experienceReport, priorData);
        if (dentalWorkbook is null && dentalServicePdf is null)
        {
            warnings.Add("Prior Dental claims detail file was not found (PDF or Excel).");
        }
        else if (dentalWorkbook is not null)
        {
            data.ImportedFiles.Add(CreateImportRecord(dentalWorkbook, "Prior Dental claims workbook"));
            ReadClaimsSummaryWorkbook(dentalWorkbook, CoverageNames.Dental, priorData, warnings, data.PriorSourcePeriodId, dentalContractSummaryPaidTotal, dentalMonthlyWorkbook);
        }
        else
        {
            ReadClaimsServicePdf(dentalServicePdf!, CoverageNames.Dental, priorData, warnings);
        }

        if (ehcWorkbook is null && dentalWorkbook is null && ehcServicePdf is null && dentalServicePdf is null)
        {
            warnings.Add("Prior folder was found, but no prior period claims experience file was found.");
        }

        if (priorData.Months.Any(month => !month.IsAnnualSummary))
        {
            priorData.Months.RemoveAll(month => month.IsAnnualSummary);
            priorData.Months.Sort((left, right) => left.Month.CompareTo(right.Month));
        }

        if (priorData.Months.Any(month => !month.IsAnnualSummary))
        {
            data.PriorMonths.Clear();
            data.PriorMonths.AddRange(priorData.Months);
        }

        data.EhcCategories.AddRange(priorData.EhcCategories);
        data.ParamedicalDetails.AddRange(priorData.ParamedicalDetails);
        data.DentalCategories.AddRange(priorData.DentalCategories);
        data.ServiceAdjudicationFacts.AddRange(priorData.ServiceAdjudicationFacts);
        data.ServiceControlTotals.AddRange(priorData.ServiceControlTotals);
        data.SourcePeriods.AddRange(priorData.SourcePeriods.Where(period => !data.SourcePeriods.Any(existing => existing.Start == period.Start && existing.End == period.End)));
        data.ImportedFiles.AddRange(priorData.ImportedFiles.Where(record => !data.ImportedFiles.Any(existing => existing.FullPath.Equals(record.FullPath, StringComparison.OrdinalIgnoreCase))));
        if (!priorData.Months.Any(month => !month.IsAnnualSummary) &&
            !data.PriorMonths.Any(month => !month.IsAnnualSummary))
        {
            warnings.Add("Prior monthly EHC/Dental experience rows could not be read.");
        }
    }

    private void ReadPriorPriorSourceData(string sourceFolder, ClaimsExperienceDataModelData data, List<string> warnings)
    {
        if (string.IsNullOrWhiteSpace(sourceFolder))
        {
            return;
        }

        var priorPriorFolder = FindPriorPriorFolder(sourceFolder);
        if (priorPriorFolder is null)
        {
            if (!data.PriorPriorMonths.Any(month => !month.IsAnnualSummary))
            {
                warnings.Add("No PriorPrior folder found; PRIOR_PRIOR_12 comparison rows will be blank until supplied.");
            }

            return;
        }

        var priorPriorFiles = ResolveSourceFilesFromFolder(priorPriorFolder);
        if (priorPriorFiles.Count == 0)
        {
            warnings.Add("PriorPrior folder was found, but no prior-prior claims PDF or Excel files were found.");
            return;
        }

        var priorPriorData = new ClaimsExperienceDataModelData
        {
            Carrier = data.Carrier,
            ReportMonthCount = data.ReportMonthCount,
            CurrentSourcePeriodId = data.CurrentSourcePeriodId,
            PriorSourcePeriodId = data.PriorSourcePeriodId,
            BenchmarkSourcePeriodId = data.BenchmarkSourcePeriodId
        };

        var experienceReport = FindExperienceReport(priorPriorFiles);

        if (experienceReport is null)
        {
            warnings.Add("Prior-prior experience report PDF was not found.");
        }
        else
        {
            data.ImportedFiles.Add(CreateImportRecord(experienceReport, "Prior-prior PDF experience report"));
            ReadExperienceReport(experienceReport, priorPriorData, warnings);
        }

        var ehcMonthlyWorkbook = FindFile(priorPriorFiles, [".xlsx", ".xlsm", ".xls"], "ehc monthly");
        var ehcContractSummaryPaidTotal = ReadContractSummaryPaidGrandTotal(ehcMonthlyWorkbook, CoverageNames.Ehc, warnings);
        if (ehcMonthlyWorkbook is not null)
        {
            data.ImportedFiles.Add(CreateImportRecord(ehcMonthlyWorkbook, "Prior-prior EHC monthly claims workbook"));
            ReadMonthlyClaimsWorkbook(ehcMonthlyWorkbook, CoverageNames.Ehc, priorPriorData, warnings);
        }

        var dentalMonthlyWorkbook = FindFile(priorPriorFiles, [".xlsx", ".xlsm", ".xls"], "dental monthly");
        var dentalContractSummaryPaidTotal = ReadContractSummaryPaidGrandTotal(dentalMonthlyWorkbook, CoverageNames.Dental, warnings);
        if (dentalMonthlyWorkbook is not null)
        {
            data.ImportedFiles.Add(CreateImportRecord(dentalMonthlyWorkbook, "Prior-prior Dental monthly claims workbook"));
            ReadMonthlyClaimsWorkbook(dentalMonthlyWorkbook, CoverageNames.Dental, priorPriorData, warnings);
        }

        var ehcWorkbook = FindClaimsWorkbook(priorPriorFiles, "ehc");
        var ehcServicePdf = FindClaimsServicePdf(priorPriorFiles, "ehc", experienceReport, priorPriorData);
        if (ehcWorkbook is not null)
        {
            data.ImportedFiles.Add(CreateImportRecord(ehcWorkbook, "Prior-prior EHC claims workbook"));
            ReadClaimsSummaryWorkbook(ehcWorkbook, CoverageNames.Ehc, priorPriorData, warnings, data.BenchmarkSourcePeriodId, ehcContractSummaryPaidTotal, ehcMonthlyWorkbook);
        }
        else if (ehcServicePdf is not null)
        {
            ReadClaimsServicePdf(ehcServicePdf, CoverageNames.Ehc, priorPriorData, warnings);
        }

        var dentalWorkbook = FindClaimsWorkbook(priorPriorFiles, "dental");
        var dentalServicePdf = FindClaimsServicePdf(priorPriorFiles, "dental", experienceReport, priorPriorData);
        if (dentalWorkbook is not null)
        {
            data.ImportedFiles.Add(CreateImportRecord(dentalWorkbook, "Prior-prior Dental claims workbook"));
            ReadClaimsSummaryWorkbook(dentalWorkbook, CoverageNames.Dental, priorPriorData, warnings, data.BenchmarkSourcePeriodId, dentalContractSummaryPaidTotal, dentalMonthlyWorkbook);
        }
        else if (dentalServicePdf is not null)
        {
            ReadClaimsServicePdf(dentalServicePdf, CoverageNames.Dental, priorPriorData, warnings);
        }

        if (priorPriorData.Months.Any(month => !month.IsAnnualSummary))
        {
            priorPriorData.Months.RemoveAll(month => month.IsAnnualSummary);
            priorPriorData.Months.Sort((left, right) => left.Month.CompareTo(right.Month));
        }

        if (priorPriorData.Months.Any(month => !month.IsAnnualSummary))
        {
            data.PriorPriorMonths.Clear();
            data.PriorPriorMonths.AddRange(priorPriorData.Months);
        }

        data.EhcCategories.AddRange(priorPriorData.EhcCategories);
        data.ParamedicalDetails.AddRange(priorPriorData.ParamedicalDetails);
        data.DentalCategories.AddRange(priorPriorData.DentalCategories);
        data.ServiceAdjudicationFacts.AddRange(priorPriorData.ServiceAdjudicationFacts);
        data.ServiceControlTotals.AddRange(priorPriorData.ServiceControlTotals);
        data.SourcePeriods.AddRange(priorPriorData.SourcePeriods.Where(period => !data.SourcePeriods.Any(existing => existing.Start == period.Start && existing.End == period.End)));
        data.ImportedFiles.AddRange(priorPriorData.ImportedFiles.Where(record => !data.ImportedFiles.Any(existing => existing.FullPath.Equals(record.FullPath, StringComparison.OrdinalIgnoreCase))));
        if (!priorPriorData.Months.Any(month => !month.IsAnnualSummary) &&
            !data.PriorPriorMonths.Any(month => !month.IsAnnualSummary))
        {
            warnings.Add("Prior-prior monthly EHC/Dental experience rows could not be read.");
        }
    }

    private void ReadExperienceReport(string pdfPath, ClaimsExperienceDataModelData data, List<string> warnings)
    {
        try
        {
            var pages = _pdfTextExtractor.Extract(pdfPath);
            var allText = ParsingHelpers.Clean(string.Join(" ", pages.Select(page => page.Text)));
            if (GroupSourcePdfParser.TryParseMonthlyClaimsReport(pdfPath, pages, out var groupSourceReport))
            {
                MergeGroupSourceMonthlyReport(data, groupSourceReport, warnings);
                return;
            }

            var identityLines = pages.SelectMany(page => page.Lines).ToList();
            var identity = ClaimsPdfIdentityExtractor.Extract(identityLines);
            data.Carrier = FirstNonBlank(data.Carrier, identity.Carrier);
            data.ClientName = FirstNonBlank(data.ClientName, identity.ClientName);
            data.PolicyNumber = FirstNonBlank(data.PolicyNumber, identity.PolicyNumber);

            if (TryExtractExperienceReportPeriod(allText, out var periodStart, out var periodEnd))
            {
                data.PeriodStart = periodStart;
                data.PeriodEnd = periodEnd;
                data.PeriodLabel = FormatPeriodLabel(periodStart, periodEnd);
            }

            var sourceLines = pages.SelectMany(page => page.Lines).ToList();
            if (sourceLines.Any(line => line.Contains("Lives/Premium/Claims", StringComparison.OrdinalIgnoreCase) &&
                                       line.Contains("Member Based", StringComparison.OrdinalIgnoreCase)))
            {
                ReadSunLifeMemberBasedMonthlyExperience(sourceLines, data, pdfPath);
                data.Months.Sort((left, right) => left.Month.CompareTo(right.Month));
                if (data.Months.Count > 0)
                {
                    return;
                }
            }

            foreach (var line in sourceLines.Select(ParsingHelpers.Clean))
            {
                var row = TryParseExperienceRow(line);
                if (row is null)
                {
                    continue;
                }

                if (row.MonthLabel.Equals("Total", StringComparison.OrdinalIgnoreCase))
                {
                    data.Totals = new ClaimsExperienceTotals
                    {
                        EhcPremium = row.EhcPremium,
                        EhcClaims = row.EhcClaims,
                        DentalPremium = row.DentalPremium,
                        DentalClaims = row.DentalClaims
                    };
                    continue;
                }

                var month = TryParseMonthLabel(row.MonthLabel);
                if (month is null || data.Months.Any(item => item.Month == month.Value))
                {
                    continue;
                }

                data.Months.Add(new ClaimsExperienceMonth
                {
                    Month = month.Value,
                    EhcPremium = row.EhcPremium,
                    EhcClaims = row.EhcClaims,
                    DentalPremium = row.DentalPremium,
                    DentalClaims = row.DentalClaims,
                    SourceFile = Path.GetFileName(pdfPath),
                    EhcPremiumMatchesMonth = row.EhcPremium != 0m,
                    DentalPremiumMatchesMonth = row.DentalPremium != 0m
                });
            }

            if (data.Months.Count == 0)
            {
                ReadSunLifeMemberBasedMonthlyExperience(pages.SelectMany(page => page.Lines), data, pdfPath);
            }

            data.Months.Sort((left, right) => left.Month.CompareTo(right.Month));

            if (data.Months.Count == 0 &&
                !TryReadAnnualExperienceSummary(pages.SelectMany(page => page.Lines), data, pdfPath))
            {
                warnings.Add("Monthly EHC/Dental experience rows could not be read from the experience report PDF.");
            }
        }
        catch (Exception ex)
        {
            warnings.Add($"Experience report PDF could not be read: {ex.Message}");
        }
    }

    private static void ReadExperienceWorkbook(
        string workbookPath,
        ClaimsExperienceDataModelData data,
        List<string> warnings)
    {
        try
        {
            using var workbook = OpenWorkbookReadOnly(workbookPath);
            if (!TryGetExperienceWorkbookSheet(workbook, out var sheet, out var headerRow))
            {
                warnings.Add($"Monthly EHC/Dental experience rows could not be read from {Path.GetFileName(workbookPath)}.");
                return;
            }

            if (TryReadWorkbookPeriod(sheet, out var periodStart, out var periodEnd))
            {
                data.PeriodStart = periodStart;
                data.PeriodEnd = periodEnd;
                data.PeriodLabel = FormatPeriodLabel(periodStart, periodEnd);
            }

            foreach (var row in sheet.RowsUsed().Where(row => row.RowNumber() < headerRow))
            {
                var policyMatch = Regex.Match(
                    CellText(row.Cell(1)),
                    @"\bPolicy\s*:\s*(?<number>[A-Za-z0-9-]+)",
                    RegexOptions.IgnoreCase);
                if (policyMatch.Success)
                {
                    data.PolicyNumber = FirstNonBlank(data.PolicyNumber, policyMatch.Groups["number"].Value);
                    break;
                }
            }

            foreach (var row in sheet.RowsUsed().Where(row => row.RowNumber() > headerRow))
            {
                var monthLabel = CellText(row.Cell(1));
                if (monthLabel.Equals("Total", StringComparison.OrdinalIgnoreCase))
                {
                    data.Totals = new ClaimsExperienceTotals
                    {
                        EhcPremium = Number(row.Cell(2)),
                        EhcClaims = Number(row.Cell(3)),
                        DentalPremium = Number(row.Cell(5)),
                        DentalClaims = Number(row.Cell(6))
                    };
                    continue;
                }

                var month = TryParseMonthLabel(monthLabel);
                if (!month.HasValue || data.Months.Any(item => item.Month == month.Value))
                {
                    continue;
                }

                var ehcPremium = Number(row.Cell(2));
                var dentalPremium = Number(row.Cell(5));
                data.Months.Add(new ClaimsExperienceMonth
                {
                    Month = month.Value,
                    EhcPremium = ehcPremium,
                    EhcClaims = Number(row.Cell(3)),
                    DentalPremium = dentalPremium,
                    DentalClaims = Number(row.Cell(6)),
                    SourceFile = Path.GetFileName(workbookPath),
                    EhcPremiumMatchesMonth = ehcPremium != 0m,
                    DentalPremiumMatchesMonth = dentalPremium != 0m
                });
            }

            data.Months.Sort((left, right) => left.Month.CompareTo(right.Month));
            if (data.Months.Count == 0)
            {
                warnings.Add($"Monthly EHC/Dental experience rows could not be read from {Path.GetFileName(workbookPath)}.");
            }
        }
        catch (Exception ex)
        {
            warnings.Add($"Experience report workbook could not be read ({Path.GetFileName(workbookPath)}): {ex.Message}");
        }
    }

    private void ReadAdditionalExperienceReports(
        IReadOnlyList<string> sourceFiles,
        string selectedExperienceReport,
        ClaimsExperienceDataModelData data,
        List<string> warnings)
    {
        var knownPeriods = new HashSet<(DateOnly Start, DateOnly End)>();
        if (data.PeriodStart is not null && data.PeriodEnd is not null)
        {
            knownPeriods.Add((data.PeriodStart.Value, data.PeriodEnd.Value));
        }

        var historicalSlot = 0;
        foreach (var pdfPath in FindExperienceReports(sourceFiles)
                     .Where(path => !path.Equals(selectedExperienceReport, StringComparison.OrdinalIgnoreCase)))
        {
            try
            {
                var pages = _pdfTextExtractor.Extract(pdfPath);
                var lines = pages.SelectMany(page => page.Lines).ToList();
                if (GroupSourcePdfParser.TryParseMonthlyClaimsReport(pdfPath, pages, out var groupSourceReport))
                {
                    var groupSourceData = new ClaimsExperienceDataModelData();
                    MergeGroupSourceMonthlyReport(groupSourceData, groupSourceReport, warnings);
                    if (IsSameExperiencePeriod(data, groupSourceData))
                    {
                        MergeSamePeriodExperienceReport(data, groupSourceData, pdfPath, warnings);
                        if (!data.ImportedFiles.Any(record => record.FullPath.Equals(pdfPath, StringComparison.OrdinalIgnoreCase)))
                        {
                            var import = CreateImportRecord(pdfPath, "Supplemental GroupSource monthly claims PDF");
                            import.RowsRead = $"{groupSourceReport.Months.Count} {groupSourceReport.Benefit} month(s)";
                            import.Status = groupSourceReport.Warnings.Count == 0 ? "Verified" : "NeedsReview";
                            import.Notes = "Carrier-reported billed premium, paid claims, and paid loss ratio.";
                            data.ImportedFiles.Add(import);
                        }
                    }

                    continue;
                }

                if (!IsSunLifeMemberBasedExperienceReport(lines))
                {
                    continue;
                }

                var fallbackData = new ClaimsExperienceDataModelData();
                var allText = ParsingHelpers.Clean(string.Join(" ", pages.Select(page => page.Text)));
                if (TryExtractExperienceReportPeriod(allText, out var periodStart, out var periodEnd))
                {
                    fallbackData.PeriodStart = periodStart;
                    fallbackData.PeriodEnd = periodEnd;
                    fallbackData.PeriodLabel = FormatPeriodLabel(periodStart, periodEnd);
                }

                ReadSunLifeMemberBasedMonthlyExperience(lines, fallbackData, pdfPath);
                MergeExperienceReportLivesFallback(data, fallbackData);

                if (IsSameExperiencePeriod(data, fallbackData))
                {
                    MergeSamePeriodExperienceReport(data, fallbackData, pdfPath, warnings);
                    if ((fallbackData.Months.Count > 0 || fallbackData.BillingInsuredLivesDetails.Count > 0) &&
                        !data.ImportedFiles.Any(record => record.FullPath.Equals(pdfPath, StringComparison.OrdinalIgnoreCase)))
                    {
                        data.ImportedFiles.Add(CreateImportRecord(pdfPath, "Supplemental PDF experience report"));
                    }

                    continue;
                }

                if (fallbackData.PeriodStart is not null &&
                    fallbackData.PeriodEnd is not null &&
                    knownPeriods.Add((fallbackData.PeriodStart.Value, fallbackData.PeriodEnd.Value)))
                {
                    var historicalMonths = fallbackData.Months
                        .Where(month => !month.IsAnnualSummary)
                        .OrderBy(month => month.Month)
                        .ToList();
                    if (historicalMonths.Count > 0)
                    {
                        if (historicalSlot == 0)
                        {
                            data.PriorMonths.Clear();
                            data.PriorMonths.AddRange(historicalMonths);
                        }
                        else if (historicalSlot == 1)
                        {
                            data.PriorPriorMonths.Clear();
                            data.PriorPriorMonths.AddRange(historicalMonths);
                        }

                        historicalSlot++;
                    }
                }

                if ((fallbackData.Months.Count > 0 || fallbackData.BillingInsuredLivesDetails.Count > 0) &&
                    !data.ImportedFiles.Any(record => record.FullPath.Equals(pdfPath, StringComparison.OrdinalIgnoreCase)))
                {
                    data.ImportedFiles.Add(CreateImportRecord(pdfPath, "Historical PDF experience report"));
                }
            }
            catch (Exception ex)
            {
                warnings.Add($"Historical monthly experience could not be read from {Path.GetFileName(pdfPath)}: {ex.Message}");
            }
        }
    }

    private static bool TryParseRenewalExperiencePeriod(string value, out DateOnly start, out DateOnly end)
    {
        start = default;
        end = default;
        var parts = Regex.Split(value ?? "", @"\s+-\s+")
            .Where(part => !string.IsNullOrWhiteSpace(part))
            .ToArray();
        var formats = new[] { "dd-MMM-yyyy", "d-MMM-yyyy", "MMMM d yyyy", "MMM d yyyy" };
        return parts.Length == 2 &&
               DateOnly.TryParseExact(parts[0].Trim(), formats, Culture, DateTimeStyles.AllowWhiteSpaces, out start) &&
               DateOnly.TryParseExact(parts[1].Trim(), formats, Culture, DateTimeStyles.AllowWhiteSpaces, out end) &&
               end >= start;
    }

    private static string NormalizeSunLifeExperienceBenefit(string value)
    {
        if (value.Contains("Dental", StringComparison.OrdinalIgnoreCase))
        {
            return CoverageNames.Dental;
        }

        return value.Contains("Health", StringComparison.OrdinalIgnoreCase) ||
               value.Equals("EHC", StringComparison.OrdinalIgnoreCase)
            ? CoverageNames.Ehc
            : "";
    }

    private static List<string> FindNearbySunLifeRenewalExperienceReports(
        string sourceFolder,
        IReadOnlyList<string> sourceFiles)
    {
        var paths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var path in sourceFiles.Where(IsSunLifeRenewalExperienceReportName))
        {
            if (File.Exists(path))
            {
                paths.Add(path);
            }
        }

        AddFolder(sourceFolder);
        AddFolder(string.IsNullOrWhiteSpace(sourceFolder)
            ? null
            : Directory.GetParent(sourceFolder)?.FullName);
        return paths.OrderBy(path => Path.GetFileName(path), StringComparer.OrdinalIgnoreCase).ToList();

        void AddFolder(string? folder)
        {
            if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
            {
                return;
            }

            foreach (var path in Directory.EnumerateFiles(folder, "*.pdf", SearchOption.TopDirectoryOnly)
                         .Where(IsSunLifeRenewalExperienceReportName))
            {
                paths.Add(path);
            }
        }
    }

    private static bool IsSunLifeRenewalExperienceReportName(string path)
    {
        return NormalizeHeader(Path.GetFileNameWithoutExtension(path))
            .Contains("renewalexperience", StringComparison.OrdinalIgnoreCase);
    }

    private static bool RenewalExperienceIdentityMatches(
        string path,
        CarrierPdfParseResult parsed,
        ClaimsExperienceDataModelData data,
        List<string> warnings)
    {
        if (!string.IsNullOrWhiteSpace(data.PolicyNumber) &&
            !string.IsNullOrWhiteSpace(parsed.Pdf.PolicyNumber) &&
            !NormalizeClaimsIdentifier(data.PolicyNumber).Equals(
                NormalizeClaimsIdentifier(parsed.Pdf.PolicyNumber),
                StringComparison.OrdinalIgnoreCase))
        {
            warnings.Add($"Ignored renewal experience report {Path.GetFileName(path)} because contract/policy '{parsed.Pdf.PolicyNumber}' does not match '{data.PolicyNumber}'.");
            return false;
        }

        if (!string.IsNullOrWhiteSpace(data.ClientName) &&
            !string.IsNullOrWhiteSpace(parsed.Pdf.ClientName) &&
            !BookletClientIdentityGuardrail.ClientNamesLikelyMatch(data.ClientName, parsed.Pdf.ClientName))
        {
            warnings.Add($"Ignored renewal experience report {Path.GetFileName(path)} because client '{parsed.Pdf.ClientName}' does not match '{data.ClientName}'.");
            return false;
        }

        return true;
    }

    private static RenewalExperienceCandidate? BuildRenewalExperienceCandidate(
        string path,
        CarrierPdfParseResult parsed)
    {
        var rows = parsed.Pdf.SunLifeExperienceRows
            .Select(row => TryParseRenewalExperiencePeriod(row.Period, out var start, out var end)
                ? new RenewalExperienceDatedRow(row, start, end)
                : null)
            .Where(row => row is not null)
            .Cast<RenewalExperienceDatedRow>()
            .ToList();
        return rows.Count == 0
            ? null
            : new RenewalExperienceCandidate(path, parsed, rows, rows.Max(row => row.End));
    }

    private RenewalExperienceCandidate? ParseRenewalExperienceCandidate(
        string path,
        ClaimsExperienceDataModelData data,
        List<string> warnings)
    {
        try
        {
            var parsed = new CarrierPdfParser(_pdfTextExtractor).Parse(path);
            if (!parsed.Carrier.Equals("Sun Life", StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            if (parsed.Pdf.SunLifeExperienceRows.Count == 0 ||
                !RenewalExperienceIdentityMatches(path, parsed, data, warnings))
            {
                return null;
            }

            return BuildRenewalExperienceCandidate(path, parsed);
        }
        catch (Exception ex)
        {
            warnings.Add($"Renewal experience comparison report {Path.GetFileName(path)} could not be read: {ex.Message}");
            return null;
        }
    }

    private RenewalExperienceCandidate? SelectSunLifeRenewalExperienceReport(
        string sourceFolder,
        IReadOnlyList<string> sourceFiles,
        ClaimsExperienceDataModelData data,
        List<string> warnings)
    {
        var parsedCandidates = new List<RenewalExperienceCandidate>();
        foreach (var path in FindNearbySunLifeRenewalExperienceReports(sourceFolder, sourceFiles))
        {
            var candidate = ParseRenewalExperienceCandidate(path, data, warnings);
            if (candidate is not null)
            {
                parsedCandidates.Add(candidate);
            }
        }

        return parsedCandidates
            .OrderByDescending(candidate => candidate.LatestEnd)
            .ThenByDescending(candidate => File.GetLastWriteTimeUtc(candidate.Path))
            .FirstOrDefault();
    }

    private static void AddRenewalExperienceFinancial(
        RenewalExperienceDatedRow item,
        string sourceFile,
        ClaimsExperienceDataModelData data)
    {
        var benefit = NormalizeSunLifeExperienceBenefit(item.Row.Benefit);
        if (string.IsNullOrWhiteSpace(benefit))
        {
            return;
        }

        if (!data.RenewalExperienceFinancials.Any(existing =>
                existing.Benefit.Equals(benefit, StringComparison.OrdinalIgnoreCase) &&
                existing.Start == item.Start &&
                existing.End == item.End))
        {
            data.RenewalExperienceFinancials.Add(new RenewalExperienceFinancial
            {
                Benefit = benefit,
                Start = item.Start,
                End = item.End,
                AverageCoveredEmployees = item.Row.AverageCoveredEmployees,
                BilledPremium = item.Row.BilledPremium,
                AdjustedPremium = item.Row.AdjustedPremium,
                PaidClaims = item.Row.PaidClaims,
                IncurredClaims = item.Row.IncurredClaims,
                AdjustedLossRatio = item.Row.AdjustedLossRatio,
                Weighting = item.Row.Weighting,
                PoolingPremium = item.Row.LargeAmountPoolingPremium.GetValueOrDefault() +
                                 item.Row.QuebecIndustryPoolingPremium.GetValueOrDefault(),
                SourceFile = sourceFile,
                SourceEvidence = item.Row.SourceReference
            });
        }

        if (!data.SourcePeriods.Any(period => period.Start == item.Start && period.End == item.End))
        {
            data.SourcePeriods.Add(new ClaimsSourcePeriod
            {
                SourceRole = "Renewal Experience",
                Start = item.Start,
                End = item.End,
                MonthCount = InclusiveMonthCount(item.Start, item.End),
                SourceFile = sourceFile
            });
        }
    }

    private static void MergeSunLifeRenewalExperienceComparison(
        RenewalExperienceCandidate candidate,
        ClaimsExperienceDataModelData data)
    {
        data.ClientName = FirstNonBlank(data.ClientName, candidate.Parsed.Pdf.ClientName);
        data.PolicyNumber = FirstNonBlank(data.PolicyNumber, candidate.Parsed.Pdf.PolicyNumber);
        data.Carrier = FirstNonBlank(data.Carrier, candidate.Parsed.Carrier, candidate.Parsed.Pdf.Carrier);
        var sourceFile = Path.GetFileName(candidate.Path);
        foreach (var row in candidate.Rows)
        {
            AddRenewalExperienceFinancial(row, sourceFile, data);
        }

        if (data.RenewalExperienceFinancials.Count == 0 ||
            data.ImportedFiles.Any(record => record.FullPath.Equals(candidate.Path, StringComparison.OrdinalIgnoreCase)))
        {
            return;
        }

        var import = CreateImportRecord(candidate.Path, "Sun Life renewal experience comparison report");
        import.RowsRead = $"{data.RenewalExperienceFinancials.Count} EHC/Dental financial row(s)";
        import.Status = "Verified";
        import.Notes = "Client and policy matched; full reported experience periods are available for PowerPoint comparison charts.";
        data.ImportedFiles.Add(import);
    }

    private void ReadSunLifeRenewalExperienceComparisons(
        string sourceFolder,
        IReadOnlyList<string> sourceFiles,
        ClaimsExperienceDataModelData data,
        List<string> warnings)
    {
        var selected = SelectSunLifeRenewalExperienceReport(sourceFolder, sourceFiles, data, warnings);
        if (selected is null)
        {
            return;
        }

        MergeSunLifeRenewalExperienceComparison(selected, data);
    }

    private static bool IsSunLifeMemberBasedExperienceReport(IEnumerable<string> lines)
    {
        return lines.Any(line => line.Contains("Lives/Premium/Claims", StringComparison.OrdinalIgnoreCase) &&
                                 line.Contains("Member Based", StringComparison.OrdinalIgnoreCase));
    }

    private static void MergeExperienceReportLivesFallback(
        ClaimsExperienceDataModelData target,
        ClaimsExperienceDataModelData fallback)
    {
        foreach (var pair in fallback.BillingMonthlyLivesByMonth.OrderBy(pair => pair.Key))
        {
            if (!target.BillingMonthlyLivesByMonth.TryGetValue(pair.Key, out var existing))
            {
                target.BillingMonthlyLivesByMonth[pair.Key] = pair.Value;
                continue;
            }

            existing.HealthInsuredLives ??= pair.Value.HealthInsuredLives;
            existing.DentalInsuredLives ??= pair.Value.DentalInsuredLives;
            existing.HealthPremium ??= pair.Value.HealthPremium;
            existing.DentalPremium ??= pair.Value.DentalPremium;
            existing.SourceFile = FirstNonBlank(existing.SourceFile, pair.Value.SourceFile);
        }

        foreach (var row in fallback.BillingInsuredLivesDetails)
        {
            var alreadyPresent = target.BillingInsuredLivesDetails.Any(existing =>
                existing.StatementMonth == row.StatementMonth &&
                existing.Benefit.Equals(row.Benefit, StringComparison.OrdinalIgnoreCase) &&
                existing.DepCode.Equals(row.DepCode, StringComparison.OrdinalIgnoreCase));
            if (!alreadyPresent)
            {
                target.BillingInsuredLivesDetails.Add(row);
            }
        }
    }

    private static bool IsSameExperiencePeriod(
        ClaimsExperienceDataModelData target,
        ClaimsExperienceDataModelData supplemental)
    {
        if (target.PeriodStart is not null &&
            target.PeriodEnd is not null &&
            supplemental.PeriodStart is not null &&
            supplemental.PeriodEnd is not null)
        {
            return target.PeriodStart == supplemental.PeriodStart &&
                   target.PeriodEnd == supplemental.PeriodEnd;
        }

        var targetMonths = target.Months
            .Where(month => !month.IsAnnualSummary)
            .Select(month => FirstDayOfMonth(month.Month))
            .ToHashSet();
        var supplementalMonths = supplemental.Months
            .Where(month => !month.IsAnnualSummary)
            .Select(month => FirstDayOfMonth(month.Month))
            .ToHashSet();

        return targetMonths.Count > 0 && targetMonths.SetEquals(supplementalMonths);
    }

    private static void MergeSamePeriodExperienceReport(
        ClaimsExperienceDataModelData target,
        ClaimsExperienceDataModelData supplemental,
        string pdfPath,
        List<string> warnings)
    {
        var targetHadEhc = target.Months.Any(month =>
            !month.IsAnnualSummary && !string.IsNullOrWhiteSpace(month.EhcSourceFile));
        var targetHadDental = target.Months.Any(month =>
            !month.IsAnnualSummary && !string.IsNullOrWhiteSpace(month.DentalSourceFile));
        var supplementalHasEhc = supplemental.Months.Any(month =>
            !month.IsAnnualSummary && !string.IsNullOrWhiteSpace(month.EhcSourceFile));
        var supplementalHasDental = supplemental.Months.Any(month =>
            !month.IsAnnualSummary && !string.IsNullOrWhiteSpace(month.DentalSourceFile));
        var conflictingBenefits = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var incoming in supplemental.Months.Where(month => !month.IsAnnualSummary))
        {
            var existing = GetOrCreateMonth(target.Months, incoming.Month);
            existing.SourceFile = FirstNonBlank(existing.SourceFile, incoming.SourceFile);

            if (!string.IsNullOrWhiteSpace(incoming.EhcSourceFile))
            {
                if (string.IsNullOrWhiteSpace(existing.EhcSourceFile))
                {
                    existing.EhcPremium = incoming.EhcPremium;
                    existing.EhcClaims = incoming.EhcClaims;
                    existing.EhcSourceFile = incoming.EhcSourceFile;
                    existing.EhcPremiumMatchesMonth = incoming.EhcPremiumMatchesMonth;
                }
                else if (Math.Abs(existing.EhcPremium - incoming.EhcPremium) > 0.01m ||
                         Math.Abs(existing.EhcClaims - incoming.EhcClaims) > 0.01m)
                {
                    conflictingBenefits.Add(CoverageNames.Ehc);
                }
            }

            if (!string.IsNullOrWhiteSpace(incoming.DentalSourceFile))
            {
                if (string.IsNullOrWhiteSpace(existing.DentalSourceFile))
                {
                    existing.DentalPremium = incoming.DentalPremium;
                    existing.DentalClaims = incoming.DentalClaims;
                    existing.DentalSourceFile = incoming.DentalSourceFile;
                    existing.DentalPremiumMatchesMonth = incoming.DentalPremiumMatchesMonth;
                }
                else if (Math.Abs(existing.DentalPremium - incoming.DentalPremium) > 0.01m ||
                         Math.Abs(existing.DentalClaims - incoming.DentalClaims) > 0.01m)
                {
                    conflictingBenefits.Add(CoverageNames.Dental);
                }
            }
        }

        target.Months.Sort((left, right) => left.Month.CompareTo(right.Month));
        target.Totals ??= new ClaimsExperienceTotals();
        if (!targetHadEhc && supplementalHasEhc)
        {
            target.Totals.EhcPremium = supplemental.Totals?.EhcPremium ??
                                       supplemental.Months.Where(month => !month.IsAnnualSummary).Sum(month => month.EhcPremium);
            target.Totals.EhcClaims = supplemental.Totals?.EhcClaims ??
                                      supplemental.Months.Where(month => !month.IsAnnualSummary).Sum(month => month.EhcClaims);
        }

        if (!targetHadDental && supplementalHasDental)
        {
            target.Totals.DentalPremium = supplemental.Totals?.DentalPremium ??
                                          supplemental.Months.Where(month => !month.IsAnnualSummary).Sum(month => month.DentalPremium);
            target.Totals.DentalClaims = supplemental.Totals?.DentalClaims ??
                                         supplemental.Months.Where(month => !month.IsAnnualSummary).Sum(month => month.DentalClaims);
        }

        if (conflictingBenefits.Count > 0)
        {
            warnings.Add(
                $"Supplemental Sun Life monthly report {Path.GetFileName(pdfPath)} contains conflicting " +
                $"{string.Join(" and ", conflictingBenefits.OrderBy(value => value))} values for the same period; " +
                "the primary report values were retained.");
        }
    }

    private DateOnly? ReadExperienceReportPeriodEnd(string pdfPath)
    {
        try
        {
            var pages = _pdfTextExtractor.Extract(pdfPath);
            var allText = ParsingHelpers.Clean(string.Join(" ", pages.Select(page => page.Text)));
            return TryExtractExperienceReportPeriod(allText, out _, out var periodEnd)
                ? periodEnd
                : null;
        }
        catch
        {
            return null;
        }
    }

    private static bool TryExtractExperienceReportPeriod(string allText, out DateOnly periodStart, out DateOnly periodEnd)
    {
        periodStart = default;
        periodEnd = default;

        var periodMatch = Regex.Match(
            allText,
            @"(?:Period:\s*)?(?<start>[A-Za-z]{3,9}\s+\d{1,2},?\s+\d{4})\s+to\s+(?<end>[A-Za-z]{3,9}\s+\d{1,2},?\s+\d{4})",
            RegexOptions.IgnoreCase);
        if (!periodMatch.Success ||
            !TryParseLongDate(periodMatch.Groups["start"].Value, out var startDate) ||
            !TryParseLongDate(periodMatch.Groups["end"].Value, out var endDate))
        {
            return false;
        }

        periodStart = DateOnly.FromDateTime(startDate);
        periodEnd = DateOnly.FromDateTime(endDate);
        return true;
    }

    private static ClaimsExperiencePdfRow? TryParseExperienceRow(string line)
    {
        var match = Regex.Match(
            line,
            @"^(?<month>(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)-\d{4}|Total)\s+(?<values>.+)$",
            RegexOptions.IgnoreCase);

        if (!match.Success)
        {
            return null;
        }

        var values = ParsingHelpers.ExtractNumbers(match.Groups["values"].Value);
        if (values.Count < 5)
        {
            return null;
        }

        return new ClaimsExperiencePdfRow
        {
            MonthLabel = match.Groups["month"].Value,
            EhcPremium = values[0],
            EhcClaims = values[1],
            DentalPremium = values[3],
            DentalClaims = values[4]
        };
    }

    private static void ReadSunLifeMemberBasedMonthlyExperience(
        IEnumerable<string> rawLines,
        ClaimsExperienceDataModelData data,
        string pdfPath)
    {
        var inSummary = false;
        var coverage = "";
        var rowsRead = 0;
        decimal? ehcPremiumTotal = null;
        decimal? ehcClaimsTotal = null;
        decimal? dentalPremiumTotal = null;
        decimal? dentalClaimsTotal = null;
        var sourceFile = Path.GetFileName(pdfPath);

        foreach (var line in rawLines.Select(ParsingHelpers.Clean).Where(line => line.Length > 0))
        {
            if (line.Equals("Summary", StringComparison.OrdinalIgnoreCase))
            {
                inSummary = true;
                coverage = "";
                continue;
            }

            if (line.Equals("Details", StringComparison.OrdinalIgnoreCase))
            {
                if (rowsRead > 0)
                {
                    break;
                }

                inSummary = false;
                coverage = "";
                continue;
            }

            if (!inSummary)
            {
                continue;
            }

            if (line.Equals("Dental", StringComparison.OrdinalIgnoreCase) ||
                line.StartsWith("Dental -", StringComparison.OrdinalIgnoreCase))
            {
                coverage = CoverageNames.Dental;
                continue;
            }

            if (line.Equals("Extended Health Care", StringComparison.OrdinalIgnoreCase) ||
                line.StartsWith("Extended Health Care -", StringComparison.OrdinalIgnoreCase))
            {
                coverage = CoverageNames.Ehc;
                continue;
            }

            if (string.IsNullOrWhiteSpace(coverage))
            {
                continue;
            }

            if (TryParseSunLifeMemberBasedMonthlyRow(
                    line,
                    out var month,
                    out var singleLives,
                    out var familyLives,
                    out var totalLives,
                    out var premium,
                    out var paidClaims))
            {
                var target = GetOrCreateMonth(data.Months, month);
                target.SourceFile = sourceFile;
                if (coverage.Equals(CoverageNames.Ehc, StringComparison.OrdinalIgnoreCase))
                {
                    target.EhcPremium = premium;
                    target.EhcClaims = paidClaims;
                    target.EhcSourceFile = sourceFile;
                    target.EhcPremiumMatchesMonth = premium != 0m;
                }
                else
                {
                    target.DentalPremium = premium;
                    target.DentalClaims = paidClaims;
                    target.DentalSourceFile = sourceFile;
                    target.DentalPremiumMatchesMonth = premium != 0m;
                }

                AddSunLifeMemberBasedLivesFallback(
                    data,
                    month,
                    coverage,
                    singleLives,
                    familyLives,
                    totalLives,
                    premium,
                    sourceFile);

                rowsRead++;
                continue;
            }

            if (TryParseSunLifeMemberBasedTotalRow(line, out premium, out paidClaims))
            {
                if (coverage.Equals(CoverageNames.Ehc, StringComparison.OrdinalIgnoreCase))
                {
                    ehcPremiumTotal = premium;
                    ehcClaimsTotal = paidClaims;
                }
                else
                {
                    dentalPremiumTotal = premium;
                    dentalClaimsTotal = paidClaims;
                }
            }
        }

        if (rowsRead > 0)
        {
            data.Totals ??= new ClaimsExperienceTotals
            {
                EhcPremium = ehcPremiumTotal ?? data.Months.Sum(month => month.EhcPremium),
                EhcClaims = ehcClaimsTotal ?? data.Months.Sum(month => month.EhcClaims),
                DentalPremium = dentalPremiumTotal ?? data.Months.Sum(month => month.DentalPremium),
                DentalClaims = dentalClaimsTotal ?? data.Months.Sum(month => month.DentalClaims)
            };
        }
    }

    private static bool TryParseSunLifeMemberBasedMonthlyRow(
        string line,
        out DateOnly month,
        out int singleLives,
        out int familyLives,
        out int totalLives,
        out decimal premium,
        out decimal paidClaims)
    {
        month = default;
        singleLives = 0;
        familyLives = 0;
        totalLives = 0;
        premium = 0m;
        paidClaims = 0m;

        var match = Regex.Match(
            line,
            @"^(?<month>Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)\s+(?<year>\d{4})\s+(?<single>\d[\d,]*)\s+(?<family>\d[\d,]*)\s+(?<total>\d[\d,]*)\s+(?<premium>-?\(?\d[\d,]*\.?\d*\)?)\s+(?<claims>-?\(?\d[\d,]*\.?\d*\)?)\s+\d+(?:\.\d+)?%$",
            RegexOptions.IgnoreCase);
        if (!match.Success)
        {
            return false;
        }

        var monthLabel = $"{NormalizeMonthAbbreviation(match.Groups["month"].Value)} {match.Groups["year"].Value}";
        var parsedMonth = TryParseMonthLabel(monthLabel);
        if (parsedMonth is null)
        {
            return false;
        }

        month = parsedMonth.Value;
        if (!int.TryParse(match.Groups["single"].Value.Replace(",", ""), NumberStyles.Integer, Culture, out singleLives) ||
            !int.TryParse(match.Groups["family"].Value.Replace(",", ""), NumberStyles.Integer, Culture, out familyLives) ||
            !int.TryParse(match.Groups["total"].Value.Replace(",", ""), NumberStyles.Integer, Culture, out totalLives))
        {
            return false;
        }

        premium = ParsingHelpers.ParseDecimal(match.Groups["premium"].Value) ?? 0m;
        paidClaims = ParsingHelpers.ParseDecimal(match.Groups["claims"].Value) ?? 0m;
        return true;
    }

    private static void AddSunLifeMemberBasedLivesFallback(
        ClaimsExperienceDataModelData data,
        DateOnly month,
        string coverage,
        int singleLives,
        int familyLives,
        int totalLives,
        decimal premium,
        string sourceFile)
    {
        var monthKey = FirstDayOfMonth(month);
        if (!data.BillingMonthlyLivesByMonth.TryGetValue(monthKey, out var monthlyLives))
        {
            monthlyLives = new BillingMonthlyLives
            {
                Month = monthKey,
                SourceFile = sourceFile
            };
            data.BillingMonthlyLivesByMonth[monthKey] = monthlyLives;
        }

        if (coverage.Equals(CoverageNames.Ehc, StringComparison.OrdinalIgnoreCase))
        {
            monthlyLives.HealthInsuredLives ??= totalLives;
            monthlyLives.HealthPremium ??= premium;
        }
        else
        {
            monthlyLives.DentalInsuredLives ??= totalLives;
            monthlyLives.DentalPremium ??= premium;
        }

        monthlyLives.SourceFile = FirstNonBlank(monthlyLives.SourceFile, sourceFile);
        var statementPeriod = $"{monthKey.ToString("MMM d yyyy", Culture)} to {EndOfMonth(monthKey).ToString("MMM d yyyy", Culture)}";
        AddDetail("Single", singleLives);
        AddDetail("Family", familyLives);

        void AddDetail(string depCode, int lives)
        {
            if (data.BillingInsuredLivesDetails.Any(existing =>
                    existing.StatementMonth == monthKey &&
                    existing.Benefit.Equals(coverage, StringComparison.OrdinalIgnoreCase) &&
                    existing.DepCode.Equals(depCode, StringComparison.OrdinalIgnoreCase)))
            {
                return;
            }

            data.BillingInsuredLivesDetails.Add(new BillingInsuredLivesDetailRow
            {
                SourceFile = sourceFile,
                StatementPeriod = statementPeriod,
                StatementMonth = monthKey,
                Benefit = coverage,
                DepCode = depCode,
                Lives = lives,
                ExtractedCategory = coverage.Equals(CoverageNames.Ehc, StringComparison.OrdinalIgnoreCase) ? "Health" : "Dental",
                IsExperienceReportFallback = true
            });
        }
    }

    private static bool TryParseSunLifeMemberBasedTotalRow(string line, out decimal premium, out decimal paidClaims)
    {
        premium = 0m;
        paidClaims = 0m;

        var match = Regex.Match(
            line,
            @"^Total\s+\d+\s+\d+\s+\d+\s+(?<premium>-?\(?\d[\d,]*\.?\d*\)?)\s+(?<claims>-?\(?\d[\d,]*\.?\d*\)?)\s+\d+(?:\.\d+)?%$",
            RegexOptions.IgnoreCase);
        if (!match.Success)
        {
            return false;
        }

        premium = ParsingHelpers.ParseDecimal(match.Groups["premium"].Value) ?? 0m;
        paidClaims = ParsingHelpers.ParseDecimal(match.Groups["claims"].Value) ?? 0m;
        return true;
    }

    private static string NormalizeMonthAbbreviation(string value)
    {
        return value.Equals("Sept", StringComparison.OrdinalIgnoreCase) ? "Sep" : value;
    }

    private static bool TryReadAnnualExperienceSummary(IEnumerable<string> rawLines, ClaimsExperienceDataModelData data, string pdfPath)
    {
        decimal? ehcPremium = null;
        decimal? ehcClaims = null;
        decimal? dentalPremium = null;
        decimal? dentalClaims = null;
        DateOnly? periodStart = null;
        DateOnly? periodEnd = null;
        string pendingBenefit = "";

        foreach (var line in rawLines.Select(ParsingHelpers.Clean).Where(line => line.Length > 0))
        {
            if (line.Contains("Extended Health Care", StringComparison.OrdinalIgnoreCase))
            {
                pendingBenefit = CoverageNames.Ehc;
                continue;
            }

            if (line.Equals("Dental", StringComparison.OrdinalIgnoreCase))
            {
                pendingBenefit = CoverageNames.Dental;
                continue;
            }

            if (string.IsNullOrWhiteSpace(pendingBenefit))
            {
                continue;
            }

            var match = Regex.Match(
                line,
                @"^(?<start>\d{2}[A-Za-z]{3}\d{4})-(?<end>\d{2}[A-Za-z]{3}\d{4})\s+(?<values>.+)$",
                RegexOptions.IgnoreCase);
            if (!match.Success)
            {
                continue;
            }

            if (TryParseCompactDate(match.Groups["start"].Value, out var startDate))
            {
                periodStart ??= DateOnly.FromDateTime(startDate);
            }

            if (TryParseCompactDate(match.Groups["end"].Value, out var endDate))
            {
                periodEnd ??= DateOnly.FromDateTime(endDate);
            }

            var values = ParsingHelpers.ExtractNumbers(match.Groups["values"].Value);
            if (values.Count >= 2)
            {
                if (pendingBenefit.Equals(CoverageNames.Ehc, StringComparison.OrdinalIgnoreCase))
                {
                    ehcPremium = values[0];
                    ehcClaims = values[1];
                }
                else if (pendingBenefit.Equals(CoverageNames.Dental, StringComparison.OrdinalIgnoreCase))
                {
                    dentalPremium = values[0];
                    dentalClaims = values[1];
                }
            }

            pendingBenefit = "";
        }

        if (!ehcPremium.HasValue && !ehcClaims.HasValue && !dentalPremium.HasValue && !dentalClaims.HasValue)
        {
            return false;
        }

        var month = periodStart ?? new DateOnly(DateTime.Today.Year, DateTime.Today.Month, 1);
        data.PeriodStart ??= periodStart;
        data.PeriodEnd ??= periodEnd;
        if (periodStart.HasValue && periodEnd.HasValue)
        {
            data.PeriodLabel = $"{periodStart.Value.ToDateTime(TimeOnly.MinValue):MMM yyyy} - {periodEnd.Value.ToDateTime(TimeOnly.MinValue):MMM yyyy}";
        }

        data.Totals ??= new ClaimsExperienceTotals
        {
            EhcPremium = ehcPremium ?? 0m,
            EhcClaims = ehcClaims ?? 0m,
            DentalPremium = dentalPremium ?? 0m,
            DentalClaims = dentalClaims ?? 0m
        };

        data.Months.Add(new ClaimsExperienceMonth
        {
            Month = FirstDayOfMonth(month),
            EhcPremium = ehcPremium ?? 0m,
            EhcClaims = ehcClaims ?? 0m,
            DentalPremium = dentalPremium ?? 0m,
            DentalClaims = dentalClaims ?? 0m,
            SourceFile = Path.GetFileName(pdfPath),
            IsAnnualSummary = true
        });

        return true;
    }

    private void ReadBillingInsuredLivesForSourceFolders(string sourceFolder, ClaimsExperienceDataModelData data, List<string> warnings)
    {
        if (string.IsNullOrWhiteSpace(sourceFolder))
        {
            warnings.Add("No billings folder found.");
            return;
        }

        var currentFolder = ResolveCurrentSourceFolder(sourceFolder);
        ReadBillingInsuredLives(currentFolder, data, warnings, "No current billings folder found.");

        var priorFolder = Path.Combine(sourceFolder, "Prior");
        if (Directory.Exists(priorFolder))
        {
            ReadBillingInsuredLives(priorFolder, data, warnings, "No prior billings folder found.");
        }

        var priorPriorFolder = FindPriorPriorFolder(sourceFolder);
        if (priorPriorFolder is not null)
        {
            ReadBillingInsuredLives(priorPriorFolder, data, warnings, "No prior-prior billings folder found.");
        }
    }

    private void ReadBillingInsuredLives(string sourceFolder, ClaimsExperienceDataModelData data, List<string> warnings, string missingFolderWarning)
    {
        if (string.IsNullOrWhiteSpace(sourceFolder))
        {
            warnings.Add(missingFolderWarning);
            return;
        }

        var billingsFolder = Path.Combine(sourceFolder, "billings");
        if (!Directory.Exists(billingsFolder))
        {
            warnings.Add(missingFolderWarning);
            return;
        }

        var extraction = extract_billing_insured_lives(billingsFolder, warnings);
        var billingDetailKeys = extraction.detail_rows
            .Where(row => row.StatementMonth.HasValue)
            .Select(row => $"{row.StatementMonth!.Value:yyyy-MM}|{row.ExtractedCategory.Trim().ToUpperInvariant()}")
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        data.BillingInsuredLivesDetails.RemoveAll(row =>
            row.IsExperienceReportFallback &&
            row.StatementMonth.HasValue &&
            billingDetailKeys.Contains($"{row.StatementMonth.Value:yyyy-MM}|{row.ExtractedCategory.Trim().ToUpperInvariant()}"));
        data.BillingInsuredLivesDetails.AddRange(extraction.detail_rows);
        foreach (var monthlyLives in extraction.monthly_lives)
        {
            if (data.BillingMonthlyLivesByMonth.TryGetValue(monthlyLives.Month, out var fallbackLives))
            {
                monthlyLives.HealthInsuredLives ??= fallbackLives.HealthInsuredLives;
                monthlyLives.DentalInsuredLives ??= fallbackLives.DentalInsuredLives;
                monthlyLives.HealthPremium ??= fallbackLives.HealthPremium;
                monthlyLives.DentalPremium ??= fallbackLives.DentalPremium;
            }

            data.BillingMonthlyLivesByMonth[monthlyLives.Month] = monthlyLives;
        }
    }

    private BillingInsuredLivesExtraction extract_billing_insured_lives(string billings_folder)
    {
        return extract_billing_insured_lives(billings_folder, []);
    }

    private BillingInsuredLivesExtraction extract_billing_insured_lives(string billings_folder, List<string> warnings)
    {
        var result = new BillingInsuredLivesExtraction();
        if (!Directory.Exists(billings_folder))
        {
            warnings.Add("No billings folder found.");
            return result;
        }

        var files = Directory.EnumerateFiles(billings_folder, "*.*", SearchOption.TopDirectoryOnly)
            .Where(path => IsAllowedSourceExtension(Path.GetExtension(path)))
            .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (files.Count == 0)
        {
            warnings.Add("No billing statement files found in the billings folder.");
            return result;
        }

        var fileResults = new List<BillingFileExtraction>();
        foreach (var file in files)
        {
            var fileResult = IsExcelExtension(Path.GetExtension(file))
                ? ExtractBillingLivesFromWorkbook(file, warnings)
                : ExtractBillingLivesFromPdf(file, warnings);

            result.detail_rows.AddRange(fileResult.DetailRows);
            fileResults.Add(fileResult);

            if (fileResult.StatementMonth is null)
            {
                warnings.Add($"Billing period could not be read from {Path.GetFileName(file)}.");
            }

            if (!fileResult.HealthInsuredLives.HasValue)
            {
                warnings.Add($"Health insured lives could not be found in {Path.GetFileName(file)}.");
            }

            if (!fileResult.DentalInsuredLives.HasValue)
            {
                warnings.Add($"Dental insured lives could not be found in {Path.GetFileName(file)}.");
            }
        }

        foreach (var group in fileResults.Where(item => item.StatementMonth.HasValue).GroupBy(item => item.StatementMonth!.Value))
        {
            var ordered = group
                .OrderByDescending(item => item.DatePrepared ?? item.FileModified)
                .ThenByDescending(item => item.FileModified)
                .ToList();

            if (ordered.Count > 1)
            {
                warnings.Add($"Multiple billing statements found for {group.Key.ToString("MMM yyyy", Culture)}. Using {ordered[0].SourceFile}; other files: {string.Join(", ", ordered.Skip(1).Select(item => item.SourceFile))}.");
            }

            var selected = ordered[0];
            result.monthly_lives.Add(new BillingMonthlyLives
            {
                Month = selected.StatementMonth!.Value,
                HealthInsuredLives = selected.HealthInsuredLives,
                DentalInsuredLives = selected.DentalInsuredLives,
                HealthPremium = selected.HealthPremium,
                DentalPremium = selected.DentalPremium,
                SourceFile = selected.SourceFile,
                DatePrepared = selected.DatePrepared,
                FileModified = selected.FileModified
            });
        }

        result.monthly_lives.Sort((left, right) => left.Month.CompareTo(right.Month));
        return result;
    }

    private BillingFileExtraction ExtractBillingLivesFromWorkbook(string path, List<string> warnings)
    {
        var result = CreateBillingFileExtraction(path);
        try
        {
            using var workbook = OpenWorkbookReadOnly(path);
            var workbookText = string.Join(" ", workbook.Worksheets.Select(ExtractWorksheetText));
            ApplyBillingMetadata(result, workbookText);

            foreach (var sheet in workbook.Worksheets)
            {
                ExtractBillingRowsFromWorksheet(sheet, result);
            }

            FinalizeBillingFileExtraction(result);
        }
        catch (Exception ex)
        {
            warnings.Add($"Billing Excel file could not be read ({Path.GetFileName(path)}): {ex.Message}");
        }

        return result;
    }

    private BillingFileExtraction ExtractBillingLivesFromPdf(string path, List<string> warnings)
    {
        var result = CreateBillingFileExtraction(path);
        try
        {
            var pages = _pdfTextExtractor.Extract(path);
            var lines = pages.SelectMany(page => page.Lines).Select(ParsingHelpers.Clean).Where(line => line.Length > 0).ToList();
            var allText = ParsingHelpers.Clean(string.Join(" ", lines));
            ApplyBillingMetadata(result, allText);

            var sawBillingHeader = false;
            foreach (var line in lines)
            {
                if (IsBillingHeaderLine(line))
                {
                    sawBillingHeader = true;
                    continue;
                }

                var detail = TryCreateBillingDetailFromTextLine(line, result, requireHeaderSeen: sawBillingHeader);
                if (detail is not null)
                {
                    result.DetailRows.Add(detail);
                }
            }

            FinalizeBillingFileExtraction(result);
        }
        catch (Exception ex)
        {
            warnings.Add($"Billing PDF file could not be read ({Path.GetFileName(path)}): {ex.Message}");
        }

        return result;
    }

    private static BillingFileExtraction CreateBillingFileExtraction(string path)
    {
        return new BillingFileExtraction
        {
            SourceFile = Path.GetFileName(path),
            FullPath = path,
            FileModified = File.GetLastWriteTime(path)
        };
    }

    private static void ExtractBillingRowsFromWorksheet(IXLWorksheet sheet, BillingFileExtraction result)
    {
        var header = FindBillingHeader(sheet);
        if (header is null)
        {
            return;
        }

        var lastRow = sheet.LastRowUsed()?.RowNumber() ?? header.HeaderRow;
        for (var rowNumber = header.HeaderRow + 1; rowNumber <= lastRow; rowNumber++)
        {
            var benefit = CellText(sheet.Cell(rowNumber, header.BenefitColumn));
            if (string.IsNullOrWhiteSpace(benefit))
            {
                continue;
            }

            var category = ClassifyBillingBenefit(benefit);
            if (category is null)
            {
                continue;
            }

            var lives = TryParseLives(CellText(sheet.Cell(rowNumber, header.LivesColumn)));
            if (!lives.HasValue)
            {
                continue;
            }

            var premiumColumn = header.TotalPremiumAdjustmentsColumn ?? header.PremiumColumn;
            var premium = premiumColumn.HasValue
                ? ParsingHelpers.ParseDecimal(CellText(sheet.Cell(rowNumber, premiumColumn.Value))) ?? 0m
                : 0m;

            result.DetailRows.Add(new BillingInsuredLivesDetailRow
            {
                SourceFile = result.SourceFile,
                StatementPeriod = result.StatementPeriod,
                StatementMonth = result.StatementMonth,
                ClientName = result.ClientName,
                GroupNumber = result.GroupNumber,
                Benefit = benefit,
                DepCode = header.DepCodeColumn.HasValue ? CellText(sheet.Cell(rowNumber, header.DepCodeColumn.Value)) : "",
                Option = header.OptionColumn.HasValue ? CellText(sheet.Cell(rowNumber, header.OptionColumn.Value)) : "",
                Lives = lives.Value,
                Premium = premium,
                ExtractedCategory = category
            });
        }
    }

    private static BillingHeaderMap? FindBillingHeader(IXLWorksheet sheet)
    {
        var rows = sheet.RowsUsed().Take(80);
        foreach (var row in rows)
        {
            var lastColumn = row.LastCellUsed()?.Address.ColumnNumber ?? 0;
            if (lastColumn == 0)
            {
                continue;
            }

            int? benefitColumn = null;
            int? depCodeColumn = null;
            int? optionColumn = null;
            int? livesColumn = null;
            int? premiumColumn = null;
            int? totalPremiumAdjustmentsColumn = null;

            for (var column = 1; column <= lastColumn; column++)
            {
                var header = NormalizeHeader(CellText(row.Cell(column)));
                if (header.Length == 0)
                {
                    continue;
                }

                if (benefitColumn is null && header.Contains("benefit", StringComparison.OrdinalIgnoreCase))
                {
                    benefitColumn = column;
                }
                else if (depCodeColumn is null &&
                         (header.Contains("depcode", StringComparison.OrdinalIgnoreCase) ||
                          header.Equals("dep", StringComparison.OrdinalIgnoreCase) ||
                          header.Contains("dependentcode", StringComparison.OrdinalIgnoreCase)))
                {
                    depCodeColumn = column;
                }
                else if (optionColumn is null && header.Contains("option", StringComparison.OrdinalIgnoreCase))
                {
                    optionColumn = column;
                }
                else if (livesColumn is null && header.Contains("lives", StringComparison.OrdinalIgnoreCase))
                {
                    livesColumn = column;
                }
                else if (totalPremiumAdjustmentsColumn is null &&
                         (header.Contains("totalpremiumadjustments", StringComparison.OrdinalIgnoreCase) ||
                          header.Contains("totalpremiumandadjustments", StringComparison.OrdinalIgnoreCase)))
                {
                    totalPremiumAdjustmentsColumn = column;
                }
                else if (premiumColumn is null && header.Equals("premium", StringComparison.OrdinalIgnoreCase))
                {
                    premiumColumn = column;
                }
            }

            if (benefitColumn.HasValue && livesColumn.HasValue)
            {
                return new BillingHeaderMap(row.RowNumber(), benefitColumn.Value, livesColumn.Value, depCodeColumn, optionColumn, premiumColumn, totalPremiumAdjustmentsColumn);
            }
        }

        return null;
    }

    private static string ExtractWorksheetText(IXLWorksheet sheet)
    {
        return string.Join(" ", sheet.RowsUsed().Take(80).Select(row =>
        {
            var lastColumn = row.LastCellUsed()?.Address.ColumnNumber ?? 0;
            return string.Join(" ", Enumerable.Range(1, Math.Min(lastColumn, 12)).Select(column => CellText(row.Cell(column))));
        }));
    }

    private static void ApplyBillingMetadata(BillingFileExtraction result, string text)
    {
        if (TryExtractBillingPeriod(text, out var month, out var periodText))
        {
            result.StatementMonth = month;
            result.StatementPeriod = periodText;
        }

        result.DatePrepared = TryExtractDatePrepared(text);
        result.ClientName = FirstNonBlank(result.ClientName, ExtractLabeledValue(text, "Client Name"), ExtractLabeledValue(text, "Client"));
        result.GroupNumber = FirstNonBlank(result.GroupNumber, ExtractLabeledValue(text, "Group Number"), ExtractLabeledValue(text, "Group No"), ExtractLabeledValue(text, "Group"));
    }

    private static void FinalizeBillingFileExtraction(BillingFileExtraction result)
    {
        foreach (var detail in result.DetailRows)
        {
            detail.StatementPeriod = FirstNonBlank(detail.StatementPeriod, result.StatementPeriod);
            detail.StatementMonth ??= result.StatementMonth;
            detail.ClientName = FirstNonBlank(detail.ClientName, result.ClientName);
            detail.GroupNumber = FirstNonBlank(detail.GroupNumber, result.GroupNumber);
        }

        var healthRows = result.DetailRows.Where(row => row.ExtractedCategory.Equals("Health", StringComparison.OrdinalIgnoreCase)).ToList();
        var dentalRows = result.DetailRows.Where(row => row.ExtractedCategory.Equals("Dental", StringComparison.OrdinalIgnoreCase)).ToList();
        result.HealthInsuredLives = healthRows.Count > 0 ? healthRows.Sum(row => row.Lives) : null;
        result.DentalInsuredLives = dentalRows.Count > 0 ? dentalRows.Sum(row => row.Lives) : null;
        result.HealthPremium = healthRows.Count > 0 ? healthRows.Sum(row => row.Premium) : null;
        result.DentalPremium = dentalRows.Count > 0 ? dentalRows.Sum(row => row.Premium) : null;
    }

    private static BillingInsuredLivesDetailRow? TryCreateBillingDetailFromTextLine(string line, BillingFileExtraction result, bool requireHeaderSeen)
    {
        if (!requireHeaderSeen && !ContainsBillingBenefitKeyword(line))
        {
            return null;
        }

        var category = ClassifyBillingBenefit(line);
        if (category is null)
        {
            return null;
        }

        var lives = ExtractLivesFromBillingLine(line);
        if (!lives.HasValue)
        {
            return null;
        }

        return new BillingInsuredLivesDetailRow
        {
            SourceFile = result.SourceFile,
            StatementPeriod = result.StatementPeriod,
            StatementMonth = result.StatementMonth,
            ClientName = result.ClientName,
            GroupNumber = result.GroupNumber,
            Benefit = ExtractBenefitLabel(line, category),
            DepCode = ExtractDepCode(line),
            Option = ExtractOption(line),
            Lives = lives.Value,
            Premium = ExtractPremiumFromBillingLine(line),
            ExtractedCategory = category
        };
    }

    private static bool IsBillingHeaderLine(string line)
    {
        var cleaned = ParsingHelpers.Clean(line);
        return cleaned.Contains("Benefit", StringComparison.OrdinalIgnoreCase) &&
               cleaned.Contains("Lives", StringComparison.OrdinalIgnoreCase) &&
               (cleaned.Contains("Premium", StringComparison.OrdinalIgnoreCase) ||
                cleaned.Contains("Volume", StringComparison.OrdinalIgnoreCase) ||
                cleaned.Contains("Dep", StringComparison.OrdinalIgnoreCase));
    }

    private static bool ContainsBillingBenefitKeyword(string text)
    {
        return Regex.IsMatch(text, @"\b(EHC|Extended\s+Health|Health|Dental)\b", RegexOptions.IgnoreCase);
    }

    private static string? ClassifyBillingBenefit(string benefit)
    {
        var cleaned = ParsingHelpers.Clean(benefit);
        if (cleaned.Length == 0 || IsNonBenefitBillingRow(cleaned))
        {
            return null;
        }

        if (Regex.IsMatch(cleaned, @"\bDental\b", RegexOptions.IgnoreCase))
        {
            return "Dental";
        }

        if (Regex.IsMatch(cleaned, @"\bEHC\b|Extended\s+Health|\bHealth\b", RegexOptions.IgnoreCase) &&
            !Regex.IsMatch(cleaned, @"Health\s+Spending|HSA", RegexOptions.IgnoreCase))
        {
            return "Health";
        }

        return null;
    }

    private static bool IsNonBenefitBillingRow(string value)
    {
        return Regex.IsMatch(value, @"\b(total|subtotal|grand\s+total|summary|tax|gst|hst|pst|qst)\b", RegexOptions.IgnoreCase);
    }

    private static int? ExtractLivesFromBillingLine(string line)
    {
        var values = Regex.Matches(line, @"(?<![A-Za-z])\$?\(?-?\d{1,3}(?:,\d{3})*(?:\.\d+)?\)?|(?<![A-Za-z])\$?\(?-?\d+(?:\.\d+)?\)?")
            .Select(match => match.Value)
            .Select(TryParseLives)
            .Where(value => value.HasValue)
            .Select(value => value!.Value)
            .ToList();

        if (values.Count == 0)
        {
            return null;
        }

        if (values.Count > 1 && values[0] <= 5 && values[1] > 5)
        {
            return values[1];
        }

        return values[0];
    }

    private static decimal ExtractPremiumFromBillingLine(string line)
    {
        var lives = ExtractLivesFromBillingLine(line);
        var numbers = ParsingHelpers.ExtractNumbers(line);
        var passedLives = !lives.HasValue;
        var amounts = new List<decimal>();

        foreach (var number in numbers)
        {
            if (!passedLives && lives.HasValue && number == lives.Value)
            {
                passedLives = true;
                continue;
            }

            if (passedLives)
            {
                amounts.Add(number);
            }
        }

        return SelectTotalPremiumAdjustmentsAmount(amounts);
    }

    private static decimal SelectTotalPremiumAdjustmentsAmount(IReadOnlyList<decimal> amounts)
    {
        if (amounts.Count == 0)
        {
            return 0m;
        }

        if (amounts.Count >= 5)
        {
            return amounts[2];
        }

        if (amounts.Count == 4)
        {
            return AmountsMatch(amounts[0] + amounts[1], amounts[2])
                ? amounts[2]
                : amounts[1];
        }

        if (amounts.Count >= 3)
        {
            return AmountsMatch(amounts[0] + amounts[1], amounts[2]) && !AmountsMatch(amounts[0], amounts[1])
                ? amounts[2]
                : amounts[1];
        }

        if (amounts.Count == 2)
        {
            return amounts[1];
        }

        return amounts[0];
    }

    private static bool AmountsMatch(decimal left, decimal right)
    {
        return Math.Abs(left - right) < 0.01m;
    }

    private static int? TryParseLives(string value)
    {
        var cleaned = ParsingHelpers.Clean(value)
            .Replace("$", "", StringComparison.Ordinal)
            .Replace(",", "", StringComparison.Ordinal)
            .Replace(" ", "", StringComparison.Ordinal);

        if (cleaned.Length == 0 || cleaned.Contains('.', StringComparison.Ordinal))
        {
            return null;
        }

        return int.TryParse(cleaned, NumberStyles.Integer, CultureInfo.InvariantCulture, out var lives) && lives >= 0
            ? lives
            : null;
    }

    private static string ExtractBenefitLabel(string line, string category)
    {
        if (category.Equals("Dental", StringComparison.OrdinalIgnoreCase))
        {
            return "Dental";
        }

        var match = Regex.Match(line, @"Extended\s+Health(?:\s+Care)?|\bEHC\b|\bHealth\b", RegexOptions.IgnoreCase);
        return match.Success ? match.Value : "Health";
    }

    private static string ExtractDepCode(string line)
    {
        var match = Regex.Match(line, @"\b(Couple|Family|Single|Employee|Subscriber|EE|ES|EF)\b", RegexOptions.IgnoreCase);
        return match.Success ? match.Value : "";
    }

    private static string ExtractOption(string line)
    {
        var match = Regex.Match(line, @"\bOption\s*[:#-]?\s*(?<option>[A-Za-z0-9-]+)", RegexOptions.IgnoreCase);
        return match.Success ? match.Groups["option"].Value : "";
    }

    private static bool TryExtractBillingPeriod(string text, out DateOnly month, out string periodText)
    {
        var datePattern = @"(?:\d{4}-\d{1,2}-\d{1,2}|(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:t(?:ember)?)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\.?\s+\d{1,2},?\s+\d{4})";
        var match = Regex.Match(
            text,
            $@"(?<start>{datePattern})\s*(?:to|through|-)\s*(?<end>{datePattern})",
            RegexOptions.IgnoreCase);

        if (match.Success && TryParseFlexibleDate(match.Groups["start"].Value, out var start))
        {
            month = new DateOnly(start.Year, start.Month, 1);
            periodText = $"{ParsingHelpers.Clean(match.Groups["start"].Value)} to {ParsingHelpers.Clean(match.Groups["end"].Value)}";
            return true;
        }

        month = default;
        periodText = "";
        return false;
    }

    private static DateTime? TryExtractDatePrepared(string text)
    {
        var datePattern = @"(?:\d{4}-\d{1,2}-\d{1,2}|(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:t(?:ember)?)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\.?\s+\d{1,2},?\s+\d{4}|\d{1,2}/\d{1,2}/\d{4})";
        var match = Regex.Match(
            text,
            $@"(?:Date\s+Prepared|Prepared\s+Date|Date\s+Generated|Generated\s+Date|Prepared\s+On)\s*[:\-]?\s*(?<date>{datePattern})",
            RegexOptions.IgnoreCase);

        return match.Success && TryParseFlexibleDate(match.Groups["date"].Value, out var date)
            ? date
            : null;
    }

    private static bool TryParseFlexibleDate(string value, out DateTime date)
    {
        var cleaned = ParsingHelpers.Clean(value).Replace(".", "", StringComparison.Ordinal);
        var formats = new[]
        {
            "yyyy-M-d",
            "yyyy-MM-dd",
            "MMM d yyyy",
            "MMM d, yyyy",
            "MMMM d yyyy",
            "MMMM d, yyyy",
            "M/d/yyyy",
            "MM/dd/yyyy"
        };

        return DateTime.TryParseExact(cleaned, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out date) ||
               DateTime.TryParse(cleaned, CultureInfo.InvariantCulture, DateTimeStyles.None, out date);
    }

    private static string ExtractLabeledValue(string text, string label)
    {
        var match = Regex.Match(
            text,
            $@"\b{Regex.Escape(label)}\b\s*[:\-]?\s*(?<value>[A-Za-z0-9&().,' /_-]{{2,80}}?)(?=\s{{2,}}|\b(?:Group|Policy|Billing|Date|Period|Benefit)\b|$)",
            RegexOptions.IgnoreCase);

        return match.Success ? ParsingHelpers.Clean(match.Groups["value"].Value) : "";
    }

    private static string NormalizeHeader(string value)
    {
        return Regex.Replace(ParsingHelpers.Clean(value), @"[^A-Za-z0-9]", "");
    }

    private static string CellText(IXLCell cell)
    {
        var text = ParsingHelpers.Clean(cell.GetFormattedString());
        return text.Length > 0 ? text : ParsingHelpers.Clean(cell.GetString());
    }

    private static XLWorkbook OpenWorkbookReadOnly(string path)
    {
        try
        {
            return OpenWorkbookFromSharedRead(path);
        }
        catch (IOException)
        {
            var tempPath = Path.Combine(
                Path.GetTempPath(),
                $"claims-experience-{Guid.NewGuid():N}{Path.GetExtension(path)}");

            File.Copy(path, tempPath, overwrite: true);
            try
            {
                return OpenWorkbookFromSharedRead(tempPath);
            }
            finally
            {
                TryDeleteFile(tempPath);
            }
        }
    }

    private static XLWorkbook OpenWorkbookFromSharedRead(string path)
    {
        using var input = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        using var memory = new MemoryStream();
        input.CopyTo(memory);
        memory.Position = 0;
        return new XLWorkbook(memory);
    }

    private static void TryDeleteFile(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch
        {
            // Temporary read copies are best-effort cleanup only.
        }
    }

    private static void WarnForMissingBillingMonths(ClaimsExperienceDataModelData data, List<string> warnings)
    {
        var warned = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var monthsToCheck = SelectExperienceMonths(data.PriorPriorMonths, 12)
            .Concat(SelectExperienceMonths(data.PriorMonths, data.ReportMonthCount))
            .Concat(SelectExperienceMonths(data.Months, data.ReportMonthCount));

        foreach (var month in monthsToCheck.Select(item => FirstDayOfMonth(item.Month)).Distinct().OrderBy(item => item))
        {
            if (!data.BillingMonthlyLivesByMonth.TryGetValue(month, out var lives))
            {
                if (warned.Add($"{month:yyyy-MM}|missing"))
                {
                    warnings.Add($"No matching billing statement found for {month.ToString("MMM yyyy", Culture)}.");
                }

                continue;
            }

            if (!lives.HealthInsuredLives.HasValue && warned.Add($"{month:yyyy-MM}|health"))
            {
                warnings.Add($"Health insured lives were blank for {month.ToString("MMM yyyy", Culture)}.");
            }

            if (!lives.DentalInsuredLives.HasValue && warned.Add($"{month:yyyy-MM}|dental"))
            {
                warnings.Add($"Dental insured lives were blank for {month.ToString("MMM yyyy", Culture)}.");
            }
        }
    }

    private static void ReadClaimsSummaryWorkbook(
        string path,
        string coverage,
        ClaimsExperienceDataModelData data,
        List<string> warnings,
        string periodId,
        decimal? contractSummaryPaidTotal = null,
        string? contractSummarySourceFile = null)
    {
        try
        {
            using var workbook = OpenWorkbookReadOnly(path);
            var sheet = workbook.Worksheets.FirstOrDefault(item => item.Name.Equals("Summary", StringComparison.OrdinalIgnoreCase))
                        ?? workbook.Worksheets.First();

            data.ClientName = FirstNonBlank(data.ClientName, sheet.Cell(4, 1).GetString().Trim());
            var headerRow = FindClaimsSummaryHeaderRow(sheet);
            if (!headerRow.HasValue)
            {
                warnings.Add($"{coverage} claims workbook did not contain the supported Summary columns: {Path.GetFileName(path)}.");
                return;
            }

            var detailRows = new List<ClaimsCategorySourceRow>();
            var categorySubtotalRows = new List<ClaimsCategorySourceRow>();
            var currentBenefit = "";
            var currentCategory = "";

            foreach (var row in sheet.RowsUsed().Where(row => row.RowNumber() > headerRow.Value))
            {
                var benefit = ParsingHelpers.Clean(sheet.Cell(row.RowNumber(), 1).GetString());
                var category = ParsingHelpers.Clean(sheet.Cell(row.RowNumber(), 2).GetString());
                var service = ParsingHelpers.Clean(sheet.Cell(row.RowNumber(), 3).GetString());

                if (!string.IsNullOrWhiteSpace(benefit))
                {
                    currentBenefit = benefit;
                }

                if (!string.IsNullOrWhiteSpace(category))
                {
                    currentCategory = category;
                }

                if (string.IsNullOrWhiteSpace(currentBenefit) &&
                    string.IsNullOrWhiteSpace(currentCategory) &&
                    string.IsNullOrWhiteSpace(service))
                {
                    continue;
                }

                if (!ClaimsBenefitMatchesCoverage(currentBenefit, coverage))
                {
                    continue;
                }

                var occurrences = Number(sheet.Cell(row.RowNumber(), 4));
                var submitted = Number(sheet.Cell(row.RowNumber(), 5));
                var eligible = Number(sheet.Cell(row.RowNumber(), 6));
                var paid = Number(sheet.Cell(row.RowNumber(), 7));

                if (occurrences == 0m && submitted == 0m && eligible == 0m && paid == 0m)
                {
                    continue;
                }

                var sourceRow = new ClaimsCategorySourceRow
                {
                    Coverage = coverage,
                    Benefit = currentBenefit,
                    Category = currentCategory,
                    Service = service,
                    Occurrences = occurrences,
                    Submitted = submitted,
                    Eligible = eligible,
                    Paid = paid,
                    SourceFile = Path.GetFileName(path)
                };

                if (service.Equals("Subtotal", StringComparison.OrdinalIgnoreCase))
                {
                    categorySubtotalRows.Add(sourceRow);
                }
                else
                {
                    detailRows.Add(sourceRow);
                }
            }

            if (detailRows.Count == 0 && categorySubtotalRows.Count == 0)
            {
                warnings.Add($"{coverage} rows were not found in the selected claims Summary workbook: {Path.GetFileName(path)}.");
                return;
            }

            var categoryRows = categorySubtotalRows.Count > 0
                ? categorySubtotalRows
                : detailRows
                    .GroupBy(row => NormalizeCategory(coverage, row.Category), StringComparer.OrdinalIgnoreCase)
                    .Select(group => new ClaimsCategorySourceRow
                    {
                        Coverage = coverage,
                        Benefit = coverage,
                        Category = group.Key,
                        Service = "Subtotal",
                        Occurrences = group.Sum(row => row.Occurrences),
                        Submitted = group.Sum(row => row.Submitted),
                        Eligible = group.Sum(row => row.Eligible),
                        Paid = group.Sum(row => row.Paid),
                        SourceFile = Path.GetFileName(path)
                    })
                    .ToList();

            ValidateCategoryPaidTotal(
                coverage,
                periodId,
                categoryRows.Sum(row => row.Paid),
                contractSummaryPaidTotal,
                contractSummarySourceFile,
                warnings);

            if (coverage == CoverageNames.Ehc)
            {
                foreach (var row in categoryRows)
                {
                    data.EhcCategories.Add(new ClaimsCategoryDetail
                    {
                        PeriodId = periodId,
                        Coverage = CoverageNames.Ehc,
                        Category = NormalizeCategory(coverage, row.Category),
                        ClaimCount = (int)Math.Round(row.Occurrences, MidpointRounding.AwayFromZero),
                        SubmittedAmount = row.Submitted,
                        EligibleAmount = row.Eligible,
                        PaidClaims = row.Paid,
                        SourceFile = row.SourceFile
                    });
                }

                foreach (var row in detailRows.Where(row => row.Category.Equals("Paramedical", StringComparison.OrdinalIgnoreCase)))
                {
                    data.ParamedicalDetails.Add(new ParamedicalDetail
                    {
                        PeriodId = periodId,
                        Service = NormalizeParamedicalService(row.Service),
                        ClaimCount = (int)Math.Round(row.Occurrences, MidpointRounding.AwayFromZero),
                        SubmittedAmount = row.Submitted,
                        EligibleAmount = row.Eligible,
                        PaidClaims = row.Paid,
                        SourceFile = row.SourceFile
                    });
                }
            }
            else if (coverage == CoverageNames.Dental)
            {
                foreach (var row in categoryRows)
                {
                    data.DentalCategories.Add(new ClaimsCategoryDetail
                    {
                        PeriodId = periodId,
                        Coverage = CoverageNames.Dental,
                        Category = NormalizeCategory(coverage, row.Category),
                        ClaimCount = (int)Math.Round(row.Occurrences, MidpointRounding.AwayFromZero),
                        SubmittedAmount = row.Submitted,
                        EligibleAmount = row.Eligible,
                        PaidClaims = row.Paid,
                        SourceFile = row.SourceFile
                    });
                }
            }
        }
        catch (Exception ex)
        {
            warnings.Add($"{coverage} claims workbook could not be read: {ex.Message}");
        }
    }

    private static void ValidateCategoryPaidTotal(
        string coverage,
        string periodId,
        decimal categoryPaidTotal,
        decimal? contractSummaryPaidTotal,
        string? contractSummarySourceFile,
        List<string> warnings)
    {
        if (!contractSummaryPaidTotal.HasValue)
        {
            return;
        }

        if (Math.Abs(categoryPaidTotal - contractSummaryPaidTotal.Value) <= 0.02m)
        {
            return;
        }

        warnings.Add(
            $"{coverage} {periodId} category paid claims total {categoryPaidTotal.ToString("C", Culture)} does not match Contract Summary grand total {contractSummaryPaidTotal.Value.ToString("C", Culture)} from {Path.GetFileName(contractSummarySourceFile ?? "")}.");
    }

    private static decimal? ReadContractSummaryPaidGrandTotal(string? path, string coverage, List<string> warnings)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return null;
        }

        try
        {
            using var workbook = OpenWorkbookReadOnly(path);
            var sheet = workbook.Worksheets.FirstOrDefault(item => item.Name.Equals("Contract Summary", StringComparison.OrdinalIgnoreCase))
                        ?? workbook.Worksheets.FirstOrDefault(item => item.Name.Contains("Contract", StringComparison.OrdinalIgnoreCase));
            if (sheet is null)
            {
                warnings.Add($"{coverage} monthly workbook did not contain a Contract Summary sheet: {Path.GetFileName(path)}.");
                return null;
            }

            var paidColumn = FindContractSummaryPaidColumn(sheet);
            if (!paidColumn.HasValue)
            {
                warnings.Add($"{coverage} Contract Summary did not contain a Total amount paid column: {Path.GetFileName(path)}.");
                return null;
            }

            var grandTotalRow = FindGrandTotalRow(sheet);
            if (!grandTotalRow.HasValue)
            {
                warnings.Add($"{coverage} Contract Summary did not contain a Grand total row: {Path.GetFileName(path)}.");
                return null;
            }

            return Number(sheet.Cell(grandTotalRow.Value, paidColumn.Value));
        }
        catch (Exception ex)
        {
            warnings.Add($"{coverage} Contract Summary grand total could not be read ({Path.GetFileName(path)}): {ex.Message}");
            return null;
        }
    }

    private static int? FindContractSummaryPaidColumn(IXLWorksheet sheet)
    {
        int? paidColumn = null;
        foreach (var row in sheet.RowsUsed().Take(30))
        {
            var lastColumn = row.LastCellUsed()?.Address.ColumnNumber ?? 0;
            for (var column = 1; column <= lastColumn; column++)
            {
                var normalized = NormalizeHeader(CellText(row.Cell(column)));
                if (normalized.Contains("Totalamountpaid", StringComparison.OrdinalIgnoreCase))
                {
                    paidColumn = column;
                }
            }
        }

        return paidColumn;
    }

    private static int? FindGrandTotalRow(IXLWorksheet sheet)
    {
        foreach (var row in sheet.RowsUsed())
        {
            var lastColumn = row.LastCellUsed()?.Address.ColumnNumber ?? 0;
            for (var column = 1; column <= Math.Min(lastColumn, 5); column++)
            {
                if (CellText(row.Cell(column)).Equals("Grand total", StringComparison.OrdinalIgnoreCase))
                {
                    return row.RowNumber();
                }
            }
        }

        return null;
    }

    private static void ReadMonthlyClaimsWorkbook(string path, string coverage, ClaimsExperienceDataModelData data, List<string> warnings)
    {
        try
        {
            using var workbook = OpenWorkbookReadOnly(path);
            var sheet = workbook.Worksheets.FirstOrDefault(item => item.Name.Equals("Sponsor Summary", StringComparison.OrdinalIgnoreCase))
                        ?? workbook.Worksheets.First();

            data.ClientName = FirstNonBlank(data.ClientName, sheet.Cell(4, 1).GetString().Trim());
            var monthColumns = FindMonthlyClaimColumns(sheet);
            if (monthColumns.Count == 0)
            {
                warnings.Add($"{coverage} monthly claims workbook did not contain month columns: {Path.GetFileName(path)}.");
                return;
            }

            var totalRow = FindMonthlyGrandTotalRow(sheet);
            if (totalRow is null)
            {
                warnings.Add($"{coverage} monthly claims workbook did not contain a Grand total row: {Path.GetFileName(path)}.");
                return;
            }

            foreach (var (column, month) in monthColumns)
            {
                var paidClaims = Number(sheet.Cell(totalRow.Value, column));
                var target = GetOrCreateMonth(data.Months, month);
                if (coverage.Equals(CoverageNames.Ehc, StringComparison.OrdinalIgnoreCase))
                {
                    target.EhcClaims = paidClaims;
                    target.EhcSourceFile = Path.GetFileName(path);
                }
                else if (coverage.Equals(CoverageNames.Dental, StringComparison.OrdinalIgnoreCase))
                {
                    target.DentalClaims = paidClaims;
                    target.DentalSourceFile = Path.GetFileName(path);
                }
            }

            data.Months.Sort((left, right) => left.Month.CompareTo(right.Month));
        }
        catch (Exception ex)
        {
            warnings.Add($"{coverage} monthly claims workbook could not be read ({Path.GetFileName(path)}): {ex.Message}");
        }
    }

    private static List<(int Column, DateOnly Month)> FindMonthlyClaimColumns(IXLWorksheet sheet)
    {
        var columns = new List<(int Column, DateOnly Month)>();
        foreach (var row in sheet.RowsUsed().Take(20))
        {
            var lastColumn = row.LastCellUsed()?.Address.ColumnNumber ?? 0;
            var rowColumns = new List<(int Column, DateOnly Month)>();
            for (var column = 1; column <= lastColumn; column++)
            {
                var month = TryParseMonthLabel(CellText(row.Cell(column)));
                if (month.HasValue)
                {
                    rowColumns.Add((column, month.Value));
                }
            }

            if (rowColumns.Count >= 2)
            {
                columns = rowColumns;
                break;
            }
        }

        return columns
            .Where(item => item.Month != default)
            .OrderBy(item => item.Month)
            .Take(12)
            .ToList();
    }

    private static int? FindMonthlyGrandTotalRow(IXLWorksheet sheet)
    {
        foreach (var row in sheet.RowsUsed())
        {
            var lastColumn = row.LastCellUsed()?.Address.ColumnNumber ?? 0;
            for (var column = 1; column <= Math.Min(lastColumn, 8); column++)
            {
                if (CellText(row.Cell(column)).Equals("Grand total", StringComparison.OrdinalIgnoreCase))
                {
                    return row.RowNumber();
                }
            }
        }

        return null;
    }

    private static ClaimsExperienceMonth GetOrCreateMonth(List<ClaimsExperienceMonth> months, DateOnly month)
    {
        var firstDay = FirstDayOfMonth(month);
        var existing = months.FirstOrDefault(item => FirstDayOfMonth(item.Month) == firstDay);
        if (existing is not null)
        {
            if (existing.IsAnnualSummary)
            {
                existing.EhcPremium = 0m;
                existing.EhcClaims = 0m;
                existing.DentalPremium = 0m;
                existing.DentalClaims = 0m;
                existing.SourceFile = "";
                existing.EhcSourceFile = "";
                existing.DentalSourceFile = "";
                existing.EhcPremiumMatchesMonth = false;
                existing.DentalPremiumMatchesMonth = false;
                existing.IsAnnualSummary = false;
            }

            return existing;
        }

        var created = new ClaimsExperienceMonth
        {
            Month = firstDay
        };
        months.Add(created);
        return created;
    }

    private static int WriteWorkbook(string outputPath, ClaimsExperienceDataModelData data, List<string> warnings)
    {
        using var workbook = new XLWorkbook();
        var tableCount = 0;

        tableCount += AddTable(
            workbook.AddWorksheet("00_Import_Log"),
            "Import Log",
            ["Import Timestamp", "Source Type", "Source File", "Full Path", "Rows Read", "Status", "Notes"],
            BuildImportRows(data),
            "tblImportLog",
            SheetKind.Input);

        tableCount += AddTable(
            workbook.AddWorksheet("01_Client_Setup"),
            "Client Setup",
            ["Field", "Value", "Notes"],
            BuildClientRows(data),
            "tblClientSetup",
            SheetKind.Input);

        tableCount += AddTable(
            workbook.AddWorksheet("02_Periods"),
            "Periods",
            ["Period ID", "Period Label", "Start Date", "End Date", "Month Count", "Comparison Role", "Is Active", "Notes"],
            BuildPeriodRows(data),
            "tblPeriods",
            SheetKind.Input);

        tableCount += AddTable(
            workbook.AddWorksheet("03_Monthly_Experience"),
            "Monthly Experience",
            ["Period ID", "Month", "Coverage", "Billed Premium", "Net Premium", "Paid Claims", "Fees", "Claims Plus Fees", "Lives", "Source File", "Notes"],
            BuildMonthlyRows(data),
            "tblMonthlyExperience",
            SheetKind.Input);

        tableCount += AddTable(
            workbook.AddWorksheet("Billing_Insured_Lives"),
            "Billing Insured Lives",
            ["Source File", "Statement Period", "Benefit", "Dep Code", "Lives", "Extracted Category"],
            BuildBillingInsuredLivesRows(data.BillingInsuredLivesDetails),
            "tblBillingInsuredLives",
            SheetKind.Input);

        tableCount += AddTable(
            workbook.AddWorksheet("04_EHC_Category_Detail"),
            "EHC Category Detail",
            ["Period ID", "Coverage", "Category", "Claim Count", "Submitted Amount", "Eligible Amount", "Paid Claims", "Source File", "Notes"],
            BuildCategoryRows(data.EhcCategories),
            "tblEHCCategory",
            SheetKind.Input);

        tableCount += AddTable(
            workbook.AddWorksheet("05_Paramedical_Detail"),
            "Paramedical Detail",
            ["Period ID", "Service", "Claim Count", "Submitted Amount", "Eligible Amount", "Paid Claims", "Source File", "Notes"],
            BuildParamedicalRows(data.ParamedicalDetails),
            "tblParamedical",
            SheetKind.Input);

        tableCount += AddTable(
            workbook.AddWorksheet("06_Dental_Category_Detail"),
            "Dental Category Detail",
            ["Period ID", "Coverage", "Category", "Claim Count", "Submitted Amount", "Eligible Amount", "Paid Claims", "Source File", "Notes"],
            BuildCategoryRows(data.DentalCategories),
            "tblDentalCategory",
            SheetKind.Input);

        tableCount += AddScheduleOfBenefitsSheet(workbook.AddWorksheet("Schedule_of_Benefits"), data);

        tableCount += AddTable(
            workbook.AddWorksheet("07_Assumptions"),
            "Assumptions",
            ["Assumption", "Value", "Applies To", "Notes"],
            BuildAssumptionRows(data),
            "tblAssumptions",
            SheetKind.Input);

        tableCount += AddTable(
            workbook.AddWorksheet("08_Report_Control"),
            "Report Control",
            ["Control", "Value", "Notes"],
            BuildReportControlRows(data),
            "tblReportControl",
            SheetKind.Control);

        tableCount += AddCalculationModelSheet(workbook.AddWorksheet("09_Calculation_Model"));
        tableCount += AddLivesAnalysisSheet(workbook.AddWorksheet("13_Lives_Analysis"));
        tableCount += AddRenewalExportSheet(workbook.AddWorksheet("10_Export_To_Renewal_Review"), data);
        tableCount += AddPowerPointExportSheet(workbook.AddWorksheet("11_Export_PowerPoint"), data);
        AddClaimsExperienceSummarySheet(workbook.AddWorksheet("ClaimsExperience_Summary"));
        AddChartsSheet(workbook.AddWorksheet("12_Charts"), data);

        foreach (var worksheet in workbook.Worksheets)
        {
            if (string.Equals(worksheet.Name, "ClaimsExperience_Summary", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (string.Equals(worksheet.Name, "13_Lives_Analysis", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            worksheet.Columns().AdjustToContents(10, 45);
        }

        ApplyCalculationModelColumnWidths(workbook);
        ApplyCategoryDetailColumnWidths(workbook);

        workbook.SaveAs(outputPath);
        RemoveCalculatedTableColumnFormulas(outputPath);
        AddClaimsExperienceCharts(outputPath, data);
        ForceWorkbookRecalculation(outputPath);

        if (data.Months.Count == 0)
        {
            warnings.Add("The claims data model was created, but monthly experience rows are blank because no monthly rows were read.");
        }

        return tableCount;
    }

    private static void ApplyCalculationModelColumnWidths(XLWorkbook workbook)
    {
        var sheet = workbook.Worksheet("09_Calculation_Model");
        sheet.Column(1).Width = 10.45;
        for (var column = 3; column <= 20; column++)
        {
            sheet.Column(column).Width = Math.Max(sheet.Column(column).Width, 16);
        }
    }

    private static void ApplyCategoryDetailColumnWidths(XLWorkbook workbook)
    {
        ApplyMinimumColumnWidths(workbook.Worksheet("04_EHC_Category_Detail"), 5, 7, 16);
        ApplyMinimumColumnWidths(workbook.Worksheet("05_Paramedical_Detail"), 4, 6, 16);
        ApplyMinimumColumnWidths(workbook.Worksheet("06_Dental_Category_Detail"), 5, 7, 16);
    }

    private static void ApplyMinimumColumnWidths(IXLWorksheet sheet, int firstColumn, int lastColumn, double minimumWidth)
    {
        for (var column = firstColumn; column <= lastColumn; column++)
        {
            sheet.Column(column).Width = Math.Max(sheet.Column(column).Width, minimumWidth);
        }
    }

    private static int AddCalculationModelSheet(IXLWorksheet sheet)
    {
        var headers = new[]
        {
            "Period ID",
            "Coverage",
            "Current Billed Premium",
            "Current Net Premium",
            "Current Paid Claims",
            "Current Fees",
            "Current Total Claims",
            "Current Loss Ratio",
            "Current Employee Count",
            "Current Average Spend Per Employee",
            "Prior Billed Premium",
            "Prior Net Premium",
            "Prior Paid Claims",
            "Prior Fees",
            "Prior Total Claims",
            "Prior Loss Ratio",
            "Prior Employee Count",
            "Prior Average Spend Per Employee",
            "Claims Change",
            "Loss Ratio Change",
            "Notes"
        };

        var rows = new List<object?[]>();
        var coverages = new[] { CoverageNames.Ehc, CoverageNames.Dental, CoverageNames.Hsa, CoverageNames.AsoFees, "Combined EHC + Dental", "Total" };
        for (var index = 0; index < coverages.Length; index++)
        {
            var coverage = coverages[index];
            var excelRow = 5 + index;
            var currentEmployeeFormula = AverageSpendEmployeeCountFormula(coverage, true);
            var priorEmployeeFormula = AverageSpendEmployeeCountFormula(coverage, false);
            rows.Add(
            [
                "='08_Report_Control'!$B$5",
                coverage,
                $"={SumMonthlyFormula("'08_Report_Control'!$B$5", coverage, "Billed Premium")}",
                $"={SumMonthlyFormula("'08_Report_Control'!$B$5", coverage, "Net Premium")}",
                $"={SumMonthlyFormula("'08_Report_Control'!$B$5", coverage, "Paid Claims")}",
                $"={SumMonthlyFormula("'08_Report_Control'!$B$5", coverage, "Fees")}",
                $"={SumMonthlyFormula("'08_Report_Control'!$B$5", coverage, "Claims Plus Fees")}",
                $"=IFERROR(G{excelRow}/D{excelRow},\"\")",
                currentEmployeeFormula,
                $"=IFERROR(E{excelRow}/I{excelRow},\"\")",
                $"={SumMonthlyFormula("'08_Report_Control'!$B$6", coverage, "Billed Premium")}",
                $"={SumMonthlyFormula("'08_Report_Control'!$B$6", coverage, "Net Premium")}",
                $"={SumMonthlyFormula("'08_Report_Control'!$B$6", coverage, "Paid Claims")}",
                $"={SumMonthlyFormula("'08_Report_Control'!$B$6", coverage, "Fees")}",
                $"={SumMonthlyFormula("'08_Report_Control'!$B$6", coverage, "Claims Plus Fees")}",
                $"=IFERROR(O{excelRow}/L{excelRow},\"\")",
                priorEmployeeFormula,
                $"=IFERROR(M{excelRow}/Q{excelRow},\"\")",
                $"=G{excelRow}-O{excelRow}",
                $"=IFERROR(H{excelRow}-P{excelRow},\"\")",
                coverage.Equals(CoverageNames.Hsa, StringComparison.OrdinalIgnoreCase) ||
                coverage.Equals(CoverageNames.AsoFees, StringComparison.OrdinalIgnoreCase)
                    ? "Ready for future source rows."
                    : ""
            ]);
        }

        return AddTable(sheet, "Calculation Model", headers, rows, "tblCalculationModel", SheetKind.Formula);
    }

    private static string AverageSpendEmployeeCountFormula(string coverage, bool current)
    {
        var periodFormula = current ? "'08_Report_Control'!$B$5" : "'08_Report_Control'!$B$6";
        var healthAverageFormula = LivesAverageFormula(periodFormula, CoverageNames.Ehc);
        var dentalAverageFormula = LivesAverageFormula(periodFormula, CoverageNames.Dental);
        return coverage switch
        {
            CoverageNames.Ehc => $"={healthAverageFormula}",
            CoverageNames.Dental => $"={dentalAverageFormula}",
            "Combined EHC + Dental" => $"=MAX({healthAverageFormula},{dentalAverageFormula})",
            "Total" => $"=MAX({healthAverageFormula},{dentalAverageFormula})",
            _ => ""
        };
    }

    private static int AddLivesAnalysisSheet(IXLWorksheet sheet)
    {
        var period = CurrentPeriodIdFormula();
        var healthCount = LivesMonthCountFormula(period, CoverageNames.Ehc);
        var dentalCount = LivesMonthCountFormula(period, CoverageNames.Dental);
        var rows = new List<object?[]>
        {
            new object?[] { "Selected Current Period", $"={period}", "Current", "tblReportControl", "Change Current Period ID in 08_Report_Control to CURR_12, CURR_9, or CURR_6." },
            new object?[] { "Health Month Count", $"={healthCount}", CoverageNames.Ehc, "tblMonthlyExperience", "Number of selected-period Health rows." },
            new object?[] { "Dental Month Count", $"={dentalCount}", CoverageNames.Dental, "tblMonthlyExperience", "Number of selected-period Dental rows." },
            new object?[] { "Average Health Lives", $"={LivesAverageFormula(period, CoverageNames.Ehc)}", CoverageNames.Ehc, "tblMonthlyExperience", "Sum of Health lives divided by selected month count." },
            new object?[] { "Average Dental Lives", $"={LivesAverageFormula(period, CoverageNames.Dental)}", CoverageNames.Dental, "tblMonthlyExperience", "Sum of Dental lives divided by selected month count." },
            new object?[] { "Current Health Lives", $"={CurrentLivesFormula(period, CoverageNames.Ehc)}", CoverageNames.Ehc, "tblMonthlyExperience", "Latest Health lives in the selected period." },
            new object?[] { "Current Dental Lives", $"={CurrentLivesFormula(period, CoverageNames.Dental)}", CoverageNames.Dental, "tblMonthlyExperience", "Latest Dental lives in the selected period." },
            new object?[] { "Health Variance %", "=IFERROR((B10-B8)/B8,\"\")", CoverageNames.Ehc, "Calculated", "Current Health lives compared to average Health lives." },
            new object?[] { "Dental Variance %", "=IFERROR((B11-B9)/B9,\"\")", CoverageNames.Dental, "Calculated", "Current Dental lives compared to average Dental lives." },
            new object?[] { "Health Variance Flag", "=IF(B12=\"\",\"\",IF(ABS(B12)>10%,\"Flag\",\"OK\"))", CoverageNames.Ehc, "Calculated", "Flags when Health variance is outside +/-10%." },
            new object?[] { "Dental Variance Flag", "=IF(B13=\"\",\"\",IF(ABS(B13)>10%,\"Flag\",\"OK\"))", CoverageNames.Dental, "Calculated", "Flags when Dental variance is outside +/-10%." }
        };

        var tableCount = AddTable(
            sheet,
            "Lives Variance Analysis",
            ["Output", "Value", "Coverage", "Source", "Notes"],
            rows,
            "tblLivesVarianceAnalysis",
            SheetKind.Formula);

        sheet.Column(1).Width = 26;
        sheet.Column(2).Width = 18;
        sheet.Column(3).Width = 14;
        sheet.Column(4).Width = 22;
        sheet.Column(5).Width = 58;
        sheet.Range("B6:B7").Style.NumberFormat.Format = "0";
        sheet.Range("B8:B11").Style.NumberFormat.Format = "0.0";
        sheet.Range("B12:B13").Style.NumberFormat.Format = "0.0%";
        sheet.Range("B12:B13").AddConditionalFormat().WhenGreaterThan(0.1).Font.SetFontColor(XLColor.FromHtml("#C00000"));
        sheet.Range("B12:B13").AddConditionalFormat().WhenLessThan(-0.1).Font.SetFontColor(XLColor.FromHtml("#C00000"));
        sheet.Range("A14:B15").Style.Font.Bold = true;

        return tableCount;
    }

    private static int AddRenewalExportSheet(IXLWorksheet sheet, ClaimsExperienceDataModelData data)
    {
        var headers = new[] { "Export Key", "Display Label", "Coverage", "Current Value", "Prior Value", "Change", "PowerPoint Table", "Notes" };
        static string ChangeFormula(int row) => $"=IF(OR(D{row}=\"\",E{row}=\"\"),\"\",D{row}-E{row})";
        var latestBillingLives = data.BillingMonthlyLivesByMonth.Values
            .OrderBy(item => item.Month)
            .LastOrDefault();

        var rows = new List<object?[]>
        {
            new object?[] { "CLIENT_NAME", "Client Name", "", "='01_Client_Setup'!$B$5", "", "", "tblPPT_ExecutiveSummary", "" },
            new object?[] { "EXPERIENCE_PERIOD", "Experience Period", "", "='01_Client_Setup'!$B$10", "", "", "tblPPT_ExecutiveSummary", "" },
            new object?[] { "HEALTH_INSURED_LIVES", "Health Insured Lives", CoverageNames.Ehc, latestBillingLives?.HealthInsuredLives, "", "", "", "From latest matching billing statement." },
            new object?[] { "DENTAL_INSURED_LIVES", "Dental Insured Lives", CoverageNames.Dental, latestBillingLives?.DentalInsuredLives, "", "", "", "From latest matching billing statement." },
            new object?[] { "EHC_PAID_CLAIMS", "EHC Paid Claims", CoverageNames.Ehc, "=SUMIFS(tblCalculationModel[Current Paid Claims],tblCalculationModel[Coverage],\"EHC\")", "=SUMIFS(tblCalculationModel[Prior Paid Claims],tblCalculationModel[Coverage],\"EHC\")", ChangeFormula(9), "tblPPT_AverageSpend", "" },
            new object?[] { "EHC_LOSS_RATIO", "EHC Loss Ratio", CoverageNames.Ehc, "=SUMIFS(tblCalculationModel[Current Loss Ratio],tblCalculationModel[Coverage],\"EHC\")", "=SUMIFS(tblCalculationModel[Prior Loss Ratio],tblCalculationModel[Coverage],\"EHC\")", ChangeFormula(10), "tblPPT_ExecutiveSummary", "" },
            new object?[] { "DENTAL_PAID_CLAIMS", "Dental Paid Claims", CoverageNames.Dental, "=SUMIFS(tblCalculationModel[Current Paid Claims],tblCalculationModel[Coverage],\"Dental\")", "=SUMIFS(tblCalculationModel[Prior Paid Claims],tblCalculationModel[Coverage],\"Dental\")", ChangeFormula(11), "tblPPT_AverageSpend", "" },
            new object?[] { "DENTAL_LOSS_RATIO", "Dental Loss Ratio", CoverageNames.Dental, "=SUMIFS(tblCalculationModel[Current Loss Ratio],tblCalculationModel[Coverage],\"Dental\")", "=SUMIFS(tblCalculationModel[Prior Loss Ratio],tblCalculationModel[Coverage],\"Dental\")", ChangeFormula(12), "tblPPT_ExecutiveSummary", "" },
            new object?[] { "COMBINED_TOTAL_CLAIMS", "Combined Total Claims", "Combined EHC + Dental", "=SUMIFS(tblCalculationModel[Current Total Claims],tblCalculationModel[Coverage],\"Combined EHC + Dental\")", "=SUMIFS(tblCalculationModel[Prior Total Claims],tblCalculationModel[Coverage],\"Combined EHC + Dental\")", ChangeFormula(13), "tblPPT_CombinedDistribution", "" },
            new object?[] { "AVG_SPEND_EHC", "Average Spend - EHC", CoverageNames.Ehc, "=SUMIFS(tblCalculationModel[Current Average Spend Per Employee],tblCalculationModel[Coverage],\"EHC\")", "=SUMIFS(tblCalculationModel[Prior Average Spend Per Employee],tblCalculationModel[Coverage],\"EHC\")", ChangeFormula(14), "tblPPT_AverageSpend", "" },
            new object?[] { "AVG_SPEND_DENTAL", "Average Spend - Dental", CoverageNames.Dental, "=SUMIFS(tblCalculationModel[Current Average Spend Per Employee],tblCalculationModel[Coverage],\"Dental\")", "=SUMIFS(tblCalculationModel[Prior Average Spend Per Employee],tblCalculationModel[Coverage],\"Dental\")", ChangeFormula(15), "tblPPT_AverageSpend", "" }
        };

        var tableCount = AddTable(sheet, "Export To Renewal Review", headers, rows, "tblExportRenewalReview", SheetKind.Export);
        sheet.Columns("D:F").Width = 21;
        return tableCount;
    }

    private static int AddScheduleOfBenefitsSheet(IXLWorksheet sheet, ClaimsExperienceDataModelData data)
    {
        FormatTitle(sheet, "Schedule of Benefits", SheetKind.Export);
        sheet.Cell(2, 1).Value = FirstNonBlank(data.ScheduleOfBenefitsClientName, data.ClientName, "Client Name");
        sheet.Cell(2, 1).Style.Font.Bold = true;
        sheet.Cell(2, 1).Style.Font.FontSize = 14;
        sheet.Cell(2, 1).Style.Font.FontColor = XLColor.FromHtml("#004751");
        sheet.Range(2, 1, 2, 8).Merge();

        sheet.Cell(3, 1).Value = "Booklet-derived plan design rows for PowerPoint slides 24, 25, and 26.";
        sheet.Cell(3, 1).Style.Font.FontColor = XLColor.FromHtml("#5F6B7A");
        sheet.Range(3, 1, 3, 8).Merge();

        var tableCount = AddTableAt(
            sheet,
            5,
            "Schedule of Benefits Export",
            ["Slide Page", "Page Total", "Section Order", "Client Name", "Carrier", "Class Name", "Section Name", "Detail Text", "Source File", "Notes"],
            BuildScheduleOfBenefitsRows(data),
            "tblScheduleOfBenefits",
            SheetKind.Export,
            out _);

        sheet.Column(1).Width = 10;
        sheet.Column(2).Width = 10;
        sheet.Column(3).Width = 12;
        sheet.Column(4).Width = 30;
        sheet.Column(5).Width = 22;
        sheet.Column(6).Width = 36;
        sheet.Column(7).Width = 30;
        sheet.Column(8).Width = 78;
        sheet.Column(9).Width = 34;
        sheet.Column(10).Width = 38;
        sheet.Columns(1, 10).Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
        sheet.Column(8).Style.Alignment.WrapText = true;
        sheet.Column(10).Style.Alignment.WrapText = true;

        return tableCount;
    }

    private static int AddPowerPointExportSheet(IXLWorksheet sheet, ClaimsExperienceDataModelData data)
    {
        FormatTitle(sheet, "PowerPoint Export Tables", SheetKind.Export);
        var tableCount = 0;
        var nextRow = 3;
        var averageSpendStartRow = nextRow;
        var selectedCurrentMonths = SelectExperienceMonths(data.Months, data.ReportMonthCount);

        tableCount += AddTableAt(
            sheet,
            nextRow,
            "Average Spend",
            ["Coverage", "Current Average Spend Per Employee", "Prior Average Spend Per Employee", "Change", "Notes"],
            [
                [CoverageNames.Ehc, "=SUMIFS(tblCalculationModel[Current Average Spend Per Employee],tblCalculationModel[Coverage],\"EHC\")", "=SUMIFS(tblCalculationModel[Prior Average Spend Per Employee],tblCalculationModel[Coverage],\"EHC\")", "=IFERROR(B5/C5-1,\"\")", ""],
                [CoverageNames.Dental, "=SUMIFS(tblCalculationModel[Current Average Spend Per Employee],tblCalculationModel[Coverage],\"Dental\")", "=SUMIFS(tblCalculationModel[Prior Average Spend Per Employee],tblCalculationModel[Coverage],\"Dental\")", "=IFERROR(B6/C6-1,\"\")", ""],
                ["Combined EHC + Dental", "=SUMIFS(tblCalculationModel[Current Average Spend Per Employee],tblCalculationModel[Coverage],\"Combined EHC + Dental\")", "=SUMIFS(tblCalculationModel[Prior Average Spend Per Employee],tblCalculationModel[Coverage],\"Combined EHC + Dental\")", "=IFERROR(B7/C7-1,\"\")", ""]
            ],
            "tblPPT_AverageSpend",
            SheetKind.Export,
            out nextRow);
        sheet.Range(averageSpendStartRow + 2, 4, averageSpendStartRow + 4, 4).Style.NumberFormat.Format = "0.0%";

        tableCount += AddMonthlyPptTable(sheet, nextRow, "Monthly Breakdown - EHC", "tblPPT_MonthlyBreakdown_EHC", CoverageNames.Ehc, selectedCurrentMonths, out nextRow);
        tableCount += AddMonthlyPptTable(sheet, nextRow, "Monthly Breakdown - Dental", "tblPPT_MonthlyBreakdown_Dental", CoverageNames.Dental, selectedCurrentMonths, out nextRow);

        tableCount += AddTableAt(
            sheet,
            nextRow,
            "Monthly Breakdown - Paramedical",
            ["Month", "Paid Claims", "Notes"],
            BuildParamedicalMonthlyPptRows(selectedCurrentMonths),
            "tblPPT_MonthlyBreakdown_Paramedical",
            SheetKind.Export,
            out nextRow);

        tableCount += AddDistributionPptTable(sheet, nextRow, "EHC Distribution", "tblPPT_EHCDistribution", "tblEHCCategory", "Category", out nextRow);
        tableCount += AddDistributionPptTable(sheet, nextRow, "Paramedical Distribution", "tblPPT_ParamedicalDistribution", "tblParamedical", "Service", out nextRow);
        tableCount += AddDistributionPptTable(sheet, nextRow, "Dental Distribution", "tblPPT_DentalDistribution", "tblDentalCategory", "Category", out nextRow);

        tableCount += AddTableAt(
            sheet,
            nextRow,
            "Combined Distribution",
            ["Category", "Paid Claims", "Percent of Combined Claims", "Notes"],
            BuildCombinedDistributionRows(nextRow),
            "tblPPT_CombinedDistribution",
            SheetKind.Export,
            out nextRow);

        tableCount += AddTableAt(
            sheet,
            nextRow,
            "Claims Distribution Chart",
            ["Category", "Current Paid Claims", "Prior Paid Claims", "Notes"],
            BuildClaimsDistributionChartRows(),
            "tblPPT_ClaimsDistributionChart",
            SheetKind.Export,
            out nextRow);

        tableCount += AddTableAt(
            sheet,
            nextRow,
            "Paramedical Breakdown Chart",
            ["Service", "Current Paid Claims", "Prior Paid Claims", "Notes"],
            BuildParamedicalBreakdownChartRows(),
            "tblPPT_ParamedicalBreakdownChart",
            SheetKind.Export,
            out nextRow);

        tableCount += AddTableAt(
            sheet,
            nextRow,
            "Dental Breakdown Chart",
            ["Category", "Current Paid Claims", "Prior Paid Claims", "Notes"],
            BuildDentalBreakdownChartRows(),
            "tblPPT_DentalBreakdownChart",
            SheetKind.Export,
            out nextRow);

        tableCount += AddTableAt(
            sheet,
            nextRow,
            "Executive Summary",
            ["Metric", "Value", "Source", "Notes"],
            [
                ["Client Name", "='01_Client_Setup'!$B$5", "tblClientSetup", ""],
                ["Carrier", "='01_Client_Setup'!$B$6", "tblClientSetup", ""],
                ["Experience Period", "='01_Client_Setup'!$B$10", "tblClientSetup", ""],
                ["EHC Loss Ratio", "=SUMIFS(tblCalculationModel[Current Loss Ratio],tblCalculationModel[Coverage],\"EHC\")", "tblCalculationModel", ""],
                ["Dental Loss Ratio", "=SUMIFS(tblCalculationModel[Current Loss Ratio],tblCalculationModel[Coverage],\"Dental\")", "tblCalculationModel", ""],
                ["Combined Claims", "=SUMIFS(tblCalculationModel[Current Total Claims],tblCalculationModel[Coverage],\"Combined EHC + Dental\")", "tblCalculationModel", ""]
            ],
            "tblPPT_ExecutiveSummary",
            SheetKind.Export,
            out nextRow);

        tableCount += AddTableAt(
            sheet,
            nextRow,
            "Claims Experience Summary",
            ["Placeholder Key", "Value"],
            BuildClaimsExperienceSummaryPptRows(),
            "tblPPT_ClaimsExperienceSummary",
            SheetKind.Export,
            out nextRow);

        tableCount += AddTableAt(
            sheet,
            nextRow,
            "Canonical Claims Periods",
            ["Sort Order", "Period Role", "Period Key", "Period ID", "Period Label", "Start Date", "End Date", "Month Count", "Source File"],
            BuildCanonicalClaimsPeriodRows(data),
            "tblPPT_CanonicalClaimsPeriods",
            SheetKind.Export,
            out nextRow);

        tableCount += AddTableAt(
            sheet,
            nextRow,
            "Canonical Claims Adjudication",
            ["Period Key", "Period Role", "Benefit", "Relationship", "Category", "Service", "Amount Submitted", "Amount Eligible", "Amount Paid", "Reported Paid Percentage", "Currency", "Source File", "Related Source Files", "Source Page", "Source Evidence", "Confidence", "Validation State"],
            BuildCanonicalClaimsAdjudicationRows(data),
            "tblPPT_CanonicalClaimsAdjudication",
            SheetKind.Export,
            out nextRow);

        tableCount += AddTableAt(
            sheet,
            nextRow,
            "Canonical Claims Control Totals",
            ["Period Key", "Period Role", "Total Type", "Benefit", "Relationship", "Category", "Service", "Amount Submitted", "Amount Eligible", "Amount Paid", "Reported Paid Percentage", "Currency", "Source File", "Source Page", "Source Evidence"],
            BuildCanonicalClaimsControlRows(data),
            "tblPPT_CanonicalClaimsControlTotals",
            SheetKind.Export,
            out nextRow);

        tableCount += AddTableAt(
            sheet,
            nextRow,
            "Canonical Claims Revisions",
            ["Period Key", "Period Role", "Natural Key", "Previous Source", "Revised Source", "Previous Submitted", "Revised Submitted", "Previous Eligible", "Revised Eligible", "Previous Paid", "Revised Paid", "Review Status"],
            BuildCanonicalClaimsRevisionRows(data),
            "tblPPT_CanonicalClaimsRevisions",
            SheetKind.Export,
            out nextRow);
        tableCount += AddTableAt(
            sheet,
            nextRow,
            "EHC Comparison Chart",
            ["Sort Order", "Period Role", "Period ID", "Period Label", "Net Premium", "Paid Claims", "Loss Ratio", "Notes"],
            BuildPremiumClaimsComparisonChartRows(data, nextRow, CoverageNames.Ehc, "Net Premium"),
            "tblPPT_EHC_ComparisonChart",
            SheetKind.Export,
            out nextRow);

        tableCount += AddTableAt(
            sheet,
            nextRow,
            "Renewal Experience Financials",
            ["Period Label", "Start Date", "End Date", "Month Count", "Benefit", "Average Covered Employees", "Billed Premium", "Adjusted Premium", "Paid Claims", "Incurred Claims", "Adjusted Loss Ratio", "Weighting", "Pooling Premium", "Source File", "Source Evidence"],
            BuildRenewalExperienceFinancialRows(data),
            "tblPPT_RenewalExperienceFinancials",
            SheetKind.Export,
            out nextRow);

        tableCount += AddTableAt(
            sheet,
            nextRow,
            "EHC Net L/R vs Break-Even Chart",
            ["Sort Order", "Period Role", "Period ID", "Period Label", "Net Loss Ratio", "Break Even Ratio", "Notes"],
            BuildLossRatioBreakEvenComparisonRows(data, nextRow, CoverageNames.Ehc, "Net Premium", "EHC Break-even Point (TLR)"),
            "tblPPT_EHC_LossRatioBreakEvenChart",
            SheetKind.Export,
            out nextRow);

        tableCount += AddTableAt(
            sheet,
            nextRow,
            "Dental Comparison Chart",
            ["Sort Order", "Period Role", "Period ID", "Period Label", "Billed Premium", "Paid Claims", "Loss Ratio", "Notes"],
            BuildPremiumClaimsComparisonChartRows(data, nextRow, CoverageNames.Dental, "Billed Premium"),
            "tblPPT_Dental_ComparisonChart",
            SheetKind.Export,
            out nextRow);

        tableCount += AddTableAt(
            sheet,
            nextRow,
            "Dental Net L/R vs Break-Even Chart",
            ["Sort Order", "Period Role", "Period ID", "Period Label", "Paid Loss Ratio", "Break Even Ratio", "Notes"],
            BuildLossRatioBreakEvenComparisonRows(data, nextRow, CoverageNames.Dental, "Billed Premium", "Dental Break-even Point (TLR)"),
            "tblPPT_Dental_LossRatioBreakEvenChart",
            SheetKind.Export,
            out nextRow);

        tableCount += AddTableAt(
            sheet,
            nextRow,
            "EHC Premium Claims Chart",
            ["Period Label", "Net Premium", "Paid Claims"],
            BuildPremiumClaimsChartRows(CoverageNames.Ehc, "Current Net Premium", "Prior Net Premium"),
            "tblPPT_EHC_PremiumClaimsChart",
            SheetKind.Export,
            out nextRow);

        tableCount += AddTableAt(
            sheet,
            nextRow,
            "Dental Premium Claims Chart",
            ["Period Label", "Billed Premium", "Paid Claims"],
            BuildPremiumClaimsChartRows(CoverageNames.Dental, "Current Billed Premium", "Prior Billed Premium"),
            "tblPPT_Dental_PremiumClaimsChart",
            SheetKind.Export,
            out nextRow);

        tableCount += AddTableAt(
            sheet,
            nextRow,
            "Schedule of Benefits",
            ["Slide Page", "Page Total", "Section Order", "Client Name", "Carrier", "Class Name", "Section Name", "Detail Text", "Source File", "Notes"],
            BuildScheduleOfBenefitsPptRows(data),
            "tblPPT_ScheduleOfBenefits",
            SheetKind.Export,
            out nextRow);

        sheet.SheetView.FreezeRows(2);
        return tableCount;
    }

    private static List<object?[]> BuildScheduleOfBenefitsRows(ClaimsExperienceDataModelData data)
    {
        return data.ScheduleOfBenefitsRows
            .OrderBy(row => row.SlidePage)
            .ThenBy(row => row.SectionOrder)
            .Select(row => new object?[]
            {
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
            })
            .ToList();
    }

    private static List<object?[]> BuildScheduleOfBenefitsPptRows(ClaimsExperienceDataModelData data)
    {
        return BuildScheduleOfBenefitsRows(data);
    }

    private static List<object?[]> BuildRenewalExperienceFinancialRows(ClaimsExperienceDataModelData data)
    {
        return data.RenewalExperienceFinancials
            .OrderBy(row => row.Start)
            .ThenBy(row => row.Benefit, StringComparer.OrdinalIgnoreCase)
            .Select(row => new object?[]
            {
                FormatPeriodLabel(row.Start, row.End),
                ToDateTime(row.Start),
                ToDateTime(row.End),
                InclusiveMonthCount(row.Start, row.End),
                row.Benefit,
                row.AverageCoveredEmployees,
                row.BilledPremium,
                row.AdjustedPremium,
                row.PaidClaims,
                row.IncurredClaims,
                row.AdjustedLossRatio / 100m,
                row.Weighting / 100m,
                row.PoolingPremium,
                row.SourceFile,
                row.SourceEvidence
            })
            .ToList();
    }

    private static List<object?[]> BuildClaimsExperienceSummaryPptRows()
    {
        return
        [
            ["{{ClientName}}", $"={ClientSetupValueFormula("Client Name", Quote(""))}"],
            ["{{ReportTitle}}", $"={ReportMonthCountFormula()}&\"-Month Claims Experience\""],
            ["{{CurrentPeriodLabel}}", $"={CurrentPeriodLabelFormula()}"],
            ["{{PriorPeriodLabel}}", $"={PriorPeriodLabelFormula()}"],
            ["{{RenewalDate}}", $"={ClientSetupValueFormula("Renewal Date", Quote(""))}"],
            ["{{EHC_Current_NetPremium}}", $"={CurrentCalculationFormula("Current Net Premium", CoverageNames.Ehc)}"],
            ["{{EHC_Current_PaidClaims}}", $"={CurrentCalculationFormula("Current Paid Claims", CoverageNames.Ehc)}"],
            ["{{EHC_Current_LossRatio}}", $"={CurrentCalculationFormula("Current Loss Ratio", CoverageNames.Ehc)}"],
            ["{{EHC_TargetLossRatio}}", $"={AssumptionValueFormula("EHC Break-even Point (TLR)", "0.763")}"],
            ["{{EHC_TrendFactor}}", $"={AssumptionValueFormula("EHC Trend Factor", "0.15")}"],
            ["{{EHC_ExpenseFactor}}", $"={AssumptionValueFormula("EHC Total Expense Factor", "0")}"],
            ["{{EHC_EstimatedRateAdjustment}}", $"={EstimatedAdjustmentFormula(CoverageNames.Ehc, "EHC Trend Factor", "EHC Total Expense Factor", "EHC Break-even Point (TLR)")}"],
            ["{{Dental_Current_BilledPremium}}", $"={CurrentCalculationFormula("Current Billed Premium", CoverageNames.Dental)}"],
            ["{{Dental_Current_PaidClaims}}", $"={CurrentCalculationFormula("Current Paid Claims", CoverageNames.Dental)}"],
            ["{{Dental_Current_LossRatio}}", $"={CurrentCalculationFormula("Current Loss Ratio", CoverageNames.Dental)}"],
            ["{{Dental_TargetLossRatio}}", $"={AssumptionValueFormula("Dental Break-even Point (TLR)", "0.763")}"],
            ["{{Dental_TrendFactor}}", $"={AssumptionValueFormula("Dental Trend Factor + Fee Guide", "0.15")}"],
            ["{{Dental_ExpenseFactor}}", $"={AssumptionValueFormula("Dental Total Expense Factor", "0")}"],
            ["{{Dental_EstimatedRateAdjustment}}", $"={EstimatedAdjustmentFormula(CoverageNames.Dental, "Dental Trend Factor + Fee Guide", "Dental Total Expense Factor", "Dental Break-even Point (TLR)")}"]
        ];
    }

    private static List<object?[]> BuildPremiumClaimsChartRows(string coverage, string currentPremiumColumn, string priorPremiumColumn)
    {
        return
        [
            [$"={PriorPeriodLabelFormula()}", $"={PriorCalculationFormula(priorPremiumColumn, coverage)}", $"={PriorCalculationFormula("Prior Paid Claims", coverage)}"],
            [$"={CurrentPeriodLabelFormula()}", $"={CurrentCalculationFormula(currentPremiumColumn, coverage)}", $"={CurrentCalculationFormula("Current Paid Claims", coverage)}"]
        ];
    }

    private static List<PremiumClaimsComparisonSelection> BuildPremiumClaimsComparisonSelections(
        ClaimsExperienceDataModelData data,
        string coverage)
    {
        if (!data.PeriodStart.HasValue || !data.PeriodEnd.HasValue)
        {
            return [];
        }

        var currentStart = data.PeriodStart.Value;
        var currentEnd = data.PeriodEnd.Value;
        var matchedPriorStart = currentStart.AddYears(-1);
        var matchedPriorEnd = currentEnd.AddYears(-1);
        var fullPrior = data.RenewalExperienceFinancials
            .Where(row => row.Benefit.Equals(coverage, StringComparison.OrdinalIgnoreCase))
            .Where(row => row.Start <= matchedPriorStart && row.End >= matchedPriorEnd)
            .Where(row => row.End < currentStart)
            .OrderByDescending(row => row.Start == matchedPriorStart)
            .ThenByDescending(row => row.End)
            .FirstOrDefault();
        if (fullPrior is null)
        {
            return [];
        }

        var selections = new List<PremiumClaimsComparisonSelection>();
        var fullMonthCount = InclusiveMonthCount(fullPrior.Start, fullPrior.End);
        var fullMatchesSelectedWindow = fullPrior.Start == matchedPriorStart && fullPrior.End == matchedPriorEnd;
        if (!fullMatchesSelectedWindow)
        {
            selections.Add(new PremiumClaimsComparisonSelection(
                1,
                "Prior Full",
                PriorPeriodId(fullMonthCount),
                FormatPeriodLabel(fullPrior.Start, fullPrior.End),
                fullPrior));
        }

        selections.Add(new PremiumClaimsComparisonSelection(
            selections.Count + 1,
            "Prior",
            data.PriorSourcePeriodId,
            FormatPeriodLabel(matchedPriorStart, matchedPriorEnd),
            fullMatchesSelectedWindow ? fullPrior : null));
        selections.Add(new PremiumClaimsComparisonSelection(
            selections.Count + 1,
            "Current",
            data.CurrentSourcePeriodId,
            FormatPeriodLabel(currentStart, currentEnd),
            null));
        return selections;
    }

    private static List<object?[]> BuildPremiumClaimsComparisonChartRows(
        ClaimsExperienceDataModelData data,
        int tableStartRow,
        string coverage,
        string premiumColumn)
    {
        var reportedSelections = BuildPremiumClaimsComparisonSelections(data, coverage);
        if (reportedSelections.Count > 0)
        {
            var reportedRows = new List<object?[]>();
            foreach (var selection in reportedSelections)
            {
                var excelRow = tableStartRow + 2 + reportedRows.Count;
                object premium = selection.ReportedFinancial?.BilledPremium is decimal reportedPremium
                    ? coverage.Equals(CoverageNames.Ehc, StringComparison.OrdinalIgnoreCase)
                        ? reportedPremium * (1m - data.RenewalTerms.EffectiveEhcPoolingCharge)
                        : reportedPremium
                    : $"={SumMonthlyFormula($"C{excelRow}", coverage, premiumColumn)}";
                object paidClaims = selection.ReportedFinancial?.PaidClaims is decimal reportedPaidClaims
                    ? reportedPaidClaims
                    : $"={CanonicalOrMonthlyPaidClaimsFormula(excelRow, coverage)}";
                reportedRows.Add(
                [
                    selection.SortOrder,
                    selection.Role,
                    selection.PeriodId,
                    selection.PeriodLabel,
                    premium,
                    paidClaims,
                    $"=IFERROR(F{excelRow}/E{excelRow},\\)",
                    selection.ReportedFinancial is null
                        ? ""
                        : $"Reported in {selection.ReportedFinancial.SourceFile}."
                ]);
            }

            return reportedRows;
        }

        var periods = BuildRankedSourcePeriodDefinitions(data, applyComparisonOverride: true)
            .OrderBy(period => period.SortOrder)
            .ToList();
        if (periods.Count == 0)
        {
            return BuildLegacyPremiumClaimsComparisonChartRows(tableStartRow, coverage, premiumColumn);
        }

        var rows = new List<object?[]>();
        foreach (var period in periods)
        {
            var excelRow = tableStartRow + 2 + rows.Count;
            rows.Add(
            [
                period.SortOrder,
                period.Role,
                $"={PeriodIdFormulaForRole(period.Role)}",
                $"={LookupFormula($"C{excelRow}", "tblPeriods[Period ID]", "tblPeriods[Period Label]", Quote(period.PeriodLabel))}",
                $"={SumMonthlyFormula($"C{excelRow}", coverage, premiumColumn)}",
                $"={CanonicalOrMonthlyPaidClaimsFormula(excelRow, coverage)}",
                $"=IFERROR(F{excelRow}/E{excelRow},\"\")",
                ""
            ]);
        }

        return rows;
    }

    private static List<object?[]> BuildLegacyPremiumClaimsComparisonChartRows(int tableStartRow, string coverage, string premiumColumn)
    {
        var rows = new List<object?[]>();
        var rowDefinitions = new[]
        {
            (SortOrder: 1, Role: "Benchmark 12 Months", PeriodIdFormula: BenchmarkPeriodIdFormula()),
            (SortOrder: 2, Role: "Prior Matched Period", PeriodIdFormula: PriorPeriodIdFormula()),
            (SortOrder: 3, Role: "Current Selected Period", PeriodIdFormula: CurrentPeriodIdFormula())
        };

        foreach (var definition in rowDefinitions)
        {
            var excelRow = tableStartRow + 2 + rows.Count;
            rows.Add(
            [
                definition.SortOrder,
                definition.Role,
                $"={definition.PeriodIdFormula}",
                $"={LookupFormula($"C{excelRow}", "tblPeriods[Period ID]", "tblPeriods[Period Label]", Quote(""))}",
                $"={SumMonthlyFormula($"C{excelRow}", coverage, premiumColumn)}",
                $"={SumMonthlyFormula($"C{excelRow}", coverage, "Paid Claims")}",
                $"=IFERROR(F{excelRow}/E{excelRow},\"\")",
                ""
            ]);
        }

        return rows;
    }

    private static List<object?[]> BuildLossRatioBreakEvenComparisonRows(
        ClaimsExperienceDataModelData data,
        int tableStartRow,
        string coverage,
        string premiumColumn,
        string targetLossRatioAssumption)
    {
        var reportedSelections = BuildPremiumClaimsComparisonSelections(data, coverage);
        if (reportedSelections.Count > 0)
        {
            var reportedRows = new List<object?[]>();
            foreach (var selection in reportedSelections)
            {
                var excelRow = tableStartRow + 2 + reportedRows.Count;
                object lossRatio;
                var reportedFinancial = selection.ReportedFinancial;
                if (reportedFinancial?.BilledPremium is > 0m &&
                    reportedFinancial.PaidClaims.HasValue)
                {
                    var premium = coverage.Equals(CoverageNames.Ehc, StringComparison.OrdinalIgnoreCase)
                        ? reportedFinancial.BilledPremium.Value * (1m - data.RenewalTerms.EffectiveEhcPoolingCharge)
                        : reportedFinancial.BilledPremium.Value;
                    lossRatio = reportedFinancial.PaidClaims.Value /
                                premium;
                }
                else
                {
                    var paidClaims = CanonicalOrMonthlyPaidClaimsFormula(excelRow, coverage);
                    var premium = SumMonthlyFormula($"C{excelRow}", coverage, premiumColumn);
                    lossRatio = $"=IFERROR(({paidClaims})/({premium}),\\)";
                }

                reportedRows.Add(
                [
                    selection.SortOrder,
                    selection.Role,
                    selection.PeriodId,
                    selection.PeriodLabel,
                    lossRatio,
                    $"={AssumptionValueFormula(targetLossRatioAssumption, "0.763")}",
                    selection.ReportedFinancial is null
                        ? ""
                        : $"Reported in {selection.ReportedFinancial.SourceFile}."
                ]);
            }

            return reportedRows;
        }

        var periods = BuildRankedSourcePeriodDefinitions(data, applyComparisonOverride: true)
            .OrderBy(period => period.SortOrder)
            .ToList();
        if (periods.Count == 0)
        {
            return BuildLegacyLossRatioBreakEvenComparisonRows(tableStartRow, coverage, premiumColumn, targetLossRatioAssumption);
        }

        var rows = new List<object?[]>();
        foreach (var period in periods)
        {
            var excelRow = tableStartRow + 2 + rows.Count;
            var paidClaims = CanonicalOrMonthlyPaidClaimsFormula(excelRow, coverage);
            var premium = SumMonthlyFormula($"C{excelRow}", coverage, premiumColumn);
            rows.Add(
            [
                period.SortOrder,
                period.Role,
                $"={PeriodIdFormulaForRole(period.Role)}",
                $"={LookupFormula($"C{excelRow}", "tblPeriods[Period ID]", "tblPeriods[Period Label]", Quote(period.PeriodLabel))}",
                $"=IFERROR(({paidClaims})/({premium}),\"\")",
                $"={AssumptionValueFormula(targetLossRatioAssumption, "0.763")}",
                ""
            ]);
        }

        return rows;
    }

    private static List<object?[]> BuildLegacyLossRatioBreakEvenComparisonRows(int tableStartRow, string coverage, string premiumColumn, string targetLossRatioAssumption)
    {
        var rows = new List<object?[]>();
        var rowDefinitions = new[]
        {
            (SortOrder: 1, Role: "Benchmark 12 Months", PeriodIdFormula: BenchmarkPeriodIdFormula()),
            (SortOrder: 2, Role: "Prior Matched Period", PeriodIdFormula: PriorPeriodIdFormula()),
            (SortOrder: 3, Role: "Current Selected Period", PeriodIdFormula: CurrentPeriodIdFormula())
        };

        foreach (var definition in rowDefinitions)
        {
            var excelRow = tableStartRow + 2 + rows.Count;
            var paidClaims = SumMonthlyFormula($"C{excelRow}", coverage, "Paid Claims");
            var premium = SumMonthlyFormula($"C{excelRow}", coverage, premiumColumn);
            rows.Add(
            [
                definition.SortOrder,
                definition.Role,
                $"={definition.PeriodIdFormula}",
                $"={LookupFormula($"C{excelRow}", "tblPeriods[Period ID]", "tblPeriods[Period Label]", Quote(""))}",
                $"=IFERROR(({paidClaims})/({premium}),\"\")",
                $"={AssumptionValueFormula(targetLossRatioAssumption, "0.763")}",
                ""
            ]);
        }

        return rows;
    }
    private static void AddChartsSheet(IXLWorksheet sheet, ClaimsExperienceDataModelData data)
    {
        FormatTitle(sheet, "Claims Experience Charts", SheetKind.Export);
        sheet.Cell(3, 1).Value = "Monthly Chart Data";
        sheet.Cell(3, 1).Style.Font.Bold = true;

        var monthlyHeaders = new[]
        {
            "Month",
            "Current Period Paid Claims",
            "Current Period No. of Employees",
            "EHC Paid Claims",
            "Dental Paid Claims"
        };

        for (var index = 0; index < monthlyHeaders.Length; index++)
        {
            sheet.Cell(4, index + 1).Value = monthlyHeaders[index];
        }

        HeaderRange(sheet.Range(4, 1, 4, monthlyHeaders.Length), SheetKind.Export);

        var selectedCurrentMonths = SelectExperienceMonths(data.Months, data.ReportMonthCount);
        var monthlyRows = selectedCurrentMonths.Count > 0 ? selectedCurrentMonths : [new ClaimsExperienceMonth()];
        for (var index = 0; index < monthlyRows.Count; index++)
        {
            var row = 5 + index;
            var month = monthlyRows[index];
            sheet.Cell(row, 1).Value = month.Month == default ? "" : month.Month.ToDateTime(TimeOnly.MinValue);
            sheet.Cell(row, 1).Style.DateFormat.Format = "mmmm";
            sheet.Cell(row, 2).Value = month.EhcClaims + month.DentalClaims;
            sheet.Cell(row, 3).FormulaA1 = MonthlyLivesForChartFormula(row, CoverageNames.Ehc);
            sheet.Cell(row, 4).Value = month.EhcClaims;
            sheet.Cell(row, 5).Value = month.DentalClaims;
        }

        sheet.Range(5, 2, 4 + monthlyRows.Count, 2).Style.NumberFormat.Format = "$#,##0";
        sheet.Range(5, 3, 4 + monthlyRows.Count, 3).Style.NumberFormat.Format = "#,##0";
        sheet.Range(5, 4, 4 + monthlyRows.Count, 5).Style.NumberFormat.Format = "$#,##0";
        sheet.Column(1).Width = 14;
        sheet.Columns(2, 5).Width = 22;

        var currentEhcCategories = data.EhcCategories.Where(item => item.PeriodId.Equals(data.CurrentSourcePeriodId, StringComparison.OrdinalIgnoreCase)).ToList();
        var currentDentalCategories = data.DentalCategories.Where(item => item.PeriodId.Equals(data.CurrentSourcePeriodId, StringComparison.OrdinalIgnoreCase)).ToList();
        var currentParamedicalDetails = data.ParamedicalDetails.Where(item => item.PeriodId.Equals(data.CurrentSourcePeriodId, StringComparison.OrdinalIgnoreCase)).ToList();

        var nextRow = 5 + monthlyRows.Count + 3;
        nextRow = AddChartDistributionData(sheet, nextRow, "EHC Distribution Data", BuildChartDistributionRows(currentEhcCategories.Select(item => (item.Category, item.PaidClaims))));
        nextRow = AddChartDistributionData(sheet, nextRow, "Dental Distribution Data", BuildChartDistributionRows(currentDentalCategories.Select(item => (item.Category, item.PaidClaims))));
        nextRow = AddChartDistributionData(sheet, nextRow, "Paramedical Distribution Data", BuildChartDistributionRows(currentParamedicalDetails.Select(item => (item.Service, item.PaidClaims))));
        AddChartDistributionData(
            sheet,
            nextRow,
            "Combined Distribution Data",
            BuildChartDistributionRows(
            [
                (CoverageNames.Ehc, currentEhcCategories.Sum(item => item.PaidClaims)),
                (CoverageNames.Dental, currentDentalCategories.Sum(item => item.PaidClaims))
            ]));

        sheet.RangeUsed()?.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
    }

    private static int AddChartDistributionData(IXLWorksheet sheet, int titleRow, string title, IReadOnlyList<ChartCategoryValue> rows)
    {
        sheet.Cell(titleRow, 1).Value = title;
        sheet.Cell(titleRow, 1).Style.Font.Bold = true;
        sheet.Cell(titleRow + 1, 1).Value = "Category";
        sheet.Cell(titleRow + 1, 2).Value = "Paid Claims";
        sheet.Cell(titleRow + 1, 3).Value = "Percent of Claims";
        HeaderRange(sheet.Range(titleRow + 1, 1, titleRow + 1, 3), SheetKind.Export);

        var dataRows = rows.Count > 0 ? rows : [new ChartCategoryValue("No data", 0m)];
        var dataStart = titleRow + 2;
        var dataEnd = dataStart + dataRows.Count - 1;
        for (var index = 0; index < dataRows.Count; index++)
        {
            var row = dataStart + index;
            sheet.Cell(row, 1).Value = dataRows[index].Category;
            sheet.Cell(row, 2).Value = dataRows[index].Value;
            sheet.Cell(row, 3).FormulaA1 = $"IFERROR(B{row}/SUM($B${dataStart}:$B${dataEnd}),\"\")";
        }

        sheet.Range(dataStart, 2, dataEnd, 2).Style.NumberFormat.Format = "$#,##0";
        sheet.Range(dataStart, 3, dataEnd, 3).Style.NumberFormat.Format = "0.0%";
        return dataEnd + 3;
    }

    private static List<ChartCategoryValue> BuildChartDistributionRows(IEnumerable<(string Category, decimal PaidClaims)> rows)
    {
        return rows
            .Where(item => !string.IsNullOrWhiteSpace(item.Category))
            .GroupBy(item => item.Category, StringComparer.OrdinalIgnoreCase)
            .Select(group => new ChartCategoryValue(group.Key, group.Sum(item => item.PaidClaims)))
            .Where(item => item.Value != 0m)
            .OrderByDescending(item => item.Value)
            .ToList();
    }

    private static void AddClaimsExperienceSummarySheet(IXLWorksheet sheet)
    {
        const string titleColor = "#34495E";
        var ehcColor = XLColor.FromHtml("#7AC943");
        var dentalColor = XLColor.FromHtml("#A6A1A1");

        sheet.Style.Font.FontName = "Calibri";
        sheet.Style.Font.FontSize = 8;
        sheet.Style.Fill.BackgroundColor = XLColor.White;
        sheet.ShowGridLines = false;
        sheet.Range("B1:L49").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        sheet.Range("B1:L49").Style.Font.FontColor = XLColor.FromHtml("#1F2D3D");

        sheet.Cell("B2").FormulaA1 = ClientSetupValueFormula("Client Name", Quote("Client"));
        sheet.Range("B2:L2").Merge();
        sheet.Cell("B2").Style.Font.Bold = true;
        sheet.Cell("B2").Style.Font.FontSize = 18;
        sheet.Cell("B2").Style.Font.FontColor = XLColor.FromHtml(titleColor);
        sheet.Cell("B2").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

        sheet.Cell("B3").FormulaA1 = $"{ReportMonthCountFormula()}&\"-Month Claims Experience\"";
        sheet.Range("B3:L3").Merge();
        sheet.Cell("B3").Style.Font.Bold = true;
        sheet.Cell("B3").Style.Font.FontSize = 12;
        sheet.Cell("B3").Style.Font.FontColor = XLColor.FromHtml(titleColor);
        sheet.Cell("B3").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

        AddClaimsExperienceSummarySection(
            sheet,
            5,
            CoverageNames.Ehc,
            "EXTENDED HEALTH CARE:",
            "Net Premium",
            "Net Loss Ratio",
            "Net loss ratio",
            "Inflation/utilization factor",
            "EHC Trend Factor",
            "EHC Total Expense Factor",
            "EHC Break-even Point (TLR)",
            "health trend",
            "net loss ratio",
            ehcColor,
            6);

        sheet.Cell("F17").FormulaA1 = $"\"The health trend is currently at \"&TEXT({AssumptionValueFormula("EHC Trend Factor", "0.15")},\"0.00%\")&\". Your net loss ratio combined with the health trend will determine your renewal rates.\"";
        sheet.Range("F17:L18").Merge();
        sheet.Cell("F17").Style.Alignment.WrapText = true;
        sheet.Cell("F20").Value = "The total paid claims may include in-province extended health care claims which will be pooled at renewal if they exceed the per-certificate threshold.";
        sheet.Range("F20:L21").Merge();
        sheet.Cell("F20").Style.Alignment.WrapText = true;

        AddClaimsExperienceSummarySection(
            sheet,
            29,
            CoverageNames.Dental,
            "DENTAL CARE:",
            "Billed Premium",
            "Paid Loss Ratio",
            "Paid loss ratio",
            "Trend factor + Fee guide",
            "Dental Trend Factor + Fee Guide",
            "Dental Total Expense Factor",
            "Dental Break-even Point (TLR)",
            "dental trend",
            "paid loss ratio",
            dentalColor,
            30);

        sheet.Cell("F41").FormulaA1 = $"\"The dental trend and fee guide is currently at \"&TEXT({AssumptionValueFormula("Dental Trend Factor + Fee Guide", "0.15")},\"0.00%\")&\". Your paid loss ratio combined with the dental trend will determine your renewal rates.\"";
        sheet.Range("F41:L43").Merge();
        sheet.Cell("F41").Style.Alignment.WrapText = true;

        sheet.Cell("B46").Value = "NEXT RENEWAL:";
        sheet.Cell("B46").Style.Font.Bold = true;
        sheet.Cell("B47").FormulaA1 = $"IF({ClientSetupValueFormula("Renewal Date", Quote(""))}=\"\",\"Your next renewal is scheduled for [Renewal Month, Year].\",\"Your next renewal is scheduled for \"&TEXT({ClientSetupValueFormula("Renewal Date", Quote(""))},\"mmmm, yyyy\")&\".\")";
        sheet.Range("B47:L47").Merge();

        sheet.Cell("B49").FormulaA1 = "TODAY()";
        sheet.Cell("B49").Style.DateFormat.Format = "d-mmm-yy";
        sheet.Cell("L49").Value = "E&OE";
        sheet.Cell("L49").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

        sheet.Columns("A:A").Width = 2;
        sheet.Columns("B:B").Width = 12;
        sheet.Columns("C:C").Width = 12;
        sheet.Columns("D:D").Width = 12;
        sheet.Columns("E:E").Width = 12;
        sheet.Columns("F:F").Width = 17;
        sheet.Columns("G:G").Width = 2;
        sheet.Columns("H:H").Width = 2;
        sheet.Columns("I:I").Width = 2;
        sheet.Columns("J:J").Width = 2;
        sheet.Columns("K:K").Width = 2;
        sheet.Columns("L:L").Width = 13;
        sheet.Columns("N:P").Width = 16;
        for (var column = 14; column <= 16; column++)
        {
            sheet.Column(column).Hide();
        }

        for (var row = 1; row <= 49; row++)
        {
            sheet.Row(row).Height = row is 2 or 3 ? 18 : 13.5;
        }

        sheet.Row(6).Height = 19.5;
        sheet.Row(30).Height = 19.5;
        sheet.Range("B1:L49").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        sheet.Range("F10:L21").Style.Alignment.WrapText = true;
        sheet.Range("F34:L43").Style.Alignment.WrapText = true;
        sheet.Cell("F16").Style.Alignment.WrapText = false;
        sheet.Cell("F40").Style.Alignment.WrapText = false;
        sheet.Range("B1:L49").Style.Border.InsideBorder = XLBorderStyleValues.None;
        sheet.Range("B1:L49").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        sheet.PageSetup.PrintAreas.Clear();
        sheet.PageSetup.PrintAreas.Add("B1:L49");
        sheet.PageSetup.PageOrientation = XLPageOrientation.Portrait;
        sheet.PageSetup.FitToPages(1, 1);
        sheet.PageSetup.Margins.Top = 0.25;
        sheet.PageSetup.Margins.Bottom = 0.25;
        sheet.PageSetup.Margins.Left = 0.25;
        sheet.PageSetup.Margins.Right = 0.25;
        sheet.SheetView.FreezeRows(4);
    }

    private static void AddClaimsExperienceSummarySection(
        IXLWorksheet sheet,
        int startRow,
        string coverage,
        string sectionTitle,
        string premiumLabel,
        string lossRatioLabel,
        string ratioLabel,
        string trendLabel,
        string trendAssumption,
        string expenseAssumption,
        string targetLossRatioAssumption,
        string trendNarrativeLabel,
        string ratioNarrativeLabel,
        XLColor headerColor,
        int chartDataHeaderRow)
    {
        var headerRow = startRow + 1;
        var dataRow = startRow + 2;
        var estimateHeaderRow = startRow + 5;
        var estimateStartRow = startRow + 6;
        var adjustmentRow = startRow + 11;
        var currentPremiumColumn = premiumLabel.Contains("Billed", StringComparison.OrdinalIgnoreCase)
            ? "Current Billed Premium"
            : "Current Net Premium";

        sheet.Cell(startRow, 2).Value = sectionTitle;
        sheet.Range(startRow, 2, startRow, 5).Merge();
        sheet.Cell(startRow, 2).Style.Font.Bold = true;
        sheet.Cell(startRow, 2).Style.Font.FontColor = headerColor;

        sheet.Cell(headerRow, 2).Value = "Current Period";
        sheet.Cell(headerRow, 3).Value = premiumLabel;
        sheet.Cell(headerRow, 4).Value = "Paid Claims";
        sheet.Cell(headerRow, 5).Value = lossRatioLabel;
        var header = sheet.Range(headerRow, 2, headerRow, 12);
        header.Style.Fill.BackgroundColor = headerColor;
        header.Style.Font.FontColor = XLColor.White;
        header.Style.Font.Bold = true;
        header.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

        sheet.Cell(dataRow, 2).FormulaA1 = CurrentPeriodLabelFormula();
        sheet.Cell(dataRow, 3).FormulaA1 = CurrentCalculationFormula(currentPremiumColumn, coverage);
        sheet.Cell(dataRow, 4).FormulaA1 = CurrentCalculationFormula("Current Paid Claims", coverage);
        sheet.Cell(dataRow, 5).FormulaA1 = $"IFERROR(D{dataRow}/C{dataRow},0)";
        sheet.Range(dataRow, 3, dataRow, 4).Style.NumberFormat.Format = "$#,##0";
        sheet.Cell(dataRow, 5).Style.NumberFormat.Format = "0.00%";
        sheet.Range(headerRow, 2, dataRow, 5).Style.Border.OutsideBorder = XLBorderStyleValues.None;
        sheet.Range(headerRow, 2, dataRow, 5).Style.Border.InsideBorder = XLBorderStyleValues.None;

        sheet.Cell(estimateHeaderRow, 10).FormulaA1 = $"\"Est. \"&IF({ClientSetupValueFormula("Renewal Date", Quote(""))}=\"\",YEAR(TODAY()),YEAR({ClientSetupValueFormula("Renewal Date", Quote(""))}))&\" Renewal\"";
        sheet.Range(estimateHeaderRow, 10, estimateHeaderRow, 12).Merge();
        sheet.Cell(estimateHeaderRow, 10).Style.Font.Bold = true;
        sheet.Cell(estimateHeaderRow, 10).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

        sheet.Cell(estimateStartRow, 6).Value = ratioLabel;
        sheet.Cell(estimateStartRow, 12).FormulaA1 = $"E{dataRow}";
        sheet.Cell(estimateStartRow + 1, 6).Value = trendLabel;
        sheet.Cell(estimateStartRow + 1, 12).FormulaA1 = AssumptionValueFormula(trendAssumption, "0.15");
        sheet.Cell(estimateStartRow + 2, 6).Value = "Total expense factors";
        sheet.Cell(estimateStartRow + 2, 12).FormulaA1 = AssumptionValueFormula(expenseAssumption, "0");
        sheet.Cell(estimateStartRow + 3, 6).Value = "Break-even point (TLR)";
        sheet.Cell(estimateStartRow + 3, 12).FormulaA1 = AssumptionValueFormula(targetLossRatioAssumption, "0.763");
        sheet.Range(estimateStartRow, 12, estimateStartRow + 3, 12).Style.NumberFormat.Format = "0.00%";
        sheet.Range(estimateStartRow, 6, estimateStartRow + 3, 12).Style.Border.BottomBorder = XLBorderStyleValues.None;

        sheet.Cell(adjustmentRow, 6).Value = "Estimated rate adjustment";
        sheet.Cell(adjustmentRow, 6).Style.Alignment.WrapText = false;
        sheet.Cell(adjustmentRow, 6).Style.Font.Bold = true;
        sheet.Cell(adjustmentRow, 6).Style.Font.FontColor = XLColor.FromHtml("#000080");
        sheet.Cell(adjustmentRow, 12).FormulaA1 = $"IFERROR((L{estimateStartRow}*(1+L{estimateStartRow + 1}+L{estimateStartRow + 2})/L{estimateStartRow + 3})-1,0)";
        sheet.Cell(adjustmentRow, 12).Style.NumberFormat.Format = "0.00%";
        sheet.Cell(adjustmentRow, 12).Style.Font.Bold = true;
        sheet.Range(adjustmentRow, 12, adjustmentRow, 12).AddConditionalFormat().WhenLessThan(0).Font.SetFontColor(XLColor.FromHtml("#00B050"));
        sheet.Range(adjustmentRow, 12, adjustmentRow, 12).AddConditionalFormat().WhenGreaterThan(0).Font.SetFontColor(XLColor.FromHtml("#C00000"));

        sheet.Cell(chartDataHeaderRow, 14).Value = $"{coverage} Chart Period";
        sheet.Cell(chartDataHeaderRow, 15).Value = "Net Premium";
        sheet.Cell(chartDataHeaderRow, 16).Value = "Paid Claims";
        HeaderRange(sheet.Range(chartDataHeaderRow, 14, chartDataHeaderRow, 16), SheetKind.Export);
        sheet.Cell(chartDataHeaderRow + 1, 14).FormulaA1 = PriorPeriodLabelFormula();
        sheet.Cell(chartDataHeaderRow + 1, 15).FormulaA1 = PriorCalculationFormula("Prior Net Premium", coverage);
        sheet.Cell(chartDataHeaderRow + 1, 16).FormulaA1 = PriorCalculationFormula("Prior Paid Claims", coverage);
        sheet.Cell(chartDataHeaderRow + 2, 14).FormulaA1 = CurrentPeriodLabelFormula();
        sheet.Cell(chartDataHeaderRow + 2, 15).FormulaA1 = CurrentCalculationFormula("Current Net Premium", coverage);
        sheet.Cell(chartDataHeaderRow + 2, 16).FormulaA1 = CurrentCalculationFormula("Current Paid Claims", coverage);
        sheet.Range(chartDataHeaderRow + 1, 15, chartDataHeaderRow + 2, 16).Style.NumberFormat.Format = "$#,##0";
    }

    private static void AddClaimsExperienceCharts(string workbookPath, ClaimsExperienceDataModelData data)
    {
        using var document = SpreadsheetDocument.Open(workbookPath, true);
        var workbookPart = document.WorkbookPart;
        var sheet = workbookPart?.Workbook.Sheets?.Elements<S.Sheet>()
            .FirstOrDefault(item => string.Equals(item.Name?.Value, "12_Charts", StringComparison.OrdinalIgnoreCase));
        if (workbookPart is null || sheet?.Id is null)
        {
            return;
        }

        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
        var drawingsPart = EnsureDrawingsPart(worksheetPart);
        var selectedCurrentMonths = SelectExperienceMonths(data.Months, data.ReportMonthCount);
        var monthlyRowCount = Math.Max(selectedCurrentMonths.Count, 1);
        var monthlyEndRow = 4 + monthlyRowCount;

        var monthlyLabels = selectedCurrentMonths.Count > 0
            ? selectedCurrentMonths.Select(item => item.Month.ToString("MMMM", Culture)).ToList()
            : [""];
        var combinedClaims = selectedCurrentMonths.Count > 0
            ? selectedCurrentMonths.Select(item => (decimal?)(item.EhcClaims + item.DentalClaims)).ToList()
            : [(decimal?)0m];
        var employeeCounts = selectedCurrentMonths.Count > 0
            ? selectedCurrentMonths.Select(item =>
                data.BillingMonthlyLivesByMonth.TryGetValue(FirstDayOfMonth(item.Month), out var lives) && lives.HealthInsuredLives.HasValue
                    ? (decimal?)lives.HealthInsuredLives.Value
                    : null).ToList()
            : [(decimal?)null];
        var ehcClaims = selectedCurrentMonths.Count > 0
            ? selectedCurrentMonths.Select(item => (decimal?)item.EhcClaims).ToList()
            : [(decimal?)0m];
        var dentalClaims = selectedCurrentMonths.Count > 0
            ? selectedCurrentMonths.Select(item => (decimal?)item.DentalClaims).ToList()
            : [(decimal?)0m];

        AddChart(
            drawingsPart,
            "Current Year Paid Claims by Month - Employee by Month",
            BuildMonthlyComboChartSpace("12_Charts", "Current Year Paid Claims by Month\nEmployee by Month", "B", "C", monthlyEndRow, monthlyLabels, combinedClaims, employeeCounts),
            7,
            2,
            20,
            24);

        AddChart(
            drawingsPart,
            "EHC Paid Claims by Month - Employee by Month",
            BuildMonthlyComboChartSpace("12_Charts", "EHC Paid Claims by Month\nEmployee by Month", "D", "C", monthlyEndRow, monthlyLabels, ehcClaims, employeeCounts),
            7,
            26,
            20,
            48);

        AddChart(
            drawingsPart,
            "Dental Paid Claims by Month - Employee by Month",
            BuildMonthlyComboChartSpace("12_Charts", "Dental Paid Claims by Month\nEmployee by Month", "E", "C", monthlyEndRow, monthlyLabels, dentalClaims, employeeCounts),
            7,
            50,
            20,
            72);

        var currentEhcCategories = data.EhcCategories.Where(item => item.PeriodId.Equals(data.CurrentSourcePeriodId, StringComparison.OrdinalIgnoreCase)).ToList();
        var currentDentalCategories = data.DentalCategories.Where(item => item.PeriodId.Equals(data.CurrentSourcePeriodId, StringComparison.OrdinalIgnoreCase)).ToList();
        var currentParamedicalDetails = data.ParamedicalDetails.Where(item => item.PeriodId.Equals(data.CurrentSourcePeriodId, StringComparison.OrdinalIgnoreCase)).ToList();
        var ehcRows = BuildChartDistributionRows(currentEhcCategories.Select(item => (item.Category, item.PaidClaims)));
        var dentalRows = BuildChartDistributionRows(currentDentalCategories.Select(item => (item.Category, item.PaidClaims)));
        var paramedicalRows = BuildChartDistributionRows(currentParamedicalDetails.Select(item => (item.Service, item.PaidClaims)));
        var combinedRows = BuildChartDistributionRows(
        [
            (CoverageNames.Ehc, currentEhcCategories.Sum(item => item.PaidClaims)),
            (CoverageNames.Dental, currentDentalCategories.Sum(item => item.PaidClaims))
        ]);

        var ehcTitleRow = 5 + monthlyRowCount + 3;
        var dentalTitleRow = ehcTitleRow + Math.Max(ehcRows.Count, 1) + 5;
        var paramedicalTitleRow = dentalTitleRow + Math.Max(dentalRows.Count, 1) + 5;
        var combinedTitleRow = paramedicalTitleRow + Math.Max(paramedicalRows.Count, 1) + 5;

        AddPieChart(drawingsPart, "12_Charts", "EHC Claims Distribution", ehcRows, ehcTitleRow, 1, 74, 7, 95);
        AddPieChart(drawingsPart, "12_Charts", "Dental Claims Distribution", dentalRows, dentalTitleRow, 8, 74, 14, 95);
        AddPieChart(drawingsPart, "12_Charts", "Paramedical Claims Distribution", paramedicalRows, paramedicalTitleRow, 15, 74, 21, 95);
        AddPieChart(drawingsPart, "12_Charts", "Combined Claims Distribution", combinedRows, combinedTitleRow, 1, 97, 7, 118);

        drawingsPart.WorksheetDrawing?.Save();
        worksheetPart.Worksheet.Save();

        AddClaimsExperienceSummaryCharts(workbookPart, data);
    }

    private static void AddClaimsExperienceSummaryCharts(WorkbookPart workbookPart, ClaimsExperienceDataModelData data)
    {
        var sheet = workbookPart.Workbook.Sheets?.Elements<S.Sheet>()
            .FirstOrDefault(item => string.Equals(item.Name?.Value, "ClaimsExperience_Summary", StringComparison.OrdinalIgnoreCase));
        if (sheet?.Id is null)
        {
            return;
        }

        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
        var drawingsPart = EnsureDrawingsPart(worksheetPart);
        var currentLabel = SummaryCurrentPeriodLabel(data);
        var priorLabel = SummaryPriorPeriodLabel(data);
        var currentMonths = SelectExperienceMonths(data.Months, data.ReportMonthCount);
        var priorMonths = SelectExperienceMonths(data.PriorMonths, data.ReportMonthCount);
        var ehcNetPremiumFactor = 1m - data.RenewalTerms.EffectiveEhcPoolingCharge;
        var ehcCurrentPremium = currentMonths.Sum(item => item.EhcPremium) * ehcNetPremiumFactor;
        var ehcCurrentPaidClaims = currentMonths.Sum(item => item.EhcClaims);
        var ehcPriorPremium = priorMonths.Sum(item => item.EhcPremium) * ehcNetPremiumFactor;
        var ehcPriorPaidClaims = priorMonths.Sum(item => item.EhcClaims);
        var dentalCurrentPremium = currentMonths.Sum(item => item.DentalPremium);
        var dentalCurrentPaidClaims = currentMonths.Sum(item => item.DentalClaims);
        var dentalPriorPremium = priorMonths.Sum(item => item.DentalPremium);
        var dentalPriorPaidClaims = priorMonths.Sum(item => item.DentalClaims);

        AddChart(
            drawingsPart,
            "EHC Net Premium vs Paid Claims",
            BuildPremiumPaidClaimsChartSpace(
                "ClaimsExperience_Summary",
                7,
                8,
                [priorLabel, currentLabel],
                [ehcPriorPremium, ehcCurrentPremium],
                [ehcPriorPaidClaims, ehcCurrentPaidClaims]),
            1,
            9,
            5,
            21);

        AddChart(
            drawingsPart,
            "Dental Billed Premium vs Paid Claims",
            BuildPremiumPaidClaimsChartSpace(
                "ClaimsExperience_Summary",
                31,
                32,
                [priorLabel, currentLabel],
                [dentalPriorPremium, dentalCurrentPremium],
                [dentalPriorPaidClaims, dentalCurrentPaidClaims]),
            1,
            33,
            5,
            45);

        drawingsPart.WorksheetDrawing?.Save();
        worksheetPart.Worksheet.Save();
    }

    private static string SummaryCurrentPeriodLabel(ClaimsExperienceDataModelData data)
    {
        var months = SelectExperienceMonths(data.Months, data.ReportMonthCount);
        if (months.Count == 0 && !string.IsNullOrWhiteSpace(data.PeriodLabel))
        {
            return data.PeriodLabel;
        }

        return months.Count == 0
            ? "Current Period"
            : FormatPeriodLabel(months.Min(item => item.Month), months.Max(item => item.Month));
    }

    private static string SummaryPriorPeriodLabel(ClaimsExperienceDataModelData data)
    {
        var priorMonths = SelectExperienceMonths(data.PriorMonths, data.ReportMonthCount);
        if (priorMonths.Count > 0)
        {
            return FormatPeriodLabel(priorMonths.Min(item => item.Month), priorMonths.Max(item => item.Month));
        }

        var months = SelectExperienceMonths(data.Months, data.ReportMonthCount);
        if (months.Count > 0)
        {
            return FormatPeriodLabel(months.Min(item => item.Month).AddYears(-1), months.Max(item => item.Month).AddYears(-1));
        }

        return "Prior Period";
    }

    private static DrawingsPart EnsureDrawingsPart(WorksheetPart worksheetPart)
    {
        var drawingsPart = worksheetPart.DrawingsPart;
        if (drawingsPart is null)
        {
            drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
            drawingsPart.WorksheetDrawing = new Xdr.WorksheetDrawing();
            InsertWorksheetDrawing(worksheetPart, worksheetPart.GetIdOfPart(drawingsPart));
        }

        drawingsPart.WorksheetDrawing ??= new Xdr.WorksheetDrawing();
        return drawingsPart;
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

    private static void AddChart(DrawingsPart drawingsPart, string chartName, C.ChartSpace chartSpace, int fromColumn, int fromRow, int toColumn, int toRow)
    {
        var chartPart = drawingsPart.AddNewPart<ChartPart>();
        chartPart.ChartSpace = chartSpace;
        chartPart.ChartSpace.Save();

        var chartRelId = drawingsPart.GetIdOfPart(chartPart);
        var drawingId = drawingsPart.WorksheetDrawing!
            .Descendants<Xdr.NonVisualDrawingProperties>()
            .Select(item => item.Id?.Value ?? 0u)
            .DefaultIfEmpty(1u)
            .Max() + 1u;

        drawingsPart.WorksheetDrawing.Append(BuildChartAnchor(drawingId, chartRelId, chartName, fromColumn, fromRow, toColumn, toRow));
    }

    private static void AddPieChart(
        DrawingsPart drawingsPart,
        string sheetName,
        string title,
        IReadOnlyList<ChartCategoryValue> rows,
        int titleRow,
        int fromColumn,
        int fromRow,
        int toColumn,
        int toRow)
    {
        var dataRows = rows.Count > 0 ? rows : [new ChartCategoryValue("No data", 0m)];
        var dataStart = titleRow + 2;
        var dataEnd = dataStart + dataRows.Count - 1;
        AddChart(
            drawingsPart,
            title,
            BuildPieChartSpace(sheetName, title, dataStart, dataEnd, dataRows),
            fromColumn,
            fromRow,
            toColumn,
            toRow);
    }

    private static C.ChartSpace BuildPremiumPaidClaimsChartSpace(
        string sheetName,
        int dataStartRow,
        int dataEndRow,
        IReadOnlyList<string> categories,
        IReadOnlyList<decimal> premiumValues,
        IReadOnlyList<decimal> paidClaimValues)
    {
        const uint categoryAxisId = 64650112u;
        const uint valueAxisId = 64672768u;
        var sheetReference = SheetReference(sheetName);
        var headerRow = dataStartRow - 1;
        var premiumSeriesValues = premiumValues.Select(value => (decimal?)value).ToList();
        var paidClaimSeriesValues = paidClaimValues.Select(value => (decimal?)value).ToList();

        var chartSpace = new C.ChartSpace(
            new C.EditingLanguage { Val = "en-US" },
            new C.Chart(
                new C.PlotArea(
                    new C.Layout(),
                    new C.BarChart(
                        new C.BarDirection { Val = C.BarDirectionValues.Column },
                        new C.BarGrouping { Val = C.BarGroupingValues.Clustered },
                        BuildBarSeries(
                            0u,
                            $"{sheetReference}!$C${headerRow}",
                            $"{sheetReference}!$N${dataStartRow}:$N${dataEndRow}",
                            $"{sheetReference}!$O${dataStartRow}:$O${dataEndRow}",
                            "Net Premium",
                            categories,
                            premiumSeriesValues,
                            Style1NetPremiumColor,
                            "$#,##0"),
                        BuildBarSeries(
                            1u,
                            $"{sheetReference}!$D${headerRow}",
                            $"{sheetReference}!$N${dataStartRow}:$N${dataEndRow}",
                            $"{sheetReference}!$P${dataStartRow}:$P${dataEndRow}",
                            "Paid Claims",
                            categories,
                            paidClaimSeriesValues,
                            Style1PaidClaimsColor,
                            "$#,##0"),
                        BuildSummaryChartDataLabels(),
                        new C.AxisId { Val = categoryAxisId },
                        new C.AxisId { Val = valueAxisId }),
                    BuildCategoryAxis(categoryAxisId, valueAxisId),
                    BuildValueAxis(valueAxisId, categoryAxisId, C.AxisPositionValues.Left, "$#,##0", C.CrossesValues.AutoZero, Style1GridlineColor)),
                new C.Legend(new C.LegendPosition { Val = C.LegendPositionValues.Bottom }, new C.Layout()),
                new C.PlotVisibleOnly { Val = false }));

        ApplyExcelStyle1Chart(chartSpace);
        return chartSpace;
    }

    private static C.ChartSpace BuildMonthlyComboChartSpace(
        string sheetName,
        string title,
        string claimsColumn,
        string employeeColumn,
        int dataEndRow,
        IReadOnlyList<string> categories,
        IReadOnlyList<decimal?> paidClaims,
        IReadOnlyList<decimal?> employeeCounts)
    {
        const uint categoryAxisId = 54650112u;
        const uint claimsAxisId = 54672768u;
        const uint employeeAxisId = 54693024u;
        var sheetReference = SheetReference(sheetName);

        return new C.ChartSpace(
            new C.EditingLanguage { Val = "en-US" },
            new C.Chart(
                BuildChartTitle(title),
                new C.PlotArea(
                    new C.Layout(),
                    new C.BarChart(
                        new C.BarDirection { Val = C.BarDirectionValues.Column },
                        new C.BarGrouping { Val = C.BarGroupingValues.Clustered },
                        BuildBarSeries(
                            0u,
                            $"{sheetReference}!${claimsColumn}$4",
                            $"{sheetReference}!$A$5:$A${dataEndRow}",
                            $"{sheetReference}!${claimsColumn}$5:${claimsColumn}${dataEndRow}",
                            "Current Period Paid Claims",
                            categories,
                            paidClaims,
                            "ED7D31",
                            "$#,##0"),
                        new C.AxisId { Val = categoryAxisId },
                        new C.AxisId { Val = claimsAxisId }),
                    new C.LineChart(
                        new C.Grouping { Val = C.GroupingValues.Standard },
                        BuildLineSeries(
                            1u,
                            $"{sheetReference}!${employeeColumn}$4",
                            $"{sheetReference}!$A$5:$A${dataEndRow}",
                            $"{sheetReference}!${employeeColumn}$5:${employeeColumn}${dataEndRow}",
                            "Current Period No. of Employees",
                            categories,
                            employeeCounts,
                            "A9C99C",
                            "#,##0"),
                        new C.AxisId { Val = categoryAxisId },
                        new C.AxisId { Val = employeeAxisId }),
                    BuildCategoryAxis(categoryAxisId, claimsAxisId),
                    BuildValueAxis(claimsAxisId, categoryAxisId, C.AxisPositionValues.Left, "$#,##0", C.CrossesValues.AutoZero),
                    BuildValueAxis(employeeAxisId, categoryAxisId, C.AxisPositionValues.Right, "#,##0", C.CrossesValues.Maximum)),
                new C.Legend(new C.LegendPosition { Val = C.LegendPositionValues.Bottom }, new C.Layout()),
                new C.PlotVisibleOnly { Val = true }));
    }

    private static void ApplyExcelStyle1Chart(C.ChartSpace chartSpace)
    {
        chartSpace.RemoveAllChildren<C.RoundedCorners>();
        chartSpace.RemoveAllChildren<C.Style>();
        chartSpace.RemoveAllChildren<C.ChartShapeProperties>();

        var language = chartSpace.GetFirstChild<C.EditingLanguage>();
        var roundedCorners = new C.RoundedCorners { Val = false };
        if (language is not null)
        {
            language.InsertAfterSelf(roundedCorners);
        }
        else
        {
            chartSpace.PrependChild(roundedCorners);
        }

        roundedCorners.InsertAfterSelf(new C.Style { Val = 1 });

        var chart = chartSpace.GetFirstChild<C.Chart>();
        if (chart is null)
        {
            return;
        }

        chart.RemoveAllChildren<C.Title>();
        chartSpace.Append(BuildWhiteAreaProperties());

        var plotArea = chart.GetFirstChild<C.PlotArea>();
        if (plotArea is not null)
        {
            plotArea.RemoveAllChildren<C.ChartShapeProperties>();
            plotArea.Append(BuildWhiteAreaProperties());

            var barChart = plotArea.GetFirstChild<C.BarChart>();
            if (barChart is not null)
            {
                barChart.GetFirstChild<C.BarDirection>()!.Val = C.BarDirectionValues.Column;
                barChart.GetFirstChild<C.BarGrouping>()!.Val = C.BarGroupingValues.Clustered;

                var series = barChart.Elements<C.BarChartSeries>().ToList();
                if (series.Count > 0)
                {
                    ApplyBarSeriesFill(series[0], Style1NetPremiumColor);
                }

                if (series.Count > 1)
                {
                    ApplyBarSeriesFill(series[1], Style1PaidClaimsColor);
                }

                RemoveChildren<C.DataLabels>(barChart);
                var firstAxisId = barChart.Elements<C.AxisId>().FirstOrDefault();
                if (firstAxisId is not null)
                {
                    firstAxisId.InsertBeforeSelf(BuildSummaryChartDataLabels());
                }
                else
                {
                    barChart.Append(BuildSummaryChartDataLabels());
                }
            }
        }

        RemoveChildren<C.Legend>(chart);
        var legend = BuildStyle1Legend();
        var plotVisibleOnly = chart.GetFirstChild<C.PlotVisibleOnly>();
        if (plotVisibleOnly is not null)
        {
            plotVisibleOnly.InsertBeforeSelf(legend);
        }
        else
        {
            chart.Append(legend);
        }
    }

    private static C.ChartShapeProperties BuildWhiteAreaProperties()
    {
        return new C.ChartShapeProperties(
            new A.SolidFill(new A.RgbColorModelHex { Val = "FFFFFF" }),
            new A.Outline(new A.NoFill()));
    }

    private static void ApplyBarSeriesFill(C.BarChartSeries series, string color)
    {
        RemoveChildren<C.ChartShapeProperties>(series);
        var properties = new C.ChartShapeProperties(
            new A.SolidFill(new A.RgbColorModelHex { Val = color }),
            new A.Outline(new A.NoFill()));
        var categoryAxisData = series.GetFirstChild<C.CategoryAxisData>();
        if (categoryAxisData is not null)
        {
            categoryAxisData.InsertBeforeSelf(properties);
        }
        else
        {
            series.Append(properties);
        }
    }

    private static C.Legend BuildStyle1Legend()
    {
        return new C.Legend(
            new C.LegendPosition { Val = C.LegendPositionValues.Bottom },
            new C.Layout(),
            new C.ChartShapeProperties(new A.Outline(new A.NoFill())),
            BuildChartTextProperties());
    }

    private static void RemoveChildren<TElement>(OpenXmlElement element)
        where TElement : OpenXmlElement
    {
        foreach (var child in element.Elements<TElement>().ToList())
        {
            child.Remove();
        }
    }

    private static C.ChartSpace BuildPieChartSpace(string sheetName, string title, int dataStartRow, int dataEndRow, IReadOnlyList<ChartCategoryValue> rows)
    {
        var sheetReference = SheetReference(sheetName);
        var categories = rows.Select(item => item.Category).ToList();
        var values = rows.Select(item => (decimal?)item.Value).ToList();

        return new C.ChartSpace(
            new C.EditingLanguage { Val = "en-US" },
            new C.Chart(
                BuildChartTitle(title),
                new C.PlotArea(
                    new C.Layout(),
                    new C.PieChart(
                        new C.VaryColors { Val = true },
                        new C.PieChartSeries(
                            new C.Index { Val = 0u },
                            new C.Order { Val = 0u },
                            BuildSeriesText(title, $"{sheetReference}!$A${dataStartRow - 2}"),
                            new C.CategoryAxisData(new C.StringReference(
                                new C.Formula($"{sheetReference}!$A${dataStartRow}:$A${dataEndRow}"),
                                BuildStringCache(categories))),
                            new C.Values(new C.NumberReference(
                                new C.Formula($"{sheetReference}!$B${dataStartRow}:$B${dataEndRow}"),
                                BuildNumberCache(values, "$#,##0")))),
                        new C.FirstSliceAngle { Val = 0 })),
                new C.Legend(new C.LegendPosition { Val = C.LegendPositionValues.Right }, new C.Layout()),
                new C.PlotVisibleOnly { Val = true }));
    }

    private static C.BarChartSeries BuildBarSeries(
        uint index,
        string nameFormula,
        string categoryFormula,
        string valueFormula,
        string name,
        IReadOnlyList<string> categories,
        IReadOnlyList<decimal?> values,
        string color,
        string formatCode)
    {
        return new C.BarChartSeries(
            new C.Index { Val = index },
            new C.Order { Val = index },
            BuildSeriesText(name, nameFormula),
            new C.ChartShapeProperties(new A.SolidFill(new A.RgbColorModelHex { Val = color })),
            new C.CategoryAxisData(new C.StringReference(new C.Formula(categoryFormula), BuildStringCache(categories))),
            new C.Values(new C.NumberReference(new C.Formula(valueFormula), BuildNumberCache(values, formatCode))));
    }

    private static C.LineChartSeries BuildLineSeries(
        uint index,
        string nameFormula,
        string categoryFormula,
        string valueFormula,
        string name,
        IReadOnlyList<string> categories,
        IReadOnlyList<decimal?> values,
        string color,
        string formatCode)
    {
        return new C.LineChartSeries(
            new C.Index { Val = index },
            new C.Order { Val = index },
            BuildSeriesText(name, nameFormula),
            new C.ChartShapeProperties(new A.Outline(new A.SolidFill(new A.RgbColorModelHex { Val = color })) { Width = 28575 }),
            new C.Marker(
                new C.Symbol { Val = C.MarkerStyleValues.Circle },
                new C.Size { Val = 6 },
                new C.ChartShapeProperties(
                    new A.SolidFill(new A.RgbColorModelHex { Val = color }),
                    new A.Outline(new A.SolidFill(new A.RgbColorModelHex { Val = "FFFFFF" })))),
            new C.CategoryAxisData(new C.StringReference(new C.Formula(categoryFormula), BuildStringCache(categories))),
            new C.Values(new C.NumberReference(new C.Formula(valueFormula), BuildNumberCache(values, formatCode))),
            new C.Smooth { Val = false });
    }

    private static C.SeriesText BuildSeriesText(string name, string formula)
    {
        return new C.SeriesText(new C.StringReference(
            new C.Formula(formula),
            new C.StringCache(
                new C.PointCount { Val = 1u },
                new C.StringPoint { Index = 0u, NumericValue = new C.NumericValue(name) })));
    }

    private static C.CategoryAxis BuildCategoryAxis(uint categoryAxisId, uint crossingAxisId)
    {
        return new C.CategoryAxis(
            new C.AxisId { Val = categoryAxisId },
            new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
            new C.AxisPosition { Val = C.AxisPositionValues.Bottom },
            new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo },
            BuildAxisLineProperties(),
            BuildChartTextProperties(),
            new C.CrossingAxis { Val = crossingAxisId },
            new C.Crosses { Val = C.CrossesValues.AutoZero },
            new C.AutoLabeled { Val = true },
            new C.LabelAlignment { Val = C.LabelAlignmentValues.Center },
            new C.LabelOffset { Val = 100 });
    }

    private static C.DataLabels BuildSummaryChartDataLabels()
    {
        return new C.DataLabels(
            new C.NumberingFormat { FormatCode = "$#,##0", SourceLinked = false },
            BuildChartTextProperties(),
            new C.DataLabelPosition { Val = C.DataLabelPositionValues.OutsideEnd },
            new C.ShowLegendKey { Val = false },
            new C.ShowValue { Val = true },
            new C.ShowCategoryName { Val = false },
            new C.ShowSeriesName { Val = false },
            new C.ShowPercent { Val = false },
            new C.ShowBubbleSize { Val = false });
    }

    private static C.ValueAxis BuildValueAxis(
        uint valueAxisId,
        uint categoryAxisId,
        C.AxisPositionValues position,
        string formatCode,
        C.CrossesValues crosses,
        string? gridlineColor = null)
    {
        return new C.ValueAxis(
            new C.AxisId { Val = valueAxisId },
            new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }, new C.MinAxisValue { Val = 0D }),
            new C.AxisPosition { Val = position },
            BuildMajorGridlines(gridlineColor),
            new C.NumberingFormat { FormatCode = formatCode, SourceLinked = false },
            new C.MajorTickMark { Val = C.TickMarkValues.None },
            new C.MinorTickMark { Val = C.TickMarkValues.None },
            new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo },
            BuildAxisLineProperties(),
            BuildChartTextProperties(),
            new C.CrossingAxis { Val = categoryAxisId },
            new C.Crosses { Val = crosses },
            new C.CrossBetween { Val = C.CrossBetweenValues.Between });
    }

    private static C.MajorGridlines BuildMajorGridlines(string? color)
    {
        if (string.IsNullOrWhiteSpace(color))
        {
            return new C.MajorGridlines();
        }

        return new C.MajorGridlines(
            new C.ChartShapeProperties(
                new A.Outline(
                    new A.SolidFill(new A.RgbColorModelHex { Val = color }))
                {
                    Width = 6350
                }));
    }

    private static C.ChartShapeProperties BuildAxisLineProperties()
    {
        return new C.ChartShapeProperties(
            new A.Outline(new A.SolidFill(new A.RgbColorModelHex { Val = Style1AxisLineColor }))
            {
                Width = 9525
            });
    }

    private static C.TextProperties BuildChartTextProperties()
    {
        var runProperties = new A.DefaultRunProperties
        {
            FontSize = 900
        };
        runProperties.Append(
            new A.SolidFill(new A.RgbColorModelHex { Val = Style1TextColor }),
            new A.LatinFont { Typeface = "Calibri" });

        return new C.TextProperties(
            new A.BodyProperties(),
            new A.ListStyle(),
            new A.Paragraph(new A.ParagraphProperties(runProperties)));
    }

    private static C.Title BuildChartTitle(string title)
    {
        var lines = title.Split('\n');
        var paragraph = new A.Paragraph();
        for (var index = 0; index < lines.Length; index++)
        {
            if (index > 0)
            {
                paragraph.Append(new A.Break());
            }

            paragraph.Append(new A.Run(
                new A.RunProperties { Language = "en-US", FontSize = 1800, Bold = true },
                new A.Text(lines[index])));
        }

        return new C.Title(
            new C.ChartText(new C.RichText(new A.BodyProperties(), new A.ListStyle(), paragraph)),
            new C.Layout(),
            new C.Overlay { Val = false });
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

    private static C.NumberingCache BuildNumberCache(IReadOnlyList<decimal?> values, string formatCode)
    {
        var cache = new C.NumberingCache(new C.FormatCode(formatCode), new C.PointCount { Val = (uint)values.Count });
        for (var index = 0; index < values.Count; index++)
        {
            if (!values[index].HasValue)
            {
                continue;
            }

            cache.Append(new C.NumericPoint
            {
                Index = (uint)index,
                NumericValue = new C.NumericValue(values[index]!.Value.ToString("0.##", Culture))
            });
        }

        return cache;
    }

    private static Xdr.TwoCellAnchor BuildChartAnchor(uint drawingId, string chartRelId, string chartName, int fromColumn, int fromRow, int toColumn, int toRow)
    {
        return new Xdr.TwoCellAnchor(
            new Xdr.FromMarker(
                new Xdr.ColumnId((fromColumn - 1).ToString(Culture)),
                new Xdr.ColumnOffset("0"),
                new Xdr.RowId((fromRow - 1).ToString(Culture)),
                new Xdr.RowOffset("0")),
            new Xdr.ToMarker(
                new Xdr.ColumnId((toColumn - 1).ToString(Culture)),
                new Xdr.ColumnOffset("0"),
                new Xdr.RowId((toRow - 1).ToString(Culture)),
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

    private static string SheetReference(string sheetName)
    {
        return $"'{sheetName.Replace("'", "''", StringComparison.Ordinal)}'";
    }

    private static string Quote(string value)
    {
        return $"\"{value.Replace("\"", "\"\"", StringComparison.Ordinal)}\"";
    }

    private static string ClientSetupValueFormula(string fieldName, string fallback)
    {
        return LookupFormula(Quote(fieldName), "tblClientSetup[Field]", "tblClientSetup[Value]", fallback);
    }

    private static string ReportControlValueFormula(string controlName, string fallback)
    {
        return LookupFormula(Quote(controlName), "tblReportControl[Control]", "tblReportControl[Value]", fallback);
    }

    private static string ReportMonthCountFormula()
    {
        return ReportControlValueFormula("Report Month Count", "12");
    }

    private static string CurrentPeriodIdFormula()
    {
        return ReportControlValueFormula("Current Period ID", Quote(PeriodIds.Current12));
    }

    private static string PriorPeriodIdFormula()
    {
        return ReportControlValueFormula("Prior Period ID", Quote(PeriodIds.Prior12));
    }

    private static string BenchmarkPeriodIdFormula()
    {
        return ReportControlValueFormula("Benchmark Period ID", Quote(PeriodIds.Prior12));
    }

    private static string CurrentPeriodLabelFormula()
    {
        return LookupFormula(CurrentPeriodIdFormula(), "tblPeriods[Period ID]", "tblPeriods[Period Label]", Quote(""));
    }

    private static string PriorPeriodLabelFormula()
    {
        return LookupFormula(PriorPeriodIdFormula(), "tblPeriods[Period ID]", "tblPeriods[Period Label]", Quote(""));
    }

    private static string BenchmarkPeriodLabelFormula()
    {
        return LookupFormula(BenchmarkPeriodIdFormula(), "tblPeriods[Period ID]", "tblPeriods[Period Label]", Quote(""));
    }

    private static string CurrentCalculationFormula(string calculationColumn, string coverage)
    {
        return $"SUMIFS(tblCalculationModel[{calculationColumn}],tblCalculationModel[Coverage],{Quote(coverage)})";
    }

    private static string PriorCalculationFormula(string calculationColumn, string coverage)
    {
        return $"SUMIFS(tblCalculationModel[{calculationColumn}],tblCalculationModel[Coverage],{Quote(coverage)})";
    }

    private static string LivesMonthCountFormula(string periodReference, string coverage)
    {
        return $"COUNTIFS(tblMonthlyExperience[Period ID],{periodReference},tblMonthlyExperience[Coverage],{Quote(coverage)})";
    }

    private static string LivesAverageFormula(string periodReference, string coverage)
    {
        var positiveLives = Quote(">0");
        return $"IFERROR(SUMIFS(tblMonthlyExperience[Lives],tblMonthlyExperience[Period ID],{periodReference},tblMonthlyExperience[Coverage],{Quote(coverage)},tblMonthlyExperience[Lives],{positiveLives})/COUNTIFS(tblMonthlyExperience[Period ID],{periodReference},tblMonthlyExperience[Coverage],{Quote(coverage)},tblMonthlyExperience[Lives],{positiveLives}),\"\")";
    }

    private static string CurrentLivesFormula(string periodReference, string coverage)
    {
        var coverageValue = Quote(coverage);
        return $"IFERROR(LOOKUP(2,1/((INDIRECT({Quote("tblMonthlyExperience[Period ID]")})={periodReference})*(INDIRECT({Quote("tblMonthlyExperience[Coverage]")})={coverageValue})*(INDIRECT({Quote("tblMonthlyExperience[Lives]")})>0)),INDIRECT({Quote("tblMonthlyExperience[Lives]")})),\"\")";
    }

    private static string MonthlyLivesForChartFormula(int row, string coverage)
    {
        var coverageValue = Quote(coverage);
        var criteria = $"tblMonthlyExperience[Period ID],'08_Report_Control'!$B$5,tblMonthlyExperience[Coverage],{coverageValue},tblMonthlyExperience[Month],A{row}";
        return $"IF(COUNTIFS({criteria})=0,\"\",SUMIFS(tblMonthlyExperience[Lives],{criteria}))";
    }

    private static string AssumptionValueFormula(string assumptionName, string fallback)
    {
        return LookupFormula(Quote(assumptionName), "tblAssumptions[Assumption]", "tblAssumptions[Value]", fallback);
    }

    private static string LookupFormula(string lookupValue, string lookupColumn, string returnColumn, string fallback)
    {
        return $"IFERROR(INDEX({returnColumn},MATCH({lookupValue},{lookupColumn},0)),{fallback})";
    }

    private static string EstimatedAdjustmentFormula(string coverage, string trendAssumption, string expenseAssumption, string targetLossRatioAssumption)
    {
        return $"IFERROR(({CurrentCalculationFormula("Current Loss Ratio", coverage)}*(1+{AssumptionValueFormula(trendAssumption, "0")}+{AssumptionValueFormula(expenseAssumption, "0")})/{AssumptionValueFormula(targetLossRatioAssumption, "0")})-1,0)";
    }

    private static int AddMonthlyPptTable(
        IXLWorksheet sheet,
        int startRow,
        string title,
        string tableName,
        string coverage,
        IReadOnlyList<ClaimsExperienceMonth> months,
        out int nextRow)
    {
        var premiumColumn = coverage.Equals(CoverageNames.Ehc, StringComparison.OrdinalIgnoreCase) ? "Net Premium" : "Billed Premium";
        var premiumHeader = coverage.Equals(CoverageNames.Ehc, StringComparison.OrdinalIgnoreCase) ? "Net Premium" : "Billed Premium";
        var rows = new List<object?[]>();
        for (var index = 0; index < Math.Max(months.Count, 1); index++)
        {
            if (months.Count == 0)
            {
                rows.Add(["", "", "", "", "No monthly rows were read."]);
                break;
            }

            var month = months[index].Month.ToDateTime(TimeOnly.MinValue);
            rows.Add(
            [
                month,
                $"=SUMIFS(tblMonthlyExperience[{premiumColumn}],tblMonthlyExperience[Period ID],'08_Report_Control'!$B$5,tblMonthlyExperience[Coverage],\"{coverage}\",tblMonthlyExperience[Month],A{startRow + 2 + index})",
                $"=SUMIFS(tblMonthlyExperience[Paid Claims],tblMonthlyExperience[Period ID],'08_Report_Control'!$B$5,tblMonthlyExperience[Coverage],\"{coverage}\",tblMonthlyExperience[Month],A{startRow + 2 + index})",
                $"=IFERROR(C{startRow + 2 + index}/B{startRow + 2 + index},\"\")",
                $"=IF(COUNTIFS(tblMonthlyExperience[Period ID],'08_Report_Control'!$B$5,tblMonthlyExperience[Coverage],\"{coverage}\",tblMonthlyExperience[Month],A{startRow + 2 + index})=0,\"\",SUMIFS(tblMonthlyExperience[Lives],tblMonthlyExperience[Period ID],'08_Report_Control'!$B$5,tblMonthlyExperience[Coverage],\"{coverage}\",tblMonthlyExperience[Month],A{startRow + 2 + index}))",
                ""
            ]);
        }

        return AddTableAt(
            sheet,
            startRow,
            title,
            ["Month", premiumHeader, "Paid Claims", "Loss Ratio", "No. of Employees", "Notes"],
            rows,
            tableName,
            SheetKind.Export,
            out nextRow);
    }

    private static int AddDistributionPptTable(
        IXLWorksheet sheet,
        int startRow,
        string title,
        string tableName,
        string sourceTableName,
        string categoryColumn,
        out int nextRow)
    {
        var rows = new List<object?[]>();
        var labels = sourceTableName switch
        {
            "tblEHCCategory" => new[] { "Drugs", "Paramedical", "Vision", "Hospital", "Medical Services and Supplies", "Out of Country", "Other EHC" },
            "tblDentalCategory" => new[] { "Basic Dental", "Endodontics", "Periodontics", "Major Dental", "Orthodontics", "Other Dental" },
            _ => new[] { "Massage", "Clinical Counsellor", "Physiotherapy", "Chiropractor", "Acupuncture", "Naturopath", "Psychologist", "Other Paramedical" }
        };

        foreach (var label in labels)
        {
            rows.Add(
            [
                label,
                $"=SUMIFS({sourceTableName}[Paid Claims],{sourceTableName}[Period ID],'08_Report_Control'!$B$5,{sourceTableName}[{categoryColumn}],A{startRow + 2 + rows.Count})",
                $"=IFERROR(B{startRow + 2 + rows.Count}/SUM(B{startRow + 2}:B{startRow + 1 + labels.Length}),\"\")",
                ""
            ]);
        }

        return AddTableAt(
            sheet,
            startRow,
            title,
            [categoryColumn, "Paid Claims", "Percent of Claims", "Notes"],
            rows,
            tableName,
            SheetKind.Export,
            out nextRow);
    }

    private static List<object?[]> BuildImportRows(ClaimsExperienceDataModelData data)
    {
        var rows = new List<object?[]>();
        foreach (var source in data.ImportedFiles)
        {
            rows.Add(
            [
                source.ImportedAt,
                source.SourceType,
                source.SourceFile,
                source.FullPath,
                source.RowsRead,
                source.Status,
                source.Notes
            ]);
        }

        return rows;
    }

    private static List<object?[]> BuildClientRows(ClaimsExperienceDataModelData data)
    {
        return
        [
            ["Client Name", data.ClientName, ""],
            ["Carrier", data.Carrier, ""],
            ["Policy Number", data.PolicyNumber, ""],
            ["Experience Start Date", ToDateTime(data.PeriodStart), ""],
            ["Experience End Date", ToDateTime(data.PeriodEnd), ""],
            ["Current Period Label", data.PeriodLabel, ""],
            ["Employee Count", data.EmployeeCount, "Optional input. Used for average spend per employee."],
            ["Renewal Date", "", "Optional input. Used by the ClaimsExperience_Summary next-renewal sentence."],
            ["Consultant", "", "Optional input."],
            ["Notes", "", ""]
        ];
    }

    private static List<object?[]> BuildPeriodRows(ClaimsExperienceDataModelData data)
    {
        return BuildDynamicPeriodRows(data);
    }

    private static List<object?[]> BuildMonthlyRows(ClaimsExperienceDataModelData data)
    {
        var rows = new List<object?[]>();

        AddMonthlyRows(data.BenchmarkSourcePeriodId, SelectExperienceMonths(data.PriorPriorMonths, 12));
        AddMonthlyRows(data.PriorSourcePeriodId, SelectExperienceMonths(data.PriorMonths, data.ReportMonthCount));
        AddMonthlyRows(data.CurrentSourcePeriodId, SelectExperienceMonths(data.Months, data.ReportMonthCount));

        return rows;

        void AddMonthlyRows(string periodId, IReadOnlyList<ClaimsExperienceMonth> months)
        {
            foreach (var month in months)
            {
                rows.Add(CreateMonthlyRow(data, periodId, month.Month, CoverageNames.Ehc, month.EhcPremium, month.EhcPremium, month.EhcClaims, FirstNonBlank(month.EhcSourceFile, month.SourceFile), month.EhcPremiumMatchesMonth, 5 + rows.Count));
                rows.Add(CreateMonthlyRow(data, periodId, month.Month, CoverageNames.Dental, month.DentalPremium, month.DentalPremium, month.DentalClaims, FirstNonBlank(month.DentalSourceFile, month.SourceFile), month.DentalPremiumMatchesMonth, 5 + rows.Count));
            }
        }
    }

    private static object?[] CreateMonthlyRow(ClaimsExperienceDataModelData data, string periodId, DateOnly month, string coverage, decimal billedPremium, decimal netPremium, decimal paidClaims, string sourceFile, bool premiumMatchesMonth, int excelRow)
    {
        var lives = data.BillingMonthlyLivesByMonth.TryGetValue(FirstDayOfMonth(month), out var billingLives)
            ? coverage.Equals(CoverageNames.Ehc, StringComparison.OrdinalIgnoreCase)
                ? billingLives.HealthInsuredLives
                : coverage.Equals(CoverageNames.Dental, StringComparison.OrdinalIgnoreCase)
                    ? billingLives.DentalInsuredLives
                    : null
            : null;
        var billingPremium = billingLives is null
            ? null
            : coverage.Equals(CoverageNames.Ehc, StringComparison.OrdinalIgnoreCase)
                ? billingLives.HealthPremium
                : coverage.Equals(CoverageNames.Dental, StringComparison.OrdinalIgnoreCase)
                    ? billingLives.DentalPremium
                    : null;
        var finalBilledPremium = premiumMatchesMonth && billedPremium != 0m
            ? billedPremium
            : billingPremium ?? billedPremium;
        object finalNetPremium = coverage.Equals(CoverageNames.Ehc, StringComparison.OrdinalIgnoreCase)
            ? $"=IFERROR(D{excelRow}*(1-{AssumptionValueFormula("EHC Pooling Charge", "0.09")}),D{excelRow})"
            : finalBilledPremium;

        return
        [
            periodId,
            month.ToDateTime(TimeOnly.MinValue),
            coverage,
            finalBilledPremium,
            finalNetPremium,
            paidClaims,
            0m,
            paidClaims,
            lives,
            FirstNonBlank(sourceFile, "Insured Experience Report.pdf"),
            ""
        ];
    }

    private static IReadOnlyList<ClaimsExperienceMonth> SelectExperienceMonths(IEnumerable<ClaimsExperienceMonth> sourceRows, int monthCount)
    {
        var selectedMonthCount = NormalizeReportMonthCount(monthCount);
        var rows = sourceRows
            .Where(month => !month.IsAnnualSummary)
            .OrderBy(month => month.Month)
            .ToList();

        if (rows.Count == 0)
        {
            rows = sourceRows
                .OrderBy(month => month.Month)
                .ToList();
        }

        return rows
            .Take(Math.Min(selectedMonthCount, rows.Count))
            .ToList();
    }

    private static List<object?[]> BuildBillingInsuredLivesRows(IEnumerable<BillingInsuredLivesDetailRow> details)
    {
        return details
            .OrderBy(row => row.StatementMonth ?? DateOnly.MaxValue)
            .ThenBy(row => row.SourceFile, StringComparer.OrdinalIgnoreCase)
            .ThenBy(row => row.ExtractedCategory, StringComparer.OrdinalIgnoreCase)
            .Select(row => new object?[]
            {
                row.SourceFile,
                row.StatementPeriod,
                row.Benefit,
                row.DepCode,
                row.Lives,
                row.ExtractedCategory
            })
            .ToList();
    }

    private static List<object?[]> BuildCategoryRows(IEnumerable<ClaimsCategoryDetail> categories)
    {
        return categories
            .OrderBy(row => PeriodSort(row.PeriodId))
            .ThenByDescending(row => row.PaidClaims)
            .Select(row => new object?[]
            {
                row.PeriodId,
                row.Coverage,
                row.Category,
                row.ClaimCount,
                row.SubmittedAmount,
                row.EligibleAmount,
                row.PaidClaims,
                row.SourceFile,
                row.Notes
            })
            .ToList();
    }

    private static List<object?[]> BuildParamedicalRows(IEnumerable<ParamedicalDetail> details)
    {
        return details
            .OrderBy(row => PeriodSort(row.PeriodId))
            .ThenByDescending(row => row.PaidClaims)
            .Select(row => new object?[]
            {
                row.PeriodId,
                row.Service,
                row.ClaimCount,
                row.SubmittedAmount,
                row.EligibleAmount,
                row.PaidClaims,
                row.SourceFile,
                row.Notes
            })
            .ToList();
    }

    private static int PeriodSort(string periodId)
    {
        if (periodId.StartsWith("PRIOR_PRIOR_", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }

        if (periodId.StartsWith("PRIOR_", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }

        return periodId.StartsWith("CURR_", StringComparison.OrdinalIgnoreCase) ? 2 : 99;
    }

    private static List<object?[]> BuildAssumptionRows(ClaimsExperienceDataModelData data)
    {
        return
        [
            ["Default Net Premium Factor", 1m, "All coverages", "Billed premium is used as net premium until a fee basis is supplied."],
            ["EHC Pooling Charge", data.RenewalTerms.EffectiveEhcPoolingCharge, CoverageNames.Ehc, "Saved renewal pooling charge; defaults to 9.00% when no client-specific term is available."],
            ["EHC Trend Factor", data.RenewalTerms.EffectiveEhcTrendFactor, CoverageNames.Ehc, "Saved renewal trend; defaults to 15.00% when no client-specific term is available."],
            ["Dental Trend Factor + Fee Guide", data.RenewalTerms.EffectiveDentalTrendFactor, CoverageNames.Dental, "Saved renewal trend; defaults to 15.00% when no client-specific term is available."],
            ["EHC Total Expense Factor", 0m, CoverageNames.Ehc, "Additional expense load used in the estimated rate adjustment."],
            ["Dental Total Expense Factor", 0m, CoverageNames.Dental, "Additional expense load used in the estimated rate adjustment."],
            ["EHC Break-even Point (TLR)", data.RenewalTerms.EffectiveEhcTargetLossRatio, CoverageNames.Ehc, "Saved renewal target loss ratio; defaults to 76.30% when no client-specific term is available."],
            ["Dental Break-even Point (TLR)", data.RenewalTerms.EffectiveDentalTargetLossRatio, CoverageNames.Dental, "Saved renewal target loss ratio; defaults to 76.30% when no client-specific term is available."],
            ["HSA Included", false, CoverageNames.Hsa, "Future source rows can be added to tblMonthlyExperience."],
            ["ASO Fees Included", false, CoverageNames.AsoFees, "Future source rows can be added to tblMonthlyExperience."],
            ["PowerPoint Source Tab", "11_Export_PowerPoint", "PowerPoint", "PowerPoint slides should map to the named export tables on this tab."],
            ["Prior Period Handling", "Blank until supplied", "Prior periods", "Prior-period formulas are ready but require prior source rows."]
        ];
    }

    private static List<object?[]> BuildReportControlRows(ClaimsExperienceDataModelData data)
    {
        return
        [
            ["Current Period ID", "=\"CURR_\"&B8", "Driven by Report Month Count."],
            ["Prior Period ID", "=\"PRIOR_\"&B8", "Uses the matched prior period with the same month count."],
            ["Report View", "=B8&\" Month vs Prior \"&B8&\" Month\"", ""],
            ["Report Month Count", data.ReportMonthCount, "Automatically detected from the experience report unless a shorter manual period was selected."],
            ["Benchmark Period ID", "=IF(B8=12,\"PRIOR_PRIOR_12\",\"PRIOR_12\")", "Used by comparison PowerPoint charts as the left-most benchmark period."],
            ["Include HSA", false, ""],
            ["Include ASO Fees", false, ""],
            ["PowerPoint Export Source", "11_Export_PowerPoint", ""],
            ["Last Refresh", DateTime.Now, ""]
        ];
    }

    private static List<object?[]> BuildParamedicalMonthlyPptRows(IReadOnlyList<ClaimsExperienceMonth> months)
    {
        var rows = new List<object?[]>();
        if (months.Count == 0)
        {
            rows.Add(["", "", "No monthly rows were read."]);
            return rows;
        }

        foreach (var month in months)
        {
            rows.Add([month.Month.ToDateTime(TimeOnly.MinValue), 0m, "Monthly paramedical detail needs a monthly source workbook."]);
        }

        return rows;
    }

    private static List<object?[]> BuildCombinedDistributionRows(int tableStartRow)
    {
        var firstDataRow = tableStartRow + 2;
        var secondDataRow = tableStartRow + 3;
        return
        [
            ["EHC", "=SUMIFS(tblCalculationModel[Current Total Claims],tblCalculationModel[Coverage],\"EHC\")", $"=IFERROR(B{firstDataRow}/SUM($B${firstDataRow}:$B${secondDataRow}),\"\")", ""],
            ["Dental", "=SUMIFS(tblCalculationModel[Current Total Claims],tblCalculationModel[Coverage],\"Dental\")", $"=IFERROR(B{secondDataRow}/SUM($B${firstDataRow}:$B${secondDataRow}),\"\")", ""]
        ];
    }

    private static List<object?[]> BuildClaimsDistributionChartRows()
    {
        var currentPeriod = "'08_Report_Control'!$B$5";
        var priorPeriod = "'08_Report_Control'!$B$6";
        return
        [
            ["Dental", ClaimsDistributionDentalFormula(currentPeriod), ClaimsDistributionDentalFormula(priorPeriod), ""],
            ["Paramedical", ClaimsDistributionEhcFormula(currentPeriod, "Paramedical"), ClaimsDistributionEhcFormula(priorPeriod, "Paramedical"), ""],
            ["Drugs", ClaimsDistributionEhcFormula(currentPeriod, "Drugs"), ClaimsDistributionEhcFormula(priorPeriod, "Drugs"), ""],
            ["Vision", ClaimsDistributionEhcFormula(currentPeriod, "Vision"), ClaimsDistributionEhcFormula(priorPeriod, "Vision"), ""],
            ["Other EHC", ClaimsDistributionOtherEhcFormula(currentPeriod), ClaimsDistributionOtherEhcFormula(priorPeriod), ""]
        ];
    }

    private static string ClaimsDistributionDentalFormula(string periodReference)
    {
        return $"=SUMIFS(tblDentalCategory[Paid Claims],tblDentalCategory[Period ID],{periodReference})";
    }

    private static string ClaimsDistributionEhcFormula(string periodReference, string category)
    {
        return $"=SUMIFS(tblEHCCategory[Paid Claims],tblEHCCategory[Period ID],{periodReference},tblEHCCategory[Category],\"{category}\")";
    }

    private static string ClaimsDistributionOtherEhcFormula(string periodReference)
    {
        return $"=MAX(0,SUMIFS(tblEHCCategory[Paid Claims],tblEHCCategory[Period ID],{periodReference})" +
               $"-SUMIFS(tblEHCCategory[Paid Claims],tblEHCCategory[Period ID],{periodReference},tblEHCCategory[Category],\"Paramedical\")" +
               $"-SUMIFS(tblEHCCategory[Paid Claims],tblEHCCategory[Period ID],{periodReference},tblEHCCategory[Category],\"Drugs\")" +
               $"-SUMIFS(tblEHCCategory[Paid Claims],tblEHCCategory[Period ID],{periodReference},tblEHCCategory[Category],\"Vision\"))";
    }

    private static List<object?[]> BuildParamedicalBreakdownChartRows()
    {
        var currentPeriod = "'08_Report_Control'!$B$5";
        var priorPeriod = "'08_Report_Control'!$B$6";
        return
        [
            ParamedicalBreakdownChartRow("Acupuncture", currentPeriod, priorPeriod, ["Acupuncture"]),
            ParamedicalBreakdownChartRow("Chiropodist", currentPeriod, priorPeriod, ["Chiropodist"]),
            ParamedicalBreakdownChartRow("Chiropractor", currentPeriod, priorPeriod, ["Chiropractor"]),
            ParamedicalBreakdownChartRow("Clinical Counsellor", currentPeriod, priorPeriod, ["Clinical Counsellor", "Clinical Counselor"]),
            ParamedicalBreakdownChartRow("Dietitian", currentPeriod, priorPeriod, ["Dietitian", "Dietician"]),
            ParamedicalBreakdownChartRow("Marriage / Family Therapist", currentPeriod, priorPeriod, ["Marriage / Family Therapist", "Marriage Family Therapist", "Family Therapist"]),
            ParamedicalBreakdownChartRow("Naturopath", currentPeriod, priorPeriod, ["Naturopath"]),
            ParamedicalBreakdownChartRow("Osteopath", currentPeriod, priorPeriod, ["Osteopath"]),
            ParamedicalBreakdownChartRow("Physio", currentPeriod, priorPeriod, ["Physiotherapy", "Physio"]),
            ParamedicalBreakdownChartRow("Podiatrist", currentPeriod, priorPeriod, ["Podiatrist"]),
            ParamedicalBreakdownChartRow("Psycho-analyst", currentPeriod, priorPeriod, ["Psycho-analyst", "Psychoanalyst"]),
            ParamedicalBreakdownChartRow("Psychologist", currentPeriod, priorPeriod, ["Psychologist"]),
            ParamedicalBreakdownChartRow("Massage", currentPeriod, priorPeriod, ["Massage"]),
            ParamedicalBreakdownChartRow("Social Worker", currentPeriod, priorPeriod, ["Social Worker"]),
            ParamedicalBreakdownChartRow("Speech Therapy", currentPeriod, priorPeriod, ["Speech Therapy", "Speech Therapist"])
        ];
    }

    private static object?[] ParamedicalBreakdownChartRow(
        string label,
        string currentPeriodReference,
        string priorPeriodReference,
        IReadOnlyList<string> serviceNames)
    {
        return
        [
            label,
            ParamedicalServiceFormula(currentPeriodReference, serviceNames),
            ParamedicalServiceFormula(priorPeriodReference, serviceNames),
            ""
        ];
    }

    private static string ParamedicalServiceFormula(string periodReference, IReadOnlyList<string> serviceNames)
    {
        return "=" + string.Join("+", serviceNames.Select(service =>
            $"SUMIFS(tblParamedical[Paid Claims],tblParamedical[Period ID],{periodReference},tblParamedical[Service],{Quote(service)})"));
    }

    private static List<object?[]> BuildDentalBreakdownChartRows()
    {
        var currentPeriod = "'08_Report_Control'!$B$5";
        var priorPeriod = "'08_Report_Control'!$B$6";
        return
        [
            DentalBreakdownChartRow("Basic Dental", currentPeriod, priorPeriod, ["Basic Dental", "Preventive", "Restorative", "Dental Diagnostic", "Basic Restorative", "Basic Surgical"]),
            DentalBreakdownChartRow("Periodontics", currentPeriod, priorPeriod, ["Periodontics", "Periodontic"]),
            DentalBreakdownChartRow("Major Dental", currentPeriod, priorPeriod, ["Major Dental", "Major Restorative", "Prosthetics", "Crowns"]),
            DentalBreakdownChartRow("Orthodontics", currentPeriod, priorPeriod, ["Orthodontics", "Orthodontic"]),
            DentalBreakdownChartRow("Endodontics", currentPeriod, priorPeriod, ["Endodontics", "Endodontic"]),
            DentalBreakdownChartRow("Dentures", currentPeriod, priorPeriod, ["Dentures"]),
            DentalBreakdownChartRow("Oral Surgery", currentPeriod, priorPeriod, ["Oral Surgery"])
        ];
    }

    private static object?[] DentalBreakdownChartRow(
        string label,
        string currentPeriodReference,
        string priorPeriodReference,
        IReadOnlyList<string> categoryNames)
    {
        return
        [
            label,
            DentalCategoryFormula(currentPeriodReference, categoryNames),
            DentalCategoryFormula(priorPeriodReference, categoryNames),
            ""
        ];
    }

    private static string DentalCategoryFormula(string periodReference, IReadOnlyList<string> categoryNames)
    {
        return "=" + string.Join("+", categoryNames.Select(category =>
            $"SUMIFS(tblDentalCategory[Paid Claims],tblDentalCategory[Period ID],{periodReference},tblDentalCategory[Category],{Quote(category)})"));
    }

    private static int AddTable(
        IXLWorksheet sheet,
        string title,
        IReadOnlyList<string> headers,
        IReadOnlyList<object?[]> rows,
        string tableName,
        SheetKind kind)
    {
        FormatTitle(sheet, title, kind);
        return AddTableAt(sheet, 3, title, headers, rows, tableName, kind, out _);
    }

    private static int AddTableAt(
        IXLWorksheet sheet,
        int startRow,
        string title,
        IReadOnlyList<string> headers,
        IReadOnlyList<object?[]> rows,
        string tableName,
        SheetKind kind,
        out int nextRow)
    {
        sheet.Cell(startRow, 1).Value = title;
        sheet.Cell(startRow, 1).Style.Font.Bold = true;
        sheet.Cell(startRow, 1).Style.Font.FontColor = XLColor.FromHtml("#17202A");

        var headerRow = startRow + 1;
        for (var column = 0; column < headers.Count; column++)
        {
            sheet.Cell(headerRow, column + 1).Value = headers[column];
        }

        var dataRows = rows.Count > 0 ? rows : [headers.Select(_ => (object?)"").ToArray()];
        for (var rowIndex = 0; rowIndex < dataRows.Count; rowIndex++)
        {
            for (var column = 0; column < headers.Count; column++)
            {
                var value = column < dataRows[rowIndex].Length ? dataRows[rowIndex][column] : "";
                SetCellValue(sheet.Cell(headerRow + 1 + rowIndex, column + 1), value);
            }
        }

        var range = sheet.Range(headerRow, 1, headerRow + dataRows.Count, headers.Count);
        var table = range.CreateTable(tableName);
        table.Theme = XLTableTheme.TableStyleMedium2;
        StyleTable(sheet, headerRow, dataRows.Count, headers, kind);
        if (tableName.Equals("tblAssumptions", StringComparison.OrdinalIgnoreCase))
        {
            sheet.Range(headerRow + 2, 2, headerRow + 8, 2).Style.NumberFormat.Format = "0.00%";
        }

        nextRow = headerRow + dataRows.Count + 4;
        return 1;
    }

    private static void FormatTitle(IXLWorksheet sheet, string title, SheetKind kind)
    {
        sheet.Cell(1, 1).Value = title;
        sheet.Cell(1, 1).Style.Font.Bold = true;
        sheet.Cell(1, 1).Style.Font.FontSize = 16;
        sheet.Cell(1, 1).Style.Font.FontColor = XLColor.White;
        sheet.Range(1, 1, 1, 6).Merge();
        sheet.Range(1, 1, 1, 6).Style.Fill.BackgroundColor = HeaderColor(kind);
        sheet.Row(1).Height = 24;
        sheet.SheetView.FreezeRows(4);
    }

    private static void HeaderRange(IXLRange range, SheetKind kind)
    {
        range.Style.Fill.BackgroundColor = HeaderColor(kind);
        range.Style.Font.FontColor = XLColor.White;
        range.Style.Font.Bold = true;
    }

    private static void StyleTable(IXLWorksheet sheet, int headerRow, int dataRowCount, IReadOnlyList<string> headers, SheetKind kind)
    {
        var headerRange = sheet.Range(headerRow, 1, headerRow, headers.Count);
        HeaderRange(headerRange, kind);

        for (var column = 1; column <= headers.Count; column++)
        {
            var header = headers[column - 1];
            var dataRange = sheet.Range(headerRow + 1, column, headerRow + dataRowCount, column);
            if (header.Contains("Date", StringComparison.OrdinalIgnoreCase) ||
                header.Equals("Month", StringComparison.OrdinalIgnoreCase) ||
                header.Equals("Last Refresh", StringComparison.OrdinalIgnoreCase))
            {
                dataRange.Style.DateFormat.Format = "mmm d, yyyy";
            }
            else if (header.Contains("Premium", StringComparison.OrdinalIgnoreCase) ||
                     header.Contains("Claims", StringComparison.OrdinalIgnoreCase) ||
                     header.Contains("Fees", StringComparison.OrdinalIgnoreCase) ||
                     header.Contains("Spend", StringComparison.OrdinalIgnoreCase) ||
                     header.Contains("Amount", StringComparison.OrdinalIgnoreCase))
            {
                dataRange.Style.NumberFormat.Format = "$#,##0.00";
            }
            else if (header.Contains("Ratio", StringComparison.OrdinalIgnoreCase) ||
                     header.Contains("Percent", StringComparison.OrdinalIgnoreCase))
            {
                dataRange.Style.NumberFormat.Format = "0.0%";
            }
            else if (header.Contains("Employee", StringComparison.OrdinalIgnoreCase) ||
                     header.Contains("Lives", StringComparison.OrdinalIgnoreCase))
            {
                dataRange.Style.NumberFormat.Format = "#,##0";
            }
        }
    }

    private static void SetCellValue(IXLCell cell, object? value)
    {
        switch (value)
        {
            case null:
                cell.Value = "";
                break;
            case string text when text.StartsWith('='):
                cell.FormulaA1 = text[1..];
                break;
            case DateTime date:
                cell.Value = date;
                break;
            case DateOnly date:
                cell.Value = date.ToDateTime(TimeOnly.MinValue);
                break;
            case bool boolean:
                cell.Value = boolean;
                break;
            case int integer:
                cell.Value = integer;
                break;
            case decimal number:
                cell.Value = number;
                break;
            case double number:
                cell.Value = number;
                break;
            default:
                cell.Value = value.ToString() ?? "";
                break;
        }
    }

    private static XLColor HeaderColor(SheetKind kind)
    {
        return kind switch
        {
            SheetKind.Input => XLColor.FromHtml("#2F6F73"),
            SheetKind.Control => XLColor.FromHtml("#5F6B7A"),
            SheetKind.Formula => XLColor.FromHtml("#1F5D99"),
            SheetKind.Export => XLColor.FromHtml("#17202A"),
            _ => XLColor.FromHtml("#17202A")
        };
    }

    private static string SumMonthlyFormula(string periodReference, string coverage, string amountColumn)
    {
        var columnReference = $"tblMonthlyExperience[{amountColumn}]";
        if (coverage.Equals("Total", StringComparison.OrdinalIgnoreCase))
        {
            return $"SUMIFS({columnReference},tblMonthlyExperience[Period ID],{periodReference})";
        }

        if (coverage.Equals("Combined EHC + Dental", StringComparison.OrdinalIgnoreCase))
        {
            return $"SUMIFS({columnReference},tblMonthlyExperience[Period ID],{periodReference},tblMonthlyExperience[Coverage],\"EHC\")+" +
                   $"SUMIFS({columnReference},tblMonthlyExperience[Period ID],{periodReference},tblMonthlyExperience[Coverage],\"Dental\")";
        }

        return $"SUMIFS({columnReference},tblMonthlyExperience[Period ID],{periodReference},tblMonthlyExperience[Coverage],\"{coverage}\")";
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

    private static void RemoveCalculatedTableColumnFormulas(string workbookPath)
    {
        using var document = SpreadsheetDocument.Open(workbookPath, true);
        var workbookPart = document.WorkbookPart;
        if (workbookPart is null)
        {
            return;
        }

        foreach (var tablePart in workbookPart.WorksheetParts.SelectMany(part => part.TableDefinitionParts))
        {
            var formulas = tablePart.Table.Descendants<S.CalculatedColumnFormula>().ToList();
            foreach (var formula in formulas)
            {
                formula.Remove();
            }

            if (formulas.Count > 0)
            {
                tablePart.Table.Save();
            }
        }
    }

    private static List<string> ResolveSourceFiles(ClaimsExperienceDataModelRequest request)
    {
        var files = new List<string>();
        if (!string.IsNullOrWhiteSpace(request.SourceFolder) && Directory.Exists(request.SourceFolder))
        {
            var currentFolder = Path.Combine(request.SourceFolder, "Current");
            if (Directory.Exists(currentFolder))
            {
                files.AddRange(ResolveSourceFilesFromFolder(currentFolder));
            }
            else
            {
                files.AddRange(Directory.EnumerateFiles(request.SourceFolder, "*.*", SearchOption.TopDirectoryOnly));
            }
        }

        files.AddRange(request.SourceFiles.Where(path => !string.IsNullOrWhiteSpace(path)));

        return files
            .Where(File.Exists)
            .Where(path => IsAllowedSourceExtension(Path.GetExtension(path)))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static List<string> ResolveSourceFilesFromFolder(string folder)
    {
        if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
        {
            return [];
        }

        return Directory.EnumerateFiles(folder, "*.*", SearchOption.TopDirectoryOnly)
            .Where(File.Exists)
            .Where(path => IsAllowedSourceExtension(Path.GetExtension(path)))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static string? FindPriorPriorFolder(string sourceFolder)
    {
        var candidateNames = new[]
        {
            "PriorPrior",
            "Prior Prior",
            "Prior_Prior",
            "Prior-Prior",
            "PriorPrior12",
            "Prior Prior 12",
            "Prior_Prior_12",
            "Prior-Prior-12"
        };

        return candidateNames
            .Select(name => Path.Combine(sourceFolder, name))
            .FirstOrDefault(Directory.Exists);
    }

    private static string ResolveCurrentSourceFolder(string sourceFolder)
    {
        var currentFolder = Path.Combine(sourceFolder, "Current");
        return Directory.Exists(currentFolder)
            ? currentFolder
            : sourceFolder;
    }

    private static bool IsAllowedSourceExtension(string extension)
    {
        return extension.Equals(".pdf", StringComparison.OrdinalIgnoreCase) ||
               extension.Equals(".xls", StringComparison.OrdinalIgnoreCase) ||
               extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) ||
               extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase);
    }

    private static ClaimsImportRecord CreateImportRecord(string path, string sourceType)
    {
        return new ClaimsImportRecord
        {
            ImportedAt = DateTime.Now,
            SourceType = sourceType,
            SourceFile = Path.GetFileName(path),
            FullPath = path,
            RowsRead = "",
            Status = "Read",
            Notes = ""
        };
    }

    private static string? FindClaimsWorkbook(
        IReadOnlyList<string> sourceFiles,
        string coverage,
        DateOnly? periodStart = null,
        DateOnly? periodEnd = null)
    {
        var candidates = new List<(string Path, int SourceOrder, bool ExactPeriod, DateOnly? PeriodEnd)>();
        var legacyCandidates = new List<(string Path, int SourceOrder)>();

        for (var index = 0; index < sourceFiles.Count; index++)
        {
            var path = sourceFiles[index];
            var extension = Path.GetExtension(path);
            if (!IsExcelExtension(extension) ||
                Path.GetFileNameWithoutExtension(path).Contains("monthly", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (extension.Equals(".xls", StringComparison.OrdinalIgnoreCase))
            {
                if (FileNameMatchesCoverage(path, coverage))
                {
                    legacyCandidates.Add((path, index));
                }

                continue;
            }

            try
            {
                using var workbook = OpenWorkbookReadOnly(path);
                if (!workbook.TryGetWorksheet("Summary", out var sheet) ||
                    !FindClaimsSummaryHeaderRow(sheet).HasValue ||
                    !ClaimsSummaryContainsCoverage(sheet, coverage))
                {
                    continue;
                }

                var hasPeriod = TryReadWorkbookPeriod(sheet, out var candidateStart, out var candidateEnd);
                if (periodStart.HasValue && periodEnd.HasValue && hasPeriod &&
                    (candidateStart != periodStart.Value || candidateEnd != periodEnd.Value))
                {
                    continue;
                }

                candidates.Add((
                    path,
                    index,
                    hasPeriod && periodStart.HasValue && periodEnd.HasValue &&
                    candidateStart == periodStart.Value && candidateEnd == periodEnd.Value,
                    hasPeriod ? candidateEnd : null));
            }
            catch
            {
                // Unsupported or unreadable workbooks are not valid claims-detail candidates.
            }
        }

        return candidates
                   .OrderByDescending(candidate => candidate.ExactPeriod)
                   .ThenByDescending(candidate => candidate.PeriodEnd.HasValue)
                   .ThenByDescending(candidate => candidate.PeriodEnd ?? DateOnly.MinValue)
                   .ThenBy(candidate => candidate.SourceOrder)
                   .Select(candidate => candidate.Path)
                   .FirstOrDefault()
               ?? legacyCandidates
                   .OrderBy(candidate => candidate.SourceOrder)
                   .Select(candidate => candidate.Path)
                   .FirstOrDefault();
    }

    private static string? FindExperienceWorkbook(IReadOnlyList<string> sourceFiles)
    {
        var candidates = new List<(string Path, int SourceOrder, DateOnly? PeriodEnd)>();
        for (var index = 0; index < sourceFiles.Count; index++)
        {
            var path = sourceFiles[index];
            var extension = Path.GetExtension(path);
            if (!extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) &&
                !extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            try
            {
                using var workbook = OpenWorkbookReadOnly(path);
                if (!TryGetExperienceWorkbookSheet(workbook, out var sheet, out _))
                {
                    continue;
                }

                candidates.Add((
                    path,
                    index,
                    TryReadWorkbookPeriod(sheet, out _, out var end) ? end : null));
            }
            catch
            {
                // Unsupported or unreadable workbooks are not valid experience candidates.
            }
        }

        return candidates
            .OrderByDescending(candidate => candidate.PeriodEnd.HasValue)
            .ThenByDescending(candidate => candidate.PeriodEnd ?? DateOnly.MinValue)
            .ThenBy(candidate => candidate.SourceOrder)
            .Select(candidate => candidate.Path)
            .FirstOrDefault();
    }

    private static bool TryGetExperienceWorkbookSheet(
        XLWorkbook workbook,
        out IXLWorksheet sheet,
        out int headerRow)
    {
        sheet = null!;
        headerRow = 0;
        if (!workbook.TryGetWorksheet("Experience Combined", out var candidate))
        {
            return false;
        }

        foreach (var row in candidate.RowsUsed().Take(20))
        {
            if (NormalizeHeader(CellText(row.Cell(1))).Equals("Month", StringComparison.OrdinalIgnoreCase) &&
                NormalizeHeader(CellText(row.Cell(2))).StartsWith("Premium", StringComparison.OrdinalIgnoreCase) &&
                NormalizeHeader(CellText(row.Cell(3))).Contains("Claim", StringComparison.OrdinalIgnoreCase) &&
                NormalizeHeader(CellText(row.Cell(5))).StartsWith("Premium", StringComparison.OrdinalIgnoreCase) &&
                NormalizeHeader(CellText(row.Cell(6))).Contains("Claim", StringComparison.OrdinalIgnoreCase))
            {
                sheet = candidate;
                headerRow = row.RowNumber();
                return true;
            }
        }

        return false;
    }

    private static bool TryReadWorkbookPeriod(
        IXLWorksheet sheet,
        out DateOnly periodStart,
        out DateOnly periodEnd)
    {
        periodStart = default;
        periodEnd = default;
        foreach (var row in sheet.RowsUsed().Take(15))
        {
            var lastColumn = Math.Min(row.LastCellUsed()?.Address.ColumnNumber ?? 0, 10);
            for (var column = 1; column <= lastColumn; column++)
            {
                if (TryExtractExperienceReportPeriod(CellText(row.Cell(column)), out periodStart, out periodEnd))
                {
                    return true;
                }
            }
        }

        return false;
    }

    private static int? FindClaimsSummaryHeaderRow(IXLWorksheet sheet)
    {
        foreach (var row in sheet.RowsUsed().Take(20))
        {
            if (NormalizeHeader(CellText(row.Cell(1))).Equals("Benefit", StringComparison.OrdinalIgnoreCase) &&
                NormalizeHeader(CellText(row.Cell(2))).Contains("Benefitcategory", StringComparison.OrdinalIgnoreCase) &&
                NormalizeHeader(CellText(row.Cell(3))).Equals("Service", StringComparison.OrdinalIgnoreCase) &&
                NormalizeHeader(CellText(row.Cell(4))).Contains("Numberofoccurrences", StringComparison.OrdinalIgnoreCase) &&
                NormalizeHeader(CellText(row.Cell(5))).Contains("Totalamountsubmitted", StringComparison.OrdinalIgnoreCase) &&
                NormalizeHeader(CellText(row.Cell(6))).Contains("Totalamounteligible", StringComparison.OrdinalIgnoreCase) &&
                NormalizeHeader(CellText(row.Cell(7))).Contains("Totalamountpaid", StringComparison.OrdinalIgnoreCase))
            {
                return row.RowNumber();
            }
        }

        return null;
    }

    private static bool ClaimsSummaryContainsCoverage(IXLWorksheet sheet, string coverage)
    {
        var headerRow = FindClaimsSummaryHeaderRow(sheet);
        if (!headerRow.HasValue)
        {
            return false;
        }

        var currentBenefit = "";
        foreach (var row in sheet.RowsUsed().Where(row => row.RowNumber() > headerRow.Value))
        {
            var benefit = ParsingHelpers.Clean(CellText(row.Cell(1)));
            if (!string.IsNullOrWhiteSpace(benefit))
            {
                currentBenefit = benefit;
            }

            if (ClaimsBenefitMatchesCoverage(currentBenefit, coverage))
            {
                return true;
            }
        }

        return false;
    }

    private static bool ClaimsBenefitMatchesCoverage(string benefit, string coverage)
    {
        var normalized = NormalizeHeader(benefit);
        if (coverage.Equals(CoverageNames.Dental, StringComparison.OrdinalIgnoreCase))
        {
            return normalized.Contains("Dental", StringComparison.OrdinalIgnoreCase);
        }

        return coverage.Equals(CoverageNames.Ehc, StringComparison.OrdinalIgnoreCase) &&
               (normalized.Equals("EHC", StringComparison.OrdinalIgnoreCase) ||
                normalized.Contains("Health", StringComparison.OrdinalIgnoreCase));
    }

    private static bool FileNameMatchesCoverage(string path, string coverage)
    {
        var fileName = Path.GetFileNameWithoutExtension(path);
        return coverage.Equals(CoverageNames.Dental, StringComparison.OrdinalIgnoreCase)
            ? fileName.Contains("dental", StringComparison.OrdinalIgnoreCase)
            : fileName.Contains("ehc", StringComparison.OrdinalIgnoreCase) ||
              fileName.Contains("health", StringComparison.OrdinalIgnoreCase);
    }

    private static string? FindMonthlyWorkbook(IReadOnlyList<string> sourceFiles, string requiredNamePart)
    {
        return sourceFiles
            .Where(path => IsExcelExtension(Path.GetExtension(path)))
            .Where(path => Path.GetFileNameWithoutExtension(path).Contains(requiredNamePart, StringComparison.OrdinalIgnoreCase))
            .Where(path => Path.GetFileNameWithoutExtension(path).Contains("monthly", StringComparison.OrdinalIgnoreCase))
            .OrderBy(path => Path.GetFileName(path), StringComparer.OrdinalIgnoreCase)
            .FirstOrDefault();
    }

    private string? FindExperienceReport(IReadOnlyList<string> sourceFiles)
    {
        return FindExperienceReports(sourceFiles).FirstOrDefault();
    }

    private IReadOnlyList<string> FindExperienceReports(IReadOnlyList<string> sourceFiles)
    {
        var pdfCandidates = sourceFiles
            .Where(path => Path.GetExtension(path).Equals(".pdf", StringComparison.OrdinalIgnoreCase))
            .Where(path => !IsClaimsServicePdfName(path))
            .Where(path => !IsHybridAsoDentalClaimsExperienceFileName(path))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        var recognizedCandidates = pdfCandidates
            .Where(path => ExperienceReportNamePriority(path) < int.MaxValue)
            .ToList();
        var candidates = recognizedCandidates.Count > 0 ? recognizedCandidates : pdfCandidates;

        return candidates
            .Select(path => new
            {
                Path = path,
                PeriodEnd = ReadExperienceReportPeriodEnd(path),
                NamePriority = ExperienceReportNamePriority(path)
            })
            .OrderByDescending(candidate => candidate.PeriodEnd.HasValue)
            .ThenByDescending(candidate => candidate.PeriodEnd ?? DateOnly.MinValue)
            .ThenBy(candidate => candidate.NamePriority)
            .ThenBy(candidate => Path.GetFileName(candidate.Path), StringComparer.OrdinalIgnoreCase)
            .Select(candidate => candidate.Path)
            .ToList();
    }

    private static bool IsHybridAsoDentalClaimsExperienceFileName(string path) =>
        Path.GetFileNameWithoutExtension(path).Contains("DEN_ClaimsExperience", StringComparison.OrdinalIgnoreCase);

    private static int ExperienceReportNamePriority(string path)
    {
        var normalized = NormalizeHeader(Path.GetFileNameWithoutExtension(path));
        if (normalized.Contains("insuredexperiencebymonth", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }

        if (normalized.Contains("experiencebymonth", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }

        if (normalized.Contains("livespremiumclaims", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("livespremiumsclaims", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }

        if (normalized.Contains("premclaims", StringComparison.OrdinalIgnoreCase))
        {
            return 3;
        }

        if (normalized.Contains("premiumclaims", StringComparison.OrdinalIgnoreCase))
        {
            return 4;
        }

        if (normalized.Contains("insuredexperience", StringComparison.OrdinalIgnoreCase))
        {
            return 5;
        }

        return normalized.Contains("experiencereport", StringComparison.OrdinalIgnoreCase)
            ? 6
            : int.MaxValue;
    }

    private static bool IsClaimsServicePdfName(string path)
    {
        if (ExperienceReportNamePriority(path) < int.MaxValue)
        {
            return false;
        }

        var normalized = NormalizeHeader(Path.GetFileNameWithoutExtension(path));
        return normalized.Contains("claim", StringComparison.OrdinalIgnoreCase) &&
               (normalized.Contains("ehc", StringComparison.OrdinalIgnoreCase) ||
                normalized.Contains("dental", StringComparison.OrdinalIgnoreCase) ||
                normalized.Contains("health", StringComparison.OrdinalIgnoreCase));
    }

    private static string? FindFileByNormalizedName(IReadOnlyList<string> sourceFiles, string[] extensions, string requiredNamePart)
    {
        return sourceFiles
            .Where(path => extensions.Contains(Path.GetExtension(path), StringComparer.OrdinalIgnoreCase))
            .Where(path => NormalizeHeader(Path.GetFileNameWithoutExtension(path)).Contains(requiredNamePart, StringComparison.OrdinalIgnoreCase))
            .OrderBy(path => Path.GetFileName(path), StringComparer.OrdinalIgnoreCase)
            .FirstOrDefault();
    }

    private static bool IsExcelExtension(string extension)
    {
        return extension.Equals(".xls", StringComparison.OrdinalIgnoreCase) ||
               extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) ||
               extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase);
    }

    private static string? FindFile(IReadOnlyList<string> sourceFiles, string[] extensions, string requiredNamePart)
    {
        return sourceFiles
            .Where(path => extensions.Contains(Path.GetExtension(path), StringComparer.OrdinalIgnoreCase))
            .Where(path => Path.GetFileNameWithoutExtension(path).Contains(requiredNamePart, StringComparison.OrdinalIgnoreCase))
            .OrderBy(path => Path.GetFileName(path), StringComparer.OrdinalIgnoreCase)
            .FirstOrDefault();
    }

    private static decimal Number(IXLCell cell)
    {
        if (cell.TryGetValue<decimal>(out var value))
        {
            return value;
        }

        return ParsingHelpers.ParseDecimal(cell.GetFormattedString()) ?? 0m;
    }

    private static string NormalizeCategory(string coverage, string category)
    {
        var cleaned = ParsingHelpers.Clean(category);
        if (coverage == CoverageNames.Ehc)
        {
            if (cleaned.Contains("drug", StringComparison.OrdinalIgnoreCase) ||
                cleaned.Contains("pharmacy", StringComparison.OrdinalIgnoreCase))
            {
                return "Drugs";
            }

            if (cleaned.Contains("paramedical", StringComparison.OrdinalIgnoreCase))
            {
                return "Paramedical";
            }

            if (cleaned.Contains("vision", StringComparison.OrdinalIgnoreCase))
            {
                return "Vision";
            }

            if (cleaned.Contains("hospital", StringComparison.OrdinalIgnoreCase))
            {
                return "Hospital";
            }

            if (cleaned.Contains("medical", StringComparison.OrdinalIgnoreCase) ||
                cleaned.Contains("supplies", StringComparison.OrdinalIgnoreCase))
            {
                return "Medical Services and Supplies";
            }

            if (cleaned.Contains("country", StringComparison.OrdinalIgnoreCase))
            {
                return "Out of Country";
            }

            return string.IsNullOrWhiteSpace(cleaned) ? "Other EHC" : cleaned;
        }

        if (cleaned.Contains("basic", StringComparison.OrdinalIgnoreCase))
        {
            return "Basic Dental";
        }

        if (cleaned.Contains("endo", StringComparison.OrdinalIgnoreCase))
        {
            return "Endodontics";
        }

        if (cleaned.Contains("perio", StringComparison.OrdinalIgnoreCase))
        {
            return "Periodontics";
        }

        if (cleaned.Contains("major", StringComparison.OrdinalIgnoreCase))
        {
            return "Major Dental";
        }

        if (cleaned.Contains("ortho", StringComparison.OrdinalIgnoreCase))
        {
            return "Orthodontics";
        }

        return string.IsNullOrWhiteSpace(cleaned) ? "Other Dental" : cleaned;
    }

    private static string NormalizeParamedicalService(string service)
    {
        var cleaned = ParsingHelpers.Clean(service);
        if (cleaned.Contains("massage", StringComparison.OrdinalIgnoreCase))
        {
            return "Massage";
        }

        if (cleaned.Contains("counsel", StringComparison.OrdinalIgnoreCase))
        {
            return "Clinical Counsellor";
        }

        if (cleaned.Contains("physio", StringComparison.OrdinalIgnoreCase))
        {
            return "Physiotherapy";
        }

        if (cleaned.Contains("chiro", StringComparison.OrdinalIgnoreCase))
        {
            return "Chiropractor";
        }

        if (cleaned.Contains("acupunct", StringComparison.OrdinalIgnoreCase))
        {
            return "Acupuncture";
        }

        if (cleaned.Contains("naturopath", StringComparison.OrdinalIgnoreCase))
        {
            return "Naturopath";
        }

        if (cleaned.Contains("psychologist", StringComparison.OrdinalIgnoreCase))
        {
            return "Psychologist";
        }

        return string.IsNullOrWhiteSpace(cleaned) ? "Other Paramedical" : cleaned;
    }

    private static DateOnly? TryParseMonthLabel(string value)
    {
        foreach (var format in MonthFormats)
        {
            if (DateTime.TryParseExact(value, format, Culture, DateTimeStyles.None, out var date))
            {
                return new DateOnly(date.Year, date.Month, 1);
            }
        }

        return null;
    }

    private static bool TryParseLongDate(string value, out DateTime date)
    {
        return DateTime.TryParseExact(
            value.Trim(),
            ["MMM d, yyyy", "MMMM d, yyyy", "MMM d yyyy", "MMMM d yyyy"],
            Culture,
            DateTimeStyles.None,
            out date);
    }

    private static bool TryParseCompactDate(string value, out DateTime date)
    {
        return DateTime.TryParseExact(
            value.Trim(),
            "ddMMMyyyy",
            Culture,
            DateTimeStyles.None,
            out date);
    }

    private static DateOnly FirstDayOfMonth(DateOnly date)
    {
        return new DateOnly(date.Year, date.Month, 1);
    }

    private static DateOnly EndOfMonth(DateOnly date)
    {
        return new DateOnly(date.Year, date.Month, DateTime.DaysInMonth(date.Year, date.Month));
    }

    private static DateOnly MinDate(DateOnly left, DateOnly right)
    {
        return left <= right ? left : right;
    }

    private static string FormatPeriodLabel(DateOnly start, DateOnly end)
    {
        return $"{start.ToString("MMM yyyy", Culture)} - {end.ToString("MMM yyyy", Culture)}";
    }

    private static DateTime? ToDateTime(DateOnly? date)
    {
        return date?.ToDateTime(TimeOnly.MinValue);
    }

    private static string FirstNonBlank(params string?[] values)
    {
        return values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value))?.Trim() ?? "";
    }

    private static string UniqueDestinationPath(string folder, string fileName)
    {
        var stem = Path.GetFileNameWithoutExtension(fileName);
        var extension = Path.GetExtension(fileName);
        var candidate = Path.Combine(folder, fileName);
        var index = 2;

        while (File.Exists(candidate))
        {
            candidate = Path.Combine(folder, $"{stem} ({index}){extension}");
            index++;
        }

        return candidate;
    }

    private static int NormalizeReportMonthCount(int monthCount)
    {
        return monthCount is >= 1 and <= 36 ? monthCount : 12;
    }

    private static string CurrentPeriodId(int monthCount)
    {
        return $"CURR_{NormalizeReportMonthCount(monthCount)}";
    }

    private static string PriorPeriodId(int monthCount)
    {
        return $"PRIOR_{NormalizeReportMonthCount(monthCount)}";
    }

    private static string BenchmarkPeriodId(int monthCount)
    {
        return NormalizeReportMonthCount(monthCount) == 12
            ? PeriodIds.PriorPrior12
            : PeriodIds.Prior12;
    }

    private static class CoverageNames
    {
        public const string Ehc = "EHC";
        public const string Dental = "Dental";
        public const string Hsa = "HSA";
        public const string AsoFees = "ASO Fees";
    }

    private static class PeriodIds
    {
        public const string Current12 = "CURR_12";
        public const string PriorPrior12 = "PRIOR_PRIOR_12";
        public const string Prior12 = "PRIOR_12";
        public const string Current9 = "CURR_9";
        public const string Prior9 = "PRIOR_9";
        public const string Current6 = "CURR_6";
        public const string Prior6 = "PRIOR_6";
        public const string CustomCurrent = "CUSTOM_CURRENT";
        public const string CustomPrior = "CUSTOM_PRIOR";
    }

    private enum SheetKind
    {
        Input,
        Control,
        Formula,
        Export
    }

    private sealed class ClaimsExperienceDataModelData
    {
        public string ClientName { get; set; } = "";
        public string Carrier { get; set; } = "";
        public ClaimsExperienceRenewalTerms RenewalTerms { get; set; } = new();
        public string PolicyNumber { get; set; } = "";
        public int ReportMonthCount { get; set; } = 12;
        public int RequestedReportMonthCount { get; set; }
        public int DetectedReportMonthCount { get; set; }
        public string CurrentSourcePeriodId { get; set; } = PeriodIds.Current12;
        public string PriorSourcePeriodId { get; set; } = PeriodIds.Prior12;
        public string BenchmarkSourcePeriodId { get; set; } = PeriodIds.PriorPrior12;
        public DateOnly? PeriodStart { get; set; }
        public DateOnly? PeriodEnd { get; set; }
        public string PeriodLabel { get; set; } = "";
        public int? EmployeeCount { get; set; }
        public ClaimsExperienceTotals? Totals { get; set; }
        public List<ClaimsImportRecord> ImportedFiles { get; } = [];
        public List<ClaimsExperienceMonth> Months { get; } = [];
        public List<ClaimsExperienceMonth> PriorMonths { get; } = [];
        public List<ClaimsExperienceMonth> PriorPriorMonths { get; } = [];
        public Dictionary<DateOnly, BillingMonthlyLives> BillingMonthlyLivesByMonth { get; } = [];
        public List<BillingInsuredLivesDetailRow> BillingInsuredLivesDetails { get; } = [];
        public List<ClaimsCategoryDetail> EhcCategories { get; } = [];
        public List<ParamedicalDetail> ParamedicalDetails { get; } = [];
        public List<ClaimsCategoryDetail> DentalCategories { get; } = [];
        public List<ClaimsSourcePeriod> SourcePeriods { get; } = [];
        public List<ClaimsServiceAdjudicationFact> ServiceAdjudicationFacts { get; } = [];
        public List<ClaimsServiceControlTotal> ServiceControlTotals { get; } = [];
        public List<RenewalExperienceFinancial> RenewalExperienceFinancials { get; } = [];
        public string ScheduleOfBenefitsSourcePath { get; set; } = "";
        public string ScheduleOfBenefitsClientName { get; set; } = "";
        public List<ScheduleOfBenefitsRow> ScheduleOfBenefitsRows { get; } = [];
    }

    private sealed record RenewalExperienceDatedRow(
        SunLifeExperienceRow Row,
        DateOnly Start,
        DateOnly End);

    private sealed record RenewalExperienceCandidate(
        string Path,
        CarrierPdfParseResult Parsed,
        List<RenewalExperienceDatedRow> Rows,
        DateOnly LatestEnd);

    private sealed record PremiumClaimsComparisonSelection(
        int SortOrder,
        string Role,
        string PeriodId,
        string PeriodLabel,
        RenewalExperienceFinancial? ReportedFinancial);

    private sealed class RenewalExperienceFinancial
    {
        public string Benefit { get; set; } = "";
        public DateOnly Start { get; set; }
        public DateOnly End { get; set; }
        public int? AverageCoveredEmployees { get; set; }
        public decimal? BilledPremium { get; set; }
        public decimal? AdjustedPremium { get; set; }
        public decimal? PaidClaims { get; set; }
        public decimal? IncurredClaims { get; set; }
        public decimal? AdjustedLossRatio { get; set; }
        public decimal? Weighting { get; set; }
        public decimal PoolingPremium { get; set; }
        public string SourceFile { get; set; } = "";
        public string SourceEvidence { get; set; } = "";
    }

    private sealed class ClaimsImportRecord
    {
        public DateTime ImportedAt { get; set; }
        public string SourceType { get; set; } = "";
        public string SourceFile { get; set; } = "";
        public string FullPath { get; set; } = "";
        public string RowsRead { get; set; } = "";
        public string Status { get; set; } = "";
        public string Notes { get; set; } = "";
    }

    private sealed class ClaimsExperienceTotals
    {
        public decimal EhcPremium { get; set; }
        public decimal EhcClaims { get; set; }
        public decimal DentalPremium { get; set; }
        public decimal DentalClaims { get; set; }
    }

    private sealed class ClaimsExperiencePdfRow
    {
        public string MonthLabel { get; set; } = "";
        public decimal EhcPremium { get; set; }
        public decimal EhcClaims { get; set; }
        public decimal DentalPremium { get; set; }
        public decimal DentalClaims { get; set; }
        public string SourceFile { get; set; } = "";
    }

    private sealed class ClaimsExperienceMonth
    {
        public DateOnly Month { get; set; }
        public decimal EhcPremium { get; set; }
        public decimal EhcClaims { get; set; }
        public decimal DentalPremium { get; set; }
        public decimal DentalClaims { get; set; }
        public string SourceFile { get; set; } = "";
        public string EhcSourceFile { get; set; } = "";
        public string DentalSourceFile { get; set; } = "";
        public bool EhcPremiumMatchesMonth { get; set; }
        public bool DentalPremiumMatchesMonth { get; set; }
        public bool IsAnnualSummary { get; set; }
    }

    private sealed class BillingInsuredLivesExtraction
    {
        public List<BillingMonthlyLives> monthly_lives { get; } = [];
        public List<BillingInsuredLivesDetailRow> detail_rows { get; } = [];
    }

    private sealed class BillingMonthlyLives
    {
        public DateOnly Month { get; set; }
        public int? HealthInsuredLives { get; set; }
        public int? DentalInsuredLives { get; set; }
        public decimal? HealthPremium { get; set; }
        public decimal? DentalPremium { get; set; }
        public string SourceFile { get; set; } = "";
        public DateTime? DatePrepared { get; set; }
        public DateTime FileModified { get; set; }
    }

    private sealed class BillingInsuredLivesDetailRow
    {
        public string SourceFile { get; set; } = "";
        public string StatementPeriod { get; set; } = "";
        public DateOnly? StatementMonth { get; set; }
        public string ClientName { get; set; } = "";
        public string GroupNumber { get; set; } = "";
        public string Benefit { get; set; } = "";
        public string DepCode { get; set; } = "";
        public string Option { get; set; } = "";
        public int Lives { get; set; }
        public decimal Premium { get; set; }
        public string ExtractedCategory { get; set; } = "";
        public bool IsExperienceReportFallback { get; set; }
    }

    private sealed class BillingFileExtraction
    {
        public string SourceFile { get; set; } = "";
        public string FullPath { get; set; } = "";
        public DateOnly? StatementMonth { get; set; }
        public string StatementPeriod { get; set; } = "";
        public string ClientName { get; set; } = "";
        public string GroupNumber { get; set; } = "";
        public int? HealthInsuredLives { get; set; }
        public int? DentalInsuredLives { get; set; }
        public decimal? HealthPremium { get; set; }
        public decimal? DentalPremium { get; set; }
        public DateTime? DatePrepared { get; set; }
        public DateTime FileModified { get; set; }
        public List<BillingInsuredLivesDetailRow> DetailRows { get; } = [];
    }

    private sealed record BillingHeaderMap(int HeaderRow, int BenefitColumn, int LivesColumn, int? DepCodeColumn, int? OptionColumn, int? PremiumColumn, int? TotalPremiumAdjustmentsColumn);

    private sealed class ClaimsCategorySourceRow
    {
        public string Coverage { get; set; } = "";
        public string Benefit { get; set; } = "";
        public string Category { get; set; } = "";
        public string Service { get; set; } = "";
        public decimal Occurrences { get; set; }
        public decimal Submitted { get; set; }
        public decimal Eligible { get; set; }
        public decimal Paid { get; set; }
        public string SourceFile { get; set; } = "";
    }

    private sealed class ClaimsCategoryDetail
    {
        public string PeriodId { get; set; } = "";
        public string Coverage { get; set; } = "";
        public string Category { get; set; } = "";
        public int? ClaimCount { get; set; }
        public decimal SubmittedAmount { get; set; }
        public decimal EligibleAmount { get; set; }
        public decimal PaidClaims { get; set; }
        public string SourceFile { get; set; } = "";
        public string Notes { get; set; } = "";
    }

    private sealed class ParamedicalDetail
    {
        public string PeriodId { get; set; } = "";
        public string Service { get; set; } = "";
        public int? ClaimCount { get; set; }
        public decimal SubmittedAmount { get; set; }
        public decimal EligibleAmount { get; set; }
        public decimal PaidClaims { get; set; }
        public string SourceFile { get; set; } = "";
        public string Notes { get; set; } = "";
    }

    private sealed class ScheduleOfBenefitsRow
    {
        public int SlidePage { get; set; }
        public int PageTotal { get; set; } = 3;
        public int SectionOrder { get; set; }
        public string ClientName { get; set; } = "";
        public string Carrier { get; set; } = "";
        public string ClassName { get; set; } = "";
        public string SectionName { get; set; } = "";
        public string DetailText { get; set; } = "";
        public string SourceFile { get; set; } = "";
        public string Notes { get; set; } = "";
    }

    private sealed record ChartCategoryValue(string Category, decimal Value);
}

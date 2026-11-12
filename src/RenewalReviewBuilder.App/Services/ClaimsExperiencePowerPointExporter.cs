using System.Globalization;
using System.IO.Compression;
using System.Xml.Linq;
using ClosedXML.Excel;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed class ClaimsExperiencePowerPointExporter
{
    private static readonly CultureInfo Culture = CultureInfo.InvariantCulture;
    private static readonly ParamedicalServiceMap[] ParamedicalBreakdownServices =
    [
        new("Acupuncture", ["Acupuncture"]),
        new("Chiropodist", ["Chiropodist"]),
        new("Chiropractor", ["Chiropractor"]),
        new("Clinical Counsellor", ["Clinical Counsellor", "Clinical Counselor", "Counsellor", "Counselor"]),
        new("Dietitian", ["Dietitian", "Dietician"]),
        new("Marriage / Family Therapist", ["Marriage", "Family Therapist"]),
        new("Naturopath", ["Naturopath"]),
        new("Osteopath", ["Osteopath"]),
        new("Physio", ["Physiotherapy", "Physio"]),
        new("Podiatrist", ["Podiatrist"]),
        new("Psycho-analyst", ["Psycho-analyst", "Psychoanalyst"]),
        new("Psychologist", ["Psychologist"]),
        new("Massage", ["Massage"]),
        new("Social Worker", ["Social Worker"]),
        new("Speech Therapy", ["Speech Therapy", "Speech Therapist", "Speech"])
    ];
    private static readonly string[] DentalBreakdownCategoryOrder =
    [
        "Basic Dental",
        "Periodontics",
        "Major Dental",
        "Orthodontics",
        "Endodontics",
        "Dentures",
        "Oral Surgery"
    ];

    public Task<ClaimsExperiencePowerPointResult> ExportAsync(ClaimsExperiencePowerPointRequest request)
    {
        return Task.Run(() => Export(request));
    }

    public ClaimsExperiencePowerPointResult Export(ClaimsExperiencePowerPointRequest request)
    {
        if (!File.Exists(request.WorkbookPath))
        {
            throw new FileNotFoundException("Claims experience model workbook was not found.", request.WorkbookPath);
        }

        if (!File.Exists(request.TemplatePath))
        {
            throw new FileNotFoundException("Claims experience PowerPoint template was not found.", request.TemplatePath);
        }

        var outputFolder = string.IsNullOrWhiteSpace(request.OutputFolder)
            ? Path.GetDirectoryName(request.WorkbookPath) ?? AppContext.BaseDirectory
            : request.OutputFolder;
        Directory.CreateDirectory(outputFolder);

        var averageSpendData = ReadAverageSpendData(request.WorkbookPath);
        var monthlyBreakdownData = ReadMonthlyBreakdownData(request.WorkbookPath);
        var claimsDistributionData = ReadClaimsDistributionData(request.WorkbookPath);
        var paramedicalBreakdownData = ReadParamedicalBreakdownData(request.WorkbookPath);
        var dentalBreakdownData = ReadDentalBreakdownData(request.WorkbookPath);
        var premiumClaimsSummaryData = ReadPremiumClaimsSummaryData(request.WorkbookPath);
        var globalPlaceholderData = ReadDeckGlobalPlaceholderData(request.WorkbookPath);
        var scheduleOfBenefitsData = ReadScheduleOfBenefitsData(request.WorkbookPath);
        var outputPath = UniqueDestinationPath(outputFolder, "Claims_Experience_Report.pptx");
        File.Copy(request.TemplatePath, outputPath, overwrite: true);

        using var archive = ZipFile.Open(outputPath, ZipArchiveMode.Update);
        var slideCountUpdated = 0;
        slideCountUpdated += UpdateGlobalDeckPlaceholders(archive, globalPlaceholderData);
        if (UpdatePerformanceSnapshotMetrics(archive, globalPlaceholderData))
        {
            slideCountUpdated++;
        }

        if (UpdateAverageSpendSlide(archive, averageSpendData))
        {
            slideCountUpdated++;
        }

        if (UpdateMonthlyBreakdownSlide(archive, monthlyBreakdownData.Health))
        {
            slideCountUpdated++;
        }

        if (UpdateMonthlyBreakdownSlide(archive, monthlyBreakdownData.Dental))
        {
            slideCountUpdated++;
        }

        if (UpdateClaimsDistributionSlide(archive, claimsDistributionData))
        {
            slideCountUpdated++;
        }

        if (UpdateParamedicalBreakdownSlide(archive, paramedicalBreakdownData))
        {
            slideCountUpdated++;
        }

        if (UpdateDentalBreakdownSlide(archive, dentalBreakdownData))
        {
            slideCountUpdated++;
        }

        if (UpdateOverallPremiumClaimsSummarySlide(archive, premiumClaimsSummaryData.Overall))
        {
            slideCountUpdated++;
        }

        if (UpdatePremiumClaimsSummarySlide(archive, premiumClaimsSummaryData.Health))
        {
            slideCountUpdated++;
        }

        if (UpdatePremiumClaimsSummarySlide(archive, premiumClaimsSummaryData.Dental))
        {
            slideCountUpdated++;
        }

        slideCountUpdated += UpdateScheduleOfBenefitsSlides(archive, scheduleOfBenefitsData);

        var averageSpendChartPath = FindChartPathsBySlideTitle(archive, "Average Amount Spent by Employee").FirstOrDefault();
        if (averageSpendChartPath is not null)
        {
            UpdateAverageSpendChartXml(archive, averageSpendChartPath, averageSpendData);

            var averageSpendEmbeddedWorkbookPath = FindEmbeddedWorkbookPath(archive, averageSpendChartPath);
            if (averageSpendEmbeddedWorkbookPath is not null)
            {
                UpdateEmbeddedAverageSpendWorkbook(archive, averageSpendEmbeddedWorkbookPath, averageSpendData);
            }
        }

        var claimsDistributionChartPath = FindChartPathsBySlideTitle(archive, "Program Performance: Claims Distribution").FirstOrDefault();
        if (claimsDistributionChartPath is not null)
        {
            UpdateClaimsDistributionChartXml(archive, claimsDistributionChartPath, claimsDistributionData);

            var claimsDistributionEmbeddedWorkbookPath = FindEmbeddedWorkbookPath(archive, claimsDistributionChartPath);
            if (claimsDistributionEmbeddedWorkbookPath is not null)
            {
                UpdateEmbeddedClaimsDistributionWorkbook(archive, claimsDistributionEmbeddedWorkbookPath, claimsDistributionData);
            }
        }

        var paramedicalBreakdownChartPath = FindChartPathsBySlideTitle(archive, "Claims Breakdown: Paramedical").FirstOrDefault();
        if (paramedicalBreakdownChartPath is not null)
        {
            UpdateParamedicalBreakdownChartXml(archive, paramedicalBreakdownChartPath, paramedicalBreakdownData);

            var paramedicalBreakdownEmbeddedWorkbookPath = FindEmbeddedWorkbookPath(archive, paramedicalBreakdownChartPath);
            if (paramedicalBreakdownEmbeddedWorkbookPath is not null)
            {
                UpdateEmbeddedParamedicalBreakdownWorkbook(archive, paramedicalBreakdownEmbeddedWorkbookPath, paramedicalBreakdownData);
            }
        }

        var dentalBreakdownChartPath = FindChartPathsBySlideTitle(archive, "Claims Breakdown: Dental").FirstOrDefault();
        if (dentalBreakdownChartPath is not null)
        {
            UpdateDentalBreakdownChartXml(archive, dentalBreakdownChartPath, dentalBreakdownData);

            var dentalBreakdownEmbeddedWorkbookPath = FindEmbeddedWorkbookPath(archive, dentalBreakdownChartPath);
            if (dentalBreakdownEmbeddedWorkbookPath is not null)
            {
                UpdateEmbeddedDentalBreakdownWorkbook(archive, dentalBreakdownEmbeddedWorkbookPath, dentalBreakdownData);
            }
        }

        var overallSummaryChartPath = FindChartPathsBySlideTextParts(archive, "Program Performance: Premium vs Claims Summary").FirstOrDefault();
        if (overallSummaryChartPath is not null)
        {
            UpdateOverallPremiumClaimsSummaryChartXml(archive, overallSummaryChartPath, premiumClaimsSummaryData.Overall);

            var overallSummaryEmbeddedWorkbookPath = FindEmbeddedWorkbookPath(archive, overallSummaryChartPath);
            if (overallSummaryEmbeddedWorkbookPath is not null)
            {
                UpdateEmbeddedOverallPremiumClaimsSummaryWorkbook(archive, overallSummaryEmbeddedWorkbookPath, premiumClaimsSummaryData.Overall);
            }
        }

        var healthSummaryChartPath = FindChartPathsBySlideTextParts(archive, "Premiums vs. Claims Summary", "Extended Healthcare").FirstOrDefault();
        if (healthSummaryChartPath is not null)
        {
            UpdatePremiumClaimsSummaryChartXml(archive, healthSummaryChartPath, premiumClaimsSummaryData.Health);

            var healthSummaryEmbeddedWorkbookPath = FindEmbeddedWorkbookPath(archive, healthSummaryChartPath);
            if (healthSummaryEmbeddedWorkbookPath is not null)
            {
                UpdateEmbeddedPremiumClaimsSummaryWorkbook(archive, healthSummaryEmbeddedWorkbookPath, premiumClaimsSummaryData.Health);
            }
        }

        var dentalSummaryChartPath = FindChartPathsBySlideTextParts(archive, "Premiums vs. Claims Summary", "Dental Care").FirstOrDefault();
        if (dentalSummaryChartPath is not null)
        {
            UpdatePremiumClaimsSummaryChartXml(archive, dentalSummaryChartPath, premiumClaimsSummaryData.Dental);

            var dentalSummaryEmbeddedWorkbookPath = FindEmbeddedWorkbookPath(archive, dentalSummaryChartPath);
            if (dentalSummaryEmbeddedWorkbookPath is not null)
            {
                UpdateEmbeddedPremiumClaimsSummaryWorkbook(archive, dentalSummaryEmbeddedWorkbookPath, premiumClaimsSummaryData.Dental);
            }
        }

        var healthComparisonBreakEvenChartPath = FindChartPathsBySlideTextParts(archive, "Premium Vs Claims over 3 periods", "Extended Health Care", "Net L/R").FirstOrDefault()
                                                 ?? FindChartPathsBySlideTextParts(archive, "Comparison Chart", "Extended Health Care", "Net L/R").FirstOrDefault();
        var dentalComparisonBreakEvenChartPath = FindChartPathsBySlideTextParts(archive, "Premium Vs Claims over 3 periods", "Dental Care", "Net L/R").FirstOrDefault();
        var healthComparisonChartPath = FindChartPathsBySlideTextParts(archive, "Premium Vs Claims over 3 periods", "Extended Health Care")
            .Concat(FindChartPathsBySlideTextParts(archive, "Comparison Chart", "Extended Health Care"))
            .Where(path => !path.Equals(healthComparisonBreakEvenChartPath, StringComparison.OrdinalIgnoreCase))
            .FirstOrDefault();
        var dentalComparisonChartPath = FindChartPathsBySlideTextParts(archive, "Premium Vs Claims over 3 periods", "Dental Care")
            .Where(path => !path.Equals(dentalComparisonBreakEvenChartPath, StringComparison.OrdinalIgnoreCase))
            .FirstOrDefault();
        if (healthComparisonChartPath is not null)
        {
            UpdatePremiumClaimsSummaryChartXml(archive, healthComparisonChartPath, premiumClaimsSummaryData.HealthComparison);
            var embeddedWorkbookPath = FindEmbeddedWorkbookPath(archive, healthComparisonChartPath);
            if (embeddedWorkbookPath is not null)
            {
                UpdateEmbeddedPremiumClaimsSummaryWorkbook(archive, embeddedWorkbookPath, premiumClaimsSummaryData.HealthComparison);
            }
        }

        if (dentalComparisonChartPath is not null)
        {
            UpdatePremiumClaimsSummaryChartXml(archive, dentalComparisonChartPath, premiumClaimsSummaryData.DentalComparison);
            var embeddedWorkbookPath = FindEmbeddedWorkbookPath(archive, dentalComparisonChartPath);
            if (embeddedWorkbookPath is not null)
            {
                UpdateEmbeddedPremiumClaimsSummaryWorkbook(archive, embeddedWorkbookPath, premiumClaimsSummaryData.DentalComparison);
            }
        }

        if (healthComparisonBreakEvenChartPath is not null)
        {
            UpdateLossRatioBreakEvenChartXml(archive, healthComparisonBreakEvenChartPath, premiumClaimsSummaryData.HealthLossRatioBreakEven);
            var embeddedWorkbookPath = FindEmbeddedWorkbookPath(archive, healthComparisonBreakEvenChartPath);
            if (embeddedWorkbookPath is not null)
            {
                UpdateEmbeddedLossRatioBreakEvenWorkbook(archive, embeddedWorkbookPath, premiumClaimsSummaryData.HealthLossRatioBreakEven);
            }
        }

        if (dentalComparisonBreakEvenChartPath is not null)
        {
            UpdateLossRatioBreakEvenChartXml(archive, dentalComparisonBreakEvenChartPath, premiumClaimsSummaryData.DentalLossRatioBreakEven);
            var embeddedWorkbookPath = FindEmbeddedWorkbookPath(archive, dentalComparisonBreakEvenChartPath);
            if (embeddedWorkbookPath is not null)
            {
                UpdateEmbeddedLossRatioBreakEvenWorkbook(archive, embeddedWorkbookPath, premiumClaimsSummaryData.DentalLossRatioBreakEven);
            }
        }

        SanitizeAllChartParts(archive);

        return new ClaimsExperiencePowerPointResult
        {
            OutputFolder = outputFolder,
            PowerPointPath = outputPath,
            SourceWorkbookPath = request.WorkbookPath,
            SlideCountUpdated = slideCountUpdated
        };
    }

    private static AverageSpendSlideData ReadAverageSpendData(string workbookPath)
    {
        using var workbook = OpenWorkbookReadOnly(workbookPath);
        var monthly = ReadMonthlyExperienceRows(workbook).ToList();
        var currentPeriodId = ReadReportControlValue(workbook, "Current Period ID", "CURR_12");
        var priorPeriodId = ReadReportControlValue(workbook, "Prior Period ID", "PRIOR_12");
        var monthCount = ReadReportControlNumber(workbook, "Report Month Count") ?? 12m;

        var currentLabel = ReadPeriodLabel(workbook, currentPeriodId);
        var priorLabel = ReadPeriodLabel(workbook, priorPeriodId);

        var currentHealthClaims = SumClaims(monthly, currentPeriodId, "EHC");
        var currentDentalClaims = SumClaims(monthly, currentPeriodId, "Dental");
        var priorHealthClaims = SumClaims(monthly, priorPeriodId, "EHC");
        var priorDentalClaims = SumClaims(monthly, priorPeriodId, "Dental");

        var currentHealthLives = AverageLives(monthly, currentPeriodId, "EHC");
        var currentDentalLives = AverageLives(monthly, currentPeriodId, "Dental");
        var priorHealthLives = AverageLives(monthly, priorPeriodId, "EHC");
        var priorDentalLives = AverageLives(monthly, priorPeriodId, "Dental");
        var currentCombinedLives = AverageCombinedLives(monthly, currentPeriodId);
        var priorCombinedLives = AverageCombinedLives(monthly, priorPeriodId);

        var rows = new List<AverageSpendRow>
        {
            new("Extended Health", SafeDivide(currentHealthClaims, currentHealthLives), SafeDivide(priorHealthClaims, priorHealthLives)),
            new("Dental", SafeDivide(currentDentalClaims, currentDentalLives), SafeDivide(priorDentalClaims, priorDentalLives))
        };
        rows.Add(new AverageSpendRow(
            "Combined Total",
            SafeDivide(currentHealthClaims + currentDentalClaims, currentCombinedLives),
            SafeDivide(priorHealthClaims + priorDentalClaims, priorCombinedLives)));

        var highestMonth = monthly
            .Where(row => row.PeriodId.Equals(currentPeriodId, StringComparison.OrdinalIgnoreCase))
            .GroupBy(row => row.Month)
            .Select(group => new
            {
                Month = group.Key,
                Claims = group.Sum(row => row.PaidClaims)
            })
            .OrderByDescending(row => row.Claims)
            .FirstOrDefault();

        return new AverageSpendSlideData(
            currentLabel,
            priorLabel,
            (int)Math.Round(monthCount, MidpointRounding.AwayFromZero),
            currentHealthLives,
            currentDentalLives,
            rows,
            highestMonth?.Month,
            highestMonth?.Claims ?? 0m);
    }

    private static MonthlyBreakdownsSlideData ReadMonthlyBreakdownData(string workbookPath)
    {
        using var workbook = OpenWorkbookReadOnly(workbookPath);
        var monthly = ReadMonthlyExperienceRows(workbook).ToList();
        var currentPeriodId = ReadReportControlValue(workbook, "Current Period ID", "CURR_12");
        var priorPeriodId = ReadReportControlValue(workbook, "Prior Period ID", "PRIOR_12");
        var ehcPoolingCharge = ReadAssumptionNumber(workbook, "EHC Pooling Charge") ?? 0.09m;

        return new MonthlyBreakdownsSlideData(
            BuildMonthlyBreakdownSlideData(
                workbook,
                monthly,
                currentPeriodId,
                priorPeriodId,
                "Monthly Breakdown: Extended Health",
                "EHC",
                "tblEHCCategory",
                "Health",
                ehcPoolingCharge),
            BuildMonthlyBreakdownSlideData(
                workbook,
                monthly,
                currentPeriodId,
                priorPeriodId,
                "Monthly Breakdown: Dental",
                "Dental",
                "tblDentalCategory",
                "Dental",
                ehcPoolingCharge));
    }

    private static ClaimsDistributionSlideData ReadClaimsDistributionData(string workbookPath)
    {
        using var workbook = OpenWorkbookReadOnly(workbookPath);
        var currentPeriodId = ReadReportControlValue(workbook, "Current Period ID", "CURR_12");
        var priorPeriodId = ReadReportControlValue(workbook, "Prior Period ID", "PRIOR_12");
        var currentPeriodLabel = ReadPeriodLabel(workbook, currentPeriodId);
        var priorPeriodLabel = ReadPeriodLabel(workbook, priorPeriodId);

        var currentDental = ReadCategoryAmounts(workbook, "tblDentalCategory", currentPeriodId);
        var priorDental = ReadCategoryAmounts(workbook, "tblDentalCategory", priorPeriodId);
        var currentEhc = ReadCategoryAmounts(workbook, "tblEHCCategory", currentPeriodId);
        var priorEhc = ReadCategoryAmounts(workbook, "tblEHCCategory", priorPeriodId);

        var rows = new List<ClaimsDistributionRow>
        {
            new("Dental", currentDental.Values.Sum(), priorDental.Values.Sum()),
            new("Paramedical", SumMatchingCategory(currentEhc, "Paramedical"), SumMatchingCategory(priorEhc, "Paramedical")),
            new("Drugs", SumMatchingCategory(currentEhc, "Drugs"), SumMatchingCategory(priorEhc, "Drugs")),
            new("Vision", SumMatchingCategory(currentEhc, "Vision"), SumMatchingCategory(priorEhc, "Vision")),
            new("Other EHC", OtherEhcClaims(currentEhc), OtherEhcClaims(priorEhc))
        };

        return new ClaimsDistributionSlideData(currentPeriodLabel, priorPeriodLabel, rows);
    }

    private static ParamedicalBreakdownSlideData ReadParamedicalBreakdownData(string workbookPath)
    {
        using var workbook = OpenWorkbookReadOnly(workbookPath);
        var currentPeriodId = ReadReportControlValue(workbook, "Current Period ID", "CURR_12");
        var priorPeriodId = ReadReportControlValue(workbook, "Prior Period ID", "PRIOR_12");
        var currentPeriodLabel = ReadPeriodLabel(workbook, currentPeriodId);
        var priorPeriodLabel = ReadPeriodLabel(workbook, priorPeriodId);

        var currentServices = ReadParamedicalServiceAmounts(workbook, currentPeriodId);
        var priorServices = ReadParamedicalServiceAmounts(workbook, priorPeriodId);
        var currentMatched = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var priorMatched = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        var rows = ParamedicalBreakdownServices
            .Select(service => new ParamedicalBreakdownRow(
                service.Label,
                SumAndMarkMatchingServices(currentServices, service.SearchTerms, currentMatched),
                SumAndMarkMatchingServices(priorServices, service.SearchTerms, priorMatched)))
            .Where(row => row.CurrentPaidClaims != 0m || row.PriorPaidClaims != 0m)
            .ToList();

        var otherCurrent = SumUnmatchedServices(currentServices, currentMatched);
        var otherPrior = SumUnmatchedServices(priorServices, priorMatched);
        if (otherCurrent != 0m || otherPrior != 0m)
        {
            rows.Add(new ParamedicalBreakdownRow("Other Paramedical", otherCurrent, otherPrior));
        }

        if (rows.Count == 0)
        {
            rows.Add(new ParamedicalBreakdownRow("No data", 0m, 0m));
        }

        return new ParamedicalBreakdownSlideData(currentPeriodLabel, priorPeriodLabel, rows);
    }

    private static DentalBreakdownSlideData ReadDentalBreakdownData(string workbookPath)
    {
        using var workbook = OpenWorkbookReadOnly(workbookPath);
        var currentPeriodId = ReadReportControlValue(workbook, "Current Period ID", "CURR_12");
        var priorPeriodId = ReadReportControlValue(workbook, "Prior Period ID", "PRIOR_12");
        var currentPeriodLabel = ReadPeriodLabel(workbook, currentPeriodId);
        var priorPeriodLabel = ReadPeriodLabel(workbook, priorPeriodId);
        var currentCategories = ReadCategoryAmounts(workbook, "tblDentalCategory", currentPeriodId);
        var priorCategories = ReadCategoryAmounts(workbook, "tblDentalCategory", priorPeriodId);
        var categoryNames = DentalBreakdownCategoryOrder
            .Concat(currentCategories.Keys)
            .Concat(priorCategories.Keys)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        var rows = categoryNames
            .Select(category => new DentalBreakdownRow(
                category,
                currentCategories.TryGetValue(category, out var currentPaidClaims) ? currentPaidClaims : 0m,
                priorCategories.TryGetValue(category, out var priorPaidClaims) ? priorPaidClaims : 0m))
            .Where(row => row.CurrentPaidClaims != 0m || row.PriorPaidClaims != 0m)
            .ToList();

        if (rows.Count == 0)
        {
            rows.Add(new DentalBreakdownRow("No data", 0m, 0m));
        }

        return new DentalBreakdownSlideData(currentPeriodLabel, priorPeriodLabel, rows);
    }

    private static PremiumClaimsSummarySlidesData ReadPremiumClaimsSummaryData(string workbookPath)
    {
        using var workbook = OpenWorkbookReadOnly(workbookPath);
        var monthly = ReadMonthlyExperienceRows(workbook).ToList();
        var monthCount = NormalizeReportMonthCount(ReadReportControlNumber(workbook, "Report Month Count") ?? 12m);
        var currentPeriodId = CurrentPeriodId(monthCount);
        var priorPeriodId = PriorPeriodId(monthCount);
        var benchmarkPeriodId = BenchmarkPeriodId(monthCount);

        var currentPeriodLabel = ReadPeriodLabel(workbook, currentPeriodId);
        var priorPeriodLabel = ReadPeriodLabel(workbook, priorPeriodId);
        var benchmarkPeriodLabel = ReadPeriodLabel(workbook, benchmarkPeriodId);
        var ehcPoolingCharge = ReadAssumptionNumber(workbook, "EHC Pooling Charge") ?? 0.09m;

        var legacyEhcCurrentRow = BuildPremiumClaimsRow(monthly, currentPeriodId, currentPeriodLabel, "EHC", useNetPremium: true, ehcPoolingCharge);
        var legacyDentalCurrentRow = BuildPremiumClaimsRow(monthly, currentPeriodId, currentPeriodLabel, "Dental", useNetPremium: false, ehcPoolingCharge);
        var canonicalHealthRows = ReadCanonicalPremiumClaimsRows(workbook, monthly, monthCount, "EHC", useNetPremium: true, ehcPoolingCharge);
        var canonicalDentalRows = ReadCanonicalPremiumClaimsRows(workbook, monthly, monthCount, "Dental", useNetPremium: false, ehcPoolingCharge);

        var healthComparisonRows = canonicalHealthRows.Count > 0
            ? canonicalHealthRows
            : BuildPremiumClaimsRows(
                monthly,
                [(benchmarkPeriodId, benchmarkPeriodLabel), (priorPeriodId, priorPeriodLabel), (currentPeriodId, currentPeriodLabel)],
                "EHC",
                useNetPremium: true,
                ehcPoolingCharge);
        var dentalComparisonRows = canonicalDentalRows.Count > 0
            ? canonicalDentalRows
            : BuildPremiumClaimsRows(
                monthly,
                [(benchmarkPeriodId, benchmarkPeriodLabel), (priorPeriodId, priorPeriodLabel), (currentPeriodId, currentPeriodLabel)],
                "Dental",
                useNetPremium: false,
                ehcPoolingCharge);
        var healthSummaryRows = BuildPremiumClaimsRows(
            monthly,
            [(priorPeriodId, priorPeriodLabel), (currentPeriodId, currentPeriodLabel)],
            "EHC",
            useNetPremium: true,
            ehcPoolingCharge);
        var dentalSummaryRows = BuildPremiumClaimsRows(
            monthly,
            [(priorPeriodId, priorPeriodLabel), (currentPeriodId, currentPeriodLabel)],
            "Dental",
            useNetPremium: false,
            ehcPoolingCharge);
        var ehcCurrentRow = healthComparisonRows.LastOrDefault() ?? legacyEhcCurrentRow;
        var dentalCurrentRow = dentalComparisonRows.LastOrDefault() ?? legacyDentalCurrentRow;

        var ehcPremium = ehcCurrentRow.Premium;
        var ehcPaidClaims = ehcCurrentRow.PaidClaims;
        var ehcLossRatio = UnderwritingEngine.CalculateLossRatioFactor(ehcPaidClaims, ehcPremium);
        var ehcTrend = ReadAssumptionNumber(workbook, "EHC Trend Factor") ?? 0.15m;
        var ehcTargetLossRatio = ReadAssumptionNumber(workbook, "EHC Break-even Point (TLR)") ?? 0.763m;
        var ehcTrendedLossRatio = UnderwritingEngine.CalculateTrendedLossRatioFactor(ehcLossRatio, ehcTrend);

        var dentalPremium = dentalCurrentRow.Premium;
        var dentalPaidClaims = dentalCurrentRow.PaidClaims;
        var dentalLossRatio = UnderwritingEngine.CalculateLossRatioFactor(dentalPaidClaims, dentalPremium);
        var dentalTrend = ReadAssumptionNumber(workbook, "Dental Trend Factor + Fee Guide") ?? 0.15m;
        var dentalTargetLossRatio = ReadAssumptionNumber(workbook, "Dental Break-even Point (TLR)") ?? 0.763m;
        var dentalTrendedLossRatio = UnderwritingEngine.CalculateTrendedLossRatioFactor(dentalLossRatio, dentalTrend);
        var healthSummary = new PremiumClaimsSummarySlideData(
            "Extended Healthcare",
            "Extended Healthcare",
            currentPeriodLabel,
            priorPeriodLabel,
            "Net Premium",
            "Net Premiums",
            ehcPremium,
            ehcPaidClaims,
            ehcLossRatio,
            ehcTrend,
            ehcTrendedLossRatio,
            ehcTargetLossRatio,
            UnderwritingEngine.CalculateRequiredChangeFactor(ehcTrendedLossRatio, ehcTargetLossRatio),
            ReadPremiumClaimsChartRows(workbook, "tblPPT_EHC_PremiumClaimsChart", "Net Premium", priorPeriodLabel, currentPeriodLabel, ehcPremium, ehcPaidClaims, fallbackRows: healthSummaryRows));
        var dentalSummary = new PremiumClaimsSummarySlideData(
            "Dental Care",
            "Dental Care",
            currentPeriodLabel,
            priorPeriodLabel,
            "Billed Premium",
            "Billed Premium",
            dentalPremium,
            dentalPaidClaims,
            dentalLossRatio,
            dentalTrend,
            dentalTrendedLossRatio,
            dentalTargetLossRatio,
            UnderwritingEngine.CalculateRequiredChangeFactor(dentalTrendedLossRatio, dentalTargetLossRatio),
            ReadPremiumClaimsChartRows(workbook, "tblPPT_Dental_PremiumClaimsChart", "Billed Premium", priorPeriodLabel, currentPeriodLabel, dentalPremium, dentalPaidClaims, fallbackRows: dentalSummaryRows));

        var healthComparison = healthSummary with
        {
            ChartRows = ReadPremiumClaimsChartRows(workbook, "tblPPT_EHC_ComparisonChart", "Net Premium", priorPeriodLabel, currentPeriodLabel, ehcPremium, ehcPaidClaims, maxRows: 4, fallbackRows: healthComparisonRows),
            PremiumSeriesLabel = "Net Premiums"
        };
        var dentalComparison = dentalSummary with
        {
            ChartRows = ReadPremiumClaimsChartRows(workbook, "tblPPT_Dental_ComparisonChart", "Billed Premium", priorPeriodLabel, currentPeriodLabel, dentalPremium, dentalPaidClaims, maxRows: 4, fallbackRows: dentalComparisonRows),
            PremiumSeriesLabel = "Billed Premium"
        };
        var healthLossRatioBreakEven = ReadLossRatioBreakEvenChartRows(workbook, "tblPPT_EHC_LossRatioBreakEvenChart", healthComparison.ChartRows, ehcTargetLossRatio);
        var dentalLossRatioBreakEven = ReadLossRatioBreakEvenChartRows(workbook, "tblPPT_Dental_LossRatioBreakEvenChart", dentalComparison.ChartRows, dentalTargetLossRatio);

        return new PremiumClaimsSummarySlidesData(
            BuildOverallPremiumClaimsSummary(currentPeriodLabel, healthSummary, dentalSummary),
            healthSummary,
            dentalSummary,
            healthComparison,
            dentalComparison,
            healthLossRatioBreakEven,
            dentalLossRatioBreakEven);
    }

    private static DeckGlobalPlaceholderData ReadDeckGlobalPlaceholderData(string workbookPath)
    {
        using var workbook = OpenWorkbookReadOnly(workbookPath);
        var monthCount = NormalizeReportMonthCount(ReadReportControlNumber(workbook, "Report Month Count") ?? 12m);
        var currentPeriodId = CurrentPeriodId(monthCount);
        var priorPeriodId = PriorPeriodId(monthCount);
        var clientName = ReadClientSetupValue(workbook, "Client Name", "Client Name");
        var carrier = ReadClientSetupValue(workbook, "Carrier", "");
        var policyNumber = ReadClientSetupValue(workbook, "Policy Number", "");
        var monthly = ReadMonthlyExperienceRows(workbook).ToList();
        var totalClaimsPaid = SumClaims(monthly, currentPeriodId, "EHC") + SumClaims(monthly, currentPeriodId, "Dental");
        var employeesCovered = AverageCombinedLives(monthly, currentPeriodId);

        return new DeckGlobalPlaceholderData(
            clientName,
            carrier,
            policyNumber,
            monthCount,
            ReadPeriodLabel(workbook, currentPeriodId),
            ReadPeriodLabel(workbook, priorPeriodId),
            DateOnly.FromDateTime(DateTime.Today),
            employeesCovered,
            totalClaimsPaid,
            SafeDivide(totalClaimsPaid, employeesCovered));
    }

    private static ScheduleOfBenefitsDeckData ReadScheduleOfBenefitsData(string workbookPath)
    {
        using var workbook = OpenWorkbookReadOnly(workbookPath);
        var clientName = ReadClientSetupValue(workbook, "Client Name", "Client Name");
        var rows = new List<ScheduleOfBenefitsRow>();

        try
        {
            var table = workbook.Table("tblPPT_ScheduleOfBenefits");
            var headers = table.Fields.Select((field, index) => new { field.Name, Index = index + 1 })
                .ToDictionary(item => item.Name, item => item.Index, StringComparer.OrdinalIgnoreCase);

            if (!headers.TryGetValue("Slide Page", out var slidePageColumn) ||
                !headers.TryGetValue("Section Name", out var sectionNameColumn) ||
                !headers.TryGetValue("Detail Text", out var detailTextColumn))
            {
                return new ScheduleOfBenefitsDeckData(clientName, []);
            }

            headers.TryGetValue("Page Total", out var pageTotalColumn);
            headers.TryGetValue("Section Order", out var sectionOrderColumn);
            headers.TryGetValue("Client Name", out var clientNameColumn);
            headers.TryGetValue("Carrier", out var carrierColumn);
            headers.TryGetValue("Class Name", out var classNameColumn);
            headers.TryGetValue("Source File", out var sourceFileColumn);
            headers.TryGetValue("Notes", out var notesColumn);

            foreach (var row in table.DataRange.Rows())
            {
                var sectionName = CellText(row.Cell(sectionNameColumn));
                var detailText = CellMultilineText(row.Cell(detailTextColumn));
                if (string.IsNullOrWhiteSpace(sectionName) && string.IsNullOrWhiteSpace(detailText))
                {
                    continue;
                }

                rows.Add(new ScheduleOfBenefitsRow(
                    NormalizeScheduleSlidePage((int)Math.Round(CellNumber(row.Cell(slidePageColumn)), MidpointRounding.AwayFromZero)),
                    pageTotalColumn == 0 ? 3 : (int)Math.Round(CellNumber(row.Cell(pageTotalColumn)), MidpointRounding.AwayFromZero),
                    sectionOrderColumn == 0 ? rows.Count + 1 : (int)Math.Round(CellNumber(row.Cell(sectionOrderColumn)), MidpointRounding.AwayFromZero),
                    clientNameColumn == 0 ? clientName : FirstNonBlank(CellText(row.Cell(clientNameColumn)), clientName),
                    carrierColumn == 0 ? "" : CellText(row.Cell(carrierColumn)),
                    classNameColumn == 0 ? "Class 1" : FirstNonBlank(CellText(row.Cell(classNameColumn)), "Class 1"),
                    FirstNonBlank(sectionName, "Schedule of Benefits"),
                    detailText,
                    sourceFileColumn == 0 ? "" : CellText(row.Cell(sourceFileColumn)),
                    notesColumn == 0 ? "" : CellText(row.Cell(notesColumn))));
            }
        }
        catch
        {
            return new ScheduleOfBenefitsDeckData(clientName, []);
        }

        var preferredClientName = rows.Select(row => row.ClientName).FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)) ?? clientName;
        return new ScheduleOfBenefitsDeckData(
            preferredClientName,
            rows
                .OrderBy(row => row.SlidePage)
                .ThenBy(row => row.SectionOrder)
                .ToList());
    }

    private static OverallPremiumClaimsSummarySlideData BuildOverallPremiumClaimsSummary(
        string currentPeriodLabel,
        PremiumClaimsSummarySlideData healthSummary,
        PremiumClaimsSummarySlideData dentalSummary)
    {
        var healthCurrent = CurrentPremiumClaimsRow(healthSummary.ChartRows, healthSummary.CurrentPeriodLabel, healthSummary.CurrentPremium, healthSummary.CurrentPaidClaims);
        var dentalCurrent = CurrentPremiumClaimsRow(dentalSummary.ChartRows, dentalSummary.CurrentPeriodLabel, dentalSummary.CurrentPremium, dentalSummary.CurrentPaidClaims);
        var rows = new List<OverallPremiumClaimsSummaryRow>
        {
            new("Extended Health", healthCurrent.Premium, healthCurrent.PaidClaims, SafeDivide(healthCurrent.PaidClaims, healthCurrent.Premium)),
            new("Dental", dentalCurrent.Premium, dentalCurrent.PaidClaims, SafeDivide(dentalCurrent.PaidClaims, dentalCurrent.Premium))
        };

        return new OverallPremiumClaimsSummarySlideData(currentPeriodLabel, rows);
    }

    private static PremiumClaimsChartRow CurrentPremiumClaimsRow(
        IReadOnlyList<PremiumClaimsChartRow> rows,
        string currentPeriodLabel,
        decimal currentPremium,
        decimal currentPaidClaims)
    {
        return rows.LastOrDefault(row => row.PeriodLabel.Equals(currentPeriodLabel, StringComparison.OrdinalIgnoreCase))
               ?? rows.LastOrDefault()
               ?? new PremiumClaimsChartRow(currentPeriodLabel, currentPremium, currentPaidClaims);
    }

    private static int NormalizeReportMonthCount(decimal monthCount)
    {
        var rounded = (int)Math.Round(monthCount, 0, MidpointRounding.AwayFromZero);
        return rounded is >= 1 and <= 36 ? rounded : 12;
    }

    private static string CurrentPeriodId(int monthCount)
    {
        return $"CURR_{monthCount}";
    }

    private static string PriorPeriodId(int monthCount)
    {
        return monthCount == 12 ? "PRIOR_12" : $"PRIOR_{monthCount}";
    }

    private static string BenchmarkPeriodId(int monthCount)
    {
        return monthCount == 12 ? "PRIOR_PRIOR_12" : "PRIOR_12";
    }

    private static IReadOnlyList<PremiumClaimsChartRow> BuildPremiumClaimsRows(
        IReadOnlyList<MonthlyExperienceRow> monthly,
        IReadOnlyList<(string PeriodId, string PeriodLabel)> periods,
        string coverage,
        bool useNetPremium,
        decimal ehcPoolingCharge)
    {
        return periods
            .Select(period => BuildPremiumClaimsRow(monthly, period.PeriodId, period.PeriodLabel, coverage, useNetPremium, ehcPoolingCharge))
            .ToList();
    }

    private static IReadOnlyList<PremiumClaimsChartRow> ReadCanonicalPremiumClaimsRows(
        XLWorkbook workbook,
        IReadOnlyList<MonthlyExperienceRow> monthly,
        int reportMonthCount,
        string coverage,
        bool useNetPremium,
        decimal ehcPoolingCharge)
    {
        try
        {
            var periodsTable = workbook.Table("tblPPT_CanonicalClaimsPeriods");
            var periodHeaders = periodsTable.Fields
                .Select((field, index) => new { field.Name, Index = index + 1 })
                .ToDictionary(item => item.Name, item => item.Index, StringComparer.OrdinalIgnoreCase);
            if (!periodHeaders.TryGetValue("Period Role", out var roleColumn) ||
                !periodHeaders.TryGetValue("Period Label", out var labelColumn) ||
                !periodHeaders.TryGetValue("Start Date", out var startColumn) ||
                !periodHeaders.TryGetValue("Month Count", out var monthCountColumn))
            {
                return [];
            }

            var periods = periodsTable.DataRange.Rows()
                .Select(row => new CanonicalComparisonPeriod(
                    CellText(row.Cell(roleColumn)),
                    CellText(row.Cell(labelColumn)),
                    CellDate(row.Cell(startColumn)),
                    (int)Math.Round(CellNumber(row.Cell(monthCountColumn)), MidpointRounding.AwayFromZero)))
                .Where(period => !string.IsNullOrWhiteSpace(period.Role) && period.SourceMonthCount > 0)
                .OrderBy(period => period.Start)
                .TakeLast(4)
                .ToList();
            if (periods.Count == 0)
            {
                return [];
            }

            var rows = new List<PremiumClaimsChartRow>();
            foreach (var period in periods)
            {
                var periodId = ComparisonPeriodId(period.Role, reportMonthCount);
                var periodLabel = FirstNonBlank(ReadPeriodLabel(workbook, periodId), period.PeriodLabel);
                var monthlyRow = BuildPremiumClaimsRow(monthly, periodId, periodLabel, coverage, useNetPremium, ehcPoolingCharge);
                var canonicalPaid = period.SourceMonthCount == reportMonthCount
                    ? ReadCanonicalPaidClaims(workbook, period.Role, coverage)
                    : null;
                rows.Add(monthlyRow with { PaidClaims = canonicalPaid ?? monthlyRow.PaidClaims });
            }

            return rows;
        }
        catch
        {
            return [];
        }
    }

    private static decimal? ReadCanonicalPaidClaims(XLWorkbook workbook, string periodRole, string coverage)
    {
        var table = workbook.Table("tblPPT_CanonicalClaimsAdjudication");
        var headers = table.Fields
            .Select((field, index) => new { field.Name, Index = index + 1 })
            .ToDictionary(item => item.Name, item => item.Index, StringComparer.OrdinalIgnoreCase);
        if (!headers.TryGetValue("Period Role", out var roleColumn) ||
            !headers.TryGetValue("Benefit", out var benefitColumn) ||
            !headers.TryGetValue("Amount Paid", out var paidColumn))
        {
            return null;
        }

        var matchingRows = table.DataRange.Rows()
            .Where(row => CellText(row.Cell(roleColumn)).Equals(periodRole, StringComparison.OrdinalIgnoreCase) &&
                          CellText(row.Cell(benefitColumn)).Equals(coverage, StringComparison.OrdinalIgnoreCase))
            .ToList();
        return matchingRows.Count == 0
            ? null
            : matchingRows.Sum(row => CellNumber(row.Cell(paidColumn)));
    }

    private static string ComparisonPeriodId(string role, int monthCount)
    {
        var normalized = new string((role ?? "").Where(char.IsLetterOrDigit).Select(char.ToUpperInvariant).ToArray());
        var prefix = normalized switch
        {
            "CURRENT" => "CURR",
            "PRIOR" => "PRIOR",
            "PRIORPRIOR" => "PRIOR_PRIOR",
            "PRIORPRIORPRIOR" => "PRIOR_PRIOR_PRIOR",
            _ => "PRIOR"
        };
        return $"{prefix}_{monthCount}";
    }
    private static PremiumClaimsChartRow BuildPremiumClaimsRow(
        IReadOnlyList<MonthlyExperienceRow> monthly,
        string periodId,
        string periodLabel,
        string coverage,
        bool useNetPremium,
        decimal ehcPoolingCharge)
    {
        var rows = monthly
            .Where(row => row.PeriodId.Equals(periodId, StringComparison.OrdinalIgnoreCase) &&
                          row.Coverage.Equals(coverage, StringComparison.OrdinalIgnoreCase))
            .ToList();
        var premium = rows.Sum(row => useNetPremium ? EffectiveEhcNetPremium(row, ehcPoolingCharge) : row.BilledPremium);
        var paidClaims = rows.Sum(row => row.PaidClaims);
        return new PremiumClaimsChartRow(periodLabel, premium, paidClaims);
    }

    private static Dictionary<string, PowerPointPlaceholderValue> ReadPowerPointPlaceholderValues(XLWorkbook workbook)
    {
        var values = new Dictionary<string, PowerPointPlaceholderValue>(StringComparer.OrdinalIgnoreCase);
        IXLTable table;
        try
        {
            table = workbook.Table("tblPPT_ClaimsExperienceSummary");
        }
        catch
        {
            return values;
        }

        var headers = table.Fields.Select((field, index) => new { field.Name, Index = index + 1 })
            .ToDictionary(item => item.Name, item => item.Index, StringComparer.OrdinalIgnoreCase);
        if (!headers.TryGetValue("Placeholder Key", out var keyColumn) ||
            !headers.TryGetValue("Value", out var valueColumn))
        {
            return values;
        }

        foreach (var row in table.DataRange.Rows())
        {
            var key = CellText(row.Cell(keyColumn));
            if (string.IsNullOrWhiteSpace(key))
            {
                continue;
            }

            values[key] = new PowerPointPlaceholderValue(
                CellText(row.Cell(valueColumn)),
                CellNumber(row.Cell(valueColumn)));
        }

        return values;
    }

    private static IReadOnlyList<PremiumClaimsChartRow> ReadPremiumClaimsChartRows(
        XLWorkbook workbook,
        string tableName,
        string premiumColumnName,
        string priorPeriodLabel,
        string currentPeriodLabel,
        decimal currentPremium,
        decimal currentPaidClaims,
        int maxRows = 2,
        IReadOnlyList<PremiumClaimsChartRow>? fallbackRows = null)
    {
        try
        {
            var table = workbook.Table(tableName);
            var headers = table.Fields.Select((field, index) => new { field.Name, Index = index + 1 })
                .ToDictionary(item => item.Name, item => item.Index, StringComparer.OrdinalIgnoreCase);
            if (!headers.TryGetValue("Period Label", out var periodColumn) ||
                !headers.TryGetValue(premiumColumnName, out var premiumColumn) ||
                !headers.TryGetValue("Paid Claims", out var paidClaimsColumn))
            {
                return fallbackRows ?? FallbackPremiumClaimsChartRows(priorPeriodLabel, currentPeriodLabel, currentPremium, currentPaidClaims);
            }

            var rows = table.DataRange.Rows()
                .Select(row => new PremiumClaimsChartRow(
                    FirstNonBlank(CellText(row.Cell(periodColumn)), row.RowNumber() == table.DataRange.FirstRow().RowNumber() ? priorPeriodLabel : currentPeriodLabel),
                    ChartCellNumber(row.Cell(premiumColumn)),
                    ChartCellNumber(row.Cell(paidClaimsColumn))))
                .Where(row => !string.IsNullOrWhiteSpace(row.PeriodLabel))
                .Take(maxRows)
                .ToList();

            if (rows.Count == 0)
            {
                return fallbackRows ?? FallbackPremiumClaimsChartRows(priorPeriodLabel, currentPeriodLabel, currentPremium, currentPaidClaims);
            }

            return fallbackRows is null
                ? rows
                : MergeFormulaBackedPremiumClaimsRows(rows, fallbackRows);
        }
        catch
        {
            return fallbackRows ?? FallbackPremiumClaimsChartRows(priorPeriodLabel, currentPeriodLabel, currentPremium, currentPaidClaims);
        }
    }

    private static IReadOnlyList<PremiumClaimsChartRow> MergeFormulaBackedPremiumClaimsRows(
        IReadOnlyList<PremiumClaimsChartRow> rows,
        IReadOnlyList<PremiumClaimsChartRow> fallbackRows)
    {
        var fallbackByPeriod = fallbackRows
            .Where(row => !string.IsNullOrWhiteSpace(row.PeriodLabel))
            .GroupBy(row => row.PeriodLabel, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.First(), StringComparer.OrdinalIgnoreCase);

        return rows
            .Select(row =>
            {
                if (!fallbackByPeriod.TryGetValue(row.PeriodLabel, out var fallback))
                {
                    return row;
                }

                return new PremiumClaimsChartRow(
                    row.PeriodLabel,
                    row.Premium == 0m && fallback.Premium != 0m ? fallback.Premium : row.Premium,
                    row.PaidClaims == 0m && fallback.PaidClaims != 0m ? fallback.PaidClaims : row.PaidClaims);
            })
            .ToList();
    }

    private static IReadOnlyList<PremiumClaimsChartRow> FallbackPremiumClaimsChartRows(
        string priorPeriodLabel,
        string currentPeriodLabel,
        decimal currentPremium,
        decimal currentPaidClaims)
    {
        return
        [
            new PremiumClaimsChartRow(priorPeriodLabel, 0m, 0m),
            new PremiumClaimsChartRow(currentPeriodLabel, currentPremium, currentPaidClaims)
        ];
    }

    private static IReadOnlyList<LossRatioBreakEvenChartRow> ReadLossRatioBreakEvenChartRows(
        XLWorkbook workbook,
        string tableName,
        IReadOnlyList<PremiumClaimsChartRow> fallbackPremiumRows,
        decimal fallbackBreakEven)
    {
        try
        {
            var table = workbook.Table(tableName);
            var headers = table.Fields.Select((field, index) => new { field.Name, Index = index + 1 })
                .ToDictionary(item => item.Name, item => item.Index, StringComparer.OrdinalIgnoreCase);
            var lossRatioHeader = new[] { "Net Loss Ratio", "Paid Loss Ratio", "Loss Ratio" }
                .FirstOrDefault(headers.ContainsKey);
            if (!headers.TryGetValue("Period Label", out var periodColumn) ||
                lossRatioHeader is null ||
                !headers.TryGetValue("Break Even Ratio", out var breakEvenColumn))
            {
                return FallbackLossRatioBreakEvenRows(fallbackPremiumRows, fallbackBreakEven);
            }

            var lossRatioColumn = headers[lossRatioHeader];
            var rows = table.DataRange.Rows()
                .Select(row => new LossRatioBreakEvenChartRow(
                    CellText(row.Cell(periodColumn)),
                    CellNumber(row.Cell(lossRatioColumn)),
                    CellNumber(row.Cell(breakEvenColumn))))
                .Where(row => !string.IsNullOrWhiteSpace(row.PeriodLabel))
                .Take(4)
                .ToList();

            var fallbackRows = FallbackLossRatioBreakEvenRows(fallbackPremiumRows, fallbackBreakEven);
            if (rows.Count == 0)
            {
                return fallbackRows;
            }

            var fallbackByPeriod = fallbackRows
                .Where(row => !string.IsNullOrWhiteSpace(row.PeriodLabel))
                .GroupBy(row => row.PeriodLabel, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.OrdinalIgnoreCase);

            return rows
                .Select(row =>
                {
                    if (!fallbackByPeriod.TryGetValue(row.PeriodLabel, out var fallback))
                    {
                        return row;
                    }

                    return new LossRatioBreakEvenChartRow(
                        row.PeriodLabel,
                        row.NetLossRatio == 0m && fallback.NetLossRatio != 0m ? fallback.NetLossRatio : row.NetLossRatio,
                        row.BreakEven == 0m && fallback.BreakEven != 0m ? fallback.BreakEven : row.BreakEven);
                })
                .ToList();
        }
        catch
        {
            return FallbackLossRatioBreakEvenRows(fallbackPremiumRows, fallbackBreakEven);
        }
    }

    private static IReadOnlyList<LossRatioBreakEvenChartRow> FallbackLossRatioBreakEvenRows(
        IReadOnlyList<PremiumClaimsChartRow> premiumRows,
        decimal breakEven)
    {
        return premiumRows
            .Select(row => new LossRatioBreakEvenChartRow(row.PeriodLabel, SafeDivide(row.PaidClaims, row.Premium), breakEven))
            .ToList();
    }

    private static Dictionary<string, decimal> ReadParamedicalServiceAmounts(XLWorkbook workbook, string periodId)
    {
        IXLTable table;
        try
        {
            table = workbook.Table("tblParamedical");
        }
        catch
        {
            return new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
        }

        var headers = table.Fields.Select((field, index) => new { field.Name, Index = index + 1 })
            .ToDictionary(item => item.Name, item => item.Index, StringComparer.OrdinalIgnoreCase);
        if (!headers.TryGetValue("Period ID", out var periodColumn) ||
            !headers.TryGetValue("Service", out var serviceColumn) ||
            !headers.TryGetValue("Paid Claims", out var paidClaimsColumn))
        {
            return new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
        }

        return table.DataRange.Rows()
            .Where(row => CellText(row.Cell(periodColumn)).Equals(periodId, StringComparison.OrdinalIgnoreCase))
            .GroupBy(row => FirstNonBlank(CellText(row.Cell(serviceColumn)), "Other Paramedical"), StringComparer.OrdinalIgnoreCase)
            .ToDictionary(
                group => group.Key,
                group => group.Sum(row => CellNumber(row.Cell(paidClaimsColumn))),
                StringComparer.OrdinalIgnoreCase);
    }

    private static decimal SumAndMarkMatchingServices(
        IReadOnlyDictionary<string, decimal> rows,
        IReadOnlyList<string> searchTerms,
        ISet<string> matchedServices)
    {
        var total = 0m;
        foreach (var item in rows)
        {
            if (matchedServices.Contains(item.Key) ||
                !searchTerms.Any(term => item.Key.Contains(term, StringComparison.OrdinalIgnoreCase)))
            {
                continue;
            }

            total += item.Value;
            matchedServices.Add(item.Key);
        }

        return total;
    }

    private static decimal SumUnmatchedServices(
        IReadOnlyDictionary<string, decimal> rows,
        ISet<string> matchedServices)
    {
        return rows
            .Where(item => !matchedServices.Contains(item.Key))
            .Sum(item => item.Value);
    }

    private static Dictionary<string, decimal> ReadCategoryAmounts(XLWorkbook workbook, string tableName, string periodId)
    {
        IXLTable table;
        try
        {
            table = workbook.Table(tableName);
        }
        catch
        {
            return new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
        }

        var headers = table.Fields.Select((field, index) => new { field.Name, Index = index + 1 })
            .ToDictionary(item => item.Name, item => item.Index, StringComparer.OrdinalIgnoreCase);
        if (!headers.TryGetValue("Period ID", out var periodColumn) ||
            !headers.TryGetValue("Category", out var categoryColumn) ||
            !headers.TryGetValue("Paid Claims", out var paidClaimsColumn))
        {
            return new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
        }

        return table.DataRange.Rows()
            .Where(row => CellText(row.Cell(periodColumn)).Equals(periodId, StringComparison.OrdinalIgnoreCase))
            .GroupBy(row => FirstNonBlank(CellText(row.Cell(categoryColumn)), "Other"), StringComparer.OrdinalIgnoreCase)
            .ToDictionary(
                group => group.Key,
                group => group.Sum(row => CellNumber(row.Cell(paidClaimsColumn))),
                StringComparer.OrdinalIgnoreCase);
    }

    private static decimal SumMatchingCategory(IReadOnlyDictionary<string, decimal> rows, string category)
    {
        return rows
            .Where(item => item.Key.Contains(category, StringComparison.OrdinalIgnoreCase))
            .Sum(item => item.Value);
    }

    private static decimal OtherEhcClaims(IReadOnlyDictionary<string, decimal> rows)
    {
        return rows
            .Where(item => !item.Key.Contains("Paramedical", StringComparison.OrdinalIgnoreCase) &&
                           !item.Key.Contains("Drugs", StringComparison.OrdinalIgnoreCase) &&
                           !item.Key.Contains("Vision", StringComparison.OrdinalIgnoreCase))
            .Sum(item => item.Value);
    }

    private static MonthlyBreakdownSlideData BuildMonthlyBreakdownSlideData(
        XLWorkbook workbook,
        IReadOnlyList<MonthlyExperienceRow> monthly,
        string currentPeriodId,
        string priorPeriodId,
        string slideTitle,
        string coverage,
        string categoryTableName,
        string employeeCoverageLabel,
        decimal ehcPoolingCharge)
    {
        var useNetPremium = coverage.Equals("EHC", StringComparison.OrdinalIgnoreCase);
        var premiumHeader = useNetPremium ? "Net Premium" : "Billed Premium";
        var rows = monthly
            .Where(row => row.PeriodId.Equals(currentPeriodId, StringComparison.OrdinalIgnoreCase) &&
                          row.Coverage.Equals(coverage, StringComparison.OrdinalIgnoreCase))
            .GroupBy(row => row.Month)
            .OrderBy(group => group.Key)
            .Select(group =>
            {
                var premium = group.Sum(row => useNetPremium ? EffectiveEhcNetPremium(row, ehcPoolingCharge) : row.BilledPremium);
                var paidClaims = group.Sum(row => row.PaidClaims);
                var lives = group.Sum(row => row.Lives);
                return new MonthlyBreakdownRow(
                    group.Key,
                    premium,
                    paidClaims,
                    SafeDivide(paidClaims, premium),
                    lives);
            })
            .ToList();

        var totalPremium = rows.Sum(row => row.Premium);
        var totalPaidClaims = rows.Sum(row => row.PaidClaims);
        var averageLives = rows.Where(row => row.EmployeeCount > 0m).Select(row => row.EmployeeCount).DefaultIfEmpty(0m).Average();
        var topCategory = ReadTopCategory(workbook, categoryTableName, currentPeriodId);
        var highestMonth = rows.OrderByDescending(row => row.PaidClaims).FirstOrDefault();
        var priorClaims = monthly
            .Where(row => row.PeriodId.Equals(priorPeriodId, StringComparison.OrdinalIgnoreCase) &&
                          row.Coverage.Equals(coverage, StringComparison.OrdinalIgnoreCase))
            .Sum(row => row.PaidClaims);

        return new MonthlyBreakdownSlideData(
            slideTitle,
            coverage,
            employeeCoverageLabel,
            premiumHeader,
            rows,
            totalPremium,
            totalPaidClaims,
            SafeDivide(totalPaidClaims, totalPremium),
            averageLives,
            topCategory.Category,
            topCategory.PaidClaims,
            highestMonth?.Month,
            highestMonth?.PaidClaims ?? 0m,
            SafeDivide(totalPaidClaims, averageLives),
            priorClaims);
    }

    private static decimal EffectiveEhcNetPremium(MonthlyExperienceRow row, decimal ehcPoolingCharge)
    {
        return row.NetPremium != 0m ? row.NetPremium : row.BilledPremium * (1m - ehcPoolingCharge);
    }

    private static decimal? ReadAssumptionNumber(XLWorkbook workbook, string assumptionName)
    {
        try
        {
            var table = workbook.Table("tblAssumptions");
            foreach (var row in table.DataRange.Rows())
            {
                if (CellText(row.Cell(1)).Equals(assumptionName, StringComparison.OrdinalIgnoreCase))
                {
                    return CellNumber(row.Cell(2));
                }
            }
        }
        catch
        {
            return null;
        }

        return null;
    }

    private static CategorySpend ReadTopCategory(XLWorkbook workbook, string tableName, string periodId)
    {
        IXLTable table;
        try
        {
            table = workbook.Table(tableName);
        }
        catch
        {
            return new CategorySpend("the top category", 0m);
        }

        var headers = table.Fields.Select((field, index) => new { field.Name, Index = index + 1 })
            .ToDictionary(item => item.Name, item => item.Index, StringComparer.OrdinalIgnoreCase);
        if (!headers.TryGetValue("Period ID", out var periodColumn) ||
            !headers.TryGetValue("Category", out var categoryColumn) ||
            !headers.TryGetValue("Paid Claims", out var paidClaimsColumn))
        {
            return new CategorySpend("the top category", 0m);
        }

        var topCategory = table.DataRange.Rows()
            .Where(row => CellText(row.Cell(periodColumn)).Equals(periodId, StringComparison.OrdinalIgnoreCase))
            .Select(row => new
            {
                Category = FirstNonBlank(CellText(row.Cell(categoryColumn)), "Other"),
                PaidClaims = CellNumber(row.Cell(paidClaimsColumn))
            })
            .Where(row => !string.IsNullOrWhiteSpace(row.Category))
            .GroupBy(row => row.Category, StringComparer.OrdinalIgnoreCase)
            .Select(group => new CategorySpend(group.Key, group.Sum(row => row.PaidClaims)))
            .OrderByDescending(row => row.PaidClaims)
            .FirstOrDefault();

        return topCategory ?? new CategorySpend("the top category", 0m);
    }

    private static IEnumerable<MonthlyExperienceRow> ReadMonthlyExperienceRows(XLWorkbook workbook)
    {
        var table = workbook.Table("tblMonthlyExperience");
        var headers = table.Fields.Select((field, index) => new { field.Name, Index = index + 1 })
            .ToDictionary(item => item.Name, item => item.Index, StringComparer.OrdinalIgnoreCase);

        foreach (var row in table.DataRange.Rows())
        {
            var periodId = CellText(row.Cell(headers["Period ID"]));
            var coverage = CellText(row.Cell(headers["Coverage"]));
            if (string.IsNullOrWhiteSpace(periodId) || string.IsNullOrWhiteSpace(coverage))
            {
                continue;
            }

            yield return new MonthlyExperienceRow(
                periodId,
                CellDate(row.Cell(headers["Month"])),
                coverage,
                CellNumber(row.Cell(headers["Billed Premium"])),
                CellNumber(row.Cell(headers["Net Premium"])),
                CellNumber(row.Cell(headers["Paid Claims"])),
                CellNumber(row.Cell(headers["Claims Plus Fees"])),
                CellNumber(row.Cell(headers["Lives"])));
        }
    }

    private static string ReadReportControlValue(XLWorkbook workbook, string controlName, string fallback)
    {
        var table = workbook.Table("tblReportControl");
        foreach (var row in table.DataRange.Rows())
        {
            if (CellText(row.Cell(1)).Equals(controlName, StringComparison.OrdinalIgnoreCase))
            {
                var value = FirstNonBlank(CellText(row.Cell(2)), "");
                if (!string.IsNullOrWhiteSpace(value))
                {
                    return value;
                }

                if (controlName.Equals("Report Month Count", StringComparison.OrdinalIgnoreCase))
                {
                    return fallback;
                }

                var monthCount = NormalizeReportMonthCount(ReadReportControlNumber(workbook, "Report Month Count") ?? 12m);
                return controlName switch
                {
                    "Current Period ID" => CurrentPeriodId(monthCount),
                    "Prior Period ID" => PriorPeriodId(monthCount),
                    "Benchmark Period ID" => BenchmarkPeriodId(monthCount),
                    _ => fallback
                };
            }
        }

        return fallback;
    }

    private static string ReadClientSetupValue(XLWorkbook workbook, string fieldName, string fallback)
    {
        try
        {
            var table = workbook.Table("tblClientSetup");
            foreach (var row in table.DataRange.Rows())
            {
                if (CellText(row.Cell(1)).Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                {
                    return FirstNonBlank(CellText(row.Cell(2)), fallback);
                }
            }
        }
        catch
        {
            return fallback;
        }

        return fallback;
    }

    private static decimal? ReadReportControlNumber(XLWorkbook workbook, string controlName)
    {
        var value = ReadReportControlValue(workbook, controlName, "");
        return ParsingHelpers.ParseDecimal(value);
    }

    private static string ReadPeriodLabel(XLWorkbook workbook, string periodId)
    {
        var table = workbook.Table("tblPeriods");
        foreach (var row in table.DataRange.Rows())
        {
            if (CellText(row.Cell(1)).Equals(periodId, StringComparison.OrdinalIgnoreCase))
            {
                return CellText(row.Cell(2));
            }
        }

        return periodId;
    }

    private static bool UpdateAverageSpendSlide(ZipArchive archive, AverageSpendSlideData data)
    {
        var slidePath = FindSlidePathByTitle(archive, "Average Amount Spent by Employee");
        if (slidePath is null)
        {
            return false;
        }

        var entry = archive.GetEntry(slidePath) ?? throw new InvalidOperationException($"Slide part not found: {slidePath}");
        var document = XDocument.Parse(ReadZipEntryText(entry));
        var a = XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/main");
        var moneyCounter = 0;

        foreach (var text in document.Descendants(a + "t"))
        {
            var value = text.Value;
            if (value.Contains("[current reporting period]", StringComparison.OrdinalIgnoreCase))
            {
                value = value.Replace("[current reporting period]", data.CurrentPeriodLabel, StringComparison.OrdinalIgnoreCase);
            }

            if (value.Contains("[prior reporting period]", StringComparison.OrdinalIgnoreCase))
            {
                value = value.Replace("[prior reporting period]", data.PriorPeriodLabel, StringComparison.OrdinalIgnoreCase);
            }

            value = value.Replace("([X]-month reporting period)", $"{data.MonthCount}-month reporting period", StringComparison.OrdinalIgnoreCase);
            value = value.Replace("([X]) employees enrolled in Extended Health", $"{FormatLives(data.HealthEmployeeCount)} employees enrolled in Extended Health", StringComparison.OrdinalIgnoreCase);
            value = value.Replace("([X]) employees enrolled in Dental", $"{FormatLives(data.DentalEmployeeCount)} employees enrolled in Dental", StringComparison.OrdinalIgnoreCase);
            value = value.Replace("[enter month]", data.HighestClaimsMonth?.ToString("MMMM", Culture) ?? "the highest month", StringComparison.OrdinalIgnoreCase);

            if (value.Contains("[$X]", StringComparison.OrdinalIgnoreCase))
            {
                moneyCounter++;
                var replacement = moneyCounter switch
                {
                    1 => FormatMoney(data.Health.CurrentAmount),
                    2 => FormatMoney(data.Dental.CurrentAmount),
                    3 => FormatMoney(data.Health.CurrentAmount),
                    4 => FormatMoney(data.Dental.CurrentAmount),
                    5 => FormatMoney(data.Combined.CurrentAmount),
                    _ => FormatMoney(data.HighestClaimsAmount)
                };
                value = value.Replace("[$X]", replacement, StringComparison.OrdinalIgnoreCase);
            }

            text.Value = value;
        }

        entry.Delete();
        var newEntry = archive.CreateEntry(slidePath, CompressionLevel.Optimal);
        using var stream = newEntry.Open();
        document.Save(stream);
        return true;
    }

    private static void UpdateAverageSpendChartXml(ZipArchive archive, string chartPath, AverageSpendSlideData data)
    {
        var entry = archive.GetEntry(chartPath) ?? throw new InvalidOperationException($"Chart part not found: {chartPath}");
        var document = XDocument.Parse(ReadZipEntryText(entry));
        var c = XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/chart");
        var series = document.Descendants(c + "ser").Take(2).ToList();
        if (series.Count < 2)
        {
            return;
        }

        var categories = data.Rows.Select(row => row.Category).ToList();
        SetStringReference(series[0].Element(c + "tx")?.Element(c + "strRef"), c, "Sheet1!$B$1", ["Current Period"]);
        SetStringReference(series[0].Element(c + "cat")?.Element(c + "strRef"), c, $"Sheet1!$A$2:$A${data.Rows.Count + 1}", categories);
        SetNumberReference(series[0].Element(c + "val")?.Element(c + "numRef"), c, $"Sheet1!$B$2:$B${data.Rows.Count + 1}", data.Rows.Select(row => row.CurrentAmount).ToList());

        SetStringReference(series[1].Element(c + "tx")?.Element(c + "strRef"), c, "Sheet1!$C$1", ["Prior Period"]);
        SetStringReference(series[1].Element(c + "cat")?.Element(c + "strRef"), c, $"Sheet1!$A$2:$A${data.Rows.Count + 1}", categories);
        SetNumberReference(series[1].Element(c + "val")?.Element(c + "numRef"), c, $"Sheet1!$C$2:$C${data.Rows.Count + 1}", data.Rows.Select(row => row.PriorAmount).ToList());

        SanitizeChartExtensions(document, c);
        entry.Delete();
        var newEntry = archive.CreateEntry(chartPath, CompressionLevel.Optimal);
        using var stream = newEntry.Open();
        document.Save(stream);
    }

    private static void UpdateEmbeddedAverageSpendWorkbook(ZipArchive archive, string workbookPath, AverageSpendSlideData data)
    {
        var entry = archive.GetEntry(workbookPath);
        if (entry is null)
        {
            return;
        }

        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("Sheet1");
        sheet.Cell(1, 1).Value = "Coverage";
        sheet.Cell(1, 2).Value = "Current Period";
        sheet.Cell(1, 3).Value = "Prior Period";
        sheet.Range(1, 1, 1, 3).Style.Font.Bold = true;

        for (var index = 0; index < data.Rows.Count; index++)
        {
            var rowNumber = index + 2;
            sheet.Cell(rowNumber, 1).Value = data.Rows[index].Category;
            sheet.Cell(rowNumber, 2).Value = data.Rows[index].CurrentAmount;
            sheet.Cell(rowNumber, 3).Value = data.Rows[index].PriorAmount;
        }

        sheet.Range(2, 2, data.Rows.Count + 1, 3).Style.NumberFormat.Format = "$#,##0.00";
        sheet.Columns(1, 3).AdjustToContents(8, 24);

        using var output = new MemoryStream();
        workbook.SaveAs(output);
        output.Position = 0;

        entry.Delete();
        var newEntry = archive.CreateEntry(workbookPath, CompressionLevel.Optimal);
        using var outputStream = newEntry.Open();
        output.CopyTo(outputStream);
    }

    private static bool UpdateClaimsDistributionSlide(ZipArchive archive, ClaimsDistributionSlideData data)
    {
        var slidePath = FindSlidePathByTitle(archive, "Program Performance: Claims Distribution");
        if (slidePath is null)
        {
            return false;
        }

        var entry = archive.GetEntry(slidePath) ?? throw new InvalidOperationException($"Slide part not found: {slidePath}");
        var document = XDocument.Parse(ReadZipEntryText(entry));
        var a = XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/main");
        var topChange = data.Rows
            .OrderByDescending(row => Math.Abs(row.CurrentPaidClaims - row.PriorPaidClaims))
            .FirstOrDefault();

        foreach (var text in document.Descendants(a + "t"))
        {
            var value = text.Value;
            value = value.Replace("[reporting period]", data.CurrentPeriodLabel, StringComparison.OrdinalIgnoreCase);
            value = value.Replace("[stable / increasing / moderating]", SpendingTrend(data.TotalCurrentPaidClaims, data.TotalPriorPaidClaims), StringComparison.OrdinalIgnoreCase);
            value = value.Replace("[the prior period / the previous year]", data.PriorPeriodLabel, StringComparison.OrdinalIgnoreCase);
            value = value.Replace("[drugs / paramedical / dental / vision]", topChange?.Category ?? "claims categories", StringComparison.OrdinalIgnoreCase);
            value = value.Replace("[normal variation in utilization / increased activity in specific areas]", CategoryChangeDescription(topChange), StringComparison.OrdinalIgnoreCase);
            value = value.Replace("[No significant / limited / several]", "Limited", StringComparison.OrdinalIgnoreCase);
            value = value.Replace("[consistent with / slightly above / slightly below]", UtilizationComparison(data.TotalCurrentPaidClaims, data.TotalPriorPaidClaims), StringComparison.OrdinalIgnoreCase);
            value = value.Replace("[this size / similar demographics]", "this size", StringComparison.OrdinalIgnoreCase);
            value = value.Replace("[forecasting upcoming renewal changes / evaluating potential plan adjustments]", "evaluating potential plan adjustments", StringComparison.OrdinalIgnoreCase);
            text.Value = value;
        }

        entry.Delete();
        var newEntry = archive.CreateEntry(slidePath, CompressionLevel.Optimal);
        using var stream = newEntry.Open();
        document.Save(stream);
        return true;
    }

    private static void UpdateClaimsDistributionChartXml(ZipArchive archive, string chartPath, ClaimsDistributionSlideData data)
    {
        var entry = archive.GetEntry(chartPath) ?? throw new InvalidOperationException($"Chart part not found: {chartPath}");
        var document = XDocument.Parse(ReadZipEntryText(entry));
        var c = XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/chart");
        var a = XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/main");
        var barChart = document.Descendants(c + "barChart").FirstOrDefault();
        if (barChart is null)
        {
            return;
        }

        var series = EnsureSeriesCount(barChart, c, 2);
        var categories = data.Rows.Select(row => row.Category).ToList();
        SetStringReference(series[0].Element(c + "tx")?.Element(c + "strRef"), c, "Sheet1!$B$1", ["Current Period"]);
        SetStringReference(series[0].Element(c + "cat")?.Element(c + "strRef"), c, $"Sheet1!$A$2:$A${data.Rows.Count + 1}", categories);
        SetNumberReference(series[0].Element(c + "val")?.Element(c + "numRef"), c, $"Sheet1!$B$2:$B${data.Rows.Count + 1}", data.Rows.Select(row => row.CurrentPaidClaims).ToList());
        SetSeriesFill(series[0], c, a, "BDE7DD");

        SetStringReference(series[1].Element(c + "tx")?.Element(c + "strRef"), c, "Sheet1!$C$1", ["Prior Period"]);
        SetStringReference(series[1].Element(c + "cat")?.Element(c + "strRef"), c, $"Sheet1!$A$2:$A${data.Rows.Count + 1}", categories);
        SetNumberReference(series[1].Element(c + "val")?.Element(c + "numRef"), c, $"Sheet1!$C$2:$C${data.Rows.Count + 1}", data.Rows.Select(row => row.PriorPaidClaims).ToList());
        SetSeriesFill(series[1], c, a, "7FCAD7");

        SanitizeChartExtensions(document, c);
        entry.Delete();
        var newEntry = archive.CreateEntry(chartPath, CompressionLevel.Optimal);
        using var stream = newEntry.Open();
        document.Save(stream);
    }

    private static List<XElement> EnsureSeriesCount(XElement barChart, XNamespace c, int count)
    {
        var series = barChart.Elements(c + "ser").ToList();
        if (series.Count == 0)
        {
            return series;
        }

        while (series.Count < count)
        {
            var clone = new XElement(series.Last());
            clone.Element(c + "extLst")?.Remove();
            series.Last().AddAfterSelf(clone);
            series.Add(clone);
        }

        foreach (var extra in series.Skip(count).ToList())
        {
            extra.Remove();
        }

        series = barChart.Elements(c + "ser").Take(count).ToList();
        for (var index = 0; index < series.Count; index++)
        {
            var idx = EnsureChild(series[index], c + "idx");
            idx.SetAttributeValue("val", index);
            var order = EnsureChild(series[index], c + "order");
            order.SetAttributeValue("val", index);
        }

        return series;
    }

    private static void SetSeriesFill(XElement series, XNamespace c, XNamespace a, string rgb)
    {
        var shapeProperties = EnsureChild(series, c + "spPr");
        shapeProperties.Elements(a + "noFill").Remove();
        shapeProperties.Elements(a + "solidFill").Remove();
        shapeProperties.Elements(a + "gradFill").Remove();
        shapeProperties.Elements(a + "blipFill").Remove();
        shapeProperties.Elements(a + "pattFill").Remove();
        shapeProperties.Elements(a + "grpFill").Remove();
        shapeProperties.AddFirst(new XElement(a + "solidFill", new XElement(a + "srgbClr", new XAttribute("val", rgb))));
    }

    private static void SetSeriesNoFill(XElement series, XNamespace c, XNamespace a)
    {
        var shapeProperties = EnsureChild(series, c + "spPr");
        shapeProperties.Elements(a + "noFill").Remove();
        shapeProperties.Elements(a + "solidFill").Remove();
        shapeProperties.Elements(a + "gradFill").Remove();
        shapeProperties.Elements(a + "blipFill").Remove();
        shapeProperties.Elements(a + "pattFill").Remove();
        shapeProperties.Elements(a + "grpFill").Remove();
        shapeProperties.AddFirst(new XElement(a + "noFill"));

        var line = shapeProperties.Element(a + "ln");
        if (line is null)
        {
            line = new XElement(a + "ln");
            shapeProperties.Add(line);
        }

        line.Elements(a + "noFill").Remove();
        line.Elements(a + "solidFill").Remove();
        line.Elements(a + "gradFill").Remove();
        line.Elements(a + "blipFill").Remove();
        line.Elements(a + "pattFill").Remove();
        line.Elements(a + "grpFill").Remove();
        line.AddFirst(new XElement(a + "noFill"));
    }

    private static void SetSeriesDataLabelsDeleted(XElement series, XNamespace c)
    {
        var labels = series.Element(c + "dLbls");
        if (labels is null)
        {
            labels = new XElement(c + "dLbls");
            var insertAfter = series.Element(c + "invertIfNegative") ??
                              series.Element(c + "spPr") ??
                              series.Element(c + "tx") ??
                              series.Element(c + "order") ??
                              series.Element(c + "idx");

            if (insertAfter is null)
            {
                series.AddFirst(labels);
            }
            else
            {
                insertAfter.AddAfterSelf(labels);
            }
        }

        labels.RemoveNodes();
        labels.Add(new XElement(c + "delete", new XAttribute("val", "1")));
    }

    private static void EnsureLinearTrendline(XElement series, XNamespace c, XNamespace a, string rgb)
    {
        var trendline = series.Element(c + "trendline");
        if (trendline is null)
        {
            trendline = new XElement(c + "trendline");
            var insertAfter = series.Elements(c + "dLbls").LastOrDefault() ??
                              series.Element(c + "invertIfNegative") ??
                              series.Element(c + "spPr") ??
                              series.Element(c + "tx") ??
                              series.Element(c + "order") ??
                              series.Element(c + "idx");

            if (insertAfter is null)
            {
                series.AddFirst(trendline);
            }
            else
            {
                insertAfter.AddAfterSelf(trendline);
            }
        }

        trendline.RemoveNodes();
        trendline.Add(
            new XElement(c + "spPr",
                new XElement(a + "ln",
                    new XAttribute("w", "28575"),
                    new XAttribute("cap", "rnd"),
                    new XElement(a + "solidFill",
                        new XElement(a + "srgbClr", new XAttribute("val", rgb))),
                    new XElement(a + "prstDash", new XAttribute("val", "solid"))),
                new XElement(a + "effectLst")),
            new XElement(c + "trendlineType", new XAttribute("val", "linear")),
            new XElement(c + "dispRSqr", new XAttribute("val", "0")),
            new XElement(c + "dispEq", new XAttribute("val", "0")));
    }

    private static void SanitizeChartExtensions(XDocument document, XNamespace c)
    {
        foreach (var extensionList in document.Descendants(c + "extLst").ToList())
        {
            extensionList.Remove();
        }

        RepairChartCacheOrdering(document, c);
        RemoveInvalidChartFillChoices(document, c);
    }

    private static void SanitizeAllChartParts(ZipArchive archive)
    {
        var chartEntries = archive.Entries
            .Where(entry => entry.FullName.StartsWith("ppt/charts/chart", StringComparison.OrdinalIgnoreCase) &&
                            entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
            .Select(entry => entry.FullName)
            .ToList();

        foreach (var chartPath in chartEntries)
        {
            var entry = archive.GetEntry(chartPath);
            if (entry is null)
            {
                continue;
            }

            var document = XDocument.Parse(ReadZipEntryText(entry));
            var c = XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/chart");
            SanitizeChartExtensions(document, c);

            entry.Delete();
            var newEntry = archive.CreateEntry(chartPath, CompressionLevel.Optimal);
            using var stream = newEntry.Open();
            document.Save(stream);
        }
    }

    private static void RepairChartCacheOrdering(XDocument document, XNamespace c)
    {
        foreach (var cache in document.Descendants(c + "numCache").Concat(document.Descendants(c + "numLit")).ToList())
        {
            RepairNumberCache(cache, c);
        }

        foreach (var cache in document.Descendants(c + "strCache").Concat(document.Descendants(c + "strLit")).ToList())
        {
            RepairStringCache(cache, c);
        }
    }

    private static void RepairNumberCache(XElement cache, XNamespace c)
    {
        var formatCode = cache.Elements(c + "formatCode").FirstOrDefault();
        var points = cache.Elements(c + "pt").Select(point => new XElement(point)).ToList();
        var pointCount = cache.Elements(c + "ptCount").FirstOrDefault();

        cache.RemoveNodes();
        if (formatCode is not null)
        {
            cache.Add(new XElement(formatCode));
        }

        cache.Add(pointCount is not null
            ? new XElement(pointCount)
            : new XElement(c + "ptCount", new XAttribute("val", points.Count)));

        foreach (var point in points)
        {
            cache.Add(point);
        }
    }

    private static void RepairStringCache(XElement cache, XNamespace c)
    {
        var points = cache.Elements(c + "pt").Select(point => new XElement(point)).ToList();
        var pointCount = cache.Elements(c + "ptCount").FirstOrDefault();

        cache.RemoveNodes();
        cache.Add(pointCount is not null
            ? new XElement(pointCount)
            : new XElement(c + "ptCount", new XAttribute("val", points.Count)));

        foreach (var point in points)
        {
            cache.Add(point);
        }
    }

    private static void RemoveInvalidChartFillChoices(XDocument document, XNamespace c)
    {
        var a = XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/main");
        var visibleFillChoices = new[]
        {
            a + "solidFill",
            a + "gradFill",
            a + "blipFill",
            a + "pattFill",
            a + "grpFill"
        };

        foreach (var shapeProperties in document.Descendants(c + "spPr").ToList())
        {
            var noFills = shapeProperties.Elements(a + "noFill").ToList();
            var visibleFills = shapeProperties.Elements()
                .Where(element => visibleFillChoices.Contains(element.Name))
                .ToList();

            if (visibleFills.Count > 0)
            {
                foreach (var noFill in noFills)
                {
                    noFill.Remove();
                }

                foreach (var duplicateFill in visibleFills.Skip(1))
                {
                    duplicateFill.Remove();
                }

                continue;
            }

            foreach (var duplicateNoFill in noFills.Skip(1))
            {
                duplicateNoFill.Remove();
            }
        }
    }

    private static void UpdateEmbeddedClaimsDistributionWorkbook(ZipArchive archive, string workbookPath, ClaimsDistributionSlideData data)
    {
        var entry = archive.GetEntry(workbookPath);
        if (entry is null)
        {
            return;
        }

        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("Sheet1");
        sheet.Cell(1, 1).Value = "Category";
        sheet.Cell(1, 2).Value = "Current Period";
        sheet.Cell(1, 3).Value = "Prior Period";
        sheet.Range(1, 1, 1, 3).Style.Font.Bold = true;

        for (var index = 0; index < data.Rows.Count; index++)
        {
            var rowNumber = index + 2;
            sheet.Cell(rowNumber, 1).Value = data.Rows[index].Category;
            sheet.Cell(rowNumber, 2).Value = data.Rows[index].CurrentPaidClaims;
            sheet.Cell(rowNumber, 3).Value = data.Rows[index].PriorPaidClaims;
        }

        sheet.Range(2, 2, data.Rows.Count + 1, 3).Style.NumberFormat.Format = "$#,##0.00";
        sheet.Columns(1, 3).AdjustToContents(8, 24);

        using var output = new MemoryStream();
        workbook.SaveAs(output);
        output.Position = 0;

        entry.Delete();
        var newEntry = archive.CreateEntry(workbookPath, CompressionLevel.Optimal);
        using var outputStream = newEntry.Open();
        output.CopyTo(outputStream);
    }

    private static bool UpdateParamedicalBreakdownSlide(ZipArchive archive, ParamedicalBreakdownSlideData data)
    {
        var slidePath = FindSlidePathByTitle(archive, "Claims Breakdown: Paramedical");
        if (slidePath is null)
        {
            return false;
        }

        var entry = archive.GetEntry(slidePath) ?? throw new InvalidOperationException($"Slide part not found: {slidePath}");
        var document = XDocument.Parse(ReadZipEntryText(entry));
        var a = XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/main");

        foreach (var text in document.Descendants(a + "t"))
        {
            var value = text.Value;
            value = value.Replace("[reporting period]", data.CurrentPeriodLabel, StringComparison.OrdinalIgnoreCase);
            value = value.Replace("$[XX]", FormatMoneyWhole(data.TotalCurrentPaidClaims), StringComparison.OrdinalIgnoreCase);
            value = value.Replace("[$XX]", FormatMoneyWhole(data.TotalCurrentPaidClaims), StringComparison.OrdinalIgnoreCase);
            text.Value = value;
        }

        entry.Delete();
        var newEntry = archive.CreateEntry(slidePath, CompressionLevel.Optimal);
        using var stream = newEntry.Open();
        document.Save(stream);
        return true;
    }

    private static void UpdateParamedicalBreakdownChartXml(ZipArchive archive, string chartPath, ParamedicalBreakdownSlideData data)
    {
        var entry = archive.GetEntry(chartPath) ?? throw new InvalidOperationException($"Chart part not found: {chartPath}");
        var document = XDocument.Parse(ReadZipEntryText(entry));
        var c = XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/chart");
        var a = XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/main");
        var barChart = document.Descendants(c + "barChart").FirstOrDefault();
        if (barChart is null)
        {
            return;
        }

        var series = EnsureSeriesCount(barChart, c, 2);
        var categories = data.Rows.Select(row => row.Service).ToList();
        SetStringReference(series[0].Element(c + "tx")?.Element(c + "strRef"), c, "Sheet1!$B$1", [data.CurrentPeriodLabel]);
        SetStringReference(series[0].Element(c + "cat")?.Element(c + "strRef"), c, $"Sheet1!$A$2:$A${data.Rows.Count + 1}", categories);
        SetNumberReference(series[0].Element(c + "val")?.Element(c + "numRef"), c, $"Sheet1!$B$2:$B${data.Rows.Count + 1}", data.Rows.Select(row => row.CurrentPaidClaims).ToList());
        SetSeriesFill(series[0], c, a, "BDE7DD");

        SetStringReference(series[1].Element(c + "tx")?.Element(c + "strRef"), c, "Sheet1!$C$1", [data.PriorPeriodLabel]);
        SetStringReference(series[1].Element(c + "cat")?.Element(c + "strRef"), c, $"Sheet1!$A$2:$A${data.Rows.Count + 1}", categories);
        SetNumberReference(series[1].Element(c + "val")?.Element(c + "numRef"), c, $"Sheet1!$C$2:$C${data.Rows.Count + 1}", data.Rows.Select(row => row.PriorPaidClaims).ToList());
        SetSeriesFill(series[1], c, a, "66C2B2");

        SanitizeChartExtensions(document, c);
        entry.Delete();
        var newEntry = archive.CreateEntry(chartPath, CompressionLevel.Optimal);
        using var stream = newEntry.Open();
        document.Save(stream);
    }

    private static void UpdateEmbeddedParamedicalBreakdownWorkbook(ZipArchive archive, string workbookPath, ParamedicalBreakdownSlideData data)
    {
        var entry = archive.GetEntry(workbookPath);
        if (entry is null)
        {
            return;
        }

        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("Sheet1");
        sheet.Cell(1, 1).Value = "Service";
        sheet.Cell(1, 2).Value = data.CurrentPeriodLabel;
        sheet.Cell(1, 3).Value = data.PriorPeriodLabel;
        sheet.Range(1, 1, 1, 3).Style.Font.Bold = true;

        for (var index = 0; index < data.Rows.Count; index++)
        {
            var rowNumber = index + 2;
            sheet.Cell(rowNumber, 1).Value = data.Rows[index].Service;
            sheet.Cell(rowNumber, 2).Value = data.Rows[index].CurrentPaidClaims;
            sheet.Cell(rowNumber, 3).Value = data.Rows[index].PriorPaidClaims;
        }

        sheet.Range(2, 2, data.Rows.Count + 1, 3).Style.NumberFormat.Format = "$#,##0.00";
        sheet.Columns(1, 3).AdjustToContents(8, 24);

        using var output = new MemoryStream();
        workbook.SaveAs(output);
        output.Position = 0;

        entry.Delete();
        var newEntry = archive.CreateEntry(workbookPath, CompressionLevel.Optimal);
        using var outputStream = newEntry.Open();
        output.CopyTo(outputStream);
    }

    private static bool UpdateDentalBreakdownSlide(ZipArchive archive, DentalBreakdownSlideData data)
    {
        var slidePath = FindSlidePathByTitle(archive, "Claims Breakdown: Dental");
        if (slidePath is null)
        {
            return false;
        }

        var entry = archive.GetEntry(slidePath) ?? throw new InvalidOperationException($"Slide part not found: {slidePath}");
        var document = XDocument.Parse(ReadZipEntryText(entry));
        var a = XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/main");

        foreach (var text in document.Descendants(a + "t"))
        {
            var value = text.Value;
            value = value.Replace("[reporting period]", data.CurrentPeriodLabel, StringComparison.OrdinalIgnoreCase);
            value = value.Replace("$[XX]", FormatMoneyWhole(data.TotalCurrentPaidClaims), StringComparison.OrdinalIgnoreCase);
            value = value.Replace("[$XX]", FormatMoneyWhole(data.TotalCurrentPaidClaims), StringComparison.OrdinalIgnoreCase);
            text.Value = value;
        }

        entry.Delete();
        var newEntry = archive.CreateEntry(slidePath, CompressionLevel.Optimal);
        using var stream = newEntry.Open();
        document.Save(stream);
        return true;
    }

    private static void UpdateDentalBreakdownChartXml(ZipArchive archive, string chartPath, DentalBreakdownSlideData data)
    {
        var entry = archive.GetEntry(chartPath) ?? throw new InvalidOperationException($"Chart part not found: {chartPath}");
        var document = XDocument.Parse(ReadZipEntryText(entry));
        var c = XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/chart");
        var a = XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/main");
        var barChart = document.Descendants(c + "barChart").FirstOrDefault();
        if (barChart is null)
        {
            return;
        }

        var series = EnsureSeriesCount(barChart, c, 2);
        var categories = data.Rows.Select(row => row.Category).ToList();
        SetStringReference(series[0].Element(c + "tx")?.Element(c + "strRef"), c, "Sheet1!$B$1", [data.CurrentPeriodLabel]);
        SetStringReference(series[0].Element(c + "cat")?.Element(c + "strRef"), c, $"Sheet1!$A$2:$A${data.Rows.Count + 1}", categories);
        SetNumberReference(series[0].Element(c + "val")?.Element(c + "numRef"), c, $"Sheet1!$B$2:$B${data.Rows.Count + 1}", data.Rows.Select(row => row.CurrentPaidClaims).ToList());
        SetSeriesFill(series[0], c, a, "BDE7DD");

        SetStringReference(series[1].Element(c + "tx")?.Element(c + "strRef"), c, "Sheet1!$C$1", [data.PriorPeriodLabel]);
        SetStringReference(series[1].Element(c + "cat")?.Element(c + "strRef"), c, $"Sheet1!$A$2:$A${data.Rows.Count + 1}", categories);
        SetNumberReference(series[1].Element(c + "val")?.Element(c + "numRef"), c, $"Sheet1!$C$2:$C${data.Rows.Count + 1}", data.Rows.Select(row => row.PriorPaidClaims).ToList());
        SetSeriesFill(series[1], c, a, "66C2B2");

        SanitizeChartExtensions(document, c);
        entry.Delete();
        var newEntry = archive.CreateEntry(chartPath, CompressionLevel.Optimal);
        using var stream = newEntry.Open();
        document.Save(stream);
    }

    private static void UpdateEmbeddedDentalBreakdownWorkbook(ZipArchive archive, string workbookPath, DentalBreakdownSlideData data)
    {
        var entry = archive.GetEntry(workbookPath);
        if (entry is null)
        {
            return;
        }

        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("Sheet1");
        sheet.Cell(1, 1).Value = "Category";
        sheet.Cell(1, 2).Value = data.CurrentPeriodLabel;
        sheet.Cell(1, 3).Value = data.PriorPeriodLabel;
        sheet.Range(1, 1, 1, 3).Style.Font.Bold = true;

        for (var index = 0; index < data.Rows.Count; index++)
        {
            var rowNumber = index + 2;
            sheet.Cell(rowNumber, 1).Value = data.Rows[index].Category;
            sheet.Cell(rowNumber, 2).Value = data.Rows[index].CurrentPaidClaims;
            sheet.Cell(rowNumber, 3).Value = data.Rows[index].PriorPaidClaims;
        }

        sheet.Range(2, 2, data.Rows.Count + 1, 3).Style.NumberFormat.Format = "$#,##0.00";
        sheet.Columns(1, 3).AdjustToContents(8, 24);

        using var output = new MemoryStream();
        workbook.SaveAs(output);
        output.Position = 0;

        entry.Delete();
        var newEntry = archive.CreateEntry(workbookPath, CompressionLevel.Optimal);
        using var outputStream = newEntry.Open();
        output.CopyTo(outputStream);
    }

    private static int UpdateGlobalDeckPlaceholders(ZipArchive archive, DeckGlobalPlaceholderData data)
    {
        var updatedSlides = 0;
        var slideEntries = archive.Entries
            .Where(entry => entry.FullName.StartsWith("ppt/slides/slide", StringComparison.OrdinalIgnoreCase) &&
                            entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
            .OrderBy(entry => SlideNumber(entry.FullName))
            .ToList();

        foreach (var entry in slideEntries)
        {
            var slidePath = entry.FullName;
            var document = XDocument.Parse(ReadZipEntryText(entry));
            var a = XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/main");
            var changed = false;

            foreach (var text in document.Descendants(a + "t"))
            {
                var original = text.Value;
                var value = ReplaceGlobalPlaceholderText(original, data);
                if (!string.Equals(original, value, StringComparison.Ordinal))
                {
                    text.Value = value;
                    changed = true;
                }
            }

            if (!changed)
            {
                continue;
            }

            entry.Delete();
            var newEntry = archive.CreateEntry(slidePath, CompressionLevel.Optimal);
            using var stream = newEntry.Open();
            document.Save(stream);
            updatedSlides++;
        }

        return updatedSlides;
    }

    private static string ReplaceGlobalPlaceholderText(string value, DeckGlobalPlaceholderData data)
    {
        var preparedDate = data.PreparedDate.ToString("MMMM d, yyyy", Culture);
        value = value.Replace("[Client Name]", data.ClientName, StringComparison.OrdinalIgnoreCase);
        value = value.Replace("[Client name]", data.ClientName, StringComparison.OrdinalIgnoreCase);
        value = value.Replace("[Enter Client Name]", data.ClientName, StringComparison.OrdinalIgnoreCase);
        value = value.Replace("[Enter Company Name]", data.ClientName, StringComparison.OrdinalIgnoreCase);
        value = value.Replace(
            "Reporting Period: [Policy Number]",
            string.IsNullOrWhiteSpace(data.PolicyNumber)
                ? $"Period: {data.CurrentPeriodLabel}"
                : $"Period: {data.CurrentPeriodLabel} | Policy: {data.PolicyNumber}",
            StringComparison.OrdinalIgnoreCase);
        value = value.Replace("[Carrier]", data.Carrier, StringComparison.OrdinalIgnoreCase);
        value = value.Replace("[Policy Number]", data.PolicyNumber, StringComparison.OrdinalIgnoreCase);
        value = value.Replace("[Contract Number]", data.PolicyNumber, StringComparison.OrdinalIgnoreCase);
        value = value.Replace("[00]", data.MonthCount.ToString("0", Culture), StringComparison.OrdinalIgnoreCase);
        value = value.Replace("[Enter Reporting Period]", data.CurrentPeriodLabel, StringComparison.OrdinalIgnoreCase);
        value = value.Replace("[Reporting Period]", data.CurrentPeriodLabel, StringComparison.OrdinalIgnoreCase);
        value = value.Replace("[reporting period]", data.CurrentPeriodLabel, StringComparison.OrdinalIgnoreCase);
        value = value.Replace("[current reporting period]", data.CurrentPeriodLabel, StringComparison.OrdinalIgnoreCase);
        value = value.Replace("[prior reporting period]", data.PriorPeriodLabel, StringComparison.OrdinalIgnoreCase);
        value = value.Replace("[Date of Presentation]", preparedDate, StringComparison.OrdinalIgnoreCase);
        value = value.Replace("[enter date]", preparedDate, StringComparison.OrdinalIgnoreCase);
        return value;
    }

    private static bool UpdatePerformanceSnapshotMetrics(ZipArchive archive, DeckGlobalPlaceholderData data)
    {
        var slidePath = FindSlidePathByTitle(archive, "Performance Snapshot");
        if (slidePath is null)
        {
            return false;
        }

        var entry = archive.GetEntry(slidePath) ?? throw new InvalidOperationException($"Slide part not found: {slidePath}");
        var document = XDocument.Parse(ReadZipEntryText(entry));
        var a = XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/main");
        var pendingMetric = "";
        var changed = false;

        foreach (var text in document.Descendants(a + "t"))
        {
            var value = text.Value;
            if (value.Contains("Employees Covered", StringComparison.OrdinalIgnoreCase))
            {
                pendingMetric = "Employees";
                continue;
            }

            if (value.Contains("Total Claims Paid", StringComparison.OrdinalIgnoreCase))
            {
                pendingMetric = "TotalClaims";
                continue;
            }

            if (value.Contains("Average Claims Per Employee", StringComparison.OrdinalIgnoreCase))
            {
                pendingMetric = "AverageClaims";
                continue;
            }

            if (value.Contains("[$XX]", StringComparison.OrdinalIgnoreCase))
            {
                text.Value = value.Replace("[$XX]", FormatMoneyWhole(data.TotalClaimsPaid), StringComparison.OrdinalIgnoreCase);
                changed = true;
                continue;
            }

            if (!value.Contains("[XX]", StringComparison.OrdinalIgnoreCase) || string.IsNullOrWhiteSpace(pendingMetric))
            {
                continue;
            }

            var replacement = pendingMetric switch
            {
                "Employees" => FormatLives(data.EmployeesCovered),
                "TotalClaims" => FormatMoneyWhole(data.TotalClaimsPaid),
                "AverageClaims" => FormatMoneyWhole(data.AverageClaimsPerEmployee),
                _ => ""
            };
            if (!string.IsNullOrWhiteSpace(replacement))
            {
                text.Value = value.Replace("[XX]", replacement, StringComparison.OrdinalIgnoreCase);
                changed = true;
            }

            pendingMetric = "";
        }

        if (!changed)
        {
            return false;
        }

        entry.Delete();
        var newEntry = archive.CreateEntry(slidePath, CompressionLevel.Optimal);
        using var stream = newEntry.Open();
        document.Save(stream);
        return true;
    }

    private static bool UpdateOverallPremiumClaimsSummarySlide(ZipArchive archive, OverallPremiumClaimsSummarySlideData data)
    {
        var slidePath = FindSlidePathByTextParts(archive, "Program Performance: Premium vs Claims Summary");
        if (slidePath is null)
        {
            return false;
        }

        var entry = archive.GetEntry(slidePath) ?? throw new InvalidOperationException($"Slide part not found: {slidePath}");
        var document = XDocument.Parse(ReadZipEntryText(entry));
        var a = XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/main");
        var highestLossRatio = data.Rows.OrderByDescending(row => row.LossRatio).FirstOrDefault();

        foreach (var text in document.Descendants(a + "t"))
        {
            var value = text.Value;
            value = value.Replace("[reporting period]", data.CurrentPeriodLabel, StringComparison.OrdinalIgnoreCase);
            value = value.Replace(
                "[Optional: Add any relevant notes highlighting the distribution of claims across the reporting period that may be important for this group.]",
                highestLossRatio is null
                    ? "Claims summary is based on the current reporting period."
                    : $"{highestLossRatio.Category} has the highest loss ratio at {FormatPercent(highestLossRatio.LossRatio)} for the current reporting period.",
                StringComparison.OrdinalIgnoreCase);
            text.Value = value;
        }

        entry.Delete();
        var newEntry = archive.CreateEntry(slidePath, CompressionLevel.Optimal);
        using var stream = newEntry.Open();
        document.Save(stream);
        return true;
    }

    private static void UpdateOverallPremiumClaimsSummaryChartXml(ZipArchive archive, string chartPath, OverallPremiumClaimsSummarySlideData data)
    {
        var entry = archive.GetEntry(chartPath) ?? throw new InvalidOperationException($"Chart part not found: {chartPath}");
        var document = XDocument.Parse(ReadZipEntryText(entry));
        var c = XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/chart");
        var barChart = document.Descendants(c + "barChart").FirstOrDefault();
        if (barChart is null)
        {
            return;
        }

        var series = EnsureSeriesCount(barChart, c, 3);
        var categories = data.Rows.Select(row => row.Category).ToList();
        SetStringReference(series[0].Element(c + "tx")?.Element(c + "strRef"), c, "Sheet1!$B$1", ["Paid Premiums"]);
        SetStringReference(series[0].Element(c + "cat")?.Element(c + "strRef"), c, $"Sheet1!$A$2:$A${data.Rows.Count + 1}", categories);
        SetNumberReference(series[0].Element(c + "val")?.Element(c + "numRef"), c, $"Sheet1!$B$2:$B${data.Rows.Count + 1}", data.Rows.Select(row => row.PaidPremiums).ToList(), "$#,##0");

        SetStringReference(series[1].Element(c + "tx")?.Element(c + "strRef"), c, "Sheet1!$C$1", ["Paid Claims"]);
        SetStringReference(series[1].Element(c + "cat")?.Element(c + "strRef"), c, $"Sheet1!$A$2:$A${data.Rows.Count + 1}", categories);
        SetNumberReference(series[1].Element(c + "val")?.Element(c + "numRef"), c, $"Sheet1!$C$2:$C${data.Rows.Count + 1}", data.Rows.Select(row => row.PaidClaims).ToList(), "$#,##0");

        SetStringReference(series[2].Element(c + "tx")?.Element(c + "strRef"), c, "Sheet1!$D$1", ["Loss Ratio"]);
        SetStringReference(series[2].Element(c + "cat")?.Element(c + "strRef"), c, $"Sheet1!$A$2:$A${data.Rows.Count + 1}", categories);
        SetNumberReference(series[2].Element(c + "val")?.Element(c + "numRef"), c, $"Sheet1!$D$2:$D${data.Rows.Count + 1}", data.Rows.Select(row => row.LossRatio).ToList(), "0.0%");

        SanitizeChartExtensions(document, c);
        entry.Delete();
        var newEntry = archive.CreateEntry(chartPath, CompressionLevel.Optimal);
        using var stream = newEntry.Open();
        document.Save(stream);
    }

    private static void UpdateEmbeddedOverallPremiumClaimsSummaryWorkbook(ZipArchive archive, string workbookPath, OverallPremiumClaimsSummarySlideData data)
    {
        var entry = archive.GetEntry(workbookPath);
        if (entry is null)
        {
            return;
        }

        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("Sheet1");
        sheet.Cell(1, 1).Value = "Coverage";
        sheet.Cell(1, 2).Value = "Paid Premiums";
        sheet.Cell(1, 3).Value = "Paid Claims";
        sheet.Cell(1, 4).Value = "Loss Ratio";
        sheet.Range(1, 1, 1, 4).Style.Font.Bold = true;

        for (var index = 0; index < data.Rows.Count; index++)
        {
            var rowNumber = index + 2;
            sheet.Cell(rowNumber, 1).Value = data.Rows[index].Category;
            sheet.Cell(rowNumber, 2).Value = data.Rows[index].PaidPremiums;
            sheet.Cell(rowNumber, 3).Value = data.Rows[index].PaidClaims;
            sheet.Cell(rowNumber, 4).Value = data.Rows[index].LossRatio;
        }

        sheet.Range(2, 2, data.Rows.Count + 1, 3).Style.NumberFormat.Format = "$#,##0";
        sheet.Range(2, 4, data.Rows.Count + 1, 4).Style.NumberFormat.Format = "0.0%";
        sheet.Columns(1, 4).AdjustToContents(8, 24);

        using var output = new MemoryStream();
        workbook.SaveAs(output);
        output.Position = 0;

        entry.Delete();
        var newEntry = archive.CreateEntry(workbookPath, CompressionLevel.Optimal);
        using var outputStream = newEntry.Open();
        output.CopyTo(outputStream);
    }

    private static bool UpdatePremiumClaimsSummarySlide(ZipArchive archive, PremiumClaimsSummarySlideData data)
    {
        var slidePath = FindSlidePathByTextParts(archive, "Premiums vs. Claims Summary", data.SlideCoverageTitle);
        if (slidePath is null)
        {
            return false;
        }

        var entry = archive.GetEntry(slidePath) ?? throw new InvalidOperationException($"Slide part not found: {slidePath}");
        var document = XDocument.Parse(ReadZipEntryText(entry));
        var a = XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/main");
        var numberCounter = 0;
        var percentCounter = 0;

        foreach (var text in document.Descendants(a + "t"))
        {
            var value = text.Value;
            value = value.Replace("[reporting period]", data.CurrentPeriodLabel, StringComparison.OrdinalIgnoreCase);
            value = value.Replace("[Reporting Period]", data.CurrentPeriodLabel, StringComparison.OrdinalIgnoreCase);

            if (value.Contains("[enter number]", StringComparison.OrdinalIgnoreCase))
            {
                numberCounter++;
                var replacement = numberCounter switch
                {
                    1 => FormatMoneyWhole(data.CurrentPremium),
                    2 => FormatMoneyWhole(data.CurrentPaidClaims),
                    3 => FormatPercent(data.CurrentLossRatio),
                    _ => ""
                };
                value = value.Replace("[enter number]", replacement, StringComparison.OrdinalIgnoreCase);
            }

            if (value.Contains("[XX]%", StringComparison.OrdinalIgnoreCase))
            {
                percentCounter++;
                var replacement = percentCounter switch
                {
                    1 => FormatPercent(data.CurrentLossRatio),
                    2 => FormatPercent(data.TrendFactor),
                    3 => FormatPercent(data.TotalExpenseFactor),
                    4 => FormatPercent(data.TargetLossRatio),
                    5 => FormatPercent(data.AnticipatedAdjustment),
                    _ => ""
                };
                value = value.Replace("[XX]%", replacement, StringComparison.OrdinalIgnoreCase);
            }

            text.Value = value;
        }

        entry.Delete();
        var newEntry = archive.CreateEntry(slidePath, CompressionLevel.Optimal);
        using var stream = newEntry.Open();
        document.Save(stream);
        return true;
    }

    private static void UpdatePremiumClaimsSummaryChartXml(ZipArchive archive, string chartPath, PremiumClaimsSummarySlideData data)
    {
        var entry = archive.GetEntry(chartPath) ?? throw new InvalidOperationException($"Chart part not found: {chartPath}");
        var document = XDocument.Parse(ReadZipEntryText(entry));
        var c = XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/chart");
        var a = XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/main");
        var barChart = document.Descendants(c + "barChart").FirstOrDefault();
        if (barChart is null)
        {
            return;
        }

        var series = EnsureSeriesCount(barChart, c, 2);
        var categories = data.ChartRows.Select(row => row.PeriodLabel).ToList();
        SetStringReference(series[0].Element(c + "tx")?.Element(c + "strRef"), c, "Sheet1!$B$1", [data.PremiumSeriesLabel]);
        SetStringReference(series[0].Element(c + "cat")?.Element(c + "strRef"), c, $"Sheet1!$A$2:$A${data.ChartRows.Count + 1}", categories);
        SetNumberReference(series[0].Element(c + "val")?.Element(c + "numRef"), c, $"Sheet1!$B$2:$B${data.ChartRows.Count + 1}", data.ChartRows.Select(row => row.Premium).ToList(), "$#,##0");
        SetSeriesFill(series[0], c, a, "004751");

        SetStringReference(series[1].Element(c + "tx")?.Element(c + "strRef"), c, "Sheet1!$C$1", ["Paid Claims"]);
        SetStringReference(series[1].Element(c + "cat")?.Element(c + "strRef"), c, $"Sheet1!$A$2:$A${data.ChartRows.Count + 1}", categories);
        SetNumberReference(series[1].Element(c + "val")?.Element(c + "numRef"), c, $"Sheet1!$C$2:$C${data.ChartRows.Count + 1}", data.ChartRows.Select(row => row.PaidClaims).ToList(), "$#,##0");
        SetSeriesFill(series[1], c, a, "CDEBF0");

        SanitizeChartExtensions(document, c);
        entry.Delete();
        var newEntry = archive.CreateEntry(chartPath, CompressionLevel.Optimal);
        using var stream = newEntry.Open();
        document.Save(stream);
    }

    private static void UpdateEmbeddedPremiumClaimsSummaryWorkbook(ZipArchive archive, string workbookPath, PremiumClaimsSummarySlideData data)
    {
        var entry = archive.GetEntry(workbookPath);
        if (entry is null)
        {
            return;
        }

        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("Sheet1");
        sheet.Cell(1, 1).Value = "Period";
        sheet.Cell(1, 2).Value = data.PremiumSeriesLabel;
        sheet.Cell(1, 3).Value = "Paid Claims";
        sheet.Range(1, 1, 1, 3).Style.Font.Bold = true;

        for (var index = 0; index < data.ChartRows.Count; index++)
        {
            var rowNumber = index + 2;
            sheet.Cell(rowNumber, 1).Value = data.ChartRows[index].PeriodLabel;
            sheet.Cell(rowNumber, 2).Value = data.ChartRows[index].Premium;
            sheet.Cell(rowNumber, 3).Value = data.ChartRows[index].PaidClaims;
        }

        sheet.Range(2, 2, data.ChartRows.Count + 1, 3).Style.NumberFormat.Format = "$#,##0";
        sheet.Columns(1, 3).AdjustToContents(8, 24);

        using var output = new MemoryStream();
        workbook.SaveAs(output);
        output.Position = 0;

        entry.Delete();
        var newEntry = archive.CreateEntry(workbookPath, CompressionLevel.Optimal);
        using var outputStream = newEntry.Open();
        output.CopyTo(outputStream);
    }

    private static void UpdateLossRatioBreakEvenChartXml(ZipArchive archive, string chartPath, IReadOnlyList<LossRatioBreakEvenChartRow> rows)
    {
        var entry = archive.GetEntry(chartPath) ?? throw new InvalidOperationException($"Chart part not found: {chartPath}");
        var document = XDocument.Parse(ReadZipEntryText(entry));
        var c = XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/chart");
        var a = XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/main");
        var barChart = document.Descendants(c + "barChart").FirstOrDefault();
        if (barChart is null)
        {
            return;
        }

        var series = EnsureSeriesCount(barChart, c, 2);
        var categories = rows.Select(row => row.PeriodLabel).ToList();
        SetStringReference(series[0].Element(c + "tx")?.Element(c + "strRef"), c, "Sheet1!$B$1", ["Net L/R"]);
        SetStringReference(series[0].Element(c + "cat")?.Element(c + "strRef"), c, $"Sheet1!$A$2:$A${rows.Count + 1}", categories);
        SetNumberReference(series[0].Element(c + "val")?.Element(c + "numRef"), c, $"Sheet1!$B$2:$B${rows.Count + 1}", rows.Select(row => row.NetLossRatio).ToList(), "0.0%");
        SetSeriesFill(series[0], c, a, "5B9BD5");

        SetStringReference(series[1].Element(c + "tx")?.Element(c + "strRef"), c, "Sheet1!$C$1", ["Break Even (TLR)"]);
        SetStringReference(series[1].Element(c + "cat")?.Element(c + "strRef"), c, $"Sheet1!$A$2:$A${rows.Count + 1}", categories);
        SetNumberReference(series[1].Element(c + "val")?.Element(c + "numRef"), c, $"Sheet1!$C$2:$C${rows.Count + 1}", rows.Select(row => row.BreakEven).ToList(), "0.0%");
        SetSeriesNoFill(series[1], c, a);
        SetSeriesDataLabelsDeleted(series[1], c);
        EnsureLinearTrendline(series[1], c, a, "004751");

        SanitizeChartExtensions(document, c);
        entry.Delete();
        var newEntry = archive.CreateEntry(chartPath, CompressionLevel.Optimal);
        using var stream = newEntry.Open();
        document.Save(stream);
    }

    private static void UpdateEmbeddedLossRatioBreakEvenWorkbook(ZipArchive archive, string workbookPath, IReadOnlyList<LossRatioBreakEvenChartRow> rows)
    {
        var entry = archive.GetEntry(workbookPath);
        if (entry is null)
        {
            return;
        }

        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("Sheet1");
        sheet.Cell(1, 1).Value = "Period";
        sheet.Cell(1, 2).Value = "Net L/R";
        sheet.Cell(1, 3).Value = "Break Even (TLR)";
        sheet.Range(1, 1, 1, 3).Style.Font.Bold = true;

        for (var index = 0; index < rows.Count; index++)
        {
            var rowNumber = index + 2;
            sheet.Cell(rowNumber, 1).Value = rows[index].PeriodLabel;
            sheet.Cell(rowNumber, 2).Value = rows[index].NetLossRatio;
            sheet.Cell(rowNumber, 3).Value = rows[index].BreakEven;
        }

        sheet.Range(2, 2, rows.Count + 1, 3).Style.NumberFormat.Format = "0.0%";
        sheet.Columns(1, 3).AdjustToContents(8, 24);

        using var output = new MemoryStream();
        workbook.SaveAs(output);
        output.Position = 0;

        entry.Delete();
        var newEntry = archive.CreateEntry(workbookPath, CompressionLevel.Optimal);
        using var outputStream = newEntry.Open();
        output.CopyTo(outputStream);
    }

    private static int UpdateScheduleOfBenefitsSlides(ZipArchive archive, ScheduleOfBenefitsDeckData data)
    {
        if (data.Rows.Count == 0)
        {
            return 0;
        }

        var updated = 0;
        for (var page = 1; page <= 3; page++)
        {
            var pageRows = data.Rows
                .Where(row => row.SlidePage == page)
                .OrderBy(row => row.SectionOrder)
                .ToList();
            if (pageRows.Count == 0)
            {
                continue;
            }

            if (UpdateScheduleOfBenefitsSlide(archive, data.ClientName, page, pageRows))
            {
                updated++;
            }
        }

        return updated;
    }

    private static bool UpdateScheduleOfBenefitsSlide(
        ZipArchive archive,
        string clientName,
        int page,
        IReadOnlyList<ScheduleOfBenefitsRow> rows)
    {
        var slidePath = FindSlidePathByTextParts(archive, "Schedule of Benefits", $"{page} of 3");
        if (slidePath is null)
        {
            return false;
        }

        var entry = archive.GetEntry(slidePath) ?? throw new InvalidOperationException($"Slide part not found: {slidePath}");
        var document = XDocument.Parse(ReadZipEntryText(entry));
        var a = XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/main");
        var table = document.Descendants(a + "tbl").FirstOrDefault();
        if (table is null)
        {
            return false;
        }

        UpdateScheduleSlideHeaderText(document, a, clientName, page);
        UpdateScheduleTable(table, a, clientName, page, rows);

        entry.Delete();
        var newEntry = archive.CreateEntry(slidePath, CompressionLevel.Optimal);
        using var stream = newEntry.Open();
        document.Save(stream);
        return true;
    }

    private static void UpdateScheduleSlideHeaderText(XDocument document, XNamespace a, string clientName, int page)
    {
        var textNodes = document.Descendants(a + "t").ToList();
        var scheduleIndex = textNodes.FindIndex(text =>
            text.Value.Contains("Schedule of Benefits", StringComparison.OrdinalIgnoreCase));

        if (scheduleIndex > 0)
        {
            var titleNode = textNodes
                .Take(scheduleIndex)
                .FirstOrDefault(text =>
                    !string.IsNullOrWhiteSpace(text.Value) &&
                    !text.Value.Contains("Westland", StringComparison.OrdinalIgnoreCase) &&
                    !text.Value.Contains("Benefits", StringComparison.OrdinalIgnoreCase));
            if (titleNode is not null)
            {
                titleNode.Value = clientName;
            }
        }

        foreach (var text in textNodes)
        {
            if (text.Value.Contains("Schedule of Benefits", StringComparison.OrdinalIgnoreCase))
            {
                text.Value = $"Schedule of Benefits | {page} of 3";
                break;
            }
        }
    }

    private static void UpdateScheduleTable(
        XElement table,
        XNamespace a,
        string clientName,
        int page,
        IReadOnlyList<ScheduleOfBenefitsRow> rows)
    {
        var tableRows = table.Elements(a + "tr").ToList();
        if (tableRows.Count == 0)
        {
            return;
        }

        var firstScheduleRow = rows.First();
        var rowIndex = 0;
        rowIndex = TryWriteScheduleHeaderRow(tableRows, rowIndex, a, clientName, firstScheduleRow);
        rowIndex = TryWriteScheduleClassRow(tableRows, rowIndex, a, firstScheduleRow);

        foreach (var scheduleRow in rows)
        {
            if (rowIndex >= tableRows.Count)
            {
                break;
            }

            var cells = tableRows[rowIndex].Elements(a + "tc").ToList();
            var templateCells = tableRows[Math.Max(0, Math.Min(rowIndex, tableRows.Count - 1))].Elements(a + "tc").ToList();
            if (cells.Count == 0)
            {
                rowIndex++;
                continue;
            }

            if (cells.Count == 1)
            {
                SetTableCellText(cells[0], a, $"{scheduleRow.SectionName}: {CompactPowerPointText(scheduleRow.DetailText)}", TemplateRunProperties(templateCells, 0, a));
            }
            else
            {
                SetTableCellText(cells[0], a, scheduleRow.SectionName, TemplateRunProperties(templateCells, 0, a));
                SetTableCellText(cells[1], a, CompactPowerPointText(scheduleRow.DetailText), TemplateRunProperties(templateCells, 1, a));
                for (var cellIndex = 2; cellIndex < cells.Count; cellIndex++)
                {
                    SetTableCellText(cells[cellIndex], a, "", TemplateRunProperties(templateCells, cellIndex, a));
                }
            }

            rowIndex++;
        }

        while (rowIndex < tableRows.Count)
        {
            foreach (var cell in tableRows[rowIndex].Elements(a + "tc"))
            {
                SetTableCellText(cell, a, "");
            }

            rowIndex++;
        }
    }

    private static int TryWriteScheduleHeaderRow(
        IReadOnlyList<XElement> rows,
        int rowIndex,
        XNamespace a,
        string clientName,
        ScheduleOfBenefitsRow firstScheduleRow)
    {
        if (rowIndex >= rows.Count)
        {
            return rowIndex;
        }

        var cells = rows[rowIndex].Elements(a + "tc").ToList();
        if (cells.Count == 0)
        {
            return rowIndex;
        }

        SetTableCellText(cells[0], a, clientName);
        for (var cellIndex = 1; cellIndex < cells.Count; cellIndex++)
        {
            SetTableCellText(cells[cellIndex], a, cellIndex == cells.Count - 1 ? firstScheduleRow.Carrier : "");
        }

        return rowIndex + 1;
    }

    private static int TryWriteScheduleClassRow(
        IReadOnlyList<XElement> rows,
        int rowIndex,
        XNamespace a,
        ScheduleOfBenefitsRow firstScheduleRow)
    {
        if (rowIndex >= rows.Count)
        {
            return rowIndex;
        }

        var cells = rows[rowIndex].Elements(a + "tc").ToList();
        if (cells.Count == 0)
        {
            return rowIndex;
        }

        SetTableCellText(cells[0], a, $"Class: {firstScheduleRow.ClassName}");
        if (cells.Count > 1)
        {
            SetTableCellText(cells[1], a, FirstNonBlank(firstScheduleRow.Carrier, "Carrier"));
        }

        for (var cellIndex = 2; cellIndex < cells.Count; cellIndex++)
        {
            SetTableCellText(cells[cellIndex], a, "");
        }

        return rowIndex + 1;
    }

    private static bool UpdateMonthlyBreakdownSlide(ZipArchive archive, MonthlyBreakdownSlideData data)
    {
        var slidePath = FindSlidePathByTitle(archive, data.SlideTitle);
        if (slidePath is null)
        {
            return false;
        }

        var entry = archive.GetEntry(slidePath) ?? throw new InvalidOperationException($"Slide part not found: {slidePath}");
        var document = XDocument.Parse(ReadZipEntryText(entry));
        var a = XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/main");
        var table = document.Descendants(a + "tbl").FirstOrDefault();
        if (table is null)
        {
            return false;
        }

        UpdateMonthlyBreakdownTable(table, a, data);
        UpdateMonthlyBreakdownNarrative(document, a, data);

        entry.Delete();
        var newEntry = archive.CreateEntry(slidePath, CompressionLevel.Optimal);
        using var stream = newEntry.Open();
        document.Save(stream);
        return true;
    }

    private static void UpdateMonthlyBreakdownTable(XElement table, XNamespace a, MonthlyBreakdownSlideData data)
    {
        var rows = table.Elements(a + "tr").ToList();
        if (rows.Count < 14)
        {
            return;
        }

        for (var index = 0; index < 12; index++)
        {
            var row = rows[index + 1];
            var cells = row.Elements(a + "tc").ToList();
            var templateCells = rows[1].Elements(a + "tc").ToList();
            if (cells.Count < 5)
            {
                continue;
            }

            if (index < data.Rows.Count)
            {
                var monthlyRow = data.Rows[index];
                SetTableCellText(cells[0], a, monthlyRow.Month.ToString("MMMM", Culture), TemplateRunProperties(templateCells, 0, a));
                SetTableCellText(cells[1], a, FormatMoneyWhole(monthlyRow.Premium), TemplateRunProperties(templateCells, 1, a));
                SetTableCellText(cells[2], a, FormatMoneyWhole(monthlyRow.PaidClaims), TemplateRunProperties(templateCells, 2, a));
                SetTableCellText(cells[3], a, FormatPercent(monthlyRow.LossRatio), TemplateRunProperties(templateCells, 3, a));
                SetTableCellText(cells[4], a, FormatLives(monthlyRow.EmployeeCount), TemplateRunProperties(templateCells, 4, a));
            }
            else
            {
                for (var cellIndex = 0; cellIndex < Math.Min(cells.Count, 5); cellIndex++)
                {
                    SetTableCellText(cells[cellIndex], a, "");
                }
            }
        }

        var totalCells = rows[13].Elements(a + "tc").ToList();
        if (totalCells.Count >= 5)
        {
            SetTableCellText(totalCells[0], a, "");
            SetTableCellText(totalCells[1], a, FormatMoneyWhole(data.TotalPremium));
            SetTableCellText(totalCells[2], a, FormatMoneyWhole(data.TotalPaidClaims));
            SetTableCellText(totalCells[3], a, FormatPercent(data.TotalLossRatio));
            SetTableCellText(totalCells[4], a, FormatLives(data.AverageEmployeeCount));
        }

        var headerCells = rows[0].Elements(a + "tc").ToList();
        if (headerCells.Count >= 2)
        {
            SetTableCellText(headerCells[1], a, data.PremiumHeader);
        }
    }

    private static XElement? TemplateRunProperties(IReadOnlyList<XElement> cells, int cellIndex, XNamespace a)
    {
        return cellIndex < cells.Count ? cells[cellIndex].Descendants(a + "rPr").FirstOrDefault() : null;
    }

    private static void SetTableCellText(XElement cell, XNamespace a, string value, XElement? templateRunProperties = null)
    {
        var textNodes = cell.Descendants(a + "t").ToList();
        if (textNodes.Count > 0)
        {
            textNodes[0].Value = value;
            foreach (var textNode in textNodes.Skip(1))
            {
                textNode.Value = "";
            }

            return;
        }

        if (string.IsNullOrEmpty(value))
        {
            return;
        }

        var paragraph = cell.Descendants(a + "p").FirstOrDefault();
        if (paragraph is null)
        {
            return;
        }

        var endProperties = paragraph.Element(a + "endParaRPr");
        var runProperties = templateRunProperties is not null
            ? new XElement(templateRunProperties)
            : endProperties is null
            ? new XElement(a + "rPr")
            : new XElement(a + "rPr", endProperties.Attributes(), endProperties.Elements());
        var run = new XElement(a + "r", runProperties, new XElement(a + "t", value));
        if (endProperties is null)
        {
            paragraph.Add(run);
        }
        else
        {
            endProperties.AddBeforeSelf(run);
        }
    }

    private static void UpdateMonthlyBreakdownNarrative(XDocument document, XNamespace a, MonthlyBreakdownSlideData data)
    {
        var moneyCounter = 0;
        foreach (var text in document.Descendants(a + "t"))
        {
            var value = text.Value;
            if (data.Coverage.Equals("EHC", StringComparison.OrdinalIgnoreCase))
            {
                value = value.Replace("[service / category]", data.TopCategory, StringComparison.OrdinalIgnoreCase);
                value = value.Replace("[Category]", data.TopCategory, StringComparison.OrdinalIgnoreCase);
                value = value.Replace("(generally aligned with / above / below)", LossRatioPosition(data.TotalLossRatio), StringComparison.OrdinalIgnoreCase);
                value = value.Replace("[X%]", FormatPercent(data.TotalLossRatio), StringComparison.OrdinalIgnoreCase);
                value = value.Replace("[$X per employee / per member]", $"{FormatMoneyWhole(data.AverageUtilization)} per employee", StringComparison.OrdinalIgnoreCase);
                value = value.Replace("(stable / trending upward / trending downward)", TrendDirection(data.TotalPaidClaims, data.PriorPaidClaims), StringComparison.OrdinalIgnoreCase);
                value = value.Replace("(the previous period / the prior year)", "the prior year", StringComparison.OrdinalIgnoreCase);

                if (value.Contains("[$X]", StringComparison.OrdinalIgnoreCase))
                {
                    moneyCounter++;
                    value = value.Replace("[$X]", FormatMoneyWhole(moneyCounter == 1 ? data.TopCategoryPaidClaims : data.TotalPaidClaims), StringComparison.OrdinalIgnoreCase);
                }
            }
            else
            {
                value = value.Replace("(basic services / preventive and diagnostic services / major restorative services)", data.TopCategory, StringComparison.OrdinalIgnoreCase);
                value = value.Replace("[month]", data.HighestClaimsMonth?.ToString("MMMM", Culture) ?? "the highest month", StringComparison.OrdinalIgnoreCase);
                value = value.Replace("(increased routine care / timing of major services / normal month-to-month variation)", "normal month-to-month variation", StringComparison.OrdinalIgnoreCase);
                value = value.Replace("(in line with / higher than / lower than)", LossRatioPosition(data.TotalLossRatio), StringComparison.OrdinalIgnoreCase);
                value = value.Replace("[X%]", FormatPercent(data.TotalLossRatio), StringComparison.OrdinalIgnoreCase);
                value = value.Replace("(stable / elevated / below expected levels)", UtilizationLevel(data.TotalLossRatio), StringComparison.OrdinalIgnoreCase);
                value = value.Replace("(this size / similar workforce demographics)", "this size", StringComparison.OrdinalIgnoreCase);

                if (value.Contains("[$X]", StringComparison.OrdinalIgnoreCase))
                {
                    moneyCounter++;
                    value = value.Replace("[$X]", FormatMoneyWhole(moneyCounter == 1 ? data.TotalPaidClaims : data.HighestClaimsAmount), StringComparison.OrdinalIgnoreCase);
                }
            }

            text.Value = value;
        }
    }

    private static string? FindSlidePathByTitle(ZipArchive archive, string slideTitle)
    {
        return FindSlidePathsByTitle(archive, slideTitle).FirstOrDefault();
    }

    private static List<string> FindSlidePathsByTitle(ZipArchive archive, string slideTitle)
    {
        var normalizedTitle = NormalizeText(slideTitle);
        return archive.Entries
            .Where(entry => entry.FullName.StartsWith("ppt/slides/slide", StringComparison.OrdinalIgnoreCase) &&
                            entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
            .Where(entry =>
            {
                var document = XDocument.Parse(ReadZipEntryText(entry));
                var a = XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/main");
                var visibleText = string.Join("", document.Descendants(a + "t").Select(element => element.Value));
                return NormalizeText(visibleText).Contains(normalizedTitle, StringComparison.OrdinalIgnoreCase);
            })
            .Select(entry => entry.FullName)
            .OrderBy(SlideNumber)
            .ToList();
    }

    private static List<string> FindChartPathsBySlideTitle(ZipArchive archive, string slideTitle)
    {
        var chartPaths = new List<string>();
        var rel = XNamespace.Get("http://schemas.openxmlformats.org/package/2006/relationships");

        foreach (var slidePath in FindSlidePathsByTitle(archive, slideTitle))
        {
            var relsPath = $"ppt/slides/_rels/{Path.GetFileName(slidePath)}.rels";
            var relsEntry = archive.GetEntry(relsPath);
            if (relsEntry is null)
            {
                continue;
            }

            var document = XDocument.Parse(ReadZipEntryText(relsEntry));
            foreach (var relationship in document.Root?.Elements(rel + "Relationship") ?? [])
            {
                if (((string?)relationship.Attribute("Type"))?.EndsWith("/chart", StringComparison.OrdinalIgnoreCase) != true)
                {
                    continue;
                }

                var target = (string?)relationship.Attribute("Target");
                if (!string.IsNullOrWhiteSpace(target))
                {
                    chartPaths.Add(ResolvePackagePath(slidePath, target));
                }
            }
        }

        return chartPaths.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
    }

    private static string? FindSlidePathByTextParts(ZipArchive archive, params string[] textParts)
    {
        return FindSlidePathsByTextParts(archive, textParts).FirstOrDefault();
    }

    private static List<string> FindChartPathsBySlideTextParts(ZipArchive archive, params string[] textParts)
    {
        var chartPaths = new List<string>();
        var rel = XNamespace.Get("http://schemas.openxmlformats.org/package/2006/relationships");

        foreach (var slidePath in FindSlidePathsByTextParts(archive, textParts))
        {
            var relsPath = $"ppt/slides/_rels/{Path.GetFileName(slidePath)}.rels";
            var relsEntry = archive.GetEntry(relsPath);
            if (relsEntry is null)
            {
                continue;
            }

            var document = XDocument.Parse(ReadZipEntryText(relsEntry));
            foreach (var relationship in document.Root?.Elements(rel + "Relationship") ?? [])
            {
                if (((string?)relationship.Attribute("Type"))?.EndsWith("/chart", StringComparison.OrdinalIgnoreCase) != true)
                {
                    continue;
                }

                var target = (string?)relationship.Attribute("Target");
                if (!string.IsNullOrWhiteSpace(target))
                {
                    chartPaths.Add(ResolvePackagePath(slidePath, target));
                }
            }
        }

        return chartPaths.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
    }

    private static List<string> FindSlidePathsByTextParts(ZipArchive archive, params string[] textParts)
    {
        var normalizedParts = textParts
            .Select(NormalizeText)
            .Where(part => part.Length > 0)
            .ToList();

        return archive.Entries
            .Where(entry => entry.FullName.StartsWith("ppt/slides/slide", StringComparison.OrdinalIgnoreCase) &&
                            entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
            .Where(entry =>
            {
                var document = XDocument.Parse(ReadZipEntryText(entry));
                var a = XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/main");
                var visibleText = NormalizeText(string.Join("", document.Descendants(a + "t").Select(element => element.Value)));
                return normalizedParts.All(part => visibleText.Contains(part, StringComparison.OrdinalIgnoreCase));
            })
            .Select(entry => entry.FullName)
            .OrderBy(SlideNumber)
            .ToList();
    }

    private static string? FindEmbeddedWorkbookPath(ZipArchive archive, string chartPath)
    {
        var chartFileName = Path.GetFileName(chartPath);
        var relsPath = $"ppt/charts/_rels/{chartFileName}.rels";
        var entry = archive.GetEntry(relsPath);
        if (entry is null)
        {
            return null;
        }

        var rel = XNamespace.Get("http://schemas.openxmlformats.org/package/2006/relationships");
        var document = XDocument.Parse(ReadZipEntryText(entry));
        var relationship = document.Root?.Elements(rel + "Relationship")
            .FirstOrDefault(element => ((string?)element.Attribute("Type"))?.EndsWith("/package", StringComparison.OrdinalIgnoreCase) == true);
        var targetPath = (string?)relationship?.Attribute("Target");
        return string.IsNullOrWhiteSpace(targetPath)
            ? null
            : ResolvePackagePath(chartPath, targetPath);
    }

    private static void SetStringReference(XElement? stringReference, XNamespace c, string formula, IReadOnlyList<string> values)
    {
        if (stringReference is null)
        {
            return;
        }

        SetFormula(stringReference, c, formula);
        var cache = EnsureChild(stringReference, c + "strCache");
        ReplaceCachePoints(cache, c, values.Select(value => (object)value).ToList());
    }

    private static void SetNumberReference(
        XElement? numberReference,
        XNamespace c,
        string formula,
        IReadOnlyList<decimal> values,
        string formatCode = "$#,##0.00")
    {
        if (numberReference is null)
        {
            return;
        }

        SetFormula(numberReference, c, formula);
        var cache = EnsureChild(numberReference, c + "numCache");
        cache.Elements(c + "formatCode").Remove();
        var formatCodeElement = new XElement(c + "formatCode");
        formatCodeElement.Value = formatCode;
        cache.AddFirst(formatCodeElement);
        ReplaceCachePoints(cache, c, values.Select(value => (object)value).ToList());
    }

    private static void SetFormula(XElement reference, XNamespace c, string formula)
    {
        var formulaElement = reference.Element(c + "f");
        if (formulaElement is null)
        {
            formulaElement = new XElement(c + "f");
            reference.AddFirst(formulaElement);
        }

        formulaElement.Value = formula;
    }

    private static XElement EnsureChild(XElement parent, XName name)
    {
        var child = parent.Element(name);
        if (child is not null)
        {
            return child;
        }

        child = new XElement(name);
        parent.Add(child);
        return child;
    }

    private static void ReplaceCachePoints(XElement cache, XNamespace c, IReadOnlyList<object> values)
    {
        cache.Elements(c + "ptCount").Remove();
        cache.Elements(c + "pt").Remove();
        var pointCount = new XElement(c + "ptCount", new XAttribute("val", values.Count));
        var formatCode = cache.Element(c + "formatCode");
        if (formatCode is null)
        {
            cache.AddFirst(pointCount);
        }
        else
        {
            formatCode.AddAfterSelf(pointCount);
        }

        for (var index = 0; index < values.Count; index++)
        {
            var value = values[index] is decimal decimalValue
                ? decimalValue.ToString("0.########", Culture)
                : values[index].ToString() ?? "";

            cache.Add(new XElement(
                c + "pt",
                new XAttribute("idx", index),
                new XElement(c + "v", value)));
        }
    }

    private static string ReadZipEntryText(ZipArchiveEntry entry)
    {
        using var stream = entry.Open();
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }

    private static string ResolvePackagePath(string sourcePath, string targetPath)
    {
        var sourceDirectory = Path.GetDirectoryName(sourcePath)?.Replace('\\', '/') ?? "";
        var combined = string.IsNullOrWhiteSpace(sourceDirectory)
            ? targetPath
            : $"{sourceDirectory}/{targetPath}";
        var segments = combined.Split(['/', '\\'], StringSplitOptions.RemoveEmptyEntries);
        var stack = new List<string>();

        foreach (var segment in segments)
        {
            if (segment == ".")
            {
                continue;
            }

            if (segment == "..")
            {
                if (stack.Count > 0)
                {
                    stack.RemoveAt(stack.Count - 1);
                }

                continue;
            }

            stack.Add(segment);
        }

        return string.Join("/", stack);
    }

    private static int SlideNumber(string slidePath)
    {
        var fileName = Path.GetFileNameWithoutExtension(slidePath);
        return int.TryParse(fileName.Replace("slide", "", StringComparison.OrdinalIgnoreCase), out var number)
            ? number
            : int.MaxValue;
    }

    private static decimal SumClaims(IEnumerable<MonthlyExperienceRow> rows, string periodId, string coverage)
    {
        return rows
            .Where(row => row.PeriodId.Equals(periodId, StringComparison.OrdinalIgnoreCase) &&
                          row.Coverage.Equals(coverage, StringComparison.OrdinalIgnoreCase))
            .Sum(row => row.PaidClaims);
    }

    private static decimal AverageLives(IEnumerable<MonthlyExperienceRow> rows, string periodId, string coverage)
    {
        var lives = rows
            .Where(row => row.PeriodId.Equals(periodId, StringComparison.OrdinalIgnoreCase) &&
                          row.Coverage.Equals(coverage, StringComparison.OrdinalIgnoreCase) &&
                          row.Lives > 0m)
            .Select(row => row.Lives)
            .ToList();
        return lives.Count == 0 ? 0m : lives.Average();
    }

    private static decimal AverageCombinedLives(IEnumerable<MonthlyExperienceRow> rows, string periodId)
    {
        var lives = rows
            .Where(row => row.PeriodId.Equals(periodId, StringComparison.OrdinalIgnoreCase) &&
                          (row.Coverage.Equals("EHC", StringComparison.OrdinalIgnoreCase) ||
                           row.Coverage.Equals("Dental", StringComparison.OrdinalIgnoreCase)) &&
                          row.Lives > 0m)
            .GroupBy(row => row.Month)
            .Select(group => group.Max(row => row.Lives))
            .ToList();
        return lives.Count == 0 ? 0m : lives.Average();
    }

    private static decimal SafeDivide(decimal numerator, decimal denominator)
    {
        return denominator == 0m ? 0m : numerator / denominator;
    }

    private static string PlaceholderText(
        IReadOnlyDictionary<string, PowerPointPlaceholderValue> placeholders,
        string key,
        string fallback)
    {
        return placeholders.TryGetValue(key, out var value) && !string.IsNullOrWhiteSpace(value.Text)
            ? value.Text
            : fallback;
    }

    private static decimal PlaceholderNumber(
        IReadOnlyDictionary<string, PowerPointPlaceholderValue> placeholders,
        string key,
        decimal fallback)
    {
        return placeholders.TryGetValue(key, out var value)
            ? value.Number
            : fallback;
    }

    private static string CellText(IXLCell cell)
    {
        try
        {
            var text = ParsingHelpers.Clean(cell.GetFormattedString());
            return text.Length > 0 ? text : ParsingHelpers.Clean(cell.GetString());
        }
        catch
        {
            return "";
        }
    }

    private static string CellMultilineText(IXLCell cell)
    {
        try
        {
            var text = cell.GetString();
            if (string.IsNullOrWhiteSpace(text))
            {
                text = cell.GetFormattedString();
            }

            return NormalizeMultiline(text);
        }
        catch
        {
            return "";
        }
    }

    private static string NormalizeMultiline(string value)
    {
        var lines = value
            .Replace("\r\n", "\n", StringComparison.Ordinal)
            .Replace('\r', '\n')
            .Split('\n')
            .Select(line => RegexNormalizeWhitespace(line.Trim()))
            .Where(line => line.Length > 0)
            .ToList();

        return string.Join("\n", lines);
    }

    private static string RegexNormalizeWhitespace(string value)
    {
        return string.Join(" ", value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
    }

    private static string CompactPowerPointText(string value)
    {
        var cleaned = NormalizeMultiline(value)
            .Replace("\uf0a1", "-", StringComparison.Ordinal)
            .Trim();
        const int maxLength = 760;
        return cleaned.Length <= maxLength ? cleaned : cleaned[..maxLength].TrimEnd() + "...";
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
                $"claims-ppt-export-{Guid.NewGuid():N}{Path.GetExtension(path)}");

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

    private static decimal ChartCellNumber(IXLCell cell)
    {
        try
        {
            if (!string.IsNullOrWhiteSpace(cell.FormulaA1) && cell.CachedValue.IsNumber)
            {
                return (decimal)cell.CachedValue.GetNumber();
            }
        }
        catch
        {
            // Fall through to the normal reader when the workbook has no usable cached value.
        }

        return CellNumber(cell);
    }

    private static decimal CellNumber(IXLCell cell)
    {
        try
        {
            if (cell.TryGetValue<decimal>(out var decimalValue))
            {
                return decimalValue;
            }

            return ParsingHelpers.ParseDecimal(cell.GetFormattedString()) ?? 0m;
        }
        catch
        {
            return 0m;
        }
    }

    private static DateOnly CellDate(IXLCell cell)
    {
        try
        {
            if (cell.TryGetValue<DateTime>(out var date))
            {
                return DateOnly.FromDateTime(date);
            }
        }
        catch
        {
            return DateOnly.FromDateTime(DateTime.Today);
        }

        return DateOnly.FromDateTime(DateTime.Today);
    }

    private static string FormatLives(decimal value)
    {
        return value == 0m ? "0" : Math.Round(value, 0, MidpointRounding.AwayFromZero).ToString("0", Culture);
    }

    private static string FormatMoney(decimal value)
    {
        return value.ToString("$#,##0.00", Culture);
    }

    private static string FormatMoneyWhole(decimal value)
    {
        return value.ToString("$#,##0", Culture);
    }

    private static string FormatPercent(decimal value)
    {
        return value.ToString("0.0%", Culture);
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

    private static string LossRatioPosition(decimal lossRatio)
    {
        if (lossRatio > 1.05m)
        {
            return "above";
        }

        if (lossRatio < 0.95m)
        {
            return "below";
        }

        return "generally aligned with";
    }

    private static string TrendDirection(decimal current, decimal prior)
    {
        if (prior == 0m)
        {
            return "stable";
        }

        var change = current / prior - 1m;
        if (change > 0.05m)
        {
            return "trending upward";
        }

        if (change < -0.05m)
        {
            return "trending downward";
        }

        return "stable";
    }

    private static string SpendingTrend(decimal current, decimal prior)
    {
        if (prior == 0m)
        {
            return "stable";
        }

        var change = current / prior - 1m;
        if (change > 0.05m)
        {
            return "increasing";
        }

        if (change < -0.05m)
        {
            return "moderating";
        }

        return "stable";
    }

    private static string CategoryChangeDescription(ClaimsDistributionRow? row)
    {
        if (row is null || row.PriorPaidClaims == 0m)
        {
            return "normal variation in utilization";
        }

        var change = row.CurrentPaidClaims / row.PriorPaidClaims - 1m;
        return Math.Abs(change) > 0.1m
            ? "increased activity in specific areas"
            : "normal variation in utilization";
    }

    private static string UtilizationComparison(decimal current, decimal prior)
    {
        if (prior == 0m)
        {
            return "consistent with";
        }

        var change = current / prior - 1m;
        if (change > 0.05m)
        {
            return "slightly above";
        }

        if (change < -0.05m)
        {
            return "slightly below";
        }

        return "consistent with";
    }

    private static string UtilizationLevel(decimal lossRatio)
    {
        if (lossRatio > 1m)
        {
            return "elevated";
        }

        if (lossRatio < 0.6m)
        {
            return "below expected levels";
        }

        return "stable";
    }

    private static string NormalizeText(string value)
    {
        return string.Concat(value.Where(char.IsLetterOrDigit)).ToUpperInvariant();
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

    private sealed record PowerPointPlaceholderValue(string Text, decimal Number);

    private sealed record DeckGlobalPlaceholderData(
        string ClientName,
        string Carrier,
        string PolicyNumber,
        int MonthCount,
        string CurrentPeriodLabel,
        string PriorPeriodLabel,
        DateOnly PreparedDate,
        decimal EmployeesCovered,
        decimal TotalClaimsPaid,
        decimal AverageClaimsPerEmployee);

    private sealed record PremiumClaimsChartRow(string PeriodLabel, decimal Premium, decimal PaidClaims);
    private sealed record CanonicalComparisonPeriod(string Role, string PeriodLabel, DateOnly Start, int SourceMonthCount);

    private sealed record PremiumClaimsSummarySlidesData(
        OverallPremiumClaimsSummarySlideData Overall,
        PremiumClaimsSummarySlideData Health,
        PremiumClaimsSummarySlideData Dental,
        PremiumClaimsSummarySlideData HealthComparison,
        PremiumClaimsSummarySlideData DentalComparison,
        IReadOnlyList<LossRatioBreakEvenChartRow> HealthLossRatioBreakEven,
        IReadOnlyList<LossRatioBreakEvenChartRow> DentalLossRatioBreakEven);

    private sealed record OverallPremiumClaimsSummaryRow(
        string Category,
        decimal PaidPremiums,
        decimal PaidClaims,
        decimal LossRatio);

    private sealed record OverallPremiumClaimsSummarySlideData(
        string CurrentPeriodLabel,
        IReadOnlyList<OverallPremiumClaimsSummaryRow> Rows);

    private sealed record PremiumClaimsSummarySlideData(
        string SlideCoverageTitle,
        string CoverageLabel,
        string CurrentPeriodLabel,
        string PriorPeriodLabel,
        string PremiumHeader,
        string PremiumSeriesLabel,
        decimal CurrentPremium,
        decimal CurrentPaidClaims,
        decimal CurrentLossRatio,
        decimal TrendFactor,
        decimal TotalExpenseFactor,
        decimal TargetLossRatio,
        decimal AnticipatedAdjustment,
        IReadOnlyList<PremiumClaimsChartRow> ChartRows);

    private sealed record LossRatioBreakEvenChartRow(string PeriodLabel, decimal NetLossRatio, decimal BreakEven);

    private sealed record ScheduleOfBenefitsDeckData(
        string ClientName,
        IReadOnlyList<ScheduleOfBenefitsRow> Rows);

    private sealed record ScheduleOfBenefitsRow(
        int SlidePage,
        int PageTotal,
        int SectionOrder,
        string ClientName,
        string Carrier,
        string ClassName,
        string SectionName,
        string DetailText,
        string SourceFile,
        string Notes);

    private sealed record MonthlyExperienceRow(string PeriodId, DateOnly Month, string Coverage, decimal BilledPremium, decimal NetPremium, decimal PaidClaims, decimal ClaimsPlusFees, decimal Lives);
    private sealed record AverageSpendRow(string Category, decimal CurrentAmount, decimal PriorAmount);
    private sealed record AverageSpendSlideData(
        string CurrentPeriodLabel,
        string PriorPeriodLabel,
        int MonthCount,
        decimal HealthEmployeeCount,
        decimal DentalEmployeeCount,
        IReadOnlyList<AverageSpendRow> Rows,
        DateOnly? HighestClaimsMonth,
        decimal HighestClaimsAmount)
    {
        public AverageSpendRow Health => Rows[0];
        public AverageSpendRow Dental => Rows[1];
        public AverageSpendRow Combined => Rows[2];
    }

    private sealed record CategorySpend(string Category, decimal PaidClaims);

    private sealed record ClaimsDistributionRow(string Category, decimal CurrentPaidClaims, decimal PriorPaidClaims);

    private sealed record ClaimsDistributionSlideData(
        string CurrentPeriodLabel,
        string PriorPeriodLabel,
        IReadOnlyList<ClaimsDistributionRow> Rows)
    {
        public decimal TotalCurrentPaidClaims => Rows.Sum(row => row.CurrentPaidClaims);
        public decimal TotalPriorPaidClaims => Rows.Sum(row => row.PriorPaidClaims);
    }

    private sealed record ParamedicalServiceMap(string Label, IReadOnlyList<string> SearchTerms);

    private sealed record ParamedicalBreakdownRow(string Service, decimal CurrentPaidClaims, decimal PriorPaidClaims);

    private sealed record ParamedicalBreakdownSlideData(
        string CurrentPeriodLabel,
        string PriorPeriodLabel,
        IReadOnlyList<ParamedicalBreakdownRow> Rows)
    {
        public decimal TotalCurrentPaidClaims => Rows.Sum(row => row.CurrentPaidClaims);
    }

    private sealed record DentalBreakdownRow(string Category, decimal CurrentPaidClaims, decimal PriorPaidClaims);

    private sealed record DentalBreakdownSlideData(
        string CurrentPeriodLabel,
        string PriorPeriodLabel,
        IReadOnlyList<DentalBreakdownRow> Rows)
    {
        public decimal TotalCurrentPaidClaims => Rows.Sum(row => row.CurrentPaidClaims);
    }

    private sealed record MonthlyBreakdownsSlideData(
        MonthlyBreakdownSlideData Health,
        MonthlyBreakdownSlideData Dental);

    private sealed record MonthlyBreakdownRow(
        DateOnly Month,
        decimal Premium,
        decimal PaidClaims,
        decimal LossRatio,
        decimal EmployeeCount);

    private sealed record MonthlyBreakdownSlideData(
        string SlideTitle,
        string Coverage,
        string EmployeeCoverageLabel,
        string PremiumHeader,
        IReadOnlyList<MonthlyBreakdownRow> Rows,
        decimal TotalPremium,
        decimal TotalPaidClaims,
        decimal TotalLossRatio,
        decimal AverageEmployeeCount,
        string TopCategory,
        decimal TopCategoryPaidClaims,
        DateOnly? HighestClaimsMonth,
        decimal HighestClaimsAmount,
        decimal AverageUtilization,
        decimal PriorPaidClaims);
}

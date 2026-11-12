using System.Globalization;
using System.Text.RegularExpressions;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public static class ClaimsDistributionExtractor
{
    private const decimal ReconciliationTolerance = 1m;
    private const string UnclassifiedParamedicalLabel = "Unclassified Practitioner - review source";

    private static readonly string[] TitleClues =
    [
        "Claims Distribution by Service",
        "Claims Distribution by Service Summary",
        "Medical Claims Distribution",
        "Service Summary",
        "Health Service Utilization",
        "Benefit Utilization Summary",
        "Claims Breakdown by Category",
        "Claims by Service Category",
        "Claims Paid by Type of Service"
    ];

    private static readonly Regex MetadataRegex = new(@"^(?<label>Policy\(s\)|Policy|Period|Benefit|Division\(s\)|Division|Class\(s\)|Class)\s+(?<value>.+)$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex ReportCreatedDateRegex = new(@"^Date\s+created:\s*(?<value>.+)$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex RowRegex = new(@"^(?<service>[A-Za-z][A-Za-z ',/&().+-]+?)\s+\$?(?<amount>-?[\d,]+(?:\.\d+)?)\s+(?<paidPercent>-?[\d.]+)%\s+(?<normPaidPercent>-?[\d.]+)%\s+(?<occurrences>\d+)\s+(?<occurrencePercent>-?[\d.]+)%\s+(?<normOccurrencePercent>-?[\d.]+)%\s+\$?(?<average>-?[\d,]+(?:\.\d+)?)$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex SunLifeServiceTitleRegex = new(@"^(?<benefit>Extended Health Care|Dental)\s+Claims\s+Paid\s+by\s+Type\s+of\s+Service$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex SunLifeContractNumberRegex = new(@"^Contract\s+Number:\s*(?<value>\d+)", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex SunLifeDateRangeRegex = new(@"\b(?<currentStart>[A-Za-z]+\s+\d{1,2},\s+\d{4})\s+to\s+(?<currentEnd>[A-Za-z]+\s+\d{1,2},\s+\d{4})\s+(?<priorStart>[A-Za-z]+\s+\d{1,2},\s+\d{4})\s+to\s+(?<priorEnd>[A-Za-z]+\s+\d{1,2},\s+\d{4})\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex SunLifeServiceRowRegex = new(@"^(?<service>[A-Za-z][A-Za-z ',/&().+-]+?)\s+(?<currentSubmitted>-?[\d,]+(?:\.\d+)?)\s+(?<currentEligible>-?[\d,]+(?:\.\d+)?)\s+(?<currentPaid>-?[\d,]+(?:\.\d+)?)\s+(?<currentPercent>-?[\d.]+)%?\s+(?<priorSubmitted>-?[\d,]+(?:\.\d+)?)\s+(?<priorEligible>-?[\d,]+(?:\.\d+)?)\s+(?<priorPaid>-?[\d,]+(?:\.\d+)?)\s+(?<priorPercent>-?[\d.]+)%?$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex SunLifeSubtotalRowRegex = new(@"^(?<currentSubmitted>-?[\d,]+(?:\.\d+)?)\s+(?<currentEligible>-?[\d,]+(?:\.\d+)?)\s+(?<currentPaid>-?[\d,]+(?:\.\d+)?)\s+(?<currentPercent>-?[\d.]+)%?\s+(?<priorSubmitted>-?[\d,]+(?:\.\d+)?)\s+(?<priorEligible>-?[\d,]+(?:\.\d+)?)\s+(?<priorPaid>-?[\d,]+(?:\.\d+)?)\s+(?<priorPercent>-?[\d.]+)%?$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex SunLifeTotalClaimsRowRegex = new(@"^Canadian\s+Dollars\s+Total\s+Claims\s+(?<currentSubmitted>-?[\d,]+(?:\.\d+)?)\s+(?<currentEligible>-?[\d,]+(?:\.\d+)?)\s+(?<currentPaid>-?[\d,]+(?:\.\d+)?)\s+(?<currentPercent>-?[\d.]+)%?\s+(?<priorSubmitted>-?[\d,]+(?:\.\d+)?)\s+(?<priorEligible>-?[\d,]+(?:\.\d+)?)\s+(?<priorPaid>-?[\d,]+(?:\.\d+)?)\s+(?<priorPercent>-?[\d.]+)%?$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex SunLifeTrailingPercentSubtotalRowRegex = new(@"^(?<currentSubmitted>-?[\d,]+(?:\.\d+)?)\s+(?<currentEligible>-?[\d,]+(?:\.\d+)?)\s+(?<currentPaid>-?[\d,]+(?:\.\d+)?)\s+(?<priorSubmitted>-?[\d,]+(?:\.\d+)?)\s+(?<priorEligible>-?[\d,]+(?:\.\d+)?)\s+(?<priorPaid>-?[\d,]+(?:\.\d+)?)\s+(?<currentPercent>-?[\d.]+)%?\s+(?<priorPercent>-?[\d.]+)%?$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex SunLifeTrailingPercentTotalClaimsRowRegex = new(@"^Canadian\s+Dollars\s+Total\s+Claims\s+(?<currentSubmitted>-?[\d,]+(?:\.\d+)?)\s+(?<currentEligible>-?[\d,]+(?:\.\d+)?)\s+(?<currentPaid>-?[\d,]+(?:\.\d+)?)\s+(?<priorSubmitted>-?[\d,]+(?:\.\d+)?)\s+(?<priorEligible>-?[\d,]+(?:\.\d+)?)\s+(?<priorPaid>-?[\d,]+(?:\.\d+)?)\s+(?<currentPercent>-?[\d.]+)%?\s+(?<priorPercent>-?[\d.]+)%?$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public static List<ClaimsDistributionRow> Extract(
        string sourceFile,
        IReadOnlyList<PdfPageText> pages,
        string carrierName,
        string policyNumber,
        string clientName)
    {
        var rows = new List<ClaimsDistributionRow>();
        foreach (var page in pages)
        {
            if (IsSunLifeServiceClaimsPage(page))
            {
                rows.AddRange(ExtractSunLifeServiceClaimsRows(page, sourceFile, carrierName, policyNumber, clientName));
                continue;
            }

            if (!IsClaimsDistributionPage(page))
            {
                continue;
            }

            var metadata = ReadMetadata(page, sourceFile, carrierName, policyNumber, clientName);
            foreach (var line in page.Lines.Select(ParsingHelpers.Clean).Where(line => line.Length > 0))
            {
                var match = RowRegex.Match(line);
                if (!match.Success)
                {
                    continue;
                }

                var originalService = ParsingHelpers.Clean(match.Groups["service"].Value);
                rows.Add(new ClaimsDistributionRow
                {
                    CarrierName = metadata.CarrierName,
                    PolicyNumber = metadata.PolicyNumber,
                    ClientName = metadata.ClientName,
                    ReportPeriod = metadata.ReportPeriod,
                    PeriodStart = metadata.PeriodStart,
                    PeriodEnd = metadata.PeriodEnd,
                    Benefit = metadata.Benefit,
                    Division = metadata.Division,
                    Class = metadata.Class,
                    ReportCreatedDate = metadata.ReportCreatedDate,
                    SourceFile = Path.GetFileName(sourceFile),
                    ServiceCategory = NormalizeServiceCategory(originalService),
                    ServiceDescriptionOriginal = originalService,
                    AmountPaid = ParseDecimal(match.Groups["amount"].Value),
                    PercentPaid = ParseDecimal(match.Groups["paidPercent"].Value),
                    NormPercentPaid = ParseDecimal(match.Groups["normPaidPercent"].Value),
                    Occurrences = ParseInt(match.Groups["occurrences"].Value),
                    PercentOccurrences = ParseDecimal(match.Groups["occurrencePercent"].Value),
                    NormPercentOccurrences = ParseDecimal(match.Groups["normOccurrencePercent"].Value),
                    AverageClaim = ParseDecimal(match.Groups["average"].Value),
                    SourceReference = $"{metadata.CarrierName} PDF page {page.PageNumber}".Trim()
                });
            }
        }

        return rows;
    }

    public static void MergeIntoWorkbook(
        WorkbookExtractionResult workbook,
        IReadOnlyList<ClaimsDistributionRow> extractedRows,
        bool updateServiceTrends = true)
    {
        if (extractedRows.Count == 0)
        {
            return;
        }

        foreach (var row in extractedRows)
        {
            var addedDistributionRow = false;
            if (!workbook.ClaimsDistributionRows.Any(current => SameDistributionRow(current, row)))
            {
                workbook.ClaimsDistributionRows.Add(row);
                addedDistributionRow = true;
            }

            if (!addedDistributionRow ||
                row.ServiceCategory.Equals("Total", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var isParamedicalDetail = IsParamedicalDetailRow(row);
            if (!isParamedicalDetail)
            {
                var claimBenefit = LegacyClaimBenefit(row.Benefit);
                if (!workbook.ClaimTrends.Any(current =>
                        NormalizeKey(current.Benefit).Equals(NormalizeKey(claimBenefit), StringComparison.OrdinalIgnoreCase) &&
                        NormalizeKey(current.Category).Equals(NormalizeKey(row.ServiceCategory), StringComparison.OrdinalIgnoreCase) &&
                        NormalizeKey(current.Period).Equals(NormalizeKey(row.ReportPeriod), StringComparison.OrdinalIgnoreCase)))
                {
                    workbook.ClaimTrends.Add(new ClaimTrend
                    {
                        Benefit = claimBenefit,
                        Category = row.ServiceCategory,
                        Period = row.ReportPeriod,
                        Amount = row.AmountPaid,
                        PercentOfTotal = row.PercentPaid,
                        PriorPeriodPercent = row.PriorPercentPaid,
                        BenchmarkPercent = row.NormPercentPaid,
                        SourceReference = row.SourceReference
                    });
                }
            }

            if (updateServiceTrends)
            {
                AddServiceTrend(workbook, row);
            }
        }

        workbook.ParamedicalServiceTrends.RemoveAll(row =>
            row.Category.Equals(UnclassifiedParamedicalLabel, StringComparison.OrdinalIgnoreCase));
        var recoveredParamedicalRows = AddParamedicalReconciliationRows(workbook.ClaimsDistributionRows);
        foreach (var recoveredRow in recoveredParamedicalRows)
        {
            AddParamedicalDetailTrend(workbook, recoveredRow);
        }

        workbook.ClaimsDistributionValidation = BuildValidation(workbook.ClaimsDistributionRows);
        workbook.Warnings.RemoveAll(IsClaimsDistributionValidationWarning);
        foreach (var warning in workbook.ClaimsDistributionValidation.Warnings)
        {
            if (!workbook.Warnings.Contains(warning, StringComparer.OrdinalIgnoreCase))
            {
                workbook.Warnings.Add(warning);
            }
        }
    }

    private static bool IsClaimsDistributionValidationWarning(string warning)
    {
        return warning.StartsWith("Claims Distribution", StringComparison.OrdinalIgnoreCase) ||
               warning.StartsWith("Paramedical row", StringComparison.OrdinalIgnoreCase) ||
               warning.StartsWith("Paramedical subtotal", StringComparison.OrdinalIgnoreCase) ||
               warning.StartsWith("Paramedical practitioner", StringComparison.OrdinalIgnoreCase) ||
               warning.StartsWith("Prior paramedical practitioner", StringComparison.OrdinalIgnoreCase) ||
               warning.StartsWith("An unexplained paramedical amount", StringComparison.OrdinalIgnoreCase) ||
               warning.StartsWith("Total row was not found in the Claims Distribution", StringComparison.OrdinalIgnoreCase);
    }

    public static ClaimsDistributionValidation BuildValidation(IReadOnlyList<ClaimsDistributionRow> rows)
    {
        var validation = new ClaimsDistributionValidation
        {
            TableFound = rows.Count > 0,
            ParamedicalFound = rows.Any(row => row.ServiceCategory.Equals("Paramedical", StringComparison.OrdinalIgnoreCase)),
            ParamedicalDetailFound = rows.Any(IsParamedicalPractitionerDetailRow),
            UnclassifiedParamedicalDetailFound = rows.Any(IsUnclassifiedParamedicalRow),
            TotalRowFound = rows.Any(row => row.ServiceCategory.Equals("Total", StringComparison.OrdinalIgnoreCase))
        };

        var latestGroup = rows
            .GroupBy(row => new
            {
                row.SourceFile,
                row.ReportPeriod,
                row.Benefit,
                row.Division,
                row.Class
            })
            .Select(group => new
            {
                Total = group.FirstOrDefault(row => row.ServiceCategory.Equals("Total", StringComparison.OrdinalIgnoreCase)),
                ServiceSum = group
                    .Where(row => !row.ServiceCategory.Equals("Total", StringComparison.OrdinalIgnoreCase))
                    .Sum(row => row.AmountPaid ?? 0m)
            })
            .FirstOrDefault(group => group.Total is not null);

        if (latestGroup is not null)
        {
            validation.TotalAmountPaid = latestGroup.Total!.AmountPaid;
            validation.SumOfServiceAmounts = latestGroup.ServiceSum;
            validation.AmountsBalanceToTotal = validation.TotalAmountPaid.HasValue &&
                                               Math.Abs(latestGroup.ServiceSum - validation.TotalAmountPaid.Value) <= 1m;
        }

        if (!validation.TableFound)
        {
            validation.Warnings.Add("Claims Distribution table not found.");
        }

        if (validation.TableFound && !validation.ParamedicalFound)
        {
            validation.Warnings.Add("Paramedical row was not found in the Claims Distribution table.");
        }

        if (validation.TableFound && !validation.TotalRowFound)
        {
            validation.Warnings.Add("Total row was not found in the Claims Distribution table.");
        }

        if (validation.TotalRowFound && !validation.AmountsBalanceToTotal)
        {
            validation.Warnings.Add("Claims Distribution service amounts do not balance to the total within $1.");
        }

        AddParamedicalValidation(rows, validation);

        return validation;
    }

    private static List<ClaimsDistributionRow> AddParamedicalReconciliationRows(List<ClaimsDistributionRow> rows)
    {
        rows.RemoveAll(IsUnclassifiedParamedicalRow);
        var recoveredRows = new List<ClaimsDistributionRow>();

        foreach (var group in BuildParamedicalGroups(rows))
        {
            if (group.Rollup is null || group.Details.Count == 0)
            {
                continue;
            }

            var currentVariance = PositiveUnexplainedVariance(
                group.Rollup.AmountPaid,
                group.Details.Sum(row => row.AmountPaid ?? 0m));
            var priorVariance = PositiveUnexplainedVariance(
                group.Rollup.PriorAmountPaid,
                group.Details.Sum(row => row.PriorAmountPaid ?? 0m));
            if (!currentVariance.HasValue && !priorVariance.HasValue)
            {
                continue;
            }

            var recovered = CreateUnclassifiedParamedicalRow(group.Rollup, currentVariance, priorVariance);
            rows.Add(recovered);
            recoveredRows.Add(recovered);
        }

        return recoveredRows;
    }

    private static void AddParamedicalValidation(
        IReadOnlyList<ClaimsDistributionRow> rows,
        ClaimsDistributionValidation validation)
    {
        var groups = BuildParamedicalGroups(rows)
            .Where(group => group.Rollup is not null)
            .OrderByDescending(group => PeriodSortValue(group.Rollup!.PeriodEnd, group.Rollup.ReportPeriod))
            .ThenByDescending(group => group.Details.Count)
            .ToList();
        if (groups.Count == 0)
        {
            return;
        }

        foreach (var group in groups)
        {
            var rollup = group.Rollup!;
            if (group.Details.Count == 0)
            {
                validation.Warnings.Add(
                    $"Paramedical subtotal {FormatMoney(rollup.AmountPaid)} was found for {ParamedicalPeriodLabel(rollup)}, but no practitioner detail rows were extracted.");
                continue;
            }

            var currentDetailSum = group.Details.Sum(row => row.AmountPaid ?? 0m);
            var currentVariance = rollup.AmountPaid.HasValue
                ? rollup.AmountPaid.Value - currentDetailSum
                : (decimal?)null;
            var priorDetailSum = group.Details.Sum(row => row.PriorAmountPaid ?? 0m);
            var priorVariance = rollup.PriorAmountPaid.HasValue
                ? rollup.PriorAmountPaid.Value - priorDetailSum
                : (decimal?)null;

            if (currentVariance.HasValue && Math.Abs(currentVariance.Value) > ReconciliationTolerance)
            {
                validation.Warnings.Add(
                    $"Paramedical practitioner detail for {ParamedicalPeriodLabel(rollup)} differs from the carrier subtotal by {currentVariance.Value:C2}.");
            }

            if (priorVariance.HasValue && Math.Abs(priorVariance.Value) > ReconciliationTolerance)
            {
                validation.Warnings.Add(
                    $"Prior paramedical practitioner detail for {ParamedicalPriorPeriodLabel(rollup)} differs from the carrier subtotal by {priorVariance.Value:C2}.");
            }

            var unclassified = group.Details.FirstOrDefault(IsUnclassifiedParamedicalRow);
            if (unclassified is not null)
            {
                validation.Warnings.Add(
                    $"An unexplained paramedical amount ({FormatMoney(unclassified.AmountPaid)} current; {FormatMoney(unclassified.PriorAmountPaid)} prior) was retained as '{UnclassifiedParamedicalLabel}'. Review the source report for a missing practitioner label.");
            }
        }

        var primary = groups.FirstOrDefault(group => group.Details.Count > 0) ?? groups[0];
        var primaryRollup = primary.Rollup!;
        validation.ParamedicalSubtotalAmountPaid = primaryRollup.AmountPaid;
        validation.SumOfParamedicalDetailAmounts = primary.Details.Count == 0
            ? 0m
            : primary.Details.Sum(row => row.AmountPaid ?? 0m);
        validation.ParamedicalDetailVariance = primaryRollup.AmountPaid.HasValue
            ? primaryRollup.AmountPaid.Value - validation.SumOfParamedicalDetailAmounts.Value
            : null;
        validation.ParamedicalDetailsBalanceToTotal = primaryRollup.AmountPaid.HasValue
            ? primary.Details.Count > 0 && Math.Abs(validation.ParamedicalDetailVariance!.Value) <= ReconciliationTolerance
            : null;

        validation.PriorParamedicalSubtotalAmountPaid = primaryRollup.PriorAmountPaid;
        validation.SumOfPriorParamedicalDetailAmounts = primary.Details.Count == 0
            ? 0m
            : primary.Details.Sum(row => row.PriorAmountPaid ?? 0m);
        validation.PriorParamedicalDetailVariance = primaryRollup.PriorAmountPaid.HasValue
            ? primaryRollup.PriorAmountPaid.Value - validation.SumOfPriorParamedicalDetailAmounts.Value
            : null;
        validation.PriorParamedicalDetailsBalanceToTotal = primaryRollup.PriorAmountPaid.HasValue
            ? primary.Details.Count > 0 && Math.Abs(validation.PriorParamedicalDetailVariance!.Value) <= ReconciliationTolerance
            : null;
    }

    private static List<ParamedicalReconciliationGroup> BuildParamedicalGroups(IReadOnlyList<ClaimsDistributionRow> rows)
    {
        return rows
            .GroupBy(ParamedicalGroupKey, StringComparer.OrdinalIgnoreCase)
            .Select(group => new ParamedicalReconciliationGroup(
                group.FirstOrDefault(IsParamedicalRollupRow),
                group.Where(IsParamedicalPractitionerDetailRow).ToList()))
            .Where(group => group.Rollup is not null || group.Details.Count > 0)
            .ToList();
    }

    private static ClaimsDistributionRow CreateUnclassifiedParamedicalRow(
        ClaimsDistributionRow rollup,
        decimal? currentVariance,
        decimal? priorVariance)
    {
        return new ClaimsDistributionRow
        {
            CarrierName = rollup.CarrierName,
            PolicyNumber = rollup.PolicyNumber,
            ClientName = rollup.ClientName,
            ReportPeriod = rollup.ReportPeriod,
            PeriodStart = rollup.PeriodStart,
            PeriodEnd = rollup.PeriodEnd,
            PriorReportPeriod = rollup.PriorReportPeriod,
            PriorPeriodStart = rollup.PriorPeriodStart,
            PriorPeriodEnd = rollup.PriorPeriodEnd,
            Benefit = "Paramedical Detail",
            Division = rollup.Division,
            Class = rollup.Class,
            ReportCreatedDate = rollup.ReportCreatedDate,
            SourceFile = rollup.SourceFile,
            ServiceCategory = "Paramedical",
            ServiceDescriptionOriginal = UnclassifiedParamedicalLabel,
            AmountPaid = currentVariance,
            PercentPaid = PercentageOf(currentVariance, rollup.AmountPaid),
            PriorAmountPaid = priorVariance,
            PriorPercentPaid = PercentageOf(priorVariance, rollup.PriorAmountPaid),
            SourceReference = MergeSourceReference(rollup.SourceReference, "Paramedical subtotal reconciliation")
        };
    }

    private static decimal? PositiveUnexplainedVariance(decimal? subtotal, decimal detailSum)
    {
        if (!subtotal.HasValue)
        {
            return null;
        }

        var variance = subtotal.Value - detailSum;
        return variance > ReconciliationTolerance ? variance : null;
    }

    private static decimal? PercentageOf(decimal? amount, decimal? total)
    {
        return amount.HasValue && total is > 0m
            ? amount.Value / total.Value * 100m
            : null;
    }

    private static bool IsClaimsDistributionPage(PdfPageText page)
    {
        var text = page.Text;
        var hasTitle = TitleClues.Any(clue => text.Contains(clue, StringComparison.OrdinalIgnoreCase));
        var hasTableShape = text.Contains("Amount Paid", StringComparison.OrdinalIgnoreCase) &&
                            text.Contains("Number of Occurrences", StringComparison.OrdinalIgnoreCase) &&
                            text.Contains("Average Amount Paid", StringComparison.OrdinalIgnoreCase);
        return hasTitle || hasTableShape;
    }

    private static ClaimsDistributionMetadata ReadMetadata(
        PdfPageText page,
        string sourceFile,
        string carrierName,
        string policyNumber,
        string clientName)
    {
        var metadata = new ClaimsDistributionMetadata
        {
            CarrierName = string.IsNullOrWhiteSpace(carrierName) ? "Unknown" : carrierName,
            PolicyNumber = policyNumber,
            ClientName = clientName,
            SourceFile = Path.GetFileName(sourceFile),
            Benefit = "Medical",
            Division = "ALL",
            Class = "ALL"
        };

        foreach (var line in page.Lines.Select(ParsingHelpers.Clean).Where(line => line.Length > 0))
        {
            var reportDate = ReportCreatedDateRegex.Match(line);
            if (reportDate.Success)
            {
                metadata.ReportCreatedDate = ParsingHelpers.Clean(reportDate.Groups["value"].Value);
                continue;
            }

            var match = MetadataRegex.Match(line);
            if (!match.Success)
            {
                continue;
            }

            var label = NormalizeKey(match.Groups["label"].Value);
            var value = ParsingHelpers.Clean(match.Groups["value"].Value);
            switch (label)
            {
                case "policys":
                case "policy":
                    metadata.PolicyNumber = string.IsNullOrWhiteSpace(metadata.PolicyNumber) ? value : metadata.PolicyNumber;
                    break;
                case "period":
                    metadata.ReportPeriod = value;
                    (metadata.PeriodStart, metadata.PeriodEnd) = SplitPeriod(value);
                    break;
                case "benefit":
                    metadata.Benefit = NormalizeBenefit(value);
                    break;
                case "divisions":
                case "division":
                    metadata.Division = value;
                    break;
                case "classs":
                case "class":
                    metadata.Class = value;
                    break;
            }
        }

        return metadata;
    }

    private static bool IsSunLifeServiceClaimsPage(PdfPageText page)
    {
        var hasPeriodShape = page.Text.Contains("Current Period", StringComparison.OrdinalIgnoreCase) &&
                             page.Text.Contains("Prior Period", StringComparison.OrdinalIgnoreCase);
        var hasContinuationTotal = page.Text.Contains("Canadian Dollars Total Claims", StringComparison.OrdinalIgnoreCase);
        return page.Text.Contains("Sun Life", StringComparison.OrdinalIgnoreCase) &&
               page.Text.Contains("Claims Paid by Type of Service", StringComparison.OrdinalIgnoreCase) &&
               (hasPeriodShape || hasContinuationTotal);
    }

    private static List<ClaimsDistributionRow> ExtractSunLifeServiceClaimsRows(
        PdfPageText page,
        string sourceFile,
        string carrierName,
        string policyNumber,
        string clientName)
    {
        var rows = new List<ClaimsDistributionRow>();
        if (page.Lines.Select(ParsingHelpers.Clean).Any(line => line.Equals("Details", StringComparison.OrdinalIgnoreCase)))
        {
            return rows;
        }

        var metadata = ReadSunLifeServiceClaimsMetadata(page, sourceFile, carrierName, policyNumber, clientName);
        var currentSection = "";
        var currentSectionOriginal = "";

        foreach (var rawLine in page.Lines)
        {
            var line = ParsingHelpers.Clean(rawLine);
            if (line.Length == 0)
            {
                continue;
            }

            if (TryMapSunLifeServiceSection(line, metadata.Benefit, out var sectionCategory, out var sectionOriginal))
            {
                currentSection = sectionCategory;
                currentSectionOriginal = sectionOriginal;
                continue;
            }

            var totalMatch = SunLifeTotalClaimsRowRegex.Match(line);
            if (!totalMatch.Success)
            {
                totalMatch = SunLifeTrailingPercentTotalClaimsRowRegex.Match(line);
            }
            if (totalMatch.Success)
            {
                rows.Add(CreateSunLifeClaimsRow(
                    metadata,
                    sourceFile,
                    page.PageNumber,
                    "Total",
                    "Total",
                    totalMatch));
                continue;
            }

            var subtotalMatch = SunLifeSubtotalRowRegex.Match(line);
            if (!subtotalMatch.Success)
            {
                subtotalMatch = SunLifeTrailingPercentSubtotalRowRegex.Match(line);
            }
            if (subtotalMatch.Success && !string.IsNullOrWhiteSpace(currentSection))
            {
                rows.Add(CreateSunLifeClaimsRow(
                    metadata,
                    sourceFile,
                    page.PageNumber,
                    currentSection,
                    currentSectionOriginal,
                    subtotalMatch));
                continue;
            }

            var detailMatch = SunLifeServiceRowRegex.Match(line);
            // Compact summaries put all eight amounts on the category row itself.
            // In detailed reports the section subtotal remains authoritative.
            if (detailMatch.Success && string.IsNullOrWhiteSpace(currentSection) &&
                TryMapSunLifeServiceSection(
                    ParsingHelpers.Clean(detailMatch.Groups["service"].Value),
                    metadata.Benefit,
                    out var summaryCategory,
                    out var summaryDescription))
            {
                rows.Add(CreateSunLifeClaimsRow(
                    metadata, sourceFile, page.PageNumber,
                    summaryCategory, summaryDescription, detailMatch));
                continue;
            }

            if (!detailMatch.Success ||
                !currentSection.Equals("Paramedical", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var service = ParsingHelpers.Clean(detailMatch.Groups["service"].Value);
            rows.Add(CreateSunLifeClaimsRow(
                metadata,
                sourceFile,
                page.PageNumber,
                "Paramedical",
                service,
                detailMatch,
                benefitOverride: "Paramedical Detail"));
        }

        return rows;
    }

    private static ClaimsDistributionMetadata ReadSunLifeServiceClaimsMetadata(
        PdfPageText page,
        string sourceFile,
        string carrierName,
        string policyNumber,
        string clientName)
    {
        var metadata = new ClaimsDistributionMetadata
        {
            CarrierName = string.IsNullOrWhiteSpace(carrierName) ? "Sun Life" : carrierName,
            PolicyNumber = policyNumber,
            ClientName = clientName,
            SourceFile = Path.GetFileName(sourceFile),
            Benefit = "Medical",
            Division = "ALL",
            Class = "ALL"
        };

        foreach (var rawLine in page.Lines)
        {
            var line = ParsingHelpers.Clean(rawLine);
            if (line.Length == 0)
            {
                continue;
            }

            var titleMatch = SunLifeServiceTitleRegex.Match(line);
            if (titleMatch.Success)
            {
                metadata.Benefit = titleMatch.Groups["benefit"].Value.Equals("Dental", StringComparison.OrdinalIgnoreCase)
                    ? "Dental"
                    : "Medical";
                continue;
            }

            var policyMatch = SunLifeContractNumberRegex.Match(line);
            if (policyMatch.Success && string.IsNullOrWhiteSpace(metadata.PolicyNumber))
            {
                metadata.PolicyNumber = policyMatch.Groups["value"].Value;
                continue;
            }

            var periodMatch = SunLifeDateRangeRegex.Match(line);
            if (periodMatch.Success)
            {
                metadata.PeriodStart = periodMatch.Groups["currentStart"].Value;
                metadata.PeriodEnd = periodMatch.Groups["currentEnd"].Value;
                metadata.ReportPeriod = $"{metadata.PeriodStart} to {metadata.PeriodEnd}";
                metadata.PriorPeriodStart = periodMatch.Groups["priorStart"].Value;
                metadata.PriorPeriodEnd = periodMatch.Groups["priorEnd"].Value;
                metadata.PriorReportPeriod = $"{metadata.PriorPeriodStart} to {metadata.PriorPeriodEnd}";
                continue;
            }

            if (line.StartsWith("Relation :", StringComparison.OrdinalIgnoreCase))
            {
                metadata.Division = FirstNonBlank(line.Split(':').LastOrDefault()?.Trim() ?? "", metadata.Division);
                metadata.Class = metadata.Division;
                continue;
            }

            if (string.IsNullOrWhiteSpace(metadata.ClientName) &&
                !line.Contains("Claims Paid by Type of Service", StringComparison.OrdinalIgnoreCase) &&
                !line.StartsWith("Contract Number:", StringComparison.OrdinalIgnoreCase) &&
                !line.Equals("Summary", StringComparison.OrdinalIgnoreCase) &&
                !line.Equals("Current Period Prior Period", StringComparison.OrdinalIgnoreCase) &&
                !line.StartsWith("Amount ", StringComparison.OrdinalIgnoreCase) &&
                !line.StartsWith("May ", StringComparison.OrdinalIgnoreCase) &&
                !line.StartsWith("Relation :", StringComparison.OrdinalIgnoreCase) &&
                !line.StartsWith("Sun Life", StringComparison.OrdinalIgnoreCase))
            {
                metadata.ClientName = line;
            }
        }

        return metadata;
    }

    private static ClaimsDistributionRow CreateSunLifeClaimsRow(
        ClaimsDistributionMetadata metadata,
        string sourceFile,
        int pageNumber,
        string serviceCategory,
        string serviceDescriptionOriginal,
        Match match,
        string? benefitOverride = null)
    {
        return new ClaimsDistributionRow
        {
            CarrierName = metadata.CarrierName,
            PolicyNumber = metadata.PolicyNumber,
            ClientName = metadata.ClientName,
            ReportPeriod = metadata.ReportPeriod,
            PeriodStart = metadata.PeriodStart,
            PeriodEnd = metadata.PeriodEnd,
            PriorReportPeriod = metadata.PriorReportPeriod,
            PriorPeriodStart = metadata.PriorPeriodStart,
            PriorPeriodEnd = metadata.PriorPeriodEnd,
            Benefit = benefitOverride ?? metadata.Benefit,
            Division = metadata.Division,
            Class = metadata.Class,
            ReportCreatedDate = metadata.ReportCreatedDate,
            SourceFile = Path.GetFileName(sourceFile),
            ServiceCategory = serviceCategory,
            ServiceDescriptionOriginal = serviceDescriptionOriginal,
            AmountPaid = ParseDecimal(match.Groups["currentPaid"].Value),
            PercentPaid = ParseDecimal(match.Groups["currentPercent"].Value),
            PriorAmountPaid = ParseDecimal(match.Groups["priorPaid"].Value),
            PriorPercentPaid = ParseDecimal(match.Groups["priorPercent"].Value),
            SourceReference = $"{metadata.CarrierName} PDF page {pageNumber}".Trim()
        };
    }

    private static bool TryMapSunLifeServiceSection(
        string line,
        string benefit,
        out string serviceCategory,
        out string serviceDescriptionOriginal)
    {
        serviceCategory = "";
        serviceDescriptionOriginal = "";
        var normalized = NormalizeKey(line);
        if (benefit.Equals("Dental", StringComparison.OrdinalIgnoreCase))
        {
            serviceCategory = normalized switch
            {
                "diagnosticpreventive" => "Preventive & Diagnostic",
                "majorrestorative" => "Major Restorative",
                "periodonticsendodontics" => "Periodontal & Endodontic",
                "surgicalrestorative" => "Surgical & Restorative",
                "orthodontic" or "orthodontics" => "Orthodontic",
                _ => ""
            };
        }
        else
        {
            serviceCategory = normalized switch
            {
                "paydirectdrugs" or "reimbursementdrugs" => "Drugs",
                "medicalservicesequipment" => "Medical Supplies",
                "paramedicalpractitioners" => "Paramedical",
                "hospital" or "hospitalincanada" => "Hospital",
                "contactlensesoreyeglasses" => "Vision",
                _ => ""
            };
        }

        serviceDescriptionOriginal = string.IsNullOrWhiteSpace(serviceCategory) ? "" : line;
        return !string.IsNullOrWhiteSpace(serviceCategory);
    }

    private static void AddServiceTrend(WorkbookExtractionResult workbook, ClaimsDistributionRow row)
    {
        if (IsParamedicalDetailRow(row))
        {
            AddParamedicalDetailTrend(workbook, row);
            return;
        }

        var trend = new ServiceCategoryTrend
        {
            Benefit = LegacyClaimBenefit(row.Benefit),
            Category = row.ServiceCategory,
            PriorPeriodLabel = row.PriorReportPeriod,
            PriorAmount = row.PriorAmountPaid ?? 0m,
            CurrentPeriodLabel = row.ReportPeriod,
            CurrentAmount = row.AmountPaid ?? 0m,
            SourceReference = row.SourceReference
        };

        var target = row.Benefit.Equals("Dental", StringComparison.OrdinalIgnoreCase)
            ? workbook.DentalSpendingTrends
            : workbook.EhcSpendingTrends;

        var existing = target.FirstOrDefault(current =>
            NormalizeKey(current.Category).Equals(NormalizeKey(trend.Category), StringComparison.OrdinalIgnoreCase) &&
            NormalizeKey(current.CurrentPeriodLabel).Equals(NormalizeKey(trend.CurrentPeriodLabel), StringComparison.OrdinalIgnoreCase));
        if (existing is null)
        {
            target.Add(trend);
            AddParamedicalDetailTrend(workbook, row);
            return;
        }

        if (row.CarrierName.Equals("Sun Life", StringComparison.OrdinalIgnoreCase))
        {
            existing.PriorPeriodLabel = FirstNonBlank(existing.PriorPeriodLabel, trend.PriorPeriodLabel);
            existing.PriorAmount += trend.PriorAmount;
            existing.CurrentPeriodLabel = FirstNonBlank(existing.CurrentPeriodLabel, trend.CurrentPeriodLabel);
            existing.CurrentAmount += trend.CurrentAmount;
            existing.SourceReference = MergeSourceReference(existing.SourceReference, trend.SourceReference);
        }

        AddParamedicalDetailTrend(workbook, row);
    }

    private static void AddParamedicalDetailTrend(WorkbookExtractionResult workbook, ClaimsDistributionRow row)
    {
        if (!IsParamedicalDetailRow(row) &&
            !row.ServiceCategory.Equals("Paramedical", StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        var category = ParamedicalCategoryNormalizer.Normalize(row.ServiceDescriptionOriginal);
        if (ParamedicalCategoryNormalizer.IsRollup(category))
        {
            return;
        }

        var trend = new ServiceCategoryTrend
        {
            Benefit = LegacyClaimBenefit(row.Benefit),
            Category = category,
            PriorPeriodLabel = row.PriorReportPeriod,
            PriorAmount = row.PriorAmountPaid ?? 0m,
            CurrentPeriodLabel = row.ReportPeriod,
            CurrentAmount = row.AmountPaid ?? 0m,
            SourceReference = row.SourceReference
        };

        var existing = workbook.ParamedicalServiceTrends.FirstOrDefault(current =>
            NormalizeKey(current.Category).Equals(NormalizeKey(trend.Category), StringComparison.OrdinalIgnoreCase) &&
            NormalizeKey(current.CurrentPeriodLabel).Equals(NormalizeKey(trend.CurrentPeriodLabel), StringComparison.OrdinalIgnoreCase));
        if (existing is null)
        {
            workbook.ParamedicalServiceTrends.Add(trend);
            return;
        }

        if (row.CarrierName.Equals("Sun Life", StringComparison.OrdinalIgnoreCase))
        {
            existing.PriorPeriodLabel = FirstNonBlank(existing.PriorPeriodLabel, trend.PriorPeriodLabel);
            existing.PriorAmount += trend.PriorAmount;
            existing.CurrentPeriodLabel = FirstNonBlank(existing.CurrentPeriodLabel, trend.CurrentPeriodLabel);
            existing.CurrentAmount += trend.CurrentAmount;
            existing.SourceReference = MergeSourceReference(existing.SourceReference, trend.SourceReference);
        }
    }

    public static string NormalizeServiceCategory(string value)
    {
        var cleaned = ParsingHelpers.Clean(value);
        var normalized = NormalizeKey(cleaned);
        if (normalized == "total")
        {
            return "Total";
        }

        if (normalized is "drug" or "drugs" or "prescriptiondrugs" or "prescriptionmedication" or "paydirectdrugs" or "reimbursementdrugs")
        {
            return "Drugs";
        }

        if (normalized is "paramedical" or "paramedicalservices" or "practitionerservices" or "healthpractitioners" or "professionalservices" or "alliedhealth" or "massagetherapy" or "physiotherapy" or "chiropractic" or "psychology" or "psychologistsofficevisit" or "naturopathic" or "acupuncture" ||
            ParamedicalCategoryNormalizer.IsKnownParamedicalCategory(cleaned))
        {
            return "Paramedical";
        }

        if (normalized is "serviceandsupplies" or "medicalsupplies" or "supplies" or "medicalservicesandsupplies" or "medicalservicesequipment")
        {
            return "Medical Supplies";
        }

        if (normalized is "outofcountry" or "travel" or "emergencytravel")
        {
            return "Out of Country";
        }

        if (normalized is "vision" or "visioncare" or "eyeglasses" or "eyeexams" or "contactlensesoreyeglasses")
        {
            return "Vision";
        }

        if (normalized is "diagnosticpreventive" or "preventivediagnostic")
        {
            return "Preventive & Diagnostic";
        }

        if (normalized is "periodonticsendodontics" or "periodontalendodontic")
        {
            return "Periodontal & Endodontic";
        }

        if (normalized is "surgicalrestorative")
        {
            return "Surgical & Restorative";
        }

        if (normalized is "dental" or "basicdental" or "majordental" or "orthodontics")
        {
            return "Dental";
        }

        if (normalized == "hospital")
        {
            return "Hospital";
        }

        if (normalized == "ambulance")
        {
            return "Ambulance";
        }

        if (normalized == "outofprovince")
        {
            return "Out of Province";
        }

        if (normalized is "other" or "miscellaneous")
        {
            return "Other";
        }

        return cleaned;
    }

    private static string NormalizeBenefit(string value)
    {
        var cleaned = ParsingHelpers.Clean(value);
        return cleaned switch
        {
            "Healthcare" or "Health" => "Medical",
            "Visioncare" => "Vision",
            _ => cleaned
        };
    }

    private static string LegacyClaimBenefit(string value)
    {
        return value.Equals("Medical", StringComparison.OrdinalIgnoreCase) ||
               value.Equals("Healthcare", StringComparison.OrdinalIgnoreCase) ||
               value.Equals("EHC", StringComparison.OrdinalIgnoreCase) ||
               value.Equals("Paramedical Detail", StringComparison.OrdinalIgnoreCase)
            ? "Health"
            : value;
    }

    private static bool IsParamedicalDetailRow(ClaimsDistributionRow row)
    {
        return row.Benefit.Equals("Paramedical Detail", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsParamedicalPractitionerDetailRow(ClaimsDistributionRow row)
    {
        return !IsParamedicalRollupRow(row) &&
               (IsParamedicalDetailRow(row) ||
                row.ServiceCategory.Equals("Paramedical", StringComparison.OrdinalIgnoreCase));
    }

    private static bool IsParamedicalRollupRow(ClaimsDistributionRow row)
    {
        if (!row.ServiceCategory.Equals("Paramedical", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        var descriptionKey = NormalizeKey(row.ServiceDescriptionOriginal);
        return ParamedicalCategoryNormalizer.IsRollup(row.ServiceDescriptionOriginal) ||
               descriptionKey is "paramedicalpractitioners" or "paramedicalpractitionerstotal" or
                   "paramedicalsubtotal" or "paramedicalsubtotals";
    }

    private static bool IsUnclassifiedParamedicalRow(ClaimsDistributionRow row)
    {
        return row.ServiceDescriptionOriginal.Equals(UnclassifiedParamedicalLabel, StringComparison.OrdinalIgnoreCase);
    }

    private static string ParamedicalGroupKey(ClaimsDistributionRow row)
    {
        return string.Join("|",
            NormalizeKey(row.SourceFile),
            NormalizeKey(row.CarrierName),
            NormalizeKey(row.ClientName),
            NormalizeKey(row.PolicyNumber),
            NormalizeKey(row.ReportPeriod),
            NormalizeKey(row.PriorReportPeriod),
            NormalizeKey(row.Division),
            NormalizeKey(row.Class));
    }

    private static DateOnly PeriodSortValue(string periodEnd, string reportPeriod)
    {
        if (DateOnly.TryParse(periodEnd, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out var end))
        {
            return end;
        }

        var candidates = Regex.Matches(reportPeriod, @"[A-Za-z]+\s+\d{1,2},\s+\d{4}|\d{1,2}[-/]\w{3}[-/]\d{4}")
            .Select(match => match.Value)
            .ToList();
        return candidates
            .Select(value => DateOnly.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out var parsed)
                ? parsed
                : DateOnly.MinValue)
            .DefaultIfEmpty(DateOnly.MinValue)
            .Max();
    }

    private static string ParamedicalPeriodLabel(ClaimsDistributionRow row)
    {
        return FirstNonBlank(row.ReportPeriod, row.PeriodEnd, row.SourceFile, "the current period");
    }

    private static string ParamedicalPriorPeriodLabel(ClaimsDistributionRow row)
    {
        return FirstNonBlank(row.PriorReportPeriod, row.PriorPeriodEnd, row.SourceFile, "the prior period");
    }

    private static string FormatMoney(decimal? value)
    {
        return value.HasValue ? value.Value.ToString("C2", CultureInfo.CurrentCulture) : "not available";
    }

    private static string FirstNonBlank(params string[] values)
    {
        return values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)) ?? "";
    }

    private static string MergeSourceReference(string left, string right)
    {
        return string.Join("; ", new[] { left, right }
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Distinct(StringComparer.OrdinalIgnoreCase));
    }

    private static bool SameDistributionRow(ClaimsDistributionRow left, ClaimsDistributionRow right)
    {
        return NormalizeKey(left.SourceFile).Equals(NormalizeKey(right.SourceFile), StringComparison.OrdinalIgnoreCase) &&
               NormalizeKey(left.Benefit).Equals(NormalizeKey(right.Benefit), StringComparison.OrdinalIgnoreCase) &&
               NormalizeKey(left.ServiceDescriptionOriginal).Equals(NormalizeKey(right.ServiceDescriptionOriginal), StringComparison.OrdinalIgnoreCase) &&
               NormalizeKey(left.ReportPeriod).Equals(NormalizeKey(right.ReportPeriod), StringComparison.OrdinalIgnoreCase) &&
               NormalizeKey(left.Division).Equals(NormalizeKey(right.Division), StringComparison.OrdinalIgnoreCase) &&
               NormalizeKey(left.Class).Equals(NormalizeKey(right.Class), StringComparison.OrdinalIgnoreCase);
    }

    private static (string Start, string End) SplitPeriod(string period)
    {
        var parts = period.Split(" - ", 2, StringSplitOptions.TrimEntries);
        return parts.Length == 2 ? (parts[0], parts[1]) : ("", "");
    }

    private static decimal? ParseDecimal(string value)
    {
        var cleaned = value
            .Replace("$", "", StringComparison.Ordinal)
            .Replace(",", "", StringComparison.Ordinal)
            .Replace("%", "", StringComparison.Ordinal)
            .Trim();
        return decimal.TryParse(cleaned, NumberStyles.Any, CultureInfo.InvariantCulture, out var parsed)
            ? parsed
            : null;
    }

    private static int? ParseInt(string value)
    {
        return int.TryParse(value.Replace(",", "", StringComparison.Ordinal).Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed)
            ? parsed
            : null;
    }

    private static string NormalizeKey(string value)
    {
        return Regex.Replace(ParsingHelpers.Clean(value), @"[^A-Za-z0-9]", "").ToLowerInvariant();
    }

    private sealed class ClaimsDistributionMetadata
    {
        public string CarrierName { get; set; } = "";
        public string PolicyNumber { get; set; } = "";
        public string ClientName { get; set; } = "";
        public string ReportPeriod { get; set; } = "";
        public string PeriodStart { get; set; } = "";
        public string PeriodEnd { get; set; } = "";
        public string PriorReportPeriod { get; set; } = "";
        public string PriorPeriodStart { get; set; } = "";
        public string PriorPeriodEnd { get; set; } = "";
        public string Benefit { get; set; } = "";
        public string Division { get; set; } = "";
        public string Class { get; set; } = "";
        public string ReportCreatedDate { get; set; } = "";
        public string SourceFile { get; set; } = "";
    }

    private sealed record ParamedicalReconciliationGroup(
        ClaimsDistributionRow? Rollup,
        List<ClaimsDistributionRow> Details);
}

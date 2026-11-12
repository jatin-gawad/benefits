using System.Globalization;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed class ManulifeWorkbookParser
{
    private static readonly string[] EhcSpendingCategories =
    [
        "Medical Services and Supplies",
        "Hospital",
        "Pooled",
        "Drugs",
        "Paramedical",
        "Vision",
        "Miscellaneous Services",
        "Total"
    ];

    private static readonly string[] DentalSpendingCategories =
    [
        "Basic",
        "Major",
        "Dentures",
        "Endodontics",
        "Orthodontic",
        "Periodontic",
        "Miscellaneous Services",
        "Total"
    ];

    private static readonly string[] ParamedicalServiceOrder =
    [
        "Acupuncture",
        "Chiropractor",
        "Registered Massage Therapist",
        "Physiotherapist",
        "Psychologist",
        "Clinical Counsellor"
    ];

    private static readonly string[] RequiredSheets =
    [
        "Dent-Hlth Ttl Claims",
        "QuartileAnalysisDentHlth",
        "ClaimsAnalysis_EHC",
        "ClaimsAnalysis_Dental",
        "TotalEHC",
        "Top_50_DIN",
        "Top50CommonIndications"
    ];

    public WorkbookExtractionResult Parse(
        string workbookPath,
        string? clientName = null,
        IEnumerable<string>? claimsWorkbookPaths = null,
        string? currentYearFolder = null,
        string? priorYearFolder = null,
        string? policyNumber = null)
    {
        var result = new WorkbookExtractionResult { SourcePath = workbookPath };

        using var workbook = OpenWorkbookSharedRead(workbookPath);
        foreach (var sheet in RequiredSheets)
        {
            if (!workbook.TryGetWorksheet(sheet, out _))
            {
                result.MissingSheets.Add(sheet);
            }
        }

        ReadClaimCategorySums(workbook, "ClaimsAnalysis_EHC", "Extended Health Care", result);
        ReadEhcSpendingTrends(workbook, workbookPath, clientName, policyNumber, claimsWorkbookPaths, currentYearFolder, priorYearFolder, result);
        ReadClaimCategorySums(workbook, "ClaimsAnalysis_Dental", "Dental", result);
        ReadDentalSpendingTrends(workbook, workbookPath, clientName, policyNumber, claimsWorkbookPaths, currentYearFolder, priorYearFolder, result);
        ReadParamedicalServiceTrends(workbookPath, clientName, policyNumber, claimsWorkbookPaths, currentYearFolder, priorYearFolder, result);
        LogClaimsTrendCaptureSummary(result);
        ReadBenchmarks(workbook, result);
        ReadTopDrugs(workbook, workbookPath, clientName, currentYearFolder, priorYearFolder, result);
        ReadPriorTopDrugs(workbook, workbookPath, clientName, claimsWorkbookPaths, currentYearFolder, priorYearFolder, result);
        ReadIndications(workbook, result);
        ReadQuartiles(workbook, result);
        ReadCertificateClaims(workbook, result);

        return result;
    }

    public int MergeSupportingClaimsWorkbookHistory(
        WorkbookExtractionResult result,
        string anchorPath,
        string? clientName,
        IEnumerable<string>? claimsWorkbookPaths,
        string? currentYearFolder,
        string? priorYearFolder,
        string? policyNumber = null)
    {
        if (string.IsNullOrWhiteSpace(anchorPath) || !File.Exists(anchorPath))
        {
            return 0;
        }

        var updated = 0;
        var legacyReader = new LegacyClaimsServiceWorkbookReader();
        foreach (var candidate in FindClaimsWorkbookCandidates(anchorPath, clientName, claimsWorkbookPaths, currentYearFolder, priorYearFolder)
                     .Where(path => Path.GetExtension(path).Equals(".xls", StringComparison.OrdinalIgnoreCase)))
        {
            var before = result.ClaimsDistributionRows.Count;
            var rows = legacyReader.Extract(candidate);
            ClaimsDistributionExtractor.MergeIntoWorkbook(result, rows);
            updated += result.ClaimsDistributionRows.Count - before;
        }

        var ehcSnapshots = ReadEhcSummarySnapshotsNearWorkbook(anchorPath, clientName, policyNumber, claimsWorkbookPaths, currentYearFolder, priorYearFolder);
        updated += MergeSupportingClaimsWorkbookHistory(result.EhcSpendingTrends, ehcSnapshots, "ClaimsAnalysis_EHC", result);

        var dentalSnapshots = ReadDentalSummarySnapshotsNearWorkbook(anchorPath, clientName, policyNumber, claimsWorkbookPaths, currentYearFolder, priorYearFolder);
        updated += MergeSupportingClaimsWorkbookHistory(result.DentalSpendingTrends, dentalSnapshots, "ClaimsAnalysis_Dental", result);

        var paramedicalSnapshots = ReadParamedicalSummarySnapshotsNearWorkbook(anchorPath, clientName, policyNumber, claimsWorkbookPaths, currentYearFolder, priorYearFolder);
        updated += MergeSupportingClaimsWorkbookHistory(result.ParamedicalServiceTrends, paramedicalSnapshots, "ClaimsAnalysis_Paramedical", result);
        updated += MergeParamedicalRollupFromEhcHistory(result.ParamedicalServiceTrends, ehcSnapshots, result);
        updated += MergeSupportingPriorTopDrugs(result, claimsWorkbookPaths, currentYearFolder, priorYearFolder);

        return updated;
    }

    private static int MergeSupportingPriorTopDrugs(
        WorkbookExtractionResult result,
        IEnumerable<string>? claimsWorkbookPaths,
        string? currentYearFolder,
        string? priorYearFolder)
    {
        if (result.DrugTrends.Count == 0 || result.PriorDrugTrends.Count > 0)
        {
            return 0;
        }

        var currentEndDate = result.DrugTrends
            .Select(drug => TryParsePeriodDatesFromLabel(drug.Period).EndDate)
            .Where(date => date.HasValue)
            .Select(date => date!.Value)
            .DefaultIfEmpty(DateOnly.MaxValue)
            .Max();
        if (currentEndDate == DateOnly.MaxValue)
        {
            return 0;
        }

        var candidatePaths = (claimsWorkbookPaths ?? Enumerable.Empty<string>())
            .Concat(EnumerateExcelFiles(currentYearFolder))
            .Concat(EnumerateExcelFiles(priorYearFolder))
            .Where(path => !string.IsNullOrWhiteSpace(path) && File.Exists(path))
            .Where(path =>
                path.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                path.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        var candidates = new List<TopDrugWorkbookCandidate>();
        foreach (var candidatePath in candidatePaths)
        {
            try
            {
                using var workbook = OpenWorkbookSharedRead(candidatePath);
                var drugs = ExtractTopDrugs(workbook, Path.GetFileName(candidatePath));
                if (drugs.Count == 0)
                {
                    continue;
                }

                drugs = drugs.Take(10).ToList();

                var period = ReadWorkbookPeriod(workbook, candidatePath);
                if (!period.EndDate.HasValue || period.EndDate.Value >= currentEndDate)
                {
                    continue;
                }

                foreach (var drug in drugs)
                {
                    drug.Period = period.Label;
                }

                candidates.Add(new TopDrugWorkbookCandidate(
                    candidatePath,
                    period.Label,
                    period.EndDate,
                    drugs,
                    GetCandidatePeriodRole(candidatePath, currentYearFolder, priorYearFolder)));
            }
            catch
            {
                // Ignore unrelated or locked supporting workbooks.
            }
        }

        var selected = candidates
            .Where(candidate => !TopDrugListsMatch(result.DrugTrends, candidate.Drugs))
            .OrderByDescending(candidate => candidate.EndDate)
            .FirstOrDefault();
        if (selected is null)
        {
            return 0;
        }

        result.PriorDrugTrends.AddRange(selected.Drugs);
        var periodText = string.IsNullOrWhiteSpace(selected.PeriodLabel)
            ? "period not identified"
            : selected.PeriodLabel;
        result.Warnings.Add($"Prior Top DIN source: {Path.GetFileName(selected.SourcePath)} ({periodText}).");
        return selected.Drugs.Count;
    }

    private static void ReadClaimCategorySums(XLWorkbook workbook, string sheetName, string benefit, WorkbookExtractionResult result)
    {
        if (!workbook.TryGetWorksheet(sheetName, out var sheet))
        {
            return;
        }

        var headerRow = FindRow(sheet, "Member/ Dependents") ?? 10;
        var headers = new Dictionary<int, string>();
        for (var col = 4; col <= 10; col++)
        {
            var header = Text(sheet.Cell(headerRow, col));
            if (!string.IsNullOrWhiteSpace(header) && !header.Contains("Miscellaneous", StringComparison.OrdinalIgnoreCase))
            {
                headers[col] = header;
            }
        }

        var sums = headers.ToDictionary(pair => pair.Key, _ => 0m);
        var total = 0m;
        var lastRow = sheet.LastRowUsed()?.RowNumber() ?? headerRow;

        for (var row = headerRow + 1; row <= lastRow; row++)
        {
            var memberType = Text(sheet.Cell(row, 3));
            if (!memberType.Equals("Total Certificate", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            foreach (var col in headers.Keys)
            {
                sums[col] += Number(sheet.Cell(row, col)) ?? 0m;
            }

            total += Number(sheet.Cell(row, 11)) ?? 0m;
        }

        foreach (var pair in sums.OrderByDescending(pair => pair.Value))
        {
            result.ClaimTrends.Add(new ClaimTrend
            {
                Benefit = benefit,
                Category = headers[pair.Key],
                Period = "Current period",
                Amount = pair.Value,
                PercentOfTotal = total > 0 ? decimal.Round(pair.Value / total * 100m, 1) : null,
                SourceReference = $"{sheetName}!{ColumnName(pair.Key)}{headerRow}:{ColumnName(pair.Key)}{lastRow}"
            });
        }
    }

    private static void ReadEhcSpendingTrends(
        XLWorkbook workbook,
        string workbookPath,
        string? clientName,
        string? policyNumber,
        IEnumerable<string>? claimsWorkbookPaths,
        string? currentYearFolder,
        string? priorYearFolder,
        WorkbookExtractionResult result)
    {
        var externalSnapshots = ReadEhcSummarySnapshotsNearWorkbook(workbookPath, clientName, policyNumber, claimsWorkbookPaths, currentYearFolder, priorYearFolder);
        var rawCurrent = TryReadEhcSpendingSnapshot(workbook, workbookPath);
        if (rawCurrent is not null)
        {
            rawCurrent = AttachSnapshotMetadata(workbook, workbookPath, rawCurrent, "Current", clientName, policyNumber);
        }
        var (current, prior) = SelectServiceCategoryPeriodPair(externalSnapshots, rawCurrent);
        if (current is null)
        {
            result.Warnings.Add("ClaimsAnalysis_EHC extraction warning: could not find EHC spending trend data in a Summary claims workbook or ClaimsAnalysis_EHC.");
            return;
        }

        var priorLabel = prior?.PeriodLabel;
        if (string.IsNullOrWhiteSpace(priorLabel))
        {
            priorLabel = BuildPriorPeriodLabel(current.StartDate, current.EndDate) ?? "Prior Period";
            result.Warnings.Add("ClaimsAnalysis_EHC extraction warning: prior EHC service-category workbook was not found; prior period values were left at $0.");
        }

        var currentLabel = string.IsNullOrWhiteSpace(current.PeriodLabel)
            ? "Current Period"
            : current.PeriodLabel;

        foreach (var category in EhcSpendingCategories)
        {
            result.EhcSpendingTrends.Add(new ServiceCategoryTrend
            {
                Benefit = "Extended Health Care",
                Category = category,
                PriorPeriodLabel = priorLabel,
                PriorAmount = prior?.Amounts.GetValueOrDefault(category) ?? 0m,
                CurrentPeriodLabel = currentLabel,
                CurrentAmount = current.Amounts.GetValueOrDefault(category),
                SourceReference = $"{current.SourceReference}; Prior: {prior?.SourceReference ?? "not found"}"
            });
        }

        RecordClaimsTrendValidation(result, "ClaimsAnalysis_EHC", current, prior, priorLabel, currentLabel);
        LogEhcSpendingTrends(result.EhcSpendingTrends, current, prior);
    }

    private static void ReadDentalSpendingTrends(
        XLWorkbook workbook,
        string workbookPath,
        string? clientName,
        string? policyNumber,
        IEnumerable<string>? claimsWorkbookPaths,
        string? currentYearFolder,
        string? priorYearFolder,
        WorkbookExtractionResult result)
    {
        var externalSnapshots = ReadDentalSummarySnapshotsNearWorkbook(workbookPath, clientName, policyNumber, claimsWorkbookPaths, currentYearFolder, priorYearFolder);
        var rawCurrent = TryReadDentalRawSpendingSnapshot(workbook, workbookPath);
        if (rawCurrent is not null)
        {
            rawCurrent = AttachSnapshotMetadata(workbook, workbookPath, rawCurrent, "Current", clientName, policyNumber);
        }
        var (current, prior) = SelectServiceCategoryPeriodPair(externalSnapshots, rawCurrent);
        if (current is null)
        {
            result.Warnings.Add("ClaimsAnalysis_Dental extraction warning: could not find Dental spending trend data in a Summary claims workbook or ClaimsAnalysis_Dental.");
            return;
        }

        var priorLabel = prior?.PeriodLabel;
        if (string.IsNullOrWhiteSpace(priorLabel))
        {
            priorLabel = BuildPriorPeriodLabel(current.StartDate, current.EndDate) ?? "Prior Period";
            result.Warnings.Add("ClaimsAnalysis_Dental extraction warning: prior Dental spending trend source was not found; prior period values were left at $0.");
        }

        var currentLabel = string.IsNullOrWhiteSpace(current.PeriodLabel)
            ? "Current Period"
            : current.PeriodLabel;

        foreach (var category in DentalSpendingCategories)
        {
            result.DentalSpendingTrends.Add(new ServiceCategoryTrend
            {
                Benefit = "Dental",
                Category = category,
                PriorPeriodLabel = priorLabel,
                PriorAmount = prior?.Amounts.GetValueOrDefault(category) ?? 0m,
                CurrentPeriodLabel = currentLabel,
                CurrentAmount = current.Amounts.GetValueOrDefault(category),
                SourceReference = $"{current.SourceReference}; Prior: {prior?.SourceReference ?? "not found"}"
            });
        }

        RecordClaimsTrendValidation(result, "ClaimsAnalysis_Dental", current, prior, priorLabel, currentLabel);
        LogDentalSpendingTrends(result.DentalSpendingTrends, current, prior);
    }

    private static void ReadParamedicalServiceTrends(
        string workbookPath,
        string? clientName,
        string? policyNumber,
        IEnumerable<string>? claimsWorkbookPaths,
        string? currentYearFolder,
        string? priorYearFolder,
        WorkbookExtractionResult result)
    {
        var referenceCurrentLabel =
            result.EhcSpendingTrends.FirstOrDefault()?.CurrentPeriodLabel ??
            result.DentalSpendingTrends.FirstOrDefault()?.CurrentPeriodLabel;
        var monthlyTargets = BuildParamedicalPeriodTargets(result);
        var snapshots = ReadParamedicalSummarySnapshotsNearWorkbook(
            workbookPath,
            clientName,
            policyNumber,
            claimsWorkbookPaths,
            currentYearFolder,
            priorYearFolder);
        snapshots.AddRange(ReadParamedicalMonthlySnapshotsNearWorkbook(
            workbookPath,
            clientName,
            policyNumber,
            claimsWorkbookPaths,
            currentYearFolder,
            priorYearFolder,
            monthlyTargets));
        var current = !string.IsNullOrWhiteSpace(referenceCurrentLabel)
            ? snapshots.FirstOrDefault(snapshot => PeriodLabelsMatch(snapshot.PeriodLabel, referenceCurrentLabel))
            : null;
        current ??= snapshots
            .OrderByDescending(snapshot => snapshot.EndDate ?? DateOnly.MinValue)
            .FirstOrDefault();
        if (current is null)
        {
            result.Warnings.Add("ClaimsAnalysis_Paramedical extraction warning: could not find Health - Paramedical service rows in a Summary claims workbook.");
            return;
        }

        var prior = SelectPriorServiceCategorySnapshot(snapshots, current);

        var priorLabel = prior?.PeriodLabel;
        if (string.IsNullOrWhiteSpace(priorLabel))
        {
            priorLabel = BuildPriorPeriodLabel(current.StartDate, current.EndDate) ?? "Prior Period";
            result.Warnings.Add("ClaimsAnalysis_Paramedical extraction warning: prior paramedical Summary workbook was not found; prior period values were left at $0.");
        }

        var currentLabel = string.IsNullOrWhiteSpace(current.PeriodLabel)
            ? "Current Period"
            : current.PeriodLabel;

        foreach (var service in BuildParamedicalServiceOrder(current, prior))
        {
            result.ParamedicalServiceTrends.Add(new ServiceCategoryTrend
            {
                Benefit = "Extended Health Care",
                Category = service,
                PriorPeriodLabel = priorLabel,
                PriorAmount = prior?.Amounts.GetValueOrDefault(service) ?? 0m,
                CurrentPeriodLabel = currentLabel,
                CurrentAmount = current.Amounts.GetValueOrDefault(service),
                SourceReference = $"{current.SourceReference}; Prior: {prior?.SourceReference ?? "not found"}"
            });
        }

        RecordClaimsTrendValidation(result, "ClaimsAnalysis_Paramedical", current, prior, priorLabel, currentLabel);
        LogParamedicalServiceTrends(result.ParamedicalServiceTrends, current, prior);
    }

    private static IReadOnlyList<ParamedicalPeriodTarget> BuildParamedicalPeriodTargets(
        WorkbookExtractionResult result)
    {
        var paramedical = result.EhcSpendingTrends.FirstOrDefault(row =>
            row.Category.Equals("Paramedical", StringComparison.OrdinalIgnoreCase));
        if (paramedical is null)
        {
            return [];
        }

        return new[]
            {
                CreateParamedicalPeriodTarget(
                    paramedical.PriorPeriodLabel,
                    paramedical.PriorAmount),
                CreateParamedicalPeriodTarget(
                    paramedical.CurrentPeriodLabel,
                    paramedical.CurrentAmount)
            }
            .OfType<ParamedicalPeriodTarget>()
            .GroupBy(target => $"{target.StartDate:yyyy-MM-dd}|{target.EndDate:yyyy-MM-dd}")
            .Select(group => group.First())
            .ToList();
    }

    private static ParamedicalPeriodTarget? CreateParamedicalPeriodTarget(
        string periodLabel,
        decimal expectedTotal)
    {
        if (string.IsNullOrWhiteSpace(periodLabel) || expectedTotal == 0m)
        {
            return null;
        }

        var dates = TryParsePeriodDatesFromLabel(periodLabel);
        return dates.StartDate.HasValue && dates.EndDate.HasValue
            ? new ParamedicalPeriodTarget(
                periodLabel,
                dates.StartDate.Value,
                dates.EndDate.Value,
                expectedTotal)
            : null;
    }

    private static List<ServiceCategorySnapshot> ReadEhcSummarySnapshotsNearWorkbook(
        string workbookPath,
        string? clientName,
        string? policyNumber,
        IEnumerable<string>? claimsWorkbookPaths,
        string? currentYearFolder,
        string? priorYearFolder)
    {
        return ReadSummarySnapshotsNearWorkbook(
            workbookPath,
            clientName,
            policyNumber,
            claimsWorkbookPaths,
            currentYearFolder,
            priorYearFolder,
            (claimsWorkbook, candidatePath) =>
                TryReadEhcSummarySnapshot(claimsWorkbook, candidatePath) ??
                TryReadEhcSpendingSnapshot(claimsWorkbook, candidatePath));
    }

    private static List<ServiceCategorySnapshot> ReadParamedicalSummarySnapshotsNearWorkbook(
        string workbookPath,
        string? clientName,
        string? policyNumber,
        IEnumerable<string>? claimsWorkbookPaths,
        string? currentYearFolder,
        string? priorYearFolder)
    {
        return ReadSummarySnapshotsNearWorkbook(workbookPath, clientName, policyNumber, claimsWorkbookPaths, currentYearFolder, priorYearFolder, TryReadParamedicalSummarySnapshot);
    }

    private static List<ServiceCategorySnapshot> ReadParamedicalMonthlySnapshotsNearWorkbook(
        string workbookPath,
        string? clientName,
        string? policyNumber,
        IEnumerable<string>? claimsWorkbookPaths,
        string? currentYearFolder,
        string? priorYearFolder,
        IReadOnlyList<ParamedicalPeriodTarget> targets)
    {
        if (targets.Count == 0)
        {
            return [];
        }

        var snapshots = new List<ServiceCategorySnapshot>();
        foreach (var candidate in FindClaimsWorkbookCandidates(
                     workbookPath,
                     clientName,
                     claimsWorkbookPaths,
                     currentYearFolder,
                     priorYearFolder))
        {
            try
            {
                using var claimsWorkbook = OpenWorkbookSharedRead(candidate);
                foreach (var target in targets)
                {
                    var snapshot = TryReadParamedicalMonthlySnapshot(
                        claimsWorkbook,
                        candidate,
                        target);
                    if (snapshot is null)
                    {
                        continue;
                    }

                    snapshots.Add(AttachSnapshotMetadata(
                        claimsWorkbook,
                        candidate,
                        snapshot,
                        GetCandidatePeriodRole(candidate, currentYearFolder, priorYearFolder),
                        clientName,
                        policyNumber));
                }
            }
            catch (InvalidDataException)
            {
                throw;
            }
            catch
            {
                // Ignore unrelated, locked, or non-Manulife claim files in the same folder.
            }
        }

        return snapshots;
    }

    private static List<ServiceCategorySnapshot> ReadSummarySnapshotsNearWorkbook(
        string workbookPath,
        string? clientName,
        string? policyNumber,
        IEnumerable<string>? claimsWorkbookPaths,
        string? currentYearFolder,
        string? priorYearFolder,
        Func<XLWorkbook, string, ServiceCategorySnapshot?> snapshotReader)
    {
        var snapshots = new List<ServiceCategorySnapshot>();
        foreach (var candidate in FindClaimsWorkbookCandidates(workbookPath, clientName, claimsWorkbookPaths, currentYearFolder, priorYearFolder))
        {
            try
            {
                using var claimsWorkbook = OpenWorkbookSharedRead(candidate);
                var snapshot = snapshotReader(claimsWorkbook, candidate);
                if (snapshot is not null)
                {
                    snapshots.Add(AttachSnapshotMetadata(
                        claimsWorkbook,
                        candidate,
                        snapshot,
                        GetCandidatePeriodRole(candidate, currentYearFolder, priorYearFolder),
                        clientName,
                        policyNumber));
                }
            }
            catch (InvalidDataException)
            {
                throw;
            }
            catch
            {
                // Ignore unrelated, locked, or non-Manulife claim files in the same folder.
            }
        }

        return snapshots
            .OrderBy(snapshot => snapshot.EndDate ?? DateOnly.MinValue)
            .ToList();
    }

    private static ServiceCategorySnapshot AttachSnapshotMetadata(
        XLWorkbook workbook,
        string sourcePath,
        ServiceCategorySnapshot snapshot,
        string periodRole,
        string? expectedClientName,
        string? expectedPolicyNumber)
    {
        var metadata = ReadClaimsWorkbookMetadata(workbook, sourcePath);
        ValidateClaimsWorkbookIdentity(sourcePath, metadata, expectedClientName, expectedPolicyNumber);

        var periodLabel = !string.IsNullOrWhiteSpace(snapshot.PeriodLabel) && !snapshot.PeriodLabel.Equals("Current Period", StringComparison.OrdinalIgnoreCase)
            ? snapshot.PeriodLabel
            : metadata.PeriodLabel;
        var startDate = snapshot.StartDate ?? metadata.StartDate;
        var endDate = snapshot.EndDate ?? metadata.EndDate;

        return snapshot with
        {
            PeriodLabel = string.IsNullOrWhiteSpace(periodLabel) ? "Current Period" : periodLabel,
            StartDate = startDate,
            EndDate = endDate,
            PeriodRole = periodRole,
            ClientName = metadata.ClientName,
            PolicyNumber = metadata.PolicyNumber
        };
    }

    private static (ServiceCategorySnapshot? Current, ServiceCategorySnapshot? Prior) SelectServiceCategoryPeriodPair(
        IReadOnlyList<ServiceCategorySnapshot> externalSnapshots,
        ServiceCategorySnapshot? rawCurrent)
    {
        var candidates = externalSnapshots
            .Concat(rawCurrent is null ? [] : [rawCurrent])
            .Where(snapshot => snapshot.Amounts.Values.Any(value => value != 0m))
            .GroupBy(SnapshotIdentity, StringComparer.OrdinalIgnoreCase)
            .Select(group => group.OrderByDescending(SnapshotSourcePriority).First())
            .ToList();

        if (candidates.Count == 0)
        {
            return (null, null);
        }

        var current = candidates
            .Where(snapshot => snapshot.EndDate.HasValue)
            .OrderByDescending(snapshot => snapshot.EndDate)
            .ThenByDescending(SnapshotSourcePriority)
            .FirstOrDefault()
            ?? candidates
                .Where(snapshot => snapshot.PeriodRole.Equals("Current", StringComparison.OrdinalIgnoreCase))
                .OrderByDescending(SnapshotSourcePriority)
                .FirstOrDefault()
            ?? candidates.OrderByDescending(SnapshotSourcePriority).First();

        return (current, SelectPriorServiceCategorySnapshot(candidates, current));
    }

    private static int SnapshotSourcePriority(ServiceCategorySnapshot snapshot)
    {
        if (snapshot.PeriodRole.Equals("Current", StringComparison.OrdinalIgnoreCase))
        {
            return 4;
        }

        if (snapshot.PeriodRole.Equals("Prior", StringComparison.OrdinalIgnoreCase))
        {
            return 3;
        }

        if (snapshot.SourceReference.StartsWith("Summary!", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }

        return 1;
    }

    private static string SnapshotIdentity(ServiceCategorySnapshot snapshot)
    {
        var source = string.IsNullOrWhiteSpace(snapshot.SourcePath)
            ? "source"
            : Path.GetFullPath(snapshot.SourcePath);
        var endDate = snapshot.EndDate?.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) ?? NormalizeClientMatchText(snapshot.PeriodLabel);
        return $"{source}|{endDate}|{NormalizeClientMatchText(snapshot.SourceReference)}";
    }

    private static IReadOnlyList<string> FindClaimsWorkbookCandidates(
        string workbookPath,
        string? clientName,
        IEnumerable<string>? claimsWorkbookPaths,
        string? currentYearFolder,
        string? priorYearFolder)
    {
        var currentFullPath = Path.GetFullPath(workbookPath);
        var folderMode = IsUsableDirectory(currentYearFolder) || IsUsableDirectory(priorYearFolder);
        // Explicit selections and designated source folders are authoritative inputs.
        // Discover them by Excel format, then let the workbook readers decide whether
        // their contents contain supported claims data.
        var currentFolderCandidates = EnumerateExcelFiles(currentYearFolder)
            .Where(path => !string.Equals(Path.GetFullPath(path), currentFullPath, StringComparison.OrdinalIgnoreCase))
            .ToList();
        var priorFolderCandidates = EnumerateExcelFiles(priorYearFolder)
            .Where(path => !string.Equals(Path.GetFullPath(path), currentFullPath, StringComparison.OrdinalIgnoreCase))
            .ToList();
        var folderCandidates = currentFolderCandidates
            .Concat(priorFolderCandidates)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        var explicitCandidates = (claimsWorkbookPaths ?? Enumerable.Empty<string>())
            .Where(path => !string.IsNullOrWhiteSpace(path) && File.Exists(path))
            .Where(path =>
                path.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                path.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase) ||
                path.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
            .Where(path => !string.Equals(Path.GetFullPath(path), currentFullPath, StringComparison.OrdinalIgnoreCase))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        if (folderMode)
        {
            return folderCandidates
                .Concat(explicitCandidates)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        var nearbyCandidates = FindNearbyClaimsWorkbookCandidates(
            workbookPath,
            clientName,
            allowBroadFallback: explicitCandidates.Count == 0 && folderCandidates.Count == 0);
        var candidates = folderCandidates
            .Concat(explicitCandidates)
            .Concat(nearbyCandidates)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        return candidates;
    }

    private static IEnumerable<string> EnumerateExcelFiles(string? folder)
    {
        return !string.IsNullOrWhiteSpace(folder) && Directory.Exists(folder)
            ? Directory.EnumerateFiles(folder, "*.xlsx", SearchOption.TopDirectoryOnly)
                .Concat(Directory.EnumerateFiles(folder, "*.xlsm", SearchOption.TopDirectoryOnly))
                .Concat(Directory.EnumerateFiles(folder, "*.xls", SearchOption.TopDirectoryOnly))
            : [];
    }

    private static bool FileNameLooksLikeClaimsWorkbook(string path)
    {
        var fileName = Path.GetFileNameWithoutExtension(path);
        return fileName.Contains("Claims", StringComparison.OrdinalIgnoreCase) ||
               fileName.Contains("Claim", StringComparison.OrdinalIgnoreCase) ||
               fileName.Contains("Client_Copy", StringComparison.OrdinalIgnoreCase) ||
               fileName.Contains("Client Copy", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsUsableDirectory(string? folder)
    {
        return !string.IsNullOrWhiteSpace(folder) && Directory.Exists(folder);
    }

    private static string GetCandidatePeriodRole(string path, string? currentYearFolder, string? priorYearFolder)
    {
        var isCurrent = IsPathInFolder(path, currentYearFolder);
        var isPrior = IsPathInFolder(path, priorYearFolder);
        if (isCurrent && isPrior)
        {
            return "";
        }

        if (isPrior)
        {
            return "Prior";
        }

        if (isCurrent)
        {
            return "Current";
        }

        return "";
    }

    private static bool IsPathInFolder(string path, string? folder)
    {
        if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
        {
            return false;
        }

        var fullPath = Path.GetFullPath(path);
        var fullFolder = Path.GetFullPath(folder).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + Path.DirectorySeparatorChar;
        return fullPath.StartsWith(fullFolder, StringComparison.OrdinalIgnoreCase);
    }

    private static IReadOnlyList<string> FindNearbyClaimsWorkbookCandidates(
        string workbookPath,
        string? clientName,
        bool allowBroadFallback = true)
    {
        var folder = Path.GetDirectoryName(workbookPath);
        if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
        {
            return [];
        }

        var currentFullPath = Path.GetFullPath(workbookPath);
        var candidates = Directory.EnumerateFiles(folder, "*.xlsx", SearchOption.TopDirectoryOnly)
            .Concat(Directory.EnumerateFiles(folder, "*.xlsm", SearchOption.TopDirectoryOnly))
            .Concat(Directory.EnumerateFiles(folder, "*.xls", SearchOption.TopDirectoryOnly))
            .Where(path => !string.Equals(Path.GetFullPath(path), currentFullPath, StringComparison.OrdinalIgnoreCase))
            .Where(path => Path.GetExtension(path).Equals(".xls", StringComparison.OrdinalIgnoreCase) || FileNameLooksLikeClaimsWorkbook(path))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        var prefixes = BuildClientFilePrefixes(clientName, workbookPath);
        List<string> clientMatches = prefixes.Count == 0
            ? []
            : candidates
                .Where(path => FileNameMatchesClientPrefix(path, prefixes))
                .ToList();

        return (clientMatches.Count > 0 || !allowBroadFallback ? clientMatches : candidates)
            .OrderBy(path => Path.GetFileName(path), StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static List<string> BuildClientFilePrefixes(string? clientName, string workbookPath)
    {
        var prefixes = new List<string>();
        AddNamePrefixCandidates(prefixes, clientName);
        AddNamePrefixCandidates(prefixes, GetClientWorkbookPrefix(workbookPath));

        return prefixes
            .Where(prefix => prefix.Length >= 3)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static void AddNamePrefixCandidates(List<string> prefixes, string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return;
        }

        var token = Regex.Matches(value, @"[A-Za-z0-9]+")
            .Cast<Match>()
            .Select(match => match.Value)
            .FirstOrDefault(token =>
                token.Any(char.IsLetter) &&
                !token.Equals("The", StringComparison.OrdinalIgnoreCase) &&
                !token.Equals("Ltd", StringComparison.OrdinalIgnoreCase) &&
                !token.Equals("Limited", StringComparison.OrdinalIgnoreCase) &&
                !token.Equals("Inc", StringComparison.OrdinalIgnoreCase) &&
                !token.Equals("Client", StringComparison.OrdinalIgnoreCase) &&
                !token.Equals("Copy", StringComparison.OrdinalIgnoreCase) &&
                !token.Equals("Claims", StringComparison.OrdinalIgnoreCase));
        if (string.IsNullOrWhiteSpace(token))
        {
            return;
        }

        var normalized = NormalizeClientMatchText(token);
        if (normalized.Length >= 3)
        {
            prefixes.Add(normalized[..3]);
        }

        if (normalized.Length >= 4)
        {
            prefixes.Add(normalized[..4]);
        }
    }

    private static bool FileNameMatchesClientPrefix(string path, IReadOnlyCollection<string> prefixes)
    {
        var fileName = NormalizeClientMatchText(Path.GetFileNameWithoutExtension(path));
        return prefixes.Any(prefix => fileName.StartsWith(prefix, StringComparison.OrdinalIgnoreCase));
    }

    private static bool FileNameLooksLikeGenericClaimsWorkbook(string path)
    {
        var tokens = Regex.Matches(Path.GetFileNameWithoutExtension(path), @"[A-Za-z0-9]+")
            .Cast<Match>()
            .Select(match => match.Value)
            .ToList();
        if (tokens.Contains("Client", StringComparer.OrdinalIgnoreCase) &&
            tokens.Contains("Copy", StringComparer.OrdinalIgnoreCase))
        {
            return true;
        }

        var claimsIndex = tokens.FindIndex(token =>
            token.Equals("Claim", StringComparison.OrdinalIgnoreCase) ||
            token.Equals("Claims", StringComparison.OrdinalIgnoreCase));
        if (claimsIndex < 0)
        {
            return false;
        }

        return tokens.Take(claimsIndex).All(IsGenericClaimsFilePrefixToken);
    }

    private static bool IsGenericClaimsFilePrefixToken(string token)
    {
        if (string.IsNullOrWhiteSpace(token) || token.All(char.IsDigit))
        {
            return true;
        }

        var alphaPrefix = Regex.Match(token, @"^[A-Za-z]+").Value;
        if (!string.IsNullOrWhiteSpace(alphaPrefix) && ParseMonthName(alphaPrefix).HasValue)
        {
            return true;
        }

        return token.Equals("Client", StringComparison.OrdinalIgnoreCase) ||
               token.Equals("Manulife", StringComparison.OrdinalIgnoreCase) ||
               token.Equals("Carrier", StringComparison.OrdinalIgnoreCase) ||
               token.Equals("Group", StringComparison.OrdinalIgnoreCase) ||
               token.Equals("Paid", StringComparison.OrdinalIgnoreCase) ||
               token.Equals("By", StringComparison.OrdinalIgnoreCase) ||
               token.Equals("Service", StringComparison.OrdinalIgnoreCase) ||
               token.Equals("Category", StringComparison.OrdinalIgnoreCase) ||
               token.Equals("To", StringComparison.OrdinalIgnoreCase) ||
               token.Equals("Report", StringComparison.OrdinalIgnoreCase);
    }

    private static bool PeriodLabelsMatch(string left, string right)
    {
        return NormalizeClientMatchText(left).Equals(NormalizeClientMatchText(right), StringComparison.OrdinalIgnoreCase);
    }

    private static IEnumerable<IXLWorksheet> WorksheetsWithPreferredFirst(
        XLWorkbook workbook,
        string preferredWorksheetName)
    {
        return workbook.Worksheets
            .OrderBy(sheet => sheet.Name.Equals(preferredWorksheetName, StringComparison.OrdinalIgnoreCase) ? 0 : 1);
    }

    private static string NormalizeClientMatchText(string value)
    {
        return new string(value.Where(char.IsLetterOrDigit).Select(char.ToUpperInvariant).ToArray());
    }

    private static ServiceCategorySnapshot? TryReadParamedicalSummarySnapshot(XLWorkbook workbook, string sourcePath)
    {
        foreach (var sheet in WorksheetsWithPreferredFirst(workbook, "Summary"))
        {
            var snapshot = TryReadParamedicalSummarySheet(sheet, sourcePath);
            if (snapshot is not null)
            {
                return snapshot;
            }
        }

        return TryReadParamedicalServiceHeaderSnapshot(workbook, sourcePath);
    }

    private static ServiceCategorySnapshot? TryReadParamedicalSummarySheet(IXLWorksheet sheet, string sourcePath)
    {
        var headerRow = FindRow(sheet, "Total amount paid") ?? FindRow(sheet, "Total amount paid ($)");
        if (headerRow is null || HasMonthlyServiceColumns(sheet, headerRow.Value))
        {
            return null;
        }

        var benefitColumn = FindColumn(sheet, headerRow.Value, "Benefit") ?? 1;
        var categoryColumn = FindColumn(sheet, headerRow.Value, "Benefit category") ?? 2;
        var serviceColumn = FindColumn(sheet, headerRow.Value, "Service") ?? 3;
        var paidColumn = FindColumn(sheet, headerRow.Value, "Total amount paid") ?? 7;
        var lastRow = sheet.LastRowUsed()?.RowNumber() ?? headerRow.Value;
        var amounts = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
        var currentBenefit = "";
        var currentCategory = "";
        var foundParamedical = false;
        decimal? paramedicalTotal = null;

        for (var row = headerRow.Value + 1; row <= lastRow; row++)
        {
            var benefit = Text(sheet.Cell(row, benefitColumn));
            if (!string.IsNullOrWhiteSpace(benefit))
            {
                currentBenefit = benefit;
            }

            if (!currentBenefit.Equals("Health", StringComparison.OrdinalIgnoreCase))
            {
                if (foundParamedical)
                {
                    break;
                }

                continue;
            }

            var rawCategory = Text(sheet.Cell(row, categoryColumn));
            if (!string.IsNullOrWhiteSpace(rawCategory))
            {
                currentCategory = rawCategory;
            }

            if (!currentCategory.Equals("Paramedical", StringComparison.OrdinalIgnoreCase))
            {
                if (foundParamedical)
                {
                    break;
                }

                continue;
            }

            var service = Text(sheet.Cell(row, serviceColumn));
            if (service.Equals("Subtotal", StringComparison.OrdinalIgnoreCase))
            {
                paramedicalTotal = Number(sheet.Cell(row, paidColumn));
                break;
            }

            if (string.IsNullOrWhiteSpace(service))
            {
                continue;
            }

            foundParamedical = true;
            var normalizedService = NormalizeParamedicalServiceName(service);
            amounts[normalizedService] = amounts.GetValueOrDefault(normalizedService) +
                                         (Number(sheet.Cell(row, paidColumn)) ?? 0m);
        }

        if (!foundParamedical)
        {
            return null;
        }

        amounts["Paramedical Total"] = paramedicalTotal ?? amounts.Values.Sum();

        var (periodLabel, startDate, endDate) = ReadExperiencePeriodLabel(sheet, sourcePath);
        return new ServiceCategorySnapshot(
            periodLabel,
            startDate,
            endDate,
            amounts,
            sourcePath,
            $"{sheet.Name}!A{headerRow.Value}:{ColumnName(sheet.LastColumnUsed()?.ColumnNumber() ?? paidColumn)}{lastRow}");
    }

    // Some Manulife exports provide the practitioner table directly instead of the
    // normalized Summary layout.  Treat that narrowly-defined Service table as
    // paramedical evidence rather than discarding it for lack of Benefit/Category columns.
    private static ServiceCategorySnapshot? TryReadParamedicalServiceHeaderSnapshot(XLWorkbook workbook, string sourcePath)
    {
        foreach (var sheet in workbook.Worksheets)
        {
            var headerRow = FindRow(sheet, "Service");
            if (headerRow is null)
            {
                continue;
            }

            var serviceColumn = FindColumn(sheet, headerRow.Value, "Service") ?? 1;
            if (HasMonthlyServiceColumns(sheet, headerRow.Value))
            {
                continue;
            }

            var percentColumn = FindColumn(sheet, headerRow.Value, "% of parameds");
            var lastColumn = sheet.LastColumnUsed()?.ColumnNumber() ?? serviceColumn;
            var amountColumns = Enumerable.Range(serviceColumn + 1, Math.Max(0, lastColumn - serviceColumn))
                .Where(column => !percentColumn.HasValue || column != percentColumn.Value)
                .ToList();
            if (amountColumns.Count == 0)
            {
                continue;
            }

            // These exports show prior then current amounts.  If only one amount
            // column is present, it is the current period.
            var paidColumn = amountColumns.Count >= 2 ? amountColumns[1] : amountColumns[0];
            var lastRow = sheet.LastRowUsed()?.RowNumber() ?? headerRow.Value;
            var amounts = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
            decimal? paramedicalTotal = null;

            for (var row = headerRow.Value + 1; row <= lastRow; row++)
            {
                var service = Text(sheet.Cell(row, serviceColumn));
                if (string.IsNullOrWhiteSpace(service))
                {
                    continue;
                }

                if (service.Equals("Paramedical Practitioners", StringComparison.OrdinalIgnoreCase) ||
                    service.Equals("Paramedical Total", StringComparison.OrdinalIgnoreCase) ||
                    service.Equals("Total", StringComparison.OrdinalIgnoreCase))
                {
                    paramedicalTotal = Number(sheet.Cell(row, paidColumn));
                    break;
                }

                var normalizedService = NormalizeParamedicalServiceName(service);
                if (!IsRecognizedParamedicalService(normalizedService))
                {
                    continue;
                }

                amounts[normalizedService] = amounts.GetValueOrDefault(normalizedService) +
                                             (Number(sheet.Cell(row, paidColumn)) ?? 0m);
            }

            if (amounts.Count == 0)
            {
                continue;
            }

            amounts["Paramedical Total"] = paramedicalTotal ?? amounts.Values.Sum();
            var (periodLabel, startDate, endDate) = ReadExperiencePeriodLabel(sheet, sourcePath);
            return new ServiceCategorySnapshot(
                periodLabel,
                startDate,
                endDate,
                amounts,
                sourcePath,
                $"{sheet.Name}!{ColumnName(serviceColumn)}{headerRow.Value}:{ColumnName(lastColumn)}{lastRow}");
        }

        return null;
    }

    private static ServiceCategorySnapshot? TryReadParamedicalMonthlySnapshot(
        XLWorkbook workbook,
        string sourcePath,
        ParamedicalPeriodTarget target)
    {
        foreach (var sheet in workbook.Worksheets)
        {
            var headerRow = FindRow(sheet, "Benefit category");
            if (!headerRow.HasValue || headerRow.Value <= 1)
            {
                continue;
            }

            var benefitColumn = FindColumn(sheet, headerRow.Value, "Benefit");
            var categoryColumn = FindColumn(sheet, headerRow.Value, "Benefit category");
            var serviceColumn = FindColumn(sheet, headerRow.Value, "Service");
            if (!benefitColumn.HasValue || !categoryColumn.HasValue || !serviceColumn.HasValue)
            {
                continue;
            }

            var lastColumn = sheet.LastColumnUsed()?.ColumnNumber() ?? serviceColumn.Value;
            var monthColumns = Enumerable.Range(1, lastColumn)
                .Select(column => new
                {
                    Column = column,
                    Month = TryReadMonthHeader(sheet.Cell(headerRow.Value - 1, column))
                })
                .Where(item =>
                    item.Month.HasValue &&
                    item.Month.Value >= target.StartDate &&
                    item.Month.Value <= target.EndDate)
                .Select(item => (item.Column, Month: item.Month!.Value))
                .ToList();
            if (monthColumns.Count == 0)
            {
                continue;
            }

            var amounts = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
            var currentBenefit = "";
            var currentCategory = "";
            var lastRow = sheet.LastRowUsed()?.RowNumber() ?? headerRow.Value;
            for (var row = headerRow.Value + 1; row <= lastRow; row++)
            {
                var benefit = Text(sheet.Cell(row, benefitColumn.Value));
                if (!string.IsNullOrWhiteSpace(benefit))
                {
                    currentBenefit = benefit;
                }

                var category = Text(sheet.Cell(row, categoryColumn.Value));
                if (!string.IsNullOrWhiteSpace(category))
                {
                    currentCategory = category;
                }

                if (!currentBenefit.Equals("Health", StringComparison.OrdinalIgnoreCase) ||
                    !currentCategory.Equals("Paramedical", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var service = NormalizeParamedicalServiceName(
                    Text(sheet.Cell(row, serviceColumn.Value)));
                if (string.IsNullOrWhiteSpace(service) ||
                    service.Equals("Subtotal", StringComparison.OrdinalIgnoreCase) ||
                    service.Equals("Total", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var amount = monthColumns.Sum(item =>
                    Number(sheet.Cell(row, item.Column)) ?? 0m);
                amounts[service] = amounts.GetValueOrDefault(service) + amount;
            }

            if (amounts.Count == 0)
            {
                continue;
            }

            var total = amounts.Values.Sum();
            if (Math.Abs(total - target.ExpectedTotal) > 0.05m)
            {
                continue;
            }

            amounts["Paramedical Total"] = total;
            var firstMonth = monthColumns.First();
            var lastMonth = monthColumns.Last();
            var selectedPeriod =
                $"{firstMonth.Month.ToString("MMM yyyy", CultureInfo.InvariantCulture)}-" +
                $"{lastMonth.Month.ToString("MMM yyyy", CultureInfo.InvariantCulture)}";
            return new ServiceCategorySnapshot(
                target.PeriodLabel,
                target.StartDate,
                target.EndDate,
                amounts,
                sourcePath,
                $"{sheet.Name}!{ColumnName(benefitColumn.Value)}{headerRow.Value}:" +
                $"{ColumnName(lastMonth.Column)}{lastRow} " +
                $"(selected {selectedPeriod}; reconciled to EHC paramedical total)");
        }

        return null;
    }

    private static bool HasMonthlyServiceColumns(IXLWorksheet sheet, int headerRow)
    {
        if (headerRow <= 1)
        {
            return false;
        }

        var lastColumn = sheet.LastColumnUsed()?.ColumnNumber() ?? 0;
        return Enumerable.Range(1, lastColumn)
            .Count(column => TryReadMonthHeader(sheet.Cell(headerRow - 1, column)).HasValue) >= 2;
    }

    private static DateOnly? TryReadMonthHeader(IXLCell cell)
    {
        if (cell.TryGetValue<DateTime>(out var dateTime))
        {
            return new DateOnly(dateTime.Year, dateTime.Month, 1);
        }

        var text = Text(cell);
        if (!DateTime.TryParse(
                text,
                CultureInfo.InvariantCulture,
                DateTimeStyles.AllowWhiteSpaces,
                out dateTime))
        {
            return null;
        }

        return new DateOnly(dateTime.Year, dateTime.Month, 1);
    }

    private static bool IsRecognizedParamedicalService(string service)
    {
        return ParamedicalServiceOrder.Contains(service, StringComparer.OrdinalIgnoreCase) ||
               service.Equals("Naturopath", StringComparison.OrdinalIgnoreCase) ||
               service.Equals("Massage", StringComparison.OrdinalIgnoreCase);
    }

    private static List<string> BuildParamedicalServiceOrder(
        ServiceCategorySnapshot current,
        ServiceCategorySnapshot? prior)
    {
        var availableServices = (prior?.Amounts.Keys ?? Enumerable.Empty<string>())
            .Concat(current.Amounts.Keys)
            .Where(service => !service.Equals("Paramedical Total", StringComparison.OrdinalIgnoreCase))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        var services = ParamedicalServiceOrder
            .Where(service => availableServices.Contains(service, StringComparer.OrdinalIgnoreCase))
            .ToList();

        foreach (var service in availableServices)
        {
            if (!services.Contains(service, StringComparer.OrdinalIgnoreCase))
            {
                services.Add(service);
            }
        }

        services.Add("Paramedical Total");
        return services;
    }

    private static string NormalizeParamedicalServiceName(string service)
    {
        var normalized = service.Trim();
        if (normalized.Contains("Massage", StringComparison.OrdinalIgnoreCase))
        {
            return "Registered Massage Therapist";
        }

        if (normalized.Contains("Physio", StringComparison.OrdinalIgnoreCase))
        {
            return "Physiotherapist";
        }

        if (normalized.Contains("Clinical Counsel", StringComparison.OrdinalIgnoreCase))
        {
            return "Clinical Counsellor";
        }

        if (normalized.Contains("Acupuncture", StringComparison.OrdinalIgnoreCase))
        {
            return "Acupuncture";
        }

        if (normalized.Contains("Chiropractor", StringComparison.OrdinalIgnoreCase))
        {
            return "Chiropractor";
        }

        if (normalized.Contains("Psychologist", StringComparison.OrdinalIgnoreCase))
        {
            return "Psychologist";
        }

        return normalized;
    }

    private static List<ServiceCategorySnapshot> ReadDentalSummarySnapshotsNearWorkbook(
        string workbookPath,
        string? clientName,
        string? policyNumber,
        IEnumerable<string>? claimsWorkbookPaths,
        string? currentYearFolder,
        string? priorYearFolder)
    {
        return ReadSummarySnapshotsNearWorkbook(
            workbookPath,
            clientName,
            policyNumber,
            claimsWorkbookPaths,
            currentYearFolder,
            priorYearFolder,
            (claimsWorkbook, candidatePath) =>
                TryReadDentalSummarySnapshot(claimsWorkbook, candidatePath) ??
                TryReadDentalRawSpendingSnapshot(claimsWorkbook, candidatePath));
    }

    private static ServiceCategorySnapshot? SelectCurrentServiceCategorySnapshot(
        IReadOnlyList<ServiceCategorySnapshot> externalSnapshots,
        ServiceCategorySnapshot? rawCurrent)
    {
        var explicitCurrent = externalSnapshots
            .Where(snapshot => snapshot.PeriodRole.Equals("Current", StringComparison.OrdinalIgnoreCase))
            .OrderByDescending(snapshot => snapshot.EndDate ?? DateOnly.MinValue)
            .FirstOrDefault();
        if (explicitCurrent is not null)
        {
            return explicitCurrent;
        }

        if (externalSnapshots.Count == 0)
        {
            return rawCurrent;
        }

        if (rawCurrent?.EndDate is not null)
        {
            var matchingExternal = externalSnapshots
                .FirstOrDefault(snapshot => snapshot.EndDate == rawCurrent.EndDate);
            if (matchingExternal is not null)
            {
                return matchingExternal;
            }
        }

        return externalSnapshots
            .OrderByDescending(snapshot => snapshot.EndDate ?? DateOnly.MinValue)
            .FirstOrDefault()
            ?? rawCurrent;
    }

    private static ServiceCategorySnapshot? SelectPriorServiceCategorySnapshot(
        IReadOnlyList<ServiceCategorySnapshot> snapshots,
        ServiceCategorySnapshot current)
    {
        var explicitPrior = snapshots
            .Where(snapshot => snapshot.PeriodRole.Equals("Prior", StringComparison.OrdinalIgnoreCase) &&
                               IsValidPriorSnapshot(snapshot, current))
            .OrderByDescending(snapshot => snapshot.EndDate ?? DateOnly.MinValue)
            .FirstOrDefault();
        if (explicitPrior is not null)
        {
            return explicitPrior;
        }

        if (current.EndDate.HasValue)
        {
            return snapshots
                .Where(snapshot =>
                    IsValidPriorSnapshot(snapshot, current) &&
                    snapshot.EndDate.HasValue &&
                    snapshot.EndDate.Value < current.EndDate.Value)
                .OrderByDescending(snapshot => snapshot.EndDate)
                .FirstOrDefault();
        }

        return null;
    }

    private static bool IsValidPriorSnapshot(ServiceCategorySnapshot candidate, ServiceCategorySnapshot current)
    {
        if (ReferenceEquals(candidate, current))
        {
            return false;
        }

        var sameSource = !string.IsNullOrWhiteSpace(candidate.SourcePath) &&
                         !string.IsNullOrWhiteSpace(current.SourcePath) &&
                         string.Equals(Path.GetFullPath(candidate.SourcePath), Path.GetFullPath(current.SourcePath), StringComparison.OrdinalIgnoreCase);
        if (sameSource)
        {
            return false;
        }

        if (candidate.EndDate.HasValue && current.EndDate.HasValue && candidate.EndDate.Value >= current.EndDate.Value)
        {
            return false;
        }

        if (!candidate.EndDate.HasValue &&
            PeriodLabelsMatch(candidate.PeriodLabel, current.PeriodLabel))
        {
            return false;
        }

        return true;
    }

    private static int MergeSupportingClaimsWorkbookHistory(
        List<ServiceCategoryTrend> currentRows,
        IReadOnlyList<ServiceCategorySnapshot> supportingSnapshots,
        string area,
        WorkbookExtractionResult result)
    {
        var snapshots = supportingSnapshots
            .Where(snapshot => snapshot.Amounts.Values.Any(value => value != 0m))
            .ToList();
        if (snapshots.Count == 0)
        {
            return 0;
        }

        if (currentRows.Count == 0)
        {
            return AddSupportingClaimsWorkbookPeriodPair(currentRows, snapshots, area, result);
        }

        var currentSnapshot = BuildSnapshotFromCurrentRows(currentRows);
        var priorSnapshot = SelectSupportingHistorySnapshot(snapshots, currentSnapshot.EndDate);
        if (priorSnapshot is null)
        {
            return 0;
        }

        var updated = 0;
        var priorLabel = string.IsNullOrWhiteSpace(priorSnapshot.PeriodLabel)
            ? "Prior Period"
            : priorSnapshot.PeriodLabel;

        foreach (var row in currentRows)
        {
            if (!TryGetSnapshotAmount(priorSnapshot, row.Category, out var priorAmount))
            {
                continue;
            }

            row.PriorPeriodLabel = priorLabel;
            row.PriorAmount = priorAmount;
            row.SourceReference = AppendSourceReference(
                row.SourceReference,
                $"Prior Excel history - {Path.GetFileName(priorSnapshot.SourcePath)} ({priorSnapshot.SourceReference})");
            updated++;
        }

        if (updated > 0)
        {
            RecordClaimsTrendValidation(result, area, currentSnapshot, priorSnapshot, priorLabel, currentSnapshot.PeriodLabel);
        }

        return updated;
    }

    private static int AddSupportingClaimsWorkbookPeriodPair(
        List<ServiceCategoryTrend> currentRows,
        IReadOnlyList<ServiceCategorySnapshot> snapshots,
        string area,
        WorkbookExtractionResult result)
    {
        var (current, prior) = SelectServiceCategoryPeriodPair(snapshots, null);
        if (current is null)
        {
            return 0;
        }

        var categories = current.Amounts.Keys
            .Concat(prior?.Amounts.Keys ?? Enumerable.Empty<string>())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        var priorLabel = string.IsNullOrWhiteSpace(prior?.PeriodLabel)
            ? BuildPriorPeriodLabel(current.StartDate, current.EndDate) ?? "Prior Period"
            : prior.PeriodLabel;
        var currentLabel = string.IsNullOrWhiteSpace(current.PeriodLabel)
            ? "Current Period"
            : current.PeriodLabel;
        var added = 0;

        foreach (var category in categories)
        {
            if (category.Equals("Paramedical Total", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            currentRows.Add(new ServiceCategoryTrend
            {
                Benefit = area.Contains("Dental", StringComparison.OrdinalIgnoreCase) ? "Dental" : "Extended Health Care",
                Category = category,
                PriorPeriodLabel = priorLabel,
                PriorAmount = prior?.Amounts.GetValueOrDefault(category) ?? 0m,
                CurrentPeriodLabel = currentLabel,
                CurrentAmount = current.Amounts.GetValueOrDefault(category),
                SourceReference = $"{Path.GetFileName(current.SourcePath)} ({current.SourceReference}); Prior Excel history: {Path.GetFileName(prior?.SourcePath ?? "")} ({prior?.SourceReference ?? "not found"})"
            });
            added++;
        }

        RecordClaimsTrendValidation(result, area, current, prior, priorLabel, currentLabel);
        return added;
    }

    private static int MergeParamedicalRollupFromEhcHistory(
        List<ServiceCategoryTrend> paramedicalRows,
        IReadOnlyList<ServiceCategorySnapshot> ehcSnapshots,
        WorkbookExtractionResult result)
    {
        if (paramedicalRows.Count == 0 || ehcSnapshots.Count == 0)
        {
            return 0;
        }

        var currentSnapshot = BuildSnapshotFromCurrentRows(paramedicalRows);
        var priorSnapshot = SelectSupportingHistorySnapshot(ehcSnapshots, currentSnapshot.EndDate);
        if (priorSnapshot is null ||
            !TryGetSnapshotAmount(priorSnapshot, "Paramedical", out var priorParamedicalAmount) ||
            priorParamedicalAmount == 0m)
        {
            return 0;
        }

        var rollup = paramedicalRows.FirstOrDefault(row =>
            NormalizeTrendCategoryKey(row.Category).Equals(NormalizeTrendCategoryKey("Paramedical Practitioners"), StringComparison.OrdinalIgnoreCase) ||
            NormalizeTrendCategoryKey(row.Category).Equals(NormalizeTrendCategoryKey("Paramedical Total"), StringComparison.OrdinalIgnoreCase));
        if (rollup is null)
        {
            rollup = new ServiceCategoryTrend
            {
                Benefit = "Extended Health Care",
                Category = "Paramedical Practitioners",
                CurrentPeriodLabel = currentSnapshot.PeriodLabel,
                CurrentAmount = paramedicalRows.Sum(row => row.CurrentAmount),
                SourceReference = "Paramedical rollup"
            };
            paramedicalRows.Add(rollup);
        }

        rollup.PriorPeriodLabel = string.IsNullOrWhiteSpace(priorSnapshot.PeriodLabel) ? "Prior Period" : priorSnapshot.PeriodLabel;
        rollup.PriorAmount = priorParamedicalAmount;
        rollup.SourceReference = AppendSourceReference(
            rollup.SourceReference,
            $"Prior Excel history - {Path.GetFileName(priorSnapshot.SourcePath)} ({priorSnapshot.SourceReference})");
        RecordClaimsTrendValidation(result, "ClaimsAnalysis_Paramedical", currentSnapshot, priorSnapshot, rollup.PriorPeriodLabel, currentSnapshot.PeriodLabel);
        return 1;
    }

    private static ServiceCategorySnapshot BuildSnapshotFromCurrentRows(IReadOnlyList<ServiceCategoryTrend> rows)
    {
        var periodLabel = rows.FirstOrDefault(row => !string.IsNullOrWhiteSpace(row.CurrentPeriodLabel))?.CurrentPeriodLabel ?? "Current Period";
        var dates = TryParsePeriodDatesFromLabel(periodLabel);
        var amounts = rows
            .GroupBy(row => row.Category, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(
                group => group.First().Category,
                group => group.Sum(row => row.CurrentAmount),
                StringComparer.OrdinalIgnoreCase);
        var sourceReference = string.Join("; ", rows
            .Select(row => row.SourceReference)
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Distinct(StringComparer.OrdinalIgnoreCase));

        return new ServiceCategorySnapshot(
            periodLabel,
            dates.StartDate,
            dates.EndDate,
            amounts,
            "PDF-derived claims detail",
            sourceReference);
    }

    private static ServiceCategorySnapshot? SelectSupportingHistorySnapshot(
        IReadOnlyList<ServiceCategorySnapshot> snapshots,
        DateOnly? currentEndDate)
    {
        if (currentEndDate.HasValue)
        {
            return snapshots
                .Where(snapshot => snapshot.EndDate.HasValue && snapshot.EndDate.Value < currentEndDate.Value)
                .OrderByDescending(snapshot => snapshot.EndDate)
                .FirstOrDefault();
        }

        return snapshots
            .OrderByDescending(snapshot => snapshot.EndDate ?? DateOnly.MinValue)
            .FirstOrDefault();
    }

    private static bool TryGetSnapshotAmount(ServiceCategorySnapshot snapshot, string category, out decimal amount)
    {
        if (snapshot.Amounts.TryGetValue(category, out amount))
        {
            return true;
        }

        var normalizedCategory = NormalizeTrendCategoryKey(category);
        foreach (var pair in snapshot.Amounts)
        {
            if (NormalizeTrendCategoryKey(pair.Key).Equals(normalizedCategory, StringComparison.OrdinalIgnoreCase))
            {
                amount = pair.Value;
                return true;
            }
        }

        amount = 0m;
        return false;
    }

    private static string NormalizeTrendCategoryKey(string category)
    {
        if (category.Contains("Prevent", StringComparison.OrdinalIgnoreCase) ||
            category.Contains("Diagnostic", StringComparison.OrdinalIgnoreCase))
        {
            return NormalizeClientMatchText("Basic");
        }

        if (category.Contains("Periodontal", StringComparison.OrdinalIgnoreCase))
        {
            return NormalizeClientMatchText("Periodontic");
        }

        if (category.Contains("Restorative", StringComparison.OrdinalIgnoreCase))
        {
            return NormalizeClientMatchText("Major");
        }

        return NormalizeClientMatchText(NormalizeParamedicalServiceName(category));
    }

    private static (DateOnly? StartDate, DateOnly? EndDate) TryParsePeriodDatesFromLabel(string label)
    {
        var named = TryParseNamedReportingPeriod(label);
        if (named.StartDate.HasValue || named.EndDate.HasValue)
        {
            return (named.StartDate, named.EndDate);
        }

        var numeric = TryParseExperiencePeriod(label);
        if (numeric.StartDate.HasValue || numeric.EndDate.HasValue)
        {
            return (numeric.StartDate, numeric.EndDate);
        }

        var fileNameStyle = TryParsePeriodFromFileName(label);
        if (fileNameStyle.StartDate.HasValue || fileNameStyle.EndDate.HasValue)
        {
            return (fileNameStyle.StartDate, fileNameStyle.EndDate);
        }

        var match = Regex.Match(
            label,
            @"(?<start>\d{1,2}[-/][A-Za-z]{3,9}[-/]\d{4}|[A-Za-z]{3,9}\s+\d{1,2},?\s+\d{4})\s*(?:to|-|â€“|â€”)\s*(?<end>\d{1,2}[-/][A-Za-z]{3,9}[-/]\d{4}|[A-Za-z]{3,9}\s+\d{1,2},?\s+\d{4})",
            RegexOptions.IgnoreCase);
        if (!match.Success ||
            !DateTime.TryParse(match.Groups["start"].Value, CultureInfo.InvariantCulture, DateTimeStyles.None, out var start) ||
            !DateTime.TryParse(match.Groups["end"].Value, CultureInfo.InvariantCulture, DateTimeStyles.None, out var end))
        {
            return (null, null);
        }

        return (DateOnly.FromDateTime(start), DateOnly.FromDateTime(end));
    }

    private static string AppendSourceReference(string current, string addition)
    {
        if (string.IsNullOrWhiteSpace(addition))
        {
            return current;
        }

        if (string.IsNullOrWhiteSpace(current))
        {
            return addition;
        }

        return current.Contains(addition, StringComparison.OrdinalIgnoreCase)
            ? current
            : $"{current}; {addition}";
    }

    private static ServiceCategorySnapshot? SelectPriorDentalSnapshot(
        string workbookPath,
        IReadOnlyList<ServiceCategorySnapshot> externalSnapshots,
        ServiceCategorySnapshot current)
    {
        var currentEnd = current.EndDate ?? DateOnly.MaxValue;
        var priorExternal = externalSnapshots
            .Where(snapshot => snapshot.EndDate.HasValue && snapshot.EndDate.Value < currentEnd)
            .OrderByDescending(snapshot => snapshot.EndDate)
            .FirstOrDefault();
        if (priorExternal is not null)
        {
            return priorExternal;
        }

        return TryReadPriorDentalRawSnapshot(workbookPath, current);
    }

    private static ServiceCategorySnapshot? TryReadPriorDentalRawSnapshot(string workbookPath, ServiceCategorySnapshot current)
    {
        var folder = Path.GetDirectoryName(workbookPath);
        if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
        {
            return null;
        }

        var currentFullPath = Path.GetFullPath(workbookPath);
        var currentPrefix = GetClientWorkbookPrefix(workbookPath);
        var snapshots = new List<ServiceCategorySnapshot>();

        foreach (var candidate in Directory.EnumerateFiles(folder, "*.xlsm", SearchOption.TopDirectoryOnly)
                     .Where(path => !string.Equals(Path.GetFullPath(path), currentFullPath, StringComparison.OrdinalIgnoreCase))
                     .Where(path => string.IsNullOrWhiteSpace(currentPrefix) ||
                                    Path.GetFileName(path).Contains(currentPrefix, StringComparison.OrdinalIgnoreCase)))
        {
            try
            {
                using var priorWorkbook = OpenWorkbookSharedRead(candidate);
                var snapshot = TryReadDentalRawSpendingSnapshot(priorWorkbook);
                if (snapshot is not null)
                {
                    snapshots.Add(snapshot);
                }
            }
            catch
            {
                // Ignore nearby workbooks that are not readable Manulife renewal workbooks.
            }
        }

        var currentEnd = current.EndDate ?? DateOnly.MaxValue;
        return snapshots
            .Where(snapshot => snapshot.EndDate.HasValue && snapshot.EndDate.Value < currentEnd)
            .OrderByDescending(snapshot => snapshot.EndDate)
            .FirstOrDefault();
    }

    private static XLWorkbook OpenWorkbookSharedRead(string path)
    {
        using var input = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        var memory = new MemoryStream();
        input.CopyTo(memory);
        memory.Position = 0;
        return new XLWorkbook(memory);
    }

    private static ServiceCategorySnapshot? TryReadEhcSummarySnapshot(XLWorkbook workbook, string sourcePath)
    {
        foreach (var sheet in WorksheetsWithPreferredFirst(workbook, "Summary"))
        {
            var snapshot = TryReadEhcSummarySheet(sheet, sourcePath);
            if (snapshot is not null)
            {
                return snapshot;
            }
        }

        return null;
    }

    private static ServiceCategorySnapshot? TryReadEhcSummarySheet(IXLWorksheet sheet, string sourcePath)
    {
        var headerRow = FindRow(sheet, "Total amount paid") ?? FindRow(sheet, "Total amount paid ($)");
        if (headerRow is null || HasMonthlyServiceColumns(sheet, headerRow.Value))
        {
            return null;
        }

        var benefitColumn = FindColumn(sheet, headerRow.Value, "Benefit") ?? 1;
        var categoryColumn = FindColumn(sheet, headerRow.Value, "Benefit category") ?? 2;
        var serviceColumn = FindColumn(sheet, headerRow.Value, "Service") ?? 3;
        var paidColumn = FindColumn(sheet, headerRow.Value, "Total amount paid") ?? 7;
        var lastRow = sheet.LastRowUsed()?.RowNumber() ?? headerRow.Value;
        var amounts = EhcSpendingCategories.ToDictionary(
            category => category,
            _ => 0m,
            StringComparer.OrdinalIgnoreCase);

        var currentBenefit = "";
        var currentCategory = "";
        var foundHealth = false;

        for (var row = headerRow.Value + 1; row <= lastRow; row++)
        {
            var benefit = Text(sheet.Cell(row, benefitColumn));
            if (!string.IsNullOrWhiteSpace(benefit))
            {
                currentBenefit = benefit;
            }

            if (!currentBenefit.Equals("Health", StringComparison.OrdinalIgnoreCase))
            {
                if (foundHealth)
                {
                    break;
                }

                continue;
            }

            foundHealth = true;
            var rawCategory = Text(sheet.Cell(row, categoryColumn));
            if (!string.IsNullOrWhiteSpace(rawCategory))
            {
                currentCategory = rawCategory;
            }

            var mappedCategory = MapEhcSummaryCategory(currentCategory);
            if (mappedCategory is null)
            {
                continue;
            }

            var service = Text(sheet.Cell(row, serviceColumn));
            if (!service.Equals("Subtotal", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            amounts[mappedCategory] += Number(sheet.Cell(row, paidColumn)) ?? 0m;
        }

        if (!foundHealth)
        {
            return null;
        }

        amounts["Total"] = amounts
            .Where(pair => !pair.Key.Equals("Total", StringComparison.OrdinalIgnoreCase))
            .Sum(pair => pair.Value);

        var (periodLabel, startDate, endDate) = ReadExperiencePeriodLabel(sheet, sourcePath);
        return new ServiceCategorySnapshot(
            periodLabel,
            startDate,
            endDate,
            amounts,
            sourcePath,
            $"{sheet.Name}!A{headerRow.Value}:{ColumnName(sheet.LastColumnUsed()?.ColumnNumber() ?? paidColumn)}{lastRow}");
    }

    private static ServiceCategorySnapshot? TryReadDentalSummarySnapshot(XLWorkbook workbook, string sourcePath)
    {
        foreach (var sheet in WorksheetsWithPreferredFirst(workbook, "Summary"))
        {
            var snapshot = TryReadDentalSummarySheet(sheet, sourcePath);
            if (snapshot is not null)
            {
                return snapshot;
            }
        }

        return null;
    }

    private static ServiceCategorySnapshot? TryReadDentalSummarySheet(IXLWorksheet sheet, string sourcePath)
    {
        var headerRow = FindRow(sheet, "Total amount paid") ?? FindRow(sheet, "Total amount paid ($)");
        if (headerRow is null || HasMonthlyServiceColumns(sheet, headerRow.Value))
        {
            return null;
        }

        var benefitColumn = FindColumn(sheet, headerRow.Value, "Benefit") ?? 1;
        var categoryColumn = FindColumn(sheet, headerRow.Value, "Benefit category") ?? 2;
        var serviceColumn = FindColumn(sheet, headerRow.Value, "Service") ?? 3;
        var paidColumn = FindColumn(sheet, headerRow.Value, "Total amount paid") ?? 7;
        var lastRow = sheet.LastRowUsed()?.RowNumber() ?? headerRow.Value;
        var amounts = DentalSpendingCategories.ToDictionary(
            category => category,
            _ => 0m,
            StringComparer.OrdinalIgnoreCase);

        var currentBenefit = "";
        var currentCategory = "";
        var foundDental = false;

        for (var row = headerRow.Value + 1; row <= lastRow; row++)
        {
            var benefit = Text(sheet.Cell(row, benefitColumn));
            if (!string.IsNullOrWhiteSpace(benefit))
            {
                currentBenefit = benefit;
            }

            if (!currentBenefit.Equals("Dental", StringComparison.OrdinalIgnoreCase))
            {
                if (foundDental)
                {
                    break;
                }

                continue;
            }

            foundDental = true;
            var rawCategory = Text(sheet.Cell(row, categoryColumn));
            var mappedCategory = MapDentalSpendingCategory(rawCategory);
            if (mappedCategory is not null)
            {
                currentCategory = mappedCategory;
            }

            var service = Text(sheet.Cell(row, serviceColumn));
            if (!service.Equals("Subtotal", StringComparison.OrdinalIgnoreCase) ||
                string.IsNullOrWhiteSpace(currentCategory))
            {
                continue;
            }

            amounts[currentCategory] += Number(sheet.Cell(row, paidColumn)) ?? 0m;
        }

        if (!foundDental)
        {
            return null;
        }

        amounts["Total"] = amounts
            .Where(pair => !pair.Key.Equals("Total", StringComparison.OrdinalIgnoreCase))
            .Sum(pair => pair.Value);

        var (periodLabel, startDate, endDate) = ReadExperiencePeriodLabel(sheet, sourcePath);
        return new ServiceCategorySnapshot(
            periodLabel,
            startDate,
            endDate,
            amounts,
            sourcePath,
            $"{sheet.Name}!A{headerRow.Value}:{ColumnName(sheet.LastColumnUsed()?.ColumnNumber() ?? paidColumn)}{lastRow}");
    }

    private static ServiceCategorySnapshot? TryReadDentalRawSpendingSnapshot(XLWorkbook workbook, string? sourcePath = null)
    {
        var sheet = GetDentalClaimsAnalysisSheet(workbook);
        if (sheet is null)
        {
            return null;
        }

        var headerRow = FindRow(sheet, "Member/ Dependents") ??
                        FindRow(sheet, "Basic") ??
                        10;
        var memberColumn = FindColumn(sheet, headerRow, "Member/ Dependents") ?? 3;
        var lastRow = sheet.LastRowUsed()?.RowNumber() ?? headerRow;
        var lastColumn = sheet.LastColumnUsed()?.ColumnNumber() ?? 0;
        var categoryColumns = new Dictionary<int, string>();

        for (var column = 1; column <= lastColumn; column++)
        {
            var category = MapDentalSpendingCategory(Text(sheet.Cell(headerRow, column)));
            if (category is not null)
            {
                categoryColumns[column] = category;
            }
        }

        if (categoryColumns.Count == 0)
        {
            return null;
        }

        var amounts = DentalSpendingCategories.ToDictionary(
            category => category,
            _ => 0m,
            StringComparer.OrdinalIgnoreCase);

        var totalRow = FindRawServiceCategoryTotalRow(sheet, headerRow, lastRow, lastColumn);
        if (totalRow.HasValue)
        {
            AddServiceCategoryRow(sheet, totalRow.Value, categoryColumns, amounts);
        }
        else
        {
            var totalCertificateRowsFound = false;
            for (var row = headerRow + 1; row <= lastRow; row++)
            {
                var memberType = Text(sheet.Cell(row, memberColumn));
                if (!memberType.Equals("Total Certificate", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                totalCertificateRowsFound = true;
                AddServiceCategoryRow(sheet, row, categoryColumns, amounts);
            }

            if (!totalCertificateRowsFound)
            {
                return null;
            }
        }

        amounts["Total"] = categoryColumns
            .Where(pair => pair.Value.Equals("Total", StringComparison.OrdinalIgnoreCase))
            .Select(pair => Number(sheet.Cell(totalRow ?? lastRow, pair.Key)) ?? 0m)
            .FirstOrDefault();
        if (amounts["Total"] == 0m)
        {
            amounts["Total"] = amounts
                .Where(pair => !pair.Key.Equals("Total", StringComparison.OrdinalIgnoreCase))
                .Sum(pair => pair.Value);
        }

        var (periodLabel, startDate, endDate) = string.IsNullOrWhiteSpace(sourcePath)
            ? ReadExperiencePeriodLabel(sheet)
            : ReadExperiencePeriodLabel(sheet, sourcePath);
        return new ServiceCategorySnapshot(
            periodLabel,
            startDate,
            endDate,
            amounts,
            string.IsNullOrWhiteSpace(sourcePath) ? "Manulife renewal workbook" : sourcePath,
            $"{sheet.Name}!A{headerRow}:{ColumnName(lastColumn)}{lastRow}");
    }

    private static ServiceCategorySnapshot? SelectPriorEhcSnapshot(
        string workbookPath,
        IReadOnlyList<ServiceCategorySnapshot> externalSnapshots,
        ServiceCategorySnapshot current)
    {
        var currentEnd = current.EndDate ?? DateOnly.MaxValue;
        var priorExternal = externalSnapshots
            .Where(snapshot => snapshot.EndDate.HasValue && snapshot.EndDate.Value < currentEnd)
            .OrderByDescending(snapshot => snapshot.EndDate)
            .FirstOrDefault();
        if (priorExternal is not null)
        {
            return priorExternal;
        }

        return TryReadPriorEhcSpendingSnapshot(workbookPath, current);
    }

    private static ServiceCategorySnapshot? TryReadPriorEhcSpendingSnapshot(string workbookPath, ServiceCategorySnapshot current)
    {
        var folder = Path.GetDirectoryName(workbookPath);
        if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
        {
            return null;
        }

        var currentFullPath = Path.GetFullPath(workbookPath);
        var currentPrefix = GetClientWorkbookPrefix(workbookPath);
        var candidates = Directory.EnumerateFiles(folder, "*.xlsm", SearchOption.TopDirectoryOnly)
            .Where(path => !string.Equals(Path.GetFullPath(path), currentFullPath, StringComparison.OrdinalIgnoreCase))
            .Where(path => string.IsNullOrWhiteSpace(currentPrefix) ||
                           Path.GetFileName(path).Contains(currentPrefix, StringComparison.OrdinalIgnoreCase))
            .ToList();

        var snapshots = new List<ServiceCategorySnapshot>();
        foreach (var candidate in candidates)
        {
            try
            {
                using var workbook = OpenWorkbookSharedRead(candidate);
                var snapshot = TryReadEhcSpendingSnapshot(workbook);
                if (snapshot is not null)
                {
                    snapshots.Add(snapshot);
                }
            }
            catch
            {
                // Ignore nearby workbooks that are open, locked, or not Manulife claim workbooks.
            }
        }

        if (snapshots.Count == 0)
        {
            return null;
        }

        var currentEnd = current.EndDate ?? DateOnly.MaxValue;
        return snapshots
            .Where(snapshot => snapshot.EndDate.HasValue && snapshot.EndDate.Value < currentEnd)
            .OrderByDescending(snapshot => snapshot.EndDate)
            .FirstOrDefault()
            ?? snapshots.OrderByDescending(snapshot => snapshot.EndDate ?? DateOnly.MinValue).FirstOrDefault();
    }

    private static string GetClientWorkbookPrefix(string workbookPath)
    {
        var fileName = Path.GetFileNameWithoutExtension(workbookPath);
        var match = Regex.Match(fileName, @"^(?<prefix>.+?)(?:[_\s-]*Client[_\s-]*Copy|$)", RegexOptions.IgnoreCase);
        return match.Success ? match.Groups["prefix"].Value.Trim(' ', '_', '-') : "";
    }

    private static ServiceCategorySnapshot? TryReadEhcSpendingSnapshot(XLWorkbook workbook, string? sourcePath = null)
    {
        var sheet = GetEhcClaimsAnalysisSheet(workbook);
        if (sheet is null)
        {
            return null;
        }

        var headerRow = FindRow(sheet, "Member/ Dependents") ??
                        FindRow(sheet, "Medical Services and Supplies") ??
                        FindRow(sheet, "Paramedical") ??
                        10;
        var memberColumn = FindColumn(sheet, headerRow, "Member/ Dependents") ?? 3;
        var lastRow = sheet.LastRowUsed()?.RowNumber() ?? headerRow;
        var lastColumn = sheet.LastColumnUsed()?.ColumnNumber() ?? 0;
        var categoryColumns = new Dictionary<int, string>();

        for (var column = 1; column <= lastColumn; column++)
        {
            var category = MapEhcSpendingCategory(Text(sheet.Cell(headerRow, column)));
            if (category is not null)
            {
                categoryColumns[column] = category;
            }
        }

        if (categoryColumns.Count == 0)
        {
            return null;
        }

        var amounts = EhcSpendingCategories.ToDictionary(
            category => category,
            _ => 0m,
            StringComparer.OrdinalIgnoreCase);

        var totalRow = FindRawServiceCategoryTotalRow(sheet, headerRow, lastRow, lastColumn);
        if (totalRow.HasValue)
        {
            AddEhcSpendingRow(sheet, totalRow.Value, categoryColumns, amounts);
        }
        else
        {
            var totalCertificateRowsFound = false;
            for (var row = headerRow + 1; row <= lastRow; row++)
            {
                var memberType = Text(sheet.Cell(row, memberColumn));
                if (!memberType.Equals("Total Certificate", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                totalCertificateRowsFound = true;
                AddEhcSpendingRow(sheet, row, categoryColumns, amounts);
            }

            if (!totalCertificateRowsFound)
            {
                return null;
            }
        }

        amounts["Total"] = categoryColumns
            .Where(pair => pair.Value.Equals("Total", StringComparison.OrdinalIgnoreCase))
            .Select(pair => Number(sheet.Cell(totalRow ?? lastRow, pair.Key)) ?? 0m)
            .FirstOrDefault();
        if (amounts["Total"] == 0m)
        {
            amounts["Total"] = amounts
                .Where(pair => !pair.Key.Equals("Total", StringComparison.OrdinalIgnoreCase))
                .Sum(pair => pair.Value);
        }

        var (periodLabel, startDate, endDate) = string.IsNullOrWhiteSpace(sourcePath)
            ? ReadExperiencePeriodLabel(sheet)
            : ReadExperiencePeriodLabel(sheet, sourcePath);
        return new ServiceCategorySnapshot(
            periodLabel,
            startDate,
            endDate,
            amounts,
            string.IsNullOrWhiteSpace(sourcePath) ? "Manulife renewal workbook" : sourcePath,
            $"{sheet.Name}!A{headerRow}:{ColumnName(lastColumn)}{lastRow}");
    }

    private static IXLWorksheet? GetEhcClaimsAnalysisSheet(XLWorkbook workbook)
    {
        if (workbook.TryGetWorksheet("ClaimsAnalysis_EHC", out var sheet))
        {
            return sheet;
        }

        return workbook.Worksheets.FirstOrDefault(worksheet =>
            worksheet.Name.Contains("EHC", StringComparison.OrdinalIgnoreCase) &&
            WorksheetContains(worksheet, "Group Claims Paid By Service Category - EHC"))
            ?? workbook.Worksheets.FirstOrDefault(worksheet =>
                WorksheetContains(worksheet, "Group Claims Paid By Service Category") &&
                WorksheetContains(worksheet, "Paramedical") &&
                WorksheetContains(worksheet, "Drugs") &&
                WorksheetContains(worksheet, "Vision"));
    }

    private static IXLWorksheet? GetDentalClaimsAnalysisSheet(XLWorkbook workbook)
    {
        if (workbook.TryGetWorksheet("ClaimsAnalysis_Dental", out var sheet))
        {
            return sheet;
        }

        return workbook.Worksheets.FirstOrDefault(worksheet =>
            worksheet.Name.Contains("Dental", StringComparison.OrdinalIgnoreCase) &&
            (WorksheetContains(worksheet, "Group Claims Paid By Service Category - Dental") ||
             WorksheetContains(worksheet, "Group Claims Paid By Service Category -Dental")))
            ?? workbook.Worksheets.FirstOrDefault(worksheet =>
                WorksheetContains(worksheet, "Group Claims Paid By Service Category") &&
                WorksheetContains(worksheet, "Basic") &&
                WorksheetContains(worksheet, "Major") &&
                WorksheetContains(worksheet, "Periodontic"));
    }

    private static bool WorksheetContains(IXLWorksheet sheet, string text)
    {
        var lastRow = Math.Min(sheet.LastRowUsed()?.RowNumber() ?? 0, 20);
        var lastColumn = Math.Min(sheet.LastColumnUsed()?.ColumnNumber() ?? 0, 15);
        for (var row = 1; row <= lastRow; row++)
        {
            for (var column = 1; column <= lastColumn; column++)
            {
                if (Text(sheet.Cell(row, column)).Contains(text, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
        }

        return false;
    }

    private static void AddEhcSpendingRow(
        IXLWorksheet sheet,
        int row,
        IReadOnlyDictionary<int, string> categoryColumns,
        Dictionary<string, decimal> amounts)
    {
        AddServiceCategoryRow(sheet, row, categoryColumns, amounts);
    }

    private static void AddServiceCategoryRow(
        IXLWorksheet sheet,
        int row,
        IReadOnlyDictionary<int, string> categoryColumns,
        Dictionary<string, decimal> amounts)
    {
        foreach (var pair in categoryColumns)
        {
            amounts[pair.Value] += Number(sheet.Cell(row, pair.Key)) ?? 0m;
        }
    }

    private static int? FindRawServiceCategoryTotalRow(IXLWorksheet sheet, int headerRow, int lastRow, int lastColumn)
    {
        for (var row = lastRow; row > headerRow; row--)
        {
            for (var column = 1; column <= Math.Min(lastColumn, 4); column++)
            {
                if (Text(sheet.Cell(row, column)).Equals("Total", StringComparison.OrdinalIgnoreCase))
                {
                    return row;
                }
            }
        }

        return null;
    }

    private static string? MapEhcSpendingCategory(string header)
    {
        if (string.IsNullOrWhiteSpace(header))
        {
            return null;
        }

        if (header.Contains("Medical Services", StringComparison.OrdinalIgnoreCase) ||
            header.Contains("Medical Service", StringComparison.OrdinalIgnoreCase) ||
            header.Contains("Supplies", StringComparison.OrdinalIgnoreCase))
        {
            return "Medical Services and Supplies";
        }

        if (header.Contains("Paramedical", StringComparison.OrdinalIgnoreCase))
        {
            return "Paramedical";
        }

        if (header.Contains("Drugs", StringComparison.OrdinalIgnoreCase))
        {
            return "Drugs";
        }

        if (header.Contains("Vision", StringComparison.OrdinalIgnoreCase))
        {
            return "Vision";
        }

        if (header.Equals("Pooled", StringComparison.OrdinalIgnoreCase) ||
            Regex.IsMatch(header, @"\bPooled\b", RegexOptions.IgnoreCase))
        {
            return "Pooled";
        }

        if (header.Contains("Miscellaneous", StringComparison.OrdinalIgnoreCase))
        {
            return "Miscellaneous Services";
        }

        if (header.Contains("Hospital", StringComparison.OrdinalIgnoreCase))
        {
            return "Hospital";
        }

        if (header.Equals("Other", StringComparison.OrdinalIgnoreCase))
        {
            return "Hospital";
        }

        if (header.Equals("Total", StringComparison.OrdinalIgnoreCase))
        {
            return "Total";
        }

        return null;
    }

    private static string? MapEhcSummaryCategory(string category)
    {
        if (string.IsNullOrWhiteSpace(category))
        {
            return null;
        }

        if (category.Contains("Pharmacy", StringComparison.OrdinalIgnoreCase) ||
            category.Contains("Drug", StringComparison.OrdinalIgnoreCase))
        {
            return "Drugs";
        }

        if (category.Contains("Vision", StringComparison.OrdinalIgnoreCase))
        {
            return "Vision";
        }

        if (category.Contains("Paramedical", StringComparison.OrdinalIgnoreCase))
        {
            return "Paramedical";
        }

        if (category.Contains("Medical Services", StringComparison.OrdinalIgnoreCase) ||
            category.Contains("Supplies", StringComparison.OrdinalIgnoreCase))
        {
            return "Medical Services and Supplies";
        }

        if (category.Contains("Pooled", StringComparison.OrdinalIgnoreCase))
        {
            return "Pooled";
        }

        if (category.Contains("Miscellaneous", StringComparison.OrdinalIgnoreCase))
        {
            return "Miscellaneous Services";
        }

        if (category.Contains("Hospital", StringComparison.OrdinalIgnoreCase))
        {
            return "Hospital";
        }

        if (category.Equals("Other", StringComparison.OrdinalIgnoreCase))
        {
            return "Hospital";
        }

        if (category.Equals("Total", StringComparison.OrdinalIgnoreCase))
        {
            return "Total";
        }

        return null;
    }

    private static string? MapDentalSpendingCategory(string header)
    {
        if (string.IsNullOrWhiteSpace(header))
        {
            return null;
        }

        var normalized = header.Trim();
        if (normalized.Equals("Basic", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("Basic Dental", StringComparison.OrdinalIgnoreCase))
        {
            return "Basic";
        }

        if (normalized.Equals("Major", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("Major Dental", StringComparison.OrdinalIgnoreCase))
        {
            return "Major";
        }

        if (normalized.Contains("Denture", StringComparison.OrdinalIgnoreCase))
        {
            return "Dentures";
        }

        if (normalized.Contains("Endodontic", StringComparison.OrdinalIgnoreCase))
        {
            return "Endodontics";
        }

        if (normalized.Contains("Orthodont", StringComparison.OrdinalIgnoreCase))
        {
            return "Orthodontic";
        }

        if (normalized.Contains("Periodont", StringComparison.OrdinalIgnoreCase))
        {
            return "Periodontic";
        }

        if (normalized.Contains("Miscellaneous", StringComparison.OrdinalIgnoreCase))
        {
            return "Miscellaneous Services";
        }

        if (normalized.Equals("Total", StringComparison.OrdinalIgnoreCase))
        {
            return "Total";
        }

        return null;
    }

    private static void LogEhcSpendingTrends(
        IReadOnlyList<ServiceCategoryTrend> rows,
        ServiceCategorySnapshot current,
        ServiceCategorySnapshot? prior)
    {
        var values = string.Join(
            "; ",
            rows.Select(row => $"{row.Category}: prior={row.PriorAmount:0.##}, current={row.CurrentAmount:0.##}"));
        AppLogger.Info(
            $"Extended Healthcare Spending Trends extraction | priorPeriod={rows.FirstOrDefault()?.PriorPeriodLabel ?? "Prior Period"} | " +
            $"currentPeriod={rows.FirstOrDefault()?.CurrentPeriodLabel ?? "Current Period"} | " +
            $"currentSource={current.SourcePath} ({current.SourceReference}) | " +
            $"priorSource={prior?.SourcePath ?? "not found"} ({prior?.SourceReference ?? "not found"}) | " +
            values);
    }

    private static void RecordClaimsTrendValidation(
        WorkbookExtractionResult result,
        string area,
        ServiceCategorySnapshot current,
        ServiceCategorySnapshot? prior,
        string priorPeriod,
        string currentPeriod)
    {
        var warnings = new List<string>();
        if (string.IsNullOrWhiteSpace(current.ClientName) && string.IsNullOrWhiteSpace(prior?.ClientName))
        {
            warnings.Add("Client name was not readable from one or more claims workbooks.");
        }

        if (string.IsNullOrWhiteSpace(current.PolicyNumber) && string.IsNullOrWhiteSpace(prior?.PolicyNumber))
        {
            warnings.Add("Policy number was not readable from one or more claims workbooks.");
        }

        if (prior is null)
        {
            warnings.Add("Prior-period source was not found.");
        }

        result.ClaimsTrendValidations.Add(new ClaimsTrendValidation
        {
            Area = area,
            ClientName = FirstNonBlank(current.ClientName, prior?.ClientName),
            PolicyNumber = FirstNonBlank(current.PolicyNumber, prior?.PolicyNumber),
            PriorPeriod = priorPeriod,
            CurrentPeriod = currentPeriod,
            SourceSheetUsed = $"Current: {Path.GetFileName(current.SourcePath)} ({current.SourceReference}); Prior: {(prior is null ? "not found" : $"{Path.GetFileName(prior.SourcePath)} ({prior.SourceReference})")}",
            Status = warnings.Count == 0 ? "Matched" : "Warning",
            Warning = string.Join(" ", warnings)
        });
    }

    private static string FirstNonBlank(params string?[] values)
    {
        return values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)) ?? "";
    }

    private static void LogDentalSpendingTrends(
        IReadOnlyList<ServiceCategoryTrend> rows,
        ServiceCategorySnapshot current,
        ServiceCategorySnapshot? prior)
    {
        var values = string.Join(
            "; ",
            rows.Select(row => $"{row.Category}: prior={row.PriorAmount:0.##}, current={row.CurrentAmount:0.##}"));
        AppLogger.Info(
            $"Dental Spending Trends extraction | priorPeriod={rows.FirstOrDefault()?.PriorPeriodLabel ?? "Prior Period"} | " +
            $"currentPeriod={rows.FirstOrDefault()?.CurrentPeriodLabel ?? "Current Period"} | " +
            $"currentSource={current.SourcePath} ({current.SourceReference}) | " +
            $"priorSource={prior?.SourcePath ?? "not found"} ({prior?.SourceReference ?? "not found"}) | " +
            values);
    }

    private static void LogParamedicalServiceTrends(
        IReadOnlyList<ServiceCategoryTrend> rows,
        ServiceCategorySnapshot current,
        ServiceCategorySnapshot? prior)
    {
        var values = string.Join(
            "; ",
            rows.Select(row => $"{row.Category}: prior={row.PriorAmount:0.##}, current={row.CurrentAmount:0.##}"));
        AppLogger.Info(
            $"Paramedical Services Utilization Trends extraction | priorPeriod={rows.FirstOrDefault()?.PriorPeriodLabel ?? "Prior Period"} | " +
            $"currentPeriod={rows.FirstOrDefault()?.CurrentPeriodLabel ?? "Current Period"} | " +
            $"currentSource={current.SourcePath} ({current.SourceReference}) | " +
            $"priorSource={prior?.SourcePath ?? "not found"} ({prior?.SourceReference ?? "not found"}) | " +
            values);
    }

    private static void LogClaimsTrendCaptureSummary(WorkbookExtractionResult result)
    {
        AppLogger.Info(
            "Claims trend capture summary | " +
            $"EHC rows={result.EhcSpendingTrends.Count} | " +
            $"Dental rows={result.DentalSpendingTrends.Count} | " +
            $"Paramedical rows={result.ParamedicalServiceTrends.Count} | " +
            $"EHC prior/current={TrendPeriodSummary(result.EhcSpendingTrends)} | " +
            $"Dental prior/current={TrendPeriodSummary(result.DentalSpendingTrends)} | " +
            $"Paramedical prior/current={TrendPeriodSummary(result.ParamedicalServiceTrends)} | " +
            $"Warnings={string.Join(" || ", result.Warnings)}");
    }

    private static string TrendPeriodSummary(IReadOnlyCollection<ServiceCategoryTrend> rows)
    {
        var first = rows.FirstOrDefault();
        return first is null
            ? "none"
            : $"{first.PriorPeriodLabel} / {first.CurrentPeriodLabel}";
    }

    private static (string Label, DateOnly? StartDate, DateOnly? EndDate) ReadExperiencePeriodLabel(IXLWorksheet sheet)
    {
        var lastRow = Math.Min(sheet.LastRowUsed()?.RowNumber() ?? 0, 20);
        var lastColumn = Math.Min(sheet.LastColumnUsed()?.ColumnNumber() ?? 0, 20);
        for (var row = 1; row <= lastRow; row++)
        {
            for (var column = 1; column <= lastColumn; column++)
            {
                var text = Text(sheet.Cell(row, column));
                var parsed = TryParseExperiencePeriod(text);
                if (string.IsNullOrWhiteSpace(parsed.Label))
                {
                    parsed = TryParseNamedReportingPeriod(text);
                }

                if (!string.IsNullOrWhiteSpace(parsed.Label))
                {
                    return parsed;
                }
            }
        }

        return ("Current Period", null, null);
    }

    private static (string Label, DateOnly? StartDate, DateOnly? EndDate) ReadExperiencePeriodLabel(IXLWorksheet sheet, string sourcePath)
    {
        var parsed = ReadExperiencePeriodLabel(sheet);
        if (parsed.StartDate.HasValue && parsed.EndDate.HasValue)
        {
            return parsed;
        }

        var fileNameParsed = TryParsePeriodFromFileName(Path.GetFileNameWithoutExtension(sourcePath));
        return !string.IsNullOrWhiteSpace(fileNameParsed.Label)
            ? fileNameParsed
            : parsed;
    }

    private static (string Label, DateOnly? StartDate, DateOnly? EndDate) TryParsePeriodFromFileName(string fileName)
    {
        var match = Regex.Match(
            fileName,
            @"(?<startMonth>[A-Za-z]{3,9})\s*(?<startYear>\d{2,4})\s*(?:to|-|–|—)\s*(?<endMonth>[A-Za-z]{3,9})\s*(?<endYear>\d{2,4})",
            RegexOptions.IgnoreCase);
        if (!match.Success)
        {
            return ("", null, null);
        }

        var startMonth = ParseMonthName(match.Groups["startMonth"].Value);
        var endMonth = ParseMonthName(match.Groups["endMonth"].Value);
        if (!startMonth.HasValue || !endMonth.HasValue)
        {
            return ("", null, null);
        }

        if (!int.TryParse(match.Groups["startYear"].Value, out var startYear) ||
            !int.TryParse(match.Groups["endYear"].Value, out var endYear))
        {
            return ("", null, null);
        }

        startYear = NormalizeYear(startYear);
        endYear = NormalizeYear(endYear);
        var start = new DateOnly(startYear, startMonth.Value, 1);
        var end = new DateOnly(endYear, endMonth.Value, 1);
        return (FormatPeriodLabel(start, end), start, end);
    }

    private static int NormalizeYear(int year)
    {
        return year < 100 ? 2000 + year : year;
    }

    private static (string Label, DateOnly? StartDate, DateOnly? EndDate) TryParseNamedReportingPeriod(string text)
    {
        var match = Regex.Match(
            text,
            @"(?<startMonth>[A-Za-z]+)\s+\d{1,2},\s*(?<startYear>\d{4})\s*(?:to|-)\s*(?<endMonth>[A-Za-z]+)\s+\d{1,2},\s*(?<endYear>\d{4})",
            RegexOptions.IgnoreCase);
        if (!match.Success)
        {
            return ("", null, null);
        }

        var startMonth = ParseMonthName(match.Groups["startMonth"].Value);
        var endMonth = ParseMonthName(match.Groups["endMonth"].Value);
        if (!startMonth.HasValue || !endMonth.HasValue)
        {
            return ("", null, null);
        }

        if (!int.TryParse(match.Groups["startYear"].Value, out var startYear) ||
            !int.TryParse(match.Groups["endYear"].Value, out var endYear))
        {
            return ("", null, null);
        }

        var start = new DateOnly(startYear, startMonth.Value, 1);
        var end = new DateOnly(endYear, endMonth.Value, 1);
        return (FormatPeriodLabel(start, end), start, end);
    }

    private static int? ParseMonthName(string month)
    {
        for (var index = 1; index <= 12; index++)
        {
            var date = new DateTime(2000, index, 1);
            if (date.ToString("MMMM", CultureInfo.InvariantCulture).Equals(month, StringComparison.OrdinalIgnoreCase) ||
                date.ToString("MMM", CultureInfo.InvariantCulture).Equals(month, StringComparison.OrdinalIgnoreCase))
            {
                return index;
            }
        }

        return null;
    }

    private static (string Label, DateOnly? StartDate, DateOnly? EndDate) TryParseExperiencePeriod(string text)
    {
        var match = Regex.Match(
            text,
            @"(?<startMonth>\d{1,2})\s*[-/]\s*(?<startYear>\d{4})\s*(?:to|-)\s*(?<endMonth>\d{1,2})\s*[-/]\s*(?<endYear>\d{4})",
            RegexOptions.IgnoreCase);
        if (!match.Success)
        {
            return ("", null, null);
        }

        if (!int.TryParse(match.Groups["startMonth"].Value, out var startMonth) ||
            !int.TryParse(match.Groups["startYear"].Value, out var startYear) ||
            !int.TryParse(match.Groups["endMonth"].Value, out var endMonth) ||
            !int.TryParse(match.Groups["endYear"].Value, out var endYear))
        {
            return ("", null, null);
        }

        var start = new DateOnly(startYear, startMonth, 1);
        var end = new DateOnly(endYear, endMonth, 1);
        return (FormatPeriodLabel(start, end), start, end);
    }

    private static string? BuildPriorPeriodLabel(DateOnly? startDate, DateOnly? endDate)
    {
        return startDate.HasValue && endDate.HasValue
            ? FormatPeriodLabel(startDate.Value.AddYears(-1), endDate.Value.AddYears(-1))
            : null;
    }

    private static string FormatPeriodLabel(DateOnly startDate, DateOnly endDate)
    {
        return $"{startDate.ToString("MMM yyyy", CultureInfo.InvariantCulture)} - {endDate.ToString("MMM yyyy", CultureInfo.InvariantCulture)}";
    }

    private static void ReadBenchmarks(XLWorkbook workbook, WorkbookExtractionResult result)
    {
        if (!workbook.TryGetWorksheet("Dent-Hlth Ttl Claims", out var sheet))
        {
            return;
        }

        var lastRow = sheet.LastRowUsed()?.RowNumber() ?? 0;
        for (var row = 1; row <= lastRow; row++)
        {
            if (!Text(sheet.Cell(row, 2)).Equals("Claim Categories", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var benefit = DetectBenchmarkBenefit(sheet, row);
            for (var dataRow = row + 1; dataRow <= Math.Min(row + 10, lastRow); dataRow++)
            {
                var category = Text(sheet.Cell(dataRow, 2));
                if (string.IsNullOrWhiteSpace(category) || category.Equals("Total", StringComparison.OrdinalIgnoreCase))
                {
                    break;
                }

                result.ClaimTrends.Add(new ClaimTrend
                {
                    Benefit = benefit,
                    Category = category,
                    Period = "Benchmark",
                    PercentOfTotal = ToPercent(Number(sheet.Cell(dataRow, 3))),
                    PriorPeriodPercent = ToPercent(Number(sheet.Cell(dataRow, 4))),
                    BenchmarkPercent = ToPercent(Number(sheet.Cell(dataRow, 5))),
                    RegionalIndustryBenchmarkPercent = ToPercent(Number(sheet.Cell(dataRow, 8))),
                    SourceReference = $"Dent-Hlth Ttl Claims!B{dataRow}:H{dataRow}"
                });
            }
        }
    }

    private static string DetectBenchmarkBenefit(IXLWorksheet sheet, int row)
    {
        for (var scan = row - 1; scan >= Math.Max(1, row - 8); scan--)
        {
            var text = string.Join(" ", Enumerable.Range(1, 8).Select(col => Text(sheet.Cell(scan, col))));
            if (text.Contains("Dental", StringComparison.OrdinalIgnoreCase))
            {
                return "Dental";
            }

            if (text.Contains("Health", StringComparison.OrdinalIgnoreCase))
            {
                return "Extended Health Care";
            }
        }

        return "Benefits";
    }

    private static void ReadTopDrugs(
        XLWorkbook workbook,
        string workbookPath,
        string? clientName,
        string? currentYearFolder,
        string? priorYearFolder,
        WorkbookExtractionResult result)
    {
        var currentFolderCandidate = SelectTopDrugWorkbookCandidate(
            EnumerateExcelFiles(currentYearFolder),
            workbookPath,
            clientName,
            currentYearFolder,
            priorYearFolder,
            "Current");
        if (currentFolderCandidate is not null)
        {
            result.DrugTrends.AddRange(currentFolderCandidate.Drugs);
            var periodText = string.IsNullOrWhiteSpace(currentFolderCandidate.PeriodLabel)
                ? "period not identified"
                : currentFolderCandidate.PeriodLabel;
            result.Warnings.Add($"Current Top DIN source: {Path.GetFileName(currentFolderCandidate.SourcePath)} ({periodText}).");
            return;
        }

        var selectedDrugs = ExtractTopDrugs(workbook, $"{Path.GetFileName(workbookPath)}:Top_50_DIN");
        result.DrugTrends.AddRange(selectedDrugs);
        if (selectedDrugs.Count > 0)
        {
            var period = ReadWorkbookPeriod(workbook, workbookPath);
            var periodText = string.IsNullOrWhiteSpace(period.Label)
                ? "period not identified"
                : period.Label;
            result.Warnings.Add($"Current Top DIN source: {Path.GetFileName(workbookPath)} ({periodText}).");
        }
    }

    private static TopDrugWorkbookCandidate? SelectTopDrugWorkbookCandidate(
        IEnumerable<string> candidatePaths,
        string workbookPath,
        string? clientName,
        string? currentYearFolder,
        string? priorYearFolder,
        string requiredRole)
    {
        var paths = candidatePaths
            .Where(path => !string.IsNullOrWhiteSpace(path) && File.Exists(path))
            .Where(path =>
                path.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                path.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase))
            .Where(path =>
                string.IsNullOrWhiteSpace(requiredRole) ||
                GetCandidatePeriodRole(path, currentYearFolder, priorYearFolder).Equals(requiredRole, StringComparison.OrdinalIgnoreCase))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        if (paths.Count == 0)
        {
            return null;
        }

        var prefixes = BuildClientFilePrefixes(clientName, workbookPath);
        var relevantPaths = SelectRelevantExplicitWorkbookCandidates(paths, prefixes);
        var candidates = new List<TopDrugWorkbookCandidate>();
        foreach (var candidatePath in relevantPaths)
        {
            try
            {
                using var candidateWorkbook = OpenWorkbookSharedRead(candidatePath);
                var drugs = ExtractTopDrugs(candidateWorkbook, $"{Path.GetFileName(candidatePath)}:Top_50_DIN");
                if (drugs.Count == 0)
                {
                    continue;
                }

                var period = ReadWorkbookPeriod(candidateWorkbook, candidatePath);
                foreach (var drug in drugs)
                {
                    drug.Period = period.Label;
                }
                candidates.Add(new TopDrugWorkbookCandidate(
                    candidatePath,
                    period.Label,
                    period.EndDate,
                    drugs,
                    GetCandidatePeriodRole(candidatePath, currentYearFolder, priorYearFolder)));
            }
            catch
            {
                // Ignore unrelated, locked, or non-Manulife workbooks in the selected folder.
            }
        }

        return candidates
            .OrderByDescending(candidate => candidate.EndDate ?? DateOnly.MinValue)
            .ThenByDescending(candidate => FileNameLooksLikeGenericRenewalWorkbook(candidate.SourcePath))
            .FirstOrDefault();
    }

    private static void ReadPriorTopDrugs(
        XLWorkbook currentWorkbook,
        string workbookPath,
        string? clientName,
        IEnumerable<string>? claimsWorkbookPaths,
        string? currentYearFolder,
        string? priorYearFolder,
        WorkbookExtractionResult result)
    {
        var currentFullPath = Path.GetFullPath(workbookPath);
        var hasPriorFolder = IsUsableDirectory(priorYearFolder);
        var folderCandidates = EnumerateExcelFiles(priorYearFolder)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        var explicitCandidates = (claimsWorkbookPaths ?? Enumerable.Empty<string>())
            .Where(path => !string.IsNullOrWhiteSpace(path) && File.Exists(path))
            .Where(path =>
                path.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                path.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase))
            .Where(path => !string.Equals(Path.GetFullPath(path), currentFullPath, StringComparison.OrdinalIgnoreCase))
            .Where(path => !IsPathInFolder(path, currentYearFolder))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        var candidatePaths = hasPriorFolder
            ? folderCandidates
            : folderCandidates
                .Concat(explicitCandidates)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        if (candidatePaths.Count == 0)
        {
            result.Warnings.Add("Prior Top DIN extraction warning: no prior-year workbook candidates were found.");
            return;
        }

        var prefixes = BuildClientFilePrefixes(clientName, workbookPath);
        var relevantCandidates = SelectRelevantExplicitWorkbookCandidates(candidatePaths, prefixes);
        var currentEndDate = ReadWorkbookPeriod(currentWorkbook, workbookPath).EndDate;
        var candidates = new List<TopDrugWorkbookCandidate>();
        foreach (var candidatePath in relevantCandidates)
        {
            try
            {
                using var candidateWorkbook = OpenWorkbookSharedRead(candidatePath);
                var drugs = ExtractTopDrugs(candidateWorkbook, $"{Path.GetFileName(candidatePath)}:Top_50_DIN");
                if (drugs.Count == 0)
                {
                    continue;
                }

                var period = ReadWorkbookPeriod(candidateWorkbook, candidatePath);
                foreach (var drug in drugs)
                {
                    drug.Period = period.Label;
                }
                candidates.Add(new TopDrugWorkbookCandidate(candidatePath, period.Label, period.EndDate, drugs, GetCandidatePeriodRole(candidatePath, currentYearFolder, priorYearFolder)));
            }
            catch
            {
                // Ignore unrelated or locked supporting files.
            }
        }

        if (candidates.Count == 0)
        {
            result.Warnings.Add("Prior Top DIN extraction warning: candidate prior-year workbooks did not contain a readable Top_50_DIN sheet.");
            return;
        }

        var priorCandidate = candidates
            .Where(candidate => candidate.PeriodRole.Equals("Prior", StringComparison.OrdinalIgnoreCase))
            .OrderByDescending(candidate => candidate.EndDate ?? DateOnly.MinValue)
            .FirstOrDefault();
        var priorFolderCandidateWasSelected = priorCandidate is not null;
        if (!priorFolderCandidateWasSelected && currentEndDate.HasValue)
        {
            var datedPriorCandidate = candidates
                .Where(candidate => candidate.EndDate.HasValue && candidate.EndDate.Value < currentEndDate.Value)
                .OrderByDescending(candidate => candidate.EndDate)
                .FirstOrDefault();
            if (datedPriorCandidate is not null &&
                (priorCandidate is null ||
                 !priorCandidate.EndDate.HasValue ||
                 priorCandidate.EndDate.Value >= currentEndDate.Value))
            {
                priorCandidate = datedPriorCandidate;
            }
        }

        if (priorCandidate is null)
        {
            result.Warnings.Add("Prior Top DIN extraction warning: no prior-year Top_50_DIN workbook could be identified. The prior drug table was left blank instead of using current-year data.");
            return;
        }

        if (TopDrugListsMatch(result.DrugTrends, priorCandidate.Drugs))
        {
            var alternatePriorCandidate = candidates
                .Where(candidate => !string.Equals(candidate.SourcePath, priorCandidate.SourcePath, StringComparison.OrdinalIgnoreCase))
                .Where(candidate => !TopDrugListsMatch(result.DrugTrends, candidate.Drugs))
                .OrderByDescending(candidate => candidate.PeriodRole.Equals("Prior", StringComparison.OrdinalIgnoreCase))
                .ThenByDescending(candidate => candidate.EndDate ?? DateOnly.MinValue)
                .FirstOrDefault();
            if (alternatePriorCandidate is null)
            {
                result.Warnings.Add($"Prior Top DIN extraction warning: skipped {Path.GetFileName(priorCandidate.SourcePath)} because its Top_50_DIN rows match the current Top DIN rows.");
                return;
            }

            priorCandidate = alternatePriorCandidate;
            priorFolderCandidateWasSelected = priorCandidate.PeriodRole.Equals("Prior", StringComparison.OrdinalIgnoreCase);
        }

        if (currentEndDate.HasValue &&
            priorCandidate.EndDate.HasValue &&
            priorCandidate.EndDate.Value >= currentEndDate.Value &&
            !priorFolderCandidateWasSelected)
        {
            result.Warnings.Add($"Prior Top DIN extraction warning: skipped {Path.GetFileName(priorCandidate.SourcePath)} because its period is not older than the current workbook.");
            return;
        }

        foreach (var drug in priorCandidate.Drugs)
        {
            result.PriorDrugTrends.Add(drug);
        }

        var periodText = string.IsNullOrWhiteSpace(priorCandidate.PeriodLabel)
            ? "period not identified"
            : priorCandidate.PeriodLabel;
        result.Warnings.Add($"Prior Top DIN source: {Path.GetFileName(priorCandidate.SourcePath)} ({periodText}).");
    }

    private static bool TopDrugListsMatch(IEnumerable<DrugTrend> currentDrugs, IEnumerable<DrugTrend> priorDrugs)
    {
        var currentSignature = BuildTopDrugSignature(currentDrugs);
        var priorSignature = BuildTopDrugSignature(priorDrugs);
        return currentSignature.Count > 0 &&
               priorSignature.Count > 0 &&
               currentSignature.SequenceEqual(priorSignature, StringComparer.OrdinalIgnoreCase);
    }

    private static List<string> BuildTopDrugSignature(IEnumerable<DrugTrend> drugs)
    {
        return drugs
            .Where(drug => !string.IsNullOrWhiteSpace(drug.DrugName))
            .Where(drug => !drug.DrugName.Equals("Grand Total", StringComparison.OrdinalIgnoreCase))
            .Where(drug => !drug.DrugName.Contains("Top 50", StringComparison.OrdinalIgnoreCase))
            .OrderByDescending(drug => drug.NetAmountPaid ?? drug.SubmittedTotal ?? 0m)
            .Take(10)
            .Select(drug => $"{NormalizeClientMatchText(drug.DrugName)}|{(drug.NetAmountPaid ?? drug.SubmittedTotal ?? 0m):0.00}")
            .ToList();
    }

    private static IReadOnlyList<string> SelectRelevantExplicitWorkbookCandidates(
        IReadOnlyList<string> explicitCandidates,
        IReadOnlyCollection<string> prefixes)
    {
        if (explicitCandidates.Count <= 1 || prefixes.Count == 0)
        {
            return explicitCandidates;
        }

        var clientMatches = explicitCandidates
            .Where(path => FileNameMatchesClientPrefix(path, prefixes))
            .ToList();
        var genericWorkbookFiles = explicitCandidates
            .Where(path => FileNameLooksLikeGenericClaimsWorkbook(path) ||
                           FileNameLooksLikeGenericRenewalWorkbook(path))
            .ToList();
        var preferred = clientMatches
            .Concat(genericWorkbookFiles)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        return preferred.Count > 0 ? preferred : explicitCandidates;
    }

    private static bool FileNameLooksLikeGenericRenewalWorkbook(string path)
    {
        var fileName = Path.GetFileNameWithoutExtension(path);
        return fileName.Contains("Renewal", StringComparison.OrdinalIgnoreCase) ||
               fileName.Contains("DetailedPackage", StringComparison.OrdinalIgnoreCase) ||
               fileName.Contains("Client_Copy", StringComparison.OrdinalIgnoreCase) ||
               fileName.Contains("Client Copy", StringComparison.OrdinalIgnoreCase);
    }

    private static ClaimsWorkbookMetadata ReadClaimsWorkbookMetadata(XLWorkbook workbook, string sourcePath)
    {
        string clientName = "";
        string policyNumber = "";
        var periodLabel = "";
        DateOnly? startDate = null;
        DateOnly? endDate = null;
        var sourceReference = "";

        var summary = workbook.Worksheets.FirstOrDefault(sheet => sheet.Name.Equals("Summary", StringComparison.OrdinalIgnoreCase));
        var sheets = summary is null
            ? workbook.Worksheets.ToList()
            : [summary, .. workbook.Worksheets.Where(sheet => !sheet.Name.Equals(summary.Name, StringComparison.OrdinalIgnoreCase))];

        foreach (var sheet in sheets)
        {
            var period = ReadExperiencePeriodLabel(sheet, sourcePath);
            if (string.IsNullOrWhiteSpace(periodLabel) && period.EndDate.HasValue)
            {
                periodLabel = period.Label;
                startDate = period.StartDate;
                endDate = period.EndDate;
                sourceReference = $"{sheet.Name}:period";
            }

            clientName = string.IsNullOrWhiteSpace(clientName)
                ? ReadClientNameFromSheetHeader(sheet)
                : clientName;
            policyNumber = string.IsNullOrWhiteSpace(policyNumber)
                ? ReadPolicyNumberFromSheetHeader(sheet)
                : policyNumber;

            if (!string.IsNullOrWhiteSpace(clientName) &&
                !string.IsNullOrWhiteSpace(policyNumber) &&
                endDate.HasValue)
            {
                break;
            }
        }

        if (!endDate.HasValue)
        {
            var filePeriod = TryParsePeriodFromFileName(Path.GetFileNameWithoutExtension(sourcePath));
            if (filePeriod.EndDate.HasValue)
            {
                periodLabel = filePeriod.Label;
                startDate = filePeriod.StartDate;
                endDate = filePeriod.EndDate;
                sourceReference = "file name";
            }
        }

        return new ClaimsWorkbookMetadata(
            clientName,
            policyNumber,
            periodLabel,
            startDate,
            endDate,
            sourceReference);
    }

    private static string ReadClientNameFromSheetHeader(IXLWorksheet sheet)
    {
        var lastRow = Math.Min(sheet.LastRowUsed()?.RowNumber() ?? 0, 12);
        var lastColumn = Math.Min(sheet.LastColumnUsed()?.ColumnNumber() ?? 0, 12);
        for (var row = 1; row <= lastRow; row++)
        {
            var rowText = string.Join(" ", Enumerable.Range(1, lastColumn).Select(column => Text(sheet.Cell(row, column))))
                .Trim();
            rowText = Regex.Replace(rowText, @"\s+", " ");
            if (LooksLikeClientNameHeader(rowText))
            {
                return rowText;
            }
        }

        return "";
    }

    private static bool LooksLikeClientNameHeader(string text)
    {
        if (string.IsNullOrWhiteSpace(text) || text.Length < 4)
        {
            return false;
        }

        if (!text.Any(char.IsLetter))
        {
            return false;
        }

        var ignored = new[]
        {
            "Manulife",
            "Summary",
            "Report",
            "Reporting period",
            "Experience Period",
            "Policy",
            "Coverage Groupings",
            "Employee Class",
            "Plan Design",
            "Group Claims Paid",
            "Claim Service Detail",
            "Benefit",
            "Total amount",
            "Number of",
            "Grand total",
            "The Manufacturers Life Insurance Company"
        };

        return !ignored.Any(item => text.Contains(item, StringComparison.OrdinalIgnoreCase));
    }

    private static string ReadPolicyNumberFromSheetHeader(IXLWorksheet sheet)
    {
        var lastRow = Math.Min(sheet.LastRowUsed()?.RowNumber() ?? 0, 12);
        var lastColumn = Math.Min(sheet.LastColumnUsed()?.ColumnNumber() ?? 0, 12);
        for (var row = 1; row <= lastRow; row++)
        {
            var rowText = string.Join(" ", Enumerable.Range(1, lastColumn).Select(column => Text(sheet.Cell(row, column))));
            var rowMatch = Regex.Match(
                rowText,
                @"(?:Policy|Group|Contract)\s*(?:Number|No\.?|#)?\s*:?\s*(?<number>\d{3,}[A-Za-z0-9-]*)",
                RegexOptions.IgnoreCase);
            if (rowMatch.Success)
            {
                return rowMatch.Groups["number"].Value;
            }

            for (var column = 1; column <= lastColumn; column++)
            {
                var text = Text(sheet.Cell(row, column));
                var match = Regex.Match(
                    text,
                    @"(?:Policy|Group|Contract)\s*(?:Number|No\.?|#)?\s*:?\s*(?<number>\d{3,}[A-Za-z0-9-]*)",
                    RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    return match.Groups["number"].Value;
                }
            }
        }

        return "";
    }

    private static void ValidateClaimsWorkbookIdentity(
        string sourcePath,
        ClaimsWorkbookMetadata metadata,
        string? expectedClientName,
        string? expectedPolicyNumber)
    {
        var policyIdentityAvailable = !string.IsNullOrWhiteSpace(expectedPolicyNumber) &&
                                      !string.IsNullOrWhiteSpace(metadata.PolicyNumber);
        var policyIdentityMatches = policyIdentityAvailable &&
                                    PolicyNumbersMatch(expectedPolicyNumber!, metadata.PolicyNumber);

        if (policyIdentityAvailable && !policyIdentityMatches)
        {
            throw new InvalidDataException(
                $"Claims workbook mismatch: {Path.GetFileName(sourcePath)} appears to be policy '{metadata.PolicyNumber}', but the renewal PDF is policy '{expectedPolicyNumber}'.");
        }

        if (!string.IsNullOrWhiteSpace(expectedClientName) &&
            !string.IsNullOrWhiteSpace(metadata.ClientName) &&
            !ClientNamesMatch(expectedClientName, metadata.ClientName) &&
            !policyIdentityMatches)
        {
            throw new InvalidDataException(
                $"Claims workbook mismatch: {Path.GetFileName(sourcePath)} appears to be for '{metadata.ClientName}', but the renewal PDF is for '{expectedClientName}'.");
        }
    }

    private static bool ClientNamesMatch(string expected, string actual)
    {
        var expectedNormalized = NormalizeClientIdentityText(expected);
        var actualNormalized = NormalizeClientIdentityText(actual);
        if (string.IsNullOrWhiteSpace(expectedNormalized) || string.IsNullOrWhiteSpace(actualNormalized))
        {
            return true;
        }

        return expectedNormalized.Equals(actualNormalized, StringComparison.OrdinalIgnoreCase) ||
               expectedNormalized.Contains(actualNormalized, StringComparison.OrdinalIgnoreCase) ||
               actualNormalized.Contains(expectedNormalized, StringComparison.OrdinalIgnoreCase) ||
               expectedNormalized[..Math.Min(6, expectedNormalized.Length)]
                   .Equals(actualNormalized[..Math.Min(6, actualNormalized.Length)], StringComparison.OrdinalIgnoreCase);
    }

    private static string NormalizeClientIdentityText(string value)
    {
        var ignored = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "THE", "AND", "LTD", "LIMITED", "INC", "INCORPORATED", "CORP", "CORPORATION", "COMPANY", "CO", "GROUP", "BENEFITS"
        };

        return string.Concat(
            Regex.Matches(value, @"[A-Za-z0-9]+")
                .Cast<Match>()
                .Select(match => match.Value.ToUpperInvariant())
                .Where(token => !ignored.Contains(token)));
    }

    private static bool PolicyNumbersMatch(string expected, string actual)
    {
        var expectedNormalized = NormalizePolicyNumber(expected);
        var actualNormalized = NormalizePolicyNumber(actual);
        return string.IsNullOrWhiteSpace(expectedNormalized) ||
               string.IsNullOrWhiteSpace(actualNormalized) ||
               expectedNormalized.Equals(actualNormalized, StringComparison.OrdinalIgnoreCase);
    }

    private static string NormalizePolicyNumber(string value)
    {
        var normalized = new string(value.Where(char.IsLetterOrDigit).Select(char.ToUpperInvariant).ToArray());
        if (normalized.Length > 0 && normalized.All(char.IsDigit))
        {
            normalized = normalized.TrimStart('0');
            return normalized.Length == 0 ? "0" : normalized;
        }

        return normalized;
    }

    private static (string Label, DateOnly? EndDate) ReadWorkbookPeriod(XLWorkbook workbook, string sourcePath)
    {
        var periods = workbook.Worksheets
            .Select(sheet => ReadExperiencePeriodLabel(sheet, sourcePath))
            .Where(period => period.EndDate.HasValue)
            .OrderByDescending(period => period.EndDate)
            .ToList();
        if (periods.Count > 0)
        {
            return (periods[0].Label, periods[0].EndDate);
        }

        var fileNamePeriod = TryParsePeriodFromFileName(Path.GetFileNameWithoutExtension(sourcePath));
        if (fileNamePeriod.EndDate.HasValue)
        {
            return (fileNamePeriod.Label, fileNamePeriod.EndDate);
        }

        var fileNameYear = TryParseYearFromFileName(Path.GetFileNameWithoutExtension(sourcePath));
        return fileNameYear.HasValue
            ? ($"Year ending {fileNameYear.Value}", new DateOnly(fileNameYear.Value, 12, 1))
            : ("", null);
    }

    private static int? TryParseYearFromFileName(string fileName)
    {
        var matches = Regex.Matches(fileName, @"(?<!\d)(20\d{2})(?!\d)")
            .Cast<Match>()
            .Select(match => int.Parse(match.Value, CultureInfo.InvariantCulture))
            .ToList();
        return matches.Count > 0 ? matches.Max() : null;
    }

    private static List<DrugTrend> ExtractTopDrugs(XLWorkbook workbook, string sourcePrefix)
    {
        var sheet = FindTopDrugWorksheet(workbook);
        if (sheet is null)
        {
            return [];
        }

        var headerRow = FindRow(sheet, "Drug Name");
        if (!headerRow.HasValue)
        {
            return [];
        }

        var standardLayout = sheet.Name.Equals("Top_50_DIN", StringComparison.OrdinalIgnoreCase);
        var dinColumn = FindColumn(sheet, headerRow.Value, "DIN") ?? (standardLayout ? 1 : null);
        var drugNameColumn = FindColumn(sheet, headerRow.Value, "Drug Name") ?? (standardLayout ? 2 : null);
        var classificationColumn = FindColumn(sheet, headerRow.Value, "Drug Classification") ??
                                    FindColumn(sheet, headerRow.Value, "Classification") ??
                                    FindColumn(sheet, headerRow.Value, "Treatment");
        var rxColumn = FindColumn(sheet, headerRow.Value, "# of Rxs") ??
                       FindColumn(sheet, headerRow.Value, "Rx Count") ??
                       FindColumn(sheet, headerRow.Value, "Number of Claims") ??
                       FindStackedHeaderColumn(sheet, headerRow.Value, "Number of Claims") ??
                       (standardLayout ? 9 : null);
        var submittedColumn = FindColumn(sheet, headerRow.Value, "Submitted Total") ?? (standardLayout ? 13 : null);
        var netAmountPaidColumn = FindColumn(sheet, headerRow.Value, "Net Amount Paid") ??
                                  FindColumn(sheet, headerRow.Value, "Total Amount Paid") ??
                                  FindColumn(sheet, headerRow.Value, "Amount Paid") ??
                                  FindStackedHeaderColumn(sheet, headerRow.Value, "Amount Paid") ??
                                  submittedColumn;
        var priorPaidColumn = FindColumn(sheet, headerRow.Value, "Prev Paid") ??
                              FindColumn(sheet, headerRow.Value, "Prior Paid") ??
                              (standardLayout ? 14 : null);
        if (!dinColumn.HasValue || !drugNameColumn.HasValue || !netAmountPaidColumn.HasValue)
        {
            return [];
        }

        var lastRow = sheet.LastRowUsed()?.RowNumber() ?? headerRow.Value;
        var lastColumn = sheet.LastColumnUsed()?.ColumnNumber() ?? netAmountPaidColumn.Value;
        var drugs = new List<DrugTrend>();
        for (var row = headerRow.Value + 1; row <= lastRow; row++)
        {
            var drugName = Text(sheet.Cell(row, drugNameColumn.Value));
            if (string.IsNullOrWhiteSpace(drugName) ||
                drugName.Contains("Total", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            drugs.Add(new DrugTrend
            {
                Din = Text(sheet.Cell(row, dinColumn.Value)),
                DrugName = drugName,
                Treats = classificationColumn.HasValue
                    ? Text(sheet.Cell(row, classificationColumn.Value))
                    : LookupDrugTreatment(drugName),
                PrescriptionCount = rxColumn.HasValue ? ToInt(Number(sheet.Cell(row, rxColumn.Value))) : null,
                SubmittedTotal = submittedColumn.HasValue ? Number(sheet.Cell(row, submittedColumn.Value)) : null,
                NetAmountPaid = Number(sheet.Cell(row, netAmountPaidColumn.Value)),
                PriorPaid = priorPaidColumn.HasValue ? Number(sheet.Cell(row, priorPaidColumn.Value)) : null,
                SourceReference = $"{sourcePrefix}:{sheet.Name}!A{row}:{sheet.Cell(row, lastColumn).Address}"
            });
        }

        return drugs;
    }

    private static IXLWorksheet? FindTopDrugWorksheet(XLWorkbook workbook)
    {
        if (workbook.TryGetWorksheet("Top_50_DIN", out var standardSheet))
        {
            return standardSheet;
        }

        return workbook.Worksheets.FirstOrDefault(sheet =>
        {
            var headerRow = FindRow(sheet, "Drug Name");
            if (!headerRow.HasValue)
            {
                return false;
            }

            return FindColumn(sheet, headerRow.Value, "DIN").HasValue &&
                   (FindColumn(sheet, headerRow.Value, "Net Amount Paid").HasValue ||
                    FindColumn(sheet, headerRow.Value, "Total Amount Paid").HasValue ||
                    FindColumn(sheet, headerRow.Value, "Amount Paid").HasValue ||
                    FindStackedHeaderColumn(sheet, headerRow.Value, "Amount Paid").HasValue);
        });
    }

    private static int? FindStackedHeaderColumn(IXLWorksheet sheet, int headerRow, string text)
    {
        var firstHeaderRow = Math.Max(1, headerRow - 2);
        var lastColumn = sheet.LastColumnUsed()?.ColumnNumber() ?? 0;
        var expected = Regex.Replace(text.Trim(), @"\s+", " ");

        for (var column = 1; column <= lastColumn; column++)
        {
            var stacked = Regex.Replace(
                string.Join(" ", Enumerable.Range(firstHeaderRow, headerRow - firstHeaderRow + 1)
                    .Select(row => Text(sheet.Cell(row, column)))
                    .Where(value => !string.IsNullOrWhiteSpace(value))),
                @"\s+",
                " ").Trim();
            if (stacked.Equals(expected, StringComparison.OrdinalIgnoreCase))
            {
                return column;
            }
        }

        for (var column = 1; column <= lastColumn; column++)
        {
            var stacked = Regex.Replace(
                string.Join(" ", Enumerable.Range(firstHeaderRow, headerRow - firstHeaderRow + 1)
                    .Select(row => Text(sheet.Cell(row, column)))
                    .Where(value => !string.IsNullOrWhiteSpace(value))),
                @"\s+",
                " ").Trim();
            if (stacked.Contains(expected, StringComparison.OrdinalIgnoreCase))
            {
                return column;
            }
        }

        return null;
    }

    private static string LookupDrugTreatment(string drugName)
    {
        if (string.IsNullOrWhiteSpace(drugName))
        {
            return "";
        }

        var normalized = NormalizeDrugName(drugName);
        foreach (var pair in DrugTreatmentLookup)
        {
            if (normalized.Contains(pair.Key, StringComparison.OrdinalIgnoreCase))
            {
                return pair.Value;
            }
        }

        return "Review required";
    }

    private static string NormalizeDrugName(string drugName)
    {
        return Regex.Replace(drugName.ToUpperInvariant(), @"[^A-Z0-9]+", " ").Trim();
    }

    private static readonly Dictionary<string, string> DrugTreatmentLookup = new(StringComparer.OrdinalIgnoreCase)
    {
        ["OZEMPIC"] = "Diabetes / weight management",
        ["WEGOVY"] = "Weight management",
        ["RYBELSUS"] = "Diabetes",
        ["TRULICITY"] = "Diabetes",
        ["MOUNJARO"] = "Diabetes / weight management",
        ["JARDIANCE"] = "Diabetes / heart failure",
        ["JANUVIA"] = "Diabetes",
        ["METFORMIN"] = "Diabetes",
        ["INSULIN"] = "Diabetes",
        ["DEXCOM"] = "Diabetes monitoring",
        ["FREESTYLE LIBRE"] = "Diabetes monitoring",
        ["HUMIRA"] = "Autoimmune inflammatory conditions",
        ["ADALIMUMAB"] = "Autoimmune inflammatory conditions",
        ["ENBREL"] = "Autoimmune inflammatory conditions",
        ["STELARA"] = "Psoriasis / Crohn's / colitis",
        ["COSENTYX"] = "Psoriasis / arthritis",
        ["SKYRIZI"] = "Psoriasis / Crohn's / colitis",
        ["RINVOQ"] = "Autoimmune inflammatory conditions",
        ["REMICADE"] = "Autoimmune inflammatory conditions",
        ["INFLECTRA"] = "Autoimmune inflammatory conditions",
        ["ENTYVIO"] = "Ulcerative colitis / Crohn's",
        ["XOLAIR"] = "Asthma / chronic hives",
        ["DUPIXENT"] = "Eczema / asthma",
        ["NUCALA"] = "Severe asthma",
        ["SYMBICORT"] = "Asthma / COPD",
        ["ADVAIR"] = "Asthma / COPD",
        ["VENTOLIN"] = "Asthma / COPD",
        ["FLOVENT"] = "Asthma",
        ["SPIRIVA"] = "COPD / asthma",
        ["DESVENLAFAXINE"] = "Depression",
        ["VENLAFAXINE"] = "Depression / anxiety",
        ["SERTRALINE"] = "Depression / anxiety",
        ["ESCITALOPRAM"] = "Depression / anxiety",
        ["BUPROPION"] = "Depression / smoking cessation",
        ["DULOXETINE"] = "Depression / anxiety / pain",
        ["VYVANSE"] = "ADHD",
        ["CONCERTA"] = "ADHD",
        ["INTUNIV"] = "ADHD",
        ["BIPHENTIN"] = "ADHD",
        ["ADDERALL"] = "ADHD",
        ["ATORVASTATIN"] = "High cholesterol",
        ["ROSUVASTATIN"] = "High cholesterol",
        ["EZETIMIBE"] = "High cholesterol",
        ["ELIQUIS"] = "Blood clot prevention",
        ["XARELTO"] = "Blood clot prevention",
        ["APIXABAN"] = "Blood clot prevention",
        ["RIVAROXABAN"] = "Blood clot prevention",
        ["PANTOPRAZOLE"] = "Acid reflux",
        ["DEXILANT"] = "Acid reflux",
        ["RABEPRAZOLE"] = "Acid reflux",
        ["PENTASA"] = "Inflammatory bowel disease",
        ["SULFASALAZINE"] = "Inflammatory bowel disease / arthritis",
        ["METHOTREXATE"] = "Autoimmune conditions / cancer",
        ["LEUCOVORIN"] = "Folate rescue / chemotherapy support",
        ["VALACYCLOVIR"] = "Viral infection",
        ["FOSFOMYCIN"] = "Urinary tract infection",
        ["COSOPT"] = "Glaucoma",
        ["COMBIGAN"] = "Glaucoma",
        ["VAGIFEM"] = "Menopause symptoms",
        ["NUVARING"] = "Birth control",
        ["CYPROTERONE ETHINYL ESTRADIOL"] = "Birth control / hormonal therapy",
        ["LANTUS"] = "Diabetes",
        ["TRESIBA"] = "Diabetes",
        ["HUMALOG"] = "Diabetes",
        ["NOVORAPID"] = "Diabetes",
        ["OFEV"] = "Pulmonary fibrosis",
        ["TRIKAFTA"] = "Cystic fibrosis",
        ["KESIMPTA"] = "Multiple sclerosis",
        ["OCREVUS"] = "Multiple sclerosis",
        ["AIMOVIG"] = "Migraine prevention",
        ["BOTOX"] = "Migraine / muscle conditions"
    };

    private static void ReadIndications(XLWorkbook workbook, WorkbookExtractionResult result)
    {
        if (!workbook.TryGetWorksheet("Top50CommonIndications", out var sheet))
        {
            return;
        }

        var headerRow = FindRow(sheet, "Common Indication") ?? 8;
        var lastRow = sheet.LastRowUsed()?.RowNumber() ?? headerRow;
        for (var row = headerRow + 1; row <= lastRow; row++)
        {
            var indication = Text(sheet.Cell(row, 3));
            if (string.IsNullOrWhiteSpace(indication) || indication.Equals("0", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            result.IndicationTrends.Add(new IndicationTrend
            {
                Indication = indication,
                PrescriptionCount = ToInt(Number(sheet.Cell(row, 5))),
                SubmittedTotal = Number(sheet.Cell(row, 9)),
                PriorPaid = Number(sheet.Cell(row, 10)),
                AllowableTotal = Number(sheet.Cell(row, 13)),
                SourceReference = $"Top50CommonIndications!C{row}:M{row}"
            });
        }
    }

    private static void ReadQuartiles(XLWorkbook workbook, WorkbookExtractionResult result)
    {
        if (!workbook.TryGetWorksheet("QuartileAnalysisDentHlth", out var sheet))
        {
            return;
        }

        var lastRow = sheet.LastRowUsed()?.RowNumber() ?? 0;
        for (var row = 1; row <= lastRow; row++)
        {
            var title = Text(sheet.Cell(row, 1));
            if (!title.Contains("% of Total Cost - by Quartile", StringComparison.OrdinalIgnoreCase) &&
                !title.Contains("Average Claim Cost - by Quartile", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var benefit = title.Contains("Dental", StringComparison.OrdinalIgnoreCase) ? "Dental" : "Extended Health Care";
            var isAverageCost = title.Contains("Average Claim Cost", StringComparison.OrdinalIgnoreCase);
            var headerRow = row + 2;

            for (var dataRow = headerRow + 1; dataRow <= Math.Min(headerRow + 5, lastRow); dataRow++)
            {
                var quartile = Text(sheet.Cell(dataRow, 1));
                if (!quartile.StartsWith("Quartile", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                for (var col = 2; col <= 9; col++)
                {
                    var category = Text(sheet.Cell(headerRow, col));
                    if (string.IsNullOrWhiteSpace(category))
                    {
                        continue;
                    }

                    var number = Number(sheet.Cell(dataRow, col));
                    if (!number.HasValue)
                    {
                        continue;
                    }

                    result.QuartileMetrics.Add(new QuartileMetric
                    {
                        Benefit = benefit,
                        Quartile = quartile,
                        Category = category,
                        PercentOfTotal = isAverageCost ? null : ToPercent(number),
                        AverageClaimCost = isAverageCost ? number : null,
                        SourceReference = $"QuartileAnalysisDentHlth!{ColumnName(col)}{dataRow}"
                    });
                }
            }
        }
    }

    private static void ReadCertificateClaims(XLWorkbook workbook, WorkbookExtractionResult result)
    {
        if (!workbook.TryGetWorksheet("TotalEHC", out var sheet))
        {
            return;
        }

        var headerRow = FindRow(sheet, "Member/ Dependents") ?? 10;
        var currentCertificate = "";

        var lastRow = sheet.LastRowUsed()?.RowNumber() ?? headerRow;
        for (var row = headerRow + 1; row <= lastRow; row++)
        {
            var certificate = Text(sheet.Cell(row, 2));
            if (!string.IsNullOrWhiteSpace(certificate))
            {
                currentCertificate = certificate;
            }

            if (!Text(sheet.Cell(row, 3)).Equals("Total Certificate", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            result.CertificateClaims.Add(new CertificateClaim
            {
                Certificate = currentCertificate,
                ExtendedHealthCare = Number(sheet.Cell(row, 4)),
                Dental = Number(sheet.Cell(row, 5)),
                Total = Number(sheet.Cell(row, 6)),
                Quartile = Text(sheet.Cell(row, 7)),
                SourceReference = $"TotalEHC!B{row}:G{row}"
            });
        }
    }

    private static int? FindRow(IXLWorksheet sheet, string text)
    {
        var lastRow = sheet.LastRowUsed()?.RowNumber() ?? 0;
        var lastCol = sheet.LastColumnUsed()?.ColumnNumber() ?? 0;
        for (var row = 1; row <= lastRow; row++)
        {
            for (var col = 1; col <= lastCol; col++)
            {
                if (Text(sheet.Cell(row, col)).Contains(text, StringComparison.OrdinalIgnoreCase))
                {
                    return row;
                }
            }
        }

        return null;
    }

    private static int? FindColumn(IXLWorksheet sheet, int row, string text)
    {
        var lastCol = sheet.LastColumnUsed()?.ColumnNumber() ?? 0;
        for (var col = 1; col <= lastCol; col++)
        {
            if (Text(sheet.Cell(row, col)).Equals(text, StringComparison.OrdinalIgnoreCase))
            {
                return col;
            }
        }

        for (var col = 1; col <= lastCol; col++)
        {
            if (Text(sheet.Cell(row, col)).Contains(text, StringComparison.OrdinalIgnoreCase))
            {
                return col;
            }
        }

        return null;
    }

    private static string Text(IXLCell cell)
    {
        try
        {
            return ParsingHelpers.Clean(cell.GetString());
        }
        catch
        {
            return ParsingHelpers.Clean(cell.GetFormattedString());
        }
    }

    private static decimal? Number(IXLCell cell)
    {
        if (cell.IsEmpty())
        {
            return null;
        }

        if (cell.TryGetValue<decimal>(out var number))
        {
            return number;
        }

        return ParsingHelpers.ParseDecimal(Text(cell));
    }

    private static decimal? ToPercent(decimal? value)
    {
        if (!value.HasValue)
        {
            return null;
        }

        return value.Value <= 1m ? decimal.Round(value.Value * 100m, 1) : decimal.Round(value.Value, 1);
    }

    private static int? ToInt(decimal? value)
    {
        return value.HasValue ? (int)Math.Round(value.Value, MidpointRounding.AwayFromZero) : null;
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

    private sealed record ServiceCategorySnapshot(
        string PeriodLabel,
        DateOnly? StartDate,
        DateOnly? EndDate,
        Dictionary<string, decimal> Amounts,
        string SourcePath,
        string SourceReference,
        string PeriodRole = "",
        string ClientName = "",
        string PolicyNumber = "");

    private sealed record ParamedicalPeriodTarget(
        string PeriodLabel,
        DateOnly StartDate,
        DateOnly EndDate,
        decimal ExpectedTotal);

    private sealed record ClaimsWorkbookMetadata(
        string ClientName,
        string PolicyNumber,
        string PeriodLabel,
        DateOnly? StartDate,
        DateOnly? EndDate,
        string SourceReference);

    private sealed record TopDrugWorkbookCandidate(
        string SourcePath,
        string PeriodLabel,
        DateOnly? EndDate,
        List<DrugTrend> Drugs,
        string PeriodRole = "");
}

using System.Globalization;
using System.Text.RegularExpressions;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed partial class ClaimsExperienceDataModelBuilder
{
    private static void MergeGroupSourceMonthlyReport(
        ClaimsExperienceDataModelData data,
        GroupSourceMonthlyClaimsReport report,
        List<string> warnings)
    {
        data.Carrier = FirstNonBlank(data.Carrier, "GroupSource");
        data.ClientName = FirstNonBlank(data.ClientName, report.ClientName);
        data.PolicyNumber = FirstNonBlank(data.PolicyNumber, report.PolicyNumber);
        data.PeriodStart ??= report.PeriodStart;
        data.PeriodEnd ??= report.PeriodEnd;
        data.PeriodLabel = FormatPeriodLabel(data.PeriodStart.Value, data.PeriodEnd.Value);
        data.Totals ??= new ClaimsExperienceTotals();

        var dental = report.Benefit.Equals(CoverageNames.Dental, StringComparison.OrdinalIgnoreCase);
        if (dental)
        {
            data.Totals.DentalPremium = report.TotalBilledPremium;
            data.Totals.DentalClaims = report.TotalPaidClaims;
        }
        else
        {
            data.Totals.EhcPremium = report.TotalBilledPremium;
            data.Totals.EhcClaims = report.TotalPaidClaims;
        }

        foreach (var source in report.Months)
        {
            var target = data.Months.FirstOrDefault(month => month.Month == source.Month);
            if (target is null)
            {
                target = new ClaimsExperienceMonth
                {
                    Month = source.Month,
                    SourceFile = Path.GetFileName(report.SourcePath)
                };
                data.Months.Add(target);
            }

            if (dental)
            {
                target.DentalPremium = source.BilledPremium;
                target.DentalClaims = source.PaidClaims;
                target.DentalSourceFile = Path.GetFileName(report.SourcePath);
                target.DentalPremiumMatchesMonth = true;
            }
            else
            {
                target.EhcPremium = source.BilledPremium;
                target.EhcClaims = source.PaidClaims;
                target.EhcSourceFile = Path.GetFileName(report.SourcePath);
                target.EhcPremiumMatchesMonth = true;
            }
        }

        data.Months.Sort((left, right) => left.Month.CompareTo(right.Month));
        warnings.AddRange(report.Warnings.Where(warning => !warnings.Contains(warning, StringComparer.OrdinalIgnoreCase)));
    }

    private void ReadHybridAsoDentalExperienceReports(
        IReadOnlyList<string> sourceFiles,
        ClaimsExperienceDataModelData data,
        List<string> warnings)
    {
        var reports = new List<HybridAsoDentalExperienceReport>();
        foreach (var path in sourceFiles
                     .Where(path => Path.GetExtension(path).Equals(".pdf", StringComparison.OrdinalIgnoreCase))
                     .Where(IsHybridAsoDentalClaimsExperienceFileName)
                     .Distinct(StringComparer.OrdinalIgnoreCase))
        {
            try
            {
                var pages = _pdfTextExtractor.Extract(path);
                var text = ParsingHelpers.Clean(string.Join(" ", pages.Select(page => page.Text)));
                if (!GroupHealthPdfParser.IsHybridAsoDentalClaimsExperienceDocument(path, text))
                {
                    continue;
                }

                var report = ParseHybridAsoDentalExperienceReport(path, pages);
                if (report is null)
                {
                    warnings.Add($"ASO dental monthly claims could not be read from {Path.GetFileName(path)}.");
                    continue;
                }

                reports.Add(report);
            }
            catch (Exception exception)
            {
                warnings.Add($"ASO dental experience report could not be read from {Path.GetFileName(path)}: {exception.Message}");
            }
        }

        foreach (var report in reports.OrderByDescending(report => report.PeriodEnd).ThenByDescending(report => report.PeriodStart))
        {
            var targetMonths = report.PeriodEnd == reports.Max(candidate => candidate.PeriodEnd)
                ? data.Months
                : report.PeriodEnd == reports.OrderByDescending(candidate => candidate.PeriodEnd).Skip(1).FirstOrDefault()?.PeriodEnd
                    ? data.PriorMonths
                    : data.PriorPriorMonths;

            MergeHybridAsoDentalMonths(targetMonths, report, warnings);
            if (ReferenceEquals(targetMonths, data.Months))
            {
                data.Carrier = FirstNonBlank(data.Carrier, "GroupHEALTH");
                data.ClientName = FirstNonBlank(data.ClientName, report.ClientName);
                data.PolicyNumber = FirstNonBlank(data.PolicyNumber, report.PolicyNumber);
                data.PeriodStart = report.PeriodStart;
                data.PeriodEnd = report.PeriodEnd;
                data.PeriodLabel = FormatPeriodLabel(report.PeriodStart, report.PeriodEnd);
                data.Totals ??= new ClaimsExperienceTotals();
                data.Totals.DentalPremium = report.Months.Sum(month => month.BilledPremium);
                data.Totals.DentalClaims = report.Months.Sum(month => month.PaidClaims);
            }

            if (!data.ImportedFiles.Any(record => record.FullPath.Equals(report.SourcePath, StringComparison.OrdinalIgnoreCase)))
            {
                var import = CreateImportRecord(report.SourcePath, "Hybrid ASO dental monthly claims PDF");
                import.RowsRead = $"{report.Months.Count} dental month(s)";
                import.Status = "Verified";
                import.Notes = "Standalone ASO dental paid-claims experience; billed premium is retained as reported.";
                data.ImportedFiles.Add(import);
            }
        }
    }

    private static HybridAsoDentalExperienceReport? ParseHybridAsoDentalExperienceReport(
        string path,
        IReadOnlyList<PdfPageText> pages)
    {
        var lines = pages.SelectMany(page => page.Lines).Select(ParsingHelpers.Clean).ToList();
        var periodLabel = GroupHealthClaimsPeriod(lines);
        if (periodLabel is null)
        {
            return null;
        }

        var months = new List<HybridAsoDentalExperienceMonth>();
        foreach (var line in lines)
        {
            var year = 0;
            var parsedMonth = default(DateTime);
            var match = Regex.Match(
                line,
                @"^(?<month>Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(?<year>20\d{2})\s+\$?(?<premium>[\d,]+\.\d{2})\s+\$?(?<claims>[\d,]+\.\d{2})(?:\s+.*)?$",
                RegexOptions.IgnoreCase);
            if (!match.Success ||
                !int.TryParse(match.Groups["year"].Value, NumberStyles.None, CultureInfo.InvariantCulture, out year) ||
                !DateTime.TryParseExact(match.Groups["month"].Value, "MMM", CultureInfo.InvariantCulture, DateTimeStyles.None, out parsedMonth) ||
                !ParsingHelpers.ParseDecimal(match.Groups["premium"].Value).HasValue ||
                !ParsingHelpers.ParseDecimal(match.Groups["claims"].Value).HasValue)
            {
                continue;
            }

            months.Add(new HybridAsoDentalExperienceMonth(
                new DateOnly(year, parsedMonth.Month, 1),
                ParsingHelpers.ParseDecimal(match.Groups["premium"].Value)!.Value,
                ParsingHelpers.ParseDecimal(match.Groups["claims"].Value)!.Value));
        }

        var expectedMonthCount = InclusiveMonthCount(periodLabel.Value.Start, periodLabel.Value.End);
        if (months.Count != expectedMonthCount || months.Select(month => month.Month).Distinct().Count() != months.Count)
        {
            return null;
        }

        var groupLine = lines.FirstOrDefault(line => line.StartsWith("Current Group ", StringComparison.OrdinalIgnoreCase)) ?? "";
        var groupValue = groupLine.StartsWith("Current Group ", StringComparison.OrdinalIgnoreCase)
            ? groupLine["Current Group ".Length..]
            : "";
        var policyMatch = Regex.Match(groupLine, @"\((?<policy>\d+)\)$");
        var clientName = Regex.Replace(groupValue, @"\s+\(\d+\)$", "").Trim();
        return new HybridAsoDentalExperienceReport(
            path,
            periodLabel.Value.Start,
            periodLabel.Value.End,
            clientName,
            policyMatch.Success ? policyMatch.Groups["policy"].Value : "",
            months);
    }

    private static (DateOnly Start, DateOnly End)? GroupHealthClaimsPeriod(IReadOnlyList<string> lines)
    {
        var from = lines.FirstOrDefault(line => line.StartsWith("Date From ", StringComparison.OrdinalIgnoreCase))?["Date From ".Length..];
        var to = lines.FirstOrDefault(line => line.StartsWith("Date To ", StringComparison.OrdinalIgnoreCase))?["Date To ".Length..];
        if (!DateTime.TryParseExact(from, "MMMM yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var start) ||
            !DateTime.TryParseExact(to, "MMMM yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var end))
        {
            return null;
        }

        return (new DateOnly(start.Year, start.Month, 1), new DateOnly(end.Year, end.Month, DateTime.DaysInMonth(end.Year, end.Month)));
    }

    private static void MergeHybridAsoDentalMonths(
        List<ClaimsExperienceMonth> targetMonths,
        HybridAsoDentalExperienceReport report,
        List<string> warnings)
    {
        foreach (var sourceMonth in report.Months)
        {
            var target = targetMonths.FirstOrDefault(month => month.Month == sourceMonth.Month);
            if (target is null)
            {
                targetMonths.Add(new ClaimsExperienceMonth
                {
                    Month = sourceMonth.Month,
                    DentalPremium = sourceMonth.BilledPremium,
                    DentalClaims = sourceMonth.PaidClaims,
                    DentalSourceFile = Path.GetFileName(report.SourcePath),
                    SourceFile = Path.GetFileName(report.SourcePath),
                    DentalPremiumMatchesMonth = sourceMonth.BilledPremium != 0m
                });
                continue;
            }

            if (target.DentalClaims != 0m && target.DentalClaims != sourceMonth.PaidClaims)
            {
                warnings.Add($"ASO dental paid claims for {sourceMonth.Month:MMM yyyy} in {Path.GetFileName(report.SourcePath)} conflict with an existing dental value; the existing value was retained.");
                continue;
            }

            target.DentalPremium = sourceMonth.BilledPremium;
            target.DentalClaims = sourceMonth.PaidClaims;
            target.DentalSourceFile = Path.GetFileName(report.SourcePath);
            target.DentalPremiumMatchesMonth = sourceMonth.BilledPremium != 0m;
        }

        targetMonths.Sort((left, right) => left.Month.CompareTo(right.Month));
    }

    private sealed record HybridAsoDentalExperienceReport(
        string SourcePath,
        DateOnly PeriodStart,
        DateOnly PeriodEnd,
        string ClientName,
        string PolicyNumber,
        IReadOnlyList<HybridAsoDentalExperienceMonth> Months);

    private sealed record HybridAsoDentalExperienceMonth(DateOnly Month, decimal BilledPremium, decimal PaidClaims);

    internal static int InclusiveMonthCount(DateOnly start, DateOnly end)
    {
        return end < start
            ? 0
            : ((end.Year - start.Year) * 12) + end.Month - start.Month + 1;
    }

    private static int NormalizeRequestedReportMonthCount(int monthCount)
    {
        return monthCount is 6 or 9 or 12 ? monthCount : 0;
    }

    private static void ResolveReportPeriodAndMonthCount(
        ClaimsExperienceDataModelData data,
        int requestedMonthCount,
        List<string> warnings)
    {
        if (data.PeriodStart is null && data.Months.Count > 0)
        {
            data.PeriodStart = data.Months.Min(month => month.Month);
        }

        if (data.PeriodEnd is null && data.Months.Count > 0)
        {
            data.PeriodEnd = EndOfMonth(data.Months.Max(month => month.Month));
        }

        var dateMonthCount = data.PeriodStart.HasValue && data.PeriodEnd.HasValue
            ? InclusiveMonthCount(data.PeriodStart.Value, data.PeriodEnd.Value)
            : 0;
        var rowMonthCount = data.Months
            .Where(month => !month.IsAnnualSummary)
            .Select(month => FirstDayOfMonth(month.Month))
            .Distinct()
            .Count();
        var detectedMonthCount = dateMonthCount > 0 ? dateMonthCount : rowMonthCount;
        data.DetectedReportMonthCount = detectedMonthCount;

        if (dateMonthCount > 0 && rowMonthCount > 0 && dateMonthCount != rowMonthCount)
        {
            warnings.Add($"The report period spans {dateMonthCount} month(s), but {rowMonthCount} distinct monthly row(s) were read. The report dates were used.");
        }

        var requested = NormalizeRequestedReportMonthCount(requestedMonthCount);
        if (requested > 0 && detectedMonthCount > 0 && requested > detectedMonthCount)
        {
            throw new InvalidDataException(
                $"The selected {requested}-month period is longer than the {detectedMonthCount}-month experience report. Select Auto or a shorter period.");
        }

        var effectiveMonthCount = requested > 0
            ? requested
            : detectedMonthCount > 0
                ? detectedMonthCount
                : 12;
        if (requested == 0 && detectedMonthCount == 0)
        {
            warnings.Add("The report month count could not be detected; 12 months was used as a fallback.");
        }

        if (requested > 0 &&
            detectedMonthCount > requested &&
            data.PeriodStart.HasValue)
        {
            data.PeriodEnd = EndOfMonth(data.PeriodStart.Value.AddMonths(requested - 1));
        }

        data.ReportMonthCount = effectiveMonthCount;
        data.CurrentSourcePeriodId = CurrentPeriodId(effectiveMonthCount);
        data.PriorSourcePeriodId = PriorPeriodId(effectiveMonthCount);
        data.BenchmarkSourcePeriodId = BenchmarkPeriodId(effectiveMonthCount);
        if (data.PeriodStart.HasValue && data.PeriodEnd.HasValue)
        {
            data.PeriodLabel = FormatPeriodLabel(data.PeriodStart.Value, data.PeriodEnd.Value);
        }
    }

    private string? FindClaimsServicePdf(
        IReadOnlyList<string> sourceFiles,
        string requiredNamePart,
        string? experienceReport,
        ClaimsExperienceDataModelData data)
    {
        var candidates = sourceFiles
            .Where(path => Path.GetExtension(path).Equals(".pdf", StringComparison.OrdinalIgnoreCase))
            .Where(path => !path.Equals(experienceReport, StringComparison.OrdinalIgnoreCase))
            .Where(path => ExperienceReportNamePriority(path) == int.MaxValue)
            .Where(path => Path.GetFileNameWithoutExtension(path).Contains(requiredNamePart, StringComparison.OrdinalIgnoreCase))
            .Where(path => Path.GetFileNameWithoutExtension(path).Contains("claim", StringComparison.OrdinalIgnoreCase))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (candidates.Count <= 1)
        {
            return candidates.FirstOrDefault();
        }

        var coverage = requiredNamePart.Equals("dental", StringComparison.OrdinalIgnoreCase)
            ? CoverageNames.Dental
            : CoverageNames.Ehc;
        return candidates
            .Select(path =>
            {
                var period = ReadClaimsServicePdfCurrentPeriod(path, coverage);
                var matchesExperiencePeriod = period is not null &&
                                              data.PeriodStart.HasValue &&
                                              data.PeriodEnd.HasValue &&
                                              period.Start == data.PeriodStart.Value &&
                                              period.End == data.PeriodEnd.Value;
                return new { Path = path, Period = period, MatchesExperiencePeriod = matchesExperiencePeriod };
            })
            .OrderByDescending(candidate => candidate.MatchesExperiencePeriod)
            .ThenByDescending(candidate => candidate.Period is not null)
            .ThenByDescending(candidate => candidate.Period?.End ?? DateOnly.MinValue)
            .ThenByDescending(candidate => candidate.Period?.Start ?? DateOnly.MinValue)
            .ThenBy(candidate => Path.GetFileName(candidate.Path), StringComparer.OrdinalIgnoreCase)
            .Select(candidate => candidate.Path)
            .FirstOrDefault();
    }

    private ClaimsSourcePeriod? ReadClaimsServicePdfCurrentPeriod(string path, string coverage)
    {
        try
        {
            return SunLifeClaimsPaidByServiceParser
                .Parse(_pdfTextExtractor.Extract(path), coverage)
                .Periods
                .OrderByDescending(period => period.End)
                .ThenByDescending(period => period.Start)
                .FirstOrDefault();
        }
        catch
        {
            return null;
        }
    }

    private void ValidateClaimsServicePdfIdentity(
        string path,
        string coverage,
        ClaimsExperienceDataModelData data,
        List<string> warnings)
    {
        ReadAndValidateClaimsServicePdf(path, coverage, data, warnings);
    }

    private void ReadClaimsServicePdf(
        string path,
        string coverage,
        ClaimsExperienceDataModelData data,
        List<string> warnings)
    {
        var pages = ReadAndValidateClaimsServicePdf(path, coverage, data, warnings);
        var parsed = SunLifeClaimsPaidByServiceParser.Parse(pages, coverage);
        var import = data.ImportedFiles.Last(record => record.FullPath.Equals(path, StringComparison.OrdinalIgnoreCase));
        if (parsed.Rows.Count == 0)
        {
            import.Status = "Needs review";
            import.Notes = "Identity matched, but Summary / Relation : All service rows were not read.";
            warnings.Add($"{coverage} Claims Paid by Type of Service rows could not be read from {Path.GetFileName(path)}.");
            return;
        }

        ApplyClaimsServicePeriodIdentity(parsed, path, data, warnings);
        AddClaimsPaidByServiceRows(parsed, path, coverage, data);
        AddClaimsServiceCanonicalRows(parsed, path, data);
        ValidateClaimsPaidByServiceTotals(parsed, path, coverage, data);
        import.RowsRead = $"{parsed.Rows.Count} summary service row(s); {parsed.ServiceFacts.Count} relationship service fact(s)";
        if (import.Status.Equals("Verified", StringComparison.OrdinalIgnoreCase))
        {
            import.Notes = $"Client and policy matched; current and prior {coverage} service totals reconciled.";
        }
    }

    private static void ApplyClaimsServicePeriodIdentity(
        SunLifeClaimsPaidByServiceResult parsed,
        string path,
        ClaimsExperienceDataModelData data,
        List<string> warnings)
    {
        if (parsed.Periods.Count == 0)
        {
            warnings.Add($"Explicit current/prior periods were not found in {Path.GetFileName(path)}; existing period selection was used.");
            return;
        }

        foreach (var period in parsed.Periods)
        {
            period.SourceFile = Path.GetFileName(path);
            if (!data.SourcePeriods.Any(existing => existing.Start == period.Start && existing.End == period.End))
            {
                data.SourcePeriods.Add(period);
            }
        }

        var ranked = parsed.Periods
            .OrderByDescending(period => period.End)
            .ThenByDescending(period => period.Start)
            .ToList();
        var current = ranked[0];
        var prior = ranked.Skip(1).FirstOrDefault();
        var detectedMonthCount = current.MonthCount;
        if (detectedMonthCount <= 0)
        {
            warnings.Add($"The report month count could not be detected from {Path.GetFileName(path)}; existing period selection was used.");
            return;
        }

        var requested = NormalizeRequestedReportMonthCount(data.RequestedReportMonthCount);
        if (requested > 0 && requested > detectedMonthCount)
        {
            throw new InvalidDataException(
                $"The selected {requested}-month period is longer than the {detectedMonthCount}-month period reported in {Path.GetFileName(path)}.");
        }

        var effectiveMonthCount = requested > 0 ? requested : detectedMonthCount;
        data.DetectedReportMonthCount = detectedMonthCount;
        data.ReportMonthCount = effectiveMonthCount;
        data.CurrentSourcePeriodId = CurrentPeriodId(effectiveMonthCount);
        data.PriorSourcePeriodId = PriorPeriodId(effectiveMonthCount);
        data.BenchmarkSourcePeriodId = BenchmarkPeriodId(effectiveMonthCount);
        data.PeriodStart = current.Start;
        data.PeriodEnd = requested > 0 && requested < detectedMonthCount
            ? EndOfMonth(current.Start.AddMonths(requested - 1))
            : current.End;
        data.PeriodLabel = FormatPeriodLabel(data.PeriodStart.Value, data.PeriodEnd.Value);

        if (prior is not null && prior.MonthCount != detectedMonthCount)
        {
            warnings.Add($"The prior period in {Path.GetFileName(path)} spans {prior.MonthCount} month(s), while the current period spans {detectedMonthCount} month(s).");
        }
    }

    private static void AddClaimsServiceCanonicalRows(
        SunLifeClaimsPaidByServiceResult parsed,
        string path,
        ClaimsExperienceDataModelData data)
    {
        var sourceFile = Path.GetFileName(path);
        foreach (var fact in parsed.ServiceFacts)
        {
            fact.SourceFile = sourceFile;
            data.ServiceAdjudicationFacts.Add(fact);
        }

        foreach (var total in parsed.ControlTotals)
        {
            total.SourceFile = sourceFile;
            data.ServiceControlTotals.Add(total);
        }
    }
    private List<PdfPageText> ReadAndValidateClaimsServicePdf(
        string path,
        string coverage,
        ClaimsExperienceDataModelData data,
        List<string> warnings)
    {
        var pages = _pdfTextExtractor.Extract(path);
        var identity = ClaimsPdfIdentityExtractor.Extract(pages);
        var import = CreateImportRecord(path, $"{coverage} Claims Paid by Type of Service PDF");
        import.RowsRead = "Header identity";
        import.Status = "Verified";
        import.Notes = "Client and policy matched the experience report.";

        if (string.IsNullOrWhiteSpace(identity.ClientName) || string.IsNullOrWhiteSpace(identity.PolicyNumber))
        {
            import.Status = "Needs review";
            import.Notes = "Client name or contract/policy number was not found in the PDF header.";
            warnings.Add($"{coverage} claims PDF identity is incomplete: {Path.GetFileName(path)}.");
        }

        if (!string.IsNullOrWhiteSpace(data.ClientName) &&
            !string.IsNullOrWhiteSpace(identity.ClientName) &&
            !BookletClientIdentityGuardrail.ClientNamesLikelyMatch(data.ClientName, identity.ClientName))
        {
            throw new InvalidDataException(
                $"{coverage} claims PDF client '{identity.ClientName}' does not match experience report client '{data.ClientName}' ({Path.GetFileName(path)}).");
        }

        if (!string.IsNullOrWhiteSpace(data.PolicyNumber) &&
            !string.IsNullOrWhiteSpace(identity.PolicyNumber) &&
            !NormalizeClaimsIdentifier(data.PolicyNumber).Equals(NormalizeClaimsIdentifier(identity.PolicyNumber), StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidDataException(
                $"{coverage} claims PDF contract/policy '{identity.PolicyNumber}' does not match '{data.PolicyNumber}' ({Path.GetFileName(path)}).");
        }

        if (!string.IsNullOrWhiteSpace(data.Carrier) &&
            !string.IsNullOrWhiteSpace(identity.Carrier) &&
            !data.Carrier.Equals(identity.Carrier, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidDataException(
                $"{coverage} claims PDF carrier '{identity.Carrier}' does not match '{data.Carrier}' ({Path.GetFileName(path)}).");
        }

        data.ClientName = FirstNonBlank(data.ClientName, identity.ClientName);
        data.PolicyNumber = FirstNonBlank(data.PolicyNumber, identity.PolicyNumber);
        data.Carrier = FirstNonBlank(data.Carrier, identity.Carrier);
        data.ImportedFiles.Add(import);
        return pages;
    }

    private static void AddClaimsPaidByServiceRows(
        SunLifeClaimsPaidByServiceResult parsed,
        string path,
        string coverage,
        ClaimsExperienceDataModelData data)
    {
        var sourceFile = Path.GetFileName(path);
        var categoryGroups = parsed.Rows
            .GroupBy(row => NormalizeLegacyClaimsCategory(coverage, row.Category), StringComparer.OrdinalIgnoreCase)
            .OrderBy(group => group.Key, StringComparer.OrdinalIgnoreCase)
            .ToList();

        foreach (var group in categoryGroups)
        {
            AddCategory(data.CurrentSourcePeriodId, group.Key,
                group.Sum(row => row.CurrentSubmitted),
                group.Sum(row => row.CurrentEligible),
                group.Sum(row => row.CurrentPaid));
            AddCategory(data.PriorSourcePeriodId, group.Key,
                group.Sum(row => row.PriorSubmitted),
                group.Sum(row => row.PriorEligible),
                group.Sum(row => row.PriorPaid));
        }

        if (coverage.Equals(CoverageNames.Ehc, StringComparison.OrdinalIgnoreCase))
        {
            foreach (var row in parsed.Rows.Where(row => row.Category.Equals("Paramedical", StringComparison.OrdinalIgnoreCase)))
            {
                data.ParamedicalDetails.Add(new ParamedicalDetail
                {
                    PeriodId = data.CurrentSourcePeriodId,
                    Service = NormalizeParamedicalService(row.Service),
                    ClaimCount = null,
                    SubmittedAmount = row.CurrentSubmitted,
                    EligibleAmount = row.CurrentEligible,
                    PaidClaims = row.CurrentPaid,
                    SourceFile = sourceFile
                });
                data.ParamedicalDetails.Add(new ParamedicalDetail
                {
                    PeriodId = data.PriorSourcePeriodId,
                    Service = NormalizeParamedicalService(row.Service),
                    ClaimCount = null,
                    SubmittedAmount = row.PriorSubmitted,
                    EligibleAmount = row.PriorEligible,
                    PaidClaims = row.PriorPaid,
                    SourceFile = sourceFile
                });
            }
        }

        return;

        void AddCategory(string periodId, string category, decimal submitted, decimal eligible, decimal paid)
        {
            var detail = new ClaimsCategoryDetail
            {
                PeriodId = periodId,
                Coverage = coverage,
                Category = category,
                ClaimCount = null,
                SubmittedAmount = submitted,
                EligibleAmount = eligible,
                PaidClaims = paid,
                SourceFile = sourceFile
            };
            if (coverage.Equals(CoverageNames.Ehc, StringComparison.OrdinalIgnoreCase))
            {
                data.EhcCategories.Add(detail);
            }
            else
            {
                data.DentalCategories.Add(detail);
            }
        }
    }

    private static string NormalizeLegacyClaimsCategory(string coverage, string category)
    {
        if (coverage.Equals(CoverageNames.Ehc, StringComparison.OrdinalIgnoreCase) &&
            (category.Equals("Pay Direct Drugs", StringComparison.OrdinalIgnoreCase) ||
             category.Equals("Reimbursement Drugs", StringComparison.OrdinalIgnoreCase)))
        {
            return "Drugs";
        }

        if (coverage.Equals(CoverageNames.Dental, StringComparison.OrdinalIgnoreCase))
        {
            if (category.Equals("Diagnostic / Preventive", StringComparison.OrdinalIgnoreCase) ||
                category.Equals("Surgical / Restorative", StringComparison.OrdinalIgnoreCase))
            {
                return "Basic Dental";
            }

            if (category.Equals("Major Restorative", StringComparison.OrdinalIgnoreCase))
            {
                return "Major Dental";
            }
        }

        return category;
    }
    private static void ValidateClaimsPaidByServiceTotals(
        SunLifeClaimsPaidByServiceResult parsed,
        string path,
        string coverage,
        ClaimsExperienceDataModelData data)
    {
        var currentServiceTotal = parsed.Rows.Sum(row => row.CurrentPaid);
        var priorServiceTotal = parsed.Rows.Sum(row => row.PriorPaid);
        Validate("current service rows", currentServiceTotal, parsed.CurrentTotalPaid);
        Validate("prior service rows", priorServiceTotal, parsed.PriorTotalPaid);

        var experiencePaidTotal = coverage.Equals(CoverageNames.Ehc, StringComparison.OrdinalIgnoreCase)
            ? data.Totals?.EhcClaims
            : data.Totals?.DentalClaims;
        Validate("current experience report", parsed.CurrentTotalPaid, experiencePaidTotal);

        return;

        void Validate(string comparison, decimal? actual, decimal? expected)
        {
            if (!actual.HasValue || !expected.HasValue || Math.Abs(actual.Value - expected.Value) <= 0.02m)
            {
                return;
            }

            throw new InvalidDataException(
                $"{coverage} paid claims from {comparison} ({actual.Value.ToString("C2", CultureInfo.CurrentCulture)}) do not match {expected.Value.ToString("C2", CultureInfo.CurrentCulture)} in {Path.GetFileName(path)}.");
        }
    }

    private static string NormalizeClaimsIdentifier(string value)
    {
        return new string((value ?? "").Where(char.IsLetterOrDigit).Select(char.ToUpperInvariant).ToArray());
    }
}


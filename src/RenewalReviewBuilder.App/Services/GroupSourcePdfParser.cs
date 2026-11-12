using System.Globalization;
using System.Text.RegularExpressions;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

/// <summary>Extracts GroupSource renewal presentations and SSQ-backed companion reports.</summary>
public sealed class GroupSourcePdfParser
{
    private const string CarrierName = "GroupSource";
    private const string Money = @"\$?[\d,]+\.\d{2}";
    private const string Amount = @"\$?[\d,]+(?:\.\d{2})?";

    private static readonly Regex StandardRateRegex = new(
        $@"^(?<benefit>Basic Life|A\s*D\s*&\s*D|Dependent Life|Long Term Disability|EFAP|Medical Second Opinion|Virtual Healthcare)\s+(?<basis>[\d,]+)\s+(?<currentRate>[\d.]+)\s+(?<currentPremium>{Money})\s+(?<renewalRate>[\d.]+)\s+(?<renewalPremium>{Money})$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex TierStartRegex = new(
        $@"^(?<benefit>Extended Health Care|Dental)\s+(?<tier>Single|Couple|Family)\s+(?<lives>\d+)\s+(?<currentRate>[\d.]+)\s+(?<currentPremium>{Money})\s+(?<renewalRate>[\d.]+)\s+(?<renewalPremium>{Money})$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex TierContinuationRegex = new(
        $@"^(?<tier>Single|Couple|Family)\s+(?<lives>\d+)\s+(?<currentRate>[\d.]+)\s+(?<currentPremium>{Money})\s+(?<renewalRate>[\d.]+)\s+(?<renewalPremium>{Money})$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex ClassTierRegex = new(
        @"^(?<class>(?:\d{2}\s*,?\s*)+)\s+(?<tier>Single|Couple|Family)$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex TierValuesOnlyRegex = new(
        $@"^(?<lives>\d+)\s+(?<currentRate>[\d.]+)\s+(?<currentPremium>{Money})\s+(?<renewalRate>[\d.]+)\s+(?<renewalPremium>{Money})$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex ExperienceRowRegex = new(
        $@"^(?<from>[A-Za-z]{{3}}\s+\d{{1,2}},\s+20\d{{2}})\s*-\s*(?<to>[A-Za-z]{{3}}\s+\d{{1,2}},\s+20\d{{2}})\s+(?<premium>{Amount})\s+(?<paid>{Amount})\s+(?<incurred>{Amount})\s+(?<ratio>[\d.]+)%$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex MonthlyRowRegex = new(
        $@"^(?<month>Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(?<year>20\d{{2}})\s+(?<premium>{Money})\s+(?<claims>{Money})\s+(?<ratio>[\d.]+)%$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex CategoryMoneyRegex = new(
        $@"^(?<category>[A-Z][A-Z /&'\-]+?)\s+(?<amount>{Money})$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex DrugRowRegex = new(
        $@"^(?<name>.+?)\s+(?<din>\d{{6,8}})\s+(?<paid>{Money})\s+(?<claims>\d+)\s+(?<percent>[\d.]+)%$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public static bool IsGroupSourceDocument(string pdfPath, string text)
    {
        var fileName = Path.GetFileNameWithoutExtension(pdfPath);
        var renewalPresentation = text.Contains("The following rates will be reflected", StringComparison.OrdinalIgnoreCase) &&
                                  text.Contains("Ledger #", StringComparison.OrdinalIgnoreCase);
        var experiencePresentation = text.Contains("EXPERIENCE CALCULATION", StringComparison.OrdinalIgnoreCase) &&
                                     text.Contains("Policy #", StringComparison.OrdinalIgnoreCase);
        var companionReport = text.Contains("Underwriter SSQ FINANCIAL", StringComparison.OrdinalIgnoreCase) &&
                              text.Contains("Claims Payor ALL", StringComparison.OrdinalIgnoreCase) &&
                              (ContainsAny(text, "Billed Premium & Claims", "Expense Information", "Paramedical Information") ||
                               IsCostCategoryReport(text));
        var p062SupportReport = fileName.StartsWith("P062_", StringComparison.OrdinalIgnoreCase) &&
                                text.Contains("Top 10 DINs by Dollar Amount", StringComparison.OrdinalIgnoreCase);
        return renewalPresentation || experiencePresentation || companionReport || p062SupportReport;
    }

    public CarrierPdfParseResult Parse(string pdfPath, List<PdfPageText> pages, string allText)
    {
        var pdf = new PdfExtractionResult
        {
            SourcePath = pdfPath,
            Carrier = CarrierName,
            ClientName = ParseClientName(pages),
            PolicyNumber = ParsePolicyNumber(pages),
            RenewalDate = ParseRenewalDate(pages)
        };
        var workbook = new WorkbookExtractionResult { SourcePath = pdfPath };

        ParseRates(pages, pdf);
        ParsePresentationExperience(pages, pdf);
        ParseExperienceCalculation(pages, pdf);
        ParseCompanionReports(pdfPath, pages, pdf, workbook);
        ParseTopDrugs(pages, workbook);

        if (pdf.RateLines.Count == 0 && pdf.ExperienceLines.Count == 0 && !HasWorkbookData(workbook))
        {
            pdf.Warnings.Add("No GroupSource renewal or claims values were extracted from this document.");
        }

        return new CarrierPdfParseResult
        {
            Carrier = CarrierName,
            Pdf = pdf,
            WorkbookFromPdf = HasWorkbookData(workbook) ? workbook : null
        };
    }

    internal static bool TryParseMonthlyClaimsReport(
        string pdfPath,
        IReadOnlyList<PdfPageText> pages,
        out GroupSourceMonthlyClaimsReport report)
    {
        report = default!;
        var page = pages.FirstOrDefault(candidate =>
            candidate.Text.Contains("Billed Premium & Claims", StringComparison.OrdinalIgnoreCase) &&
            candidate.Text.Contains("Date From", StringComparison.OrdinalIgnoreCase) &&
            candidate.Text.Contains("Date To", StringComparison.OrdinalIgnoreCase));
        if (page is null)
        {
            return false;
        }

        var lines = page.Lines.Select(ParsingHelpers.Clean).ToList();
        if (!TryParsePeriod(lines, out var start, out var end))
        {
            return false;
        }

        var benefit = ParseBenefit(lines);
        if (benefit.Length == 0)
        {
            return false;
        }

        var months = new List<GroupSourceMonthlyClaim>();
        foreach (var line in lines)
        {
            var match = MonthlyRowRegex.Match(line);
            if (!match.Success ||
                !int.TryParse(match.Groups["year"].Value, NumberStyles.None, CultureInfo.InvariantCulture, out var year) ||
                !DateTime.TryParseExact(match.Groups["month"].Value, "MMM", CultureInfo.InvariantCulture, DateTimeStyles.None, out var month))
            {
                continue;
            }

            months.Add(new GroupSourceMonthlyClaim(
                new DateOnly(year, month.Month, 1),
                Value(match, "premium") ?? 0m,
                Value(match, "claims") ?? 0m,
                Value(match, "ratio")));
        }

        var totalLine = lines.Select(line => Regex.Match(
                line,
                $@"^Grand Total\s+(?<premium>{Money})\s+(?<claims>{Money})(?:\s+(?<ratio>[\d.]+)%)?$",
                RegexOptions.IgnoreCase))
            .FirstOrDefault(match => match.Success);
        if (totalLine is null)
        {
            return false;
        }

        var warnings = new List<string>();
        var expectedMonths = ((end.Year - start.Year) * 12) + end.Month - start.Month + 1;
        if (months.Count != expectedMonths || months.Select(row => row.Month).Distinct().Count() != months.Count)
        {
            warnings.Add($"GroupSource monthly report page {page.PageNumber} covers {expectedMonths} month(s), but {months.Count} unique row(s) were read.");
        }

        var totalPremium = Value(totalLine, "premium") ?? 0m;
        var totalClaims = Value(totalLine, "claims") ?? 0m;
        if (Math.Abs(months.Sum(row => row.BilledPremium) - totalPremium) > 0.02m)
        {
            warnings.Add($"GroupSource {benefit} monthly billed premium does not reconcile to the source grand total.");
        }
        if (Math.Abs(months.Sum(row => row.PaidClaims) - totalClaims) > 0.02m)
        {
            warnings.Add($"GroupSource {benefit} monthly paid claims do not reconcile to the source grand total.");
        }

        report = new GroupSourceMonthlyClaimsReport(
            pdfPath,
            page.PageNumber,
            ParseClientName(pages),
            ParsePolicyNumber(pages),
            benefit,
            start,
            end,
            totalPremium,
            totalClaims,
            Value(totalLine, "ratio"),
            months,
            warnings);
        return true;
    }

    private static void ParseRates(IEnumerable<PdfPageText> pages, PdfExtractionResult pdf)
    {
        var page = pages.FirstOrDefault(candidate =>
            candidate.Text.Contains("The following rates will be reflected", StringComparison.OrdinalIgnoreCase) &&
            candidate.Text.Contains("Benefit Volume Current Rate", StringComparison.OrdinalIgnoreCase));
        if (page is null)
        {
            return;
        }

        var tierBenefit = "";
        var pendingTier = "";
        var pendingClass = "";
        foreach (var raw in page.Lines.Select(ParsingHelpers.Clean))
        {
            if (pendingTier.Length > 0)
            {
                var values = TierValuesOnlyRegex.Match(raw);
                if (values.Success)
                {
                    AddTierRate(pdf, page.PageNumber, tierBenefit, pendingTier, pendingClass, values);
                    pendingTier = "";
                    pendingClass = "";
                    continue;
                }
            }

            var standard = StandardRateRegex.Match(raw);
            if (standard.Success)
            {
                var benefit = NormalizeBenefit(standard.Groups["benefit"].Value);
                var basis = Value(standard, "basis");
                var livesBased = IsLivesBasedBenefit(benefit);
                var row = new RateLine
                {
                    Benefit = benefit,
                    TierOrClass = livesBased ? "Total" : "",
                    Lives = livesBased ? DecimalToInt(basis) : null,
                    Volume = livesBased ? null : basis,
                    CurrentRate = Value(standard, "currentRate"),
                    CurrentPremium = Value(standard, "currentPremium"),
                    RenewalRate = Value(standard, "renewalRate"),
                    RenewalPremium = Value(standard, "renewalPremium"),
                    SourceReference = $"GroupSource PDF page {page.PageNumber}"
                };
                ApplyCalculatedPremiums(row);
                pdf.RateLines.Add(row);
                tierBenefit = "";
                continue;
            }

            var tierStart = TierStartRegex.Match(raw);
            if (tierStart.Success)
            {
                tierBenefit = NormalizeBenefit(tierStart.Groups["benefit"].Value);
                AddTierRate(pdf, page.PageNumber, tierBenefit, tierStart);
                continue;
            }

            var tier = TierContinuationRegex.Match(raw);
            if (tier.Success && tierBenefit.Length > 0)
            {
                AddTierRate(pdf, page.PageNumber, tierBenefit, tier);
                continue;
            }

            var classTier = ClassTierRegex.Match(raw);
            if (classTier.Success && tierBenefit.Length > 0)
            {
                pendingTier = ParsingHelpers.Clean(classTier.Groups["tier"].Value);
                pendingClass = ParsingHelpers.Clean(classTier.Groups["class"].Value);
                foreach (var row in pdf.RateLines.Where(row =>
                             row.Benefit.Equals(tierBenefit, StringComparison.OrdinalIgnoreCase) &&
                             string.IsNullOrWhiteSpace(row.AccountClass)))
                {
                    row.AccountClass = pendingClass;
                }
                continue;
            }

            if (Regex.IsMatch(raw, @"^(?:\d{2}\s*,?\s*)+$") && tierBenefit.Length > 0)
            {
                var accountClass = Regex.Replace(raw, @"\s+", " ").Trim();
                foreach (var row in pdf.RateLines.Where(row =>
                             row.Benefit.Equals(tierBenefit, StringComparison.OrdinalIgnoreCase) &&
                             string.IsNullOrWhiteSpace(row.AccountClass)))
                {
                    row.AccountClass = accountClass;
                }
                continue;
            }

            var total = Regex.Match(raw, $@"^TOTAL\s+(?<current>{Money})\s+(?<renewal>{Money})$", RegexOptions.IgnoreCase);
            if (total.Success)
            {
                pdf.CurrentMonthlyPremium = Value(total, "current");
                pdf.RenewalMonthlyPremium = Value(total, "renewal");
                pdf.MonthlyPremiumChange = pdf.RenewalMonthlyPremium - pdf.CurrentMonthlyPremium;
                pdf.MonthlyPremiumChangePercent = Change(pdf.CurrentMonthlyPremium, pdf.RenewalMonthlyPremium);
            }
        }

        pdf.CurrentEhcPremium = SumPremium(pdf.RateLines, "Extended Health Care", true);
        pdf.RenewalEhcPremium = SumPremium(pdf.RateLines, "Extended Health Care", false);
        pdf.CurrentDentalPremium = SumPremium(pdf.RateLines, "Dental", true);
        pdf.RenewalDentalPremium = SumPremium(pdf.RateLines, "Dental", false);
        ValidateRateTotals(pdf);
    }

    private static void AddTierRate(PdfExtractionResult pdf, int pageNumber, string benefit, Match match)
    {
        AddTierRate(pdf, pageNumber, benefit, ParsingHelpers.Clean(match.Groups["tier"].Value), "", match);
    }

    private static void AddTierRate(
        PdfExtractionResult pdf,
        int pageNumber,
        string benefit,
        string tier,
        string accountClass,
        Match match)
    {
        var row = new RateLine
        {
            Benefit = benefit,
            TierOrClass = tier,
            AccountClass = accountClass,
            Lives = Int(match, "lives"),
            CurrentRate = Value(match, "currentRate"),
            CurrentPremium = Value(match, "currentPremium"),
            RenewalRate = Value(match, "renewalRate"),
            RenewalPremium = Value(match, "renewalPremium"),
            SourceReference = $"GroupSource PDF page {pageNumber}"
        };
        ApplyCalculatedPremiums(row);
        pdf.RateLines.Add(row);
    }

    private static void ApplyCalculatedPremiums(RateLine row)
    {
        var basis = row.Lives.HasValue ? row.Lives.Value : row.Volume;
        if (!basis.HasValue)
        {
            return;
        }

        var divisor = row.Benefit.Equals("LTD", StringComparison.OrdinalIgnoreCase)
            ? 100m
            : row.Lives.HasValue ? 1m : 1000m;
        row.CalculatedCurrentPremium = row.CurrentRate.HasValue
            ? decimal.Round(basis.Value * row.CurrentRate.Value / divisor, 2)
            : null;
        row.CalculatedRenewalPremium = row.RenewalRate.HasValue
            ? decimal.Round(basis.Value * row.RenewalRate.Value / divisor, 2)
            : null;
        row.CurrentVariance = row.CurrentPremium - row.CalculatedCurrentPremium;
        row.RenewalVariance = row.RenewalPremium - row.CalculatedRenewalPremium;
        row.PercentChange = Change(row.CurrentPremium, row.RenewalPremium);
    }

    private static void ParsePresentationExperience(IEnumerable<PdfPageText> pages, PdfExtractionResult pdf)
    {
        foreach (var page in pages)
        {
            var matches = page.Lines
                .Select(ParsingHelpers.Clean)
                .Select(line => ExperienceRowRegex.Match(line))
                .Where(match => match.Success)
                .ToList();
            if (matches.Count == 0)
            {
                continue;
            }

            var benefit = page.Text.Contains("Extended Health Care", StringComparison.OrdinalIgnoreCase)
                ? "Extended Health Care"
                : page.Text.Contains("Dental Care", StringComparison.OrdinalIgnoreCase) ? "Dental" : "";
            if (benefit.Length == 0)
            {
                continue;
            }

            foreach (var match in matches)
            {
                pdf.ExperienceLines.Add(new ExperienceLine
                {
                    Benefit = benefit,
                    Period = $"{match.Groups["from"].Value} to {match.Groups["to"].Value}",
                    BilledPremium = Value(match, "premium"),
                    PaidClaims = Value(match, "paid"),
                    IncurredClaims = Value(match, "incurred"),
                    PaidLossRatio = UnderwritingEngine.CalculateLossRatioPercent(Value(match, "paid"), Value(match, "premium")),
                    IncurredLossRatio = Value(match, "ratio"),
                    SourceReference = $"GroupSource PDF page {page.PageNumber}"
                });
            }
        }
    }

    private static void ParseExperienceCalculation(IEnumerable<PdfPageText> pages, PdfExtractionResult pdf)
    {
        var page = pages.FirstOrDefault(candidate =>
            candidate.Text.Contains("EXPERIENCE CALCULATION", StringComparison.OrdinalIgnoreCase) &&
            candidate.Text.Contains("Requested Rate Adjustment", StringComparison.OrdinalIgnoreCase));
        if (page is null)
        {
            return;
        }

        var lines = page.Lines.Select(ParsingHelpers.Clean).ToList();
        var trends = FindPercentPair(lines, "1 Annual Trend Factor");
        var credibility = FindPercentPair(lines, "2 Credibility Applied");
        var adjustments = FindPercentPair(lines, "F. Requested Rate Adjustment");
        EnrichCurrentExperience(pdf, "Extended Health Care", trends, credibility, adjustments, 0, page.PageNumber);
        EnrichCurrentExperience(pdf, "Dental", trends, credibility, adjustments, 1, page.PageNumber);
    }

    private static void EnrichCurrentExperience(
        PdfExtractionResult pdf,
        string benefit,
        IReadOnlyList<decimal> trends,
        IReadOnlyList<decimal> credibility,
        IReadOnlyList<decimal> adjustments,
        int index,
        int pageNumber)
    {
        var row = pdf.ExperienceLines.FirstOrDefault(candidate => candidate.Benefit.Equals(benefit, StringComparison.OrdinalIgnoreCase));
        if (row is null)
        {
            return;
        }

        row.TrendFactor = At(trends, index);
        row.Credibility = At(credibility, index);
        row.ProposedAdjustment = At(adjustments, index);
        row.SourceReference = $"{row.SourceReference}; GroupSource PDF page {pageNumber} renewal calculation";
    }

    private static void ParseCompanionReports(
        string pdfPath,
        IReadOnlyList<PdfPageText> pages,
        PdfExtractionResult pdf,
        WorkbookExtractionResult workbook)
    {
        if (TryParseMonthlyClaimsReport(pdfPath, pages, out var monthly))
        {
            pdf.Warnings.AddRange(monthly.Warnings);
            pdf.ExperienceLines.Add(new ExperienceLine
            {
                Benefit = monthly.Benefit,
                Period = PeriodLabel(monthly.PeriodStart, monthly.PeriodEnd),
                BilledPremium = monthly.TotalBilledPremium,
                PaidClaims = monthly.TotalPaidClaims,
                PaidLossRatio = monthly.ReportedPaidLossRatio ?? UnderwritingEngine.CalculateLossRatioPercent(monthly.TotalPaidClaims, monthly.TotalBilledPremium),
                SourceReference = $"GroupSource claims PDF page {monthly.PageNumber}"
            });
        }

        foreach (var page in pages)
        {
            var target = page.Text.Contains("Paramedical Information", StringComparison.OrdinalIgnoreCase)
                ? workbook.ParamedicalServiceTrends
                : IsCostCategoryReport(page.Text)
                    ? workbook.DentalSpendingTrends
                    : page.Text.Contains("Expense Information", StringComparison.OrdinalIgnoreCase)
                        ? workbook.EhcSpendingTrends
                        : null;
            if (target is null)
            {
                continue;
            }

            var lines = page.Lines.Select(ParsingHelpers.Clean).ToList();
            if (!TryParsePeriod(lines, out var start, out var end))
            {
                pdf.Warnings.Add($"GroupSource claims detail page {page.PageNumber} has no readable report period.");
                continue;
            }

            var period = PeriodLabel(start, end);
            var benefit = ParseBenefit(lines);
            var added = new List<(string Category, decimal Amount)>();
            foreach (var raw in lines)
            {
                var match = CategoryMoneyRegex.Match(raw);
                if (!match.Success || raw.StartsWith("Grand Total", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var category = NormalizeCategory(match.Groups["category"].Value);
                var amount = Value(match, "amount") ?? 0m;
                added.Add((category, amount));
                AddCategoryPeriod(target, benefit, category, period, amount, $"GroupSource claims PDF page {page.PageNumber}");
            }

            var grandTotal = lines.Select(line => Regex.Match(line, $@"^Grand Total\s+(?<amount>{Money})$", RegexOptions.IgnoreCase))
                .FirstOrDefault(match => match.Success);
            var total = grandTotal is null ? (decimal?)null : Value(grandTotal, "amount");
            foreach (var item in added)
            {
                workbook.ClaimTrends.Add(new ClaimTrend
                {
                    Benefit = benefit,
                    Category = item.Category,
                    Period = period,
                    Amount = item.Amount,
                    PercentOfTotal = total is > 0m ? decimal.Round(item.Amount / total.Value * 100m, 2) : null,
                    SourceReference = $"GroupSource claims PDF page {page.PageNumber}"
                });
            }

            if (total.HasValue && Math.Abs(added.Sum(row => row.Amount) - total.Value) > 0.02m)
            {
                workbook.Warnings.Add($"GroupSource {benefit} detail on page {page.PageNumber} does not reconcile to the source grand total.");
            }
        }
    }

    private static void ParseTopDrugs(IEnumerable<PdfPageText> pages, WorkbookExtractionResult workbook)
    {
        foreach (var page in pages.Where(candidate => candidate.Text.Contains("Top 10 DINs by Dollar Amount", StringComparison.OrdinalIgnoreCase)))
        {
            var periodMatch = Regex.Match(
                page.Text,
                @"From\s+(?<from>[A-Za-z]+\s+\d{1,2},\s+\d{4})\s*-\s*(?<to>[A-Za-z]+\s+\d{1,2},\s+\d{4})",
                RegexOptions.IgnoreCase);
            var period = periodMatch.Success
                ? $"{periodMatch.Groups["from"].Value} to {periodMatch.Groups["to"].Value}"
                : "";
            foreach (var raw in page.Lines.Select(ParsingHelpers.Clean))
            {
                var match = DrugRowRegex.Match(raw);
                if (!match.Success)
                {
                    continue;
                }

                workbook.DrugTrends.Add(new DrugTrend
                {
                    Din = match.Groups["din"].Value,
                    DrugName = ParsingHelpers.Clean(match.Groups["name"].Value),
                    Period = period,
                    PrescriptionCount = Int(match, "claims"),
                    NetAmountPaid = Value(match, "paid"),
                    SourceReference = $"GroupSource drug PDF page {page.PageNumber}"
                });
            }
        }
    }

    private static void ValidateRateTotals(PdfExtractionResult pdf)
    {
        if (!pdf.CurrentMonthlyPremium.HasValue || !pdf.RenewalMonthlyPremium.HasValue)
        {
            pdf.Warnings.Add("GroupSource rate table total was not found.");
            return;
        }

        var current = pdf.RateLines.Sum(row => row.CurrentPremium ?? 0m);
        var renewal = pdf.RateLines.Sum(row => row.RenewalPremium ?? 0m);
        if (Math.Abs(current - pdf.CurrentMonthlyPremium.Value) > 0.02m)
        {
            pdf.Warnings.Add($"GroupSource current rate rows total {current:C2}, but the source reports {pdf.CurrentMonthlyPremium:C2}.");
        }
        if (Math.Abs(renewal - pdf.RenewalMonthlyPremium.Value) > 0.02m)
        {
            pdf.Warnings.Add($"GroupSource renewal rate rows total {renewal:C2}, but the source reports {pdf.RenewalMonthlyPremium:C2}.");
        }
    }

    private static string ParseClientName(IEnumerable<PdfPageText> pages)
    {
        var lines = pages.SelectMany(page => page.Lines).Select(ParsingHelpers.Clean).ToList();
        var group = lines.Select(line => Regex.Match(line, @"^Group\s+(?<client>.+?)\s+\(\d+\)$", RegexOptions.IgnoreCase))
            .FirstOrDefault(match => match.Success);
        if (group is not null)
        {
            return ParsingHelpers.Clean(group.Groups["client"].Value);
        }

        var policy = lines.Select(line => Regex.Match(line, @"^Policy\s*#\s*\d+\s*-\s*(?<client>.+)$", RegexOptions.IgnoreCase))
            .FirstOrDefault(match => match.Success);
        if (policy is not null)
        {
            return ParsingHelpers.Clean(policy.Groups["client"].Value);
        }

        var firstPage = pages.FirstOrDefault()?.Lines.Select(ParsingHelpers.Clean).ToList() ?? [];
        var toIndex = firstPage.FindIndex(line => line.Equals("to", StringComparison.OrdinalIgnoreCase));
        if (toIndex >= 0 && toIndex + 1 < firstPage.Count)
        {
            return firstPage[toIndex + 1];
        }

        return "";
    }

    private static string ParsePolicyNumber(IEnumerable<PdfPageText> pages)
    {
        var lines = pages.SelectMany(page => page.Lines).Select(ParsingHelpers.Clean).ToList();
        var group = lines.Select(line => Regex.Match(line, @"^Group\s+.+?\s+\((?<policy>\d+)\)$", RegexOptions.IgnoreCase))
            .FirstOrDefault(match => match.Success);
        if (group is not null)
        {
            return group.Groups["policy"].Value;
        }

        var policy = lines.Select(line => Regex.Match(line, @"^Policy\s*#?\s*(?<policy>\d+)", RegexOptions.IgnoreCase))
            .FirstOrDefault(match => match.Success);
        return policy?.Groups["policy"].Value ?? "";
    }

    private static DateOnly? ParseRenewalDate(IEnumerable<PdfPageText> pages)
    {
        foreach (var line in pages.SelectMany(page => page.Lines).Select(ParsingHelpers.Clean))
        {
            var iso = Regex.Match(line, @"Renewal Date\s+(?<date>20\d{2}-\d{2}-\d{2})", RegexOptions.IgnoreCase);
            if (iso.Success && DateOnly.TryParseExact(iso.Groups["date"].Value, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var isoDate))
            {
                return isoDate;
            }

            var longDate = Regex.Match(line, @"\b(?<date>(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+20\d{2})\b", RegexOptions.IgnoreCase);
            if (longDate.Success && ParsingHelpers.ParseDateOnly(longDate.Groups["date"].Value) is { } date)
            {
                return date;
            }
        }

        return null;
    }

    private static bool TryParsePeriod(IReadOnlyList<string> lines, out DateOnly start, out DateOnly end)
    {
        start = default;
        end = default;
        var from = lines.FirstOrDefault(line => line.StartsWith("Date From ", StringComparison.OrdinalIgnoreCase))?["Date From ".Length..];
        var to = lines.FirstOrDefault(line => line.StartsWith("Date To ", StringComparison.OrdinalIgnoreCase))?["Date To ".Length..];
        if (!DateTime.TryParseExact(from, "MMMM yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var startDate) ||
            !DateTime.TryParseExact(to, "MMMM yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var endDate))
        {
            return false;
        }

        start = new DateOnly(startDate.Year, startDate.Month, 1);
        end = new DateOnly(endDate.Year, endDate.Month, DateTime.DaysInMonth(endDate.Year, endDate.Month));
        return true;
    }

    private static string ParseBenefit(IEnumerable<string> lines)
    {
        var values = lines.ToList();
        return values.Any(line => line.Contains("DENTAL", StringComparison.OrdinalIgnoreCase)) ? "Dental" :
            values.Any(line => line.Contains("EXTENDED HEALTH CARE", StringComparison.OrdinalIgnoreCase)) ? "Extended Health Care" : "";
    }

    private static List<decimal> FindPercentPair(IEnumerable<string> lines, string startsWith)
    {
        var line = lines.FirstOrDefault(candidate => candidate.StartsWith(startsWith, StringComparison.OrdinalIgnoreCase)) ?? "";
        return Regex.Matches(line, @"-?[\d.]+%")
            .Select(match => ParsingHelpers.ParseDecimal(match.Value))
            .Where(value => value.HasValue)
            .Select(value => value!.Value)
            .Take(2)
            .ToList();
    }

    private static void AddCategoryPeriod(
        List<ServiceCategoryTrend> rows,
        string benefit,
        string category,
        string period,
        decimal amount,
        string source)
    {
        var row = rows.FirstOrDefault(candidate =>
            Same(candidate.Benefit, benefit) && Same(candidate.Category, category));
        if (row is null)
        {
            row = new ServiceCategoryTrend { Benefit = benefit, Category = category, SourceReference = source };
            rows.Add(row);
        }
        ServiceCategoryHistory.Upsert(row, period, amount, source);
    }

    private static string PeriodLabel(DateOnly start, DateOnly end) =>
        $"{start:MMMM d, yyyy} to {end:MMMM d, yyyy}";

    private static string NormalizeBenefit(string value) => ParsingHelpers.Clean(value) switch
    {
        var item when Regex.IsMatch(item, @"^A\s*D\s*&\s*D$", RegexOptions.IgnoreCase) => "AD&D",
        var item when item.Equals("Basic Life", StringComparison.OrdinalIgnoreCase) => "Life",
        var item when item.Equals("Long Term Disability", StringComparison.OrdinalIgnoreCase) => "LTD",
        var item => item
    };

    private static string NormalizeCategory(string value) => ParsingHelpers.Clean(value).ToUpperInvariant() switch
    {
        "DRUGS" => "Drugs",
        "MEDICAL APPLIANCES / SUPPLIES" => "Medical Services and Supplies",
        "PARAMEDICALS" => "Paramedicals",
        "VISION CARE" => "Vision",
        "BASIC SERVICES" => "Basic",
        var item => CultureInfo.InvariantCulture.TextInfo.ToTitleCase(item.ToLowerInvariant())
    };

    private static bool IsLivesBasedBenefit(string benefit) =>
        benefit.Equals("Dependent Life", StringComparison.OrdinalIgnoreCase) ||
        benefit.Equals("EFAP", StringComparison.OrdinalIgnoreCase) ||
        benefit.Equals("Medical Second Opinion", StringComparison.OrdinalIgnoreCase) ||
        benefit.Equals("Virtual Healthcare", StringComparison.OrdinalIgnoreCase);

    private static bool HasWorkbookData(WorkbookExtractionResult workbook) =>
        workbook.ClaimTrends.Count > 0 ||
        workbook.EhcSpendingTrends.Count > 0 ||
        workbook.DentalSpendingTrends.Count > 0 ||
        workbook.ParamedicalServiceTrends.Count > 0 ||
        workbook.DrugTrends.Count > 0;

    private static decimal? SumPremium(IEnumerable<RateLine> rows, string benefit, bool current) =>
        rows.Where(row => row.Benefit.Equals(benefit, StringComparison.OrdinalIgnoreCase))
            .Sum(row => current ? row.CurrentPremium : row.RenewalPremium);

    private static decimal? Change(decimal? current, decimal? renewal) =>
        !current.HasValue || current == 0m || !renewal.HasValue
            ? null
            : decimal.Round((renewal.Value - current.Value) / current.Value * 100m, 1);

    private static decimal? Value(Match match, string name) => ParsingHelpers.ParseDecimal(match.Groups[name].Value);
    private static int? Int(Match match, string name) => ParsingHelpers.ParseInt(match.Groups[name].Value);
    private static int? DecimalToInt(decimal? value) => value.HasValue ? decimal.ToInt32(value.Value) : null;
    private static decimal? At(IReadOnlyList<decimal> values, int index) => index < values.Count ? values[index] : null;
    private static bool Same(string left, string right) => ParsingHelpers.Clean(left).Equals(ParsingHelpers.Clean(right), StringComparison.OrdinalIgnoreCase);
    private static bool ContainsAny(string text, params string[] values) => values.Any(value => text.Contains(value, StringComparison.OrdinalIgnoreCase));
    internal static bool IsCostCategoryReport(string text) =>
        Regex.IsMatch(text, @"Cost\s+Categor\s*y", RegexOptions.IgnoreCase);
}

internal sealed record GroupSourceMonthlyClaimsReport(
    string SourcePath,
    int PageNumber,
    string ClientName,
    string PolicyNumber,
    string Benefit,
    DateOnly PeriodStart,
    DateOnly PeriodEnd,
    decimal TotalBilledPremium,
    decimal TotalPaidClaims,
    decimal? ReportedPaidLossRatio,
    IReadOnlyList<GroupSourceMonthlyClaim> Months,
    IReadOnlyList<string> Warnings);

internal sealed record GroupSourceMonthlyClaim(
    DateOnly Month,
    decimal BilledPremium,
    decimal PaidClaims,
    decimal? ReportedPaidLossRatio);

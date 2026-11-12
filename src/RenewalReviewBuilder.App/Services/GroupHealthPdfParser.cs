using System.Globalization;
using System.Text.RegularExpressions;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

/// <summary>Extracts GroupHEALTH/GroupPLUS renewal and companion-report layouts into shared renewal models.</summary>
public sealed class GroupHealthPdfParser
{
    private static readonly Regex MoneyPairRegex = new(@"^\$?(?<current>[\d,]+\.\d{2})\s+\$?(?<renewal>[\d,]+\.\d{2})$", RegexOptions.Compiled);
    private static readonly Regex StandardRateRegex = new(@"^(?<benefit>Life Insurance|Accident & Serious Illness \(ASI\)|Dependent Life|Critical Illness|Long Term Disability)\s+(?<class>(?:\d{2}\s+)+)(?<lives>\d+)\s+(?<volume>[\d,]+)\s+(?<current>[\d.]+)\s+(?<renewal>[\d.]+)$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex VirtualRateRegex = new(@"^(?<benefit>Virtual Healthcare)\s+(?<class>(?:\d{2}\s+)+)(?<lives>\d+)\s+(?<current>[\d.]+)\s+(?<renewal>[\d.]+)$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex TierStartRegex = new(@"^(?<benefit>Extended Health Care|Dental Care|Evolve Drug|Evolve Major Bundle)\s+(?<class>(?:\d{2}\s+)+)(?<totalLives>\d+)\s+(?<tier>Single|Couple|Family)\s+(?<lives>\d+)\s+(?<current>[\d.]+)\s+(?<renewal>[\d.]+)$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex TierContinuationRegex = new(@"^(?:(?:\d{2}\s+)+)?(?<tier>Single|Couple|Family)\s+(?<lives>\d+)\s+(?<current>[\d.]+)\s+(?<renewal>[\d.]+)$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex CategoryMoneyRegex = new(@"^(?<category>[A-Z][A-Z /&-]+?)\s+\$?(?<amount>[\d,]+\.\d{2})$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex DrugRowRegex = new(@"^(?<name>.+?)\s+(?<din>\d{6,8})\s+\$(?<paid>[\d,]+\.\d{2})\s+(?<claims>\d+)\s+(?<percent>[\d.]+)%$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex AgeRowRegex = new(@"^(?<age><25|\d{2}-\d{2})\s*(?<values>.*)$", RegexOptions.Compiled);
    private const string AssurePlusRatingPattern =
        @"(?:(?:Top|Above|Bottom|Under|Below)\s*Performing|Not\s+Performing|Needs\s+Improvement|Performing)";
    private static readonly Regex AssurePlusBlockRegex = new(
        @"(?<benefit>Health|Dental)\s+Your Claims Experience\s+Rate Change effective\s+(?<date>(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4})(?<body>.*?)(?=(?:Health|Dental)\s+Your Claims Experience\s+Rate Change effective|$)",
        RegexOptions.IgnoreCase | RegexOptions.Singleline | RegexOptions.Compiled);

    public static bool IsGroupHealthDocument(string text)
    {
        return ContainsAny(text, "GroupPLUS", "GroupHEALTH", "clienthelp.grouphealth.ca") ||
               (ContainsAny(text, "Premium Comparison", "Billed Premium & Claims", "Paramedical Information") &&
                ContainsAny(text, "GROUPPLUS", "Current Underwriter SSQ FINANCIAL")) ||
               (text.Contains("Top 10 DINs by Dollar Amount", StringComparison.OrdinalIgnoreCase) &&
                text.Contains("Policy #", StringComparison.OrdinalIgnoreCase) &&
                ContainsAny(text, "# Paid Claims", "Total $ Paid"));
    }

    public CarrierPdfParseResult Parse(string pdfPath, List<PdfPageText> pages, string allText)
    {
        var pdf = new PdfExtractionResult
        {
            SourcePath = pdfPath,
            Carrier = "GroupHEALTH",
            ClientName = ParseClientName(pages),
            PolicyNumber = ParsePolicyNumber(pdfPath, pages),
            RenewalDate = ParseRenewalDate(pages)
        };
        var workbook = new WorkbookExtractionResult { SourcePath = pdfPath };

        ParsePremiumComparison(pages, pdf);
        ParseAsoBudgetedBenefits(pages, pdf);
        ParseAssurePlusPricing(pages, pdf);
        ParseRenewalExperience(pages, pdf);
        ParseDemographics(pages, pdf);
        ParseClaimsCompanionPages(pages, pdf, workbook);
        ParseTopDrugs(pages, workbook);
        ParsePoolingSupport(pages, pdf);

        if (pdf.RateLines.Count == 0 && pdf.ExperienceLines.Count == 0 &&
            pdf.AssurePlusPricingReports.Count == 0 && !HasWorkbookData(workbook))
        {
            pdf.Warnings.Add("No GroupHEALTH renewal or claims values were extracted from this document.");
        }

        return new CarrierPdfParseResult
        {
            Carrier = "GroupHEALTH",
            Pdf = pdf,
            WorkbookFromPdf = HasWorkbookData(workbook) ? workbook : null
        };
    }

    /// <summary>
    /// AssurePLUS groups can have ASO dental administered on the GroupHEALTH platform.  Those
    /// dental experience reports do not carry the usual GroupHEALTH branding, but have a stable
    /// date-range and paid-claims layout.
    /// </summary>
    public static bool IsHybridAsoDentalClaimsExperienceDocument(string pdfPath, string text)
    {
        var fileName = Path.GetFileNameWithoutExtension(pdfPath);
        var hasDentalSignal = fileName.Contains("DEN", StringComparison.OrdinalIgnoreCase) ||
                              text.Contains("Benefit DENTAL", StringComparison.OrdinalIgnoreCase) ||
                              text.Contains("Cost Category", StringComparison.OrdinalIgnoreCase);
        var hasMonthlyClaimsLayout = text.Contains("Billed Premium", StringComparison.OrdinalIgnoreCase) &&
                                     text.Contains("Paid Claims", StringComparison.OrdinalIgnoreCase) &&
                                     Regex.IsMatch(text, @"\bDate\s+From\s+[A-Za-z]+\s+20\d{2}", RegexOptions.IgnoreCase) &&
                                     Regex.IsMatch(text, @"\bDate\s+To\s+[A-Za-z]+\s+20\d{2}", RegexOptions.IgnoreCase);
        return hasDentalSignal && hasMonthlyClaimsLayout;
    }

    private static void ParseAssurePlusPricing(IEnumerable<PdfPageText> pages, PdfExtractionResult pdf)
    {
        foreach (var page in pages.Where(candidate =>
                     candidate.Text.Contains("How Your Health and Dental Pricing Is Determined", StringComparison.OrdinalIgnoreCase) &&
                     candidate.Text.Contains("Your Claims Experience", StringComparison.OrdinalIgnoreCase) &&
                     candidate.Text.Contains("Rate Change effective", StringComparison.OrdinalIgnoreCase)))
        {
            var normalized = ParsingHelpers.Clean(page.Text);
            var matches = AssurePlusBlockRegex.Matches(normalized);
            if (matches.Count == 0)
            {
                pdf.Warnings.Add($"AssurePLUS pricing page {page.PageNumber} was found, but its Health and Dental pricing table could not be read.");
                continue;
            }

            var report = new AssurePlusPricingReport
            {
                SourcePath = pdf.SourcePath,
                SourcePage = page.PageNumber
            };
            foreach (Match match in matches)
            {
                var body = match.Groups["body"].Value;
                var ratingsSection = Regex.Match(
                    body,
                    @"Previous Year\s+Recent Year(?<ratings>.*?)Adjustment for your experience",
                    RegexOptions.IgnoreCase | RegexOptions.Singleline);
                var ratings = ratingsSection.Success
                    ? Regex.Matches(ratingsSection.Groups["ratings"].Value, AssurePlusRatingPattern, RegexOptions.IgnoreCase)
                        .Select(item => NormalizeAssurePlusRating(item.Value))
                        .Take(2)
                        .ToList()
                    : [];
                var standard = Regex.Match(body, @"Standard:\s*(?<value>[-−–—]?[\d.]+)%", RegexOptions.IgnoreCase);
                var adjustment = Regex.Match(body, @"Adjustment for your experience:\s*(?<value>[-−–—]?[\d.]+)%", RegexOptions.IgnoreCase);
                var total = Regex.Match(body, @"Total change:\s*(?<value>[-−–—]?[\d.]+)%", RegexOptions.IgnoreCase);
                var effectiveDate = ParsingHelpers.ParseDateOnly(match.Groups["date"].Value);
                report.EffectiveDate ??= effectiveDate;
                if (report.EffectiveDate.HasValue && effectiveDate.HasValue && report.EffectiveDate != effectiveDate)
                {
                    report.Warnings.Add($"AssurePLUS pricing page {page.PageNumber} contains more than one effective date.");
                }

                var benefit = match.Groups["benefit"].Value.Equals("Health", StringComparison.OrdinalIgnoreCase)
                    ? "Health"
                    : "Dental";
                var row = new AssurePlusBenefitPricing
                {
                    Benefit = benefit,
                    PreviousYearRating = ratings.ElementAtOrDefault(0) ?? "",
                    RecentYearRating = ratings.ElementAtOrDefault(1) ?? "",
                    StandardChangePercent = AssurePlusPercent(standard, "value"),
                    ExperienceAdjustmentPercent = AssurePlusPercent(adjustment, "value"),
                    TotalChangePercent = AssurePlusPercent(total, "value"),
                    SourceReference = $"{Path.GetFileName(pdf.SourcePath)} page {page.PageNumber}"
                };
                report.Benefits.Add(row);
                if (ratings.Count < 2)
                {
                    report.Warnings.Add($"AssurePLUS {benefit} previous/recent experience statuses are incomplete in {report.SourceReference}.");
                }
                ValidateAssurePlusRateArithmetic(report, row);
            }

            ParseAssurePlusDistribution(
                report,
                normalized,
                "Health",
                ["Drugs", "Paramedicals", "Supplies", "Vision", "Other"]);
            ParseAssurePlusDistribution(
                report,
                normalized,
                "Dental",
                ["Basic Other", "Periodontics", "Endodontics", "Major", "Ortho"]);

            if (report.Benefits.Count < 2)
            {
                report.Warnings.Add($"AssurePLUS pricing page {page.PageNumber} did not contain both Health and Dental pricing rows.");
            }

            if (report.EffectiveDate.HasValue && pdf.RenewalDate.HasValue && report.EffectiveDate != pdf.RenewalDate)
            {
                report.Warnings.Add(
                    $"AssurePLUS effective date {report.EffectiveDate:MMMM d, yyyy} does not match the renewal date {pdf.RenewalDate:MMMM d, yyyy}.");
            }

            pdf.AssurePlusPricingReports.Add(report);
        }
    }

    private static void ParseAssurePlusDistribution(
        AssurePlusPricingReport report,
        string normalizedPageText,
        string benefit,
        IReadOnlyList<string> categories)
    {
        var categoryPattern = string.Join(@"\s+", categories.Select(Regex.Escape));
        var numberPattern = @"(?<value>-?[\d.]+)%";
        var fiveValues = string.Join(@"\s+", Enumerable.Repeat(numberPattern, categories.Count));
        var match = Regex.Match(
            normalizedPageText,
            $@"{categoryPattern}\s+Your Claims\s+{fiveValues}\s+Benchmark Claims\s+{fiveValues}",
            RegexOptions.IgnoreCase);
        if (!match.Success || match.Groups["value"].Captures.Count != categories.Count * 2)
        {
            report.Warnings.Add($"AssurePLUS {benefit} claims distribution could not be read from {report.SourceReference}.");
            return;
        }

        var values = match.Groups["value"].Captures
            .Select(capture => ParsingHelpers.ParseDecimal(capture.Value))
            .ToList();
        for (var index = 0; index < categories.Count; index++)
        {
            report.DistributionRows.Add(new AssurePlusDistributionRow
            {
                Benefit = benefit,
                Category = categories[index],
                YourClaimsPercent = values[index],
                BenchmarkClaimsPercent = values[index + categories.Count],
                SourceReference = report.SourceReference
            });
        }

        ValidateAssurePlusDistribution(report, benefit, "Your Claims", values.Take(categories.Count));
        ValidateAssurePlusDistribution(report, benefit, "Benchmark Claims", values.Skip(categories.Count));
    }

    private static void ValidateAssurePlusRateArithmetic(
        AssurePlusPricingReport report,
        AssurePlusBenefitPricing row)
    {
        if (!row.StandardChangePercent.HasValue ||
            !row.ExperienceAdjustmentPercent.HasValue ||
            !row.TotalChangePercent.HasValue)
        {
            report.Warnings.Add($"AssurePLUS {row.Benefit} rate change is incomplete in {report.SourceReference}.");
            return;
        }

        var calculated = row.StandardChangePercent.Value + row.ExperienceAdjustmentPercent.Value;
        if (Math.Abs(calculated - row.TotalChangePercent.Value) > 0.01m)
        {
            report.Warnings.Add(
                $"AssurePLUS {row.Benefit} rate change does not reconcile in {report.SourceReference}: " +
                $"{row.StandardChangePercent:0.00}% + {row.ExperienceAdjustmentPercent:0.00}% does not equal {row.TotalChangePercent:0.00}%.");
        }
    }

    private static void ValidateAssurePlusDistribution(
        AssurePlusPricingReport report,
        string benefit,
        string series,
        IEnumerable<decimal?> values)
    {
        var items = values.ToList();
        if (items.Any(value => !value.HasValue))
        {
            report.Warnings.Add($"AssurePLUS {benefit} {series} distribution is incomplete in {report.SourceReference}.");
            return;
        }

        var total = items.Sum(value => value!.Value);
        if (Math.Abs(total - 100m) > 0.05m)
        {
            report.Warnings.Add($"AssurePLUS {benefit} {series} distribution totals {total:0.00}% instead of 100.00% in {report.SourceReference}.");
        }
    }

    private static string NormalizeAssurePlusRating(string value)
    {
        var normalized = string.Join(' ',
            value.Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .Select(word => CultureInfo.InvariantCulture.TextInfo.ToTitleCase(word.ToLowerInvariant())));
        foreach (var prefix in new[] { "Top", "Above", "Bottom", "Under", "Below" })
        {
            if (normalized.Equals($"{prefix}performing", StringComparison.OrdinalIgnoreCase))
            {
                return $"{prefix} Performing";
            }
        }

        return normalized;
    }

    private static decimal? AssurePlusPercent(Match match, string name) =>
        ParsingHelpers.ParseDecimal(match.Groups[name].Value
            .Replace('−', '-')
            .Replace('–', '-')
            .Replace('—', '-'));

    private static void ParsePremiumComparison(IEnumerable<PdfPageText> pages, PdfExtractionResult pdf)
    {
        var page = pages.FirstOrDefault(candidate =>
            candidate.Text.Contains("Premium Comparison", StringComparison.OrdinalIgnoreCase) &&
            candidate.Text.Contains("Total Monthly Premiums", StringComparison.OrdinalIgnoreCase));
        if (page is null) return;

        RateLine? pending = null;
        var tierBenefit = "";
        var tierClass = "";
        foreach (var raw in page.Lines.Select(ParsingHelpers.Clean))
        {
            if (pending is not null)
            {
                var premiums = MoneyPairRegex.Match(raw);
                if (premiums.Success)
                {
                    pending.CurrentPremium = Value(premiums, "current");
                    pending.RenewalPremium = Value(premiums, "renewal");
                    pending.PercentChange = Change(pending.CurrentPremium, pending.RenewalPremium);
                    pdf.RateLines.Add(pending);
                    pending = null;
                    continue;
                }
            }

            var standard = StandardRateRegex.Match(raw);
            if (standard.Success)
            {
                pending = new RateLine
                {
                    Benefit = NormalizeBenefit(standard.Groups["benefit"].Value),
                    TierOrClass = ParsingHelpers.Clean(standard.Groups["class"].Value),
                    AccountClass = ParsingHelpers.Clean(standard.Groups["class"].Value),
                    Lives = Int(standard, "lives"),
                    Volume = Value(standard, "volume"),
                    CurrentRate = Value(standard, "current"),
                    RenewalRate = Value(standard, "renewal"),
                    SourceReference = $"GroupHEALTH PDF page {page.PageNumber}"
                };
                continue;
            }

            var virtualRate = VirtualRateRegex.Match(raw);
            if (virtualRate.Success)
            {
                pending = new RateLine
                {
                    Benefit = NormalizeBenefit(virtualRate.Groups["benefit"].Value),
                    TierOrClass = "Total",
                    AccountClass = ParsingHelpers.Clean(virtualRate.Groups["class"].Value),
                    Lives = Int(virtualRate, "lives"),
                    CurrentRate = Value(virtualRate, "current"),
                    RenewalRate = Value(virtualRate, "renewal"),
                    SourceReference = $"GroupHEALTH PDF page {page.PageNumber}"
                };
                continue;
            }

            var tierStart = TierStartRegex.Match(raw);
            if (tierStart.Success)
            {
                tierBenefit = NormalizeBenefit(tierStart.Groups["benefit"].Value);
                tierClass = ParsingHelpers.Clean(tierStart.Groups["class"].Value);
                AddTierRate(pdf, page.PageNumber, tierBenefit, tierClass, tierStart);
                continue;
            }

            var tier = TierContinuationRegex.Match(raw);
            if (tier.Success && tierBenefit.Length > 0)
            {
                AddTierRate(pdf, page.PageNumber, tierBenefit, tierClass, tier);
                continue;
            }

            var monthly = Regex.Match(raw, @"^Total Monthly Premiums\s+\$(?<current>[\d,]+\.\d{2})\s+\$(?<renewal>[\d,]+\.\d{2})$", RegexOptions.IgnoreCase);
            if (monthly.Success)
            {
                pdf.CurrentMonthlyPremium = Value(monthly, "current");
                pdf.RenewalMonthlyPremium = Value(monthly, "renewal");
                pdf.MonthlyPremiumChange = pdf.RenewalMonthlyPremium - pdf.CurrentMonthlyPremium;
            }

            var percent = Regex.Match(raw, @"^Percentage difference from Current Plan\s+(?<value>[\d.]+)%$", RegexOptions.IgnoreCase);
            if (percent.Success) pdf.MonthlyPremiumChangePercent = Value(percent, "value");
        }

        pdf.CurrentEhcPremium = SumPremium(pdf.RateLines, "Extended Health Care", current: true);
        pdf.RenewalEhcPremium = SumPremium(pdf.RateLines, "Extended Health Care", current: false);
        pdf.CurrentDentalPremium = SumPremium(pdf.RateLines, "Dental", current: true);
        pdf.RenewalDentalPremium = SumPremium(pdf.RateLines, "Dental", current: false);
    }

    private static void ParseAsoBudgetedBenefits(IEnumerable<PdfPageText> pages, PdfExtractionResult pdf)
    {
        var pageList = pages.ToList();
        var dentalArrangementPage = pageList.FirstOrDefault(page =>
            page.Text.Contains("Administrative Services Only Benefits", StringComparison.OrdinalIgnoreCase) &&
            page.Text.Contains("Dental Care", StringComparison.OrdinalIgnoreCase) &&
            Regex.IsMatch(page.Text, @"Underwriting\s+Arrangement\s*:\s*ASO\s*[\-–—]?\s*Budgeted", RegexOptions.IgnoreCase));

        if (dentalArrangementPage is null)
        {
            return;
        }

        var expensePage = pageList.FirstOrDefault(page =>
            page.Text.Contains("ASO Expenses", StringComparison.OrdinalIgnoreCase) &&
            page.Text.Contains("Dental Care Expenses", StringComparison.OrdinalIgnoreCase));
        var expenseMatch = expensePage is null
            ? Match.Empty
            : Regex.Match(
                expensePage.Text,
                @"Dental\s+Care\s+Expenses\s+(?<current>[\d.]+)\s*%\s+(?<renewal>[\d.]+)\s*%",
                RegexOptions.IgnoreCase);

        var sourcePage = expensePage ?? dentalArrangementPage;
        pdf.AsoBudgetedBenefits.Add(new AsoBudgetedBenefit
        {
            Benefit = "Dental",
            UnderwritingArrangement = "ASO - Budgeted",
            CurrentExpensePercent = expenseMatch.Success ? Value(expenseMatch, "current") : null,
            RenewalExpensePercent = expenseMatch.Success ? Value(expenseMatch, "renewal") : null,
            SourceReference = $"GroupHEALTH PDF page {sourcePage.PageNumber}"
        });

        foreach (var dentalRate in pdf.RateLines.Where(line =>
                     line.Benefit.Equals("Dental", StringComparison.OrdinalIgnoreCase)))
        {
            dentalRate.FundingArrangement = "ASO - Budgeted";
        }

        foreach (var healthRate in pdf.RateLines.Where(line =>
                     line.Benefit.Equals("Extended Health Care", StringComparison.OrdinalIgnoreCase)))
        {
            healthRate.FundingArrangement = "Insured";
        }
    }

    private static void AddTierRate(PdfExtractionResult pdf, int page, string benefit, string accountClass, Match match)
    {
        var lives = Int(match, "lives");
        var currentRate = Value(match, "current");
        var renewalRate = Value(match, "renewal");
        decimal? currentPremium = lives.HasValue && currentRate.HasValue ? lives.Value * currentRate.Value : null;
        decimal? renewalPremium = lives.HasValue && renewalRate.HasValue ? lives.Value * renewalRate.Value : null;
        pdf.RateLines.Add(new RateLine
        {
            Benefit = benefit,
            TierOrClass = ParsingHelpers.Clean(match.Groups["tier"].Value),
            AccountClass = accountClass,
            Lives = lives,
            CurrentRate = currentRate,
            RenewalRate = renewalRate,
            CurrentPremium = currentPremium,
            RenewalPremium = renewalPremium,
            PercentChange = Change(currentPremium, renewalPremium),
            SourceReference = $"GroupHEALTH PDF page {page}"
        });
    }

    private static void ParseRenewalExperience(IEnumerable<PdfPageText> pages, PdfExtractionResult pdf)
    {
        foreach (var page in pages.Where(candidate => candidate.Text.Contains("Claims Experience Period", StringComparison.OrdinalIgnoreCase)))
        {
            var benefit = page.Text.Contains("Extended Health Care", StringComparison.OrdinalIgnoreCase) ? "Extended Health Care" :
                page.Text.Contains("Dental Care", StringComparison.OrdinalIgnoreCase) ? "Dental" : "";
            if (benefit.Length == 0) continue;

            var lines = page.Lines.Select(ParsingHelpers.Clean).ToList();
            var from = FindTwoDates(lines, "Claims Experience Period From:");
            var to = FindTwoDates(lines, "To:");
            var periodCount = Math.Min(from.Count, to.Count);
            if (periodCount == 0) continue;

            var premiums = FindTwoMoney(lines, "Paid Premium");
            var paidClaims = FindTwoMoney(lines, "Paid Claims", exclude: "Net of Pooled");
            var incurredClaims = FindTwoMoney(lines, "B. Incurred Claims");
            var incurredRatios = FindTwoPercent(lines, "C. Incurred Loss Ratio");
            var trend = FindTwoPercent(lines, "D. Annual Trend");
            var weighting = FindTwoPercent(lines, "Weighting");
            var target = FindOnePercent(lines, "G. TLR") ??
                         FindOnePercent(lines, "F. TLR") ??
                         FindOnePercentContaining(lines, "TLR");
            var adjustment = FindOnePercent(lines, "Required Rate Change");
            for (var index = 0; index < periodCount; index++)
            {
                var period = $"{from[index]:MMMM d, yyyy} to {to[index]:MMMM d, yyyy}";
                var billed = At(premiums, index);
                var paid = At(paidClaims, index);
                pdf.ExperienceLines.Add(new ExperienceLine
                {
                    Benefit = benefit,
                    Period = period,
                    BilledPremium = billed,
                    PaidClaims = paid,
                    IncurredClaims = At(incurredClaims, index),
                    PaidLossRatio = UnderwritingEngine.CalculateLossRatioPercent(paid, billed),
                    IncurredLossRatio = At(incurredRatios, index),
                    TrendFactor = At(trend, index),
                    Weighting = At(weighting, index),
                    TargetLossRatio = target,
                    ProposedAdjustment = adjustment,
                    SourceReference = $"GroupHEALTH PDF page {page.PageNumber}"
                });
            }

            if (benefit.Equals("Extended Health Care", StringComparison.OrdinalIgnoreCase))
            {
                var pooled = FindTwoMoney(lines, "Pooled Premium");
                pdf.PooledPremium ??= At(pooled, 0);
                if (pdf.PooledPremium.HasValue && At(premiums, 0).GetValueOrDefault() > 0m)
                {
                    pdf.PoolingCharge ??= Math.Round(pdf.PooledPremium.Value / At(premiums, 0)!.Value * 100m, 2);
                }
            }
        }
    }

    private static void ParseClaimsCompanionPages(IEnumerable<PdfPageText> pages, PdfExtractionResult pdf, WorkbookExtractionResult workbook)
    {
        var aggregateCategoryReports = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var page in pages)
        {
            var lines = page.Lines.Select(ParsingHelpers.Clean).ToList();
            var period = ParseClaimsPeriod(lines);
            if (period.Length == 0) continue;
            var benefit = IsDentalClaimsPage(pdf.SourcePath, page) ? "Dental" : "Extended Health Care";

            if (page.Text.Contains("Billed Premium & Claims", StringComparison.OrdinalIgnoreCase) ||
                IsMonthlyClaimsExperiencePage(page.Text))
            {
                var total = lines.Select(line => Regex.Match(line, @"^Grand Total\s+\$(?<premium>[\d,]+\.\d{2})\s+\$(?<incurred>[\d,]+\.\d{2})", RegexOptions.IgnoreCase)).FirstOrDefault(match => match.Success);
                if (total is not null)
                {
                    var billed = Value(total, "premium");
                    var paid = Value(total, "incurred");
                    if (!pdf.ExperienceLines.Any(row => Same(row.Benefit, benefit) && Same(row.Period, period)))
                    {
                        pdf.ExperienceLines.Add(new ExperienceLine
                        {
                            Benefit = benefit,
                            Period = period,
                            BilledPremium = billed,
                            PaidClaims = paid,
                            IncurredClaims = paid,
                            PaidLossRatio = UnderwritingEngine.CalculateLossRatioPercent(paid, billed),
                            IncurredLossRatio = UnderwritingEngine.CalculateLossRatioPercent(paid, billed),
                            SourceReference = $"GroupHEALTH claims PDF page {page.PageNumber}"
                        });
                    }
                }
                continue;
            }

            var target = page.Text.Contains("Paramedical Information", StringComparison.OrdinalIgnoreCase) ? workbook.ParamedicalServiceTrends :
                page.Text.Contains("Cost Category", StringComparison.OrdinalIgnoreCase) ? workbook.DentalSpendingTrends :
                page.Text.Contains("Expense Information", StringComparison.OrdinalIgnoreCase) ? workbook.EhcSpendingTrends : null;
            if (target is null) continue;

            var reportType = ReferenceEquals(target, workbook.ParamedicalServiceTrends) ? "Paramedical" :
                ReferenceEquals(target, workbook.DentalSpendingTrends) ? "Dental" : "EHC";
            var reportKey = $"{reportType}|{benefit}|{period}";
            var isPolicyDetail = lines.Any(line => line.StartsWith("Current Policy ", StringComparison.OrdinalIgnoreCase));
            if (isPolicyDetail && aggregateCategoryReports.Contains(reportKey)) continue;
            if (!isPolicyDetail) aggregateCategoryReports.Add(reportKey);

            foreach (var raw in lines)
            {
                var match = CategoryMoneyRegex.Match(raw);
                if (!match.Success || raw.StartsWith("Grand Total", StringComparison.OrdinalIgnoreCase)) continue;
                var category = NormalizeCategory(match.Groups["category"].Value);
                var amount = Value(match, "amount") ?? 0m;
                var source = $"GroupHEALTH claims PDF page {page.PageNumber}";
                AddCategoryPeriod(target, benefit, category, period, amount, source, sumWithExisting: isPolicyDetail);
                AddClaimTrendPeriod(workbook.ClaimTrends, benefit, category, period, amount, source, sumWithExisting: isPolicyDetail);
            }
        }
    }

    private static bool IsMonthlyClaimsExperiencePage(string text) =>
        text.Contains("Billed Premium", StringComparison.OrdinalIgnoreCase) &&
        text.Contains("Paid Claims", StringComparison.OrdinalIgnoreCase);

    private static bool IsDentalClaimsPage(string sourcePath, PdfPageText page)
    {
        var text = NormalizePdfLigatures(page.Text);
        return !text.Contains("Benefit EXTENDED HEALTH CARE", StringComparison.OrdinalIgnoreCase) &&
               (text.Contains("Benefit DENTAL", StringComparison.OrdinalIgnoreCase) ||
                text.Contains("Cost Category", StringComparison.OrdinalIgnoreCase) ||
                IsDentalOnlyClaimsFile(sourcePath));
    }

    private static string NormalizePdfLigatures(string value) =>
        value.Replace("ﬁ", "fi", StringComparison.Ordinal)
             .Replace("ﬂ", "fl", StringComparison.Ordinal);

    private static bool IsDentalOnlyClaimsFile(string sourcePath)
    {
        var fileName = Path.GetFileNameWithoutExtension(sourcePath);
        return fileName.Contains("DEN", StringComparison.OrdinalIgnoreCase) &&
               !fileName.Contains("EHL", StringComparison.OrdinalIgnoreCase) &&
               !fileName.Contains("EHC", StringComparison.OrdinalIgnoreCase);
    }

    private static void AddCategoryPeriod(
        List<ServiceCategoryTrend> rows,
        string benefit,
        string category,
        string period,
        decimal amount,
        string source,
        bool sumWithExisting = false)
    {
        var row = rows.FirstOrDefault(candidate => Same(candidate.Benefit, benefit) && Same(candidate.Category, category));
        if (row is null)
        {
            row = new ServiceCategoryTrend { Benefit = benefit, Category = category, SourceReference = source };
            rows.Add(row);
        }
        else if (sumWithExisting)
        {
            amount += ServiceCategoryHistory.GetPeriods(row, int.MaxValue)
                .FirstOrDefault(existing => Same(existing.PeriodLabel, period))?.Amount ?? 0m;
        }
        ServiceCategoryHistory.Upsert(row, period, amount, source);
    }

    private static void AddClaimTrendPeriod(
        List<ClaimTrend> rows,
        string benefit,
        string category,
        string period,
        decimal amount,
        string source,
        bool sumWithExisting)
    {
        var row = rows.FirstOrDefault(candidate =>
            Same(candidate.Benefit, benefit) && Same(candidate.Category, category) && Same(candidate.Period, period));
        if (row is null)
        {
            rows.Add(new ClaimTrend
            {
                Benefit = benefit,
                Category = category,
                Period = period,
                Amount = amount,
                SourceReference = source
            });
            return;
        }

        row.Amount = sumWithExisting ? row.Amount.GetValueOrDefault() + amount : amount;
        row.SourceReference = source;
    }

    private static void ParseTopDrugs(IEnumerable<PdfPageText> pages, WorkbookExtractionResult workbook)
    {
        foreach (var page in pages.Where(candidate => candidate.Text.Contains("Top 10 DINs by Dollar Amount", StringComparison.OrdinalIgnoreCase)))
        {
            var periodMatch = Regex.Match(page.Text, @"From\s+(?<from>[A-Za-z]+\s+\d{1,2},\s+\d{4})\s+-\s+(?<to>[A-Za-z]+\s+\d{1,2},\s+\d{4})", RegexOptions.IgnoreCase);
            var period = periodMatch.Success ? $"{periodMatch.Groups["from"].Value} to {periodMatch.Groups["to"].Value}" : "";
            foreach (var line in page.Lines.Select(ParsingHelpers.Clean))
            {
                var match = DrugRowRegex.Match(line);
                if (!match.Success || match.Groups["name"].Value.Equals("Other", StringComparison.OrdinalIgnoreCase)) continue;
                workbook.DrugTrends.Add(new DrugTrend { Din = match.Groups["din"].Value, DrugName = ParsingHelpers.Clean(match.Groups["name"].Value), Period = period, PrescriptionCount = Int(match, "claims"), NetAmountPaid = Value(match, "paid"), SourceReference = $"GroupHEALTH drug PDF page {page.PageNumber}" });
            }
        }
    }

    private static void ParsePoolingSupport(IEnumerable<PdfPageText> pages, PdfExtractionResult pdf)
    {
        var text = string.Join(" ", pages.Select(page => page.Text));
        var threshold = Regex.Match(text, @"\$(?<value>[\d,]+(?:\.\d{2})?)\s+(?:per insured person|pooling level|EP3)", RegexOptions.IgnoreCase);
        if (!threshold.Success) threshold = Regex.Match(text, @"pooling level\s+(?:of\s+)?\$(?<value>[\d,]+(?:\.\d{2})?)", RegexOptions.IgnoreCase);
        if (threshold.Success) pdf.PoolingThreshold ??= Value(threshold, "value");
    }

    private static void ParseDemographics(IEnumerable<PdfPageText> pages, PdfExtractionResult pdf)
    {
        foreach (var page in pages.Where(candidate => candidate.Text.Contains("Age & Volume Distribution", StringComparison.OrdinalIgnoreCase)))
        {
            var benefit = page.Text.Contains("LTD Age", StringComparison.OrdinalIgnoreCase) ? "Long-Term Disability" : "Life";
            var lines = page.Lines.Select(ParsingHelpers.Clean).ToList();
            var priorStart = lines.FindIndex(line => line.StartsWith("PRIOR:", StringComparison.OrdinalIgnoreCase));
            var currentStart = lines.FindIndex(line => line.StartsWith("CURRENT:", StringComparison.OrdinalIgnoreCase));
            if (priorStart < 0 || currentStart < 0) continue;
            var prior = ParseDemographicBlock(lines.Skip(priorStart + 1).Take(currentStart - priorStart - 1).ToList());
            var current = ParseDemographicBlock(lines.Skip(currentStart + 1).ToList());
            var bands = prior.Keys.Concat(current.Keys).Distinct(StringComparer.OrdinalIgnoreCase);
            foreach (var age in bands)
            {
                prior.TryGetValue(age, out var p);
                current.TryGetValue(age, out var c);
                var normalized = benefit.Equals("Life", StringComparison.OrdinalIgnoreCase) ? DemographicAgeBandNormalizer.NormalizeLifeAgeBand(age) : DemographicAgeBandNormalizer.NormalizeLtdAgeBand(age);
                var target = benefit.Equals("Life", StringComparison.OrdinalIgnoreCase) ? pdf.LifeDemographics : pdf.LtdDemographics;
                target.Add(new DemographicLine { Benefit = benefit, AgeBand = normalized, PriorMaleLives = p?.MaleLives, PriorMaleVolume = p?.MaleVolume, PriorFemaleLives = p?.FemaleLives, PriorFemaleVolume = p?.FemaleVolume, CurrentMaleLives = c?.MaleLives, CurrentMaleVolume = c?.MaleVolume, CurrentFemaleLives = c?.FemaleLives, CurrentFemaleVolume = c?.FemaleVolume, SourceReference = $"GroupHEALTH PDF page {page.PageNumber}" });
            }
        }
    }

    private static Dictionary<string, DemographicValues> ParseDemographicBlock(IReadOnlyList<string> lines)
    {
        var result = new Dictionary<string, DemographicValues>(StringComparer.OrdinalIgnoreCase);
        var totalIndex = lines.ToList().FindIndex(line => line.StartsWith("TOTAL ", StringComparison.OrdinalIgnoreCase));
        var totalLine = totalIndex >= 0 ? lines[totalIndex] : "";
        var totals = ParsingHelpers.ExtractNumbers(totalLine["TOTAL".Length..]);
        var expectedMaleLives = totals.Count >= 6 ? (int)totals[0] : 0;
        var expectedMaleVolume = totals.Count >= 6 ? totals[1] : 0m;
        var expectedFemaleLives = totals.Count >= 6 ? (int)totals[2] : 0;
        var expectedFemaleVolume = totals.Count >= 6 ? totals[3] : 0m;
        var ambiguous = new List<(string Age, int Lives, decimal Volume)>();
        var tableLines = totalIndex >= 0 ? lines.Take(totalIndex) : lines;
        foreach (var line in tableLines)
        {
            var match = AgeRowRegex.Match(line);
            if (!match.Success) continue;
            var values = ParsingHelpers.ExtractNumbers(match.Groups["values"].Value);
            if (values.Count == 6)
            {
                result[match.Groups["age"].Value] = new DemographicValues((int)values[0], values[1], (int)values[2], values[3]);
            }
            else if (values.Count == 4)
            {
                ambiguous.Add((match.Groups["age"].Value, (int)values[0], values[1]));
            }
            else if (values.Count == 0)
            {
                result[match.Groups["age"].Value] = new DemographicValues(null, null, null, null);
            }
        }

        var fixedMaleLives = result.Values.Sum(value => value.MaleLives ?? 0);
        var fixedMaleVolume = result.Values.Sum(value => value.MaleVolume ?? 0m);
        var fixedFemaleLives = result.Values.Sum(value => value.FemaleLives ?? 0);
        var fixedFemaleVolume = result.Values.Sum(value => value.FemaleVolume ?? 0m);
        var selectedMask = 0;
        for (var mask = 0; mask < (1 << ambiguous.Count); mask++)
        {
            var maleLives = fixedMaleLives; var maleVolume = fixedMaleVolume; var femaleLives = fixedFemaleLives; var femaleVolume = fixedFemaleVolume;
            for (var i = 0; i < ambiguous.Count; i++)
            {
                if ((mask & (1 << i)) != 0) { maleLives += ambiguous[i].Lives; maleVolume += ambiguous[i].Volume; }
                else { femaleLives += ambiguous[i].Lives; femaleVolume += ambiguous[i].Volume; }
            }
            if (maleLives == expectedMaleLives && maleVolume == expectedMaleVolume && femaleLives == expectedFemaleLives && femaleVolume == expectedFemaleVolume) { selectedMask = mask; break; }
        }
        for (var i = 0; i < ambiguous.Count; i++)
        {
            var item = ambiguous[i];
            result[item.Age] = (selectedMask & (1 << i)) != 0 ? new DemographicValues(item.Lives, item.Volume, null, null) : new DemographicValues(null, null, item.Lives, item.Volume);
        }
        return result;
    }

    private static List<DateOnly> FindTwoDates(IEnumerable<string> lines, string startsWith)
    {
        var line = lines.FirstOrDefault(candidate => candidate.StartsWith(startsWith, StringComparison.OrdinalIgnoreCase)) ?? "";
        return Regex.Matches(line, @"[A-Za-z]+\s+\d{1,2},\s+\d{4}").Select(match => ParsingHelpers.ParseDateOnly(match.Value)).Where(date => date.HasValue).Select(date => date!.Value).Take(2).ToList();
    }
    private static List<decimal> FindTwoMoney(IEnumerable<string> lines, string startsWith, string exclude = "") => FindNumbers(lines, startsWith, exclude, @"\(?\$?[\d,]+\.\d{2}\)?");
    private static List<decimal> FindTwoPercent(IEnumerable<string> lines, string startsWith) => FindNumbers(lines, startsWith, "", @"\(?-?[\d.]+%\)?");
    private static decimal? FindOnePercent(IEnumerable<string> lines, string startsWith)
    {
        var values = FindTwoPercent(lines, startsWith);
        return values.Count > 0 ? values[0] : null;
    }
    private static decimal? FindOnePercentContaining(IEnumerable<string> lines, string label)
    {
        var line = lines.FirstOrDefault(candidate => candidate.Contains(label, StringComparison.OrdinalIgnoreCase)) ?? "";
        return Regex.Matches(line, @"\(?-?[\d.]+%\)?")
            .Select(match => ParsingHelpers.ParseDecimal(match.Value))
            .FirstOrDefault(value => value.HasValue);
    }
    private static List<decimal> FindNumbers(IEnumerable<string> lines, string startsWith, string exclude, string pattern)
    {
        var line = lines.FirstOrDefault(candidate => candidate.StartsWith(startsWith, StringComparison.OrdinalIgnoreCase) && (exclude.Length == 0 || !candidate.Contains(exclude, StringComparison.OrdinalIgnoreCase))) ?? "";
        return Regex.Matches(line, pattern).Select(match => ParsingHelpers.ParseDecimal(match.Value)).Where(value => value.HasValue).Select(value => value!.Value).Take(2).ToList();
    }

    private static string ParseClaimsPeriod(IReadOnlyList<string> lines)
    {
        var from = lines.FirstOrDefault(line => line.StartsWith("Date From ", StringComparison.OrdinalIgnoreCase))?["Date From ".Length..];
        var to = lines.FirstOrDefault(line => line.StartsWith("Date To ", StringComparison.OrdinalIgnoreCase))?["Date To ".Length..];
        if (!DateTime.TryParseExact(from, "MMMM yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var start) || !DateTime.TryParseExact(to, "MMMM yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var end)) return "";
        return $"{start:MMMM} 1, {start:yyyy} to {end:MMMM} {DateTime.DaysInMonth(end.Year, end.Month)}, {end:yyyy}";
    }
    private static string ParseClientName(IEnumerable<PdfPageText> pages)
    {
        var lines = pages.SelectMany(page => page.Lines).Select(ParsingHelpers.Clean).ToList();
        var policyClient = lines
            .Select(line => Regex.Match(line, @"^Policy\s*(?:#|Number|No\.?)?\s*[:\-]?\s*\d+\s*-\s*(?<client>.+)$", RegexOptions.IgnoreCase))
            .FirstOrDefault(match => match.Success);
        if (policyClient is not null)
        {
            return ParsingHelpers.Clean(policyClient.Groups["client"].Value);
        }

        var group = lines.FirstOrDefault(line => line.StartsWith("Current Group ", StringComparison.OrdinalIgnoreCase));
        if (group is not null)
        {
            return Regex.Replace(group["Current Group ".Length..], @"\s+\(\d+\)$", "", RegexOptions.IgnoreCase);
        }

        var first = pages.FirstOrDefault()?.Lines.Select(ParsingHelpers.Clean).ToList() ?? [];
        return first
            .SkipWhile(line => !line.Equals("Group Benefits Plan Renewal", StringComparison.OrdinalIgnoreCase))
            .Skip(1)
            .FirstOrDefault() ?? "";
    }
    private static string ParsePolicyNumber(string path, IEnumerable<PdfPageText> pages)
    {
        var lines = pages.SelectMany(page => page.Lines).Select(ParsingHelpers.Clean).ToList();
        var policy = lines
            .Select(line => Regex.Match(line, @"^Policy\s*(?:#|Number|No\.?)?\s*[:\-]?\s*(?<policy>\d+)\b", RegexOptions.IgnoreCase))
            .FirstOrDefault(match => match.Success);
        if (policy is not null)
        {
            return policy.Groups["policy"].Value;
        }

        var claimsFilePolicy = Regex.Match(Path.GetFileNameWithoutExtension(path), @"^(?<policy>\d{4,})(?:\(\k<policy>\))?__", RegexOptions.IgnoreCase);
        if (claimsFilePolicy.Success)
        {
            return claimsFilePolicy.Groups["policy"].Value;
        }

        var group = lines.FirstOrDefault(line => line.StartsWith("Current Group ", StringComparison.OrdinalIgnoreCase));
        var match = Regex.Match(group ?? "", @"\((?<policy>\d+)\)$");
        if (match.Success) return match.Groups["policy"].Value;
        // Broker renewal filenames often contain only the effective year. Returning no
        // policy is safer than treating that year as a conflicting client identity.
        return Regex.Matches(Path.GetFileNameWithoutExtension(path), @"\d+")
            .Select(item => item.Value)
            .Where(value => !IsCalendarYearToken(value))
            .LastOrDefault() ?? "";
    }

    private static bool IsCalendarYearToken(string value) =>
        value.Length == 4 &&
        int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out var year) &&
        year is >= 1900 and <= 2100;
    private static DateOnly? ParseRenewalDate(IEnumerable<PdfPageText> pages)
    {
        foreach (var line in pages.SelectMany(page => page.Lines).Select(ParsingHelpers.Clean))
        {
            var match = Regex.Match(line, @"\b(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4}\b");
            if (match.Success && ParsingHelpers.ParseDateOnly(match.Value) is { } date) return date;
        }
        return null;
    }

    private static decimal? SumPremium(IEnumerable<RateLine> rows, string benefit, bool current) => rows.Where(row => row.Benefit.Equals(benefit, StringComparison.OrdinalIgnoreCase)).Sum(row => current ? row.CurrentPremium : row.RenewalPremium);
    private static decimal? Change(decimal? current, decimal? renewal) => !current.HasValue || current.Value == 0m || !renewal.HasValue ? null : Math.Round((renewal.Value - current.Value) / current.Value * 100m, 1);
    private static decimal? Value(Match match, string name) => ParsingHelpers.ParseDecimal(match.Groups[name].Value);
    private static int? Int(Match match, string name) => ParsingHelpers.ParseInt(match.Groups[name].Value);
    private static decimal? At(IReadOnlyList<decimal> values, int index) => index < values.Count ? values[index] : null;
    private static string NormalizeBenefit(string value) => value.Trim() switch { var v when v.Equals("Accident & Serious Illness (ASI)", StringComparison.OrdinalIgnoreCase) => "AD&D", var v when v.Equals("Long Term Disability", StringComparison.OrdinalIgnoreCase) => "LTD", var v when v.Equals("Dental Care", StringComparison.OrdinalIgnoreCase) => "Dental", var v when v.Equals("Life Insurance", StringComparison.OrdinalIgnoreCase) => "Life", _ => ParsingHelpers.Clean(value) };
    private static string NormalizeCategory(string value) => ParsingHelpers.Clean(value).ToUpperInvariant() switch { "DRUGS" => "Drugs", "MEDICAL APPLIANCES / SUPPLIES" => "Medical Services and Supplies", "PARAMEDICALS" => "Paramedicals", "VISION CARE" => "Vision", "BASIC SERVICES" => "Basic", "MASSAGE THERAPIST" => "Massage Therapist", "PHYSIOTHERAPY" => "Physiotherapy", var other => CultureInfo.InvariantCulture.TextInfo.ToTitleCase(other.ToLowerInvariant()) };
    private static bool HasWorkbookData(WorkbookExtractionResult workbook) => workbook.ClaimTrends.Count > 0 || workbook.EhcSpendingTrends.Count > 0 || workbook.DentalSpendingTrends.Count > 0 || workbook.ParamedicalServiceTrends.Count > 0 || workbook.DrugTrends.Count > 0;
    private static bool Same(string left, string right) => ParsingHelpers.Clean(left).Equals(ParsingHelpers.Clean(right), StringComparison.OrdinalIgnoreCase);
    private static bool ContainsAny(string text, params string[] values) => values.Any(value => text.Contains(value, StringComparison.OrdinalIgnoreCase));
    private sealed record DemographicValues(int? MaleLives, decimal? MaleVolume, int? FemaleLives, decimal? FemaleVolume);
}

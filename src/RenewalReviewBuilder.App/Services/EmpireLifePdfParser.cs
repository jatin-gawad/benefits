using System.Globalization;
using System.Text.RegularExpressions;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed class EmpireLifePdfParser
{
    private static readonly Regex PolicyNumberRegex = new(
        @"(?:Group\s+)?Policy Number:\s*(?<value>[A-Z0-9-]+)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex PolicyNameRegex = new(
        @"(?:Group\s+)?Policy Name:\s*(?<value>.+)$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex GroupNumberRegex = new(
        @"^Group Number\s+(?<value>[A-Z0-9-]+)$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex GroupNameRegex = new(
        @"^Group Name\s+(?<value>.+)$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex RenewalDateRegex = new(
        @"(?:Renewal\s+)?Effective Date:\s*(?<value>[A-Za-z]+\s+\d{1,2},\s+\d{4})",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex TotalPremiumRegex = new(
        @"^TOTAL\s+\$(?<current>[\d,]+\.\d{2})\s+\$(?<renewal>[\d,]+\.\d{2})\s+\*?(?<change>[+-]?[\d.]+)%$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex RateRowRegex = new(
        @"^(?<benefit>Accidental Death\s*&\s*Dismemberment|Dependant Life|Dependent Life|Extended Health|Dental|Life)\s+(?<class>.+?)\s+(?:(?<tier>Single Parent|Single|Couple|Family)\s+)?(?<volume>[\d,]+)\s+\$(?<currentRate>[\d.]+)\s+\$(?<currentPremium>[\d,]+\.\d{2})\s+\$(?<renewalRate>[\d.]+)\s+\$(?<renewalPremium>[\d,]+\.\d{2})\s+(?<change>[+-]?[\d.]+)%$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex PooledBenefitRegex = new(
        @"^(?<benefit>Accidental Death\s*&\s*Dismemberment|Dependant Life|Dependent Life|Healthcare Pooling|Extended Health\s*-\s*Out of Country|Life)\s+\$(?<premium>[\d,]+\.\d{2})\s+\$(?<claims>[\d,]+\.\d{2})$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex PoolingThresholdRegex = new(
        @"pooling threshold is \$(?<value>[\d,]+(?:\.\d{2})?) per insured",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex RenewalExperienceRowRegex = new(
        @"^(?<period>[A-Za-z]{3}\s+\d{4}\s*-\s*[A-Za-z]{3}\s+\d{4})\s+\$(?<nonPooled>[\d,]+\.\d{2})\s+\$(?<adjustedPremium>[\d,]+\.\d{2})\s+\$(?<paid>[\d,]+\.\d{2})\s+\$(?<trended>[\d,]+\.\d{2})\s+\$(?<incurred>[\d,]+\.\d{2})\s+(?<ratio>[\d.]+)%$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex SummaryPeriodHeaderRegex = new(
        @"(?<currentStart>\d{2}\s+[A-Za-z]{3}\s+\d{4})\s*-\s*(?<currentEnd>\d{2}\s+[A-Za-z]{3}\s+\d{4})\s+(?<priorStart>\d{2}\s+[A-Za-z]{3}\s+\d{4})\s*-\s*(?<priorEnd>\d{2}\s+[A-Za-z]{3}\s+\d{4})",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex SummaryExperienceRowRegex = new(
        @"^(?<benefit>Extended Health(?:\s+\*\*)?|Dental(?:\s+\*\*)?)\s+(?<currentPremium>[\d,]+(?:\.\d+)?)\s+(?<currentPaid>[\d,]+(?:\.\d+)?)\s+(?<currentIncurred>[\d,]+(?:\.\d+)?)\s+(?<currentRatio>[+-]?[\d.]+)\s+(?<priorPremium>[\d,]+(?:\.\d+)?)\s+(?<priorPaid>[\d,]+(?:\.\d+)?)\s+(?<priorIncurred>[\d,]+(?:\.\d+)?)\s+(?<priorRatio>[+-]?[\d.]+)$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex ExperienceSummaryRegex = new(
        @"^(?<benefit>Extended Health|Dental)\s+(?<weighted>[\d.]+)%\s+(?<target>[\d.]+)%\s+(?<credibility>[\d.]+)%$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex ExperienceWeightingRegex = new(
        @"using a\s+(?<current>\d+)\s*/\s*(?<prior>\d+)\s+weighting",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex TrendRegex = new(
        @"annual trend adjustments for Extended Health and Dental are\s+(?<health>[\d.]+)%\s+and\s+(?<dental>[\d.]+)%",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex DemographicRowRegex = new(
        @"^(?<age><25|\d{2}-\d{2}|65\+)\s+(?<currentMaleLives>\d+)\s+(?<currentMaleVolume>[\d,]+)\s+(?<currentFemaleLives>\d+)\s+(?<currentFemaleVolume>[\d,]+)\s+(?<priorMaleLives>\d+)\s+(?<priorMaleVolume>[\d,]+)\s+(?<priorFemaleLives>\d+)\s+(?<priorFemaleVolume>[\d,]+)$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex UtilizationRowRegex = new(
        @"^(?<category>.+?)\s+(?<currentOccurrences>\d+)\s+(?<currentPaid>[\d,]+(?:\.\d+)?)\s+(?<currentPercent>[\d.]+)\s+(?<priorOccurrences>\d+)\s+(?<priorPaid>[\d,]+(?:\.\d+)?)\s+(?<priorPercent>[\d.]+)$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex UtilizationCompactRowRegex = new(
        @"^(?<category>.+?)\s+(?<currentOccurrences>\d+)\s+(?<currentPaid>[\d,]+(?:\.\d+)?)\s+(?<priorOccurrences>\d+)\s+(?<priorPaid>[\d,]+(?:\.\d+)?)$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex UtilizationPeriodHeaderRegex = new(
        @"(?<currentStart>\d{2}\s+[A-Za-z]{3}\s+\d{4})\s+To\s+(?<currentEnd>\d{2}\s+[A-Za-z]{3}\s+\d{4})\s+(?<priorStart>\d{2}\s+[A-Za-z]{3}\s+\d{4})\s+To\s+(?<priorEnd>\d{2}\s+[A-Za-z]{3}\s+\d{4})",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex ReportCreatedDateRegex = new(
        @"(?<value>[A-Za-z]+\s+\d{1,2},\s+\d{4})\s+Page\s+\d+\s+of\s+\d+",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public static bool IsEmpireLifeDocument(string pdfPath, string text)
    {
        if (text.Contains("Empire Life", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("group.advisor@empire.ca", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        var searchable = $"{Path.GetFileNameWithoutExtension(pdfPath)} {text}";
        return searchable.Contains("Group Renewal Summary", StringComparison.OrdinalIgnoreCase) ||
               (searchable.Contains("Group Policy Name:", StringComparison.OrdinalIgnoreCase) &&
                searchable.Contains("Group Policy Number:", StringComparison.OrdinalIgnoreCase) &&
                (searchable.Contains("Health Claims by Major Medical", StringComparison.OrdinalIgnoreCase) ||
                 searchable.Contains("Age Volume Distribution - Life", StringComparison.OrdinalIgnoreCase) ||
                 searchable.Contains("Dental Claims Summary", StringComparison.OrdinalIgnoreCase))) ||
               (searchable.Contains("Group Name", StringComparison.OrdinalIgnoreCase) &&
                searchable.Contains("Group Number", StringComparison.OrdinalIgnoreCase) &&
                searchable.Contains("Health Claims by Major Medical", StringComparison.OrdinalIgnoreCase));
    }

    public CarrierPdfParseResult Parse(string pdfPath, List<PdfPageText> pages, string allText)
    {
        var pdf = new PdfExtractionResult
        {
            SourcePath = pdfPath,
            Carrier = "Empire Life",
            ClientName = FirstNonBlank(ParseLabeledValue(pages, PolicyNameRegex), ParseLabeledValue(pages, GroupNameRegex)),
            PolicyNumber = FirstNonBlank(ParseLabeledValue(pages, PolicyNumberRegex), ParseLabeledValue(pages, GroupNumberRegex)),
            RenewalDate = ParseRenewalDate(pages)
        };

        ParseRates(pages, pdf);
        ParsePooledBenefits(pages, pdf);
        ParseRenewalExperience(pages, allText, pdf);
        ParseSummaryExperience(pages, pdf);
        ParseDemographics(pages, pdf);
        var workbook = ParseUtilization(pdfPath, pages, pdf);

        if (pdf.RateLines.Count == 0 &&
            pdf.ExperienceLines.Count == 0 &&
            pdf.LifeDemographics.Count == 0 &&
            workbook.ClaimsDistributionRows.Count == 0)
        {
            pdf.Warnings.Add("No Empire Life renewal, experience, demographic, or utilization values were extracted from the PDF.");
        }

        return new CarrierPdfParseResult
        {
            Carrier = "Empire Life",
            Pdf = pdf,
            WorkbookFromPdf = HasWorkbookData(workbook) ? workbook : null
        };
    }

    private static void ParseRates(IEnumerable<PdfPageText> pages, PdfExtractionResult result)
    {
        foreach (var page in pages.Where(page =>
                     page.Text.Contains("PROPOSED RENEWAL RATES", StringComparison.OrdinalIgnoreCase)))
        {
            foreach (var line in page.Lines.Select(ParsingHelpers.Clean))
            {
                var total = TotalPremiumRegex.Match(line);
                if (total.Success)
                {
                    result.CurrentMonthlyPremium = Money(total, "current");
                    result.RenewalMonthlyPremium = Money(total, "renewal");
                    result.MonthlyPremiumChange = result.RenewalMonthlyPremium - result.CurrentMonthlyPremium;
                    result.MonthlyPremiumChangePercent = Money(total, "change");
                    continue;
                }

                var match = RateRowRegex.Match(line);
                if (!match.Success)
                {
                    continue;
                }

                var benefit = NormalizeBenefit(match.Groups["benefit"].Value);
                var tier = ParsingHelpers.Clean(match.Groups["tier"].Value);
                var countOrVolume = Money(match, "volume");
                var isLives = tier.Length > 0 || benefit.Equals("Dependent Life", StringComparison.OrdinalIgnoreCase);
                result.RateLines.Add(new RateLine
                {
                    Benefit = benefit,
                    AccountClass = ParsingHelpers.Clean(match.Groups["class"].Value),
                    TierOrClass = tier,
                    Lives = isLives && countOrVolume.HasValue ? decimal.ToInt32(countOrVolume.Value) : null,
                    Volume = isLives ? null : countOrVolume,
                    CurrentRate = Money(match, "currentRate"),
                    CurrentPremium = Money(match, "currentPremium"),
                    RenewalRate = Money(match, "renewalRate"),
                    RenewalPremium = Money(match, "renewalPremium"),
                    PercentChange = Money(match, "change"),
                    SourceReference = $"Empire Life PDF page {page.PageNumber}"
                });
            }
        }

        result.CurrentEhcPremium = SumPremium(result.RateLines, "Extended Health", current: true);
        result.RenewalEhcPremium = SumPremium(result.RateLines, "Extended Health", current: false);
        result.CurrentDentalPremium = SumPremium(result.RateLines, "Dental", current: true);
        result.RenewalDentalPremium = SumPremium(result.RateLines, "Dental", current: false);
    }

    private static void ParsePooledBenefits(IEnumerable<PdfPageText> pages, PdfExtractionResult result)
    {
        foreach (var page in pages.Where(page =>
                     page.Text.Contains("POOLED BENEFITS", StringComparison.OrdinalIgnoreCase)))
        {
            foreach (var line in page.Lines.Select(ParsingHelpers.Clean))
            {
                var match = PooledBenefitRegex.Match(line);
                if (!match.Success)
                {
                    continue;
                }

                var row = new EmpireLifePooledBenefit
                {
                    Benefit = NormalizeBenefit(match.Groups["benefit"].Value),
                    BilledPremium = Money(match, "premium"),
                    PaidClaims = Money(match, "claims"),
                    SourceReference = $"Empire Life PDF page {page.PageNumber}"
                };
                result.EmpireLifePooledBenefits.Add(row);
                if (row.Benefit.Equals("Healthcare Pooling", StringComparison.OrdinalIgnoreCase))
                {
                    result.PooledPremium = row.BilledPremium;
                }
            }
        }

        result.PoolingThreshold = Money(PoolingThresholdRegex.Match(string.Join(" ", pages.Select(page => page.Text))), "value");
    }

    private static void ParseRenewalExperience(List<PdfPageText> pages, string allText, PdfExtractionResult result)
    {
        var summaries = ParseExperienceSummaries(pages);
        var weighting = ExperienceWeightingRegex.Match(allText);
        var trend = TrendRegex.Match(allText);

        foreach (var page in pages.Where(page =>
                     page.Text.Contains("EXPERIENCE RATED BENEFITS", StringComparison.OrdinalIgnoreCase)))
        {
            var currentBenefit = "";
            var benefitRows = new Dictionary<string, List<ExperienceLine>>(StringComparer.OrdinalIgnoreCase);
            foreach (var line in page.Lines.Select(ParsingHelpers.Clean))
            {
                if (line.Equals("Extended Health", StringComparison.OrdinalIgnoreCase) ||
                    line.Equals("Dental", StringComparison.OrdinalIgnoreCase))
                {
                    currentBenefit = NormalizeBenefit(line);
                    continue;
                }

                var match = RenewalExperienceRowRegex.Match(line);
                if (!match.Success || currentBenefit.Length == 0)
                {
                    continue;
                }

                summaries.TryGetValue(currentBenefit, out var summary);
                var adjustedPremium = Money(match, "adjustedPremium");
                var paidClaims = Money(match, "paid");
                var generic = new ExperienceLine
                {
                    Benefit = currentBenefit,
                    Period = NormalizePeriod(match.Groups["period"].Value),
                    BilledPremium = adjustedPremium,
                    PaidClaims = paidClaims,
                    IncurredClaims = Money(match, "incurred"),
                    PaidLossRatio = LossRatio(adjustedPremium, paidClaims),
                    IncurredLossRatio = Money(match, "ratio"),
                    TrendFactor = currentBenefit.Equals("Dental", StringComparison.OrdinalIgnoreCase)
                        ? Money(trend, "dental")
                        : Money(trend, "health"),
                    Credibility = summary?.Credibility,
                    TargetLossRatio = summary?.TargetLossRatio,
                    SourceReference = $"Empire Life PDF page {page.PageNumber}"
                };
                result.ExperienceLines.Add(generic);
                if (!benefitRows.TryGetValue(currentBenefit, out var rows))
                {
                    rows = [];
                    benefitRows[currentBenefit] = rows;
                }
                rows.Add(generic);

                result.EmpireLifeExperienceRows.Add(new EmpireLifeExperienceRow
                {
                    Benefit = currentBenefit,
                    Period = generic.Period,
                    NonPooledPremium = Money(match, "nonPooled"),
                    AdjustedPremium = adjustedPremium,
                    PaidClaims = paidClaims,
                    TrendedClaims = Money(match, "trended"),
                    IncurredClaims = generic.IncurredClaims,
                    IncurredLossRatio = generic.IncurredLossRatio,
                    WeightedLossRatio = summary?.WeightedLossRatio,
                    TargetLossRatio = summary?.TargetLossRatio,
                    Credibility = summary?.Credibility,
                    SourceReference = generic.SourceReference
                });
            }

            foreach (var rows in benefitRows.Values)
            {
                var ordered = rows.OrderByDescending(row => ParsePeriodEnd(row.Period) ?? DateOnly.MinValue).ToList();
                if (ordered.Count > 0) ordered[0].Weighting = Money(weighting, "current");
                if (ordered.Count > 1) ordered[1].Weighting = Money(weighting, "prior");
            }
        }
    }

    private static void ParseSummaryExperience(List<PdfPageText> pages, PdfExtractionResult result)
    {
        foreach (var page in pages.Where(page =>
                     page.Text.Contains("Experience Summary", StringComparison.OrdinalIgnoreCase)))
        {
            var header = SummaryPeriodHeaderRegex.Match(page.Text);
            if (!header.Success)
            {
                continue;
            }

            var currentPeriod = FormatExperiencePeriod(header, "currentStart", "currentEnd");
            var priorPeriod = FormatExperiencePeriod(header, "priorStart", "priorEnd");
            foreach (var line in page.Lines.Select(ParsingHelpers.Clean))
            {
                var match = SummaryExperienceRowRegex.Match(line);
                if (!match.Success)
                {
                    continue;
                }

                var benefit = NormalizeBenefit(match.Groups["benefit"].Value.Replace("**", "", StringComparison.Ordinal));
                AddSummaryExperienceRow(result, benefit, currentPeriod, match, "current", page.PageNumber);
                AddSummaryExperienceRow(result, benefit, priorPeriod, match, "prior", page.PageNumber);
            }
        }
    }

    private static void AddSummaryExperienceRow(
        PdfExtractionResult result,
        string benefit,
        string period,
        Match match,
        string prefix,
        int pageNumber)
    {
        if (result.ExperienceLines.Any(row =>
                row.Benefit.Equals(benefit, StringComparison.OrdinalIgnoreCase) &&
                row.Period.Equals(period, StringComparison.OrdinalIgnoreCase)))
        {
            return;
        }

        var premium = Money(match, prefix + "Premium");
        var paid = Money(match, prefix + "Paid");
        var incurred = Money(match, prefix + "Incurred");
        var ratio = Money(match, prefix + "Ratio");
        var source = $"Empire Life Group Renewal Summary PDF page {pageNumber}";
        result.ExperienceLines.Add(new ExperienceLine
        {
            Benefit = benefit,
            Period = period,
            BilledPremium = premium,
            PaidClaims = paid,
            IncurredClaims = incurred,
            PaidLossRatio = LossRatio(premium, paid),
            IncurredLossRatio = ratio,
            SourceReference = source
        });
        result.EmpireLifeExperienceRows.Add(new EmpireLifeExperienceRow
        {
            Benefit = benefit,
            Period = period,
            AdjustedPremium = premium,
            PaidClaims = paid,
            IncurredClaims = incurred,
            IncurredLossRatio = ratio,
            SourceReference = source
        });
    }

    private static void ParseDemographics(IEnumerable<PdfPageText> pages, PdfExtractionResult result)
    {
        foreach (var page in pages.Where(page =>
                     page.Text.Contains("Age Volume Distribution - Life", StringComparison.OrdinalIgnoreCase)))
        {
            foreach (var line in page.Lines.Select(ParsingHelpers.Clean))
            {
                var match = DemographicRowRegex.Match(line);
                if (!match.Success)
                {
                    continue;
                }

                result.LifeDemographics.Add(new DemographicLine
                {
                    Benefit = "Life",
                    AgeBand = DemographicAgeBandNormalizer.NormalizeLifeAgeBand(match.Groups["age"].Value),
                    CurrentMaleLives = Int(match, "currentMaleLives"),
                    CurrentMaleVolume = Money(match, "currentMaleVolume"),
                    CurrentFemaleLives = Int(match, "currentFemaleLives"),
                    CurrentFemaleVolume = Money(match, "currentFemaleVolume"),
                    PriorMaleLives = Int(match, "priorMaleLives"),
                    PriorMaleVolume = Money(match, "priorMaleVolume"),
                    PriorFemaleLives = Int(match, "priorFemaleLives"),
                    PriorFemaleVolume = Money(match, "priorFemaleVolume"),
                    SourceReference = $"Empire Life Group Renewal Summary PDF page {page.PageNumber}"
                });
            }
        }
    }

    private static WorkbookExtractionResult ParseUtilization(
        string pdfPath,
        IEnumerable<PdfPageText> pages,
        PdfExtractionResult pdf)
    {
        var workbook = new WorkbookExtractionResult { SourcePath = pdfPath };
        var rows = new List<ClaimsDistributionRow>();
        foreach (var page in pages.Where(page =>
                     page.Text.Contains("Health Claims by Major Medical", StringComparison.OrdinalIgnoreCase) ||
                     page.Text.Contains("Dental Claims Summary", StringComparison.OrdinalIgnoreCase)))
        {
            var period = UtilizationPeriodHeaderRegex.Match(page.Text);
            if (!period.Success)
            {
                continue;
            }

            var currentPeriod = FormatPeriod(period, "currentStart", "currentEnd");
            var priorPeriod = FormatPeriod(period, "priorStart", "priorEnd");
            var isDental = page.Text.Contains("Dental Claims Summary", StringComparison.OrdinalIgnoreCase);
            var created = ReportCreatedDateRegex.Match(page.Text);
            foreach (var line in page.Lines.Select(ParsingHelpers.Clean))
            {
                var match = UtilizationRowRegex.Match(line);
                var hasPercentages = match.Success;
                if (!match.Success)
                {
                    match = UtilizationCompactRowRegex.Match(line);
                }
                if (!match.Success)
                {
                    continue;
                }

                var category = ParsingHelpers.Clean(match.Groups["category"].Value);
                var mapped = MapUtilizationCategory(category, isDental);
                rows.Add(new ClaimsDistributionRow
                {
                    CarrierName = "Empire Life",
                    PolicyNumber = pdf.PolicyNumber,
                    ClientName = pdf.ClientName,
                    ReportPeriod = currentPeriod,
                    PeriodStart = DateText(period, "currentStart"),
                    PeriodEnd = DateText(period, "currentEnd"),
                    PriorReportPeriod = priorPeriod,
                    PriorPeriodStart = DateText(period, "priorStart"),
                    PriorPeriodEnd = DateText(period, "priorEnd"),
                    Benefit = mapped.Benefit,
                    ServiceCategory = mapped.ServiceCategory,
                    ServiceDescriptionOriginal = category,
                    ReportCreatedDate = created.Success ? DateText(created, "value") : "",
                    SourceFile = Path.GetFileName(pdfPath),
                    AmountPaid = Money(match, "currentPaid"),
                    PercentPaid = hasPercentages ? Money(match, "currentPercent") : null,
                    PriorAmountPaid = Money(match, "priorPaid"),
                    PriorPercentPaid = hasPercentages ? Money(match, "priorPercent") : null,
                    Occurrences = Int(match, "currentOccurrences"),
                    PriorOccurrences = Int(match, "priorOccurrences"),
                    SourceReference = $"Empire Life PDF page {page.PageNumber}"
                });
            }
        }

        ClaimsDistributionExtractor.MergeIntoWorkbook(workbook, rows);
        return workbook;
    }

    private static (string Benefit, string ServiceCategory) MapUtilizationCategory(string category, bool isDental)
    {
        if (category.Equals("Total Major Services", StringComparison.OrdinalIgnoreCase))
        {
            return (isDental ? "Dental" : "Extended Health Care", "Total");
        }

        if (isDental)
        {
            return ("Dental", category.Replace("/ ", "/", StringComparison.Ordinal));
        }

        if (category.Contains("Paramedical", StringComparison.OrdinalIgnoreCase))
        {
            return ("Extended Health Care", "Paramedical");
        }

        if (ParamedicalCategoryNormalizer.Normalize(category) is var normalized &&
            !normalized.Equals(category, StringComparison.OrdinalIgnoreCase))
        {
            return ("Paramedical Detail", "Paramedical");
        }

        if (category.Contains("Physio", StringComparison.OrdinalIgnoreCase) ||
            category.Contains("Massage", StringComparison.OrdinalIgnoreCase) ||
            category.Contains("Chiropr", StringComparison.OrdinalIgnoreCase) ||
            category.Contains("Psycholog", StringComparison.OrdinalIgnoreCase))
        {
            return ("Paramedical Detail", "Paramedical");
        }

        return ("Extended Health Care", category);
    }

    private static Dictionary<string, ExperienceSummary> ParseExperienceSummaries(IEnumerable<PdfPageText> pages)
    {
        var summaries = new Dictionary<string, ExperienceSummary>(StringComparer.OrdinalIgnoreCase);
        foreach (var line in pages.SelectMany(page => page.Lines).Select(ParsingHelpers.Clean))
        {
            var match = ExperienceSummaryRegex.Match(line);
            if (!match.Success)
            {
                continue;
            }

            summaries[NormalizeBenefit(match.Groups["benefit"].Value)] = new ExperienceSummary(
                Money(match, "weighted"),
                Money(match, "target"),
                Money(match, "credibility"));
        }
        return summaries;
    }

    private static string ParseLabeledValue(IEnumerable<PdfPageText> pages, Regex regex)
    {
        foreach (var line in pages.SelectMany(page => page.Lines).Select(ParsingHelpers.Clean))
        {
            var match = regex.Match(line);
            if (match.Success)
            {
                return ParsingHelpers.Clean(match.Groups["value"].Value);
            }
        }
        return "";
    }

    private static DateOnly? ParseRenewalDate(IEnumerable<PdfPageText> pages)
    {
        var value = ParseLabeledValue(pages, RenewalDateRegex);
        return DateOnly.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out var date)
            ? date
            : null;
    }

    private static string FormatPeriod(Match match, string startGroup, string endGroup)
    {
        return $"{ParsingHelpers.Clean(match.Groups[startGroup].Value)} - {ParsingHelpers.Clean(match.Groups[endGroup].Value)}";
    }

    private static string FormatExperiencePeriod(Match match, string startGroup, string endGroup)
    {
        if (!DateOnly.TryParse(match.Groups[startGroup].Value, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out var start) ||
            !DateOnly.TryParse(match.Groups[endGroup].Value, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out var end))
        {
            return NormalizePeriod(FormatPeriod(match, startGroup, endGroup));
        }

        return $"{start:MMM yyyy}-{end:MMM yyyy}";
    }

    private static string DateText(Match match, string groupName)
    {
        return DateOnly.TryParse(match.Groups[groupName].Value, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out var date)
            ? date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
            : "";
    }

    private static DateOnly? ParsePeriodEnd(string period)
    {
        var parts = period.Split('-', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length < 2)
        {
            return null;
        }

        return DateOnly.TryParseExact($"01 {parts[^1]}", "dd MMM yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var date)
            ? date
            : null;
    }

    private static string NormalizePeriod(string period)
    {
        return Regex.Replace(ParsingHelpers.Clean(period), @"\s*-\s*", "-");
    }

    private static string NormalizeBenefit(string value)
    {
        value = ParsingHelpers.Clean(value);
        if (value.Equals("Extended Health", StringComparison.OrdinalIgnoreCase)) return "Extended Health Care";
        if (value.Equals("Dependant Life", StringComparison.OrdinalIgnoreCase)) return "Dependent Life";
        if (value.Contains("Accidental Death", StringComparison.OrdinalIgnoreCase)) return "Accidental Death and Dismemberment";
        return value;
    }

    private static decimal? SumPremium(IEnumerable<RateLine> lines, string benefit, bool current)
    {
        var values = lines
            .Where(row => row.Benefit.Contains(benefit, StringComparison.OrdinalIgnoreCase))
            .Select(row => current ? row.CurrentPremium : row.RenewalPremium)
            .Where(value => value.HasValue)
            .Select(value => value!.Value)
            .ToList();
        return values.Count == 0 ? null : decimal.Round(values.Sum(), 2);
    }

    private static decimal? LossRatio(decimal? premium, decimal? claims)
    {
        return premium is > 0m && claims.HasValue
            ? decimal.Round(claims.Value / premium.Value * 100m, 2)
            : null;
    }

    private static bool HasWorkbookData(WorkbookExtractionResult workbook)
    {
        return workbook.ClaimsDistributionRows.Count > 0 ||
               workbook.ClaimTrends.Count > 0 ||
               workbook.EhcSpendingTrends.Count > 0 ||
               workbook.DentalSpendingTrends.Count > 0 ||
               workbook.ParamedicalServiceTrends.Count > 0;
    }

    private static decimal? Money(Match match, string groupName)
    {
        return match.Success ? ParsingHelpers.ParseDecimal(match.Groups[groupName].Value) : null;
    }

    private static string FirstNonBlank(params string[] values)
    {
        return values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)) ?? "";
    }

    private static int? Int(Match match, string groupName)
    {
        return match.Success ? ParsingHelpers.ParseInt(match.Groups[groupName].Value) : null;
    }

    private sealed record ExperienceSummary(
        decimal? WeightedLossRatio,
        decimal? TargetLossRatio,
        decimal? Credibility);
}

using System.Globalization;
using System.Text.RegularExpressions;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed class CarrierPdfParser
{
    private readonly PdfTextExtractor _extractor;
    private readonly ManulifePdfParser _manulifePdfParser;

    public CarrierPdfParser(PdfTextExtractor? extractor = null)
    {
        _extractor = extractor ?? new PdfTextExtractor();
        _manulifePdfParser = new ManulifePdfParser(_extractor);
    }

    public CarrierPdfParseResult Parse(string pdfPath)
    {
        var pages = _extractor.Extract(pdfPath);
        var allText = ParsingHelpers.Clean(string.Join(" ", pages.Select(page => page.Text)));

        CarrierPdfParseResult result;
        if (AllstateCriticalIllnessPdfParser.IsAllstateCriticalIllnessApplication(pdfPath, allText))
        {
            result = new AllstateCriticalIllnessPdfParser().Parse(pdfPath, pages, allText);
            AttachClaimsDistributionRows(result, pdfPath, pages);
            return result;
        }

        if (GroupSourcePdfParser.IsGroupSourceDocument(pdfPath, allText))
        {
            result = new GroupSourcePdfParser().Parse(pdfPath, pages, allText);
            AttachClaimsDistributionRows(result, pdfPath, pages);
            return result;
        }

        if (RwamRateSheetPdfParser.IsRwamRateSheet(allText))
        {
            result = new RwamRateSheetPdfParser().Parse(pdfPath, pages, allText);
            AttachClaimsDistributionRows(result, pdfPath, pages);
            return result;
        }

        if (GroupHealthPdfParser.IsGroupHealthDocument(allText) ||
            GroupHealthPdfParser.IsHybridAsoDentalClaimsExperienceDocument(pdfPath, allText))
        {
            result = new GroupHealthPdfParser().Parse(pdfPath, pages, allText);
            AttachClaimsDistributionRows(result, pdfPath, pages);
            return result;
        }

        if (EmpireLifePdfParser.IsEmpireLifeDocument(pdfPath, allText))
        {
            result = new EmpireLifePdfParser().Parse(pdfPath, pages, allText);
            AttachClaimsDistributionRows(result, pdfPath, pages);
            return result;
        }

        if (SunLifePdfParser.IsSunLifeRenewal(allText))
        {
            result = new SunLifePdfParser().Parse(pdfPath, pages, allText);
            AttachClaimsDistributionRows(result, pdfPath, pages);
            return result;
        }

        if (VictorPdfParser.IsVictorRenewal(allText))
        {
            result = new VictorPdfParser().Parse(pdfPath, pages, allText);
            AttachClaimsDistributionRows(result, pdfPath, pages);
            return result;
        }

        if (EquitablePdfParser.IsEquitableRenewal(allText))
        {
            result = new EquitablePdfParser().Parse(pdfPath, pages, allText);
            AttachClaimsDistributionRows(result, pdfPath, pages);
            return result;
        }

        if (CanadaLifePdfParser.IsCanadaLifeRenewal(allText))
        {
            result = new CanadaLifePdfParser().Parse(pdfPath, pages, allText);
            AttachClaimsDistributionRows(result, pdfPath, pages);
            return result;
        }

        result = new CarrierPdfParseResult
        {
            Pdf = _manulifePdfParser.Parse(pdfPath),
            Carrier = "Manulife"
        };
        AttachClaimsDistributionRows(result, pdfPath, pages);
        return result;
    }

    private static void AttachClaimsDistributionRows(CarrierPdfParseResult result, string pdfPath, IReadOnlyList<PdfPageText> pages)
    {
        var rows = ClaimsDistributionExtractor.Extract(
            pdfPath,
            pages,
            FirstNonBlank(result.Carrier, result.Pdf.Carrier),
            result.Pdf.PolicyNumber,
            result.Pdf.ClientName);
        if (rows.Count == 0)
        {
            return;
        }

        result.WorkbookFromPdf ??= new WorkbookExtractionResult
        {
            SourcePath = pdfPath
        };
        ClaimsDistributionExtractor.MergeIntoWorkbook(result.WorkbookFromPdf, rows);
    }

    private static string FirstNonBlank(params string[] values)
    {
        return values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)) ?? "";
    }
}

public sealed class CarrierPdfParseResult
{
    public PdfExtractionResult Pdf { get; set; } = new();
    public WorkbookExtractionResult? WorkbookFromPdf { get; set; }
    public string Carrier { get; set; } = "";
}

public sealed class EquitablePdfParser
{
    private static readonly Regex ClientPolicyLineRegex = new(@"^(?:RE:\s*)?(?<client>.+?)\s+Policy Number:\s*(?<policy>\d+)", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex PolicyNumberRegex = new(@"Policy\s*(?:Number|#):\s*(?<value>\d+)", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex PolicyNameRegex = new(@"Policy Name:\s*(?<value>.+?)(?:\s+GAE:|$)", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex RenewalDateRegex = new(@"Renewal Date:\s*(?<date>[A-Za-z]+\s+\d{1,2},\s+\d{4})", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex FooterIdentityRegex = new(@"\b(?<policy>\d{4,})\s*\|\s*(?<client>[^|]+?)\s*\|\s*Renewal\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex FooterRenewalDateRegex = new(@"\|\s*Renewal\s+(?<date>[A-Za-z]+\s+\d{1,2},\s+\d{4})", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex TotalPremiumRegex = new(@"^Total Premium\s+\$(?<current>[\d,]+\.\d{2})\s+\$(?<renewal>[\d,]+\.\d{2})\s+(?<change>[+-]?[\d.]+)%$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex VolumeRateRegex = new(@"^(?<benefit>.+?)\s+(?:\$(?<unitValue>[\d,]+)\s+)?(?<class>[A-Z0-9]+(?:\s*,\s*[A-Z0-9]+)*)\s+(?<lives>\d+)\s+\$(?<volume>[\d,]+)\s+\$(?<currentRate>[\d.]+)\s+\$(?<currentPremium>[\d,]+\.\d{2})\s+\$(?<renewalRate>[\d.]+)\s+\$(?<renewalPremium>[\d,]+\.\d{2})\s+(?<change>[+-]?[\d.]+)%$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex UnitRateRegex = new(@"^(?<benefit>.+?)\s+(?<class>[A-Z0-9]+(?:\s*,\s*[A-Z0-9]+)*)\s+(?<lives>\d+)\s+\$(?<currentRate>[\d.]+)\s+\$(?<currentPremium>[\d,]+\.\d{2})\s+\$(?<renewalRate>[\d.]+)\s+\$(?<renewalPremium>[\d,]+\.\d{2})\s+(?<change>[+-]?[\d.]+)%$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex PendingUnitRateRegex = new(@"^(?<class>[A-Z0-9]+(?:\s*,\s*[A-Z0-9]+)*)\s+(?<lives>\d+)\s+\$(?<currentRate>[\d.]+)\s+\$(?<currentPremium>[\d,]+\.\d{2})\s+\$(?<renewalRate>[\d.]+)\s+\$(?<renewalPremium>[\d,]+\.\d{2})\s+(?<change>[+-]?[\d.]+)%$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex TierRateRegex = new(@"^(?<benefit>Health|Dental)\s+(?<tier>Single|Couple|Family)\s+(?<class>[A-Z0-9]+(?:\s*,\s*[A-Z0-9]+)*)\s+(?<lives>\d+)\s+\$(?<currentRate>[\d.]+)\s+\$(?<currentPremium>[\d,]+\.\d{2})\s+\$(?<renewalRate>[\d.]+)\s+\$(?<renewalPremium>[\d,]+\.\d{2})\s+(?<change>[+-]?[\d.]+)%$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex ExperiencePeriodRegex = new(@"From:\s*(?<from>\d{2}-[A-Za-z]{3}-\d{4})\s+to\s+(?<to>\d{2}-[A-Za-z]{3}-\d{4})", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex PlainExperiencePeriodRegex = new(@"\b(?<from>\d{2}-[A-Za-z]{3}-\d{4})\s+to\s+(?<to>\d{2}-[A-Za-z]{3}-\d{4})\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex ExperienceSummaryRegex = new(@"^(?<benefit>Health|Dental)\s+\$(?<paid>[\d,]+\.\d{2})\s+\$(?<future>[\d,]+\.\d{2})\s+\$(?<closing>[\d,]+\.\d{2})\s+\$(?<opening>[\d,]+\.\d{2})\s+\$(?<incurred>[\d,]+\.\d{2})\s+\$(?<premium>[\d,]+\.\d{2})\s+(?<ratio>[\d.]+)$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex UnderwritingExperienceRowRegex = new(@"^(?<from>\d{2}\s+[A-Za-z]{3}\s+\d{4})\s+(?<to>\d{2}\s+[A-Za-z]{3}\s+\d{4})\s+\$(?<premium>[\d,]+(?:\.\d{2})?)\s+\$(?<paid>[\d,]+(?:\.\d{2})?)\s+\$(?<incurred>[\d,]+(?:\.\d{2})?)\s+(?<ratio>[\d.]+)%$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex TwoPeriodHeaderRegex = new(@"(?<priorFrom>\d{2}-[A-Za-z]{3}-\d{4})\s+to\s+(?<priorTo>\d{2}-[A-Za-z]{3}-\d{4})\s+(?<currentFrom>\d{2}-[A-Za-z]{3}-\d{4})\s+to\s+(?<currentTo>\d{2}-[A-Za-z]{3}-\d{4})", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex TwoPeriodClaimsRowRegex = new(@"^(?<category>.+?)\s+(?<priorOccurrences>\d+)\s+\$(?<priorPaid>[\d,]+\.\d{2})\s+\$(?<priorAverage>[\d,]+\.\d{2})\s+(?<priorPercent>[\d.]+)%\s+(?<currentOccurrences>\d+)\s+\$(?<currentPaid>[\d,]+\.\d{2})\s+\$(?<currentAverage>[\d,]+\.\d{2})\s+(?<currentPercent>[\d.]+)%$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex WeightingRatioRegex = new(@"based on a\s+(?<current>\d+)\s*/\s*(?<prior>\d+)\s+weighting", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex TargetLossRatioRegex = new(@"\bTarget Loss Ratio\s+(?<value>[\d.]+)%", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex CredibilityRegex = new(@"\bCredibility\s+(?<value>[\d.]+)%", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex AnnualInflationRegex = new(@"\bAnnual (?:Inflation|Trend)\s+(?<value>[\d.]+)%", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex EquitablePoolingChargePercentRegex = new(@"\bPool charges of\s+(?<value>[\d.]+)%\s+of\s+Premium", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex EquitablePoolingThresholdRegex = new(@"\bpooling level of\s+\$(?<value>[\d,]+(?:\.\d{2})?)\s+per individual", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex EquitablePooledPremiumRegex = new(@"\bPooled Premium\s+\$(?<value>[\d,]+(?:\.\d{2})?)", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex EquitableTotalBilledPremiumRegex = new(@"\bTotal Billed Premium\s+\$(?<value>[\d,]+(?:\.\d{2})?)", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex BenefitSummaryTotalRegex = new(@"^(?<category>Dental Accident|Drug|Major Services|Vision)\s+Total\s+\$(?<charged>[\d,]+\.\d{2})\s+\$[\d,]+\.\d{2}\s+\$[\d,]+\.\d{2}\s+\$[\d,]+\.\d{2}\s+\(?\$?-?[\d,]+\.\d{2}\)?\s+\$(?<paid>[\d,]+\.\d{2})$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex HealthClaimsRowRegex = new(@"^(?<category>.+?)\s+(?<occurrences>\d+)\s+\$(?<paid>[\d,]+\.\d{2})\s+\$[\d,]+\.\d{2}\s+(?<percent>[\d.]+)%$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex HealthMajorServicesRowRegex = new(@"^(?<category>.+?)\s+(?<occurrences>\d+)\s+\$(?<paid>[\d,]+\.\d{2})\s+\$[\d,]+\.\d{2}\s+(?<percent>[\d.]+)%$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex TelusTopDinPeriodRegex = new(@"From:\s*(?<from>\d{2}-[A-Za-z]{3}-\d{4}).*?To:\s*(?<to>\d{2}-[A-Za-z]{3}-\d{4})", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex TelusTopDinRowStartRegex = new(@"^(?<rank>\d{1,2})\s+(?<din>\d{7,8})\b", RegexOptions.Compiled);
    private static readonly Regex TelusTopDinRowRegex = new(@"^(?<rank>\d{1,2})\s+(?<din>\d{7,8})\s+(?<detail>.+?)\s+(?<rx>\d+)\s+[\d.]+\s+\d+\s+[\d.]+\s+\$?(?<eligible>[\d,]+\.\d{2})\s+[\d.]+\s+\$?(?<paid>[\d,]+\.\d{2})\s+[\d.]+$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex EquitableTopPaidPeriodRegex = new(@"From:\s*(?<from>\d{2}-[A-Za-z]{3}-\d{4})\s+to\s+(?<to>\d{2}-[A-Za-z]{3}-\d{4})", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex EquitableTopPaidRankRegex = new(@"^(?<rank>\d{1,2})$", RegexOptions.Compiled);
    private static readonly Regex EquitableTopPaidDinRegex = new(@"^\d{7,8}$", RegexOptions.Compiled);
    private static readonly Regex EquitableTopPaidAmountRegex = new(@"^\$?(?<value>[\d,]+\.\d{2})$", RegexOptions.Compiled);
    private static readonly Regex EquitableTopPaidPercentRegex = new(@"^[\d.]+%$", RegexOptions.Compiled);
    private static readonly Regex EquitableTopPaidRowRegex = new(@"^(?<rank>\d{1,2})\s+(?<din>\d{7,8})\s+(?<name>.+?)\s+\$?(?<amount>[\d,]+\.\d{2})\s+(?<percent>[\d.]+)%$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex EquitableLifeVolumeRowRegex = new(@"^(?<priorAge>0-19|20-24|25-29|30-34|35-39|40-44|45-49|50-54|55-59|60-64|65-69|70\+)\s+\$(?<priorMale>[\d,]+)\s+\$(?<priorFemale>[\d,]+)\s+\$(?<priorTotal>[\d,]+)\s+[\d.]+%\s+(?<currentAge>0-19|20-24|25-29|30-34|35-39|40-44|45-49|50-54|55-59|60-64|65-69|70\+)\s+\$(?<currentMale>[\d,]+)\s+\$(?<currentFemale>[\d,]+)\s+\$(?<currentTotal>[\d,]+)\s+[\d.]+%$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex EquitableLivesRowRegex = new(@"^(?<priorAge>0-19|20-24|25-29|30-34|35-39|40-44|45-49|50-54|55-59|60-64|65-69|70\+)\s+(?<priorMale>\d+)\s+(?<priorFemale>\d+)\s+(?<priorTotal>\d+)\s+[\d.]+%\s+(?<currentAge>0-19|20-24|25-29|30-34|35-39|40-44|45-49|50-54|55-59|60-64|65-69|70\+)\s+(?<currentMale>\d+)\s+(?<currentFemale>\d+)\s+(?<currentTotal>\d+)\s+[\d.]+%$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly string[] TelusDrugDiseaseCategories =
    [
        "Fertility disorders",
        "ADD/Narcolepsy",
        "Mental disorders",
        "Vitamins and Minerals",
        "Allergy conditions",
        "Birth control",
        "Cholesterol disorders",
        "Thyroid & Glandular",
        "Hypertension",
        "Menopause",
        "Diabetes",
        "Epilepsy",
        "Insomnia",
        "Migraines",
        "All other"
    ];

    public static bool IsEquitableRenewal(string text)
    {
        return text.Contains("The Equitable Life Insurance Company of Canada", StringComparison.OrdinalIgnoreCase) ||
               text.Contains("equitable.ca", StringComparison.OrdinalIgnoreCase) ||
               text.Contains("RenewalStProc_EL_", StringComparison.OrdinalIgnoreCase) ||
               IsEquitableTopPaidReport(text) ||
               IsTelusTopDinReport(text);
    }

    public CarrierPdfParseResult Parse(string pdfPath, List<PdfPageText> pages, string allText)
    {
        var pdf = new PdfExtractionResult
        {
            SourcePath = pdfPath,
            Carrier = "Equitable",
            PolicyNumber = ParsePolicyNumber(pages, allText),
            ClientName = ParseClientName(pages, allText),
            RenewalDate = ParseRenewalDate(allText)
        };

        ParseMonthlyCostSummary(pages, pdf);
        ParseExperienceSummary(pages, pdf);
        ParseUnderwritingExperienceHistory(pages, pdf);
        ParsePoolingArrangement(pages, pdf);
        ParseDemographics(pages, pdf);

        var workbook = ParseWorkbookDetailFromPdf(pdfPath, pages);

        if (pdf.RateLines.Count == 0 &&
            pdf.ExperienceLines.Count == 0 &&
            !HasWorkbookData(workbook))
        {
            pdf.Warnings.Add("No Equitable renewal values were extracted from the PDF.");
        }

        return new CarrierPdfParseResult
        {
            Pdf = pdf,
            WorkbookFromPdf = HasWorkbookData(workbook) ? workbook : null,
            Carrier = "Equitable"
        };
    }

    private static string ParseClientName(List<PdfPageText> pages, string allText)
    {
        foreach (var line in pages.SelectMany(page => page.Lines).Select(ParsingHelpers.Clean))
        {
            var match = ClientPolicyLineRegex.Match(line);
            if (match.Success)
            {
                return ParsingHelpers.Clean(match.Groups["client"].Value);
            }

            match = PolicyNameRegex.Match(line);
            if (match.Success)
            {
                return ParsingHelpers.Clean(match.Groups["value"].Value);
            }
        }

        var footerMatch = FooterIdentityRegex.Match(allText);
        if (footerMatch.Success)
        {
            return ParsingHelpers.Clean(footerMatch.Groups["client"].Value);
        }

        return ParseEquitableTopPaidClientName(pages);
    }

    private static string ParsePolicyNumber(List<PdfPageText> pages, string allText)
    {
        foreach (var line in pages.SelectMany(page => page.Lines).Select(ParsingHelpers.Clean))
        {
            var match = PolicyNumberRegex.Match(line);
            if (match.Success)
            {
                return ParsingHelpers.Clean(match.Groups["value"].Value);
            }
        }

        var allTextMatch = PolicyNumberRegex.Match(allText);
        if (allTextMatch.Success)
        {
            return ParsingHelpers.Clean(allTextMatch.Groups["value"].Value);
        }

        var footerMatch = FooterIdentityRegex.Match(allText);
        return footerMatch.Success ? ParsingHelpers.Clean(footerMatch.Groups["policy"].Value) : "";
    }

    private static DateOnly? ParseRenewalDate(string text)
    {
        var match = RenewalDateRegex.Match(text);
        if (match.Success &&
            DateOnly.TryParse(match.Groups["date"].Value, CultureInfo.InvariantCulture, DateTimeStyles.None, out var renewalDate))
        {
            return renewalDate;
        }

        match = FooterRenewalDateRegex.Match(text);
        if (match.Success &&
            DateOnly.TryParse(match.Groups["date"].Value, CultureInfo.InvariantCulture, DateTimeStyles.None, out renewalDate))
        {
            return renewalDate;
        }

        return null;
    }

    private static void ParseMonthlyCostSummary(List<PdfPageText> pages, PdfExtractionResult result)
    {
        var page = pages.FirstOrDefault(page => page.Text.Contains("Total monthly cost summary", StringComparison.OrdinalIgnoreCase));
        if (page is null)
        {
            return;
        }

        var pendingBenefit = "";
        foreach (var rawLine in page.Lines)
        {
            var line = ParsingHelpers.Clean(rawLine);
            if (line.Length == 0)
            {
                continue;
            }

            var totalMatch = TotalPremiumRegex.Match(line);
            if (totalMatch.Success)
            {
                result.CurrentMonthlyPremium = Money(totalMatch, "current");
                result.RenewalMonthlyPremium = Money(totalMatch, "renewal");
                result.MonthlyPremiumChange = result.RenewalMonthlyPremium - result.CurrentMonthlyPremium;
                result.MonthlyPremiumChangePercent = Money(totalMatch, "change");
                continue;
            }

            if (line.Equals("Employee and Family", StringComparison.OrdinalIgnoreCase))
            {
                pendingBenefit = "Employee and Family Assistance Program";
                continue;
            }

            if (line.Equals("Assistance Program", StringComparison.OrdinalIgnoreCase))
            {
                pendingBenefit = "Employee and Family Assistance Program";
                continue;
            }

            if (TryNormalizeEquitableRateHeading(line, out var headingBenefit))
            {
                pendingBenefit = headingBenefit;
                continue;
            }

            var rateLine = TryParseRateLine(line, pendingBenefit, page.PageNumber);
            if (rateLine is not null)
            {
                result.RateLines.Add(rateLine);
                pendingBenefit = "";
            }
        }

        result.CurrentEhcPremium = SumPremium(result.RateLines, "Extended Health Care", current: true);
        result.RenewalEhcPremium = SumPremium(result.RateLines, "Extended Health Care", current: false);
        result.CurrentDentalPremium = SumPremium(result.RateLines, "Dental", current: true);
        result.RenewalDentalPremium = SumPremium(result.RateLines, "Dental", current: false);
    }

    private static RateLine? TryParseRateLine(string line, string pendingBenefit, int pageNumber)
    {
        var match = TierRateRegex.Match(line);
        if (match.Success)
        {
            var benefit = NormalizeBenefit(match.Groups["benefit"].Value);
            return BuildRateLine(match, benefit, pageNumber, hasVolume: false, tier: match.Groups["tier"].Value);
        }

        match = VolumeRateRegex.Match(line);
        if (match.Success && IsEquitableRateBenefit(match.Groups["benefit"].Value))
        {
            return BuildRateLine(match, NormalizeBenefit(match.Groups["benefit"].Value), pageNumber, hasVolume: true);
        }

        match = UnitRateRegex.Match(line);
        if (match.Success && IsEquitableRateBenefit(match.Groups["benefit"].Value))
        {
            return BuildRateLine(match, NormalizeBenefit(match.Groups["benefit"].Value), pageNumber, hasVolume: false);
        }

        match = PendingUnitRateRegex.Match(line);
        if (match.Success && !string.IsNullOrWhiteSpace(pendingBenefit))
        {
            return BuildRateLine(match, pendingBenefit, pageNumber, hasVolume: false);
        }

        return null;
    }

    private static bool TryNormalizeEquitableRateHeading(string line, out string benefit)
    {
        benefit = NormalizeBenefit(line);
        return IsEquitableRateBenefit(benefit) &&
               !line.Contains('$') &&
               !line.Any(char.IsDigit);
    }

    private static bool IsEquitableRateBenefit(string value)
    {
        var benefit = NormalizeBenefit(value);
        return benefit.Equals("Life", StringComparison.OrdinalIgnoreCase) ||
               benefit.Equals("Accidental Death and Dismemberment", StringComparison.OrdinalIgnoreCase) ||
               benefit.Equals("LTD", StringComparison.OrdinalIgnoreCase) ||
               benefit.Equals("STD", StringComparison.OrdinalIgnoreCase) ||
               benefit.Equals("Dependent Life", StringComparison.OrdinalIgnoreCase) ||
               benefit.Equals("Critical Illness", StringComparison.OrdinalIgnoreCase) ||
               benefit.Equals("Employee and Family Assistance Program", StringComparison.OrdinalIgnoreCase) ||
               benefit.Equals("Virtual Healthcare", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsEquitableTopPaidReport(string text)
    {
        return text.Contains("Top 50 Drugs By Total Paid Amount", StringComparison.OrdinalIgnoreCase) &&
               text.Contains("Policy #:", StringComparison.OrdinalIgnoreCase);
    }

    private static string ParseEquitableTopPaidClientName(List<PdfPageText> pages)
    {
        foreach (var page in pages.Where(page => IsEquitableTopPaidReport(page.Text)))
        {
            var footerIndex = page.Lines.FindIndex(line =>
                ParsingHelpers.Clean(line).StartsWith("Page:", StringComparison.OrdinalIgnoreCase));
            if (footerIndex >= 0 && footerIndex + 1 < page.Lines.Count)
            {
                var candidate = ParsingHelpers.Clean(page.Lines[footerIndex + 1]);
                if (!string.IsNullOrWhiteSpace(candidate))
                {
                    return candidate;
                }
            }
        }

        return "";
    }

    private static RateLine BuildRateLine(Match match, string benefit, int pageNumber, bool hasVolume, string? tier = null)
    {
        return new RateLine
        {
            Benefit = benefit,
            TierOrClass = tier ?? "",
            AccountClass = match.Groups["class"].Value,
            Lives = Int(match, "lives"),
            Volume = hasVolume ? Money(match, "volume") : null,
            CurrentRate = Money(match, "currentRate"),
            CurrentPremium = Money(match, "currentPremium"),
            RenewalRate = Money(match, "renewalRate"),
            RenewalPremium = Money(match, "renewalPremium"),
            PercentChange = Money(match, "change"),
            SourceReference = $"Equitable PDF page {pageNumber}"
        };
    }

    private static void ParseExperienceSummary(List<PdfPageText> pages, PdfExtractionResult result)
    {
        foreach (var page in pages.Where(page => page.Text.Contains("Experience Summary", StringComparison.OrdinalIgnoreCase)))
        {
            var period = ParseExperiencePeriod(page.Text);
            foreach (var line in page.Lines.Select(ParsingHelpers.Clean))
            {
                var match = ExperienceSummaryRegex.Match(line);
                if (!match.Success)
                {
                    continue;
                }

                var premium = Money(match, "premium");
                var paid = Money(match, "paid");
                var incurred = Money(match, "incurred");
                result.ExperienceLines.Add(new ExperienceLine
                {
                    Benefit = NormalizeBenefit(match.Groups["benefit"].Value),
                    Period = period,
                    BilledPremium = premium,
                    PaidClaims = paid,
                    IncurredClaims = incurred,
                    PaidLossRatio = premium is > 0m && paid.HasValue ? decimal.Round(paid.Value / premium.Value * 100m, 1) : null,
                    IncurredLossRatio = Money(match, "ratio"),
                    SourceReference = $"Equitable Experience Summary PDF page {page.PageNumber}"
                });
            }
        }
    }

    private static void ParseUnderwritingExperienceHistory(List<PdfPageText> pages, PdfExtractionResult result)
    {
        foreach (var page in pages.Where(page =>
                     page.Text.Contains("Target Loss Ratio", StringComparison.OrdinalIgnoreCase) &&
                     (page.Text.Contains("Billed net premium", StringComparison.OrdinalIgnoreCase) ||
                      page.Text.Contains("Billed premium", StringComparison.OrdinalIgnoreCase))))
        {
            var parsedRows = page.Lines
                .Select(ParsingHelpers.Clean)
                .Select(line => UnderwritingExperienceRowRegex.Match(line))
                .Where(match => match.Success)
                .Select(match => new
                {
                    Match = match,
                    PeriodStart = ParseEquitableTableDate(match.Groups["from"].Value)
                })
                .OrderBy(row => row.PeriodStart ?? DateOnly.MinValue)
                .ToList();
            if (parsedRows.Count == 0)
            {
                continue;
            }

            var trend = Money(AnnualInflationRegex.Match(page.Text), "value");
            var credibility = Money(CredibilityRegex.Match(page.Text), "value");
            var targetLossRatio = Money(TargetLossRatioRegex.Match(page.Text), "value");
            var weighting = ParseWeightingRatio(page.Text);
            var benefit = DetectEquitableExperienceBenefit(page.Text);

            for (var index = 0; index < parsedRows.Count; index++)
            {
                var match = parsedRows[index].Match;
                var premium = Money(match, "premium");
                var paid = Money(match, "paid");
                result.ExperienceLines.Add(new ExperienceLine
                {
                    Benefit = benefit,
                    Period = FormatEquitablePeriod(match.Groups["from"].Value, match.Groups["to"].Value),
                    BilledPremium = premium,
                    PaidClaims = paid,
                    IncurredClaims = Money(match, "incurred"),
                    PaidLossRatio = premium is > 0m && paid.HasValue
                        ? decimal.Round(paid.Value / premium.Value * 100m, 1)
                        : null,
                    IncurredLossRatio = Money(match, "ratio"),
                    Weighting = GetPeriodWeight(index, parsedRows.Count, weighting),
                    TrendFactor = trend,
                    Credibility = credibility,
                    TargetLossRatio = targetLossRatio,
                    SourceReference = $"Equitable renewal report PDF page {page.PageNumber}"
                });
            }
        }
    }

    private static string DetectEquitableExperienceBenefit(string text)
    {
        return text.Contains("Dental benefits", StringComparison.OrdinalIgnoreCase) ||
               text.Contains("Dental rates", StringComparison.OrdinalIgnoreCase)
            ? "Dental"
            : "Extended Health Care";
    }

    private static (decimal Current, decimal Prior)? ParseWeightingRatio(string text)
    {
        var match = WeightingRatioRegex.Match(text);
        var current = Money(match, "current");
        var prior = Money(match, "prior");
        var total = current.GetValueOrDefault() + prior.GetValueOrDefault();
        return match.Success && total > 0m
            ? (current!.Value / total * 100m, prior!.Value / total * 100m)
            : null;
    }

    private static decimal? GetPeriodWeight(int index, int count, (decimal Current, decimal Prior)? weighting)
    {
        if (!weighting.HasValue)
        {
            return null;
        }

        if (index == count - 1)
        {
            return weighting.Value.Current;
        }

        return index == count - 2 ? weighting.Value.Prior : 0m;
    }

    private static void ParsePoolingArrangement(List<PdfPageText> pages, PdfExtractionResult result)
    {
        var text = string.Join(" ", pages.Select(page => page.Text));
        result.PoolingThreshold ??= Money(EquitablePoolingThresholdRegex.Match(text), "value");
        result.PooledPremium ??= Money(EquitablePooledPremiumRegex.Match(text), "value");

        var explicitCharge = Money(EquitablePoolingChargePercentRegex.Match(text), "value");
        if (explicitCharge.HasValue)
        {
            result.PoolingCharge = explicitCharge;
            return;
        }

        var totalBilledPremium = Money(EquitableTotalBilledPremiumRegex.Match(text), "value");
        if (!result.PoolingCharge.HasValue && result.PooledPremium.HasValue && totalBilledPremium is > 0m)
        {
            result.PoolingCharge = decimal.Round(result.PooledPremium.Value / totalBilledPremium.Value * 100m, 2);
        }
    }

    private static WorkbookExtractionResult ParseWorkbookDetailFromPdf(string pdfPath, List<PdfPageText> pages)
    {
        var workbook = new WorkbookExtractionResult
        {
            SourcePath = pdfPath
        };

        ParseHealthClaimsSummary(pages, workbook);
        ParseHealthMajorServicesClaims(pages, workbook);
        ParseDentalClaimsTables(pages, workbook);
        ParseDentalExperienceSummary(pages, workbook);
        ParseTelusTopDinReport(pages, workbook);
        ParseEquitableTopPaidReport(pages, workbook);
        return workbook;
    }

    private static void ParseTelusTopDinReport(List<PdfPageText> pages, WorkbookExtractionResult workbook)
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var currentPeriod = "Current";
        var inTopDinReport = false;
        foreach (var page in pages)
        {
            if (IsTelusTopDinReport(page.Text))
            {
                currentPeriod = ParseTelusTopDinPeriod(page.Text);
                inTopDinReport = true;
            }

            if (!inTopDinReport || !HasTelusTopDinRows(page))
            {
                continue;
            }

            var pending = "";
            foreach (var rawLine in page.Lines.Select(ParsingHelpers.Clean))
            {
                if (TelusTopDinRowStartRegex.IsMatch(rawLine))
                {
                    TryAddTelusTopDinRow(pending, currentPeriod, page.PageNumber, workbook, seen);
                    pending = rawLine;
                }
                else if (!string.IsNullOrWhiteSpace(pending))
                {
                    pending = $"{pending} {rawLine}";
                }
                else
                {
                    continue;
                }

                if (TryAddTelusTopDinRow(pending, currentPeriod, page.PageNumber, workbook, seen))
                {
                    pending = "";
                }
            }

            TryAddTelusTopDinRow(pending, currentPeriod, page.PageNumber, workbook, seen);
        }
    }

    private static bool HasTelusTopDinRows(PdfPageText page)
    {
        return page.Lines.Any(line => TelusTopDinRowStartRegex.IsMatch(ParsingHelpers.Clean(line)));
    }

    private static bool TryAddTelusTopDinRow(
        string line,
        string period,
        int pageNumber,
        WorkbookExtractionResult workbook,
        ISet<string> seen)
    {
        if (string.IsNullOrWhiteSpace(line))
        {
            return false;
        }

        var match = TelusTopDinRowRegex.Match(ParsingHelpers.Clean(line));
        if (!match.Success)
        {
            return false;
        }

        var rank = Int(match, "rank") ?? 0;
        if (rank is < 1 or > 10)
        {
            return true;
        }

        var din = match.Groups["din"].Value;
        var key = $"{period}|{rank}|{din}";
        if (!seen.Add(key))
        {
            return true;
        }

        var detail = SplitTelusDrugDetail(match.Groups["detail"].Value);
        workbook.DrugTrends.Add(new DrugTrend
        {
            Din = din,
            DrugName = detail.DrugName,
            Treats = detail.Category,
            Period = period,
            PrescriptionCount = Int(match, "rx"),
            SubmittedTotal = Money(match, "eligible"),
            NetAmountPaid = Money(match, "paid"),
            SourceReference = $"Equitable TELUS Top DIN PDF page {pageNumber}"
        });

        return true;
    }

    private static string ParseTelusTopDinPeriod(string text)
    {
        var match = TelusTopDinPeriodRegex.Match(text);
        return match.Success
            ? $"{match.Groups["from"].Value} - {match.Groups["to"].Value}"
            : "Current";
    }

    private static (string DrugName, string Category) SplitTelusDrugDetail(string value)
    {
        var detail = ParsingHelpers.Clean(value);
        foreach (var category in TelusDrugDiseaseCategories)
        {
            if (!detail.EndsWith(category, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var drugName = ParsingHelpers.Clean(detail[..^category.Length]);
            return (drugName, category);
        }

        return (detail, "");
    }

    private static bool IsTelusTopDinReport(string text)
    {
        return text.Contains("Standard Report 13b", StringComparison.OrdinalIgnoreCase) &&
               text.Contains("Top 50 DIN", StringComparison.OrdinalIgnoreCase) &&
               text.Contains("Insurer: EQUITABLE", StringComparison.OrdinalIgnoreCase);
    }

    private static void ParseEquitableTopPaidReport(
        List<PdfPageText> pages,
        WorkbookExtractionResult workbook)
    {
        var inReport = false;
        var period = "Current";
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        int? rank = null;
        var din = "";
        var drugNameParts = new List<string>();
        decimal? amount = null;

        foreach (var page in pages)
        {
            if (IsEquitableTopPaidReport(page.Text))
            {
                inReport = true;
                period = ParseEquitableTopPaidPeriod(page.Text);
            }

            if (!inReport)
            {
                continue;
            }

            foreach (var rawLine in page.Lines)
            {
                var line = ParsingHelpers.Clean(rawLine);
                var rowMatch = EquitableTopPaidRowRegex.Match(line);
                if (rowMatch.Success)
                {
                    var rowRank = Int(rowMatch, "rank") ?? 0;
                    if (rowRank > 10)
                    {
                        return;
                    }

                    if (rowRank >= 1)
                    {
                        AddEquitableTopPaidDrug(
                            workbook,
                            seen,
                            period,
                            page.PageNumber,
                            rowRank,
                            rowMatch.Groups["din"].Value,
                            rowMatch.Groups["name"].Value,
                            Money(rowMatch, "amount"));
                    }

                    rank = null;
                    din = "";
                    drugNameParts.Clear();
                    amount = null;
                    continue;
                }

                var rankMatch = EquitableTopPaidRankRegex.Match(line);
                if (rankMatch.Success)
                {
                    rank = Int(rankMatch, "rank");
                    din = "";
                    drugNameParts.Clear();
                    amount = null;
                    if (rank is > 10)
                    {
                        return;
                    }

                    continue;
                }

                if (!rank.HasValue || rank.Value < 1 || rank.Value > 10)
                {
                    continue;
                }

                if (string.IsNullOrWhiteSpace(din))
                {
                    if (EquitableTopPaidDinRegex.IsMatch(line))
                    {
                        din = line;
                    }

                    continue;
                }

                if (!amount.HasValue)
                {
                    var amountMatch = EquitableTopPaidAmountRegex.Match(line);
                    if (amountMatch.Success)
                    {
                        amount = Money(amountMatch, "value");
                    }
                    else if (!EquitableTopPaidPercentRegex.IsMatch(line))
                    {
                        drugNameParts.Add(line);
                    }

                    continue;
                }

                if (!EquitableTopPaidPercentRegex.IsMatch(line))
                {
                    continue;
                }

                var drugName = ParsingHelpers.Clean(string.Join(" ", drugNameParts));
                AddEquitableTopPaidDrug(
                    workbook,
                    seen,
                    period,
                    page.PageNumber,
                    rank.Value,
                    din,
                    drugName,
                    amount);

                rank = null;
                din = "";
                drugNameParts.Clear();
                amount = null;
            }
        }
    }

    private static void AddEquitableTopPaidDrug(
        WorkbookExtractionResult workbook,
        ISet<string> seen,
        string period,
        int pageNumber,
        int rank,
        string din,
        string drugName,
        decimal? amount)
    {
        var normalizedDin = ParsingHelpers.Clean(din);
        var normalizedName = ParsingHelpers.Clean(drugName);
        var key = $"{period}|{rank}|{normalizedDin}";
        if (string.IsNullOrWhiteSpace(normalizedName) || !seen.Add(key))
        {
            return;
        }

        workbook.DrugTrends.Add(new DrugTrend
        {
            Din = normalizedDin,
            DrugName = normalizedName,
            Period = period,
            NetAmountPaid = amount,
            SourceReference = $"Equitable Top Paid Drugs PDF page {pageNumber}"
        });
    }

    private static string ParseEquitableTopPaidPeriod(string text)
    {
        var match = EquitableTopPaidPeriodRegex.Match(text);
        return match.Success
            ? $"{match.Groups["from"].Value} - {match.Groups["to"].Value}"
            : "Current";
    }

    private static void ParseHealthClaimsSummary(List<PdfPageText> pages, WorkbookExtractionResult workbook)
    {
        foreach (var page in pages.Where(page => page.Text.Contains("Health claims", StringComparison.OrdinalIgnoreCase) &&
                                                 !page.Text.Contains("Health major services claims", StringComparison.OrdinalIgnoreCase)))
        {
            var twoPeriod = ParseTwoPeriodHeader(page.Text);
            if (twoPeriod.HasValue)
            {
                var foundTwoPeriodRows = false;
                foreach (var line in page.Lines.Select(ParsingHelpers.Clean))
                {
                    var match = TwoPeriodClaimsRowRegex.Match(line);
                    var category = match.Success ? NormalizeHealthCategory(match.Groups["category"].Value) : "";
                    if (!match.Success || !IsEquitableHealthClaimsCategory(category))
                    {
                        continue;
                    }

                    foundTwoPeriodRows = true;
                    AddEquitableHealthClaimsPeriod(
                        workbook,
                        category,
                        twoPeriod.Value.Prior,
                        Money(match, "priorPaid") ?? 0m,
                        Money(match, "priorPercent"),
                        page.PageNumber);
                    AddEquitableHealthClaimsPeriod(
                        workbook,
                        category,
                        twoPeriod.Value.Current,
                        Money(match, "currentPaid") ?? 0m,
                        Money(match, "currentPercent"),
                        page.PageNumber);
                }

                if (foundTwoPeriodRows)
                {
                    continue;
                }
            }

            var period = ParseExperiencePeriod(page.Text);
            foreach (var line in page.Lines.Select(ParsingHelpers.Clean))
            {
                var match = HealthClaimsRowRegex.Match(line);
                if (!match.Success)
                {
                    continue;
                }

                var category = NormalizeHealthCategory(match.Groups["category"].Value);
                var paid = Money(match, "paid") ?? 0m;
                var percent = Money(match, "percent");
                workbook.ClaimTrends.Add(new ClaimTrend
                {
                    Benefit = "Health",
                    Category = category,
                    Period = period,
                    Amount = paid,
                    PercentOfTotal = percent,
                    SourceReference = $"Equitable PDF page {page.PageNumber}"
                });
                workbook.EhcSpendingTrends.Add(new ServiceCategoryTrend
                {
                    Benefit = "Health",
                    Category = category,
                    CurrentPeriodLabel = period,
                    CurrentAmount = paid,
                    SourceReference = $"Equitable PDF page {page.PageNumber}"
                });
            }
        }
    }

    private static void ParseHealthMajorServicesClaims(List<PdfPageText> pages, WorkbookExtractionResult workbook)
    {
        foreach (var page in pages.Where(page => page.Text.Contains("Health major services claims", StringComparison.OrdinalIgnoreCase)))
        {
            var twoPeriod = ParseTwoPeriodHeader(page.Text);
            if (twoPeriod.HasValue)
            {
                var foundTwoPeriodRows = false;
                foreach (var line in page.Lines.Select(ParsingHelpers.Clean))
                {
                    if (IsEquitableParamedicalSubtotal(line))
                    {
                        break;
                    }

                    var match = TwoPeriodClaimsRowRegex.Match(line);
                    var category = match.Success ? NormalizeParamedicalCategory(match.Groups["category"].Value) : "";
                    if (!match.Success || !IsEquitableParamedicalCategory(category))
                    {
                        continue;
                    }

                    foundTwoPeriodRows = true;
                    var row = workbook.ParamedicalServiceTrends.FirstOrDefault(existing =>
                        existing.Category.Equals(category, StringComparison.OrdinalIgnoreCase));
                    if (row is null)
                    {
                        row = new ServiceCategoryTrend
                        {
                            Benefit = "Health",
                            Category = category,
                            SourceReference = $"Equitable PDF page {page.PageNumber}"
                        };
                        workbook.ParamedicalServiceTrends.Add(row);
                    }

                    ServiceCategoryHistory.Upsert(
                        row,
                        twoPeriod.Value.Prior,
                        Money(match, "priorPaid") ?? 0m,
                        $"Equitable PDF page {page.PageNumber}");
                    ServiceCategoryHistory.Upsert(
                        row,
                        twoPeriod.Value.Current,
                        Money(match, "currentPaid") ?? 0m,
                        $"Equitable PDF page {page.PageNumber}");
                }

                if (foundTwoPeriodRows)
                {
                    continue;
                }
            }

            var period = ParseExperiencePeriod(page.Text);
            foreach (var line in page.Lines.Select(ParsingHelpers.Clean))
            {
                if (IsEquitableParamedicalSubtotal(line))
                {
                    break;
                }

                var match = HealthMajorServicesRowRegex.Match(line);
                var category = match.Success ? NormalizeParamedicalCategory(match.Groups["category"].Value) : "";
                if (!match.Success || !IsEquitableParamedicalCategory(category))
                {
                    continue;
                }

                workbook.ParamedicalServiceTrends.Add(new ServiceCategoryTrend
                {
                    Benefit = "Health",
                    Category = category,
                    CurrentPeriodLabel = period,
                    CurrentAmount = Money(match, "paid") ?? 0m,
                    SourceReference = $"Equitable PDF page {page.PageNumber}"
                });
            }
        }
    }

    private static void AddEquitableHealthClaimsPeriod(
        WorkbookExtractionResult workbook,
        string category,
        string period,
        decimal paid,
        decimal? percent,
        int pageNumber)
    {
        workbook.ClaimTrends.Add(new ClaimTrend
        {
            Benefit = "Health",
            Category = category,
            Period = period,
            Amount = paid,
            PercentOfTotal = percent,
            SourceReference = $"Equitable PDF page {pageNumber}"
        });

        var row = workbook.EhcSpendingTrends.FirstOrDefault(existing =>
            existing.Category.Equals(category, StringComparison.OrdinalIgnoreCase));
        if (row is null)
        {
            row = new ServiceCategoryTrend
            {
                Benefit = "Health",
                Category = category,
                SourceReference = $"Equitable PDF page {pageNumber}"
            };
            workbook.EhcSpendingTrends.Add(row);
        }

        ServiceCategoryHistory.Upsert(row, period, paid, $"Equitable PDF page {pageNumber}");
    }

    private static bool IsEquitableHealthClaimsCategory(string category)
    {
        return !string.IsNullOrWhiteSpace(category) &&
               !category.Contains("Type of expense", StringComparison.OrdinalIgnoreCase) &&
               !category.Contains("Amount paid", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsEquitableParamedicalCategory(string category)
    {
        return !string.IsNullOrWhiteSpace(category) &&
               !category.Contains("Type of expense", StringComparison.OrdinalIgnoreCase) &&
               !IsEquitableParamedicalSubtotal(category);
    }

    private static bool IsEquitableParamedicalSubtotal(string value)
    {
        return ParsingHelpers.Clean(value)
            .StartsWith("Paramedical Sub Total", StringComparison.OrdinalIgnoreCase);
    }

    private static void ParseDentalClaimsTables(List<PdfPageText> pages, WorkbookExtractionResult workbook)
    {
        var dentalPages = pages
            .Where(page => page.Text.Contains("Dental claims", StringComparison.OrdinalIgnoreCase))
            .ToList();
        var hasBasicDetailPage = dentalPages.Any(page =>
            page.Text.Contains("Basic dental claims", StringComparison.OrdinalIgnoreCase));
        var hasSummaryPage = dentalPages.Any(page =>
            !page.Text.Contains("Basic dental claims", StringComparison.OrdinalIgnoreCase));

        foreach (var page in dentalPages)
        {
            var isBasicDetail = page.Text.Contains("Basic dental claims", StringComparison.OrdinalIgnoreCase);
            var twoPeriod = ParseTwoPeriodHeader(page.Text);
            if (twoPeriod.HasValue)
            {
                foreach (var line in page.Lines.Select(ParsingHelpers.Clean))
                {
                    var match = TwoPeriodClaimsRowRegex.Match(line);
                    if (!match.Success)
                    {
                        continue;
                    }

                    var category = NormalizeDentalCategory(match.Groups["category"].Value);
                    if (!ShouldIncludeEquitableDentalCategory(category, isBasicDetail, hasSummaryPage, hasBasicDetailPage))
                    {
                        continue;
                    }

                    AddEquitableDentalClaimsPeriod(
                        workbook,
                        category,
                        twoPeriod.Value.Prior,
                        Money(match, "priorPaid") ?? 0m,
                        page.PageNumber);
                    AddEquitableDentalClaimsPeriod(
                        workbook,
                        category,
                        twoPeriod.Value.Current,
                        Money(match, "currentPaid") ?? 0m,
                        page.PageNumber);
                }

                continue;
            }

            var period = ParseExperiencePeriod(page.Text);
            foreach (var line in page.Lines.Select(ParsingHelpers.Clean))
            {
                var match = HealthClaimsRowRegex.Match(line);
                if (!match.Success)
                {
                    continue;
                }

                var category = NormalizeDentalCategory(match.Groups["category"].Value);
                if (!ShouldIncludeEquitableDentalCategory(category, isBasicDetail, hasSummaryPage, hasBasicDetailPage))
                {
                    continue;
                }

                AddEquitableDentalClaimsPeriod(
                    workbook,
                    category,
                    period,
                    Money(match, "paid") ?? 0m,
                    page.PageNumber);
            }
        }
    }

    private static bool ShouldIncludeEquitableDentalCategory(
        string category,
        bool isBasicDetail,
        bool hasSummaryPage,
        bool hasBasicDetailPage)
    {
        if (string.IsNullOrWhiteSpace(category) ||
            category.Contains("Type of expense", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        if (isBasicDetail)
        {
            return !category.Equals("Total", StringComparison.OrdinalIgnoreCase) || !hasSummaryPage;
        }

        return !category.Equals("Basic", StringComparison.OrdinalIgnoreCase) || !hasBasicDetailPage;
    }

    private static void AddEquitableDentalClaimsPeriod(
        WorkbookExtractionResult workbook,
        string category,
        string period,
        decimal paid,
        int pageNumber)
    {
        workbook.ClaimTrends.Add(new ClaimTrend
        {
            Benefit = "Dental",
            Category = category,
            Period = period,
            Amount = paid,
            SourceReference = $"Equitable PDF page {pageNumber}"
        });

        var row = workbook.DentalSpendingTrends.FirstOrDefault(existing =>
            existing.Category.Equals(category, StringComparison.OrdinalIgnoreCase));
        if (row is null)
        {
            row = new ServiceCategoryTrend
            {
                Benefit = "Dental",
                Category = category,
                SourceReference = $"Equitable PDF page {pageNumber}"
            };
            workbook.DentalSpendingTrends.Add(row);
        }

        ServiceCategoryHistory.Upsert(row, period, paid, $"Equitable PDF page {pageNumber}");
    }

    private static void ParseDentalExperienceSummary(List<PdfPageText> pages, WorkbookExtractionResult workbook)
    {
        if (workbook.DentalSpendingTrends.Count > 0)
        {
            return;
        }

        foreach (var page in pages.Where(page => page.Text.Contains("Experience Summary", StringComparison.OrdinalIgnoreCase)))
        {
            foreach (var line in page.Lines.Select(ParsingHelpers.Clean))
            {
                var match = ExperienceSummaryRegex.Match(line);
                if (!match.Success || !match.Groups["benefit"].Value.Equals("Dental", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var paid = Money(match, "paid") ?? 0m;
                var period = ParseExperiencePeriod(page.Text);
                workbook.ClaimTrends.Add(new ClaimTrend
                {
                    Benefit = "Dental",
                    Category = "Dental",
                    Period = period,
                    Amount = paid,
                    SourceReference = $"Equitable PDF page {page.PageNumber}"
                });
                workbook.DentalSpendingTrends.Add(new ServiceCategoryTrend
                {
                    Benefit = "Dental",
                    Category = "Dental",
                    CurrentPeriodLabel = period,
                    CurrentAmount = paid,
                    SourceReference = $"Equitable PDF page {page.PageNumber}"
                });
            }
        }
    }

    private static (string Prior, string Current)? ParseTwoPeriodHeader(string text)
    {
        var match = TwoPeriodHeaderRegex.Match(text);
        if (!match.Success)
        {
            return null;
        }

        return (
            FormatEquitablePeriod(match.Groups["priorFrom"].Value, match.Groups["priorTo"].Value),
            FormatEquitablePeriod(match.Groups["currentFrom"].Value, match.Groups["currentTo"].Value));
    }

    private static string FormatEquitablePeriod(string from, string to)
    {
        return $"{FormatEquitableTableDate(from)} - {FormatEquitableTableDate(to)}";
    }

    private static string FormatEquitableTableDate(string value)
    {
        return ParseEquitableTableDate(value)?.ToString("dd-MMM-yyyy", CultureInfo.InvariantCulture) ?? ParsingHelpers.Clean(value);
    }

    private static DateOnly? ParseEquitableTableDate(string value)
    {
        var formats = new[] { "dd MMM yyyy", "d MMM yyyy", "dd-MMM-yyyy", "d-MMM-yyyy" };
        return DateOnly.TryParseExact(
            ParsingHelpers.Clean(value),
            formats,
            CultureInfo.InvariantCulture,
            DateTimeStyles.None,
            out var date)
            ? date
            : null;
    }

    private static string ParseExperiencePeriod(string text)
    {
        var match = ExperiencePeriodRegex.Match(text);
        if (!match.Success)
        {
            match = PlainExperiencePeriodRegex.Match(text);
        }

        return match.Success
            ? $"{match.Groups["from"].Value} - {match.Groups["to"].Value}"
            : "Current";
    }

    private static string NormalizeBenefit(string value)
    {
        value = ParsingHelpers.Clean(value);
        return value switch
        {
            "AD&D" or "AD+D" => "Accidental Death and Dismemberment",
            "Health" => "Extended Health Care",
            "Employee Critical Illness" => "Critical Illness",
            "Long Term Disability" or "Long-Term Disability" => "LTD",
            "Employee and Family Assistance Program" or "EAP" => "Employee and Family Assistance Program",
            _ => value
        };
    }

    private static string NormalizeHealthCategory(string value)
    {
        value = ParsingHelpers.Clean(value);
        return value switch
        {
            "Drug" => "Drugs",
            "Hospital" or "HOSPITAL" => "Hospital",
            "Total" or "TOTAL" => "Total",
            _ => value
        };
    }

    private static string NormalizeDentalCategory(string value)
    {
        value = ParsingHelpers.Clean(value);
        return value switch
        {
            "Preventive" => "Preventative",
            "Periodontal" => "Periodontic",
            "Oral Surgery" => "Oral Surgical",
            "Total" or "TOTAL" => "Total",
            _ => value
        };
    }

    private static string NormalizeParamedicalCategory(string value)
    {
        value = ParsingHelpers.Clean(value);
        return value switch
        {
            "Clinical Counselors" => "Clinical Counsellors",
            _ => value
        };
    }

    private static void ParseDemographics(List<PdfPageText> pages, PdfExtractionResult result)
    {
        foreach (var page in pages)
        {
            if (page.Text.Contains("Changes in life", StringComparison.OrdinalIgnoreCase))
            {
                AddEquitableDemographicRows(
                    page,
                    "Life",
                    result.LifeDemographics,
                    page.Lines.SkipWhile(line => !line.Contains("Changes in life demographics", StringComparison.OrdinalIgnoreCase)),
                    parseVolumes: false);

                AddEquitableDemographicRows(
                    page,
                    "Life",
                    result.LifeDemographics,
                    page.Lines.SkipWhile(line => !line.Contains("Changes in life volumes", StringComparison.OrdinalIgnoreCase)),
                    parseVolumes: true);
            }

            if (page.Text.Contains("Changes in LTD", StringComparison.OrdinalIgnoreCase))
            {
                AddEquitableDemographicRows(
                    page,
                    "LTD",
                    result.LtdDemographics,
                    page.Lines.SkipWhile(line => !line.Contains("Changes in LTD demographics", StringComparison.OrdinalIgnoreCase)),
                    parseVolumes: false);

                AddEquitableDemographicRows(
                    page,
                    "LTD",
                    result.LtdDemographics,
                    page.Lines.SkipWhile(line => !line.Contains("Changes in LTD volumes", StringComparison.OrdinalIgnoreCase)),
                    parseVolumes: true);
            }
        }

        SortDemographicRows(result.LifeDemographics);
        SortDemographicRows(result.LtdDemographics);
    }

    private static void AddEquitableDemographicRows(
        PdfPageText page,
        string benefit,
        List<DemographicLine> rows,
        IEnumerable<string> sourceLines,
        bool parseVolumes)
    {
        foreach (var rawLine in sourceLines.Select(ParsingHelpers.Clean))
        {
            if (rawLine.StartsWith("Total ", StringComparison.OrdinalIgnoreCase) ||
                rawLine.Contains(" | ", StringComparison.Ordinal))
            {
                continue;
            }

            var match = parseVolumes
                ? EquitableLifeVolumeRowRegex.Match(rawLine)
                : EquitableLivesRowRegex.Match(rawLine);
            if (!match.Success)
            {
                continue;
            }

            var ageBand = NormalizeEquitableAgeBand(match.Groups["priorAge"].Value);
            if (string.IsNullOrWhiteSpace(ageBand))
            {
                continue;
            }

            var row = rows.FirstOrDefault(item => item.AgeBand.Equals(ageBand, StringComparison.OrdinalIgnoreCase));
            if (row is null)
            {
                row = new DemographicLine
                {
                    Benefit = benefit,
                    AgeBand = ageBand,
                    SourceReference = $"Equitable PDF page {page.PageNumber}"
                };
                rows.Add(row);
            }

            if (parseVolumes)
            {
                row.PriorMaleVolume = Add(row.PriorMaleVolume, Money(match, "priorMale"));
                row.PriorFemaleVolume = Add(row.PriorFemaleVolume, Money(match, "priorFemale"));
                row.CurrentMaleVolume = Add(row.CurrentMaleVolume, Money(match, "currentMale"));
                row.CurrentFemaleVolume = Add(row.CurrentFemaleVolume, Money(match, "currentFemale"));
            }
            else
            {
                row.PriorMaleLives = Add(row.PriorMaleLives, Int(match, "priorMale"));
                row.PriorFemaleLives = Add(row.PriorFemaleLives, Int(match, "priorFemale"));
                row.CurrentMaleLives = Add(row.CurrentMaleLives, Int(match, "currentMale"));
                row.CurrentFemaleLives = Add(row.CurrentFemaleLives, Int(match, "currentFemale"));
            }
        }
    }

    private static void SortDemographicRows(List<DemographicLine> rows)
    {
        var sorted = rows
            .OrderBy(row => EquitableAgeBandSort(row.AgeBand))
            .ThenBy(row => row.AgeBand, StringComparer.OrdinalIgnoreCase)
            .ToList();

        rows.Clear();
        rows.AddRange(sorted);
    }

    private static string NormalizeEquitableAgeBand(string ageBand)
    {
        var normalized = ParsingHelpers.Clean(ageBand);
        return normalized switch
        {
            "0-19" or "20-24" => "Under 25",
            "70+" => "70 & over",
            _ => normalized
        };
    }

    private static int EquitableAgeBandSort(string ageBand)
    {
        return NormalizeEquitableAgeBand(ageBand) switch
        {
            "Under 25" => 0,
            "25-29" => 1,
            "30-34" => 2,
            "35-39" => 3,
            "40-44" => 4,
            "45-49" => 5,
            "50-54" => 6,
            "55-59" => 7,
            "60-64" => 8,
            "65-69" => 9,
            "70 & over" => 10,
            _ => 99
        };
    }

    private static decimal? Add(decimal? current, decimal? value)
    {
        if (!current.HasValue)
        {
            return value;
        }

        return current.Value + (value ?? 0m);
    }

    private static int? Add(int? current, int? value)
    {
        if (!current.HasValue)
        {
            return value;
        }

        return current.Value + (value ?? 0);
    }

    private static bool HasWorkbookData(WorkbookExtractionResult workbook)
    {
        return workbook.ClaimTrends.Count > 0 ||
               workbook.EhcSpendingTrends.Count > 0 ||
               workbook.DentalSpendingTrends.Count > 0 ||
               workbook.ParamedicalServiceTrends.Count > 0 ||
               workbook.DrugTrends.Count > 0 ||
               workbook.IndicationTrends.Count > 0;
    }

    private static decimal? SumPremium(IEnumerable<RateLine> lines, string benefit, bool current)
    {
        var values = lines
            .Where(line => line.Benefit.Equals(benefit, StringComparison.OrdinalIgnoreCase) ||
                           line.Benefit.StartsWith(benefit + " ", StringComparison.OrdinalIgnoreCase))
            .Select(line => current ? line.CurrentPremium : line.RenewalPremium)
            .Where(value => value.HasValue)
            .Select(value => value!.Value)
            .ToList();

        return values.Count == 0 ? null : decimal.Round(values.Sum(), 2);
    }

    private static decimal? Money(Match match, string groupName)
    {
        return match.Success ? ParsingHelpers.ParseDecimal(match.Groups[groupName].Value) : null;
    }

    private static int? Int(Match match, string groupName)
    {
        return ParsingHelpers.ParseInt(match.Groups[groupName].Value);
    }
}

public sealed class CanadaLifePdfParser
{
    private static readonly Regex PolicyNumberRegex = new(@"Policy Number:\s*(?<value>[\d-]+)", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex GroupNetPolicyNumberRegex = new(@"\bPolicy\(s\)\s*(?<value>[\d-]+)", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex EffectiveDateRegex = new(@"Effective:\s*(?<date>[A-Za-z]+\s+\d{1,2},\s+\d{4})", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex PresentedByRegex = new(@"Presented by:\s*(?<value>.+?)(?:\s+\*|$)", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex CurrentExperiencePeriodRegex = new(@"Current Experience Period\s+(?<period>[A-Za-z]+\s+\d{2},\s+\d{4}\s+-\s+[A-Za-z]+\s+\d{2},\s+\d{4})", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex PriorExperiencePeriodRegex = new(@"Prior Year Experience Period\s+(?<period>[A-Za-z]+\s+\d{2},\s+\d{4}\s+-\s+[A-Za-z]+\s+\d{2},\s+\d{4})", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex VolumeRateRegex = new(@"^(?<benefit>Employee Term Life|AD&D|Accidental Death\s*&\s*Dismemberment|LTD|Long[\s-]+Term\s+Disability|STD|Optional Life|Dependent Life|Critical Illness|Employee Critical Illness|Virtual Health(?:care)?|Teladoc Medical Experts|Global Medical Assistance)\s+(?<volume>[\d,]+|-)\s+\$?(?<currentRate>-?[\d,]+\.\d+)\s+\$?(?<currentPremium>-?[\d,]+\.\d{2})\s+\$?(?<renewalRate>-?[\d,]+\.\d+)\s+\$?(?<renewalPremium>-?[\d,]+\.\d{2})\s+(?<change>-?[\d.]+)%$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex GenericVolumeRateRegex = new(@"^(?<benefit>[A-Za-z][A-Za-z0-9 &'./()+-]+?)\s+(?<volume>[\d,]+|-)\s+\$?(?<currentRate>-?[\d,]+\.\d+)\s+\$?(?<currentPremium>-?[\d,]+\.\d{2})\s+\$?(?<renewalRate>-?[\d,]+\.\d+)\s+\$?(?<renewalPremium>-?[\d,]+\.\d{2})\s+(?<change>-?[\d.]+)%$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex TierRateRegex = new(@"^(?<tier>Single|Family|Couple|Single Parent|Employee|Employee Only|Employee \+ Spouse|Employee \+ Child(?:ren)?|Composite)\s+(?<lives>\d+|-)\s+\$?(?<currentRate>-?[\d,]+\.\d+)\s+\$?(?<currentPremium>-?[\d,]+\.\d{2})\s+\$?(?<renewalRate>-?[\d,]+\.\d+)\s+\$?(?<renewalPremium>-?[\d,]+\.\d{2})\s+(?<change>-?[\d.]+)%$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex TotalPremiumRegex = new(@"^Total:\s+\$?(?<current>[\d,]+\.\d{2})\s+\$?(?<renewal>[\d,]+\.\d{2})\s+(?<change>-?[\d.]+)%$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex ExperienceRowRegex = new(@"^(?<period>[A-Za-z]+\s+\d{2},\s+\d{4}\s+-\s+[A-Za-z]+\s+\d{2},\s+\d{4})\s+(?<premium>\(?-?\$?[\d,]+\.\d{2}\)?)\s+(?<paid>\(?-?\$?[\d,]+\.\d{2}\)?)\s+(?<incurred>\(?-?\$?[\d,]+\.\d{2}\)?)\s+(?<ratio>-?[\d.]+)%(?:\s+(?<weight>-?[\d.]+)\s*(?<weightUnit>%?))?$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex ClaimsDistributionBenefitRegex = new(@"^Benefit\s+(?<value>.+)$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex ClaimsDistributionPeriodRegex = new(@"^Period\s+(?<value>.+)$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex ClaimsDistributionRowRegex = new(@"^(?<category>[A-Za-z][A-Za-z ,/&().+-]+?)\s+\$?(?<amount>[\d,]+(?:\.\d{2})?)\s+(?<percent>[\d.]+)%\s+(?<normPercent>[\d.]+)%\s+(?<occurrences>\d+)\s+(?<occurrencePercent>[\d.]+)%\s+(?<occurrenceNormPercent>[\d.]+)%\s+\$?(?<average>[\d,]+(?:\.\d{2})?)$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex ParamedicalRowRegex = new(@"^(?<category>[A-Za-z][A-Za-z ,/&().+-]+?)\s+\$?(?<amount>[\d,]+(?:\.\d{2})?)\s+(?<percent>[\d.]+)%$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex GroupNetTopDrugRowRegex = new(@"^(?<descriptor>.+?)\s+(?<din>\d{7,8})\s+(?<rank>\d{1,3})\s+\$?(?<amount>[\d,]+(?:\.\d{2})?)\s+(?<amountPercent>[\d.]+)%\s+(?<claims>\d+)\s+(?<claimsPercent>[\d.]+)%\s+(?<claimRank>\d{1,3})\s+\$?(?<ingredientCost>[\d,]+(?:\.\d{2})?)$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex GroupNetTreatmentRegex = new(@"\b(?<value>[A-Z][a-z]+(?:\s+(?:[A-Z][a-z]+|and|or|of|the|&|/|-))*?)$", RegexOptions.Compiled);
    private static readonly Regex KnownParamedicalRowRegex = new(@"(?<category>Registered Massage Therapist|Massage Therapy|Masseur|Massage|Physiotherapist|Physiotherapy|Chiropractic Treatment|Chiropractor|Chiropractic|Naturopath Treatment|Naturopath|Psychologist's Office Visit|Psychologist Office Visit|Psychologist|Psychology|Psychotherapist|Clinical Counsellor|Social Worker|Acupuncture|Osteopath|Podiatrist|Chiropodist|Dietitian|Dietician|Speech Therapy)\s+\$?(?<amount>[\d,]+(?:\.\d{2})?)\s+(?<percent>[\d.]+)%", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex DemographicRowRegex = new(@"^(?<age><\s*25|Under\s+25|25-29|30-34|35-39|40-44|45-49|50-54|55-59|60-64|60\s*(?:\+|&\s*over|and\s+over)|65-69\+?|65\s*(?:\+|&\s*over|and\s+over))\s+(?<maleLives>\d+|-)\s+(?<maleVolume>[\d,]+|-)\s+(?<femaleLives>\d+|-)\s+(?<femaleVolume>[\d,]+|-)\s+(?<otherLives>\d+|-)\s+(?<otherVolume>[\d,]+|-)\s+(?<undisclosedLives>\d+|-)\s+(?<undisclosedVolume>[\d,]+|-)$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex PoolingChargeRegex = new(@"\bPooling\s+Charge\s+(?<value>[\d.]+)%", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex PoolingLevelRegex = new(@"\bPooling\s+(?:Level|Threshold)\s+\$?\s*(?<value>[\d,\s]+)", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex SplitMoneyTokenRegex = new(@"(?<![\d,])(?<prefix>\d{1,2})\s+(?<rest>\d{1,2},\d{3})(?![\d,])", RegexOptions.Compiled);
    private static readonly Regex SpacedCommaMoneyTokenRegex = new(@"(?<=\d)\s+,\s*(?=\d{3}(?:\D|$))", RegexOptions.Compiled);

    public static bool IsCanadaLifeRenewal(string text)
    {
        return text.Contains("Canada Life", StringComparison.OrdinalIgnoreCase) ||
               text.Contains("Great-West Life Assurance Co.", StringComparison.OrdinalIgnoreCase) ||
               text.Contains("Summary of Renewal Rates and Premium", StringComparison.OrdinalIgnoreCase);
    }

    public CarrierPdfParseResult Parse(string pdfPath, List<PdfPageText> pages, string allText)
    {
        var isStandaloneTopDrugReport = IsStandaloneTopDrugReport(allText);
        var pdf = new PdfExtractionResult
        {
            SourcePath = pdfPath,
            Carrier = "Canada Life",
            ClientName = isStandaloneTopDrugReport ? "" : ParseClientName(pdfPath, pages, allText),
            PolicyNumber = ParsePolicyNumber(allText),
            RenewalDate = ParseRenewalDate(allText)
        };

        ParseRateLines(pages, pdf);
        ParseExperienceLines(pages, pdf);
        ParseDemographics(pages, pdf);
        ParsePoolingArrangement(pages, pdf);

        var workbook = ParseWorkbookDetailFromPdf(pdfPath, pages);

        if (pdf.RateLines.Count == 0)
        {
            pdf.Warnings.Add("No Canada Life renewal rate rows were extracted from the PDF.");
        }

        if (pdf.ExperienceLines.Count == 0)
        {
            pdf.Warnings.Add("No Canada Life experience history rows were extracted from the PDF.");
        }

        if (workbook.ClaimTrends.Count == 0 && workbook.DrugTrends.Count == 0)
        {
            workbook.Warnings.Add("No Canada Life claims distribution rows were extracted from the PDF.");
        }

        return new CarrierPdfParseResult
        {
            Pdf = pdf,
            WorkbookFromPdf = workbook,
            Carrier = "Canada Life"
        };
    }

    private static bool IsCanadaLifeTopDrugReport(string text)
    {
        return text.Contains("Top 100 Drugs By Amount Paid", StringComparison.OrdinalIgnoreCase) ||
               text.Contains("Top 10 Drugs By Amount Paid", StringComparison.OrdinalIgnoreCase) ||
               text.Contains("Top Drugs By Amount Paid", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsStandaloneTopDrugReport(string text)
    {
        return IsCanadaLifeTopDrugReport(text) &&
               !text.Contains("Summary of Renewal Rates and Premium", StringComparison.OrdinalIgnoreCase) &&
               !text.Contains("Presented by:", StringComparison.OrdinalIgnoreCase);
    }

    private static string ParseClientName(string pdfPath, List<PdfPageText> pages, string allText)
    {
        var fileName = Path.GetFileNameWithoutExtension(pdfPath);
        var fileNameMatch = Regex.Match(fileName, @"Renewal\s+-\s+\d+\s+(?<client>.+?)\s+-\s+Policy", RegexOptions.IgnoreCase);
        if (fileNameMatch.Success)
        {
            return ParsingHelpers.Clean(fileNameMatch.Groups["client"].Value);
        }

        fileNameMatch = Regex.Match(
            fileName,
            @"^(?:Revised\s+)?(?:Canada Life\s+)?Renewal\s*-\s*(?:Policy\s*#?\s*)?(?<policy>[\d-]{5,})\s*-\s*(?<client>.+?)\s*-\s*(?:\d{8}|\d{4})$",
            RegexOptions.IgnoreCase);
        if (fileNameMatch.Success)
        {
            return ParsingHelpers.Clean(fileNameMatch.Groups["client"].Value);
        }

        fileNameMatch = Regex.Match(
            fileName,
            @"^(?:Canada Life\s+)?Renewal\s*-\s*(?:\d{5,}\s*-\s*)?(?<client>.+?)\s*-\s*(?:Policy\s*#?\s*\d+(?:\s*-\s*\d{4})?|\d{5,}\s*-\s*\d{4}|\d{8})$",
            RegexOptions.IgnoreCase);
        if (fileNameMatch.Success)
        {
            return ParsingHelpers.Clean(fileNameMatch.Groups["client"].Value);
        }

        var headerMatch = Regex.Match(allText, @"\b\d{5,}\s+(?<client>[A-Z0-9][A-Z0-9 &'.,-]+?)\s+Policy Number:", RegexOptions.IgnoreCase);
        if (headerMatch.Success)
        {
            return ParsingHelpers.Clean(headerMatch.Groups["client"].Value);
        }

        var firstPageLines = pages.FirstOrDefault()?.Lines.Select(ParsingHelpers.Clean).Where(line => line.Length > 0).ToList() ?? [];
        for (var index = 1; index < firstPageLines.Count; index++)
        {
            if (firstPageLines[index].StartsWith("Policy Number:", StringComparison.OrdinalIgnoreCase))
            {
                var candidate = firstPageLines[index - 1];
                if (IsPlausibleClientName(candidate))
                {
                    return candidate;
                }
            }
        }

        for (var index = 0; index < firstPageLines.Count; index++)
        {
            if (firstPageLines[index].StartsWith("Presented by:", StringComparison.OrdinalIgnoreCase))
            {
                return firstPageLines.Skip(index + 1).FirstOrDefault(IsPlausibleClientName) ?? "";
            }
        }

        return firstPageLines.LastOrDefault(IsPlausibleClientName) ?? "";
    }

    private static bool IsPlausibleClientName(string value)
    {
        var candidate = ParsingHelpers.Clean(value);
        return candidate.Length is > 1 and <= 160 &&
               candidate.Any(char.IsLetter) &&
               !candidate.StartsWith("*", StringComparison.OrdinalIgnoreCase) &&
               !candidate.StartsWith("Policy Number:", StringComparison.OrdinalIgnoreCase) &&
               !candidate.StartsWith("Effective:", StringComparison.OrdinalIgnoreCase) &&
               !candidate.StartsWith("Presented by:", StringComparison.OrdinalIgnoreCase) &&
               !candidate.StartsWith("Total", StringComparison.OrdinalIgnoreCase) &&
               !candidate.StartsWith("Summary of Renewal", StringComparison.OrdinalIgnoreCase) &&
               !candidate.Equals("Benefit", StringComparison.OrdinalIgnoreCase) &&
               !candidate.Contains('$') &&
               !candidate.Contains('%');
    }

    private static string ParsePolicyNumber(string text)
    {
        var policyNumber = MatchValue(PolicyNumberRegex, text);
        return string.IsNullOrWhiteSpace(policyNumber)
            ? MatchValue(GroupNetPolicyNumberRegex, text)
            : policyNumber;
    }

    private static DateOnly? ParseRenewalDate(string text)
    {
        var match = EffectiveDateRegex.Match(text);
        return match.Success &&
               DateOnly.TryParse(match.Groups["date"].Value, CultureInfo.InvariantCulture, DateTimeStyles.None, out var effectiveDate)
            ? effectiveDate
            : null;
    }

    private static void ParseRateLines(List<PdfPageText> pages, PdfExtractionResult result)
    {
        var page = pages.FirstOrDefault(item => item.Text.Contains("Summary of Renewal Rates and Premium", StringComparison.OrdinalIgnoreCase));
        if (page is null)
        {
            return;
        }

        var currentBenefit = "";
        var currentAccountClass = "";
        foreach (var rawLine in page.Lines.Select(ParsingHelpers.Clean))
        {
            if (IsRateHeaderLine(rawLine))
            {
                continue;
            }

            var totalMatch = TotalPremiumRegex.Match(rawLine);
            if (totalMatch.Success)
            {
                result.CurrentMonthlyPremium = Money(totalMatch, "current");
                result.RenewalMonthlyPremium = Money(totalMatch, "renewal");
                result.MonthlyPremiumChange = result.RenewalMonthlyPremium - result.CurrentMonthlyPremium;
                result.MonthlyPremiumChangePercent = Money(totalMatch, "change");
                continue;
            }

            if (IsBenefitHeading(rawLine))
            {
                currentBenefit = NormalizeBenefit(rawLine);
                currentAccountClass = "";
                continue;
            }

            if (TryParseAccountClassHeading(rawLine, out var accountClass))
            {
                currentAccountClass = accountClass;
                continue;
            }

            var rateLine = TryParseRateLine(
                rawLine,
                page.PageNumber,
                currentBenefit,
                currentAccountClass,
                out var benefitFromLine);
            if (rateLine is not null)
            {
                result.RateLines.Add(rateLine);
                if (!string.IsNullOrWhiteSpace(benefitFromLine))
                {
                    currentBenefit = benefitFromLine;
                    currentAccountClass = rateLine.AccountClass;
                }
                else if (!string.IsNullOrWhiteSpace(rateLine.AccountClass))
                {
                    currentAccountClass = rateLine.AccountClass;
                }
            }
        }

        result.CurrentEhcPremium = SumPremium(result.RateLines, "Extended Health Care", current: true);
        result.RenewalEhcPremium = SumPremium(result.RateLines, "Extended Health Care", current: false);
        result.CurrentDentalPremium = SumPremium(result.RateLines, "Dental", current: true);
        result.RenewalDentalPremium = SumPremium(result.RateLines, "Dental", current: false);
        FillMonthlyPremiumTotalsFromRateLines(result);
    }

    private static void ParsePoolingArrangement(List<PdfPageText> pages, PdfExtractionResult result)
    {
        foreach (var page in pages.Where(page => page.Text.Contains("Pooling", StringComparison.OrdinalIgnoreCase)))
        {
            foreach (var line in page.Lines.Select(ParsingHelpers.Clean))
            {
                var chargeMatch = PoolingChargeRegex.Match(line);
                if (chargeMatch.Success)
                {
                    result.PoolingCharge = Money(chargeMatch, "value");
                }

                var levelMatch = PoolingLevelRegex.Match(line);
                if (levelMatch.Success)
                {
                    result.PoolingThreshold = Money(levelMatch, "value");
                }
            }

            result.PoolingCharge ??= Money(PoolingChargeRegex.Match(page.Text), "value");
            result.PoolingThreshold ??= Money(PoolingLevelRegex.Match(page.Text), "value");
        }
    }

    private static bool IsRateHeaderLine(string line)
    {
        return line.Length == 0 ||
               line.StartsWith("Benefit Volume", StringComparison.OrdinalIgnoreCase) ||
               line.StartsWith("Current Renewal", StringComparison.OrdinalIgnoreCase) ||
               line.StartsWith("Summary of Renewal Rates", StringComparison.OrdinalIgnoreCase) ||
               line.StartsWith("*", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsBenefitHeading(string line)
    {
        return line.Equals("Employee Term Life", StringComparison.OrdinalIgnoreCase) ||
               line.Equals("AD&D", StringComparison.OrdinalIgnoreCase) ||
               line.Equals("Accidental Death & Dismemberment", StringComparison.OrdinalIgnoreCase) ||
               line.Equals("Dependent Life", StringComparison.OrdinalIgnoreCase) ||
               line.Equals("LTD", StringComparison.OrdinalIgnoreCase) ||
               line.Equals("Long Term Disability", StringComparison.OrdinalIgnoreCase) ||
               line.Equals("Long-Term Disability", StringComparison.OrdinalIgnoreCase) ||
               line.Equals("STD", StringComparison.OrdinalIgnoreCase) ||
               line.Equals("Optional Life", StringComparison.OrdinalIgnoreCase) ||
               line.Equals("Critical Illness", StringComparison.OrdinalIgnoreCase) ||
               line.Equals("Employee Critical Illness", StringComparison.OrdinalIgnoreCase) ||
               line.Equals("Contact", StringComparison.OrdinalIgnoreCase) ||
               line.Equals("EAP", StringComparison.OrdinalIgnoreCase) ||
               line.Equals("Employee and Family Assistance Program", StringComparison.OrdinalIgnoreCase) ||
               line.Equals("Healthcare", StringComparison.OrdinalIgnoreCase) ||
               line.Equals("Health", StringComparison.OrdinalIgnoreCase) ||
               line.Equals("Medical", StringComparison.OrdinalIgnoreCase) ||
               line.Equals("Drugs", StringComparison.OrdinalIgnoreCase) ||
               line.Equals("Dental", StringComparison.OrdinalIgnoreCase) ||
               line.Equals("Vision", StringComparison.OrdinalIgnoreCase) ||
               line.Equals("Visioncare", StringComparison.OrdinalIgnoreCase);
    }

    private static bool TryParseAccountClassHeading(string line, out string accountClass)
    {
        if (line.Equals("All Classes", StringComparison.OrdinalIgnoreCase))
        {
            accountClass = "All Classes";
            return true;
        }

        if (Regex.IsMatch(line, @"^Class\s+[A-Za-z0-9][A-Za-z0-9 ,&/-]*$", RegexOptions.IgnoreCase))
        {
            accountClass = ParsingHelpers.Clean(line);
            return true;
        }

        accountClass = "";
        return false;
    }

    private static RateLine? TryParseRateLine(
        string line,
        int pageNumber,
        string currentBenefit,
        string currentAccountClass,
        out string benefitFromLine)
    {
        benefitFromLine = "";
        var match = VolumeRateRegex.Match(line);
        if (match.Success)
        {
            benefitFromLine = NormalizeBenefit(match.Groups["benefit"].Value);
            return new RateLine
            {
                Benefit = benefitFromLine,
                Volume = Money(match, "volume"),
                CurrentRate = Money(match, "currentRate"),
                CurrentPremium = Money(match, "currentPremium"),
                RenewalRate = Money(match, "renewalRate"),
                RenewalPremium = Money(match, "renewalPremium"),
                PercentChange = Money(match, "change"),
                SourceReference = $"Canada Life PDF page {pageNumber}"
            };
        }

        match = TierRateRegex.Match(line);
        if (match.Success && !string.IsNullOrWhiteSpace(currentBenefit))
        {
            return new RateLine
            {
                Benefit = currentBenefit,
                AccountClass = currentAccountClass,
                TierOrClass = match.Groups["tier"].Value,
                Lives = Int(match, "lives"),
                CurrentRate = Money(match, "currentRate"),
                CurrentPremium = Money(match, "currentPremium"),
                RenewalRate = Money(match, "renewalRate"),
                RenewalPremium = Money(match, "renewalPremium"),
                PercentChange = Money(match, "change"),
                SourceReference = $"Canada Life PDF page {pageNumber}"
            };
        }

        match = GenericVolumeRateRegex.Match(line);
        if (!match.Success)
        {
            return null;
        }

        var genericLabel = ParsingHelpers.Clean(match.Groups["benefit"].Value);
        if (!string.IsNullOrWhiteSpace(currentBenefit) &&
            TryParseAccountClassHeading(genericLabel, out var accountClassFromLine))
        {
            return new RateLine
            {
                Benefit = currentBenefit,
                AccountClass = accountClassFromLine,
                Volume = Money(match, "volume"),
                CurrentRate = Money(match, "currentRate"),
                CurrentPremium = Money(match, "currentPremium"),
                RenewalRate = Money(match, "renewalRate"),
                RenewalPremium = Money(match, "renewalPremium"),
                PercentChange = Money(match, "change"),
                SourceReference = $"Canada Life PDF page {pageNumber}"
            };
        }

        benefitFromLine = NormalizeBenefit(genericLabel);
        return new RateLine
        {
            Benefit = benefitFromLine,
            Volume = Money(match, "volume"),
            CurrentRate = Money(match, "currentRate"),
            CurrentPremium = Money(match, "currentPremium"),
            RenewalRate = Money(match, "renewalRate"),
            RenewalPremium = Money(match, "renewalPremium"),
            PercentChange = Money(match, "change"),
            SourceReference = $"Canada Life PDF page {pageNumber}"
        };
    }

    private static void ParseExperienceLines(List<PdfPageText> pages, PdfExtractionResult result)
    {
        foreach (var page in pages.Where(page => page.Text.Contains("Experience Analysis", StringComparison.OrdinalIgnoreCase)))
        {
            var benefit = ParseExperienceBenefit(page.Text);
            if (string.IsNullOrWhiteSpace(benefit))
            {
                continue;
            }

            var factors = ParseExperienceFactors(page.Text);
            var pageRows = new List<(ExperienceLine Line, bool WeightIsPercent)>();
            foreach (var line in page.Lines.Select(ParsingHelpers.Clean))
            {
                var match = ExperienceRowRegex.Match(line);
                if (!match.Success)
                {
                    continue;
                }

                pageRows.Add((new ExperienceLine
                {
                    Benefit = benefit,
                    Period = ParsingHelpers.Clean(match.Groups["period"].Value),
                    BilledPremium = Money(match, "premium"),
                    PaidClaims = Money(match, "paid"),
                    IncurredClaims = Money(match, "incurred"),
                    IncurredLossRatio = Money(match, "ratio"),
                    PaidLossRatio = CalculateLossRatio(Money(match, "premium"), Money(match, "paid")),
                    ProposedAdjustment = factors.FeeGuideAdjustment,
                    Weighting = Money(match, "weight"),
                    TrendFactor = factors.AnnualTrendFactor,
                    Credibility = factors.Credibility,
                    TargetLossRatio = factors.TargetLossRatio,
                    SourceReference = $"Canada Life PDF page {page.PageNumber}"
                }, match.Groups["weightUnit"].Value == "%"));
            }

            if (page.Text.Contains("Weights Used", StringComparison.OrdinalIgnoreCase))
            {
                NormalizeRelativeExperienceWeights(pageRows);
            }

            result.ExperienceLines.AddRange(pageRows.Select(row => row.Line));
        }
    }

    private static void NormalizeRelativeExperienceWeights(List<(ExperienceLine Line, bool WeightIsPercent)> rows)
    {
        var weightedRows = rows
            .Where(row => row.Line.Weighting.HasValue)
            .ToList();

        if (weightedRows.Count == 0 || weightedRows.Any(row => row.WeightIsPercent))
        {
            return;
        }

        var totalWeight = weightedRows.Sum(row => row.Line.Weighting!.Value);
        if (totalWeight <= 0m)
        {
            return;
        }

        foreach (var row in weightedRows)
        {
            row.Line.Weighting = Math.Round(
                row.Line.Weighting!.Value / totalWeight * 100m,
                4,
                MidpointRounding.AwayFromZero);
        }
    }

    private static string ParseExperienceBenefit(string text)
    {
        if (text.Contains("Drug Experience Analysis", StringComparison.OrdinalIgnoreCase))
        {
            return "Drugs";
        }

        if (text.Contains("Health Experience Analysis", StringComparison.OrdinalIgnoreCase))
        {
            return "Extended Health Care";
        }

        if (text.Contains("Dental Experience Analysis", StringComparison.OrdinalIgnoreCase))
        {
            return "Dental";
        }

        if (text.Contains("Vision Experience Analysis", StringComparison.OrdinalIgnoreCase))
        {
            return "Vision";
        }

        if (text.Contains("STD Experience Analysis", StringComparison.OrdinalIgnoreCase))
        {
            return "STD";
        }

        return "";
    }

    private static (decimal? TargetLossRatio, decimal? AnnualTrendFactor, decimal? Credibility, decimal? FeeGuideAdjustment) ParseExperienceFactors(string text)
    {
        var target = Regex.Match(text, @"Target Loss Ratio\s+(?<value>[\d.]+)%", RegexOptions.IgnoreCase);
        var annualTrend = Regex.Match(text, @"Annual Trend Factor\s+(?<value>[\d.]+)%", RegexOptions.IgnoreCase);
        var credibility = Regex.Match(text, @"Credibility\s+(?<value>[\d.]+)%", RegexOptions.IgnoreCase);
        var feeGuide = Regex.Match(text, @"Fee Guide Adjustment\s+(?<value>-?[\d.]+)%", RegexOptions.IgnoreCase);
        return (Money(target, "value"), Money(annualTrend, "value"), Money(credibility, "value"), Money(feeGuide, "value"));
    }

    private static void ParseDemographics(List<PdfPageText> pages, PdfExtractionResult result)
    {
        foreach (var page in pages.Where(page => page.Text.Contains("Demographic Summary - ", StringComparison.OrdinalIgnoreCase)))
        {
            if (page.Text.Contains("Demographic Summary - Life Insurance", StringComparison.OrdinalIgnoreCase))
            {
                ParseDemographicPage(page, "Life", result.LifeDemographics);
            }

            if (page.Text.Contains("Demographic Summary - Long Term Disability", StringComparison.OrdinalIgnoreCase) ||
                page.Text.Contains("Demographic Summary - Long-Term Disability", StringComparison.OrdinalIgnoreCase))
            {
                ParseDemographicPage(page, "LTD", result.LtdDemographics);
            }
        }
    }

    private static void ParseDemographicPage(
        PdfPageText page,
        string benefit,
        List<DemographicLine> target)
    {
        var period = "";
        var rows = new Dictionary<string, DemographicLine>(StringComparer.OrdinalIgnoreCase);
        foreach (var line in page.Lines.Select(ParsingHelpers.Clean))
        {
            if (line.Equals("Prior Year", StringComparison.OrdinalIgnoreCase))
            {
                period = "Prior";
                continue;
            }

            if (line.Equals("Current Year", StringComparison.OrdinalIgnoreCase) ||
                line.StartsWith("Current Year ", StringComparison.OrdinalIgnoreCase))
            {
                period = "Current";
                continue;
            }

            var match = DemographicRowRegex.Match(line);
            if (!match.Success || string.IsNullOrWhiteSpace(period))
            {
                continue;
            }

            var ageBand = NormalizeAgeBand(match.Groups["age"].Value);
            if (!rows.TryGetValue(ageBand, out var row))
            {
                row = new DemographicLine
                {
                    Benefit = benefit,
                    AgeBand = ageBand,
                    SourceReference = $"Canada Life PDF page {page.PageNumber}"
                };
                rows[ageBand] = row;
            }

            if (period.Equals("Prior", StringComparison.OrdinalIgnoreCase))
            {
                row.PriorMaleLives = Int(match, "maleLives");
                row.PriorMaleVolume = Money(match, "maleVolume");
                row.PriorFemaleLives = Int(match, "femaleLives");
                row.PriorFemaleVolume = Money(match, "femaleVolume");
            }
            else
            {
                row.CurrentMaleLives = Int(match, "maleLives");
                row.CurrentMaleVolume = Money(match, "maleVolume");
                row.CurrentFemaleLives = Int(match, "femaleLives");
                row.CurrentFemaleVolume = Money(match, "femaleVolume");
            }
        }

        target.AddRange(rows.Values.OrderBy(row => AgeBandSort(row.AgeBand)));
    }

    private static WorkbookExtractionResult ParseWorkbookDetailFromPdf(string pdfPath, List<PdfPageText> pages)
    {
        var workbook = new WorkbookExtractionResult
        {
            SourcePath = pdfPath
        };

        ParseClaimsDistribution(pages, workbook);
        ParseParamedicalDetail(pages, workbook);
        ParseGroupNetTopDrugPages(pages, workbook);
        return workbook;
    }

    private static void ParseGroupNetTopDrugPages(List<PdfPageText> pages, WorkbookExtractionResult workbook)
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var page in pages.Where(page => IsCanadaLifeTopDrugReport(page.Text)))
        {
            var period = "";
            foreach (var rawLine in page.Lines.Select(ParsingHelpers.Clean))
            {
                var periodMatch = ClaimsDistributionPeriodRegex.Match(rawLine);
                if (periodMatch.Success)
                {
                    period = ParsingHelpers.Clean(periodMatch.Groups["value"].Value);
                    continue;
                }

                var match = GroupNetTopDrugRowRegex.Match(rawLine);
                if (!match.Success)
                {
                    continue;
                }

                var rank = Int(match, "rank") ?? 0;
                if (rank is < 1 or > 10)
                {
                    continue;
                }

                var (drugName, treatment) = SplitGroupNetDrugDescriptor(match.Groups["descriptor"].Value);
                if (string.IsNullOrWhiteSpace(drugName))
                {
                    continue;
                }

                var din = match.Groups["din"].Value;
                var key = $"{page.PageNumber}|{period}|{rank}|{din}";
                if (!seen.Add(key))
                {
                    continue;
                }

                workbook.DrugTrends.Add(new DrugTrend
                {
                    Din = din,
                    DrugName = drugName,
                    Treats = treatment,
                    Period = period,
                    PrescriptionCount = Int(match, "claims"),
                    NetAmountPaid = Money(match, "amount"),
                    SourceReference = $"Canada Life GroupNet Top Drugs PDF page {page.PageNumber}"
                });
            }
        }
    }

    private static (string DrugName, string Treatment) SplitGroupNetDrugDescriptor(string descriptor)
    {
        var value = ParsingHelpers.Clean(descriptor);
        var treatmentMatch = GroupNetTreatmentRegex.Match(value);
        if (!treatmentMatch.Success)
        {
            return (value, "");
        }

        var drugName = ParsingHelpers.Clean(value[..treatmentMatch.Index]);
        var treatment = ParsingHelpers.Clean(treatmentMatch.Groups["value"].Value);
        return string.IsNullOrWhiteSpace(drugName) ? (value, "") : (drugName, treatment);
    }

    private static void ParseClaimsDistribution(List<PdfPageText> pages, WorkbookExtractionResult workbook)
    {
        foreach (var page in pages.Where(page => page.Text.Contains("Claims Distribution by Service - Summary", StringComparison.OrdinalIgnoreCase)))
        {
            var benefit = "";
            var period = "";
            foreach (var line in page.Lines.Select(ParsingHelpers.Clean))
            {
                var benefitMatch = ClaimsDistributionBenefitRegex.Match(line);
                if (benefitMatch.Success)
                {
                    benefit = NormalizeClaimBenefit(benefitMatch.Groups["value"].Value);
                    continue;
                }

                var periodMatch = ClaimsDistributionPeriodRegex.Match(line);
                if (periodMatch.Success)
                {
                    period = periodMatch.Groups["value"].Value;
                    continue;
                }

                var rowMatch = ClaimsDistributionRowRegex.Match(line);
                if (!rowMatch.Success || line.StartsWith("Total ", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var category = NormalizeClaimCategory(rowMatch.Groups["category"].Value);
                var amount = Money(rowMatch, "amount");
                var percent = Money(rowMatch, "percent");
                var benchmark = Money(rowMatch, "normPercent");
                workbook.ClaimTrends.Add(new ClaimTrend
                {
                    Benefit = benefit,
                    Category = category,
                    Period = period,
                    Amount = amount,
                    PercentOfTotal = percent,
                    BenchmarkPercent = benchmark,
                    SourceReference = $"Canada Life PDF page {page.PageNumber}"
                });

                AddServiceCategoryTrend(workbook, benefit, category, period, amount, page.PageNumber);
            }
        }
    }

    private static void ParseParamedicalDetail(List<PdfPageText> pages, WorkbookExtractionResult workbook)
    {
        foreach (var page in pages.Where(page => page.Text.Contains("Claims Distribution - Paramedical And Related Services - Detail", StringComparison.OrdinalIgnoreCase)))
        {
            var period = "";
            var addedKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var line in page.Lines.Select(ParsingHelpers.Clean))
            {
                var periodMatch = ClaimsDistributionPeriodRegex.Match(line);
                if (periodMatch.Success)
                {
                    period = periodMatch.Groups["value"].Value;
                    continue;
                }

                AddParamedicalDetailRows(workbook, line, period, page.PageNumber, addedKeys, allowGeneric: true);
            }

            AddParamedicalDetailRows(workbook, page.Text, period, page.PageNumber, addedKeys, allowGeneric: false);
        }
    }

    private static void AddParamedicalDetailRows(
        WorkbookExtractionResult workbook,
        string text,
        string period,
        int pageNumber,
        HashSet<string> addedKeys,
        bool allowGeneric)
    {
        foreach (Match match in KnownParamedicalRowRegex.Matches(ParsingHelpers.Clean(text)))
        {
            AddParamedicalDetailRow(workbook, match.Groups["category"].Value, Money(match, "amount"), period, pageNumber, addedKeys);
        }

        var genericMatch = allowGeneric ? ParamedicalRowRegex.Match(ParsingHelpers.Clean(text)) : Match.Empty;
        if (genericMatch.Success && !text.StartsWith("Total ", StringComparison.OrdinalIgnoreCase))
        {
            AddParamedicalDetailRow(workbook, genericMatch.Groups["category"].Value, Money(genericMatch, "amount"), period, pageNumber, addedKeys);
        }
    }

    private static void AddParamedicalDetailRow(
        WorkbookExtractionResult workbook,
        string category,
        decimal? amount,
        string period,
        int pageNumber,
        HashSet<string> addedKeys)
    {
        var normalizedCategory = NormalizeParamedicalCategory(category);
        if (ParamedicalCategoryNormalizer.IsRollup(normalizedCategory))
        {
            return;
        }

        var key = $"{normalizedCategory}|{amount}|{period}";
        if (!addedKeys.Add(key))
        {
            return;
        }

        workbook.ParamedicalServiceTrends.Add(new ServiceCategoryTrend
        {
            Benefit = "Health",
            Category = normalizedCategory,
            CurrentPeriodLabel = period,
            CurrentAmount = amount ?? 0m,
            SourceReference = $"Canada Life PDF page {pageNumber}"
        });
    }

    private static void AddServiceCategoryTrend(WorkbookExtractionResult workbook, string benefit, string category, string period, decimal? amount, int pageNumber)
    {
        var row = new ServiceCategoryTrend
        {
            Benefit = benefit,
            Category = category,
            CurrentPeriodLabel = period,
            CurrentAmount = amount ?? 0m,
            SourceReference = $"Canada Life PDF page {pageNumber}"
        };

        if (benefit.Equals("Dental", StringComparison.OrdinalIgnoreCase))
        {
            workbook.DentalSpendingTrends.Add(row);
        }
        else if (benefit.Equals("Health", StringComparison.OrdinalIgnoreCase))
        {
            workbook.EhcSpendingTrends.Add(row);
        }
    }

    private static string NormalizeBenefit(string value)
    {
        value = ParsingHelpers.Clean(value);
        return value switch
        {
            "Employee Term Life" => "Basic Life",
            "AD&D" or "Accidental Death & Dismemberment" => "Accidental Death and Dismemberment",
            "Employee Critical Illness" => "Critical Illness",
            "Long Term Disability" or "Long-Term Disability" => "LTD",
            "Healthcare" or "Health" or "Medical" => "Extended Health Care",
            "Contact" or "Employee and Family Assistance Program" => "EAP",
            "Visioncare" => "Vision",
            "Virtual Healthcare" => "Virtual Health",
            _ => value
        };
    }

    private static string NormalizeClaimBenefit(string value)
    {
        value = ParsingHelpers.Clean(value);
        return value switch
        {
            "Medical" or "Healthcare" or "Health" => "Health",
            "Visioncare" => "Vision",
            _ => value
        };
    }

    private static string NormalizeClaimCategory(string value)
    {
        value = ParsingHelpers.Clean(value);
        var normalized = Regex.Replace(value, @"[^A-Za-z0-9]", "").ToLowerInvariant();
        if (normalized == "total")
        {
            return "Total";
        }

        if (normalized is "drug" or "drugs" or "prescriptiondrugs" or "prescriptionmedication")
        {
            return "Drugs";
        }

        if (normalized is "paramedical" or "paramedicalservices" or "practitionerservices" or "healthpractitioners" or "professionalservices" or "alliedhealth" or "massagetherapy" or "physiotherapy" or "chiropractic" or "psychology" or "psychologistsofficevisit" or "naturopathic" or "acupuncture" ||
            ParamedicalCategoryNormalizer.IsKnownParamedicalCategory(value))
        {
            return "Paramedical";
        }

        if (normalized is "serviceandsupplies" or "medicalsupplies" or "supplies" or "medicalservicesandsupplies")
        {
            return "Medical Supplies";
        }

        if (normalized is "outofcountry" or "travel" or "emergencytravel")
        {
            return "Out of Country";
        }

        if (normalized is "vision" or "visioncare" or "eyeglasses" or "eyeexams")
        {
            return "Vision";
        }

        if (normalized is "dental" or "basicdental" or "majordental" or "orthodontics")
        {
            return "Dental";
        }

        return normalized switch
        {
            "hospital" => "Hospital",
            "ambulance" => "Ambulance",
            "outofprovince" => "Out of Province",
            "other" or "miscellaneous" => "Other",
            "masseur" => "Massage",
            _ => value
        };
    }

    private static string NormalizeParamedicalCategory(string value)
    {
        return ParamedicalCategoryNormalizer.Normalize(value);
    }

    private static string NormalizeAgeBand(string value)
    {
        return DemographicAgeBandNormalizer.NormalizeLtdAgeBand(ParsingHelpers.Clean(value));
    }

    private static int AgeBandSort(string value)
    {
        return value switch
        {
            "Under 25" => 0,
            "<25" => 0,
            "25-29" => 1,
            "30-34" => 2,
            "35-39" => 3,
            "40-44" => 4,
            "45-49" => 5,
            "50-54" => 6,
            "55-59" => 7,
            "60-64" => 8,
            "60 & over" => 9,
            "65-69" => 10,
            "65+" => 11,
            "70 & over" => 12,
            _ => 99
        };
    }

    private static decimal? CalculateLossRatio(decimal? premium, decimal? claims)
    {
        return premium is > 0m && claims.HasValue
            ? decimal.Round(claims.Value / premium.Value * 100m, 1)
            : null;
    }

    private static decimal? SumPremium(IEnumerable<RateLine> lines, string benefit, bool current)
    {
        var values = lines
            .Where(line => line.Benefit.Equals(benefit, StringComparison.OrdinalIgnoreCase))
            .Select(line => current ? line.CurrentPremium : line.RenewalPremium)
            .Where(value => value.HasValue)
            .Select(value => value!.Value)
            .ToList();

        return values.Count == 0 ? null : values.Sum();
    }

    private static void FillMonthlyPremiumTotalsFromRateLines(PdfExtractionResult result)
    {
        if (result.RateLines.Count == 0)
        {
            return;
        }

        result.CurrentMonthlyPremium ??= SumAllPremium(result.RateLines, current: true);
        result.RenewalMonthlyPremium ??= SumAllPremium(result.RateLines, current: false);

        if (!result.MonthlyPremiumChange.HasValue &&
            result.CurrentMonthlyPremium.HasValue &&
            result.RenewalMonthlyPremium.HasValue)
        {
            result.MonthlyPremiumChange = decimal.Round(result.RenewalMonthlyPremium.Value - result.CurrentMonthlyPremium.Value, 2);
        }

        result.MonthlyPremiumChangePercent ??= UnderwritingEngine.CalculatePercentChange(
            result.CurrentMonthlyPremium,
            result.RenewalMonthlyPremium);
    }

    private static decimal? SumAllPremium(IEnumerable<RateLine> lines, bool current)
    {
        var values = lines
            .Select(line => current ? line.CurrentPremium : line.RenewalPremium)
            .Where(value => value.HasValue)
            .Select(value => value!.Value)
            .ToList();

        return values.Count == 0 ? null : decimal.Round(values.Sum(), 2);
    }

    private static string MatchValue(Regex regex, string text)
    {
        var match = regex.Match(text);
        return match.Success ? ParsingHelpers.Clean(match.Groups["value"].Value) : "";
    }

    private static decimal? Money(Match match, string groupName)
    {
        return ParsingHelpers.ParseDecimal(NormalizeSplitNumberText(match.Groups[groupName].Value));
    }

    private static string NormalizeSplitNumberText(string value)
    {
        value = SpacedCommaMoneyTokenRegex.Replace(ParsingHelpers.Clean(value), ",");
        var previous = "";
        while (!previous.Equals(value, StringComparison.Ordinal))
        {
            previous = value;
            value = SplitMoneyTokenRegex.Replace(value, match =>
                match.Groups["prefix"].Value + match.Groups["rest"].Value);
        }

        return value;
    }

    private static int? Int(Match match, string groupName)
    {
        var text = match.Groups[groupName].Value;
        return text.Equals("-", StringComparison.Ordinal) ? null : ParsingHelpers.ParseInt(text);
    }
}

public sealed class VictorPdfParser
{
    private static readonly Regex ClientNumberRegex = new(@"Client (?:No\.|Number):\s*(?<value>[\d-]+)", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex EffectiveDateRegex = new(@"effective\s+(?<date>[A-Za-z]+\s+\d{1,2},\s+\d{4})", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex GrandTotalRegex = new(@"Grand Total\s+\$\s*(?<current>[\d,]+\.\d{2})\s+\$\s*(?<renewal>[\d,]+\.\d{2})", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex VolumeRateRegex = new(@"^(?<benefit>Basic Life|Basic AD&D|Employee Critical Illness)\s+(?<class>[A-Z])\s+\$\s*(?<volume>[\d,]+)\s+\$\s*(?<currentRate>[\d.]+)\s*/\$\s*(?<currentBasis>[\d,]+)\s+\$\s*(?<currentPremium>[\d,]+\.\d{2})\s+\$\s*(?<renewalRate>[\d.]+)\s*/\$\s*(?<renewalBasis>[\d,]+)\s+\$\s*(?<renewalPremium>[\d,]+\.\d{2})$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex UnitRateRegex = new(@"^Dependent Life\s+(?<class>[A-Z])\s+(?<lives>\d+)\s+\$\s*(?<currentRate>[\d.]+)\s*/\s*Unit\s+\$\s*(?<currentPremium>[\d,]+\.\d{2})\s+\$\s*(?<renewalRate>[\d.]+)\s*/\s*Unit\s+\$\s*(?<renewalPremium>[\d,]+\.\d{2})$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex LivesRateRegex = new(@"^(?:(?<benefit>Extended Health Care|Emergency Travel Assistance|Dental Care)\s+)?(?<tier>Single|Family)\s+(?<class>[A-Z])\s+(?<lives>\d+)\s+\$\s*(?<currentRate>[\d.]+)\s+\$\s*(?<currentPremium>[\d,]+\.\d{2})\s+\$\s*(?<renewalRate>[\d.]+)\s+\$\s*(?<renewalPremium>[\d,]+\.\d{2})$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex OptionalRateRegex = new(@"^Optional AD&D\s+\$\s*(?<volume>[\d,]+)\s+Single/Family\s+\$\s*(?<currentPremium>[\d.]+)\s+Single/Family\s+\$\s*(?<renewalPremium>[\d.]+)$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex PaidLossRatioRegex = new(@"For the Period\s+(?<period>.*?)(?:\s+Premium\s+Claims Paid.*?| )\s+\$\s*(?<premium>[\d,]+\.\d{2})\s+\$\s*(?<claims>[\d,]+\.\d{2})\s+(?<claimCount>\d+)\s+\$\s*(?<average>[\d,]+\.\d{2})\s+(?<lossRatio>[\d.]+)%", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex RenewalCalculationTitleRegex = new(@"Renewal Rate Calculations - (?<benefit>.+?) Client No", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex TripleMoneyRegex = new(@"\s+\$?\s*(?<current>[\d,]+\.\d{2})\s+\$?\s*(?<prior>[\d,]+\.\d{2})\s+\$?\s*(?<weighted>[\d,]+\.\d{2})", RegexOptions.Compiled);
    private static readonly Regex TriplePercentRegex = new(@"\s+(?<current>[\d.]+)%\s+(?<prior>[\d.]+)%\s+(?<weighted>[\d.]+)%", RegexOptions.Compiled);
    private static readonly Regex RateActionRegex = new(@"Rate Action is\s+(?<value>[+-]?[\d.]+)%|Rate Action is\s+0%", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex TargetLossRatioRegex = new(@"Target Loss Ratio.*?is\s+(?<value>[\d.]+)%", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex TotalCategoryRegex = new(@"^(?<category>.+?)\s+Total:\s+\$(?<amount>[\d,]+\.\d{2})\s+(?<percent>[\d.]+)%$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex CategoryDetailRegex = new(@"^(?<category>[A-Za-z][A-Za-z .&/-]+?)\s+\$(?<amount>[\d,]+\.\d{2})\s+(?<percent>[\d.]+)%$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex TopDrugPeriodRegex = new(@"For the Period\s+(?<period>[A-Za-z]+\s+\d{1,2},\s+\d{4}\s+to\s+[A-Za-z]+\s+\d{1,2},\s+\d{4})", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex DrugRegex = new(@"^(?<rank>\d+)\s+(?<din>\d+)\s+(?<name>.+?)\s+(?<claims>\d+)\s+\$(?<submitted>[\d,]+\.\d{2})\s+\$(?<paid>[\d,]+\.\d{2})$", RegexOptions.Compiled);
    private static readonly Regex AgeDistributionRowRegex = new(
        @"^(?<priorAge><20|20-24|25-29|30-34|35-39|40-44|45-49|50-54|55-59|60-64|65-69|>=70|>70)\s+" +
        @"(?<priorMaleLives>\d+)\s+\$?(?<priorMaleVolume>[\d,]+(?:\.\d{2})?)\s+" +
        @"(?<priorFemaleLives>\d+)\s+\$?(?<priorFemaleVolume>[\d,]+(?:\.\d{2})?)\s+" +
        @"(?<priorTotalLives>\d+)\s+\$?(?<priorTotalVolume>[\d,]+(?:\.\d{2})?)\s+" +
        @"(?<currentAge><20|20-24|25-29|30-34|35-39|40-44|45-49|50-54|55-59|60-64|65-69|>=70|>70)\s+" +
        @"(?<currentMaleLives>\d+)\s+\$?(?<currentMaleVolume>[\d,]+(?:\.\d{2})?)\s+" +
        @"(?<currentFemaleLives>\d+)\s+\$?(?<currentFemaleVolume>[\d,]+(?:\.\d{2})?)\s+" +
        @"(?<currentTotalLives>\d+)\s+\$?(?<currentTotalVolume>[\d,]+(?:\.\d{2})?)$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public static bool IsVictorRenewal(string text)
    {
        return text.Contains("Victor Insurance Managers", StringComparison.OrdinalIgnoreCase) ||
               text.Contains("victorinsurance.ca", StringComparison.OrdinalIgnoreCase);
    }

    public CarrierPdfParseResult Parse(string pdfPath, List<PdfPageText> pages, string allText)
    {
        var pdf = new PdfExtractionResult
        {
            SourcePath = pdfPath,
            Carrier = "Victor",
            PolicyNumber = MatchValue(ClientNumberRegex, allText),
            ClientName = ParseClientName(pages, allText),
            RenewalDate = ParseRenewalDate(allText)
        };

        ParseTotals(allText, pdf);
        ParseRateLines(pages, pdf);
        ParseExperienceLines(pages, pdf);
        ParseRenewalCalculationLines(pages, pdf);
        ParseDemographics(pages, pdf);

        var workbook = ParseWorkbookDetailFromPdf(pdfPath, pages);

        if (pdf.RateLines.Count == 0)
        {
            pdf.Warnings.Add("No Victor monthly premium comparison rows were extracted from the PDF.");
        }

        if (pdf.ExperienceLines.Count == 0)
        {
            pdf.Warnings.Add("No Victor paid loss ratio rows were extracted from the PDF.");
        }

        return new CarrierPdfParseResult
        {
            Pdf = pdf,
            WorkbookFromPdf = workbook,
            Carrier = "Victor"
        };
    }

    private static string ParseClientName(List<PdfPageText> pages, string allText)
    {
        var schedulePage = pages.FirstOrDefault(page => page.Text.Contains("Schedule of Benefits", StringComparison.OrdinalIgnoreCase));
        if (schedulePage is not null)
        {
            var lines = schedulePage.Lines.Select(ParsingHelpers.Clean).Where(line => line.Length > 0).ToList();
            for (var i = 0; i < lines.Count; i++)
            {
                if (lines[i].Equals("Schedule of Benefits", StringComparison.OrdinalIgnoreCase) &&
                    i + 1 < lines.Count &&
                    !lines[i + 1].Equals("Schedule of Benefits", StringComparison.OrdinalIgnoreCase))
                {
                    return lines[i + 1];
                }

                if (lines[i].Equals("Schedule of Benefits", StringComparison.OrdinalIgnoreCase) &&
                    i + 2 < lines.Count)
                {
                    return lines[i + 2];
                }
            }
        }

        var match = Regex.Match(allText, @"(?<name>[A-Za-z0-9 &'.,-]+?)\s+\(\d{3}-\d{5}\)", RegexOptions.IgnoreCase);
        return match.Success ? ParsingHelpers.Clean(match.Groups["name"].Value) : "";
    }

    private static DateOnly? ParseRenewalDate(string text)
    {
        var match = EffectiveDateRegex.Match(text);
        if (match.Success &&
            DateOnly.TryParse(match.Groups["date"].Value, CultureInfo.InvariantCulture, DateTimeStyles.None, out var effectiveDate))
        {
            return effectiveDate;
        }

        return null;
    }

    private static void ParseTotals(string text, PdfExtractionResult result)
    {
        var match = GrandTotalRegex.Match(text);
        if (!match.Success)
        {
            return;
        }

        result.CurrentMonthlyPremium = Money(match, "current");
        result.RenewalMonthlyPremium = Money(match, "renewal");
        if (result.CurrentMonthlyPremium.HasValue && result.RenewalMonthlyPremium.HasValue)
        {
            result.MonthlyPremiumChange = decimal.Round(result.RenewalMonthlyPremium.Value - result.CurrentMonthlyPremium.Value, 2);
            result.MonthlyPremiumChangePercent = UnderwritingEngine.CalculatePercentChange(result.CurrentMonthlyPremium, result.RenewalMonthlyPremium);
        }
    }

    private static void ParseRateLines(List<PdfPageText> pages, PdfExtractionResult result)
    {
        var page = pages.FirstOrDefault(item => item.Text.Contains("Monthly Premium Comparison", StringComparison.OrdinalIgnoreCase));
        if (page is null)
        {
            return;
        }

        var currentBenefit = "";
        foreach (var rawLine in page.Lines)
        {
            var line = ParsingHelpers.Clean(rawLine);
            if (line.Length == 0 ||
                line.StartsWith("Benefit Coverage/Class", StringComparison.OrdinalIgnoreCase) ||
                line.StartsWith("Subtotal", StringComparison.OrdinalIgnoreCase) ||
                line.StartsWith("Grand Total", StringComparison.OrdinalIgnoreCase) ||
                line.StartsWith("*", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var rateLine = TryParseRateLine(line, page.PageNumber, ref currentBenefit);
            if (rateLine is not null)
            {
                result.RateLines.Add(rateLine);
            }
        }

        result.CurrentEhcPremium = SumPremium(result.RateLines, "Extended Health Care", current: true);
        result.RenewalEhcPremium = SumPremium(result.RateLines, "Extended Health Care", current: false);
        result.CurrentDentalPremium = SumPremium(result.RateLines, "Dental", current: true);
        result.RenewalDentalPremium = SumPremium(result.RateLines, "Dental", current: false);
    }

    private static RateLine? TryParseRateLine(string line, int pageNumber, ref string currentBenefit)
    {
        var match = VolumeRateRegex.Match(line);
        if (match.Success)
        {
            var benefit = NormalizeBenefit(match.Groups["benefit"].Value);
            currentBenefit = benefit;
            return new RateLine
            {
                Benefit = benefit,
                AccountClass = match.Groups["class"].Value,
                Volume = Money(match, "volume"),
                CurrentRate = Money(match, "currentRate"),
                CurrentPremium = Money(match, "currentPremium"),
                RenewalRate = Money(match, "renewalRate"),
                RenewalPremium = Money(match, "renewalPremium"),
                SourceReference = $"Victor PDF page {pageNumber}"
            };
        }

        match = UnitRateRegex.Match(line);
        if (match.Success)
        {
            currentBenefit = "Dependent Life";
            return new RateLine
            {
                Benefit = "Dependent Life",
                AccountClass = match.Groups["class"].Value,
                Lives = Int(match, "lives"),
                CurrentRate = Money(match, "currentRate"),
                CurrentPremium = Money(match, "currentPremium"),
                RenewalRate = Money(match, "renewalRate"),
                RenewalPremium = Money(match, "renewalPremium"),
                SourceReference = $"Victor PDF page {pageNumber}"
            };
        }

        match = OptionalRateRegex.Match(line);
        if (match.Success)
        {
            currentBenefit = "Optional AD&D";
            return new RateLine
            {
                Benefit = "Optional AD&D",
                TierOrClass = "Single/Family",
                Volume = Money(match, "volume"),
                CurrentRate = Money(match, "currentPremium"),
                CurrentPremium = Money(match, "currentPremium"),
                RenewalRate = Money(match, "renewalPremium"),
                RenewalPremium = Money(match, "renewalPremium"),
                SourceReference = $"Victor PDF page {pageNumber}"
            };
        }

        match = LivesRateRegex.Match(line);
        if (!match.Success)
        {
            return null;
        }

        var benefitFromLine = ParsingHelpers.Clean(match.Groups["benefit"].Value);
        var benefitName = string.IsNullOrWhiteSpace(benefitFromLine) ? currentBenefit : NormalizeBenefit(benefitFromLine);
        if (string.IsNullOrWhiteSpace(benefitName))
        {
            return null;
        }

        currentBenefit = benefitName;
        return new RateLine
        {
            Benefit = benefitName,
            TierOrClass = match.Groups["tier"].Value,
            AccountClass = match.Groups["class"].Value,
            Lives = Int(match, "lives"),
            CurrentRate = Money(match, "currentRate"),
            CurrentPremium = Money(match, "currentPremium"),
            RenewalRate = Money(match, "renewalRate"),
            RenewalPremium = Money(match, "renewalPremium"),
            SourceReference = $"Victor PDF page {pageNumber}"
        };
    }

    private static void ParseExperienceLines(List<PdfPageText> pages, PdfExtractionResult result)
    {
        foreach (var page in pages.Where(page => page.Text.Contains("Paid Loss Ratio -", StringComparison.OrdinalIgnoreCase)))
        {
            var benefit = NormalizeBenefit(Regex.Match(page.Text, @"Paid Loss Ratio - (?<benefit>.+?) Client No", RegexOptions.IgnoreCase).Groups["benefit"].Value);
            if (string.IsNullOrWhiteSpace(benefit))
            {
                continue;
            }

            foreach (Match match in PaidLossRatioRegex.Matches(page.Text))
            {
                result.ExperienceLines.Add(new ExperienceLine
                {
                    Benefit = benefit,
                    Period = ParsingHelpers.Clean(match.Groups["period"].Value),
                    BilledPremium = Money(match, "premium"),
                    PaidClaims = Money(match, "claims"),
                    IncurredClaims = Money(match, "claims"),
                    PaidLossRatio = Money(match, "lossRatio"),
                    IncurredLossRatio = Money(match, "lossRatio"),
                    SourceReference = $"Victor PDF page {page.PageNumber}"
                });
            }
        }
    }

    private static void ParseRenewalCalculationLines(List<PdfPageText> pages, PdfExtractionResult result)
    {
        foreach (var page in pages.Where(page => page.Text.Contains("Renewal Rate Calculations -", StringComparison.OrdinalIgnoreCase)))
        {
            var benefit = NormalizeBenefit(RenewalCalculationTitleRegex.Match(page.Text).Groups["benefit"].Value);
            if (string.IsNullOrWhiteSpace(benefit))
            {
                continue;
            }

            var adjustedPremium = ParseTripleMoneyLine(page.Lines, "Adjusted Premium");
            var adjustedClaims = ParseTripleMoneyLine(page.Lines, "Adjusted Claims");
            var adjustedLossRatio = ParseTriplePercentLine(page.Lines, "Adjusted Loss Ratio");
            var targetLossRatio = Money(TargetLossRatioRegex.Match(page.Text), "value");
            var rateAction = ParseRateAction(page.Text);

            if (adjustedPremium.Weighted.HasValue || adjustedClaims.Weighted.HasValue || adjustedLossRatio.Weighted.HasValue)
            {
                result.ExperienceLines.Add(new ExperienceLine
                {
                    Benefit = benefit,
                    Period = "Weighted renewal calculation",
                    BilledPremium = adjustedPremium.Weighted,
                    PaidClaims = adjustedClaims.Weighted,
                    IncurredClaims = adjustedClaims.Weighted,
                    PaidLossRatio = adjustedLossRatio.Weighted,
                    IncurredLossRatio = adjustedLossRatio.Weighted,
                    ProposedAdjustment = rateAction,
                    Weighting = targetLossRatio,
                    SourceReference = $"Victor PDF page {page.PageNumber}"
                });
            }
        }
    }

    private static void ParseDemographics(List<PdfPageText> pages, PdfExtractionResult result)
    {
        var rows = new Dictionary<string, DemographicLine>(StringComparer.OrdinalIgnoreCase);

        foreach (var page in pages.Where(page => page.Text.Contains("Age Distribution Detail Report", StringComparison.OrdinalIgnoreCase)))
        {
            foreach (var line in page.Lines.Select(ParsingHelpers.Clean))
            {
                var match = AgeDistributionRowRegex.Match(line);
                if (!match.Success)
                {
                    continue;
                }

                AddAgeDistributionRow(rows, match, page.PageNumber);
            }
        }

        result.LifeDemographics.AddRange(rows.Values.OrderBy(row => VictorLifeAgeBandSort(row.AgeBand)));
    }

    private static void AddAgeDistributionRow(Dictionary<string, DemographicLine> rows, Match match, int pageNumber)
    {
        var ageBand = NormalizeVictorLifeAgeBand(match.Groups["currentAge"].Value);
        if (string.IsNullOrWhiteSpace(ageBand))
        {
            return;
        }

        if (!rows.TryGetValue(ageBand, out var row))
        {
            row = new DemographicLine
            {
                Benefit = "Life",
                AgeBand = ageBand,
                SourceReference = $"Victor PDF page {pageNumber}"
            };
            rows[ageBand] = row;
        }

        row.PriorMaleLives = Add(row.PriorMaleLives, Int(match, "priorMaleLives"));
        row.PriorMaleVolume = Add(row.PriorMaleVolume, Money(match, "priorMaleVolume"));
        row.PriorFemaleLives = Add(row.PriorFemaleLives, Int(match, "priorFemaleLives"));
        row.PriorFemaleVolume = Add(row.PriorFemaleVolume, Money(match, "priorFemaleVolume"));
        row.CurrentMaleLives = Add(row.CurrentMaleLives, Int(match, "currentMaleLives"));
        row.CurrentMaleVolume = Add(row.CurrentMaleVolume, Money(match, "currentMaleVolume"));
        row.CurrentFemaleLives = Add(row.CurrentFemaleLives, Int(match, "currentFemaleLives"));
        row.CurrentFemaleVolume = Add(row.CurrentFemaleVolume, Money(match, "currentFemaleVolume"));
    }

    private static string NormalizeVictorLifeAgeBand(string ageBand)
    {
        var normalized = ParsingHelpers.Clean(ageBand);
        return normalized switch
        {
            "<20" or "20-24" => "Under 25",
            ">=70" or ">70" => "70 & over",
            _ => normalized
        };
    }

    private static int VictorLifeAgeBandSort(string ageBand)
    {
        return NormalizeVictorLifeAgeBand(ageBand) switch
        {
            "Under 25" => 0,
            "25-29" => 1,
            "30-34" => 2,
            "35-39" => 3,
            "40-44" => 4,
            "45-49" => 5,
            "50-54" => 6,
            "55-59" => 7,
            "60-64" => 8,
            "65-69" => 9,
            "70 & over" => 10,
            _ => 99
        };
    }

    private static WorkbookExtractionResult ParseWorkbookDetailFromPdf(string pdfPath, List<PdfPageText> pages)
    {
        var workbook = new WorkbookExtractionResult
        {
            SourcePath = pdfPath
        };

        ParseCategoryPages(pages, workbook);
        ParseTopDrugPage(pages, workbook);
        return workbook;
    }

    private static void ParseCategoryPages(List<PdfPageText> pages, WorkbookExtractionResult workbook)
    {
        foreach (var page in pages)
        {
            var isEhc = page.Text.Contains("Extended Health Care Claims by Category", StringComparison.OrdinalIgnoreCase);
            var isDental = page.Text.Contains("Dental Care Claims by Category", StringComparison.OrdinalIgnoreCase);
            if (!isEhc && !isDental)
            {
                continue;
            }

            var benefit = isEhc ? "Health" : "Dental";
            foreach (var rawLine in page.Lines)
            {
                var line = ParsingHelpers.Clean(rawLine);
                var totalMatch = TotalCategoryRegex.Match(line);
                if (totalMatch.Success && !line.StartsWith("All Benefit", StringComparison.OrdinalIgnoreCase))
                {
                    AddClaimTrend(workbook, benefit, totalMatch.Groups["category"].Value, Money(totalMatch, "amount"), Money(totalMatch, "percent"), page.PageNumber);
                    AddSpendingTrend(workbook, benefit, totalMatch.Groups["category"].Value, Money(totalMatch, "amount"), page.PageNumber);
                    continue;
                }

                if (!isEhc || line.Contains("Total:", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var detailMatch = CategoryDetailRegex.Match(line);
                if (detailMatch.Success)
                {
                    var category = detailMatch.Groups["category"].Value;
                    if (IsHeaderOrSummaryCategory(category))
                    {
                        continue;
                    }

                    if (page.Lines.TakeWhile(item => !ReferenceEquals(item, rawLine)).LastOrDefault()?.Contains("Paramedical", StringComparison.OrdinalIgnoreCase) == true ||
                        IsParamedicalCategory(category))
                    {
                        workbook.ParamedicalServiceTrends.Add(new ServiceCategoryTrend
                        {
                            Benefit = "Health",
                            Category = category,
                            CurrentPeriodLabel = "Current",
                            CurrentAmount = Money(detailMatch, "amount") ?? 0m,
                            SourceReference = $"Victor PDF page {page.PageNumber}"
                        });
                    }
                }
            }
        }
    }

    private static void ParseTopDrugPage(List<PdfPageText> pages, WorkbookExtractionResult workbook)
    {
        var sections = new List<VictorTopDrugSection>();
        foreach (var page in pages.Where(page => page.Text.Contains("Top Drug Claims by DIN", StringComparison.OrdinalIgnoreCase)))
        {
            var period = ParseTopDrugPeriod(page);
            var rows = new List<DrugTrend>();
            foreach (var line in page.Lines.Select(ParsingHelpers.Clean))
            {
                var match = DrugRegex.Match(line);
                if (!match.Success)
                {
                    continue;
                }

                rows.Add(new DrugTrend
                {
                    Din = match.Groups["din"].Value,
                    DrugName = match.Groups["name"].Value,
                    Period = period,
                    PrescriptionCount = Int(match, "claims"),
                    SubmittedTotal = Money(match, "submitted"),
                    NetAmountPaid = Money(match, "paid"),
                    SourceReference = $"Victor PDF page {page.PageNumber}"
                });
            }

            if (rows.Count > 0)
            {
                sections.Add(new VictorTopDrugSection(page.PageNumber, period, ParseTopDrugPeriodEnd(period), rows));
            }
        }

        if (sections.Count == 0)
        {
            return;
        }

        var currentSection = sections
            .OrderByDescending(section => section.PeriodEndDate ?? DateOnly.MinValue)
            .ThenByDescending(section => section.PageNumber)
            .First();

        workbook.DrugTrends.AddRange(currentSection.Rows);
        foreach (var priorSection in sections.Where(section => !ReferenceEquals(section, currentSection)))
        {
            workbook.PriorDrugTrends.AddRange(priorSection.Rows);
        }
    }

    private static string ParseTopDrugPeriod(PdfPageText page)
    {
        var match = TopDrugPeriodRegex.Match(page.Text);
        if (match.Success)
        {
            return ParsingHelpers.Clean(match.Groups["period"].Value);
        }

        foreach (var line in page.Lines.Select(ParsingHelpers.Clean))
        {
            match = TopDrugPeriodRegex.Match(line);
            if (match.Success)
            {
                return ParsingHelpers.Clean(match.Groups["period"].Value);
            }
        }

        return "";
    }

    private static DateOnly? ParseTopDrugPeriodEnd(string period)
    {
        var parts = period.Split(" to ", StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parts.Length != 2)
        {
            return null;
        }

        return DateOnly.TryParse(parts[1], CultureInfo.InvariantCulture, DateTimeStyles.None, out var endDate)
            ? endDate
            : null;
    }

    private static void AddClaimTrend(WorkbookExtractionResult workbook, string benefit, string category, decimal? amount, decimal? percent, int pageNumber)
    {
        workbook.ClaimTrends.Add(new ClaimTrend
        {
            Benefit = benefit,
            Category = ParsingHelpers.Clean(category),
            Period = "Current",
            Amount = amount,
            PercentOfTotal = percent,
            SourceReference = $"Victor PDF page {pageNumber}"
        });
    }

    private static void AddSpendingTrend(WorkbookExtractionResult workbook, string benefit, string category, decimal? amount, int pageNumber)
    {
        var trend = new ServiceCategoryTrend
        {
            Benefit = benefit,
            Category = ParsingHelpers.Clean(category),
            CurrentPeriodLabel = "Current",
            CurrentAmount = amount ?? 0m,
            SourceReference = $"Victor PDF page {pageNumber}"
        };

        if (benefit.Equals("Dental", StringComparison.OrdinalIgnoreCase))
        {
            workbook.DentalSpendingTrends.Add(trend);
        }
        else
        {
            workbook.EhcSpendingTrends.Add(trend);
        }
    }

    private static bool IsHeaderOrSummaryCategory(string category)
    {
        return category.Contains("Service Category", StringComparison.OrdinalIgnoreCase) ||
               category.Contains("Amount", StringComparison.OrdinalIgnoreCase) ||
               category.Equals("All Benefit", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsParamedicalCategory(string category)
    {
        return category.Contains("Acupuncture", StringComparison.OrdinalIgnoreCase) ||
               category.Contains("Chiropr", StringComparison.OrdinalIgnoreCase) ||
               category.Contains("Massage", StringComparison.OrdinalIgnoreCase) ||
               category.Contains("Osteopathy", StringComparison.OrdinalIgnoreCase) ||
               category.Contains("Physiotherapy", StringComparison.OrdinalIgnoreCase) ||
               category.Contains("Psychology", StringComparison.OrdinalIgnoreCase);
    }

    private static (decimal? Current, decimal? Prior, decimal? Weighted) ParseTripleMoneyLine(IEnumerable<string> lines, string label)
    {
        var line = lines.Select(ParsingHelpers.Clean).FirstOrDefault(item => item.StartsWith(label, StringComparison.OrdinalIgnoreCase));
        if (string.IsNullOrWhiteSpace(line))
        {
            return default;
        }

        var match = TripleMoneyRegex.Match(line);
        return match.Success ? (Money(match, "current"), Money(match, "prior"), Money(match, "weighted")) : default;
    }

    private static (decimal? Current, decimal? Prior, decimal? Weighted) ParseTriplePercentLine(IEnumerable<string> lines, string label)
    {
        var line = lines.Select(ParsingHelpers.Clean).FirstOrDefault(item => item.StartsWith(label, StringComparison.OrdinalIgnoreCase));
        if (string.IsNullOrWhiteSpace(line))
        {
            return default;
        }

        var match = TriplePercentRegex.Match(line);
        return match.Success ? (Money(match, "current"), Money(match, "prior"), Money(match, "weighted")) : default;
    }

    private static decimal? ParseRateAction(string text)
    {
        var match = RateActionRegex.Match(text);
        if (!match.Success)
        {
            return null;
        }

        if (string.IsNullOrWhiteSpace(match.Groups["value"].Value))
        {
            return 0m;
        }

        return Money(match, "value");
    }

    private static decimal? SumPremium(IEnumerable<RateLine> lines, string benefit, bool current)
    {
        var values = lines
            .Where(line => line.Benefit.Equals(benefit, StringComparison.OrdinalIgnoreCase) ||
                           line.Benefit.StartsWith(benefit + " ", StringComparison.OrdinalIgnoreCase))
            .Select(line => current ? line.CurrentPremium : line.RenewalPremium)
            .Where(value => value.HasValue)
            .Select(value => value!.Value)
            .ToList();

        return values.Count == 0 ? null : decimal.Round(values.Sum(), 2);
    }

    private static string NormalizeBenefit(string value)
    {
        value = ParsingHelpers.Clean(value);
        return value switch
        {
            "Basic Life" => "Life",
            "Basic AD&D" => "Accidental Death and Dismemberment",
            "Dental Care" => "Dental",
            "Employee Critical Illness" => "Critical Illness",
            _ => value
        };
    }

    private static string MatchValue(Regex regex, string text)
    {
        var match = regex.Match(text);
        return match.Success ? ParsingHelpers.Clean(match.Groups["value"].Value) : "";
    }

    private static decimal? Money(Match match, string groupName)
    {
        return match.Success ? ParsingHelpers.ParseDecimal(match.Groups[groupName].Value) : null;
    }

    private static int? Int(Match match, string groupName)
    {
        return ParsingHelpers.ParseInt(match.Groups[groupName].Value);
    }

    private static int? Add(int? left, int? right)
    {
        return (left ?? 0) + (right ?? 0);
    }

    private static decimal? Add(decimal? left, decimal? right)
    {
        return (left ?? 0m) + (right ?? 0m);
    }

    private sealed record VictorTopDrugSection(
        int PageNumber,
        string Period,
        DateOnly? PeriodEndDate,
        List<DrugTrend> Rows);
}

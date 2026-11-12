using System.Text.RegularExpressions;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed class ManulifePdfParser
{
    private static readonly Regex DateRegex = new(@"\d{1,2}-[A-Za-z]{3}-\d{4}", RegexOptions.Compiled);

    private static readonly string[] BenefitNames =
    [
        "Non Pooled Extended Health Care - Subtotal",
        "Extended Health Care - Drugs and Other",
        "Extended Health Care - Hospital",
        "Extended Health Care - Pooled",
        "Extended Health Care - Total",
        "Extended Health Care - Vision",
        "Extended Health Care",
        "Accidental Death and Dismemberment",
        "Employee and Family Assistance",
        "Group Critical Illness",
        "Critical Illness",
        "Long Term Disability",
        "Short Term Disability",
        "Dependent Life",
        "Healthcare Online",
        "Dental",
        "STD",
        "Life",
        "Pooled",
        "Hospital",
        "Vision",
        "Drugs and Other"
    ];

    private static readonly string[] RateSectionBenefits =
    [
        "Accidental Death and Dismemberment",
        "Employee and Family Assistance",
        "Group Critical Illness",
        "Long Term Disability",
        "Short Term Disability",
        "Extended Health Care - Drugs and Other",
        "Extended Health Care - Hospital",
        "Extended Health Care - Pooled",
        "Extended Health Care - Total",
        "Extended Health Care - Vision",
        "Healthcare Online",
        "Critical Illness",
        "STD",
        "Dependent Life",
        "Dental"
    ];

    private static readonly HashSet<string> TierNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "Single",
        "Couple",
        "Family"
    };

    private readonly PdfTextExtractor _extractor;

    public ManulifePdfParser(PdfTextExtractor? extractor = null)
    {
        _extractor = extractor ?? new PdfTextExtractor();
    }

    public PdfExtractionResult Parse(string pdfPath)
    {
        var pages = _extractor.Extract(pdfPath);
        return ParseExtractedPages(pages, pdfPath);
    }

    internal static PdfExtractionResult ParseExtractedPages(IReadOnlyCollection<PdfPageText> pageTexts, string pdfPath = "")
    {
        var pages = pageTexts.ToList();
        var allText = ParsingHelpers.Clean(string.Join(" ", pages.Select(page => page.Text)));

        var result = new PdfExtractionResult
        {
            SourcePath = pdfPath,
            Carrier = "Manulife",
            PolicyNumber = ParsePolicyNumber(allText),
            ClientName = ParseClientName(pages, allText),
            RenewalDate = ParseRenewalDate(allText)
        };

        ParseRenewalTotals(pages, allText, result);
        ParseRateLines(pages, result);
        ParseDemographicLines(pages, result);
        ParseExperienceLines(pages, result);
        var pooledPremium = ParsePooledPremiumLines(pages, result);
        CalculatePoolingCharge(result, pooledPremium);
        ValidateNonPooledEhcSubtotalExperience(result);
        ValidateDentalExperience(result);

        if (result.RateLines.Count == 0)
        {
            result.Warnings.Add("No rate illustration rows were extracted from the PDF. The PDF may be scanned or its table layout may need a new parser rule.");
        }

        if (result.ExperienceLines.Count == 0)
        {
            result.Warnings.Add("No experience history rows were extracted from the PDF.");
        }

        return result;
    }

    private static void ParseDemographicLines(List<PdfPageText> pages, PdfExtractionResult result)
    {
        foreach (var page in pages)
        {
            if (!page.Text.Contains("Age Volume Distribution", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var benefit = DetectDemographicBenefit(page);
            if (benefit is null)
            {
                continue;
            }

            var target = benefit.Equals("Life", StringComparison.OrdinalIgnoreCase)
                ? result.LifeDemographics
                : result.LtdDemographics;

            var section = "";
            foreach (var rawLine in page.Lines)
            {
                var line = ParsingHelpers.Clean(rawLine);
                if (line.Contains("Age Volume Distribution", StringComparison.OrdinalIgnoreCase) &&
                    line.Contains("CURRENT RENEWAL", StringComparison.OrdinalIgnoreCase))
                {
                    section = "Current";
                    continue;
                }

                if (line.Contains("Age Volume Distribution", StringComparison.OrdinalIgnoreCase) &&
                    line.Contains("LAST RENEWAL", StringComparison.OrdinalIgnoreCase))
                {
                    section = "Prior";
                    continue;
                }

                if (line.StartsWith("Group Total", StringComparison.OrdinalIgnoreCase) ||
                    line.StartsWith("Ratio", StringComparison.OrdinalIgnoreCase) ||
                    line.StartsWith("Summary", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (section.Length == 0 || !TryParseDemographicRow(line, out var ageBand, out var maleLives, out var maleVolume, out var femaleLives, out var femaleVolume))
                {
                    continue;
                }

                var demographicLine = target.FirstOrDefault(row => row.AgeBand.Equals(ageBand, StringComparison.OrdinalIgnoreCase));
                if (demographicLine is null)
                {
                    demographicLine = new DemographicLine
                    {
                        Benefit = benefit,
                        AgeBand = ageBand,
                        SourceReference = $"PDF page {page.PageNumber}"
                    };
                    target.Add(demographicLine);
                }

                if (section == "Current")
                {
                    demographicLine.CurrentMaleLives = maleLives;
                    demographicLine.CurrentMaleVolume = maleVolume;
                    demographicLine.CurrentFemaleLives = femaleLives;
                    demographicLine.CurrentFemaleVolume = femaleVolume;
                }
                else
                {
                    demographicLine.PriorMaleLives = maleLives;
                    demographicLine.PriorMaleVolume = maleVolume;
                    demographicLine.PriorFemaleLives = femaleLives;
                    demographicLine.PriorFemaleVolume = femaleVolume;
                }
            }
        }
    }

    private static string? DetectDemographicBenefit(PdfPageText page)
    {
        if (page.Lines.Any(line => NormalizeBenefitText(line).Contains("LONG TERM DISABILITY", StringComparison.OrdinalIgnoreCase)))
        {
            return "Long Term Disability";
        }

        if (page.Lines.Any(line =>
            NormalizeBenefitText(line).Equals("LIFE", StringComparison.OrdinalIgnoreCase) ||
            NormalizeBenefitText(line).Contains("LIFE INSURANCE", StringComparison.OrdinalIgnoreCase)))
        {
            return "Life";
        }

        return null;
    }

    private static string NormalizeBenefitText(string text)
    {
        var normalized = ParsingHelpers.Clean(text)
            .Replace("\u2013", "-", StringComparison.Ordinal)
            .Replace("\u2014", "-", StringComparison.Ordinal)
            .Replace("-", " ", StringComparison.Ordinal);

        return Regex.Replace(normalized, @"\s+", " ").Trim().ToUpperInvariant();
    }

    private static bool TryParseDemographicRow(
        string line,
        out string ageBand,
        out int? maleLives,
        out decimal? maleVolume,
        out int? femaleLives,
        out decimal? femaleVolume)
    {
        ageBand = "";
        maleLives = null;
        maleVolume = null;
        femaleLives = null;
        femaleVolume = null;

        var match = Regex.Match(
            line,
            @"^(?<age>Under 25|25-29|30-34|35-39|40-44|45-49|50-54|55-59|60-64|65-69\+?|70\s*&\s*over|60\s*&\s*over)\s+(?<values>.+)$",
            RegexOptions.IgnoreCase);

        if (!match.Success)
        {
            return false;
        }

        var numbers = ParsingHelpers.ExtractNumbers(match.Groups["values"].Value);
        if (numbers.Count < 6)
        {
            return false;
        }

        ageBand = NormalizeAgeBand(match.Groups["age"].Value);
        maleLives = (int)numbers[0];
        maleVolume = numbers[1];
        femaleLives = (int)numbers[2];
        femaleVolume = numbers[3];
        return true;
    }

    private static string NormalizeAgeBand(string ageBand)
    {
        return ParsingHelpers.Clean(ageBand)
            .Replace(" & ", " & ", StringComparison.Ordinal)
            .Replace("  ", " ", StringComparison.Ordinal);
    }

    private static string ParsePolicyNumber(string text)
    {
        var match = Regex.Match(text, @"Policy[:\s]+(?<policy>\d+)", RegexOptions.IgnoreCase);
        return match.Success ? match.Groups["policy"].Value : "";
    }

    private static string ParseClientName(List<PdfPageText> pages, string allText)
    {
        var firstPageForLine = pages.FirstOrDefault()?.Lines
            .FirstOrDefault(line => line.StartsWith("for ", StringComparison.OrdinalIgnoreCase));

        if (!string.IsNullOrWhiteSpace(firstPageForLine))
        {
            return firstPageForLine[4..].Trim();
        }

        foreach (var line in pages.SelectMany(page => page.Lines).Take(25))
        {
            var match = Regex.Match(line, @"^(?<client>.+?)\s*Policy\s+\d+", RegexOptions.IgnoreCase);
            if (match.Success)
            {
                return match.Groups["client"].Value.Trim();
            }
        }

        var fallback = Regex.Match(allText, @"for\s+(?<client>[A-Z0-9 &.'\-]+?)\s+(?:Policy|Renewal)", RegexOptions.IgnoreCase);
        return fallback.Success ? fallback.Groups["client"].Value.Trim() : "";
    }

    private static DateOnly? ParseRenewalDate(string text)
    {
        var match = Regex.Match(text, @"Renewal Date:\s*(?<date>\d{1,2}-[A-Za-z]{3}-\d{4})", RegexOptions.IgnoreCase);
        return match.Success ? ParsingHelpers.ParseDateOnly(match.Groups["date"].Value) : null;
    }

    private static void ParseRenewalTotals(List<PdfPageText> pages, string text, PdfExtractionResult result)
    {
        ReadRateIllustrationMoneyPair(pages, text, @"Total Monthly Premium", out var currentTotal, out var renewalTotal);
        result.CurrentMonthlyPremium = currentTotal;
        result.RenewalMonthlyPremium = renewalTotal;

        var change = Regex.Match(
            text,
            @"Change in Total Monthly Premium\s*\$?(?<amount>[\d,]+\.\d{2})\s*(?<percent>-?\d+(?:\.\d+)?)\s*%",
            RegexOptions.IgnoreCase);

        if (change.Success)
        {
            result.MonthlyPremiumChange = ParsingHelpers.ParseDecimal(change.Groups["amount"].Value);
            result.MonthlyPremiumChangePercent = ParsingHelpers.ParseDecimal(change.Groups["percent"].Value);
        }

        ReadRateIllustrationMoneyPair(pages, text, @"Total Extended Health Care", out var currentEhc, out var renewalEhc);
        result.CurrentEhcPremium = currentEhc;
        result.RenewalEhcPremium = renewalEhc;

        ReadRateIllustrationMoneyPair(pages, text, @"Total Dental", out var currentDental, out var renewalDental);
        result.CurrentDentalPremium = currentDental;
        result.RenewalDentalPremium = renewalDental;
    }

    private static void ReadRateIllustrationMoneyPair(
        IEnumerable<PdfPageText> pages,
        string fallbackText,
        string label,
        out decimal? current,
        out decimal? renewal)
    {
        var ratePages = pages
            .Where(page => page.Text.Contains("Renewal Rate Illustration", StringComparison.OrdinalIgnoreCase))
            .ToList();
        var summaryPairs = ratePages
            .Where(IsRateIllustrationSummaryPage)
            .Select(page => ReadMoneyPair(page.Text, label))
            .Where(pair => pair.Current.HasValue && pair.Renewal.HasValue)
            .ToList();

        if (summaryPairs.Count > 0)
        {
            var summary = summaryPairs[^1];
            current = summary.Current;
            renewal = summary.Renewal;
            return;
        }

        var detailPairs = ratePages
            .Where(page => !IsRateIllustrationSummaryPage(page))
            .Select(page => ReadMoneyPair(page.Text, label))
            .Where(pair => pair.Current.HasValue && pair.Renewal.HasValue)
            .ToList();

        if (detailPairs.Count > 0)
        {
            current = detailPairs.Sum(pair => pair.Current!.Value);
            renewal = detailPairs.Sum(pair => pair.Renewal!.Value);
            return;
        }

        ReadMoneyPair(fallbackText, label, out current, out renewal);
    }

    private static (decimal? Current, decimal? Renewal) ReadMoneyPair(string text, string label)
    {
        ReadMoneyPair(text, label, out var current, out var renewal);
        return (current, renewal);
    }

    private static void ReadMoneyPair(string text, string label, out decimal? current, out decimal? renewal)
    {
        current = null;
        renewal = null;
        var match = Regex.Match(
            text,
            $"{Regex.Escape(label)}\\s*\\$?(?<current>[\\d,]+\\.\\d{{2}})\\s*\\$?(?<renewal>[\\d,]+\\.\\d{{2}})",
            RegexOptions.IgnoreCase);

        if (match.Success)
        {
            current = ParsingHelpers.ParseDecimal(match.Groups["current"].Value);
            renewal = ParsingHelpers.ParseDecimal(match.Groups["renewal"].Value);
        }
    }

    private static void ParseRateLines(List<PdfPageText> pages, PdfExtractionResult result)
    {
        string? parentBenefit = null;
        var accountClass = "";
        var ratePages = pages
            .Where(page => page.Text.Contains("Renewal Rate Illustration", StringComparison.OrdinalIgnoreCase))
            .ToList();
        var pagesToParse = ratePages.Any(page => !IsRateIllustrationSummaryPage(page))
            ? ratePages.Where(page => !IsRateIllustrationSummaryPage(page))
            : ratePages;

        foreach (var page in pagesToParse)
        {
            accountClass = FirstNonBlank(ParseAccountClass(page), accountClass);

            foreach (var rawLine in ExpandRateLineCandidates(page.Lines))
            {
                var line = ParsingHelpers.Clean(rawLine);
                if (line.Length == 0)
                {
                    continue;
                }

                if (line.Contains("Extended Health Care - Hospital", StringComparison.OrdinalIgnoreCase))
                {
                    parentBenefit = "Extended Health Care - Hospital";
                }
                else if (line.Contains("Extended Health Care - Vision", StringComparison.OrdinalIgnoreCase))
                {
                    parentBenefit = "Extended Health Care - Vision";
                }
                else if (line.Contains("Extended Health Care - Drugs and Other", StringComparison.OrdinalIgnoreCase))
                {
                    parentBenefit = "Extended Health Care - Drugs and Other";
                }
                else if (line.Contains("Extended Health Care - Pooled", StringComparison.OrdinalIgnoreCase))
                {
                    parentBenefit = "Extended Health Care - Pooled";
                }
                else if (line.Contains("Extended Health Care - Total", StringComparison.OrdinalIgnoreCase))
                {
                    parentBenefit = "Extended Health Care - Total";
                }
                else if (line.Equals("Extended Health Care", StringComparison.OrdinalIgnoreCase))
                {
                    parentBenefit = "Extended Health Care";
                }
                else if (line.StartsWith("Dental", StringComparison.OrdinalIgnoreCase))
                {
                    parentBenefit = "Dental";
                }

                if (TryParseRateLine(line, parentBenefit, page.PageNumber, accountClass, out var rateLine))
                {
                    result.RateLines.Add(rateLine);
                }

                if (line.Contains("Total Monthly Premium", StringComparison.OrdinalIgnoreCase))
                {
                    break;
                }
            }
        }
    }

    private static bool IsRateIllustrationSummaryPage(PdfPageText page)
    {
        return page.Text.Contains("Renewal Rate Illustration Summary", StringComparison.OrdinalIgnoreCase);
    }

    private static string FirstNonBlank(params string[] values)
    {
        return values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)) ?? "";
    }

    private static string ParseAccountClass(PdfPageText page)
    {
        var line = page.Lines
            .Select(ParsingHelpers.Clean)
            .FirstOrDefault(value => value.StartsWith("ACCOUNT/CLASS:", StringComparison.OrdinalIgnoreCase) ||
                                     value.StartsWith("ACCOUNT:", StringComparison.OrdinalIgnoreCase));

        return line ?? "";
    }

    private static IEnumerable<string> ExpandRateLineCandidates(IEnumerable<string> lines)
    {
        foreach (var rawLine in lines)
        {
            var line = ParsingHelpers.Clean(rawLine);
            if (line.Length == 0)
            {
                continue;
            }

            foreach (var sectionBenefit in BenefitNames.Where(benefit => benefit.StartsWith("Extended Health Care", StringComparison.OrdinalIgnoreCase) || benefit.Equals("Dental", StringComparison.OrdinalIgnoreCase)))
            {
                foreach (var tier in TierNames)
                {
                    line = Regex.Replace(
                        line,
                        $"{Regex.Escape(sectionBenefit)}(?={tier}\\b)",
                        $"{sectionBenefit} ",
                        RegexOptions.IgnoreCase);
                }
            }

            foreach (var benefit in RateSectionBenefits)
            {
                line = Regex.Replace(
                    line,
                    $"(?<!^)(?<![A-Za-z])(?={Regex.Escape(benefit)}\\b)",
                    "\n",
                    RegexOptions.IgnoreCase);
            }

            line = Regex.Replace(line, "(?<=CHANGE)Life\\b", "\nLife", RegexOptions.IgnoreCase);
            line = Regex.Replace(line, "(?<=[%\\d])(?=Couple\\b|Family\\b)", "\n", RegexOptions.IgnoreCase);

            foreach (var candidate in line.Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
            {
                yield return ParsingHelpers.Clean(candidate);
            }
        }
    }

    private static bool TryParseRateLine(string line, string? parentBenefit, int pageNumber, string accountClass, out RateLine rateLine)
    {
        rateLine = new RateLine();

        line = TrimRateLineTail(line);

        if (line.StartsWith("Total ", StringComparison.OrdinalIgnoreCase) ||
            line.StartsWith("*Total ", StringComparison.OrdinalIgnoreCase) ||
            line.StartsWith("Extended Health Care - Total", StringComparison.OrdinalIgnoreCase) ||
            line.Contains("Renewal Rate Illustration", StringComparison.OrdinalIgnoreCase) ||
            line.Contains("Monthly Premium", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        var benefit = "";
        var tier = "";
        var remainder = line;

        foreach (var knownBenefit in BenefitNames.OrderByDescending(value => value.Length))
        {
            if (line.StartsWith(knownBenefit, StringComparison.OrdinalIgnoreCase))
            {
                benefit = knownBenefit;
                remainder = line[knownBenefit.Length..].Trim();
                break;
            }
        }

        if (string.IsNullOrWhiteSpace(benefit))
        {
            var firstWord = line.Split(' ', StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() ?? "";
            if (!TierNames.Contains(firstWord) || string.IsNullOrWhiteSpace(parentBenefit))
            {
                return false;
            }

            benefit = parentBenefit;
            tier = firstWord;
            remainder = line[firstWord.Length..].Trim();

            if (benefit.Equals("Extended Health Care - Total", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }
        }
        else
        {
            if (benefit.Equals("Extended Health Care - Total", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            foreach (var tierName in TierNames)
            {
                if (remainder.StartsWith(tierName, StringComparison.OrdinalIgnoreCase))
                {
                    tier = tierName;
                    remainder = remainder[tierName.Length..].Trim();
                    break;
                }
            }
        }

        if (!TryParseRateValues(benefit, remainder, out var volume, out var lives, out var currentRate, out var currentPremium, out var renewalRate, out var renewalPremium, out var percentChange))
        {
            var numbers = ParsingHelpers.ExtractNumbers(remainder);
            if (numbers.Count < 5)
            {
                return false;
            }

            if (numbers.Count >= 7)
            {
                volume = numbers[0];
                lives = (int)numbers[1];
                currentRate = numbers[2];
                currentPremium = numbers[3];
                renewalRate = numbers[4];
                renewalPremium = numbers[5];
                percentChange = numbers[6];
            }
            else
            {
                lives = (int)numbers[0];
                currentRate = numbers[1];
                currentPremium = numbers[2];
                renewalRate = numbers[3];
                renewalPremium = numbers[4];
                percentChange = numbers.Count > 5 ? numbers[5] : null;
            }
        }

        rateLine.Benefit = benefit;
        rateLine.TierOrClass = tier;
        rateLine.Volume = volume;
        rateLine.Lives = lives;
        rateLine.CurrentRate = currentRate;
        rateLine.CurrentPremium = currentPremium;
        rateLine.RenewalRate = renewalRate;
        rateLine.RenewalPremium = renewalPremium;
        rateLine.PercentChange = percentChange;
        rateLine.SourceReference = $"PDF page {pageNumber}";
        rateLine.AccountClass = accountClass;

        return true;
    }

    private static string TrimRateLineTail(string line)
    {
        var cleaned = ParsingHelpers.Clean(line);
        foreach (var marker in new[]
                 {
                     "*Total Extended Health Care",
                     "Total Extended Health Care",
                     "Total Dental",
                     "Total Monthly Premium",
                     "Change in Total Monthly Premium",
                     "Date Created:",
                     "The Manufacturers Life Insurance Company"
                 })
        {
            var index = cleaned.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
            if (index > 0)
            {
                cleaned = cleaned[..index].Trim();
            }
        }

        return cleaned;
    }

    private static bool TryParseRateValues(
        string benefit,
        string remainder,
        out decimal? volume,
        out int? lives,
        out decimal? currentRate,
        out decimal? currentPremium,
        out decimal? renewalRate,
        out decimal? renewalPremium,
        out decimal? percentChange)
    {
        volume = null;
        lives = null;
        currentRate = null;
        currentPremium = null;
        renewalRate = null;
        renewalPremium = null;
        percentChange = null;

        var text = TrimRateLineTail(remainder);
        if (text.Length == 0)
        {
            return false;
        }

        var hasPercentMarker = text.Contains('%');
        if (BenefitUsesVolume(benefit))
        {
            var volumeMatch = Regex.Match(text, @"^\s*(?<volume>\d{1,3}(?:,\d{3})+|\d{4,})\s*(?<rest>.*)$");
            if (!volumeMatch.Success)
            {
                return false;
            }

            volume = ParsingHelpers.ParseDecimal(volumeMatch.Groups["volume"].Value);
            text = volumeMatch.Groups["rest"].Value;
        }

        var compact = Regex.Replace(text, @"[\s,$%]", "", RegexOptions.None);
        if (compact.Length == 0)
        {
            return false;
        }

        for (var prefixLength = 1; prefixLength <= Math.Min(4, compact.Length); prefixLength++)
        {
            var livesText = compact[..prefixLength];
            if (!livesText.All(char.IsDigit))
            {
                break;
            }

            if (!int.TryParse(livesText, out var livesCandidate) || livesCandidate <= 0)
            {
                continue;
            }

            var basis = volume.HasValue ? volume.Value / 1000m : livesCandidate;
            var valueText = compact[prefixLength..];
            if (!TrySplitRatePremiumPairs(valueText, basis, out var parsedCurrentRate, out var parsedCurrentPremium, out var parsedRenewalRate, out var parsedRenewalPremium, out var trailing))
            {
                continue;
            }

            var parsedPercent = string.IsNullOrWhiteSpace(trailing) ? null : ParsingHelpers.ParseDecimal(trailing);
            if (hasPercentMarker && !parsedPercent.HasValue)
            {
                continue;
            }

            if (!hasPercentMarker && !string.IsNullOrWhiteSpace(trailing))
            {
                continue;
            }

            lives = livesCandidate;
            currentRate = parsedCurrentRate;
            currentPremium = parsedCurrentPremium;
            renewalRate = parsedRenewalRate;
            renewalPremium = parsedRenewalPremium;
            percentChange = parsedPercent;
            return true;
        }

        return false;
    }

    private static bool TrySplitRatePremiumPairs(
        string compact,
        decimal basis,
        out decimal currentRate,
        out decimal currentPremium,
        out decimal renewalRate,
        out decimal renewalPremium,
        out string trailing)
    {
        currentRate = 0m;
        currentPremium = 0m;
        renewalRate = 0m;
        renewalPremium = 0m;
        trailing = "";

        foreach (var currentRateCandidate in EnumerateRatePrefixes(compact))
        {
            var expectedCurrentPremium = decimal.Round(basis * currentRateCandidate.Value, 2, MidpointRounding.AwayFromZero);
            var currentPremiumText = FormatCompactMoney(expectedCurrentPremium);
            var afterCurrentRate = compact[currentRateCandidate.Text.Length..];
            if (!afterCurrentRate.StartsWith(currentPremiumText, StringComparison.Ordinal))
            {
                continue;
            }

            var afterCurrentPremium = afterCurrentRate[currentPremiumText.Length..];
            foreach (var renewalRateCandidate in EnumerateRatePrefixes(afterCurrentPremium))
            {
                var expectedRenewalPremium = decimal.Round(basis * renewalRateCandidate.Value, 2, MidpointRounding.AwayFromZero);
                var renewalPremiumText = FormatCompactMoney(expectedRenewalPremium);
                var afterRenewalRate = afterCurrentPremium[renewalRateCandidate.Text.Length..];
                if (!afterRenewalRate.StartsWith(renewalPremiumText, StringComparison.Ordinal))
                {
                    continue;
                }

                currentRate = currentRateCandidate.Value;
                currentPremium = expectedCurrentPremium;
                renewalRate = renewalRateCandidate.Value;
                renewalPremium = expectedRenewalPremium;
                trailing = afterRenewalRate[renewalPremiumText.Length..];
                return true;
            }
        }

        return false;
    }

    private static IEnumerable<(string Text, decimal Value)> EnumerateRatePrefixes(string valueText)
    {
        for (var length = 3; length <= Math.Min(8, valueText.Length); length++)
        {
            var candidate = valueText[..length];
            if (!Regex.IsMatch(candidate, @"^\d{1,3}\.\d{1,3}$"))
            {
                continue;
            }

            var value = ParsingHelpers.ParseDecimal(candidate);
            if (value.HasValue)
            {
                yield return (candidate, value.Value);
            }
        }
    }

    private static string FormatCompactMoney(decimal value)
    {
        return value.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture);
    }

    private static bool BenefitUsesVolume(string benefit)
    {
        return benefit.Equals("Life", StringComparison.OrdinalIgnoreCase) ||
               benefit.Equals("Accidental Death and Dismemberment", StringComparison.OrdinalIgnoreCase) ||
               benefit.Equals("Long Term Disability", StringComparison.OrdinalIgnoreCase) ||
               benefit.Equals("Short Term Disability", StringComparison.OrdinalIgnoreCase) ||
               benefit.Equals("STD", StringComparison.OrdinalIgnoreCase) ||
               benefit.Equals("Group Critical Illness", StringComparison.OrdinalIgnoreCase) ||
               benefit.Equals("Critical Illness", StringComparison.OrdinalIgnoreCase);
    }

    private static void ParseExperienceLines(List<PdfPageText> pages, PdfExtractionResult result)
    {
        string? currentBenefit = null;
        string? pendingBenefit = null;
        var currentAccountCanSupplyExperience = false;

        foreach (var page in pages)
        {
            var hasAverageLivesColumn = PageHasAverageLivesColumn(page);

            for (var index = 0; index < page.Lines.Count; index++)
            {
                var rawLine = page.Lines[index];
                var line = ParsingHelpers.Clean(rawLine);

                if (TryDetectAccountScope(line, out var accountIsAll, out var isAccountClass))
                {
                    currentAccountCanSupplyExperience = accountIsAll || isAccountClass;
                    currentBenefit = currentAccountCanSupplyExperience ? pendingBenefit : null;
                    continue;
                }

                if (IsNonPooledEhcSubtotalHeader(page.Lines, index))
                {
                    pendingBenefit = "Non Pooled Extended Health Care - Subtotal";
                    currentBenefit = currentAccountCanSupplyExperience ? pendingBenefit : null;
                    continue;
                }

                if (IsDentalExperienceHeader(line))
                {
                    pendingBenefit = "Dental";
                    currentBenefit = currentAccountCanSupplyExperience ? pendingBenefit : null;
                    continue;
                }

                var benefitFromLine = DetectBenefit(line);
                if (!string.IsNullOrWhiteSpace(benefitFromLine))
                {
                    if (benefitFromLine.Equals("Dental", StringComparison.OrdinalIgnoreCase))
                    {
                        currentBenefit = null;
                        continue;
                    }

                    pendingBenefit = benefitFromLine;
                    currentBenefit = benefitFromLine;
                }

                if (currentBenefit is not null &&
                    TryParseExperienceLine(line, currentBenefit, page.PageNumber, hasAverageLivesColumn, out var experienceLine))
                {
                    AddOrReplaceExperienceLine(result, experienceLine);
                }
            }

            ParseCurrentRenewalExperienceBlock(page, result);
            ApplyRenewalFactors(page, result);
        }
    }

    private static void ValidateNonPooledEhcSubtotalExperience(PdfExtractionResult result)
    {
        ValidateExperienceRows(
            result,
            "Non Pooled Extended Health Care - Subtotal",
            "Non Pooled Extended Health Care - Subtotal experience");
    }

    private static void ValidateDentalExperience(PdfExtractionResult result)
    {
        ValidateExperienceRows(result, "Dental", "Dental experience");
    }

    private static void ValidateExperienceRows(PdfExtractionResult result, string benefit, string warningLabel)
    {
        var rows = result.ExperienceLines
            .Where(row => row.Benefit.Equals(benefit, StringComparison.OrdinalIgnoreCase))
            .GroupBy(row => row.Period, StringComparer.OrdinalIgnoreCase)
            .Select(group => group.Last())
            .Select(row => new
            {
                Row = row,
                StartDate = ParsingHelpers.ParseDateOnly(row.Period.Split(" - ", StringSplitOptions.None).FirstOrDefault())
            })
            .OrderBy(item => item.StartDate)
            .ToList();

        if (rows.Count == 0)
        {
            return;
        }

        var validRows = rows
            .Where(item => item.StartDate.HasValue)
            .TakeLast(3)
            .ToList();

        result.ExperienceLines.RemoveAll(row => row.Benefit.Equals(benefit, StringComparison.OrdinalIgnoreCase));

        if (validRows.Count == 0)
        {
            result.Warnings.Add($"{warningLabel} was not used because the policy period dates could not be read.");
            return;
        }

        result.ExperienceLines.AddRange(validRows.Select(item => item.Row));

        var invalidDateCount = rows.Count(item => !item.StartDate.HasValue);
        if (invalidDateCount > 0)
        {
            result.Warnings.Add($"{warningLabel} included {invalidDateCount} row(s) whose dates could not be read; those row(s) were skipped.");
        }

        if (validRows.Count < 3)
        {
            result.Warnings.Add($"{warningLabel} has {validRows.Count} extracted period(s). The current renewal values will be used where available; missing prior years need manual review.");
        }
    }

    private static bool TryDetectAccountScope(string line, out bool isAll, out bool isAccountClass)
    {
        var normalized = NormalizeSectionText(line)
            .Replace(":", " ", StringComparison.Ordinal)
            .Replace("/", " ", StringComparison.Ordinal);

        if (!Regex.IsMatch(normalized, @"\bACCOUNT\b", RegexOptions.IgnoreCase))
        {
            isAll = false;
            isAccountClass = false;
            return false;
        }

        isAll = Regex.IsMatch(normalized, @"\bACCOUNT\b\s+ALL\b", RegexOptions.IgnoreCase);
        isAccountClass = Regex.IsMatch(normalized, @"\bACCOUNT\s+CLASS\b", RegexOptions.IgnoreCase);
        return true;
    }

    private static bool IsNonPooledEhcSubtotalHeader(IReadOnlyList<string> lines, int index)
    {
        var current = NormalizeSectionText(lines[index]);
        var next = index + 1 < lines.Count ? NormalizeSectionText(lines[index + 1]) : "";
        var combined = $"{current} {next}".Trim();

        return combined.Contains("NON POOLED EXTENDED HEALTH CARE SUBTOTAL", StringComparison.OrdinalIgnoreCase) ||
               combined.Contains("NON-POOLED EXTENDED HEALTH CARE SUBTOTAL", StringComparison.OrdinalIgnoreCase) ||
               combined.Contains("EXPERIENCE HISTORY NON POOLED EXTENDED HEALTH CARE", StringComparison.OrdinalIgnoreCase) ||
               combined.Contains("EXPERIENCE HISTORY NON-POOLED EXTENDED HEALTH CARE", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsDentalExperienceHeader(string line)
    {
        return NormalizeSectionText(line).Equals("DENTAL", StringComparison.OrdinalIgnoreCase);
    }

    private static void ParseCurrentRenewalExperienceBlock(PdfPageText page, PdfExtractionResult result)
    {
        var normalizedPage = NormalizeSectionText(page.Text);

        if (IsHealthRenewalAnalysisPage(normalizedPage))
        {
            ParseFirstExperienceCandidate(page, result, "Non Pooled Extended Health Care - Subtotal");
        }

        if (IsDentalRenewalAnalysisPage(normalizedPage))
        {
            ParseFirstExperienceCandidate(page, result, "Dental");
        }
    }

    private static bool IsHealthRenewalAnalysisPage(string normalizedPage)
    {
        return normalizedPage.Contains("EXPERIENCE HISTORY NON POOLED EXTENDED HEALTH CARE", StringComparison.OrdinalIgnoreCase) ||
               normalizedPage.Contains("EXPERIENCE HISTORY NON-POOLED EXTENDED HEALTH CARE", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsDentalRenewalAnalysisPage(string normalizedPage)
    {
        return Regex.IsMatch(normalizedPage, @"\bDENTAL\b", RegexOptions.IgnoreCase) &&
               normalizedPage.Contains("EXPERIENCE HISTORY", StringComparison.OrdinalIgnoreCase) &&
               normalizedPage.Contains("RENEWAL FACTORS", StringComparison.OrdinalIgnoreCase);
    }

    private static void ParseFirstExperienceCandidate(PdfPageText page, PdfExtractionResult result, string benefit)
    {
        var hasAverageLivesColumn = PageHasAverageLivesColumn(page);

        for (var index = 0; index < page.Lines.Count; index++)
        {
            var candidate = CombineLineWindow(page.Lines, index, 3);
            if (DateRegex.Matches(candidate).Count < 2)
            {
                continue;
            }

            if (TryParseExperienceLine(candidate, benefit, page.PageNumber, hasAverageLivesColumn, out var experienceLine))
            {
                AddOrReplaceExperienceLine(result, experienceLine);
                return;
            }
        }
    }

    private static bool PageHasAverageLivesColumn(PdfPageText page)
    {
        return NormalizeSectionText(page.Text)
            .Contains("AVERAGE NUMBER OF LIVES", StringComparison.OrdinalIgnoreCase);
    }

    private static void ApplyRenewalFactors(PdfPageText page, PdfExtractionResult result)
    {
        var normalizedPage = NormalizeSectionText(page.Text);
        var benefit = IsHealthRenewalAnalysisPage(normalizedPage)
            ? "Non Pooled Extended Health Care - Subtotal"
            : IsDentalRenewalAnalysisPage(normalizedPage)
                ? "Dental"
                : null;

        if (benefit is null)
        {
            return;
        }

        var trendFactor = ReadPercentValueAfterLabel(page.Lines, "Trend Factor");
        var credibility = ReadPercentValueAfterLabel(page.Lines, "Credibility");
        var targetLossRatio = ReadPercentValueAfterLabel(page.Lines, "Target Loss Ratio");
        var rateAdjustment = ReadPercentValueAfterLabel(page.Lines, "Rate Adjustment");

        if (!trendFactor.HasValue &&
            !credibility.HasValue &&
            !targetLossRatio.HasValue &&
            !rateAdjustment.HasValue)
        {
            return;
        }

        var sourceReference = $"PDF page {page.PageNumber}";
        var targetRow = result.ExperienceLines
            .Where(row => row.Benefit.Equals(benefit, StringComparison.OrdinalIgnoreCase) &&
                          row.SourceReference.Contains(sourceReference, StringComparison.OrdinalIgnoreCase))
            .OrderBy(row => ParsingHelpers.ParseDateOnly(row.Period.Split(" - ", StringSplitOptions.None).FirstOrDefault()))
            .LastOrDefault()
            ?? result.ExperienceLines
                .Where(row => row.Benefit.Equals(benefit, StringComparison.OrdinalIgnoreCase))
                .OrderBy(row => ParsingHelpers.ParseDateOnly(row.Period.Split(" - ", StringSplitOptions.None).FirstOrDefault()))
                .LastOrDefault();

        if (targetRow is null)
        {
            return;
        }

        targetRow.TrendFactor = trendFactor ?? targetRow.TrendFactor;
        targetRow.Credibility = credibility ?? targetRow.Credibility;
        targetRow.TargetLossRatio = targetLossRatio ?? targetRow.TargetLossRatio;
        targetRow.ProposedAdjustment = rateAdjustment ?? targetRow.ProposedAdjustment;
    }

    private static decimal? ReadPercentValueAfterLabel(IReadOnlyList<string> lines, string label)
    {
        for (var index = 0; index < lines.Count; index++)
        {
            var line = ParsingHelpers.Clean(lines[index]);
            var labelIndex = line.IndexOf(label, StringComparison.OrdinalIgnoreCase);
            if (labelIndex < 0)
            {
                continue;
            }

            var tail = line[(labelIndex + label.Length)..];
            var values = ParsingHelpers.ExtractNumbers(tail);
            if (values.Count == 0 && index + 1 < lines.Count)
            {
                values = ParsingHelpers.ExtractNumbers(lines[index + 1]);
            }

            if (values.Count > 0)
            {
                return values[0];
            }
        }

        return null;
    }

    private static PooledPremiumCandidate? ParsePooledPremiumLines(
        List<PdfPageText> pages,
        PdfExtractionResult result)
    {
        var candidates = new List<PooledPremiumCandidate>();

        foreach (var page in pages)
        {
            var inPooledEhcSection = false;

            for (var index = 0; index < page.Lines.Count; index++)
            {
                var line = ParsingHelpers.Clean(page.Lines[index]);
                var normalizedLine = NormalizeSectionText(line);

                if ((normalizedLine.Equals("POOLED", StringComparison.OrdinalIgnoreCase) ||
                     normalizedLine.Contains("POOLED EXTENDED HEALTH CARE", StringComparison.OrdinalIgnoreCase)) &&
                    !normalizedLine.Contains("NON POOLED", StringComparison.OrdinalIgnoreCase))
                {
                    inPooledEhcSection = true;
                    continue;
                }

                if (!inPooledEhcSection)
                {
                    continue;
                }

                var detectedBenefit = DetectBenefit(line);
                if (!string.IsNullOrWhiteSpace(detectedBenefit) &&
                    !detectedBenefit.Contains("Pooled", StringComparison.OrdinalIgnoreCase) &&
                    !detectedBenefit.Contains("Extended Health Care", StringComparison.OrdinalIgnoreCase))
                {
                    inPooledEhcSection = false;
                    continue;
                }

                var candidate = CombineLineWindow(page.Lines, index, 3);
                if (TryParsePooledPremiumLine(candidate, page.PageNumber, out var pooledPremiumCandidate) &&
                    !candidates.Any(existing =>
                        existing.Period.Equals(pooledPremiumCandidate.Period, StringComparison.OrdinalIgnoreCase) &&
                        existing.PooledPremium == pooledPremiumCandidate.PooledPremium))
                {
                    candidates.Add(pooledPremiumCandidate);
                }
            }
        }

        var selected = candidates
            .Where(candidate => candidate.PeriodStart.HasValue)
            .OrderByDescending(candidate => candidate.PeriodStart)
            .FirstOrDefault()
            ?? candidates.LastOrDefault();

        if (selected is null)
        {
            return null;
        }

        result.PooledPremium = selected.PooledPremium;
        result.DrugStopLossClaims = selected.DrugStopLossClaims;
        result.OtherPooledClaims = selected.OtherPooledClaims;
        return selected;
    }

    private static bool TryParsePooledPremiumLine(
        string line,
        int pageNumber,
        out PooledPremiumCandidate candidate)
    {
        candidate = null!;
        var dateMatches = DateRegex.Matches(line);
        if (dateMatches.Count < 2)
        {
            return false;
        }

        var secondDate = dateMatches[1];
        var tail = line[(secondDate.Index + secondDate.Length)..];
        var numbers = ParsingHelpers.ExtractNumbers(tail);
        if (numbers.Count < 3)
        {
            return false;
        }

        var period = $"{dateMatches[0].Value} - {dateMatches[1].Value}";
        candidate = new PooledPremiumCandidate(
            period,
            ParsingHelpers.ParseDateOnly(dateMatches[0].Value),
            numbers[0],
            numbers[1],
            numbers[2],
            pageNumber);
        return true;
    }

    private static void CalculatePoolingCharge(
        PdfExtractionResult result,
        PooledPremiumCandidate? pooledPremium)
    {
        if (result.PoolingCharge.HasValue ||
            pooledPremium is null ||
            result.PooledPremium is not > 0m)
        {
            return;
        }

        var latestNonPooled = result.ExperienceLines
            .Where(row => row.Benefit.Equals(
                "Non Pooled Extended Health Care - Subtotal",
                StringComparison.OrdinalIgnoreCase))
            .Where(row => row.Period.Equals(pooledPremium.Period, StringComparison.OrdinalIgnoreCase))
            .Where(row => row.BilledPremium is > 0m)
            .FirstOrDefault();

        if (latestNonPooled is null)
        {
            result.Warnings.Add(
                $"Pooling charge was not calculated because pooled period {pooledPremium.Period} on PDF page {pooledPremium.SourcePage} did not match an extracted non-pooled Extended Health Care billed-premium period.");
            return;
        }

        result.PoolingCharge = decimal.Round(
            result.PooledPremium.Value / latestNonPooled.BilledPremium!.Value * 100m,
            2,
            MidpointRounding.AwayFromZero);
    }

    private sealed record PooledPremiumCandidate(
        string Period,
        DateOnly? PeriodStart,
        decimal PooledPremium,
        decimal DrugStopLossClaims,
        decimal OtherPooledClaims,
        int SourcePage);

    private static string CombineLineWindow(IReadOnlyList<string> lines, int startIndex, int lineCount)
    {
        return ParsingHelpers.Clean(string.Join(" ", lines.Skip(startIndex).Take(lineCount)));
    }

    private static void AddOrReplaceExperienceLine(PdfExtractionResult result, ExperienceLine candidate)
    {
        var existingIndex = result.ExperienceLines.FindIndex(row =>
            row.Benefit.Equals(candidate.Benefit, StringComparison.OrdinalIgnoreCase) &&
            row.Period.Equals(candidate.Period, StringComparison.OrdinalIgnoreCase));

        if (existingIndex < 0)
        {
            result.ExperienceLines.Add(candidate);
            return;
        }

        result.ExperienceLines[existingIndex] = MergeExperienceLine(result.ExperienceLines[existingIndex], candidate);
    }

    private static ExperienceLine MergeExperienceLine(ExperienceLine existing, ExperienceLine candidate)
    {
        return new ExperienceLine
        {
            Benefit = candidate.Benefit,
            Period = candidate.Period,
            AverageLives = candidate.AverageLives ?? existing.AverageLives,
            BilledPremium = candidate.BilledPremium ?? existing.BilledPremium,
            PaidClaims = candidate.PaidClaims ?? existing.PaidClaims,
            IncurredClaims = candidate.IncurredClaims ?? existing.IncurredClaims,
            PaidLossRatio = candidate.PaidLossRatio ?? existing.PaidLossRatio,
            IncurredLossRatio = candidate.IncurredLossRatio ?? existing.IncurredLossRatio,
            Weighting = candidate.Weighting ?? existing.Weighting,
            ProposedAdjustment = candidate.ProposedAdjustment ?? existing.ProposedAdjustment,
            TrendFactor = candidate.TrendFactor ?? existing.TrendFactor,
            Credibility = candidate.Credibility ?? existing.Credibility,
            TargetLossRatio = candidate.TargetLossRatio ?? existing.TargetLossRatio,
            SourceReference = string.IsNullOrWhiteSpace(candidate.SourceReference)
                ? existing.SourceReference
                : candidate.SourceReference
        };
    }

    private static string NormalizeSectionText(string text)
    {
        var normalized = ParsingHelpers.Clean(text)
            .Replace("\u2013", "-", StringComparison.Ordinal)
            .Replace("\u2014", "-", StringComparison.Ordinal)
            .Replace("-", " ", StringComparison.Ordinal)
            .ToUpperInvariant();

        return Regex.Replace(normalized, @"\s+", " ").Trim();
    }

    private static string? DetectBenefit(string line)
    {
        foreach (var benefit in BenefitNames.OrderByDescending(value => value.Length))
        {
            if (line.Contains(benefit, StringComparison.OrdinalIgnoreCase))
            {
                return benefit switch
                {
                    "Hospital" => "Extended Health Care - Hospital",
                    "Vision" => "Extended Health Care - Vision",
                    "Drugs and Other" => "Extended Health Care - Drugs and Other",
                    "Pooled" => "Extended Health Care - Pooled",
                    _ => benefit
                };
            }
        }

        return null;
    }

    private static bool TryParseExperienceLine(
        string line,
        string benefit,
        int pageNumber,
        bool hasAverageLivesColumn,
        out ExperienceLine experienceLine)
    {
        experienceLine = new ExperienceLine();
        var dateMatches = DateRegex.Matches(line);
        if (dateMatches.Count < 2)
        {
            return false;
        }

        var secondDate = dateMatches[1];
        var tail = line[(secondDate.Index + secondDate.Length)..];
        var numbers = ParsingHelpers.ExtractNumbers(tail);
        if (numbers.Count < 4)
        {
            return false;
        }

        experienceLine.Benefit = benefit;
        experienceLine.Period = $"{dateMatches[0].Value} - {dateMatches[1].Value}";
        experienceLine.SourceReference = $"PDF page {pageNumber}";

        if (numbers.Count >= 5 &&
            (numbers.Count >= 6 || hasAverageLivesColumn || BenefitUsesAverageLivesExperience(benefit)) &&
            LooksLikeDollarAmount(numbers[3]) &&
            LooksLikePercentNumber(numbers[4]))
        {
            experienceLine.AverageLives = (int)numbers[0];
            experienceLine.BilledPremium = numbers[1];
            experienceLine.PaidClaims = numbers[2];
            experienceLine.IncurredClaims = numbers[3];
            experienceLine.IncurredLossRatio = numbers[4];
            experienceLine.Weighting = numbers.Count > 5 ? numbers[5] : null;
        }
        else if (numbers.Count >= 5 && LooksLikeDollarAmount(numbers[2]) && LooksLikePercentNumber(numbers[3]))
        {
            experienceLine.BilledPremium = numbers[0];
            experienceLine.PaidClaims = numbers[1];
            experienceLine.IncurredClaims = numbers[2];
            experienceLine.IncurredLossRatio = numbers[3];
            experienceLine.Weighting = numbers[4];
        }
        else if (numbers.Count >= 8)
        {
            experienceLine.AverageLives = (int)numbers[0];
            experienceLine.BilledPremium = numbers[1];
            experienceLine.PaidClaims = numbers[3];
            experienceLine.IncurredClaims = numbers[4];
            experienceLine.IncurredLossRatio = numbers[6];
            experienceLine.Weighting = numbers[7];
        }
        else if (numbers.Count >= 5)
        {
            experienceLine.BilledPremium = numbers[0];
            experienceLine.PaidClaims = numbers[1];
            experienceLine.PaidLossRatio = numbers[2];
            experienceLine.IncurredClaims = numbers[3];
            experienceLine.IncurredLossRatio = numbers[4];
        }
        else
        {
            experienceLine.BilledPremium = numbers[0];
            experienceLine.PaidClaims = numbers[1];
            experienceLine.IncurredClaims = numbers.Count > 2 ? numbers[2] : null;
            experienceLine.IncurredLossRatio = numbers.Count > 3 ? numbers[3] : null;
        }

        if (!experienceLine.PaidLossRatio.HasValue && experienceLine.BilledPremium is > 0 && experienceLine.PaidClaims.HasValue)
        {
            experienceLine.PaidLossRatio = decimal.Round(experienceLine.PaidClaims.Value / experienceLine.BilledPremium.Value * 100m, 1);
        }

        return true;
    }

    private static bool BenefitUsesAverageLivesExperience(string benefit)
    {
        return benefit.Equals("Long Term Disability", StringComparison.OrdinalIgnoreCase) ||
               benefit.Equals("Short Term Disability", StringComparison.OrdinalIgnoreCase) ||
               benefit.Equals("STD", StringComparison.OrdinalIgnoreCase);
    }

    private static bool LooksLikeDollarAmount(decimal value)
    {
        return Math.Abs(value) >= 100m;
    }

    private static bool LooksLikePercentNumber(decimal value)
    {
        return value is >= -300m and <= 300m;
    }
}

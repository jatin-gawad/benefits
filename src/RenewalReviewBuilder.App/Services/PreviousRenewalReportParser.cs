using System.Globalization;
using System.Text.RegularExpressions;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed class PreviousRenewalReportParser
{
    private static readonly Regex ClientPolicyRegex = new(
        @"(?<client>[A-Za-z0-9 &'.,()/\-]+?)\s*\|\s*Policy\s+(?<policy>[A-Za-z0-9\-]+)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex ClientRegex = new(
        @"(?:Client|Client Name)\s*[:\-]\s*(?<value>[A-Za-z0-9 &'.,()/\-]+?)(?=\s+(?:Policy|Carrier|Renewal|Current|Generated)\b|$)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex PolicyRegex = new(
        @"(?:Policy|Policy Number|Policy No\.?)\s*[:\-]?\s*(?<value>[A-Za-z0-9\-]+)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex CarrierRegex = new(
        @"(?:Carrier|Insurer)\s*[:\-]\s*(?<value>[A-Za-z][A-Za-z0-9 &'./\-]+?)(?=\s+(?:Renewal|Current|Policy|Generated)\b|$)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex RenewalDateRegex = new(
        @"(?:Renewal date|Renewal Date|Renewal Effective|Effective Date|Implementation Date)\s*[:\-]\s*(?<value>[A-Za-z]{3,9}\.?\s+\d{1,2},?\s+\d{4}|\d{1,2}[-/][A-Za-z]{3,9}[-/]\d{2,4}|\d{1,2}/\d{1,2}/\d{2,4}|\d{4}-\d{1,2}-\d{1,2})",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex CurrentMonthlyPremiumRegex = new(
        @"(?:Current monthly premium|Current Monthly Premium|Current Premium)\s*[:\-]?\s*(?<value>\(?\$?\s*[\d,]+(?:\.\d{2})?\)?)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex RenewalMonthlyPremiumRegex = new(
        @"(?:Renewal monthly premium|Renewal Monthly Premium|Renewal Premium|Carrier Proposal|Carrier Proposed|Proposed Premium)\s*[:\-]?\s*(?<value>\(?\$?\s*[\d,]+(?:\.\d{2})?\)?)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex MonthlyChangeRegex = new(
        @"(?:Monthly change|Monthly Change|Premium Change)\s*[:\-]?\s*(?<money>\(?\$?\s*[\d,]+(?:\.\d{2})?\)?)(?:\s*\((?<percent>[+-]?\d+(?:\.\d+)?)%\))?",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex PercentChangeRegex = new(
        @"(?:Monthly Change %|Percent Change|Premium Change %|Renewal Change|Final Renewal Change)\s*[:\-]?\s*\+?(?<value>[+-]?\d+(?:\.\d+)?)%",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private readonly PdfTextExtractor _extractor;

    public PreviousRenewalReportParser(PdfTextExtractor? extractor = null)
    {
        _extractor = extractor ?? new PdfTextExtractor();
    }

    public PreviousRenewalSnapshot? Parse(string pdfPath)
    {
        var pages = _extractor.Extract(pdfPath);
        return ParsePages(pdfPath, pages);
    }

    public PreviousRenewalSnapshot? ParsePages(string pdfPath, IReadOnlyList<PdfPageText> pages)
    {
        var allText = ParsingHelpers.Clean(string.Join(" ", pages.Select(page => page.Text)));
        if (!LooksLikeRenewalReview(allText, pdfPath))
        {
            return null;
        }

        var snapshot = new PreviousRenewalSnapshot
        {
            SourcePath = pdfPath,
            SourceFileName = Path.GetFileName(pdfPath),
            SourceType = "Previous renewal report",
            ClientName = Match(ClientRegex, allText),
            PolicyNumber = Match(PolicyRegex, allText),
            Carrier = Match(CarrierRegex, allText),
            RenewalDate = ParseDate(Match(RenewalDateRegex, allText)),
            CurrentMonthlyPremium = MatchMoney(CurrentMonthlyPremiumRegex, allText),
            RenewalMonthlyPremium = MatchMoney(RenewalMonthlyPremiumRegex, allText),
            SourceReference = BuildSourceReference(pdfPath, pages)
        };

        var clientPolicy = ClientPolicyRegex.Match(allText);
        if (clientPolicy.Success)
        {
            snapshot.ClientName = FirstNonBlank(snapshot.ClientName, ParsingHelpers.Clean(clientPolicy.Groups["client"].Value));
            snapshot.PolicyNumber = FirstNonBlank(snapshot.PolicyNumber, ParsingHelpers.Clean(clientPolicy.Groups["policy"].Value));
        }

        var change = MonthlyChangeRegex.Match(allText);
        if (change.Success)
        {
            snapshot.MonthlyPremiumChange = ParsingHelpers.ParseDecimal(change.Groups["money"].Value);
            snapshot.MonthlyPremiumChangePercent = ParsingHelpers.ParseDecimal(change.Groups["percent"].Value);
        }

        snapshot.MonthlyPremiumChangePercent ??= MatchPercent(PercentChangeRegex, allText);
        snapshot.MonthlyPremiumChange ??= snapshot.CurrentMonthlyPremium.HasValue && snapshot.RenewalMonthlyPremium.HasValue
            ? decimal.Round(snapshot.RenewalMonthlyPremium.Value - snapshot.CurrentMonthlyPremium.Value, 2)
            : null;
        snapshot.MonthlyPremiumChangePercent ??= UnderwritingEngine.CalculatePercentChange(
            snapshot.CurrentMonthlyPremium,
            snapshot.RenewalMonthlyPremium);

        return snapshot.HasUsefulValues ? snapshot : null;
    }

    public static PreviousRenewalSnapshot? FromCarrierExtraction(string pdfPath, PdfExtractionResult extraction)
    {
        var snapshot = new PreviousRenewalSnapshot
        {
            SourcePath = pdfPath,
            SourceFileName = Path.GetFileName(pdfPath),
            SourceType = string.IsNullOrWhiteSpace(extraction.Carrier) ? "Previous carrier renewal PDF" : $"Previous {extraction.Carrier} renewal PDF",
            Carrier = extraction.Carrier,
            ClientName = extraction.ClientName,
            PolicyNumber = extraction.PolicyNumber,
            RenewalDate = extraction.RenewalDate,
            CurrentMonthlyPremium = extraction.RateLines.Count > 0
                ? extraction.RateLines.Sum(line => line.CurrentPremium ?? 0m)
                : extraction.CurrentMonthlyPremium,
            RenewalMonthlyPremium = extraction.RateLines.Count > 0
                ? extraction.RateLines.Sum(line => line.RenewalPremium ?? 0m)
                : extraction.RenewalMonthlyPremium,
            MonthlyPremiumChange = extraction.MonthlyPremiumChange,
            MonthlyPremiumChangePercent = extraction.MonthlyPremiumChangePercent,
            SourceReference = $"{Path.GetFileName(pdfPath)}"
        };

        if (!snapshot.MonthlyPremiumChange.HasValue &&
            snapshot.CurrentMonthlyPremium.HasValue &&
            snapshot.RenewalMonthlyPremium.HasValue)
        {
            snapshot.MonthlyPremiumChange = decimal.Round(snapshot.RenewalMonthlyPremium.Value - snapshot.CurrentMonthlyPremium.Value, 2);
        }

        snapshot.MonthlyPremiumChangePercent ??= UnderwritingEngine.CalculatePercentChange(
            snapshot.CurrentMonthlyPremium,
            snapshot.RenewalMonthlyPremium);

        return snapshot.HasUsefulValues ? snapshot : null;
    }

    private static bool LooksLikeRenewalReview(string text, string pdfPath)
    {
        var fileName = Path.GetFileNameWithoutExtension(pdfPath);
        return text.Contains("Renewal Review", StringComparison.OrdinalIgnoreCase) ||
               text.Contains("Renewal Analysis", StringComparison.OrdinalIgnoreCase) ||
               text.Contains("Renewal Review Summary", StringComparison.OrdinalIgnoreCase) ||
               fileName.Contains("Renewal_Review", StringComparison.OrdinalIgnoreCase) ||
               fileName.Contains("Renewal Review", StringComparison.OrdinalIgnoreCase) ||
               fileName.Contains("Client_Renewal", StringComparison.OrdinalIgnoreCase);
    }

    private static string BuildSourceReference(string pdfPath, IReadOnlyList<PdfPageText> pages)
    {
        var page = pages.FirstOrDefault(item =>
            item.Text.Contains("Renewal Review", StringComparison.OrdinalIgnoreCase) ||
            item.Text.Contains("Renewal Analysis", StringComparison.OrdinalIgnoreCase));

        return page is null
            ? Path.GetFileName(pdfPath)
            : $"{Path.GetFileName(pdfPath)} page {page.PageNumber}";
    }

    private static string Match(Regex regex, string text)
    {
        var match = regex.Match(text);
        return match.Success ? ParsingHelpers.Clean(match.Groups["value"].Value) : "";
    }

    private static decimal? MatchMoney(Regex regex, string text)
    {
        var match = regex.Match(text);
        return match.Success ? ParsingHelpers.ParseDecimal(match.Groups["value"].Value) : null;
    }

    private static decimal? MatchPercent(Regex regex, string text)
    {
        var match = regex.Match(text);
        return match.Success ? ParsingHelpers.ParseDecimal(match.Groups["value"].Value) : null;
    }

    private static DateOnly? ParseDate(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        var cleaned = value.Replace(",", "", StringComparison.Ordinal).Replace(".", "", StringComparison.Ordinal);
        var formats = new[]
        {
            "MMM d yyyy",
            "MMMM d yyyy",
            "d-MMM-yyyy",
            "d-MMMM-yyyy",
            "d/MM/yyyy",
            "M/d/yyyy",
            "MM/dd/yyyy",
            "yyyy-M-d",
            "yyyy-MM-dd"
        };

        foreach (var format in formats)
        {
            if (DateOnly.TryParseExact(cleaned, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out var exactDate))
            {
                return exactDate;
            }
        }

        return DateOnly.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.None, out var date)
            ? date
            : null;
    }

    private static string FirstNonBlank(params string[] values)
    {
        return values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)) ?? "";
    }
}

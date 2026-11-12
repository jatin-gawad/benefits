using System.Globalization;
using System.Text.RegularExpressions;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

/// <summary>
/// Parses the unbranded RWAM-administered current/renewal rate comparison layout.
/// Classification is based on the table's content fingerprint, never the filename.
/// </summary>
public sealed class RwamRateSheetPdfParser
{
    public const string CarrierName = "RWAM";
    public const string SupportedStatus = RateLineValidationStatuses.Supported;
    public const string NeedsReviewStatus = RateLineValidationStatuses.NeedsReview;

    private static readonly Regex GroupClassRegex = new(
        @"^GROUP\s*#\s*(?<group>[A-Z0-9-]+)\s+(?<class>[A-Z0-9-]+)$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex RenewalDateRegex = new(
        @"^RENEWAL\s+EFFECTIVE:\s*(?<date>.+)$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex TotalRegex = new(
        @"^TOTAL\s+(?<current>[\d,]+\.\d{2})\s+(?<renewal>[\d,]+\.\d{2})$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex NumericRowRegex = new(
        @"^(?<label>.+?)\s+(?<quantity>[\d,]+)\s+(?<currentRate>\d+\.\d+)\s+(?<currentPremium>[\d,]+\.\d{2})\s+(?<renewalRate>\d+\.\d+)\s+(?<renewalPremium>[\d,]+\.\d{2})(?:\s+(?<carrier>[A-Za-z][A-Za-z &.-]*))?$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly string[] RequiredFingerprintAnchors =
    [
        "GROUP #",
        "RENEWAL EFFECTIVE:",
        "CURRENT RENEWAL",
        "BENEFIT VOLUME RATE PREMIUM RATE PREMIUM CARRIER",
        "STOP LOSS ($25,000)",
        "OUT OF CANADA",
        "TOTAL EHC",
        "Overall Adjustment to Group Premium",
        "Premiums are subject to applicable taxes."
    ];

    public static bool IsRwamRateSheet(string text)
    {
        var cleaned = ParsingHelpers.Clean(text);
        return RequiredFingerprintAnchors.All(anchor =>
            cleaned.Contains(anchor, StringComparison.OrdinalIgnoreCase));
    }

    public CarrierPdfParseResult Parse(string pdfPath, List<PdfPageText> pages, string allText)
    {
        if (!IsRwamRateSheet(allText))
        {
            throw new InvalidDataException("The PDF does not match the supported RWAM rate-sheet content fingerprint.");
        }

        var result = new PdfExtractionResult
        {
            SourcePath = pdfPath,
            Carrier = CarrierName
        };

        foreach (var page in pages)
        {
            ParsePage(page, result);
        }

        PopulateDocumentIdentity(pages, result);
        PopulateTotals(result);
        ValidateParsedDocument(result);

        return new CarrierPdfParseResult
        {
            Pdf = result,
            Carrier = CarrierName
        };
    }

    private static void ParsePage(PdfPageText page, PdfExtractionResult result)
    {
        var section = "";
        var sectionCarrier = "";
        foreach (var rawLine in page.Lines)
        {
            var line = ParsingHelpers.Clean(rawLine);
            if (string.IsNullOrWhiteSpace(line))
            {
                continue;
            }

            var totalMatch = TotalRegex.Match(line);
            if (totalMatch.Success)
            {
                result.CurrentMonthlyPremium = Decimal(totalMatch, "current");
                result.RenewalMonthlyPremium = Decimal(totalMatch, "renewal");
                continue;
            }

            if (TryMapSection(line, out var mappedSection))
            {
                section = mappedSection;
                sectionCarrier = "";
                continue;
            }

            var rowMatch = NumericRowRegex.Match(line);
            if (!rowMatch.Success)
            {
                continue;
            }

            var sourceLabel = ParsingHelpers.Clean(rowMatch.Groups["label"].Value);
            var tier = "";
            var benefit = "";
            if (sourceLabel.StartsWith("-", StringComparison.Ordinal))
            {
                tier = ParsingHelpers.Clean(sourceLabel.TrimStart('-'));
                benefit = section;
            }
            else
            {
                benefit = MapUntieredBenefit(sourceLabel);
            }

            if (string.IsNullOrWhiteSpace(benefit))
            {
                result.Warnings.Add($"RWAM page {page.PageNumber}: unsupported rate row '{sourceLabel}'.");
                continue;
            }

            var sourceCarrier = ParsingHelpers.Clean(rowMatch.Groups["carrier"].Value);
            if (!string.IsNullOrWhiteSpace(sourceCarrier))
            {
                sectionCarrier = sourceCarrier;
            }
            else if (!string.IsNullOrWhiteSpace(tier))
            {
                sourceCarrier = sectionCarrier;
            }

            result.RateLines.Add(BuildRateLine(
                result,
                page.PageNumber,
                line,
                benefit,
                tier,
                sourceCarrier,
                rowMatch));
        }
    }

    private static RateLine BuildRateLine(
        PdfExtractionResult document,
        int pageNumber,
        string sourceText,
        string benefit,
        string tier,
        string sourceCarrier,
        Match match)
    {
        var quantity = Decimal(match, "quantity") ?? 0m;
        var rateBasis = ResolveRateBasis(benefit);
        var isVolumeBased = UnderwritingEngine.RateVolumeDivisor(benefit).HasValue;
        var isStopLoss = benefit.Equals("Extended Health Care - Stop Loss", StringComparison.OrdinalIgnoreCase);
        var isComponent = benefit.Equals("Extended Health Care - ASO", StringComparison.OrdinalIgnoreCase) ||
                          benefit.Equals("Extended Health Care - Out of Canada", StringComparison.OrdinalIgnoreCase);
        var isSummary = benefit.Equals("Extended Health Care - Total", StringComparison.OrdinalIgnoreCase);
        var zeroEvidenceNeedsReview = quantity == 0m &&
                                      (benefit.Equals("WorldCare", StringComparison.OrdinalIgnoreCase) ||
                                       benefit.Equals("Employee Assistance Program", StringComparison.OrdinalIgnoreCase) ||
                                       benefit.Equals("Critical Illness", StringComparison.OrdinalIgnoreCase) ||
                                       benefit.Equals("Virtual Care", StringComparison.OrdinalIgnoreCase));
        var needsReview = isStopLoss || zeroEvidenceNeedsReview;

        var accountClass = ParseAccountClass(document.PolicyNumber);
        var sourceReference = $"RWAM rate sheet | {accountClass} | page {pageNumber}";
        return new RateLine
        {
            Benefit = benefit,
            TierOrClass = tier,
            AccountClass = accountClass,
            Volume = isVolumeBased ? quantity : null,
            Lives = isVolumeBased ? null : SafeInt(quantity),
            CurrentRate = Decimal(match, "currentRate"),
            CurrentPremium = Decimal(match, "currentPremium"),
            RenewalRate = Decimal(match, "renewalRate"),
            RenewalPremium = Decimal(match, "renewalPremium"),
            SourceReference = sourceReference,
            FundingArrangement = sourceCarrier.Equals("ASO", StringComparison.OrdinalIgnoreCase) ? "ASO" : "",
            RateBasis = isStopLoss ? "percent_of_net_paid_ehc_claims" : rateBasis,
            IncludeInPremiumTotal = !isStopLoss && !isComponent,
            SourceRowRole = isStopLoss ? "Conditional" : isSummary ? "Summary" : "Detail",
            ValidationStatus = needsReview ? NeedsReviewStatus : SupportedStatus,
            ValidationMessage = isStopLoss
                ? "The source states 2.5% of Net Paid EHC Claims, but the claim base is absent; printed 0.00 cells are preserved and are not treated as a per-life rate."
                : zeroEvidenceNeedsReview
                    ? "The source prints a zero quantity, zero rate, and zero premium, so the application basis cannot be independently confirmed from this sheet."
                    : "Printed current and renewal values map to an established application rate basis.",
            SourceCarrier = sourceCarrier,
            SourceFileName = SanitizedSourceName(accountClass),
            SourcePage = pageNumber,
            SourceText = sourceText,
            CurrentRatePrecision = PrintedPrecision(match.Groups["currentRate"].Value),
            RenewalRatePrecision = PrintedPrecision(match.Groups["renewalRate"].Value),
            CurrentPeriodLabel = "Current (effective date not stated)",
            RenewalPeriodLabel = document.RenewalDate?.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) ?? "Renewal",
            RenewalEffectiveDate = document.RenewalDate
        };
    }

    private static void PopulateDocumentIdentity(IReadOnlyList<PdfPageText> pages, PdfExtractionResult result)
    {
        foreach (var page in pages)
        {
            var lines = page.Lines.Select(ParsingHelpers.Clean).Where(line => line.Length > 0).ToList();
            for (var index = 0; index < lines.Count; index++)
            {
                var groupMatch = GroupClassRegex.Match(lines[index]);
                if (groupMatch.Success)
                {
                    var group = groupMatch.Groups["group"].Value;
                    var accountClass = groupMatch.Groups["class"].Value;
                    result.PolicyNumber = $"{group}|{accountClass}";
                    result.ClientName = index > 0 ? lines[index - 1] : "";
                    continue;
                }

                var dateMatch = RenewalDateRegex.Match(lines[index]);
                if (dateMatch.Success &&
                    DateOnly.TryParse(dateMatch.Groups["date"].Value, CultureInfo.InvariantCulture, DateTimeStyles.None, out var date))
                {
                    result.RenewalDate = date;
                }
            }
        }

        var groupNumber = ParseGroupNumber(result.PolicyNumber);
        var account = ParseAccountClass(result.PolicyNumber);
        result.PolicyNumber = groupNumber;
        foreach (var line in result.RateLines)
        {
            line.AccountClass = account;
            line.SourceReference = $"RWAM rate sheet | {account} | page {line.SourcePage ?? 1}";
            line.SourceFileName = SanitizedSourceName(account);
            line.RenewalPeriodLabel = result.RenewalDate?.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) ?? "Renewal";
            line.RenewalEffectiveDate = result.RenewalDate;
        }
    }

    private static void PopulateTotals(PdfExtractionResult result)
    {
        result.CurrentEhcPremium = SumPremium(result.RateLines, "Extended Health Care - Total", current: true);
        result.RenewalEhcPremium = SumPremium(result.RateLines, "Extended Health Care - Total", current: false);
        result.CurrentDentalPremium = SumPremium(result.RateLines, "Dental", current: true);
        result.RenewalDentalPremium = SumPremium(result.RateLines, "Dental", current: false);
        if (result.CurrentMonthlyPremium.HasValue && result.RenewalMonthlyPremium.HasValue)
        {
            result.MonthlyPremiumChange = result.RenewalMonthlyPremium.Value - result.CurrentMonthlyPremium.Value;
            result.MonthlyPremiumChangePercent = UnderwritingEngine.CalculatePercentChange(
                result.CurrentMonthlyPremium,
                result.RenewalMonthlyPremium);
        }
    }

    private static void ValidateParsedDocument(PdfExtractionResult result)
    {
        if (string.IsNullOrWhiteSpace(result.ClientName) || string.IsNullOrWhiteSpace(result.PolicyNumber))
        {
            result.Warnings.Add("RWAM client/group identity requires review.");
        }

        if (!result.RenewalDate.HasValue)
        {
            result.Warnings.Add("RWAM renewal effective date was not extracted.");
        }

        if (result.RateLines.Count != 21)
        {
            result.Warnings.Add($"RWAM rate-sheet layout requires review: expected 21 displayed rate rows but extracted {result.RateLines.Count}.");
        }

        if (!result.CurrentMonthlyPremium.HasValue || !result.RenewalMonthlyPremium.HasValue)
        {
            result.Warnings.Add("RWAM printed grand totals were not extracted.");
        }

        result.Warnings.Add("RWAM attribution uses the supplied rate-sheet content fingerprint; the issuer name is not printed on the page.");
    }

    private static bool TryMapSection(string line, out string section)
    {
        section = line.ToUpperInvariant() switch
        {
            "E.H.C." => "Extended Health Care - ASO",
            "OUT OF CANADA" => "Extended Health Care - Out of Canada",
            "TOTAL EHC" => "Extended Health Care - Total",
            "DENTAL" => "Dental",
            "CRITICAL ILLNESS" => "Critical Illness",
            "VIRTUAL CARE" => "Virtual Care",
            _ when line.StartsWith("STOP LOSS ($25,000)", StringComparison.OrdinalIgnoreCase) => "Extended Health Care - Stop Loss",
            _ => ""
        };
        return section.Length > 0;
    }

    private static string MapUntieredBenefit(string label)
    {
        return label.ToUpperInvariant() switch
        {
            "LIFE" => "Life",
            "A.D.&D." => "Accidental Death and Dismemberment",
            "DEP. LIFE" => "Dependent Life",
            "S.T.D." => "Short Term Disability",
            "L.T.D." => "Long Term Disability",
            "WORLDCARE" => "WorldCare",
            "EAP" => "Employee Assistance Program",
            _ => ""
        };
    }

    public static string ResolveRateBasis(string benefit)
    {
        return UnderwritingEngine.RateVolumeDivisor(benefit) switch
        {
            1000 => "per_1000_volume",
            100 => "per_100_volume",
            10 => "per_10_volume",
            _ => "per_life"
        };
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

    private static decimal? Decimal(Match match, string groupName)
    {
        return match.Success ? ParsingHelpers.ParseDecimal(match.Groups[groupName].Value) : null;
    }

    private static int SafeInt(decimal value)
    {
        return value is >= int.MinValue and <= int.MaxValue && value == decimal.Truncate(value)
            ? decimal.ToInt32(value)
            : throw new InvalidDataException($"RWAM lives/count value '{value}' is not a valid whole number.");
    }

    private static int PrintedPrecision(string value)
    {
        var separator = value.IndexOf('.');
        return separator < 0 ? 0 : value.Length - separator - 1;
    }

    private static string ParseGroupNumber(string compound)
    {
        return compound.Split('|', 2, StringSplitOptions.TrimEntries).FirstOrDefault() ?? "";
    }

    private static string ParseAccountClass(string compound)
    {
        var parts = compound.Split('|', 2, StringSplitOptions.TrimEntries);
        return parts.Length == 2 && parts[1].Length > 0 ? $"RWAM Class {parts[1]}" : "RWAM Class Needs Review";
    }

    private static string SanitizedSourceName(string accountClass)
    {
        var className = accountClass.Replace("RWAM Class ", "", StringComparison.OrdinalIgnoreCase);
        return $"[GROUP] {className} Rate Sheet.pdf";
    }
}

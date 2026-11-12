using System.Globalization;
using System.Text.RegularExpressions;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed class SunLifePdfParser
{
    private static readonly Regex ContractNumberRegex = new(@"(?:Contract\s+numbers?|Contract\s+No\.?|Re:\s*Contract)\s*:?\s*(?<value>\d+)", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex ContractHolderRegex = new(@"(?:Contract\s+holder|Policyholder\s+Name)\s*:?\s*(?<value>[A-Za-z0-9 &'.,\-]+)", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex RateEffectiveDateRegex = new(@"Rate\s+effective\s+date\s*:?\s*(?<date>[A-Za-z]{3,9}\s+\d{1,2},?\s+\d{4}|\d{1,2}-[A-Za-z]{3,9}-\d{2,4})", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex VolumeRateRowRegex = new(@"^\$?(?<volume>[\d,]+)\s+(?<lives>\d+)\s+\$?(?<currentRate>[\d.]+)/(?:\$)?(?<currentBasis>[\d,]+)\s+\$?(?<renewalRate>[\d.]+)/(?:\$)?(?<renewalBasis>[\d,]+)\s+\$?(?<currentPremium>[\d,]+\.\d{2})\s+\$?(?<renewalPremium>[\d,]+\.\d{2})$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex TierRateRowRegex = new(@"^(?<lives>\d+)\s+\$?(?<currentRate>[\d.]+)/(?<currentTier>single|couple|family)\s+\$?(?<renewalRate>[\d.]+)/(?<renewalTier>single|couple|family)\s+\$?(?<currentPremium>[\d,]+\.\d{2})\s+\$?(?<renewalPremium>[\d,]+\.\d{2})$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex TotalPremiumRegex = new(@"Total premium excluding sales tax\s+\$?(?<current>[\d,]+\.\d{2})\s+\$?(?<renewal>[\d,]+\.\d{2})", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex DetailedTotalPremiumRegex = new(@"^(?:Grand\s+)?Total Monthly Premium\s+\$?(?<current>[\d,]+\.\d{2})\s+\$?(?<renewal>[\d,]+\.\d{2})(?:\s+[+-]?[\d.]+%)?$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex DetailedPremiumRowRegex = new(
        @"^(?<prefix>.+?)\s+(?<volume>[\d,]+)\s+\$?(?<currentRate>[\d.]+)\s+\$?(?<currentPremium>[\d,]+\.\d{2})\s+\$?(?<renewalRate>[\d.]+)\s+\$?(?<renewalPremium>[\d,]+\.\d{2})\s+(?<change>[+-]?[\d.]+)%$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex DetailedRateBasisRegex = new(
        @"\bper\s+\$?\s*(?<basis>[\d,]+)\b",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex SunLifeDemographicRowRegex = new(
        @"^(?<age>Up\s+to\s+24|Under\s+25|<25|<20|0-19|20-24|25-29|30-34|35-39|40-44|45-49|50-54|55-59|60-64|65\s*\+|65-69|70\s*(?:\+|&\s*over|and\s+over))\s+" +
        @"(?<currentFemaleLives>-?\d+)\s+(?<currentFemaleVolume>-?[\d,]+)\s+" +
        @"(?<currentMaleLives>-?\d+)\s+(?<currentMaleVolume>-?[\d,]+)\s+" +
        @"(?<priorFemaleLives>-?\d+)\s+(?<priorFemaleVolume>-?[\d,]+)\s+" +
        @"(?<priorMaleLives>-?\d+)\s+(?<priorMaleVolume>-?[\d,]+)\s+" +
        @"(?<changeFemaleLives>-?\d+)\s+(?<changeFemaleVolume>-?[\d,]+)\s+" +
        @"(?<changeMaleLives>-?\d+)\s+(?<changeMaleVolume>-?[\d,]+)$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex ExperienceDateLineRegex = new(@"^(?<date>\d{2}-[A-Za-z]{3}-\d{4})$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex ExperienceDataRowRegex = new(@"^(?<average>\d+)\s+\$?(?<billed>[\d,]+)\s+\$?(?<adjusted>[\d,]+)\s+\$?(?<paid>[\d,]+)\s+\$?(?<incurred>[\d,]+)\s+(?<loss>[\d.]+)%\s+(?<weighting>[\d.]+)%$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex PoolingDataRowRegex = new(@"^\$?(?<largePremium>[\d,]+)\s+\$?(?<quebecPremium>[\d,]+)\s+\$?(?<pooledClaims>[\d,]+)\s+\$?(?<quebecPooledClaims>[\d,]+)$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex PoolingThresholdRegex = new(@"(?:claims\s+in\s+excess\s+of|Pooling\s+(?:Level|Threshold))\s+\$?\s*(?<value>[\d,\s]+)", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex CombinedFromLineRegex = new(@"^From\s+(?<prior>\d{1,2}-[A-Za-z]{3}-\d{4})\s+(?<current>\d{1,2}-[A-Za-z]{3}-\d{4})$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex CombinedToLineRegex = new(@"^To\s+(?<prior>\d{1,2}-[A-Za-z]{3}-\d{4})\s+(?<current>\d{1,2}-[A-Za-z]{3}-\d{4})$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex CombinedDemographicRowRegex = new(@"^(?<age>UP\s+TO\s+24|UNDER\s+25|<25|<20|0-19|20-24|25-29|30-34|35-39|40-44|45-49|50-54|55-59|60-64|65-69|70-74|75-79|80-84|85-89|90-94|95-99|65\+|70\+)\s+(?<values>.+)$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex CombinedDemographicValueRegex = new(@"\([\d,]+\)|-|[\d,]+", RegexOptions.Compiled);
    private static readonly Regex TopDinRowRegex = new(
        @"^(?<rank>\d{1,3})\s+(?<din>\d{6,9})\s+(?<detail>.+?)\s+(?<claims>\d+)\s+(?<claimsPercent>[\d.]+)%\s+\$?(?<paid>[\d,]+\.\d{2})\s+(?<paidPercent>[\d.]+)%$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex TopDinDetailRegex = new(@"^(?<name>.+?)\s+(?<type>Single Source|Brand Name|Generic|Unknown)\s+(?<classification>.+)$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex SplitMoneyTokenRegex = new(@"(?<![\d,])(?<prefix>\d{1,2})\s+(?<rest>\d{1,2},\d{3})(?![\d,])", RegexOptions.Compiled);
    private static readonly Regex SpacedCommaMoneyTokenRegex = new(@"(?<=\d)\s+,\s*(?=\d{3}(?:\D|$))", RegexOptions.Compiled);

    public static bool IsSunLifeRenewal(string text)
    {
        return text.Contains("Sun Life", StringComparison.OrdinalIgnoreCase) &&
               (text.Contains("Summary of rates and estimated monthly premium", StringComparison.OrdinalIgnoreCase) ||
                text.Contains("DETAILED PREMIUM ILLUSTRATION", StringComparison.OrdinalIgnoreCase) ||
                text.Contains("Experience analysis", StringComparison.OrdinalIgnoreCase) ||
                 text.Contains("Extended Health Care Renewal Factors", StringComparison.OrdinalIgnoreCase) ||
                 text.Contains("Dental Renewal Factors", StringComparison.OrdinalIgnoreCase) ||
                 text.Contains("Top DIN Listing", StringComparison.OrdinalIgnoreCase) ||
                 text.Contains("Claims Paid by Type of Service", StringComparison.OrdinalIgnoreCase));
    }

    public CarrierPdfParseResult Parse(string pdfPath, List<PdfPageText> pages, string allText)
    {
        var pdf = new PdfExtractionResult
        {
            SourcePath = pdfPath,
            Carrier = "Sun Life",
            PolicyNumber = MatchValue(ContractNumberRegex, allText),
            ClientName = ParseContractHolder(pages),
            RenewalDate = ParseRateEffectiveDate(allText)
        };

        ParseRateLines(pages, pdf);
        ParseExperienceLines(pages, pdf);
        ParseDemographics(pages, pdf);
        ParsePoolingThreshold(allText, pdf);
        ParseCombinedPoolingDetails(pages, pdf);
        ApplySunLifePoolingAssumptions(pdf);
        var workbookFromPdf = ParseTopDins(pages, pdf);

        if (pdf.RateLines.Count == 0)
        {
            pdf.Warnings.Add("No Sun Life renewal rate rows were extracted from this PDF.");
        }

        if (pdf.ExperienceLines.Count == 0)
        {
            pdf.Warnings.Add("No Sun Life experience-analysis rows were extracted from this PDF.");
        }

        return new CarrierPdfParseResult
        {
            Pdf = pdf,
            WorkbookFromPdf = workbookFromPdf,
            Carrier = "Sun Life"
        };
    }

    public static void ApplySunLifePoolingAssumptions(PdfExtractionResult result)
    {
        if (!result.Carrier.Equals("Sun Life", StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        if (result.PoolingCharge.HasValue)
        {
            return;
        }

        var currentEhcPoolingRow = result.SunLifeExperienceRows
            .Where(row => row.Benefit.Equals("Extended Health Care", StringComparison.OrdinalIgnoreCase) &&
                          row.AdjustedPremium.GetValueOrDefault() > 0m &&
                          (row.LargeAmountPoolingPremium.HasValue || row.QuebecIndustryPoolingPremium.HasValue))
            .OrderByDescending(row => ParseExperiencePeriodEnd(row.Period) ?? DateOnly.MinValue)
            .FirstOrDefault();
        if (currentEhcPoolingRow is null)
        {
            return;
        }

        var poolingPremium = currentEhcPoolingRow.LargeAmountPoolingPremium.GetValueOrDefault() +
                             currentEhcPoolingRow.QuebecIndustryPoolingPremium.GetValueOrDefault();
        if (poolingPremium < 0m)
        {
            return;
        }

        result.PoolingCharge = decimal.Round(poolingPremium / currentEhcPoolingRow.AdjustedPremium!.Value * 100m, 2);
    }

    private static string ParseContractHolder(IEnumerable<PdfPageText> pages)
    {
        foreach (var line in pages.SelectMany(page => page.Lines).Select(ParsingHelpers.Clean))
        {
            var match = ContractHolderRegex.Match(line);
            if (!match.Success)
            {
                continue;
            }

            var value = ParsingHelpers.Clean(match.Groups["value"].Value);
            if (!string.IsNullOrWhiteSpace(value))
            {
                return value;
            }
        }

        return "";
    }

    private static DateOnly? ParseRateEffectiveDate(string text)
    {
        var match = RateEffectiveDateRegex.Match(text);
        return match.Success && TryParseDate(match.Groups["date"].Value, out var date) ? date : null;
    }

    private static void ParseRateLines(IEnumerable<PdfPageText> pages, PdfExtractionResult result)
    {
        foreach (var page in pages.Where(ContainsRateSection))
        {
            var currentBenefit = "";
            var isDetailedIllustration =
                page.Text.Contains("DETAILED PREMIUM ILLUSTRATION", StringComparison.OrdinalIgnoreCase) ||
                (page.Text.Contains("Average rate / Total monthly premium", StringComparison.OrdinalIgnoreCase) &&
                 page.Text.Contains("Total Monthly Premium", StringComparison.OrdinalIgnoreCase));
            foreach (var line in page.Lines.Select(ParsingHelpers.Clean).Where(line => line.Length > 0))
            {
                var totalMatch = TotalPremiumRegex.Match(line);
                if (!totalMatch.Success)
                {
                    totalMatch = DetailedTotalPremiumRegex.Match(line);
                }

                if (totalMatch.Success)
                {
                    result.CurrentMonthlyPremium = Money(totalMatch, "current");
                    result.RenewalMonthlyPremium = Money(totalMatch, "renewal");
                    result.MonthlyPremiumChange = result.RenewalMonthlyPremium - result.CurrentMonthlyPremium;
                    result.MonthlyPremiumChangePercent = UnderwritingEngine.CalculatePercentChange(result.CurrentMonthlyPremium, result.RenewalMonthlyPremium);
                    continue;
                }

                if (TryMapRateBenefitHeading(line, out var benefit))
                {
                    currentBenefit = benefit;
                    continue;
                }

                var rateLine = TryParseRateLine(line, currentBenefit, page.PageNumber, isDetailedIllustration);
                if (rateLine is not null)
                {
                    result.RateLines.Add(rateLine);
                }
            }
        }

        result.CurrentEhcPremium = SumPremium(result.RateLines, "Extended Health Care", current: true);
        result.RenewalEhcPremium = SumPremium(result.RateLines, "Extended Health Care", current: false);
        result.CurrentDentalPremium = SumPremium(result.RateLines, "Dental", current: true);
        result.RenewalDentalPremium = SumPremium(result.RateLines, "Dental", current: false);

        if (result.RateLines.Count > 0)
        {
            result.CurrentMonthlyPremium ??= SumAllPremium(result.RateLines, current: true);
            result.RenewalMonthlyPremium ??= SumAllPremium(result.RateLines, current: false);
            result.MonthlyPremiumChange ??= result.RenewalMonthlyPremium - result.CurrentMonthlyPremium;
            result.MonthlyPremiumChangePercent ??= UnderwritingEngine.CalculatePercentChange(result.CurrentMonthlyPremium, result.RenewalMonthlyPremium);
        }
    }

    private static RateLine? TryParseRateLine(string line, string currentBenefit, int pageNumber, bool isDetailedIllustration)
    {
        if (string.IsNullOrWhiteSpace(currentBenefit))
        {
            return null;
        }

        var volumeMatch = VolumeRateRowRegex.Match(line);
        if (volumeMatch.Success)
        {
            var currentRateBasis = FormatRateBasis(volumeMatch.Groups["currentBasis"].Value);
            var renewalRateBasis = FormatRateBasis(volumeMatch.Groups["renewalBasis"].Value);
            var hasMismatchedRateBasis = !currentRateBasis.Equals(renewalRateBasis, StringComparison.OrdinalIgnoreCase);
            var rateLine = new RateLine
            {
                Benefit = currentBenefit,
                Volume = Money(volumeMatch, "volume"),
                Lives = Int(volumeMatch, "lives"),
                CurrentRate = Money(volumeMatch, "currentRate"),
                CurrentPremium = Money(volumeMatch, "currentPremium"),
                RenewalRate = Money(volumeMatch, "renewalRate"),
                RenewalPremium = Money(volumeMatch, "renewalPremium"),
                RateBasis = currentRateBasis,
                SourceReference = $"Sun Life PDF page {pageNumber}",
                ValidationStatus = hasMismatchedRateBasis ? RateLineValidationStatuses.NeedsReview : "",
                ValidationMessage = hasMismatchedRateBasis
                    ? $"Current rate basis '{currentRateBasis}' does not match renewal rate basis '{renewalRateBasis}'."
                    : ""
            };
            rateLine.PercentChange = UnderwritingEngine.CalculatePercentChange(rateLine.CurrentPremium, rateLine.RenewalPremium);
            return rateLine;
        }

        if (isDetailedIllustration)
        {
            var detailedMatch = DetailedPremiumRowRegex.Match(line);
            if (detailedMatch.Success &&
                (detailedMatch.Groups["prefix"].Value.Contains("Class", StringComparison.OrdinalIgnoreCase) ||
                 detailedMatch.Groups["prefix"].Value.Contains("guaranteed amount", StringComparison.OrdinalIgnoreCase)))
            {
                var prefix = detailedMatch.Groups["prefix"].Value;
                var detailedTierMatch = Regex.Match(prefix, @"\b(?<tier>Single|Couple|Family)\b", RegexOptions.IgnoreCase);
                var classMatch = Regex.Match(prefix, @"(?:Classes?\s+|,\s*)(?<class>[A-Z](?:\s*,\s*[A-Z])*)(?=\s*(?:-|$))", RegexOptions.IgnoreCase);
                var volume = Money(detailedMatch, "volume");
                var detailedLine = new RateLine
                {
                    Benefit = currentBenefit,
                    TierOrClass = detailedTierMatch.Success
                        ? CultureInfo.InvariantCulture.TextInfo.ToTitleCase(detailedTierMatch.Groups["tier"].Value.ToLowerInvariant())
                        : classMatch.Success
                            ? Regex.Replace(classMatch.Groups["class"].Value, @"\s+", "")
                            : "",
                    Volume = UsesLivesAsDetailedVolume(currentBenefit, detailedTierMatch.Success) ? null : volume,
                    Lives = UsesLivesAsDetailedVolume(currentBenefit, detailedTierMatch.Success) && volume.HasValue
                        ? decimal.ToInt32(volume.Value)
                        : null,
                    CurrentRate = Money(detailedMatch, "currentRate"),
                    CurrentPremium = Money(detailedMatch, "currentPremium"),
                    RenewalRate = Money(detailedMatch, "renewalRate"),
                    RenewalPremium = Money(detailedMatch, "renewalPremium"),
                    RateBasis = ParseDetailedRateBasis(prefix),
                    SourceReference = $"Sun Life PDF page {pageNumber}"
                };
                detailedLine.PercentChange = UnderwritingEngine.CalculatePercentChange(detailedLine.CurrentPremium, detailedLine.RenewalPremium);
                return detailedLine;
            }
        }

        var tierMatch = TierRateRowRegex.Match(line);
        if (!tierMatch.Success)
        {
            return null;
        }

        var tier = CultureInfo.InvariantCulture.TextInfo.ToTitleCase(tierMatch.Groups["currentTier"].Value.ToLowerInvariant());
        var tierRateLine = new RateLine
        {
            Benefit = currentBenefit,
            TierOrClass = tier,
            Lives = Int(tierMatch, "lives"),
            CurrentRate = Money(tierMatch, "currentRate"),
            CurrentPremium = Money(tierMatch, "currentPremium"),
            RenewalRate = Money(tierMatch, "renewalRate"),
            RenewalPremium = Money(tierMatch, "renewalPremium"),
            SourceReference = $"Sun Life PDF page {pageNumber}"
        };
        tierRateLine.PercentChange = UnderwritingEngine.CalculatePercentChange(tierRateLine.CurrentPremium, tierRateLine.RenewalPremium);
        return tierRateLine;
    }

    private static bool TryMapRateBenefitHeading(string line, out string benefit)
    {
        var heading = line.TrimEnd('*').Trim();
        benefit = heading switch
        {
            "Employee Life" => "Life",
            "Employee AD&D" => "Accidental Death and Dismemberment",
            "Dependent Life" => "Dependent Life",
            "Long-Term Disability" => "LTD",
            "Extended Health Care" => "Extended Health Care",
            "Dental Care" => "Dental",
            _ => ""
        };

        if (!string.IsNullOrWhiteSpace(benefit))
        {
            return true;
        }

        benefit = heading switch
        {
            var value when value.StartsWith("Employee Life", StringComparison.OrdinalIgnoreCase) => "Life",
            var value when value.StartsWith("Optional Spousal Life", StringComparison.OrdinalIgnoreCase) => "Optional Spousal Life",
            var value when value.StartsWith("Optional Life", StringComparison.OrdinalIgnoreCase) => "Optional Life",
            var value when value.StartsWith("Dependent Life", StringComparison.OrdinalIgnoreCase) => "Dependent Life",
            var value when value.StartsWith("Optional Accidental Death and Dismemberment", StringComparison.OrdinalIgnoreCase) => "Optional Accidental Death and Dismemberment",
            var value when value.StartsWith("Accidental Death and Dismemberment", StringComparison.OrdinalIgnoreCase) => "Accidental Death and Dismemberment",
            var value when value.StartsWith("Long Term Disability", StringComparison.OrdinalIgnoreCase) => "LTD",
            var value when value.StartsWith("Extended Health Care", StringComparison.OrdinalIgnoreCase) => "Extended Health Care",
            var value when value.StartsWith("Dental", StringComparison.OrdinalIgnoreCase) => "Dental",
            var value when value.StartsWith("Critical Illness - Basic", StringComparison.OrdinalIgnoreCase) => "Critical Illness",
            var value when value.StartsWith("Optional Member Critical Illness", StringComparison.OrdinalIgnoreCase) => "Optional Member Critical Illness",
            var value when value.StartsWith("Optional Spousal Critical Illness", StringComparison.OrdinalIgnoreCase) => "Optional Spousal Critical Illness",
            var value when value.StartsWith("Optional Child Critical Illness", StringComparison.OrdinalIgnoreCase) => "Optional Child Critical Illness",
            _ => ""
        };

        return !string.IsNullOrWhiteSpace(benefit);
    }

    private static bool UsesLivesAsDetailedVolume(string benefit, bool hasTier)
    {
        return hasTier || benefit.Equals("Dependent Life", StringComparison.OrdinalIgnoreCase);
    }

    private static void ParseExperienceLines(IEnumerable<PdfPageText> pages, PdfExtractionResult result)
    {
        var pageList = pages.ToList();
        foreach (var page in pageList.Where(ContainsExperienceSection))
        {
            var benefit = page.Text.Contains("Extended Health Care", StringComparison.OrdinalIgnoreCase)
                ? "Extended Health Care"
                : page.Text.Contains("Dental", StringComparison.OrdinalIgnoreCase)
                    ? "Dental"
                    : "";
            if (string.IsNullOrWhiteSpace(benefit))
            {
                continue;
            }

            var factors = ParseRenewalFactors(page.Text);
            var rawRows = ParseSunLifeExperienceRows(page, benefit, factors);
            foreach (var raw in rawRows)
            {
                result.SunLifeExperienceRows.Add(raw);
                result.ExperienceLines.Add(new ExperienceLine
                {
                    Benefit = raw.Benefit,
                    Period = raw.Period,
                    AverageLives = raw.AverageCoveredEmployees,
                    BilledPremium = raw.BilledPremium,
                    PaidClaims = raw.PaidClaims,
                    IncurredClaims = raw.IncurredClaims,
                    PaidLossRatio = CalculateLossRatio(raw.PaidClaims, raw.BilledPremium),
                    IncurredLossRatio = raw.AdjustedLossRatio,
                    Weighting = raw.Weighting,
                    ProposedAdjustment = raw.RateAdjustment,
                    TrendFactor = raw.AnnualTrendFactor,
                    Credibility = raw.Credibility,
                    TargetLossRatio = raw.TargetLossRatio,
                    SourceReference = raw.SourceReference
                });
            }
        }

        ParseCombinedExperienceLines(pageList, result);
    }

    private static void ParseCombinedExperienceLines(IReadOnlyList<PdfPageText> pages, PdfExtractionResult result)
    {
        var allText = string.Join(Environment.NewLine, pages.Select(page => page.Text));
        foreach (var page in pages.Where(page =>
                     page.Text.Contains("experience analysis", StringComparison.OrdinalIgnoreCase) &&
                     page.Text.Contains("A) Paid Premiums", StringComparison.OrdinalIgnoreCase) &&
                     page.Text.Contains("B) Adjusted Premiums", StringComparison.OrdinalIgnoreCase)))
        {
            var benefit = page.Text.Contains("Extended health care", StringComparison.OrdinalIgnoreCase)
                ? "Extended Health Care"
                : page.Text.Contains("Dental", StringComparison.OrdinalIgnoreCase)
                    ? "Dental"
                    : "";
            if (string.IsNullOrWhiteSpace(benefit))
            {
                continue;
            }

            var lines = page.Lines.Select(ParsingHelpers.Clean).Where(line => line.Length > 0).ToList();
            var fromMatch = lines.Select(line => CombinedFromLineRegex.Match(line)).FirstOrDefault(match => match.Success);
            var toMatch = lines.Select(line => CombinedToLineRegex.Match(line)).FirstOrDefault(match => match.Success);
            if (fromMatch is null || toMatch is null || !fromMatch.Success || !toMatch.Success)
            {
                continue;
            }

            var incurredClaimsPattern = benefit.Equals("Dental", StringComparison.OrdinalIgnoreCase)
                ? @"E\)\s*Incurred Claims"
                : @"C\)\s*(?:Non Pooled )?Incurred Claims";
            if (!TryMatchPair(page.Text, @"Average Number of Insureds", out var priorAverage, out var currentAverage) ||
                !TryMatchPair(page.Text, @"A\)\s*Paid Premiums", out var priorBilled, out var currentBilled) ||
                !TryMatchPair(page.Text, @"B\)\s*Adjusted Premiums(?:\s*\([^)]*\))?", out var priorAdjusted, out var currentAdjusted) ||
                !TryMatchPair(page.Text, @"Paid Claims(?:\s*\(Net of Cost Plus\))?", out var priorPaid, out var currentPaid) ||
                !TryMatchPair(page.Text, incurredClaimsPattern, out var priorIncurred, out var currentIncurred))
            {
                continue;
            }

            TryMatchPair(page.Text, @"Pooling Premiums", out var priorPoolingPremium, out var currentPoolingPremium);
            TryMatchPair(page.Text, @"F\)\s*Loss Ratio[^\r\n]*", out var priorAdjustedLossRatio, out var currentAdjustedLossRatio, percentages: true);
            TryMatchPair(page.Text, @"G\)\s*Weighting", out var priorWeighting, out var currentWeighting, percentages: true);

            var factors = ParseCombinedRenewalFactors(page.Text, allText, benefit);
            var priorPeriod = FormatPeriod(fromMatch.Groups["prior"].Value, toMatch.Groups["prior"].Value);
            var currentPeriod = FormatPeriod(fromMatch.Groups["current"].Value, toMatch.Groups["current"].Value);
            AddCombinedExperienceRow(result, benefit, priorPeriod, priorAverage, priorBilled, priorAdjusted, priorPaid, priorIncurred,
                priorAdjustedLossRatio, priorWeighting, priorPoolingPremium, factors, page.PageNumber);
            AddCombinedExperienceRow(result, benefit, currentPeriod, currentAverage, currentBilled, currentAdjusted, currentPaid, currentIncurred,
                currentAdjustedLossRatio, currentWeighting, currentPoolingPremium, factors, page.PageNumber);

            if (benefit.Equals("Extended Health Care", StringComparison.OrdinalIgnoreCase) && currentPoolingPremium.HasValue)
            {
                result.PooledPremium = currentPoolingPremium;
            }
        }
    }

    private static void AddCombinedExperienceRow(
        PdfExtractionResult result,
        string benefit,
        string period,
        decimal? averageLives,
        decimal? billedPremium,
        decimal? adjustedPremium,
        decimal? paidClaims,
        decimal? incurredClaims,
        decimal? adjustedLossRatio,
        decimal? weighting,
        decimal? poolingPremium,
        SunLifeRenewalFactors factors,
        int pageNumber)
    {
        if (result.SunLifeExperienceRows.Any(row =>
                row.Benefit.Equals(benefit, StringComparison.OrdinalIgnoreCase) &&
                row.Period.Equals(period, StringComparison.OrdinalIgnoreCase)))
        {
            return;
        }

        var raw = new SunLifeExperienceRow
        {
            Benefit = benefit,
            Period = period,
            AverageCoveredEmployees = averageLives.HasValue ? decimal.ToInt32(averageLives.Value) : null,
            BilledPremium = billedPremium,
            AdjustedPremium = adjustedPremium,
            PaidClaims = paidClaims,
            IncurredClaims = incurredClaims,
            AdjustedLossRatio = adjustedLossRatio,
            Weighting = weighting,
            LargeAmountPoolingPremium = poolingPremium,
            WeightedAdjustedLossRatio = factors.WeightedAdjustedLossRatio,
            Credibility = factors.Credibility,
            Ibnr = factors.Ibnr,
            TargetLossRatio = factors.TargetLossRatio,
            AnnualTrendFactor = factors.AnnualTrendFactor,
            RateAdjustment = factors.RateAdjustment,
            SourceReference = $"Sun Life PDF page {pageNumber}"
        };
        result.SunLifeExperienceRows.Add(raw);
        result.ExperienceLines.Add(new ExperienceLine
        {
            Benefit = benefit,
            Period = period,
            AverageLives = raw.AverageCoveredEmployees,
            BilledPremium = raw.BilledPremium,
            PaidClaims = raw.PaidClaims,
            IncurredClaims = raw.IncurredClaims,
            PaidLossRatio = CalculateLossRatio(raw.PaidClaims, raw.BilledPremium),
            IncurredLossRatio = raw.AdjustedLossRatio,
            Weighting = raw.Weighting,
            ProposedAdjustment = raw.RateAdjustment,
            TrendFactor = raw.AnnualTrendFactor,
            Credibility = raw.Credibility,
            TargetLossRatio = raw.TargetLossRatio,
            SourceReference = raw.SourceReference
        });
    }

    private static SunLifeRenewalFactors ParseCombinedRenewalFactors(string pageText, string allText, string benefit)
    {
        var local = ParseRenewalFactors(pageText);
        var benefitToken = benefit.Equals("Extended Health Care", StringComparison.OrdinalIgnoreCase) ? "EHC" : "Dental";
        return new SunLifeRenewalFactors(
            local.WeightedAdjustedLossRatio ?? MatchPercent(pageText, @"Weighted (?:Adjusted )?Loss Ratio\s+(?<value>[\d.]+)%"),
            local.Credibility ?? MatchPercent(pageText, @"Credibility\s+(?<value>[\d.]+)%"),
            local.Ibnr ?? MatchPercent(allText, $@"Incurred but\s+{benefitToken}\s+(?<value>[\d.]+)%"),
            local.TargetLossRatio ?? MatchPercent(allText, $@"Target Loss\s+{benefitToken}\s+(?<value>[\d.]+)%"),
            local.AnnualTrendFactor ?? MatchPercent(allText, $@"Trend Factors?\s+{benefitToken}\s+(?<value>[\d.]+)%"),
            MatchPercent(pageText, @"Final Rate Adjustment\s+(?<value>[+-]?[\d.]+)%") ?? local.RateAdjustment);
    }

    private static bool TryMatchPair(string text, string labelPattern, out decimal? prior, out decimal? current, bool percentages = false)
    {
        var suffix = percentages ? "%" : "";
        var pattern = $@"{labelPattern}\s+\$?\(?(?<prior>[\d,]+(?:\.\d+)?)\)?{suffix}\s+\$?\(?(?<current>[\d,]+(?:\.\d+)?)\)?{suffix}";
        var match = Regex.Match(text, pattern, RegexOptions.IgnoreCase);
        prior = match.Success ? ParsingHelpers.ParseDecimal(match.Groups["prior"].Value) : null;
        current = match.Success ? ParsingHelpers.ParseDecimal(match.Groups["current"].Value) : null;
        return match.Success;
    }

    private static string FormatPeriod(string fromText, string toText)
    {
        return TryParseDate(fromText, out var from) && TryParseDate(toText, out var to)
            ? $"{from:dd-MMM-yyyy} - {to:dd-MMM-yyyy}"
            : $"{fromText} - {toText}";
    }

    private static decimal? CalculateLossRatio(decimal? claims, decimal? premium)
    {
        return claims.HasValue && premium is > 0m
            ? decimal.Round(claims.Value / premium.Value * 100m, 2)
            : null;
    }

    private static void ParsePoolingThreshold(string text, PdfExtractionResult result)
    {
        var match = PoolingThresholdRegex.Match(text);
        if (match.Success)
        {
            result.PoolingThreshold = Money(match, "value");
        }
    }

    private static void ParseCombinedPoolingDetails(IEnumerable<PdfPageText> pages, PdfExtractionResult result)
    {
        foreach (var page in pages.Where(page => page.Text.Contains("Pooling details", StringComparison.OrdinalIgnoreCase)))
        {
            var lines = page.Lines.Select(ParsingHelpers.Clean).Where(line => line.Length > 0).ToList();
            if (!result.PoolingThreshold.HasValue)
            {
                var levelLine = lines.FirstOrDefault(line =>
                    Regex.IsMatch(line, @"^\$?[\d,]+\s+(?:Individual|Family|Certificate|Per)", RegexOptions.IgnoreCase));
                var levelMatch = Regex.Match(levelLine ?? "", @"^\$?(?<value>[\d,]+)", RegexOptions.IgnoreCase);
                if (levelMatch.Success)
                {
                    result.PoolingThreshold = Money(levelMatch, "value");
                }
            }

            if (!result.PoolingCharge.HasValue)
            {
                var percentageMatches = lines
                    .Select(line => Regex.Match(line, @"(?<value>[\d.]+)%\s*$"))
                    .Where(match => match.Success)
                    .ToList();
                if (percentageMatches.Count > 0)
                {
                    result.PoolingCharge = Money(percentageMatches[^1], "value");
                }
            }
        }
    }

    private static WorkbookExtractionResult? ParseTopDins(IEnumerable<PdfPageText> pages, PdfExtractionResult pdf)
    {
        var pageList = pages.ToList();
        var workbook = new WorkbookExtractionResult
        {
            SourcePath = pdf.SourcePath
        };

        var claimsRankedRows = new Dictionary<string, TopDinParsedRow>(StringComparer.OrdinalIgnoreCase);
        foreach (var page in pageList.Where(page =>
                     page.Text.Contains("Top DIN Listing", StringComparison.OrdinalIgnoreCase) &&
                     page.Text.Contains("by Number of Claims", StringComparison.OrdinalIgnoreCase)))
        {
            if (!TopDinPageMatchesPolicy(page, pdf, workbook))
            {
                continue;
            }

            foreach (var line in page.Lines.Select(ParsingHelpers.Clean).Where(line => line.Length > 0))
            {
                var row = ParseTopDinRow(line, page.PageNumber);
                if (row is not null && !string.IsNullOrWhiteSpace(row.DrugName))
                {
                    claimsRankedRows[row.Din] = row;
                }
            }
        }

        foreach (var page in pageList.Where(page =>
                     page.Text.Contains("Top DIN Listing", StringComparison.OrdinalIgnoreCase) &&
                     page.Text.Contains("by Amount Paid", StringComparison.OrdinalIgnoreCase)))
        {
            if (!TopDinPageMatchesPolicy(page, pdf, workbook))
            {
                continue;
            }

            var lines = page.Lines.Select(ParsingHelpers.Clean).Where(line => line.Length > 0).ToList();
            var period = lines.FirstOrDefault(line => Regex.IsMatch(line,
                @"^[A-Za-z]+\s+\d{1,2},\s+\d{4}\s+to\s+[A-Za-z]+\s+\d{1,2},\s+\d{4}$",
                RegexOptions.IgnoreCase)) ?? "";

            foreach (var line in lines)
            {
                var row = ParseTopDinRow(line, page.PageNumber);
                if (row is null || row.Rank is < 1 or > 10)
                {
                    continue;
                }

                claimsRankedRows.TryGetValue(row.Din, out var claimsRankedRow);
                var drugName = !string.IsNullOrWhiteSpace(row.DrugName)
                    ? row.DrugName
                    : claimsRankedRow?.DrugName ?? "";
                var classification = !string.IsNullOrWhiteSpace(row.Classification)
                    ? row.Classification
                    : claimsRankedRow?.Classification ?? "";
                if (string.IsNullOrWhiteSpace(drugName))
                {
                    continue;
                }

                workbook.DrugTrends.Add(new DrugTrend
                {
                    Din = row.Din,
                    DrugName = drugName,
                    Treats = classification,
                    Period = period,
                    PrescriptionCount = row.Claims,
                    NetAmountPaid = row.AmountPaid,
                    SourceReference = claimsRankedRow is null
                        ? $"Sun Life PDF page {page.PageNumber}"
                        : $"Sun Life PDF page {page.PageNumber}; classification cross-check page {claimsRankedRow.PageNumber}"
                });
            }
        }

        var distinct = workbook.DrugTrends
            .GroupBy(row => $"{row.Din}|{row.Period}", StringComparer.OrdinalIgnoreCase)
            .Select(group => group.First())
            .ToList();
        workbook.DrugTrends.Clear();
        workbook.DrugTrends.AddRange(distinct);
        return workbook.DrugTrends.Count > 0 || workbook.Warnings.Count > 0 ? workbook : null;
    }

    private static bool TopDinPageMatchesPolicy(
        PdfPageText page,
        PdfExtractionResult pdf,
        WorkbookExtractionResult workbook)
    {
        var pagePolicy = MatchValue(ContractNumberRegex, page.Text);
        if (string.IsNullOrWhiteSpace(pdf.PolicyNumber) ||
            string.IsNullOrWhiteSpace(pagePolicy) ||
            pagePolicy.Equals(pdf.PolicyNumber, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        workbook.Warnings.Add($"Sun Life Top DIN page {page.PageNumber} was skipped because contract {pagePolicy} does not match {pdf.PolicyNumber}.");
        return false;
    }

    private static TopDinParsedRow? ParseTopDinRow(string line, int pageNumber)
    {
        var rowMatch = TopDinRowRegex.Match(line);
        if (!rowMatch.Success || !int.TryParse(rowMatch.Groups["rank"].Value, out var rank))
        {
            return null;
        }

        var detailMatch = TopDinDetailRegex.Match(rowMatch.Groups["detail"].Value);
        return new TopDinParsedRow(
            rank,
            rowMatch.Groups["din"].Value,
            detailMatch.Success ? ParsingHelpers.Clean(detailMatch.Groups["name"].Value) : "",
            detailMatch.Success ? ParsingHelpers.Clean(detailMatch.Groups["classification"].Value) : "",
            Int(rowMatch, "claims"),
            Money(rowMatch, "paid"),
            pageNumber);
    }

    private sealed record TopDinParsedRow(
        int Rank,
        string Din,
        string DrugName,
        string Classification,
        int? Claims,
        decimal? AmountPaid,
        int PageNumber);

    private static List<SunLifeExperienceRow> ParseSunLifeExperienceRows(
        PdfPageText page,
        string benefit,
        SunLifeRenewalFactors factors)
    {
        var rows = new List<SunLifeExperienceRow>();
        var lines = page.Lines.Select(ParsingHelpers.Clean).Where(line => line.Length > 0).ToList();
        var poolingByPeriod = ParsePoolingRows(lines);

        for (var index = 0; index < lines.Count - 2; index++)
        {
            var fromMatch = ExperienceDateLineRegex.Match(lines[index]);
            if (!fromMatch.Success)
            {
                continue;
            }

            var valuesMatch = ExperienceDataRowRegex.Match(lines[index + 1]);
            var toMatch = ExperienceDateLineRegex.Match(lines[index + 2]);
            if (!valuesMatch.Success || !toMatch.Success)
            {
                continue;
            }

            var period = $"{fromMatch.Groups["date"].Value} - {toMatch.Groups["date"].Value}";
            poolingByPeriod.TryGetValue(period, out var pooling);

            rows.Add(new SunLifeExperienceRow
            {
                Benefit = benefit,
                Period = period,
                AverageCoveredEmployees = Int(valuesMatch, "average"),
                BilledPremium = Money(valuesMatch, "billed"),
                AdjustedPremium = Money(valuesMatch, "adjusted"),
                PaidClaims = Money(valuesMatch, "paid"),
                IncurredClaims = Money(valuesMatch, "incurred"),
                AdjustedLossRatio = Money(valuesMatch, "loss"),
                Weighting = Money(valuesMatch, "weighting"),
                LargeAmountPoolingPremium = pooling?.LargeAmountPoolingPremium,
                QuebecIndustryPoolingPremium = pooling?.QuebecIndustryPoolingPremium,
                PooledClaims = pooling?.PooledClaims,
                QuebecIndustryPooledClaims = pooling?.QuebecIndustryPooledClaims,
                WeightedAdjustedLossRatio = factors.WeightedAdjustedLossRatio,
                Credibility = factors.Credibility,
                Ibnr = factors.Ibnr,
                TargetLossRatio = factors.TargetLossRatio,
                AnnualTrendFactor = factors.AnnualTrendFactor,
                RateAdjustment = factors.RateAdjustment,
                SourceReference = $"Sun Life EXP PDF page {page.PageNumber}"
            });
        }

        return rows
            .GroupBy(row => $"{row.Benefit}|{row.Period}", StringComparer.OrdinalIgnoreCase)
            .Select(group => group.First())
            .ToList();
    }

    private static Dictionary<string, SunLifePoolingValues> ParsePoolingRows(List<string> lines)
    {
        var rows = new Dictionary<string, SunLifePoolingValues>(StringComparer.OrdinalIgnoreCase);
        var poolingHeaderIndex = lines.FindIndex(line => line.Contains("Large Quebec Large Quebec", StringComparison.OrdinalIgnoreCase));
        if (poolingHeaderIndex < 0)
        {
            return rows;
        }

        for (var index = poolingHeaderIndex + 1; index < lines.Count - 2; index++)
        {
            var fromMatch = ExperienceDateLineRegex.Match(lines[index]);
            if (!fromMatch.Success)
            {
                continue;
            }

            var valuesMatch = PoolingDataRowRegex.Match(lines[index + 1]);
            var toMatch = ExperienceDateLineRegex.Match(lines[index + 2]);
            if (!valuesMatch.Success || !toMatch.Success)
            {
                continue;
            }

            var period = $"{fromMatch.Groups["date"].Value} - {toMatch.Groups["date"].Value}";
            rows[period] = new SunLifePoolingValues(
                Money(valuesMatch, "largePremium"),
                Money(valuesMatch, "quebecPremium"),
                Money(valuesMatch, "pooledClaims"),
                Money(valuesMatch, "quebecPooledClaims"));
        }

        return rows;
    }

    private static SunLifeRenewalFactors ParseRenewalFactors(string text)
    {
        return new SunLifeRenewalFactors(
            MatchPercent(text, @"Weighted Adjusted Loss Ratio\s+(?<value>[\d.]+)%"),
            MatchPercent(text, @"Credibility\s+(?<value>[\d.]+)%"),
            MatchPercent(text, @"IBNR\s*\(% of premium\)\s+(?<value>[\d.]+)%"),
            MatchPercent(text, @"Target Loss Ratio\s+(?<value>[\d.]+)%"),
            MatchPercent(text, @"Annual Trend Factor\s+(?<value>[\d.]+)%"),
            MatchPercent(text, @"Rate Adjustment\s+(?<value>[+-]?[\d.]+)%"));
    }

    private static void ParseDemographics(IEnumerable<PdfPageText> pages, PdfExtractionResult result)
    {
        foreach (var page in pages.Where(ContainsDemographicSection))
        {
            if (page.Text.Contains("CURRENT", StringComparison.OrdinalIgnoreCase) &&
                (page.Text.Contains("PRIOR", StringComparison.OrdinalIgnoreCase) ||
                 page.Text.Contains("MARKETING", StringComparison.OrdinalIgnoreCase)) &&
                page.Text.Contains("VARIATIONS", StringComparison.OrdinalIgnoreCase) &&
                page.Text.Contains("BASIC LIFE", StringComparison.OrdinalIgnoreCase))
            {
                ParseCombinedDemographicPage(page, result);
                continue;
            }

            if (page.Text.Contains("Employee Life", StringComparison.OrdinalIgnoreCase))
            {
                ParseDemographicPage(page, "Life", result.LifeDemographics, DemographicAgeBandNormalizer.NormalizeLifeAgeBand);
            }

            if (page.Text.Contains("Long-Term Disability", StringComparison.OrdinalIgnoreCase) ||
                page.Text.Contains("Long Term Disability", StringComparison.OrdinalIgnoreCase))
            {
                ParseDemographicPage(page, "LTD", result.LtdDemographics, DemographicAgeBandNormalizer.NormalizeLtdAgeBand);
            }
        }
    }

    private static bool ContainsRateSection(PdfPageText page)
    {
        return page.Text.Contains("Summary of rates and estimated monthly premium", StringComparison.OrdinalIgnoreCase) ||
               page.Text.Contains("DETAILED PREMIUM ILLUSTRATION", StringComparison.OrdinalIgnoreCase) ||
               (page.Text.Contains("Average rate / Total monthly premium", StringComparison.OrdinalIgnoreCase) &&
                page.Text.Contains("Total Monthly Premium", StringComparison.OrdinalIgnoreCase)) ||
               page.Text.Contains("Coverage and estimated monthly premium by employee", StringComparison.OrdinalIgnoreCase) ||
               page.Text.Contains("Total premium excluding sales tax", StringComparison.OrdinalIgnoreCase);
    }

    private static bool ContainsExperienceSection(PdfPageText page)
    {
        return page.Text.Contains("Experience analysis", StringComparison.OrdinalIgnoreCase) ||
               page.Text.Contains("Extended Health Care Renewal Factors", StringComparison.OrdinalIgnoreCase) ||
               page.Text.Contains("Dental Renewal Factors", StringComparison.OrdinalIgnoreCase) ||
               page.Text.Contains("Weighted Adjusted Loss Ratio", StringComparison.OrdinalIgnoreCase) ||
               page.Text.Contains("Adjusted Loss Ratio", StringComparison.OrdinalIgnoreCase);
    }

    private static bool ContainsDemographicSection(PdfPageText page)
    {
        return page.Text.Contains("Demographic distribution", StringComparison.OrdinalIgnoreCase) ||
               (page.Text.Contains("Age Band", StringComparison.OrdinalIgnoreCase) &&
                (page.Text.Contains("Employee Life", StringComparison.OrdinalIgnoreCase) ||
                 page.Text.Contains("Long-Term Disability", StringComparison.OrdinalIgnoreCase) ||
                 page.Text.Contains("Long Term Disability", StringComparison.OrdinalIgnoreCase)));
    }

    private static void ParseCombinedDemographicPage(PdfPageText page, PdfExtractionResult result)
    {
        var benefit = "";
        foreach (var line in page.Lines.Select(ParsingHelpers.Clean).Where(line => line.Length > 0))
        {
            if (line.Equals("LONG TERM DISABILITY", StringComparison.OrdinalIgnoreCase))
            {
                benefit = "LTD";
                continue;
            }

            if (line.Equals("BASIC LIFE", StringComparison.OrdinalIgnoreCase))
            {
                benefit = "Life";
                continue;
            }

            if (string.IsNullOrWhiteSpace(benefit) || line.StartsWith("TOTAL", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var rowMatch = CombinedDemographicRowRegex.Match(line);
            if (!rowMatch.Success)
            {
                continue;
            }

            var values = CombinedDemographicValueRegex.Matches(rowMatch.Groups["values"].Value)
                .Select(match => match.Value)
                .Take(8)
                .ToList();
            if (values.Count < 8)
            {
                continue;
            }

            var ageBand = benefit.Equals("Life", StringComparison.OrdinalIgnoreCase)
                ? DemographicAgeBandNormalizer.NormalizeLifeAgeBand(rowMatch.Groups["age"].Value.Replace("UP TO 24", "Under 25", StringComparison.OrdinalIgnoreCase))
                : DemographicAgeBandNormalizer.NormalizeLtdAgeBand(rowMatch.Groups["age"].Value.Replace("UP TO 24", "Under 25", StringComparison.OrdinalIgnoreCase));
            var target = benefit.Equals("Life", StringComparison.OrdinalIgnoreCase)
                ? result.LifeDemographics
                : result.LtdDemographics;
            target.Add(new DemographicLine
            {
                Benefit = benefit,
                AgeBand = ageBand,
                CurrentMaleLives = ParseDemographicInt(values[0]),
                CurrentMaleVolume = ParseDemographicDecimal(values[1]),
                CurrentFemaleLives = ParseDemographicInt(values[2]),
                CurrentFemaleVolume = ParseDemographicDecimal(values[3]),
                PriorMaleLives = ParseDemographicInt(values[4]),
                PriorMaleVolume = ParseDemographicDecimal(values[5]),
                PriorFemaleLives = ParseDemographicInt(values[6]),
                PriorFemaleVolume = ParseDemographicDecimal(values[7]),
                SourceReference = $"Sun Life PDF page {page.PageNumber}"
            });
        }
    }

    private static int? ParseDemographicInt(string value)
    {
        if (value == "-")
        {
            return 0;
        }

        var negative = value.StartsWith('(') && value.EndsWith(')');
        var parsed = ParsingHelpers.ParseInt(value.Trim('(', ')'));
        return negative && parsed.HasValue ? -parsed.Value : parsed;
    }

    private static decimal? ParseDemographicDecimal(string value)
    {
        if (value == "-")
        {
            return 0m;
        }

        var negative = value.StartsWith('(') && value.EndsWith(')');
        var parsed = ParsingHelpers.ParseDecimal(value.Trim('(', ')'));
        return negative && parsed.HasValue ? -parsed.Value : parsed;
    }

    private static void ParseDemographicPage(
        PdfPageText page,
        string benefit,
        List<DemographicLine> target,
        Func<string, string> normalizeAgeBand)
    {
        foreach (var line in page.Lines.Select(ParsingHelpers.Clean))
        {
            var match = SunLifeDemographicRowRegex.Match(line);
            if (!match.Success)
            {
                continue;
            }

            target.Add(new DemographicLine
            {
                Benefit = benefit,
                AgeBand = normalizeAgeBand(match.Groups["age"].Value.Replace("Up to 24", "Under 25", StringComparison.OrdinalIgnoreCase)),
                CurrentFemaleLives = Int(match, "currentFemaleLives"),
                CurrentFemaleVolume = Money(match, "currentFemaleVolume"),
                CurrentMaleLives = Int(match, "currentMaleLives"),
                CurrentMaleVolume = Money(match, "currentMaleVolume"),
                PriorFemaleLives = Int(match, "priorFemaleLives"),
                PriorFemaleVolume = Money(match, "priorFemaleVolume"),
                PriorMaleLives = Int(match, "priorMaleLives"),
                PriorMaleVolume = Money(match, "priorMaleVolume"),
                SourceReference = $"Sun Life PDF page {page.PageNumber}"
            });
        }
    }

    private static decimal? SumPremium(IEnumerable<RateLine> lines, string benefit, bool current)
    {
        var values = lines
            .Where(line => line.Benefit.Equals(benefit, StringComparison.OrdinalIgnoreCase))
            .Select(line => current ? line.CurrentPremium : line.RenewalPremium)
            .Where(value => value.HasValue)
            .Select(value => value!.Value)
            .ToList();

        return values.Count == 0 ? null : decimal.Round(values.Sum(), 2);
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

    private static decimal? MatchPercent(string text, string pattern)
    {
        var match = Regex.Match(text, pattern, RegexOptions.IgnoreCase);
        return match.Success ? ParsingHelpers.ParseDecimal(match.Groups["value"].Value) : null;
    }

    private static decimal? Money(Match match, string groupName)
    {
        return match.Success ? ParsingHelpers.ParseDecimal(NormalizeMoneyText(match.Groups[groupName].Value)) : null;
    }

    private static string ParseDetailedRateBasis(string prefix)
    {
        var match = DetailedRateBasisRegex.Match(prefix);
        return match.Success ? FormatRateBasis(match.Groups["basis"].Value) : "";
    }

    private static string FormatRateBasis(string value)
    {
        var divisor = ParsingHelpers.ParseInt(value);
        return divisor is > 0
            ? $"Per ${divisor.Value.ToString("N0", CultureInfo.InvariantCulture)}"
            : "";
    }

    private static int? Int(Match match, string groupName)
    {
        return ParsingHelpers.ParseInt(match.Groups[groupName].Value);
    }

    private static bool TryParseDate(string value, out DateOnly date)
    {
        var cleaned = value.Replace(",", "", StringComparison.Ordinal).Replace(".", "", StringComparison.Ordinal);
        var formats = new[] { "MMMM d yyyy", "MMM d yyyy", "d-MMMM-yyyy", "d-MMM-yyyy", "dd-MMM-yyyy" };
        foreach (var format in formats)
        {
            if (DateOnly.TryParseExact(cleaned, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
            {
                return true;
            }
        }

        return DateOnly.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.None, out date);
    }

    private static DateOnly? ParseExperiencePeriodEnd(string period)
    {
        var parts = period.Split(" - ", 2, StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length != 2)
        {
            return null;
        }

        return TryParseDate(parts[1], out var date) ? date : null;
    }

    private static string NormalizeMoneyText(string value)
    {
        var cleaned = SpacedCommaMoneyTokenRegex.Replace(value, ",");
        cleaned = SplitMoneyTokenRegex.Replace(cleaned, match => match.Groups["prefix"].Value + match.Groups["rest"].Value);
        return cleaned;
    }

    private sealed record SunLifeRenewalFactors(
        decimal? WeightedAdjustedLossRatio,
        decimal? Credibility,
        decimal? Ibnr,
        decimal? TargetLossRatio,
        decimal? AnnualTrendFactor,
        decimal? RateAdjustment);

    private sealed record SunLifePoolingValues(
        decimal? LargeAmountPoolingPremium,
        decimal? QuebecIndustryPoolingPremium,
        decimal? PooledClaims,
        decimal? QuebecIndustryPooledClaims);
}

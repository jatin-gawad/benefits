using System.Globalization;
using System.Text.RegularExpressions;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

/// <summary>
/// Reads an Allstate/AICC employer-paid Group Critical Illness master application.
/// The sold premiums are monthly tier rates per enrolled employee, not rates per
/// $1,000 of insured benefit.
/// </summary>
public sealed class AllstateCriticalIllnessPdfParser
{
    public const string CarrierName = "Allstate";
    public const string PerEmployeeRateBasis = "per_employee_per_month";

    private static readonly Regex ClientNameRegex = new(
        @"Legal\s*Business\s*Name:\s*(?<value>.+?)(?:\s+City\s*,?\s*Province|\s+Physical\s*Address|\s+Type\s*of\s*Business)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex EligibleEmployeesRegex = new(
        @"Total\s*number\s*of\s*employees(?:\s*/\s*members)?\s*eligible\s*for\s*coverage:\s*(?<value>\d+)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex BenefitAmountRegex = new(
        @"Employer\s*Paid\s*Basic\s*Benefit\s*Amount\s*\$?\s*\$?\s*(?<value>[\d,]+)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex TerminationAgeRegex = new(
        @"Termination\s*of\s*Coverage\s*(?<value>\d{2})\b",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex EmployeeRateRegex = new(
        @"(?:^|\s)EE\s*\$\s*(?<value>\d+(?:\.\d{1,4})?)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex FamilyRateRegex = new(
        @"(?:^|\s)F\s*\$\s*(?<value>\d+(?:\.\d{1,4})?)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex DateRegex = new(
        @"(?<month>January|February|March|April|May|June|July|August|September|October|November|December)\s*(?<day>\d{1,2})\s*,\s*(?<year>20\d{2})",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public static bool IsAllstateCriticalIllnessApplication(string filePath, string allText)
    {
        var fileName = Path.GetFileName(filePath);
        var hasAllstateIdentity = allText.Contains("Allstate Insurance Company of Canada", StringComparison.OrdinalIgnoreCase) ||
                                  allText.Contains("ALLSTATEINSURANCECOMPANYOFCANADA", StringComparison.OrdinalIgnoreCase) ||
                                  fileName.Contains("AICC", StringComparison.OrdinalIgnoreCase) ||
                                  fileName.Contains("Allstate", StringComparison.OrdinalIgnoreCase);
        var hasCriticalIllness = allText.Contains("Group Critical Illness", StringComparison.OrdinalIgnoreCase) ||
                                 allText.Contains("GroupCriticalIllness", StringComparison.OrdinalIgnoreCase);
        var hasSoldPremium = allText.Contains("Monthly premium based on sold premium", StringComparison.OrdinalIgnoreCase) ||
                             allText.Contains("Monthlypremiumbasedonsoldpremium", StringComparison.OrdinalIgnoreCase);

        return hasAllstateIdentity && hasCriticalIllness && hasSoldPremium;
    }

    public CarrierPdfParseResult Parse(string pdfPath, IReadOnlyList<PdfPageText> pages, string allText)
    {
        var normalized = ParsingHelpers.Clean(allText);
        var pageFour = pages.FirstOrDefault(page =>
            page.Text.Contains("Group Critical Illness", StringComparison.OrdinalIgnoreCase) ||
            page.Text.Contains("GroupCriticalIllness", StringComparison.OrdinalIgnoreCase));
        var rateText = pageFour?.Text ?? normalized;
        var effectiveDates = ExtractDates(normalized);

        var pdf = new PdfExtractionResult
        {
            SourcePath = pdfPath,
            Carrier = CarrierName,
            ClientName = MatchValue(ClientNameRegex, normalized),
            RenewalDate = effectiveDates.FirstOrDefault()
        };

        var eligibleEmployees = ParseInt(EligibleEmployeesRegex, normalized);
        var benefitAmount = ParseDecimal(BenefitAmountRegex, rateText);
        var terminationAge = ParseInt(TerminationAgeRegex, rateText);
        var employeeRate = ParseDecimal(EmployeeRateRegex, rateText);
        var familyRate = ParseDecimal(FamilyRateRegex, rateText);
        var planEndDate = effectiveDates.Skip(1).FirstOrDefault();

        AddTierRateLine(pdf, pdfPath, pageFour?.PageNumber, "Single", employeeRate,
            eligibleEmployees, benefitAmount, terminationAge, pdf.RenewalDate, planEndDate);
        AddTierRateLine(pdf, pdfPath, pageFour?.PageNumber, "Family", familyRate,
            eligibleEmployees, benefitAmount, terminationAge, pdf.RenewalDate, planEndDate);

        if (pdf.RateLines.Count == 0)
        {
            pdf.Warnings.Add("Allstate Group Critical Illness was identified, but the EE and Family sold monthly rates could not be extracted.");
        }

        return new CarrierPdfParseResult
        {
            Carrier = CarrierName,
            Pdf = pdf
        };
    }

    private static void AddTierRateLine(
        PdfExtractionResult pdf,
        string pdfPath,
        int? sourcePage,
        string tier,
        decimal? rate,
        int? eligibleEmployees,
        decimal? benefitAmount,
        int? terminationAge,
        DateOnly? planStart,
        DateOnly? planEnd)
    {
        if (!rate.HasValue)
        {
            return;
        }

        var planPeriod = planStart.HasValue && planEnd.HasValue
            ? $"Allstate plan year {planStart:yyyy-MM-dd} to {planEnd:yyyy-MM-dd}"
            : "Allstate sold plan";
        var eligibilityNote = eligibleEmployees.HasValue
            ? $" The application states {eligibleEmployees} eligible employees, but that is not an enrolled tier count."
            : "";
        var benefitNote = benefitAmount.HasValue ? $" Employer-paid basic benefit: {ParsingHelpers.Money(benefitAmount)}." : "";
        var terminationNote = terminationAge.HasValue ? $" Coverage termination age: {terminationAge}." : "";

        pdf.RateLines.Add(new RateLine
        {
            Benefit = "Critical Illness",
            TierOrClass = tier,
            CurrentRate = rate,
            RenewalRate = rate,
            CarrierProposedRate = rate,
            NegotiatedRate = rate,
            RateBasis = PerEmployeeRateBasis,
            FundingArrangement = "Employer-paid stand-alone policy",
            IncludeInPremiumTotal = false,
            SourceRowRole = "Conditional",
            ValidationStatus = RateLineValidationStatuses.NeedsReview,
            ValidationMessage = "Monthly rate is per enrolled employee in this tier, not per $1,000 of benefit. Enter the EE and Family enrollment counts before including premium in consolidated totals." + eligibilityNote,
            SourceReference = $"{Path.GetFileName(pdfPath)}, page {sourcePage ?? 4}: Allstate monthly sold premium per enrolled employee.{benefitNote}{terminationNote}",
            SourceCarrier = CarrierName,
            SourceFileName = Path.GetFileName(pdfPath),
            SourcePage = sourcePage ?? 4,
            SourceText = $"{tier} rate {rate:C2} per enrolled employee per month; {planPeriod}.{benefitNote}{terminationNote}{eligibilityNote}",
            CurrentRatePrecision = 2,
            RenewalRatePrecision = 2,
            CurrentPeriodLabel = planPeriod,
            RenewalPeriodLabel = planPeriod
        });
    }

    private static string MatchValue(Regex regex, string text) =>
        ParsingHelpers.Clean(regex.Match(text).Groups["value"].Value).TrimEnd('.', ',');

    private static decimal? ParseDecimal(Regex regex, string text) =>
        ParsingHelpers.ParseDecimal(regex.Match(text).Groups["value"].Value);

    private static int? ParseInt(Regex regex, string text) =>
        ParsingHelpers.ParseInt(regex.Match(text).Groups["value"].Value);

    private static List<DateOnly> ExtractDates(string text)
    {
        var candidates = new[] { text, CollapseRepeatedGlyphs(text) };
        foreach (var candidate in candidates)
        {
            var dates = DateRegex.Matches(candidate)
                .Select(match => DateOnly.TryParseExact(
                    $"{match.Groups["month"].Value} {match.Groups["day"].Value}, {match.Groups["year"].Value}",
                    "MMMM d, yyyy",
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.None,
                    out var date)
                    ? (DateOnly?)date
                    : null)
                .Where(date => date.HasValue)
                .Select(date => date!.Value)
                .Distinct()
                .ToList();
            if (dates.Count > 0)
            {
                return dates;
            }
        }

        return [];
    }

    private static string CollapseRepeatedGlyphs(string value)
    {
        var collapsed = new List<char>(value.Length);
        for (var index = 0; index < value.Length; index++)
        {
            if (index + 1 < value.Length && value[index] == value[index + 1])
            {
                collapsed.Add(value[index]);
                index++;
            }
            else
            {
                collapsed.Add(value[index]);
            }
        }

        return new string(collapsed.ToArray());
    }
}

using System.Globalization;
using System.Text.RegularExpressions;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

internal sealed record SunLifeClaimsPaidServiceRow(
    string Category,
    string Service,
    decimal CurrentSubmitted,
    decimal CurrentEligible,
    decimal CurrentPaid,
    decimal PriorSubmitted,
    decimal PriorEligible,
    decimal PriorPaid);

internal sealed class SunLifeClaimsPaidByServiceResult
{
    public List<SunLifeClaimsPaidServiceRow> Rows { get; } = [];
    public List<ClaimsSourcePeriod> Periods { get; } = [];
    public List<ClaimsServiceAdjudicationFact> ServiceFacts { get; } = [];
    public List<ClaimsServiceControlTotal> ControlTotals { get; } = [];
    public decimal? CurrentTotalPaid { get; set; }
    public decimal? PriorTotalPaid { get; set; }
}

internal static class SunLifeClaimsPaidByServiceParser
{
    private static readonly Regex PeriodPairRegex = new(
        @"(?<currentStart>[A-Za-z]{3,9}\s+\d{1,2},?\s+\d{4})\s+to\s+(?<currentEnd>[A-Za-z]{3,9}\s+\d{1,2},?\s+\d{4})\s+(?<priorStart>[A-Za-z]{3,9}\s+\d{1,2},?\s+\d{4})\s+to\s+(?<priorEnd>[A-Za-z]{3,9}\s+\d{1,2},?\s+\d{4})",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex RelationRegex = new(
        @"^Relation\s*:\s*(?<relationship>.+)$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex ServiceRowRegex = new(
        @"^(?<service>[A-Za-z][A-Za-z ',/&().+-]+?)\s+(?<currentSubmitted>-?[\d,]+(?:\.\d+)?|\([\d,]+(?:\.\d+)?\))\s+(?<currentEligible>-?[\d,]+(?:\.\d+)?|\([\d,]+(?:\.\d+)?\))\s+(?<currentPaid>-?[\d,]+(?:\.\d+)?|\([\d,]+(?:\.\d+)?\))\s+(?<currentPercent>-?[\d.]+|\([\d.]+\))%?\s+(?<priorSubmitted>-?[\d,]+(?:\.\d+)?|\([\d,]+(?:\.\d+)?\))\s+(?<priorEligible>-?[\d,]+(?:\.\d+)?|\([\d,]+(?:\.\d+)?\))\s+(?<priorPaid>-?[\d,]+(?:\.\d+)?|\([\d,]+(?:\.\d+)?\))\s+(?<priorPercent>-?[\d.]+|\([\d.]+\))%?$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex TotalClaimsRegex = new(
        @"^Canadian\s+Dollars\s+Total\s+Claims\s+(?<currentSubmitted>-?[\d,]+(?:\.\d+)?)\s+(?<currentEligible>-?[\d,]+(?:\.\d+)?)\s+(?<currentPaid>-?[\d,]+(?:\.\d+)?)\s+(?<currentPercent>-?[\d.]+)%?\s+(?<priorSubmitted>-?[\d,]+(?:\.\d+)?)\s+(?<priorEligible>-?[\d,]+(?:\.\d+)?)\s+(?<priorPaid>-?[\d,]+(?:\.\d+)?)\s+(?<priorPercent>-?[\d.]+)%?$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex RelationshipTotalRegex = new(
        @"^Total\s+for\s+(?<relationship>[A-Za-z ]+)\s+(?<currentSubmitted>-?[\d,]+(?:\.\d+)?)\s+(?<currentEligible>-?[\d,]+(?:\.\d+)?)\s+(?<currentPaid>-?[\d,]+(?:\.\d+)?)\s+(?<currentPercent>-?[\d.]+)%?\s+(?<priorSubmitted>-?[\d,]+(?:\.\d+)?)\s+(?<priorEligible>-?[\d,]+(?:\.\d+)?)\s+(?<priorPaid>-?[\d,]+(?:\.\d+)?)\s+(?<priorPercent>-?[\d.]+)%?$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex CategorySubtotalRegex = new(
        @"^(?<currentSubmitted>-?[\d,]+(?:\.\d+)?)\s+(?<currentEligible>-?[\d,]+(?:\.\d+)?)\s+(?<currentPaid>-?[\d,]+(?:\.\d+)?)\s+(?<currentPercent>-?[\d.]+)%?\s+(?<priorSubmitted>-?[\d,]+(?:\.\d+)?)\s+(?<priorEligible>-?[\d,]+(?:\.\d+)?)\s+(?<priorPaid>-?[\d,]+(?:\.\d+)?)\s+(?<priorPercent>-?[\d.]+)%?$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public static SunLifeClaimsPaidByServiceResult Parse(IReadOnlyList<PdfPageText> pages, string coverage)
    {
        var result = new SunLifeClaimsPaidByServiceResult();
        foreach (var page in pages)
        {
            var lines = page.Lines.Select(ParsingHelpers.Clean).Where(line => line.Length > 0).ToList();
            if (!lines.Any(line => RelationRegex.IsMatch(line)))
            {
                continue;
            }

            var pagePeriods = ExtractPeriods(lines, result, page);
            var currentPeriod = pagePeriods.Current;
            var priorPeriod = pagePeriods.Prior;
            var relationship = "Unknown";
            var category = "";
            foreach (var line in lines)
            {
                var relationMatch = RelationRegex.Match(line);
                if (relationMatch.Success)
                {
                    relationship = NormalizeRelationshipLabel(relationMatch.Groups["relationship"].Value);
                    continue;
                }

                if (TryMapSection(line, coverage, out var mappedCategory))
                {
                    category = mappedCategory;
                    continue;
                }

                var totalMatch = TotalClaimsRegex.Match(line);
                if (totalMatch.Success && currentPeriod is not null && priorPeriod is not null)
                {
                    result.CurrentTotalPaid = Parse(totalMatch.Groups["currentPaid"].Value);
                    result.PriorTotalPaid = Parse(totalMatch.Groups["priorPaid"].Value);
                    AddControlTotals(result, coverage, "All", "", "", "GrandTotal", totalMatch, currentPeriod, priorPeriod, page, line);
                    continue;
                }

                var relationshipTotalMatch = RelationshipTotalRegex.Match(line);
                if (relationshipTotalMatch.Success && currentPeriod is not null && priorPeriod is not null)
                {
                    AddControlTotals(result, coverage, ParsingHelpers.Clean(relationshipTotalMatch.Groups["relationship"].Value), "", "", "RelationshipTotal", relationshipTotalMatch, currentPeriod, priorPeriod, page, line);
                    continue;
                }

                var categorySubtotalMatch = CategorySubtotalRegex.Match(line);
                if (categorySubtotalMatch.Success && !string.IsNullOrWhiteSpace(category) && currentPeriod is not null && priorPeriod is not null)
                {
                    AddControlTotals(result, coverage, relationship, category, "", "CategorySubtotal", categorySubtotalMatch, currentPeriod, priorPeriod, page, line);
                    continue;
                }

                var serviceMatch = ServiceRowRegex.Match(line);
                if (!serviceMatch.Success || string.IsNullOrWhiteSpace(category) || currentPeriod is null || priorPeriod is null)
                {
                    continue;
                }

                var service = ParsingHelpers.Clean(serviceMatch.Groups["service"].Value);
                if (service.StartsWith("Total for", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var mappedCategoryForService = MapServiceCategory(category, service, coverage);
                if (relationship.Equals("All", StringComparison.OrdinalIgnoreCase))
                {
                    result.Rows.Add(new SunLifeClaimsPaidServiceRow(
                        mappedCategoryForService,
                        service,
                        Parse(serviceMatch.Groups["currentSubmitted"].Value),
                        Parse(serviceMatch.Groups["currentEligible"].Value),
                        Parse(serviceMatch.Groups["currentPaid"].Value),
                        Parse(serviceMatch.Groups["priorSubmitted"].Value),
                        Parse(serviceMatch.Groups["priorEligible"].Value),
                        Parse(serviceMatch.Groups["priorPaid"].Value)));
                    AddControlTotals(result, coverage, relationship, mappedCategoryForService, service, "AllService", serviceMatch, currentPeriod, priorPeriod, page, line);
                    continue;
                }

                AddServiceFacts(result, coverage, relationship, mappedCategoryForService, service, serviceMatch, currentPeriod, priorPeriod, page, line);
            }
        }

        return result;
    }

    private static (ClaimsSourcePeriod? Current, ClaimsSourcePeriod? Prior) ExtractPeriods(List<string> lines, SunLifeClaimsPaidByServiceResult result, PdfPageText page)
    {
        var text = ParsingHelpers.Clean(string.Join(" ", lines));
        var match = PeriodPairRegex.Match(text);
        if (!match.Success)
        {
            return (null, null);
        }

        var current = AddPeriod(result, "Current", match.Groups["currentStart"].Value, match.Groups["currentEnd"].Value, page);
        var prior = AddPeriod(result, "Prior", match.Groups["priorStart"].Value, match.Groups["priorEnd"].Value, page);
        return (current, prior);
    }

    private static ClaimsSourcePeriod AddPeriod(SunLifeClaimsPaidByServiceResult result, string role, string startText, string endText, PdfPageText page)
    {
        var start = ParseDate(startText);
        var end = ParseDate(endText);
        var existing = result.Periods.FirstOrDefault(period => period.Start == start && period.End == end);
        if (existing is not null)
        {
            return existing;
        }

        var period = new ClaimsSourcePeriod
        {
            SourceRole = role,
            Start = start,
            End = end,
            MonthCount = ClaimsExperienceDataModelBuilder.InclusiveMonthCount(start, end),
            SourceFile = $"Page {page.PageNumber}"
        };
        result.Periods.Add(period);
        return period;
    }

    private static void AddServiceFacts(SunLifeClaimsPaidByServiceResult result, string coverage, string relationship, string category, string service, Match match, ClaimsSourcePeriod currentPeriod, ClaimsSourcePeriod priorPeriod, PdfPageText page, string rawLine)
    {
        AddServiceFact(result, coverage, relationship, category, service, currentPeriod, page, rawLine, Parse(match.Groups["currentSubmitted"].Value), Parse(match.Groups["currentEligible"].Value), Parse(match.Groups["currentPaid"].Value), Parse(match.Groups["currentPercent"].Value));
        AddServiceFact(result, coverage, relationship, category, service, priorPeriod, page, rawLine, Parse(match.Groups["priorSubmitted"].Value), Parse(match.Groups["priorEligible"].Value), Parse(match.Groups["priorPaid"].Value), Parse(match.Groups["priorPercent"].Value));
    }

    private static void AddServiceFact(SunLifeClaimsPaidByServiceResult result, string coverage, string relationship, string category, string service, ClaimsSourcePeriod period, PdfPageText page, string rawLine, decimal submitted, decimal eligible, decimal paid, decimal percent)
    {
        if (submitted == 0m && eligible == 0m && paid == 0m && percent == 0m)
        {
            return;
        }

        result.ServiceFacts.Add(new ClaimsServiceAdjudicationFact
        {
            Benefit = coverage,
            Relationship = relationship,
            Category = category,
            Service = service,
            PeriodStart = period.Start,
            PeriodEnd = period.End,
            AmountSubmitted = submitted,
            AmountEligible = eligible,
            AmountPaid = paid,
            ReportedPaidPercentage = percent,
            Currency = "CAD",
            SourcePage = page.PageNumber,
            SourceEvidence = rawLine,
            Confidence = 0.96m,
            ValidationState = submitted >= eligible && eligible >= paid ? "Valid" : "NeedsReview"
        });
    }

    private static void AddControlTotals(SunLifeClaimsPaidByServiceResult result, string coverage, string relationship, string category, string service, string totalType, Match match, ClaimsSourcePeriod currentPeriod, ClaimsSourcePeriod priorPeriod, PdfPageText page, string rawLine)
    {
        AddControlTotal(result, coverage, relationship, category, service, totalType, currentPeriod, page, rawLine, Parse(match.Groups["currentSubmitted"].Value), Parse(match.Groups["currentEligible"].Value), Parse(match.Groups["currentPaid"].Value), Parse(match.Groups["currentPercent"].Value));
        AddControlTotal(result, coverage, relationship, category, service, totalType, priorPeriod, page, rawLine, Parse(match.Groups["priorSubmitted"].Value), Parse(match.Groups["priorEligible"].Value), Parse(match.Groups["priorPaid"].Value), Parse(match.Groups["priorPercent"].Value));
    }

    private static void AddControlTotal(SunLifeClaimsPaidByServiceResult result, string coverage, string relationship, string category, string service, string totalType, ClaimsSourcePeriod period, PdfPageText page, string rawLine, decimal submitted, decimal eligible, decimal paid, decimal percent)
    {
        result.ControlTotals.Add(new ClaimsServiceControlTotal
        {
            Benefit = coverage,
            Relationship = relationship,
            Category = category,
            Service = service,
            TotalType = totalType,
            PeriodStart = period.Start,
            PeriodEnd = period.End,
            AmountSubmitted = submitted,
            AmountEligible = eligible,
            AmountPaid = paid,
            ReportedPaidPercentage = percent,
            Currency = "CAD",
            SourcePage = page.PageNumber,
            SourceEvidence = rawLine
        });
    }
    private static bool TryMapSection(string line, string coverage, out string category)
    {
        category = "";
        var normalized = new string(line.Where(char.IsLetterOrDigit).Select(char.ToUpperInvariant).ToArray());
        if (coverage.Equals("Dental", StringComparison.OrdinalIgnoreCase))
        {
            category = normalized switch
            {
                "DIAGNOSTICPREVENTIVE" => "Diagnostic / Preventive",
                "SURGICALRESTORATIVE" => "Surgical / Restorative",
                "MAJORRESTORATIVE" => "Major Restorative",
                "PERIODONTICSENDODONTICS" => "Periodontics / Endodontics",
                "ORTHODONTIC" or "ORTHODONTICS" => "Orthodontics",
                _ => ""
            };
        }
        else
        {
            category = normalized switch
            {
                "HOSPITALINCANADA" or "HOSPITAL" => "Hospital",
                "PAYDIRECTDRUGS" => "Pay Direct Drugs",
                "REIMBURSEMENTDRUGS" => "Reimbursement Drugs",
                "MEDICALSERVICESEQUIPMENT" => "Medical Services and Supplies",
                "PARAMEDICALPRACTITIONERS" => "Paramedical",
                "CONTACTLENSESOREYEGLASSES" => "Vision",
                "MISCELLANEOUS" => "Other EHC",
                _ => ""
            };
        }

        return category.Length > 0;
    }

    internal static string NormalizeRelationshipLabel(string value)
    {
        var normalized = new string((value ?? "").Where(char.IsLetterOrDigit).Select(char.ToUpperInvariant).ToArray());
        return normalized switch
        {
            "EMPLOYEE" or "MEMBER" or "CERTIFICATEHOLDER" => "Employee",
            "SPOUSE" or "PARTNER" => "Spouse",
            "CHILD" or "CHILDREN" => "Child",
            "OTHERDEPENDANT" or "OTHERDEPENDENT" or "DEPENDANT" or "DEPENDENT" => "OtherDependant",
            "UNALLOCATED" => "Unallocated",
            "ALL" or "TOTAL" => "All",
            _ => "Unknown"
        };
    }
    private static string MapServiceCategory(string category, string service, string coverage)
    {
        if (!coverage.Equals("Dental", StringComparison.OrdinalIgnoreCase) ||
            !category.Equals("Periodontics / Endodontics", StringComparison.OrdinalIgnoreCase))
        {
            return category;
        }

        return service.Contains("period", StringComparison.OrdinalIgnoreCase)
            ? "Periodontics"
            : "Endodontics";
    }

    private static DateOnly ParseDate(string value)
    {
        var cleaned = ParsingHelpers.Clean(value).Replace(",", "", StringComparison.Ordinal);
        foreach (var format in new[] { "MMMM d yyyy", "MMM d yyyy" })
        {
            if (DateOnly.TryParseExact(cleaned, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out var date))
            {
                return date;
            }
        }

        return DateOnly.Parse(value, CultureInfo.InvariantCulture);
    }

    private static decimal Parse(string value)
    {
        return ParsingHelpers.ParseDecimal(value) ?? 0m;
    }
}



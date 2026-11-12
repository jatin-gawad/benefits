using System.Text.RegularExpressions;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services.PolicySchedules;

/// <summary>Reads Manulife schedule layout into source-shaped facts; it has no SQLite knowledge.</summary>
public sealed class ManulifePolicyScheduleExtractor
{
    private static readonly (string Benefit, string[] Headings)[] Benefits =
    [
        ("Employee Life", ["Employee Life Insurance", "Employee Life"]),
        ("Dependent Life", ["Dependent Life Insurance", "Dependent Life"]),
        ("AD&D", ["Accidental Death and Dismemberment", "AD&D"]),
        ("LTD", ["Long Term Disability", "Long-Term Disability", "LTD"])
    ];

    private static readonly string[] KnownLabels =
    [
        "Benefit Amount", "Non-Evidence Limit", "Participation Basis", "Waiting Period", "Qualifying Period",
        "Elimination Period", "Termination Age", "Waiver of Premium", "Maximum Benefit Period",
        "Maximum Monthly Benefit", "Maximum Benefit", "Reduction", "Survivor Benefit", "CPP/QPP",
        "Earnings", "Rounding", "Spouse", "Child", "Rehabilitation", "Repatriation", "Transportation",
        "Education", "Occupational Training", "Seat Belt", "Day-Care", "Home Alteration", "Vehicle Modification"
    ];

    public PolicyScheduleExtractionResult Extract(string sourcePath, IReadOnlyList<PdfPageText> pages)
    {
        var result = new PolicyScheduleExtractionResult { Carrier = "Manulife", SourcePath = sourcePath };
        PolicyScheduleSection? current = null;
        foreach (var page in pages)
        {
            var lines = page.Lines.Select(ParsingHelpers.Clean).Where(line => line.Length > 0).ToList();
            foreach (var line in lines)
            {
                var heading = FindHeading(line);
                if (heading is not null)
                {
                    current = new PolicyScheduleSection { Benefit = heading.Value.Benefit, Heading = line, StartPage = page.PageNumber };
                    result.Sections.Add(current);
                    continue;
                }

                if (current is null || IsRepeatedHeader(line)) continue;
                AddClassifications(current, line);
                var label = KnownLabels.FirstOrDefault(candidate => line.Contains(candidate, StringComparison.OrdinalIgnoreCase));
                if (label is not null)
                {
                    var value = ExtractValue(line, label);
                    if (!string.IsNullOrWhiteSpace(value))
                    {
                        current.Fields.Add(Field(label, value, page.PageNumber, current.Heading, line, 0.86m));
                    }
                    continue;
                }

                if (LooksLikeFormula(line))
                {
                    current.Fields.Add(Field("Benefit Formula", line, page.PageNumber, current.Heading, line, 0.82m));
                }
                else if (line.Contains("benefit", StringComparison.OrdinalIgnoreCase) || line.Contains("expense", StringComparison.OrdinalIgnoreCase))
                {
                    current.Fields.Add(Field("Supplementary Benefit", line, page.PageNumber, current.Heading, line, 0.65m));
                    result.Warnings.Add(new PolicyScheduleWarning { Code = ValidationIssueCodes.PolicyScheduleUnsupportedTerm, Message = $"Unmapped Manulife schedule term: {line}", PageNumber = page.PageNumber, SourceExcerpt = line });
                }
            }
        }

        if (result.Sections.Count == 0)
        {
            result.Warnings.Add(new PolicyScheduleWarning { Code = ValidationIssueCodes.PolicyScheduleNoBenefitSections, Message = "No supported Manulife benefit sections were found." });
        }
        return result;
    }

    private static (string Benefit, string Heading)? FindHeading(string line)
    {
        foreach (var item in Benefits)
        {
            if (item.Headings.Any(heading => line.Equals(heading, StringComparison.OrdinalIgnoreCase) ||
                                             (line.Length < 90 && line.Contains(heading, StringComparison.OrdinalIgnoreCase))))
                return (item.Benefit, line);
        }
        return null;
    }

    private static void AddClassifications(PolicyScheduleSection section, string line)
    {
        var match = Regex.Match(line, @"(?:Classification\s+Code|Class)\s*[:\-]?\s*(?<codes>[A-Z](?:\s*(?:,|and|&)\s*[A-Z])*)", RegexOptions.IgnoreCase);
        if (!match.Success && !line.Contains("All Employees", StringComparison.OrdinalIgnoreCase)) return;
        var raw = match.Success ? match.Groups["codes"].Value : "All Employees";
        if (!section.RawClassifications.Contains(raw, StringComparer.OrdinalIgnoreCase)) section.RawClassifications.Add(raw);
        foreach (Match code in Regex.Matches(raw, @"\b[A-Z]\b", RegexOptions.IgnoreCase))
            if (!section.ClassificationCodes.Contains(code.Value, StringComparer.OrdinalIgnoreCase)) section.ClassificationCodes.Add(code.Value.ToUpperInvariant());
    }

    private static string ExtractValue(string line, string label)
    {
        var index = line.IndexOf(label, StringComparison.OrdinalIgnoreCase);
        var value = index < 0 ? "" : line[(index + label.Length)..].Trim(' ', ':', '-', '–');
        return value.Length == 0 ? line : value;
    }

    private static bool LooksLikeFormula(string line) =>
        Regex.IsMatch(line, @"\b\d+(?:\.\d+)?\s*x\s+(?:annual|monthly)\s+earnings\b|\b\d+(?:\.\d+)?%\s+of\s+", RegexOptions.IgnoreCase);
    private static bool IsRepeatedHeader(string line) => line.Contains("Manulife", StringComparison.OrdinalIgnoreCase) && line.Length < 70;
    private static PolicyScheduleField Field(string label, string value, int page, string section, string excerpt, decimal confidence) => new()
    { Label = label, Value = value, PageNumber = page, SectionHeading = section, SourceExcerpt = excerpt, Confidence = confidence };
}

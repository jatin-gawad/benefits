using System.Text.RegularExpressions;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

internal sealed record ClaimsPdfIdentity(string Carrier, string ClientName, string PolicyNumber);

internal static class ClaimsPdfIdentityExtractor
{
    private static readonly Regex PolicyRegex = new(
        @"\b(?:Policy|Contract)\s*(?:Number|No\.?|#)?\s*:?\s*(?<value>[A-Za-z0-9][A-Za-z0-9-]*)\b",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public static ClaimsPdfIdentity Extract(IReadOnlyList<PdfPageText> pages)
    {
        return Extract(pages.SelectMany(page => page.Lines));
    }

    public static ClaimsPdfIdentity Extract(IEnumerable<string> rawLines)
    {
        var lines = rawLines
            .Select(ParsingHelpers.Clean)
            .Where(line => line.Length > 0)
            .ToList();
        var allText = string.Join(" ", lines);
        var carrier = DetectCarrier(allText);
        var policyNumber = "";
        var policyLineIndex = -1;

        for (var index = 0; index < lines.Count; index++)
        {
            var match = PolicyRegex.Match(lines[index]);
            if (!match.Success)
            {
                continue;
            }

            policyNumber = match.Groups["value"].Value.Trim();
            policyLineIndex = index;
            break;
        }

        var clientName = policyLineIndex < 0
            ? ""
            : FindClientName(lines, policyLineIndex);
        return new ClaimsPdfIdentity(carrier, clientName, policyNumber);
    }

    private static string DetectCarrier(string allText)
    {
        if (allText.Contains("Sun Life", StringComparison.OrdinalIgnoreCase))
        {
            return "Sun Life";
        }

        if (allText.Contains("Manulife", StringComparison.OrdinalIgnoreCase))
        {
            return "Manulife";
        }

        if (allText.Contains("Canada Life", StringComparison.OrdinalIgnoreCase))
        {
            return "Canada Life";
        }

        return "";
    }

    private static string FindClientName(IReadOnlyList<string> lines, int policyLineIndex)
    {
        for (var index = policyLineIndex - 1; index >= Math.Max(0, policyLineIndex - 6); index--)
        {
            var candidate = lines[index];
            if (IsClientCandidate(candidate))
            {
                return candidate;
            }
        }

        return "";
    }

    private static bool IsClientCandidate(string value)
    {
        if (value.Length < 3 ||
            value.Equals("Summary", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("Details", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("Current Period", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("Prior Period", StringComparison.OrdinalIgnoreCase) ||
            value.Contains("Claims Paid by Type of Service", StringComparison.OrdinalIgnoreCase) ||
            value.Contains("Lives/Premium/Claims", StringComparison.OrdinalIgnoreCase) ||
            value.Contains("Member Based", StringComparison.OrdinalIgnoreCase) ||
            value.Contains("Sun Life", StringComparison.OrdinalIgnoreCase) ||
            value.Contains("Manulife", StringComparison.OrdinalIgnoreCase) ||
            value.Contains("Canada Life", StringComparison.OrdinalIgnoreCase) ||
            value.Contains(" to ", StringComparison.OrdinalIgnoreCase) ||
            value.Contains("Period:", StringComparison.OrdinalIgnoreCase) ||
            Regex.IsMatch(value, @"^[A-Za-z]{3,9}\s+\d{1,2},?\s+\d{4}$"))
        {
            return false;
        }

        return value.Any(char.IsLetter);
    }
}

using System.Text.RegularExpressions;
using System.Globalization;

namespace RenewalReviewBuilder.App.Services;

public static class ClientNameResolver
{
    private static readonly string[] InvalidTextFragments =
    [
        "Unknown Client",
        "each benefit",
        "detailed premium illustration",
        "summary of rates chart",
        "Great-West Life Assurance",
        "Great West Life Assurance",
        "The Canada Life Assurance Company",
        "Canada Life Assurance Company",
        "Sun Life Assurance Company",
        "Manulife Financial"
    ];

    private static readonly HashSet<string> CarrierOnlyNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "Canada Life",
        "Canada Life Assurance",
        "Great-West Life",
        "Great West Life",
        "Great-West Life Assurance Co.",
        "Great West Life Assurance Co.",
        "Manulife",
        "Sun Life",
        "Equitable",
        "Desjardins",
        "Victor",
        "Empire Life"
    };

    private static readonly Regex RenewalEnvelopeRegex = new(
        @"^(?:(?:Canada Life|Manulife|Sun Life|Equitable|Desjardins|Victor|Empire Life)\s+)?Renewal\s*-\s*(?:\d{5,}\s*-\s*)?(?<client>.+?)\s*-\s*(?:Policy\s*#?\s*\d+(?:\s*-\s*\d{4})?|\d{5,}\s*-\s*\d{4}|\d{8})$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly HashSet<string> NonClientFolderNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "Batch Processing",
        "Batch Output",
        "Output",
        "Outputs",
        "Export",
        "Exports",
        "Input",
        "Inputs",
        "Downloads",
        "Desktop",
        "Temp",
        "Renewal",
        "Renewals",
        "Documents",
        "Build Renewal",
        "Renewal App - Test",
        "OneDrive - WESTLANDCORP"
    };

    public static string Resolve(string? primary, string? secondary = null, string? sourceFilePath = null)
    {
        var folderName = ResolveFromSourceFolder(sourceFilePath);
        foreach (var candidate in new[] { primary, secondary })
        {
            var cleaned = CleanCandidate(candidate);
            if (!IsInvalidClientName(cleaned))
            {
                return ApplyNumberedFolderPrefix(cleaned, folderName);
            }
        }

        var fileName = ResolveFromSourceFileName(sourceFilePath);
        var sourceFileName = Path.GetFileNameWithoutExtension(sourceFilePath ?? "");
        if (!IsInvalidClientName(fileName) && RenewalEnvelopeRegex.IsMatch(sourceFileName))
        {
            if (!IsInvalidClientName(folderName) &&
                Normalize(fileName).Equals(Normalize(folderName), StringComparison.OrdinalIgnoreCase))
            {
                return folderName;
            }

            return ApplyNumberedFolderPrefix(fileName, folderName);
        }

        if (!IsInvalidClientName(folderName))
        {
            return folderName;
        }

        return IsInvalidClientName(fileName) ? "" : fileName;
    }

    private static string ApplyNumberedFolderPrefix(string candidate, string folderName)
    {
        if (string.IsNullOrWhiteSpace(candidate) ||
            string.IsNullOrWhiteSpace(folderName) ||
            Regex.IsMatch(candidate, @"^\s*\d{4,}\b"))
        {
            return candidate;
        }

        var folderMatch = Regex.Match(folderName, @"^\s*(\d{4,})\s+(.+?)\s*$");
        if (!folderMatch.Success)
        {
            return candidate;
        }

        var folderNumber = folderMatch.Groups[1].Value;
        var folderLegalName = CleanCandidate(folderMatch.Groups[2].Value);
        if (string.IsNullOrWhiteSpace(folderLegalName))
        {
            return candidate;
        }

        var normalizedCandidate = Normalize(candidate);
        var normalizedFolderLegalName = Normalize(folderLegalName);
        return normalizedCandidate.StartsWith(normalizedFolderLegalName, StringComparison.OrdinalIgnoreCase)
            ? $"{folderNumber} {candidate}"
            : candidate;
    }

    public static bool IsInvalidClientName(string? value)
    {
        var cleaned = CleanCandidate(value);
        if (string.IsNullOrWhiteSpace(cleaned))
        {
            return true;
        }

        if (cleaned.Length > 140)
        {
            return true;
        }

        if (cleaned.All(character => char.IsDigit(character) || char.IsWhiteSpace(character) || character is '-' or '_' or '.'))
        {
            return true;
        }

        if (CarrierOnlyNames.Contains(cleaned))
        {
            return true;
        }

        var normalized = Normalize(cleaned);
        return InvalidTextFragments.Any(fragment => normalized.Contains(Normalize(fragment), StringComparison.OrdinalIgnoreCase));
    }

    public static string ToDisplayName(string? value)
    {
        var cleaned = CleanCandidate(value);
        if (string.IsNullOrWhiteSpace(cleaned))
        {
            return "";
        }

        var letters = cleaned.Where(char.IsLetter).ToList();
        if (letters.Count == 0 || letters.Any(char.IsLower))
        {
            return cleaned;
        }

        var lower = cleaned.ToLower(CultureInfo.CurrentCulture);
        return CultureInfo.CurrentCulture.TextInfo.ToTitleCase(lower);
    }

    private static string ResolveFromSourceFolder(string? sourceFilePath)
    {
        if (string.IsNullOrWhiteSpace(sourceFilePath))
        {
            return "";
        }

        var directoryPath = Path.GetDirectoryName(sourceFilePath);
        if (string.IsNullOrWhiteSpace(directoryPath))
        {
            return "";
        }

        var directory = new DirectoryInfo(directoryPath);
        for (var depth = 0; directory is not null && depth < 5; depth++, directory = directory.Parent)
        {
            var candidate = CleanCandidate(directory.Name);
            if (NonClientFolderNames.Contains(candidate))
            {
                continue;
            }

            if (!IsInvalidClientName(candidate))
            {
                return candidate;
            }
        }

        return "";
    }

    private static string ResolveFromSourceFileName(string? sourceFilePath)
    {
        if (string.IsNullOrWhiteSpace(sourceFilePath))
        {
            return "";
        }

        var candidate = Path.GetFileNameWithoutExtension(sourceFilePath);
        if (string.IsNullOrWhiteSpace(candidate))
        {
            return "";
        }

        var renewalEnvelope = RenewalEnvelopeRegex.Match(candidate);
        if (renewalEnvelope.Success)
        {
            return CleanCandidate(renewalEnvelope.Groups["client"].Value);
        }

        candidate = Regex.Replace(candidate, @"^\s*(Canada Life|Manulife|Sun Life|Equitable|Desjardins|Victor|Empire Life)\s+Renewal\s*-\s*", "", RegexOptions.IgnoreCase);
        candidate = Regex.Replace(candidate, @"\s*-\s*Policy\s*#?.*$", "", RegexOptions.IgnoreCase);
        candidate = Regex.Replace(candidate, @"\s*-\s*\d{4}.*$", "", RegexOptions.IgnoreCase);
        candidate = Regex.Replace(candidate, @"\b(Group Information Report|Renewal|Detailed Package|DetailedPackage)\b", "", RegexOptions.IgnoreCase);
        candidate = Regex.Replace(candidate, @"\s{2,}", " ");
        return CleanCandidate(candidate.Trim(' ', '-', '_', '.'));
    }

    private static string CleanCandidate(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return "";
        }

        var cleaned = Regex.Replace(value.Trim(), @"\s+", " ");
        cleaned = Regex.Split(cleaned, @"\[\[PAGE\s+\d+\]\]", RegexOptions.IgnoreCase)[0];
        cleaned = Regex.Split(cleaned, @"\bIntroduction\s+Policyholder\b", RegexOptions.IgnoreCase)[0];
        return cleaned.Trim(' ', ':', '-', ',');
    }

    private static string Normalize(string value)
    {
        return new string(value
            .ToUpperInvariant()
            .Where(char.IsLetterOrDigit)
            .ToArray());
    }
}

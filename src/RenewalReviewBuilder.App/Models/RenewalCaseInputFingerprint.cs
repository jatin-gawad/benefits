using System.Security.Cryptography;
using System.Text;

namespace RenewalReviewBuilder.App.Models;

/// <summary>
/// Captures the exact source set used for an analysis. The output folder is deliberately
/// excluded because changing where a pack is saved does not change the reviewed evidence.
/// </summary>
public static class RenewalCaseInputFingerprint
{
    private static readonly HashSet<string> SourceFolderExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".pdf",
        ".xlsx",
        ".xlsm",
        ".xls"
    };

    public static string Create(RenewalCase? renewalCase)
    {
        renewalCase ??= new RenewalCase();
        var parts = new List<string>
        {
            $"client={NormalizeText(renewalCase.ClientName)}",
            $"policy={NormalizeText(renewalCase.PolicyNumber)}",
            $"renewal-date={renewalCase.RenewalDate?.ToString("yyyy-MM-dd") ?? ""}"
        };

        AddFile(parts, "primary-pdf", renewalCase.PdfPath);
        AddFiles(parts, "additional-pdf", renewalCase.AdditionalPdfPaths);
        AddFile(parts, "primary-workbook", renewalCase.WorkbookPath);
        AddFiles(parts, "additional-workbook", renewalCase.AdditionalWorkbookPaths);
        AddSourceFolder(parts, "current-folder", renewalCase.CurrentYearFolder);
        AddSourceFolder(parts, "prior-folder", renewalCase.PriorYearFolder);
        AddSourceFolder(parts, "aso-billing-folder", renewalCase.AsoBillingFolder);
        AddHcsaCalculation(parts, renewalCase.HcsaCalculation);
        AddAllstateCriticalIllnessEnrollment(parts, renewalCase.AllstateCriticalIllnessEnrollment);
        AddFile(parts, "powerpoint-template", renewalCase.PowerPointTemplatePath);

        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(string.Join('\n', parts))));
    }

    private static void AddFiles(List<string> parts, string role, IEnumerable<string> paths)
    {
        foreach (var path in paths
                     .Where(path => !string.IsNullOrWhiteSpace(path))
                     .Select(NormalizePath)
                     .Distinct(StringComparer.OrdinalIgnoreCase)
                     .OrderBy(path => path, StringComparer.OrdinalIgnoreCase))
        {
            AddNormalizedFile(parts, role, path);
        }
    }

    private static void AddFile(List<string> parts, string role, string path) =>
        AddNormalizedFile(parts, role, NormalizePath(path));

    private static void AddNormalizedFile(List<string> parts, string role, string path)
    {
        if (string.IsNullOrEmpty(path))
        {
            parts.Add($"{role}=<none>");
            return;
        }

        try
        {
            var file = new FileInfo(path);
            parts.Add(file.Exists
                ? $"{role}={path}|{file.Length}|{file.LastWriteTimeUtc.Ticks}"
                : $"{role}={path}|<missing>");
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
            parts.Add($"{role}={path}|<unavailable>");
        }
    }

    private static void AddSourceFolder(List<string> parts, string role, string path)
    {
        var normalized = NormalizePath(path);
        if (string.IsNullOrEmpty(normalized))
        {
            parts.Add($"{role}=<none>");
            return;
        }

        if (!Directory.Exists(normalized))
        {
            parts.Add($"{role}={normalized}|<missing>");
            return;
        }

        parts.Add($"{role}={normalized}");
        try
        {
            foreach (var sourcePath in Directory
                         .EnumerateFiles(normalized, "*", SearchOption.TopDirectoryOnly)
                         .Where(file => SourceFolderExtensions.Contains(Path.GetExtension(file)))
                         .Select(NormalizePath)
                         .OrderBy(file => file, StringComparer.OrdinalIgnoreCase))
            {
                AddNormalizedFile(parts, $"{role}-source", sourcePath);
            }
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
            parts.Add($"{role}-source=<unavailable>");
        }
    }

    private static void AddHcsaCalculation(List<string> parts, HcsaCalculation? calculation)
    {
        var entries = calculation?.Classes
            .Where(item => item.IsIncluded)
            .OrderBy(item => item.ClassName, StringComparer.OrdinalIgnoreCase)
            .ThenBy(item => item.EligibleEmployees)
            .ToList() ?? [];

        if (entries.Count == 0)
        {
            parts.Add("hcsa=<none>");
            return;
        }

        foreach (var item in entries)
        {
            parts.Add($"hcsa={NormalizeText(item.ClassName)}|{item.EligibleEmployees}|{item.CurrentAnnualCredit:0.####}|{item.RenewalAnnualCredit:0.####}|{item.UsagePercent:0.####}|{item.AdministrationFeePercent:0.####}");
        }
    }

    private static void AddAllstateCriticalIllnessEnrollment(
        List<string> parts,
        AllstateCriticalIllnessEnrollment? enrollment)
    {
        parts.Add(enrollment?.HasEntries == true
            ? $"allstate-ci-enrollment={enrollment.EmployeeOnlyCount}|{enrollment.FamilyCount}"
            : "allstate-ci-enrollment=<none>");
    }

    private static string NormalizePath(string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return "";
        }

        try
        {
            return Path.GetFullPath(path.Trim()).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                .ToUpperInvariant();
        }
        catch (Exception exception) when (exception is ArgumentException or NotSupportedException or PathTooLongException)
        {
            return path.Trim().ToUpperInvariant();
        }
    }

    private static string NormalizeText(string? value) =>
        string.Join(' ', (value ?? "").Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries))
            .ToUpperInvariant();
}

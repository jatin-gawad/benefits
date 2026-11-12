using System.Globalization;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

/// <summary>
/// Reads Equitable Health/Dental claims-banding tables. Unlike the GroupHEALTH chart parser,
/// Equitable places prior and current periods in a table on the same renewal-report page.
/// </summary>
public sealed class EquitableClaimsRangePdfParser
{
    private static readonly Regex PeriodHeaderRegex = new(
        @"(?<priorFrom>\d{2}-[A-Za-z]{3}-\d{4})\s+to\s+(?<priorTo>\d{2}-[A-Za-z]{3}-\d{4})\s+(?<currentFrom>\d{2}-[A-Za-z]{3}-\d{4})\s+to\s+(?<currentTo>\d{2}-[A-Za-z]{3}-\d{4})",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex BandRowRegex = new(
        @"^(?<label>Under\s+\$[\d,]+(?:\.\d+)?|\$[\d,]+(?:\.\d+)?\s*-\s*\$[\d,]+(?:\.\d+)?|\$[\d,]+(?:\.\d+)?\s*\+)\s+(?<priorCount>\d+)\s+\$[\d,]+\.\d{2}\s+[\d.]+%\s+(?<currentCount>\d+)\s+\$[\d,]+\.\d{2}\s+[\d.]+%$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex FooterIdentityRegex = new(
        @"\b(?<policy>\d{4,})\s*\|\s*(?<client>[^|]+?)\s*\|\s*Renewal\b",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex MoneyNumberRegex = new(
        @"(?<value>[\d,]+(?:\.\d+)?)",
        RegexOptions.Compiled);

    public IReadOnlyList<ClaimsRangeReport> Parse(string pdfPath)
    {
        if (string.IsNullOrWhiteSpace(pdfPath))
        {
            throw new ArgumentException("An Equitable renewal PDF path is required.", nameof(pdfPath));
        }

        if (!File.Exists(pdfPath))
        {
            throw new FileNotFoundException("The Equitable renewal PDF was not found.", pdfPath);
        }

        var fullPath = Path.GetFullPath(pdfPath);
        var before = CaptureSourceSnapshot(fullPath);
        var pages = new PdfTextExtractor().Extract(fullPath);
        var reports = new Dictionary<string, ClaimsRangeReport>(StringComparer.OrdinalIgnoreCase);

        foreach (var page in pages)
        {
            var benefit = DetectBenefit(page.Text);
            if (benefit is null)
            {
                continue;
            }

            var header = PeriodHeaderRegex.Match(page.Text);
            if (!header.Success ||
                !TryParseDate(header.Groups["priorFrom"].Value, out var priorStart) ||
                !TryParseDate(header.Groups["priorTo"].Value, out var priorEnd) ||
                !TryParseDate(header.Groups["currentFrom"].Value, out var currentStart) ||
                !TryParseDate(header.Groups["currentTo"].Value, out var currentEnd))
            {
                continue;
            }

            var identity = FooterIdentityRegex.Match(page.Text);
            var policyNumber = identity.Success ? ParsingHelpers.Clean(identity.Groups["policy"].Value) : "";
            var clientName = identity.Success ? ParsingHelpers.Clean(identity.Groups["client"].Value) : "";
            var priorPage = CreatePage(benefit.Value, page.PageNumber, fullPath);
            var currentPage = CreatePage(benefit.Value, page.PageNumber, fullPath);
            var sortOrder = 0;

            foreach (var line in page.Lines.Select(ParsingHelpers.Clean))
            {
                var rowMatch = BandRowRegex.Match(line);
                if (!rowMatch.Success)
                {
                    continue;
                }

                var band = ParseBand(rowMatch.Groups["label"].Value, sortOrder++);
                priorPage.Rows.Add(CreateRow(
                    benefit.Value.Key,
                    band,
                    ParsingHelpers.ParseInt(rowMatch.Groups["priorCount"].Value) ?? 0,
                    page.PageNumber,
                    priorPage.SourceReference));
                currentPage.Rows.Add(CreateRow(
                    benefit.Value.Key,
                    band,
                    ParsingHelpers.ParseInt(rowMatch.Groups["currentCount"].Value) ?? 0,
                    page.PageNumber,
                    currentPage.SourceReference));
            }

            AddPage(reports, priorStart, priorEnd, clientName, policyNumber, priorPage, fullPath, before);
            AddPage(reports, currentStart, currentEnd, clientName, policyNumber, currentPage, fullPath, before);
        }

        foreach (var report in reports.Values)
        {
            Validate(report);
        }

        var after = CaptureSourceSnapshot(fullPath);
        if (before != after)
        {
            foreach (var report in reports.Values)
            {
                report.Errors.Add("The Equitable claims-banding source PDF changed while it was being analyzed.");
            }
        }

        return reports.Values
            .OrderBy(report => report.PeriodStart)
            .ToList();
    }

    private static (string Key, string DisplayName)? DetectBenefit(string text)
    {
        if (text.Contains("Health claims banding", StringComparison.OrdinalIgnoreCase))
        {
            return (ClaimsRangeBenefitKeys.Health, "Extended Health Care");
        }

        if (text.Contains("Dental claims banding", StringComparison.OrdinalIgnoreCase))
        {
            return (ClaimsRangeBenefitKeys.Dental, "Dental Care");
        }

        return null;
    }

    private static ClaimsRangePage CreatePage(
        (string Key, string DisplayName) benefit,
        int pageNumber,
        string sourcePath)
    {
        return new ClaimsRangePage
        {
            Benefit = benefit.Key,
            BenefitDisplayName = benefit.DisplayName,
            SourcePage = pageNumber,
            SourceReference = $"{Path.GetFileName(sourcePath)} page {pageNumber}"
        };
    }

    private static ClaimsRangeBandDefinition ParseBand(string rawLabel, int sortOrder)
    {
        var label = ParsingHelpers.Clean(rawLabel);
        var values = MoneyNumberRegex.Matches(label)
            .Select(match => decimal.Parse(
                match.Groups["value"].Value.Replace(",", "", StringComparison.Ordinal),
                CultureInfo.InvariantCulture))
            .ToList();
        var lower = 0;
        int? upper = null;

        if (label.StartsWith("Under", StringComparison.OrdinalIgnoreCase) && values.Count > 0)
        {
            upper = Math.Max(0, decimal.ToInt32(decimal.Ceiling(values[0])) - 1);
        }
        else if (label.Contains('+') && values.Count > 0)
        {
            lower = decimal.ToInt32(decimal.Ceiling(values[0]));
        }
        else if (values.Count >= 2)
        {
            lower = decimal.ToInt32(decimal.Ceiling(values[0]));
            upper = decimal.ToInt32(decimal.Floor(values[1]));
        }

        var code = upper.HasValue
            ? $"RANGE_{lower}_{upper.Value}"
            : $"RANGE_{lower}_PLUS";
        return new ClaimsRangeBandDefinition
        {
            RangeCode = code,
            DisplayLabel = label.Replace(" - ", "-", StringComparison.Ordinal),
            SortOrder = sortOrder,
            LowerBoundInclusive = lower,
            UpperBoundInclusive = upper,
            PdfMatchKey = label
        };
    }

    private static ClaimsRangeRow CreateRow(
        string benefit,
        ClaimsRangeBandDefinition band,
        int count,
        int pageNumber,
        string sourceReference)
    {
        return new ClaimsRangeRow
        {
            Benefit = benefit,
            RangeCode = band.RangeCode,
            RangeLabel = band.DisplayLabel,
            RawRangeLabel = band.PdfMatchKey,
            SortOrder = band.SortOrder,
            LowerBoundInclusive = band.LowerBoundInclusive,
            UpperBoundInclusive = band.UpperBoundInclusive,
            CertificateCount = count,
            SourcePage = pageNumber,
            SourceReference = sourceReference
        };
    }

    private static void AddPage(
        IDictionary<string, ClaimsRangeReport> reports,
        DateOnly periodStart,
        DateOnly periodEnd,
        string clientName,
        string policyNumber,
        ClaimsRangePage page,
        string sourcePath,
        SourceSnapshot snapshot)
    {
        var key = $"{periodStart:yyyyMMdd}|{periodEnd:yyyyMMdd}";
        if (!reports.TryGetValue(key, out var report))
        {
            report = new ClaimsRangeReport
            {
                SourcePath = sourcePath,
                SourceFileName = Path.GetFileName(sourcePath),
                SourceSnapshotIsAuthoritative = true,
                SourceFileLength = snapshot.Length,
                SourceLastWriteTimeUtc = snapshot.LastWriteTimeUtc,
                SourceContentSha256 = snapshot.ContentSha256,
                Carrier = "Equitable",
                ClientName = clientName,
                PolicyNumber = policyNumber,
                PeriodStart = periodStart,
                PeriodEnd = periodEnd
            };
            reports[key] = report;
        }

        if (string.IsNullOrWhiteSpace(report.ClientName)) report.ClientName = clientName;
        if (string.IsNullOrWhiteSpace(report.PolicyNumber)) report.PolicyNumber = policyNumber;
        report.Pages.RemoveAll(existing =>
            existing.Benefit.Equals(page.Benefit, StringComparison.OrdinalIgnoreCase));
        report.Pages.Add(page);
    }

    private static void Validate(ClaimsRangeReport report)
    {
        if (report.Pages.Count == 0 || report.Pages.Any(page => page.Rows.Count == 0))
        {
            report.Errors.Add("No complete Equitable claims-banding rows were found.");
        }

        if (string.IsNullOrWhiteSpace(report.PolicyNumber))
        {
            report.Errors.Add("Equitable claims-banding policy number was not found.");
        }

        if (string.IsNullOrWhiteSpace(report.ClientName))
        {
            report.Errors.Add("Equitable claims-banding client name was not found.");
        }

        foreach (var page in report.Pages)
        {
            if (page.Rows.Select(row => row.RangeCode).Distinct(StringComparer.OrdinalIgnoreCase).Count() != page.Rows.Count)
            {
                report.Errors.Add($"Duplicate {page.BenefitDisplayName} claims-banding rows were found on page {page.SourcePage}.");
            }
        }
    }

    private static bool TryParseDate(string value, out DateOnly date)
    {
        return DateOnly.TryParseExact(
            value,
            "dd-MMM-yyyy",
            CultureInfo.InvariantCulture,
            DateTimeStyles.None,
            out date);
    }

    private static SourceSnapshot CaptureSourceSnapshot(string path)
    {
        var info = new FileInfo(path);
        var length = info.Length;
        var lastWrite = info.LastWriteTimeUtc;
        using var stream = File.OpenRead(path);
        var hash = Convert.ToHexString(SHA256.HashData(stream));
        return new SourceSnapshot(length, lastWrite, hash);
    }

    private sealed record SourceSnapshot(long Length, DateTime LastWriteTimeUtc, string ContentSha256);
}

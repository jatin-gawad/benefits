using System.Globalization;
using System.Text.RegularExpressions;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

/// <summary>
/// Reads GroupHEALTH Summary of Monthly Activity reports from a year/month billing folder.
/// It deliberately uses the actual billed fee and GST rather than assuming a fixed administration percentage.
/// </summary>
public sealed class AsoMonthlyActivityPdfParser
{
    private static readonly Regex FilePeriodRegex = new(@"_(?<year>20\d{2})(?<month>0[1-9]|1[0-2])_", RegexOptions.Compiled);
    private static readonly Regex ChargesPeriodRegex = new(
        @"Charges\s+for\s+the\s+Period\s*:\s*\d{1,2}\s+(?<month>[A-Za-z]{3,9})\s+(?<year>20\d{2})",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private readonly PdfTextExtractor _extractor;

    public AsoMonthlyActivityPdfParser(PdfTextExtractor? extractor = null)
    {
        _extractor = extractor ?? new PdfTextExtractor();
    }

    public AsoBillingFolderExtractionResult ExtractFolder(string folderPath)
    {
        var result = new AsoBillingFolderExtractionResult();
        if (string.IsNullOrWhiteSpace(folderPath) || !Directory.Exists(folderPath))
        {
            return result;
        }

        foreach (var path in Directory.EnumerateFiles(folderPath, "*.pdf", SearchOption.AllDirectories)
                     .Where(IsMonthlyActivityReport)
                     .OrderBy(path => path, StringComparer.OrdinalIgnoreCase))
        {
            try
            {
                var record = Extract(path);
                if (record is not null)
                {
                    result.Records.Add(record);
                }
            }
            catch (Exception ex)
            {
                result.Warnings.Add($"{Path.GetFileName(path)} could not be read: {ex.Message}");
            }
        }

        var distinctRecords = result.Records
            .GroupBy(record => record.ServiceMonth)
            .Select(group => group.OrderByDescending(record => record.SourcePath, StringComparer.OrdinalIgnoreCase).First())
            .OrderBy(record => record.ServiceMonth)
            .ToList();
        result.Records.Clear();
        result.Records.AddRange(distinctRecords);
        return result;
    }

    public AsoBillingActivityRecord? Extract(string pdfPath)
    {
        if (!IsMonthlyActivityReport(pdfPath))
        {
            return null;
        }

        var pages = _extractor.Extract(pdfPath);
        var text = string.Join(" ", pages.Select(page => page.Text));
        if (!text.Contains("Summary of Monthly Activity", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        var lines = pages.SelectMany(page => page.Lines)
            .Select(ParsingHelpers.Clean)
            .Where(line => line.Length > 0)
            .ToList();
        var paidClaims = ReadFirstAmount(lines, "Dental");
        if (!paidClaims.HasValue)
        {
            return null;
        }

        var administrationFee = ReadFirstAmount(lines, "Fees");
        var gst = ReadFirstAmount(lines, "GST Tax");
        var total = ReadFirstAmount(lines, "Claim Total") ?? paidClaims.Value + (administrationFee ?? 0m) + (gst ?? 0m);
        var serviceMonth = ResolveServiceMonth(Path.GetFileName(pdfPath), text);
        if (!serviceMonth.HasValue)
        {
            throw new InvalidDataException("The service month could not be read from the file name or charges period.");
        }

        return new AsoBillingActivityRecord
        {
            ServiceMonth = serviceMonth.Value,
            PaidDentalClaims = paidClaims.Value,
            AdministrationFee = administrationFee,
            Gst = gst,
            TotalAsoCost = total,
            SourcePath = pdfPath
        };
    }

    private static bool IsMonthlyActivityReport(string path) =>
        Path.GetFileName(path).Contains("Summary", StringComparison.OrdinalIgnoreCase) &&
        Path.GetFileName(path).Contains("Activity", StringComparison.OrdinalIgnoreCase);

    private static decimal? ReadFirstAmount(IEnumerable<string> lines, string label)
    {
        var pattern = new Regex($@"^{Regex.Escape(label)}\s+\$?(?<amount>-?[\d,]+(?:\.\d{{2}})?)\b", RegexOptions.IgnoreCase);
        foreach (var line in lines)
        {
            var match = pattern.Match(line);
            if (!match.Success)
            {
                continue;
            }

            if (decimal.TryParse(match.Groups["amount"].Value.Replace(",", ""), NumberStyles.Number, CultureInfo.InvariantCulture, out var amount))
            {
                return amount;
            }
        }

        return null;
    }

    private static DateOnly? ResolveServiceMonth(string fileName, string text)
    {
        var fileMatch = FilePeriodRegex.Match(fileName);
        if (fileMatch.Success &&
            int.TryParse(fileMatch.Groups["year"].Value, out var fileYear) &&
            int.TryParse(fileMatch.Groups["month"].Value, out var fileMonth))
        {
            return new DateOnly(fileYear, fileMonth, 1);
        }

        var chargesMatch = ChargesPeriodRegex.Match(text);
        if (!chargesMatch.Success)
        {
            return null;
        }

        return DateTime.TryParseExact(
            $"{chargesMatch.Groups["month"].Value} {chargesMatch.Groups["year"].Value}",
            ["MMM yyyy", "MMMM yyyy"],
            CultureInfo.InvariantCulture,
            DateTimeStyles.None,
            out var parsed)
            ? new DateOnly(parsed.Year, parsed.Month, 1)
            : null;
    }
}

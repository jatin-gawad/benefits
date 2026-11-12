using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

/// <summary>Deterministic, read-only extraction of client history workbook rows. It does not save or classify with AI.</summary>
public sealed class ClientHistoryWorkbookPreviewParser
{
    private static readonly string[] SectionNames = ["Benefit Plan Amendments", "Market Survey", "Change In Carrier", "Meetings", "Other"];
    private static readonly Regex OrdinalDate = new(@"\b(?<day>\d{1,2})(st|nd|rd|th)\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public bool LooksLikeClientHistoryWorkbook(string filePath)
    {
        if (!File.Exists(filePath) || !Path.GetExtension(filePath).Equals(".xlsx", StringComparison.OrdinalIgnoreCase)) return false;
        try
        {
            using var workbook = new XLWorkbook(filePath);
            return workbook.Worksheets.Any(sheet => IsLikelyHistorySheet(sheet));
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            return false;
        }
    }

    public ClientHistoryWorkbookPreview Parse(string filePath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);
        if (!File.Exists(filePath)) throw new FileNotFoundException("Client history workbook was not found.", filePath);

        var preview = new ClientHistoryWorkbookPreview { SourceFile = Path.GetFullPath(filePath) };
        try
        {
            using var workbook = new XLWorkbook(filePath);
            var sheets = workbook.Worksheets.Where(IsLikelyHistorySheet).ToList();
            if (sheets.Count == 0)
            {
                preview.Records.Add(Failed(Path.GetFullPath(filePath), "", 0, "No likely client history worksheet was detected."));
                return preview;
            }

            foreach (var sheet in sheets)
            {
                ParseSheet(sheet, preview);
            }
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            preview.Records.Add(Failed(Path.GetFullPath(filePath), "", 0, $"Workbook could not be read: {ex.Message}"));
        }

        return preview;
    }

    private static bool IsLikelyHistorySheet(IXLWorksheet sheet)
    {
        if (sheet.Name.Contains("history", StringComparison.OrdinalIgnoreCase)) return true;
        var text = SheetText(sheet, 80);
        return SectionNames.Any(section => text.Contains(section, StringComparison.OrdinalIgnoreCase));
    }

    private static void ParseSheet(IXLWorksheet sheet, ClientHistoryWorkbookPreview preview)
    {
        var lastRow = sheet.LastRowUsed()?.RowNumber() ?? 0;
        var section = "";
        DateOnly? mostRecentDate = null;
        for (var row = 1; row <= lastRow; row++)
        {
            try
            {
                var text = RowText(sheet, row);
                if (string.IsNullOrWhiteSpace(text)) continue;
                var detectedSection = SectionNames.FirstOrDefault(name => text.Equals(name, StringComparison.OrdinalIgnoreCase));
                if (detectedSection is not null)
                {
                    section = detectedSection;
                    mostRecentDate = null;
                    continue;
                }
                if (string.IsNullOrEmpty(section) || IsColumnHeader(text)) continue;

                var date = FindDate(sheet, row, text);
                if (date.HasValue) mostRecentDate = date;
                var effectiveDate = date ?? mostRecentDate;
                var body = RemoveDate(text);
                if (string.IsNullOrWhiteSpace(body)) continue;

                // Undated Other rows are attributes: retain the evidence but deliberately do not create timeline events.
                if (section.Equals("Other", StringComparison.OrdinalIgnoreCase) && !effectiveDate.HasValue)
                {
                    preview.Records.Add(Create(preview.SourceFile, sheet, row, section, null, body, ClientHistoryPreviewStatuses.Parsed,
                        "other-client-attribute", 0.90m, "Undated Other item retained as a client attribute, not a timeline event."));
                    continue;
                }

                var status = effectiveDate.HasValue ? ClientHistoryPreviewStatuses.Parsed : ClientHistoryPreviewStatuses.NeedsReview;
                var message = effectiveDate.HasValue ? "" : "No effective date was found for this history row.";
                preview.Records.Add(Create(preview.SourceFile, sheet, row, section, effectiveDate, body, status,
                    date.HasValue ? "section-row-date" : "section-row-inherited-date", date.HasValue ? 0.95m : (effectiveDate.HasValue ? 0.82m : 0.45m), message));
            }
            catch (Exception ex)
            {
                AppLogger.Error(ex);
                preview.Records.Add(Failed(preview.SourceFile, sheet.Name, row, $"Row could not be parsed: {ex.Message}"));
            }
        }
    }

    private static ClientHistoryPreviewEvent Create(string sourceFile, IXLWorksheet sheet, int row, string section, DateOnly? date, string text, string status, string method, decimal confidence, string message) => new()
    {
        EffectiveDate = date,
        EventType = section,
        Title = FirstSentence(text),
        Description = text,
        BenefitCategory = BenefitCategory(text),
        SourceSheet = sheet.Name,
        SourceRow = row,
        SourceFile = sourceFile,
        OriginalText = RowText(sheet, row),
        ExtractionMethod = method,
        Confidence = confidence,
        ValidationMessages = message,
        PreviewStatus = status
    };

    private static ClientHistoryPreviewEvent Failed(string file, string sheet, int row, string message) => new()
    {
        SourceFile = file, SourceSheet = sheet, SourceRow = row, EventType = "Unclassified", ExtractionMethod = "row-error", Confidence = 0m,
        ValidationMessages = message, PreviewStatus = ClientHistoryPreviewStatuses.Failed
    };

    private static DateOnly? FindDate(IXLWorksheet sheet, int row, string rowText)
    {
        foreach (var cell in sheet.Row(row).CellsUsed())
        {
            if (cell.TryGetValue<DateTime>(out var dateTime)) return DateOnly.FromDateTime(dateTime);
            if (TryParseDate(cell.GetFormattedString(), out var parsed)) return parsed;
        }
        return TryParseDate(rowText, out var value) ? value : null;
    }

    private static bool TryParseDate(string text, out DateOnly value)
    {
        var normalized = OrdinalDate.Replace(text, "${day}");
        normalized = normalized.Replace("Sept", "Sep", StringComparison.OrdinalIgnoreCase)
            .Replace("Febuary", "February", StringComparison.OrdinalIgnoreCase)
            .Replace("Octobr", "October", StringComparison.OrdinalIgnoreCase);
        foreach (var format in new[] { "d/M/yyyy", "M/d/yyyy", "d-M-yyyy", "d MMM yyyy", "d MMMM yyyy", "MMMM d yyyy", "MMM d yyyy" })
        {
            if (DateOnly.TryParseExact(normalized, format, CultureInfo.GetCultureInfo("en-CA"), DateTimeStyles.AllowWhiteSpaces, out value)) return true;
        }
        return DateOnly.TryParse(normalized, CultureInfo.GetCultureInfo("en-CA"), DateTimeStyles.AllowWhiteSpaces, out value);
    }

    private static string SheetText(IXLWorksheet sheet, int rows) => string.Join(" ", Enumerable.Range(1, Math.Min(rows, sheet.LastRowUsed()?.RowNumber() ?? 0)).Select(row => RowText(sheet, row)));
    private static string RowText(IXLWorksheet sheet, int row) => string.Join(" | ", sheet.Row(row).CellsUsed().Select(cell => cell.GetFormattedString().Trim()).Where(value => value.Length > 0));
    private static bool IsColumnHeader(string text) =>
        (text.Contains("date", StringComparison.OrdinalIgnoreCase) || text.Contains("effective", StringComparison.OrdinalIgnoreCase)) &&
        (text.Contains("description", StringComparison.OrdinalIgnoreCase) || text.Contains("details", StringComparison.OrdinalIgnoreCase) || text.Contains("type of change", StringComparison.OrdinalIgnoreCase));
    private static string RemoveDate(string text) => Regex.Replace(OrdinalDate.Replace(text, "${day}"), @"\b\d{4}-\d{1,2}-\d{1,2}\b|\b\d{1,2}([/-])\d{1,2}\1\d{2,4}\b|\b\d{1,2}\s+[A-Za-z]+\s+\d{4}\b|\b[A-Za-z]+\s+\d{1,2},?\s+\d{4}\b", "", RegexOptions.IgnoreCase).Trim(' ', '|', '-', ':');
    private static string FirstSentence(string text) => text.Split(['.', ';', '|'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).FirstOrDefault() ?? text;
    private static string BenefitCategory(string text) => new[] { "Health", "Dental", "Life", "Disability", "Vision", "HCSA", "EAP" }.FirstOrDefault(item => text.Contains(item, StringComparison.OrdinalIgnoreCase)) ?? "";
}

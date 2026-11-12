using System.Globalization;
using System.Text;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed class BillingCsvWriter : IManulifeBillingCsvExporter
{
    public BillingCsvExportResult Write(BillingExtractionResult result, string outputFolder)
    {
        Directory.CreateDirectory(outputFolder);

        var export = new BillingCsvExportResult
        {
            OutputFolder = outputFolder,
            ClassesDetectedPath = Path.Combine(outputFolder, "billing_classes_detected.csv"),
            ReconciliationResultsPath = Path.Combine(outputFolder, "billing_reconciliation_results.csv"),
            ValidationErrorsPath = Path.Combine(outputFolder, "billing_validation_errors.csv")
        };

        WriteClassesDetected(result, export.ClassesDetectedPath);
        WriteReconciliationResults(result, export.ReconciliationResultsPath);
        WriteValidationErrors(result, export.ValidationErrorsPath);

        result.CsvExport = export;
        return export;
    }

    private static void WriteClassesDetected(BillingExtractionResult result, string path)
    {
        var rows = new List<string[]>
        {
            new[] { "province", "benefit", "class_code", "option_label", "lives", "monthly_premium" }
        };

        rows.AddRange(result.OptionSummaries.Select(row => new[]
        {
            row.Province,
            row.Benefit,
            row.ClassCode,
            row.OptionLabel,
            row.Lives.ToString(CultureInfo.InvariantCulture),
            FormatDecimal(row.MonthlyPremium)
        }));

        WriteRows(path, rows);
    }

    private static void WriteReconciliationResults(BillingExtractionResult result, string path)
    {
        var rows = new List<string[]>
        {
            new[]
            {
                "client_id",
                "billing_period",
                "source_file",
                "reconciliation_level",
                "province",
                "benefit",
                "class_code",
                "pdf_lives",
                "calculated_lives",
                "lives_difference",
                "lives_status",
                "lives_warning_message",
                "pdf_premium",
                "calculated_premium",
                "premium_difference",
                "pdf_tax",
                "calculated_tax",
                "tax_difference",
                "pdf_total_with_tax",
                "calculated_total_with_tax",
                "total_difference",
                "status",
                "warning_message",
                "source_page",
                "created_at"
            }
        };

        rows.AddRange(result.ReconciliationResults.Select(row => new[]
        {
            row.ClientId.ToString(CultureInfo.InvariantCulture),
            row.BillingPeriod,
            row.SourceFile,
            row.ReconciliationLevel,
            row.Province,
            row.Benefit,
            row.ClassCode,
            FormatInt(row.PdfLives),
            FormatInt(row.CalculatedLives),
            FormatInt(row.LivesDifference),
            row.LivesStatus,
            row.LivesWarningMessage,
            FormatDecimal(row.PdfPremium),
            FormatDecimal(row.CalculatedPremium),
            FormatDecimal(row.PremiumDifference),
            FormatDecimal(row.PdfTax),
            FormatDecimal(row.CalculatedTax),
            FormatDecimal(row.TaxDifference),
            FormatDecimal(row.PdfTotalWithTax),
            FormatDecimal(row.CalculatedTotalWithTax),
            FormatDecimal(row.TotalDifference),
            row.Status,
            row.WarningMessage,
            FormatInt(row.SourcePage),
            row.CreatedAt.ToString("O", CultureInfo.InvariantCulture)
        }));

        WriteRows(path, rows);
    }

    private static void WriteValidationErrors(BillingExtractionResult result, string path)
    {
        var rows = new List<string[]>
        {
            new[]
            {
                "billing_period",
                "reconciliation_level",
                "province",
                "benefit",
                "class_code",
                "pdf_lives",
                "calculated_lives",
                "lives_difference",
                "lives_status",
                "lives_warning_message",
                "pdf_premium",
                "calculated_premium",
                "premium_difference",
                "pdf_tax",
                "calculated_tax",
                "tax_difference",
                "pdf_total_with_tax",
                "calculated_total_with_tax",
                "total_difference",
                "status",
                "warning_message",
                "source_page"
            }
        };

        rows.AddRange(result.ReconciliationResults
            .Where(row =>
                row.Status.Equals(BillingReconciliationStatuses.Failed, StringComparison.OrdinalIgnoreCase) ||
                row.Status.Equals(BillingReconciliationStatuses.Warning, StringComparison.OrdinalIgnoreCase))
            .Select(row => new[]
            {
                row.BillingPeriod,
                row.ReconciliationLevel,
                row.Province,
                row.Benefit,
                row.ClassCode,
                FormatInt(row.PdfLives),
                FormatInt(row.CalculatedLives),
                FormatInt(row.LivesDifference),
                row.LivesStatus,
                row.LivesWarningMessage,
                FormatDecimal(row.PdfPremium),
                FormatDecimal(row.CalculatedPremium),
                FormatDecimal(row.PremiumDifference),
                FormatDecimal(row.PdfTax),
                FormatDecimal(row.CalculatedTax),
                FormatDecimal(row.TaxDifference),
                FormatDecimal(row.PdfTotalWithTax),
                FormatDecimal(row.CalculatedTotalWithTax),
                FormatDecimal(row.TotalDifference),
                row.Status,
                row.WarningMessage,
                FormatInt(row.SourcePage)
            }));

        WriteRows(path, rows);
    }

    private static void WriteRows(string path, IEnumerable<string[]> rows)
    {
        var builder = new StringBuilder();
        foreach (var row in rows)
        {
            builder.AppendLine(string.Join(",", row.Select(EscapeCsv)));
        }

        File.WriteAllText(path, builder.ToString(), Encoding.UTF8);
    }

    private static string EscapeCsv(string? value)
    {
        var text = value ?? "";
        return text.Contains('"', StringComparison.Ordinal) ||
               text.Contains(',', StringComparison.Ordinal) ||
               text.Contains('\r', StringComparison.Ordinal) ||
               text.Contains('\n', StringComparison.Ordinal)
            ? $"\"{text.Replace("\"", "\"\"", StringComparison.Ordinal)}\""
            : text;
    }

    private static string FormatDecimal(decimal? value)
    {
        return value?.ToString("0.00", CultureInfo.InvariantCulture) ?? "";
    }

    private static string FormatInt(int? value)
    {
        return value?.ToString(CultureInfo.InvariantCulture) ?? "";
    }
}

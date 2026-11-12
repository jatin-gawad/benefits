using System.Globalization;
using ClosedXML.Excel;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed class NetGrossRateScheduleImporter
{
    private static readonly string[] SupportedExtensions = [".xlsx", ".xlsm"];

    public NetGrossRateScheduleImportResult Import(string filePath)
    {
        var errors = new List<string>();
        if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
        {
            errors.Add("Choose an existing Excel rate schedule.");
            return NetGrossRateScheduleImportResult.Failure(errors);
        }

        var extension = Path.GetExtension(filePath);
        if (!SupportedExtensions.Contains(extension, StringComparer.OrdinalIgnoreCase))
        {
            errors.Add("Net and Gross rate schedules must be an .xlsx or .xlsm workbook.");
            return NetGrossRateScheduleImportResult.Failure(errors);
        }

        using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        using var workbook = new XLWorkbook(stream);
        foreach (var worksheet in workbook.Worksheets)
        {
            var detected = TryReadWorksheet(filePath, worksheet);
            if (detected.Schedule is not null || detected.Errors.Count > 0)
            {
                return detected;
            }
        }

        errors.Add("No dual-basis rate table was found. Expected Current, Renewal, and Negotiated sections with Net rates and Gross rates columns.");
        return NetGrossRateScheduleImportResult.Failure(errors);
    }

    private static NetGrossRateScheduleImportResult TryReadWorksheet(string filePath, IXLWorksheet worksheet)
    {
        var used = worksheet.RangeUsed();
        if (used is null)
        {
            return NetGrossRateScheduleImportResult.NotDetected();
        }

        var firstRow = used.RangeAddress.FirstAddress.RowNumber;
        var lastRow = used.RangeAddress.LastAddress.RowNumber;
        var firstColumn = used.RangeAddress.FirstAddress.ColumnNumber;
        var lastColumn = used.RangeAddress.LastAddress.ColumnNumber;

        for (var headerRow = firstRow; headerRow <= Math.Min(lastRow, firstRow + 30); headerRow++)
        {
            var benefitColumn = Enumerable.Range(firstColumn, lastColumn - firstColumn + 1)
                .FirstOrDefault(column => Normalize(worksheet.Cell(headerRow, column).GetString()) == "benefit");
            if (benefitColumn == 0 || headerRow <= firstRow)
            {
                continue;
            }

            var columns = DetectRateColumns(worksheet, headerRow - 1, headerRow, firstColumn, lastColumn);
            if (columns.Count == 0)
            {
                continue;
            }

            var missingColumns = RequiredColumns()
                .Where(required => !columns.ContainsKey(required))
                .Select(ColumnDisplayName)
                .ToList();
            if (missingColumns.Count > 0)
            {
                return NetGrossRateScheduleImportResult.Failure(
                [
                    $"The sheet '{worksheet.Name}' looks like a Net and Gross rate schedule but is missing: {string.Join(", ", missingColumns)}."
                ]);
            }

            var rows = new List<NetGrossRateScheduleRow>();
            var errors = new List<string>();
            for (var rowNumber = headerRow + 1; rowNumber <= lastRow; rowNumber++)
            {
                var sourceBenefit = worksheet.Cell(rowNumber, benefitColumn).GetString().Trim();
                if (string.IsNullOrWhiteSpace(sourceBenefit))
                {
                    continue;
                }

                var values = new Dictionary<RateScheduleColumn, decimal>();
                foreach (var required in RequiredColumns())
                {
                    var cell = worksheet.Cell(rowNumber, columns[required]);
                    if (!TryReadDisplayedRate(cell, out var value))
                    {
                        errors.Add($"{worksheet.Name} row {rowNumber}: {sourceBenefit} is missing a numeric {ColumnDisplayName(required)} value.");
                        continue;
                    }

                    if (value < 0m)
                    {
                        errors.Add($"{worksheet.Name} row {rowNumber}: {sourceBenefit} has a negative {ColumnDisplayName(required)} value.");
                        continue;
                    }

                    values[required] = value;
                }

                if (values.Count != 6)
                {
                    continue;
                }

                var identity = CanonicalIdentity(sourceBenefit);
                rows.Add(new NetGrossRateScheduleRow
                {
                    SourceBenefit = sourceBenefit,
                    CanonicalBenefit = identity.Benefit,
                    TierOrClass = identity.TierOrClass,
                    CurrentNetRate = values[RateScheduleColumn.CurrentNet],
                    CurrentGrossRate = values[RateScheduleColumn.CurrentGross],
                    RenewalNetRate = values[RateScheduleColumn.RenewalNet],
                    RenewalGrossRate = values[RateScheduleColumn.RenewalGross],
                    NegotiatedNetRate = values[RateScheduleColumn.NegotiatedNet],
                    NegotiatedGrossRate = values[RateScheduleColumn.NegotiatedGross],
                    SourceRow = rowNumber
                });
            }

            if (rows.Count == 0)
            {
                errors.Add($"The sheet '{worksheet.Name}' contains the required headers but no complete benefit rate rows.");
            }

            var duplicates = rows
                .GroupBy(row => $"{BenefitKey(row.CanonicalBenefit)}|{TierKey(row.SourceBenefit, row.TierOrClass)}")
                .Where(group => group.Count() > 1)
                .Select(group => string.Join(", ", group.Select(row => $"'{row.SourceBenefit}' (row {row.SourceRow})")));
            foreach (var duplicate in duplicates)
            {
                errors.Add($"The schedule has duplicate benefit/tier mappings: {duplicate}.");
            }

            if (errors.Count > 0)
            {
                return NetGrossRateScheduleImportResult.Failure(errors);
            }

            return NetGrossRateScheduleImportResult.Success(new NetGrossRateSchedule
            {
                SourceFilePath = Path.GetFullPath(filePath),
                SourceFileName = Path.GetFileName(filePath),
                SheetName = worksheet.Name,
                PresentationRateBasis = RateValueBasisNames.Gross,
                Rows = rows
            });
        }

        return NetGrossRateScheduleImportResult.NotDetected();
    }

    private static Dictionary<RateScheduleColumn, int> DetectRateColumns(
        IXLWorksheet worksheet,
        int sectionHeaderRow,
        int rateHeaderRow,
        int firstColumn,
        int lastColumn)
    {
        var result = new Dictionary<RateScheduleColumn, int>();
        var section = "";
        for (var column = firstColumn; column <= lastColumn; column++)
        {
            var sectionText = worksheet.Cell(sectionHeaderRow, column).GetString().Trim();
            if (!string.IsNullOrWhiteSpace(sectionText))
            {
                section = SectionKey(sectionText);
            }

            var rateType = Normalize(worksheet.Cell(rateHeaderRow, column).GetString());
            var key = (section, rateType) switch
            {
                ("current", var type) when type.Contains("netrate", StringComparison.Ordinal) => RateScheduleColumn.CurrentNet,
                ("current", var type) when type.Contains("grossrate", StringComparison.Ordinal) => RateScheduleColumn.CurrentGross,
                ("renewal", var type) when type.Contains("netrate", StringComparison.Ordinal) => RateScheduleColumn.RenewalNet,
                ("renewal", var type) when type.Contains("grossrate", StringComparison.Ordinal) => RateScheduleColumn.RenewalGross,
                ("negotiated", var type) when type.Contains("netrate", StringComparison.Ordinal) => RateScheduleColumn.NegotiatedNet,
                ("negotiated", var type) when type.Contains("grossrate", StringComparison.Ordinal) => RateScheduleColumn.NegotiatedGross,
                _ => (RateScheduleColumn?)null
            };

            if (key.HasValue)
            {
                result[key.Value] = column;
            }
        }

        return result;
    }

    private static bool TryReadDisplayedRate(IXLCell cell, out decimal value)
    {
        var displayed = cell.GetFormattedString().Trim();
        if (decimal.TryParse(displayed, NumberStyles.Number | NumberStyles.AllowExponent, CultureInfo.CurrentCulture, out value) ||
            decimal.TryParse(displayed, NumberStyles.Number | NumberStyles.AllowExponent, CultureInfo.InvariantCulture, out value))
        {
            return true;
        }

        return cell.TryGetValue(out value);
    }

    private static string SectionKey(string value)
    {
        var normalized = Normalize(value);
        if (normalized.Contains("current", StringComparison.Ordinal))
        {
            return "current";
        }

        if (normalized.Contains("negotiated", StringComparison.Ordinal) || normalized.Contains("final", StringComparison.Ordinal))
        {
            return "negotiated";
        }

        if (normalized.Contains("renewal", StringComparison.Ordinal) || normalized.Contains("proposed", StringComparison.Ordinal))
        {
            return "renewal";
        }

        return "";
    }

    private static (string Benefit, string TierOrClass) CanonicalIdentity(string sourceBenefit)
    {
        var normalized = Normalize(sourceBenefit);
        var tier = normalized.Contains("single", StringComparison.Ordinal)
            ? "Single"
            : normalized.Contains("family", StringComparison.Ordinal)
                ? "Family"
                : normalized.Contains("couple", StringComparison.Ordinal)
                    ? "Couple"
                    : "";

        if (normalized.Contains("dependentlife", StringComparison.Ordinal) || normalized.StartsWith("deplife", StringComparison.Ordinal))
        {
            return ("Dependent Life", tier);
        }

        if (normalized.Equals("add", StringComparison.Ordinal) ||
            normalized.Contains("adandd", StringComparison.Ordinal) ||
            normalized.Contains("accidentaldeath", StringComparison.Ordinal))
        {
            return ("AD&D", tier);
        }

        if (normalized.Contains("health", StringComparison.Ordinal) || normalized.Contains("ehc", StringComparison.Ordinal))
        {
            return ("Extended Health Care", tier);
        }

        if (normalized.Contains("dental", StringComparison.Ordinal))
        {
            return ("Dental", tier);
        }

        if (normalized.Contains("life", StringComparison.Ordinal))
        {
            return ("Life", tier);
        }

        return (sourceBenefit.Trim(), tier);
    }

    internal static string BenefitKey(string value)
    {
        return CanonicalIdentity(value).Benefit switch
        {
            "Dependent Life" => "dependentlife",
            "AD&D" => "add",
            "Extended Health Care" => "health",
            "Dental" => "dental",
            "Life" => "life",
            var other => Normalize(other)
        };
    }

    internal static string TierKey(params string[] values)
    {
        var normalized = Normalize(string.Join(" ", values));
        if (normalized.Contains("single", StringComparison.Ordinal) || normalized.Contains("employeeonly", StringComparison.Ordinal))
        {
            return "single";
        }

        if (normalized.Contains("family", StringComparison.Ordinal))
        {
            return "family";
        }

        if (normalized.Contains("couple", StringComparison.Ordinal) || normalized.Contains("employeeandspouse", StringComparison.Ordinal))
        {
            return "couple";
        }

        return "";
    }

    private static string Normalize(string value)
    {
        return new string((value ?? "")
            .Where(char.IsLetterOrDigit)
            .Select(char.ToLowerInvariant)
            .ToArray());
    }

    private static IReadOnlyList<RateScheduleColumn> RequiredColumns()
    {
        return
        [
            RateScheduleColumn.CurrentNet,
            RateScheduleColumn.CurrentGross,
            RateScheduleColumn.RenewalNet,
            RateScheduleColumn.RenewalGross,
            RateScheduleColumn.NegotiatedNet,
            RateScheduleColumn.NegotiatedGross
        ];
    }

    private static string ColumnDisplayName(RateScheduleColumn column)
    {
        return column switch
        {
            RateScheduleColumn.CurrentNet => "Current Net rate",
            RateScheduleColumn.CurrentGross => "Current Gross rate",
            RateScheduleColumn.RenewalNet => "Renewal Net rate",
            RateScheduleColumn.RenewalGross => "Renewal Gross rate",
            RateScheduleColumn.NegotiatedNet => "Negotiated Net rate",
            RateScheduleColumn.NegotiatedGross => "Negotiated Gross rate",
            _ => "rate"
        };
    }

    private enum RateScheduleColumn
    {
        CurrentNet,
        CurrentGross,
        RenewalNet,
        RenewalGross,
        NegotiatedNet,
        NegotiatedGross
    }
}

public sealed class NetGrossRateScheduleImportResult
{
    public NetGrossRateSchedule? Schedule { get; private init; }
    public IReadOnlyList<string> Errors { get; private init; } = [];
    public bool IsDetected => Schedule is not null || Errors.Count > 0;
    public bool Succeeded => Schedule is not null && Errors.Count == 0;

    public static NetGrossRateScheduleImportResult Success(NetGrossRateSchedule schedule) => new() { Schedule = schedule };
    public static NetGrossRateScheduleImportResult Failure(IReadOnlyList<string> errors) => new() { Errors = errors };
    public static NetGrossRateScheduleImportResult NotDetected() => new();
}

public sealed class NetGrossRateSchedule
{
    public string SourceFilePath { get; init; } = "";
    public string SourceFileName { get; init; } = "";
    public string SheetName { get; init; } = "";
    public string PresentationRateBasis { get; init; } = RateValueBasisNames.Gross;
    public IReadOnlyList<NetGrossRateScheduleRow> Rows { get; init; } = [];
}

public sealed class NetGrossRateScheduleRow
{
    public string SourceBenefit { get; init; } = "";
    public string CanonicalBenefit { get; init; } = "";
    public string TierOrClass { get; init; } = "";
    public decimal CurrentNetRate { get; init; }
    public decimal CurrentGrossRate { get; init; }
    public decimal RenewalNetRate { get; init; }
    public decimal RenewalGrossRate { get; init; }
    public decimal NegotiatedNetRate { get; init; }
    public decimal NegotiatedGrossRate { get; init; }
    public int SourceRow { get; init; }
}

using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using ExcelDataReader;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed class LegacyClaimsServiceWorkbookReader
{
    private static readonly Regex ContractNumberRegex = new(
        @"(?:Contract|Policy)\s+Number\s*:\s*(?<value>[A-Za-z0-9-]+)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public IReadOnlyList<ClaimsDistributionRow> Extract(string path)
    {
        if (!File.Exists(path) || !Path.GetExtension(path).Equals(".xls", StringComparison.OrdinalIgnoreCase))
        {
            return [];
        }

        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        using var reader = ExcelReaderFactory.CreateReader(stream);

        do
        {
            var rows = ReadRows(reader);
            var extracted = ExtractSheet(path, reader.Name, rows);
            if (extracted.Count > 0)
            {
                return extracted;
            }
        }
        while (reader.NextResult());

        return [];
    }

    private static List<ClaimsDistributionRow> ExtractSheet(
        string path,
        string sheetName,
        IReadOnlyList<LegacyWorkbookRow> rows)
    {
        var titleRow = rows
            .Select((row, index) => new { Row = row, Index = index })
            .FirstOrDefault(item => item.Row.Text.Contains("Claims Paid by Type of Service", StringComparison.OrdinalIgnoreCase));
        if (titleRow is null)
        {
            return [];
        }

        var title = titleRow.Row.Text;
        var benefit = title.Contains("Dental", StringComparison.OrdinalIgnoreCase) ? "Dental" :
            title.Contains("Health", StringComparison.OrdinalIgnoreCase) || title.Contains("Medical", StringComparison.OrdinalIgnoreCase)
                ? "Medical"
                : "";
        if (string.IsNullOrWhiteSpace(benefit))
        {
            return [];
        }

        var periodHeaderRow = rows
            .Select((row, index) => new { Row = row, Index = index })
            .Skip(titleRow.Index)
            .FirstOrDefault(item =>
                item.Row.Cells.Any(cell => CellText(cell).Equals("Current Period", StringComparison.OrdinalIgnoreCase)) &&
                item.Row.Cells.Any(cell => CellText(cell).Equals("Prior Period", StringComparison.OrdinalIgnoreCase)));
        if (periodHeaderRow is null)
        {
            return [];
        }

        var currentStartColumn = FindCell(periodHeaderRow.Row, "Current Period");
        var priorStartColumn = FindCell(periodHeaderRow.Row, "Prior Period");
        var periodRowIndex = Enumerable.Range(periodHeaderRow.Index + 1, Math.Min(4, rows.Count - periodHeaderRow.Index - 1))
            .FirstOrDefault(index =>
                TryParseDateRange(CellText(rows[index].Cell(currentStartColumn)), out _, out _) &&
                TryParseDateRange(CellText(rows[index].Cell(priorStartColumn)), out _, out _), -1);
        if (periodRowIndex < 0 ||
            !TryParseDateRange(CellText(rows[periodRowIndex].Cell(currentStartColumn)), out var currentStart, out var currentEnd) ||
            !TryParseDateRange(CellText(rows[periodRowIndex].Cell(priorStartColumn)), out var priorStart, out var priorEnd))
        {
            return [];
        }

        var currentPaidColumn = FindPaidColumn(rows, periodRowIndex + 1, currentStartColumn, priorStartColumn);
        var priorPaidColumn = FindPaidColumn(rows, periodRowIndex + 1, priorStartColumn, rows.Max(row => row.Cells.Count));
        if (currentPaidColumn < 0 || priorPaidColumn < 0)
        {
            return [];
        }

        var summaryStart = rows
            .Select((row, index) => new { Row = row, Index = index })
            .Skip(periodRowIndex)
            .FirstOrDefault(item =>
                CellText(item.Row.Cell(0)).StartsWith("Relation", StringComparison.OrdinalIgnoreCase) &&
                CellText(item.Row.Cell(0)).Contains("All", StringComparison.OrdinalIgnoreCase));
        if (summaryStart is null)
        {
            return [];
        }

        var clientName = ReadClientName(rows, titleRow.Index);
        var policyNumber = ReadPolicyNumber(rows);
        var carrier = DetectCarrier(rows);
        var reportDate = ReadReportDate(rows, titleRow.Index);
        var currentPeriod = PeriodLabel(currentStart, currentEnd);
        var priorPeriod = PeriodLabel(priorStart, priorEnd);
        var sourceReference = $"{carrier} Excel {sheetName}!A{summaryStart.Index + 1}:J";

        var sectionRows = new List<ClaimsDistributionRow>();
        var paramedicalRows = new List<ClaimsDistributionRow>();
        SectionAccumulator? section = null;

        for (var index = summaryStart.Index + 1; index < rows.Count; index++)
        {
            var row = rows[index];
            var label = CellText(row.Cell(0));
            var currentPaid = CellDecimal(row.Cell(currentPaidColumn));
            var priorPaid = CellDecimal(row.Cell(priorPaidColumn));
            var currentPercent = CellDecimal(row.Cell(currentPaidColumn + 1));
            var priorPercent = CellDecimal(row.Cell(priorPaidColumn + 1));

            if (label.Contains("Canadian Dollars Total Claims", StringComparison.OrdinalIgnoreCase))
            {
                AddSection(sectionRows, section, benefit, currentPeriod, priorPeriod, currentStart, currentEnd, priorStart, priorEnd,
                    carrier, policyNumber, clientName, reportDate, path, $"{sourceReference}{index + 1}");
                sectionRows.Add(CreateRow(
                    benefit,
                    "Total",
                    "Total",
                    currentPaid,
                    currentPercent,
                    priorPaid,
                    priorPercent,
                    currentPeriod,
                    priorPeriod,
                    currentStart,
                    currentEnd,
                    priorStart,
                    priorEnd,
                    carrier,
                    policyNumber,
                    clientName,
                    reportDate,
                    path,
                    $"{sourceReference}{index + 1}"));
                break;
            }

            if (!string.IsNullOrWhiteSpace(label) && !currentPaid.HasValue && !priorPaid.HasValue)
            {
                AddSection(sectionRows, section, benefit, currentPeriod, priorPeriod, currentStart, currentEnd, priorStart, priorEnd,
                    carrier, policyNumber, clientName, reportDate, path, $"{sourceReference}{index}");
                section = new SectionAccumulator(label);
                continue;
            }

            if (section is null || (!currentPaid.HasValue && !priorPaid.HasValue))
            {
                continue;
            }

            if (string.IsNullOrWhiteSpace(label))
            {
                section.SetSubtotal(currentPaid, currentPercent, priorPaid, priorPercent);
                AddSection(sectionRows, section, benefit, currentPeriod, priorPeriod, currentStart, currentEnd, priorStart, priorEnd,
                    carrier, policyNumber, clientName, reportDate, path, $"{sourceReference}{index + 1}");
                section = null;
                continue;
            }

            section.AddDetail(currentPaid, currentPercent, priorPaid, priorPercent);
            if (NormalizeSectionCategory(benefit, section.Name).Equals("Paramedical", StringComparison.OrdinalIgnoreCase))
            {
                paramedicalRows.Add(CreateRow(
                    "Paramedical Detail",
                    "Paramedical",
                    label,
                    currentPaid,
                    currentPercent,
                    priorPaid,
                    priorPercent,
                    currentPeriod,
                    priorPeriod,
                    currentStart,
                    currentEnd,
                    priorStart,
                    priorEnd,
                    carrier,
                    policyNumber,
                    clientName,
                    reportDate,
                    path,
                    $"{sourceReference}{index + 1}"));
            }
        }

        return AggregateSections(sectionRows)
            .Concat(paramedicalRows)
            .ToList();
    }

    private static List<LegacyWorkbookRow> ReadRows(IExcelDataReader reader)
    {
        var rows = new List<LegacyWorkbookRow>();
        while (reader.Read())
        {
            var cells = new object?[reader.FieldCount];
            for (var column = 0; column < reader.FieldCount; column++)
            {
                cells[column] = reader.GetValue(column);
            }

            rows.Add(new LegacyWorkbookRow(cells));
        }

        return rows;
    }

    private static void AddSection(
        List<ClaimsDistributionRow> rows,
        SectionAccumulator? section,
        string benefit,
        string currentPeriod,
        string priorPeriod,
        DateOnly currentStart,
        DateOnly currentEnd,
        DateOnly priorStart,
        DateOnly priorEnd,
        string carrier,
        string policyNumber,
        string clientName,
        DateOnly? reportDate,
        string path,
        string sourceReference)
    {
        if (section is null || !section.HasAmounts)
        {
            return;
        }

        rows.Add(CreateRow(
            benefit,
            NormalizeSectionCategory(benefit, section.Name),
            section.Name,
            section.CurrentPaid,
            section.CurrentPercent,
            section.PriorPaid,
            section.PriorPercent,
            currentPeriod,
            priorPeriod,
            currentStart,
            currentEnd,
            priorStart,
            priorEnd,
            carrier,
            policyNumber,
            clientName,
            reportDate,
            path,
            sourceReference));
    }

    private static ClaimsDistributionRow CreateRow(
        string benefit,
        string category,
        string original,
        decimal? currentPaid,
        decimal? currentPercent,
        decimal? priorPaid,
        decimal? priorPercent,
        string currentPeriod,
        string priorPeriod,
        DateOnly currentStart,
        DateOnly currentEnd,
        DateOnly priorStart,
        DateOnly priorEnd,
        string carrier,
        string policyNumber,
        string clientName,
        DateOnly? reportDate,
        string path,
        string sourceReference)
    {
        return new ClaimsDistributionRow
        {
            CarrierName = carrier,
            PolicyNumber = policyNumber,
            ClientName = clientName,
            ReportPeriod = currentPeriod,
            PeriodStart = DateLabel(currentStart),
            PeriodEnd = DateLabel(currentEnd),
            PriorReportPeriod = priorPeriod,
            PriorPeriodStart = DateLabel(priorStart),
            PriorPeriodEnd = DateLabel(priorEnd),
            Benefit = benefit,
            Division = "ALL",
            Class = "ALL",
            ReportCreatedDate = reportDate?.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) ?? "",
            SourceFile = Path.GetFileName(path),
            ServiceCategory = category,
            ServiceDescriptionOriginal = original,
            AmountPaid = currentPaid,
            PercentPaid = currentPercent,
            PriorAmountPaid = priorPaid,
            PriorPercentPaid = priorPercent,
            SourceReference = sourceReference
        };
    }

    private static IReadOnlyList<ClaimsDistributionRow> AggregateSections(IReadOnlyList<ClaimsDistributionRow> rows)
    {
        return rows
            .GroupBy(
                row => $"{row.Benefit}\u001F{row.ServiceCategory}",
                StringComparer.OrdinalIgnoreCase)
            .Select(group =>
            {
                var first = group.First();
                if (group.Count() == 1 || first.ServiceCategory.Equals("Total", StringComparison.OrdinalIgnoreCase))
                {
                    return first;
                }

                first.ServiceDescriptionOriginal = string.Join("; ", group
                    .Select(row => row.ServiceDescriptionOriginal)
                    .Where(value => !string.IsNullOrWhiteSpace(value))
                    .Distinct(StringComparer.OrdinalIgnoreCase));
                first.AmountPaid = group.Sum(row => row.AmountPaid ?? 0m);
                first.PercentPaid = group.Sum(row => row.PercentPaid ?? 0m);
                first.PriorAmountPaid = group.Sum(row => row.PriorAmountPaid ?? 0m);
                first.PriorPercentPaid = group.Sum(row => row.PriorPercentPaid ?? 0m);
                first.SourceReference = string.Join("; ", group.Select(row => row.SourceReference).Distinct(StringComparer.OrdinalIgnoreCase));
                return first;
            })
            .ToList();
    }

    private static string NormalizeSectionCategory(string benefit, string section)
    {
        var normalized = Regex.Replace(ParsingHelpers.Clean(section), @"[^A-Za-z0-9]", "").ToLowerInvariant();
        if (benefit.Equals("Dental", StringComparison.OrdinalIgnoreCase))
        {
            return normalized switch
            {
                "diagnosticpreventive" => "Preventive & Diagnostic",
                "majorrestorative" => "Major Restorative",
                "orthodontic" or "orthodontics" => "Orthodontic",
                "periodonticsendodontics" => "Periodontal & Endodontic",
                "surgicalrestorative" => "Surgical & Restorative",
                _ => ClaimsDistributionExtractor.NormalizeServiceCategory(section)
            };
        }

        return normalized switch
        {
            "hospitalincanada" => "Hospital",
            "paydirectdrugs" or "reimbursementdrugs" => "Drugs",
            "medicalservicesequipment" => "Medical Services and Supplies",
            "paramedicalpractitioners" => "Paramedical",
            "contactlensesoreyeglasses" => "Vision",
            "healthspendingaccount" or "miscellaneous" => "Miscellaneous Services",
            _ => ClaimsDistributionExtractor.NormalizeServiceCategory(section)
        };
    }

    private static int FindPaidColumn(IReadOnlyList<LegacyWorkbookRow> rows, int startRow, int startColumn, int endColumn)
    {
        for (var rowIndex = startRow; rowIndex < Math.Min(startRow + 5, rows.Count); rowIndex++)
        {
            for (var column = startColumn; column < Math.Min(endColumn, rows[rowIndex].Cells.Count); column++)
            {
                if (CellText(rows[rowIndex].Cell(column)).Equals("Paid", StringComparison.OrdinalIgnoreCase))
                {
                    return column;
                }
            }
        }

        return -1;
    }

    private static int FindCell(LegacyWorkbookRow row, string value)
    {
        for (var index = 0; index < row.Cells.Count; index++)
        {
            if (CellText(row.Cell(index)).Equals(value, StringComparison.OrdinalIgnoreCase))
            {
                return index;
            }
        }

        return -1;
    }

    private static string ReadClientName(IReadOnlyList<LegacyWorkbookRow> rows, int titleRow)
    {
        for (var index = titleRow + 1; index < Math.Min(titleRow + 5, rows.Count); index++)
        {
            var value = CellText(rows[index].Cell(0));
            if (!string.IsNullOrWhiteSpace(value) &&
                !value.StartsWith("Contract", StringComparison.OrdinalIgnoreCase) &&
                !value.StartsWith("Policy", StringComparison.OrdinalIgnoreCase) &&
                !value.Equals("Summary", StringComparison.OrdinalIgnoreCase))
            {
                return value;
            }
        }

        return "";
    }

    private static string ReadPolicyNumber(IEnumerable<LegacyWorkbookRow> rows)
    {
        foreach (var row in rows)
        {
            var match = ContractNumberRegex.Match(row.Text);
            if (match.Success)
            {
                return match.Groups["value"].Value;
            }
        }

        return "";
    }

    private static string DetectCarrier(IEnumerable<LegacyWorkbookRow> rows)
    {
        var text = string.Join(" ", rows.Select(row => row.Text));
        var carriers = new[]
        {
            "Sun Life", "Canada Life", "Great-West Life", "Manulife", "Equitable", "RBC", "GreenShield",
            "GroupHEALTH", "Desjardins", "Beneva", "Empire Life", "Industrial Alliance", "iA Financial"
        };
        return carriers.FirstOrDefault(carrier => text.Contains(carrier, StringComparison.OrdinalIgnoreCase)) ?? "Carrier";
    }

    private static DateOnly? ReadReportDate(IReadOnlyList<LegacyWorkbookRow> rows, int titleRow)
    {
        for (var index = 0; index < titleRow; index++)
        {
            if (TryParseDate(CellText(rows[index].Cell(0)), out var date))
            {
                return date;
            }
        }

        return null;
    }

    private static bool TryParseDateRange(string value, out DateOnly start, out DateOnly end)
    {
        start = default;
        end = default;
        var parts = Regex.Split(ParsingHelpers.Clean(value), @"\s+to\s+", RegexOptions.IgnoreCase);
        return parts.Length == 2 && TryParseDate(parts[0], out start) && TryParseDate(parts[1], out end);
    }

    private static bool TryParseDate(string value, out DateOnly date)
    {
        date = default;
        return DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out var parsed) &&
               (date = DateOnly.FromDateTime(parsed)) != default;
    }

    private static string PeriodLabel(DateOnly start, DateOnly end) => $"{DateLabel(start)} to {DateLabel(end)}";

    private static string DateLabel(DateOnly date) => date.ToString("MMMM d, yyyy", CultureInfo.InvariantCulture);

    private static decimal? CellDecimal(object? value)
    {
        return value switch
        {
            null => null,
            decimal decimalValue => decimalValue,
            double doubleValue => Convert.ToDecimal(doubleValue, CultureInfo.InvariantCulture),
            float floatValue => Convert.ToDecimal(floatValue, CultureInfo.InvariantCulture),
            int intValue => intValue,
            long longValue => longValue,
            _ => ParsingHelpers.ParseDecimal(CellText(value))
        };
    }

    private static string CellText(object? value)
    {
        return value switch
        {
            null => "",
            DateTime date => date.ToString("MMMM d, yyyy", CultureInfo.InvariantCulture),
            _ => ParsingHelpers.Clean(Convert.ToString(value, CultureInfo.InvariantCulture) ?? "")
        };
    }

    private sealed record LegacyWorkbookRow(IReadOnlyList<object?> Cells)
    {
        public object? Cell(int index) => index >= 0 && index < Cells.Count ? Cells[index] : null;

        public string Text => string.Join(" ", Cells.Select(CellText).Where(value => !string.IsNullOrWhiteSpace(value)));
    }

    private sealed class SectionAccumulator(string name)
    {
        private decimal? _subtotalCurrentPaid;
        private decimal? _subtotalCurrentPercent;
        private decimal? _subtotalPriorPaid;
        private decimal? _subtotalPriorPercent;
        private decimal _detailCurrentPaid;
        private decimal _detailCurrentPercent;
        private decimal _detailPriorPaid;
        private decimal _detailPriorPercent;
        private bool _hasDetails;

        public string Name { get; } = name;
        public bool HasAmounts => _subtotalCurrentPaid.HasValue || _subtotalPriorPaid.HasValue || _hasDetails;
        public decimal? CurrentPaid => _subtotalCurrentPaid ?? (_hasDetails ? _detailCurrentPaid : null);
        public decimal? CurrentPercent => _subtotalCurrentPercent ?? (_hasDetails ? _detailCurrentPercent : null);
        public decimal? PriorPaid => _subtotalPriorPaid ?? (_hasDetails ? _detailPriorPaid : null);
        public decimal? PriorPercent => _subtotalPriorPercent ?? (_hasDetails ? _detailPriorPercent : null);

        public void AddDetail(decimal? currentPaid, decimal? currentPercent, decimal? priorPaid, decimal? priorPercent)
        {
            _detailCurrentPaid += currentPaid ?? 0m;
            _detailCurrentPercent += currentPercent ?? 0m;
            _detailPriorPaid += priorPaid ?? 0m;
            _detailPriorPercent += priorPercent ?? 0m;
            _hasDetails = true;
        }

        public void SetSubtotal(decimal? currentPaid, decimal? currentPercent, decimal? priorPaid, decimal? priorPercent)
        {
            _subtotalCurrentPaid = currentPaid;
            _subtotalCurrentPercent = currentPercent;
            _subtotalPriorPaid = priorPaid;
            _subtotalPriorPercent = priorPercent;
        }
    }
}

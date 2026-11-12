using System.Globalization;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed class AsoFinancialActivityWorkbookParser
{
    private static readonly Regex PeriodRegex = new(
        @"(?<start>\d{1,2}/\d{1,2}/\d{4})\s+to\s+(?<end>\d{1,2}/\d{1,2}/\d{4})",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    public bool LooksLikeAsoFinancialActivityWorkbook(string filePath)
    {
        if (!File.Exists(filePath) || !IsWorkbook(filePath))
        {
            return false;
        }

        if (Path.GetFileName(filePath).Contains("ASOActivity", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        try
        {
            using var workbook = new XLWorkbook(filePath);
            return workbook.Worksheets.Any(HasFinancialActivitySignature);
        }
        catch
        {
            return false;
        }
    }

    public AsoFinancialActivityStatement Parse(string filePath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);
        if (!File.Exists(filePath))
        {
            throw new FileNotFoundException("ASO financial activity workbook was not found.", filePath);
        }

        using var workbook = new XLWorkbook(filePath);
        var worksheet = workbook.Worksheets.FirstOrDefault(HasFinancialActivitySignature)
            ?? throw new InvalidDataException("No 'Financial Summary of Activity' worksheet was found.");

        var headerRow = FindHeaderRow(worksheet);
        if (headerRow <= 0)
        {
            throw new InvalidDataException("The ASO activity header row could not be identified.");
        }

        var statement = new AsoFinancialActivityStatement
        {
            ClientName = FindClientName(worksheet),
            ContractNumber = ReadMetadata(worksheet, "Contract"),
            Carrier = FindCarrier(worksheet),
            BillingArrangement = ReadMetadata(worksheet, "Billing Arrangement"),
            Divisions = ReadMetadata(worksheet, "Division"),
            Benefits = ReadMetadata(worksheet, "Benefit"),
            SourceFilePath = Path.GetFullPath(filePath),
            SourceFileName = Path.GetFileName(filePath),
            SourceSheet = worksheet.Name,
            ImportedAt = DateTime.Now
        };

        ReadPeriodRange(worksheet, statement);
        var columns = BuildColumnMap(worksheet, headerRow);
        var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? headerRow;
        var displayOrder = 0;

        for (var rowNumber = headerRow + 1; rowNumber <= lastRow; rowNumber++)
        {
            var label = ReadText(worksheet.Cell(rowNumber, Column(columns, "month")));
            if (string.IsNullOrWhiteSpace(label))
            {
                continue;
            }

            var isTotal = label.Equals("Total", StringComparison.OrdinalIgnoreCase);
            var activity = ParseRow(worksheet, rowNumber, label, isTotal, columns, displayOrder++);
            if (isTotal)
            {
                statement.Total = activity;
                break;
            }

            if (activity.ActivityMonth is not null)
            {
                statement.Periods.Add(activity);
            }
        }

        Validate(statement);
        statement.ImportStatus = statement.ValidationIssues.Count == 0
            ? AsoFinancialImportStatuses.Complete
            : AsoFinancialImportStatuses.NeedsReview;
        statement.ValidationSummary = string.Join(" | ", statement.ValidationIssues);
        return statement;
    }

    private static AsoFinancialActivityPeriod ParseRow(
        IXLWorksheet worksheet,
        int row,
        string label,
        bool isTotal,
        IReadOnlyDictionary<string, int> columns,
        int displayOrder)
    {
        var period = new AsoFinancialActivityPeriod
        {
            PeriodLabel = label,
            ActivityMonth = isTotal ? null : ParseMonth(worksheet.Cell(row, Column(columns, "month")), label),
            IsTotal = isTotal,
            DisplayOrder = displayOrder,
            OpeningBalance = Number(worksheet, row, columns, "openingbalancesurplusdeficit"),
            Transfers = Number(worksheet, row, columns, "transfers"),
            DeficitPaymentRefund = Number(worksheet, row, columns, "deficitpaymentrefund"),
            TotalDeposit = Number(worksheet, row, columns, "totaldeposit"),
            OrstDeposit = Number(worksheet, row, columns, "orstdeposit"),
            QstDeposit = Number(worksheet, row, columns, "qstdeposit"),
            HstDeposit = Number(worksheet, row, columns, "hstdeposit"),
            TvqDeposit = Number(worksheet, row, columns, "tvqdeposit"),
            SkpstDeposit = Number(worksheet, row, columns, "skpstdeposit"),
            NetDeposit = Number(worksheet, row, columns, "netdeposit"),
            UnpooledBenefitPayments = Number(worksheet, row, columns, "unpooledbenefitpayments"),
            PooledBenefitPayments = Number(worksheet, row, columns, "pooledbenefitpayments"),
            TotalBenefitPayment = Number(worksheet, row, columns, "totalbenefitpayment"),
            LossRatio = Number(worksheet, row, columns, "lossratio"),
            PoolingCharges = Number(worksheet, row, columns, "poolingcharges"),
            AdminFees = Number(worksheet, row, columns, "adminfees"),
            OtherCharges = Number(worksheet, row, columns, "othercharges"),
            TotalTaxes = Number(worksheet, row, columns, "totaltaxes"),
            Gst = Number(worksheet, row, columns, "gst"),
            Hst = Number(worksheet, row, columns, "hst"),
            Tvq = Number(worksheet, row, columns, "tvq"),
            InsuredPremiumTax = Number(worksheet, row, columns, "insuredpremiumtax"),
            PremiumTaxQuebec = Number(worksheet, row, columns, "premiumtaxquebec"),
            PremiumTaxOntario = Number(worksheet, row, columns, "premiumtaxontario"),
            PremiumTaxNewfoundland = Number(worksheet, row, columns, "premiumtaxesnewfoundland"),
            RetroactiveTaxQst = Number(worksheet, row, columns, "retroactivetaxqst"),
            RetroactiveTaxOrst = Number(worksheet, row, columns, "retroactivetaxorst"),
            RetroactiveTaxSkpst = Number(worksheet, row, columns, "retroactivetaxskpst"),
            InterestCreditCharge = Number(worksheet, row, columns, "interestcreditcharge"),
            TotalCharges = Number(worksheet, row, columns, "totalcharges"),
            SubTotal = Number(worksheet, row, columns, "subtotal"),
            RequiredFloat = Number(worksheet, row, columns, "requiredfloat"),
            CurrentBalance = Number(worksheet, row, columns, "currentbalance"),
            BalanceStatus = Text(worksheet, row, columns, "status")
        };

        if (!isTotal && string.IsNullOrWhiteSpace(period.BalanceStatus))
        {
            period.BalanceStatus = period.CurrentBalance switch
            {
                < 0m => "Surplus",
                > 0m => "Deficit",
                0m => "Balanced",
                _ => ""
            };
        }

        return period;
    }

    private static void Validate(AsoFinancialActivityStatement statement)
    {
        if (string.IsNullOrWhiteSpace(statement.ClientName))
        {
            statement.ValidationIssues.Add("Client name was not detected.");
        }

        if (string.IsNullOrWhiteSpace(statement.ContractNumber))
        {
            statement.ValidationIssues.Add("Contract number was not detected.");
        }

        if (statement.Periods.Count == 0)
        {
            statement.ValidationIssues.Add("No monthly activity rows were detected.");
        }

        if (statement.Total is null)
        {
            statement.ValidationIssues.Add("The workbook total row was not detected.");
            return;
        }

        ValidateTotal(statement, p => p.TotalBenefitPayment, statement.Total.TotalBenefitPayment, "total benefit payments");
        ValidateTotal(statement, p => p.PoolingCharges, statement.Total.PoolingCharges, "pooling charges");
        ValidateTotal(statement, p => p.AdminFees, statement.Total.AdminFees, "admin fees");
        ValidateTotal(statement, p => p.TotalTaxes, statement.Total.TotalTaxes, "total taxes");
        ValidateTotal(statement, p => p.TotalCharges, statement.Total.TotalCharges, "total charges");

        foreach (var period in statement.Periods)
        {
            if (period.TotalBenefitPayment is not null &&
                period.UnpooledBenefitPayments is not null &&
                period.PooledBenefitPayments is not null &&
                Math.Abs(period.TotalBenefitPayment.Value -
                         (period.UnpooledBenefitPayments.Value + period.PooledBenefitPayments.Value)) > 0.02m)
            {
                statement.ValidationIssues.Add($"{period.PeriodLabel}: benefit payment components do not reconcile.");
            }

            var expectedStatus = period.CurrentBalance switch
            {
                < 0m => "Surplus",
                > 0m => "Deficit",
                0m => "Balanced",
                _ => ""
            };
            if (!string.IsNullOrWhiteSpace(expectedStatus) &&
                !string.IsNullOrWhiteSpace(period.BalanceStatus) &&
                !period.BalanceStatus.Equals(expectedStatus, StringComparison.OrdinalIgnoreCase))
            {
                statement.ValidationIssues.Add(
                    $"{period.PeriodLabel}: balance sign indicates {expectedStatus}, but the source status is {period.BalanceStatus}.");
            }
        }
    }

    private static void ValidateTotal(
        AsoFinancialActivityStatement statement,
        Func<AsoFinancialActivityPeriod, decimal?> selector,
        decimal? reported,
        string label)
    {
        if (reported is null)
        {
            return;
        }

        var calculated = statement.Periods.Sum(period => selector(period) ?? 0m);
        if (Math.Abs(calculated - reported.Value) > 0.05m)
        {
            statement.ValidationIssues.Add(
                $"{label}: monthly sum {calculated:C2} does not match the reported total {reported.Value:C2}.");
        }
    }

    private static bool HasFinancialActivitySignature(IXLWorksheet worksheet)
    {
        var used = worksheet.RangeUsed();
        if (used is null)
        {
            return false;
        }

        var sample = used.Cells()
            .Take(400)
            .Select(ReadText)
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .ToList();
        return sample.Any(value => value.Contains("Financial Summary of Activity", StringComparison.OrdinalIgnoreCase)) &&
               sample.Any(value => value.Contains("Total Benefit Payment", StringComparison.OrdinalIgnoreCase)) &&
               sample.Any(value => value.Contains("Current Balance", StringComparison.OrdinalIgnoreCase));
    }

    private static int FindHeaderRow(IXLWorksheet worksheet)
    {
        var lastRow = Math.Min(worksheet.LastRowUsed()?.RowNumber() ?? 0, 60);
        for (var row = 1; row <= lastRow; row++)
        {
            var values = worksheet.Row(row).CellsUsed().Select(ReadText).ToList();
            if (values.Any(value => value.Equals("Month", StringComparison.OrdinalIgnoreCase)) &&
                values.Any(value => value.Contains("Total Benefit Payment", StringComparison.OrdinalIgnoreCase)) &&
                values.Any(value => value.Contains("Current Balance", StringComparison.OrdinalIgnoreCase)))
            {
                return row;
            }
        }

        return 0;
    }

    private static Dictionary<string, int> BuildColumnMap(IXLWorksheet worksheet, int headerRow)
    {
        var result = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (var cell in worksheet.Row(headerRow).CellsUsed())
        {
            var key = NormalizeHeader(ReadText(cell));
            if (!string.IsNullOrWhiteSpace(key) && !result.ContainsKey(key))
            {
                result[key] = cell.Address.ColumnNumber;
            }
        }

        return result;
    }

    private static string FindClientName(IXLWorksheet worksheet)
    {
        var marker = worksheet.CellsUsed().FirstOrDefault(cell =>
            ReadText(cell).Contains("Financial Summary of Activity", StringComparison.OrdinalIgnoreCase));
        if (marker is null)
        {
            return "";
        }

        for (var row = marker.Address.RowNumber - 1; row >= 1 && row >= marker.Address.RowNumber - 4; row--)
        {
            var candidate = ReadText(worksheet.Cell(row, marker.Address.ColumnNumber));
            if (!string.IsNullOrWhiteSpace(candidate))
            {
                return candidate.Trim();
            }
        }

        return "";
    }

    private static string ReadMetadata(IXLWorksheet worksheet, string label)
    {
        foreach (var cell in worksheet.CellsUsed())
        {
            var value = ReadText(cell);
            if (!value.StartsWith(label, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var separator = value.IndexOf(':');
            if (separator >= 0 && separator + 1 < value.Length)
            {
                return value[(separator + 1)..].Trim();
            }

            var next = ReadText(worksheet.Cell(cell.Address.RowNumber, cell.Address.ColumnNumber + 1));
            if (!string.IsNullOrWhiteSpace(next))
            {
                return next.Trim();
            }
        }

        return "";
    }

    private static void ReadPeriodRange(IXLWorksheet worksheet, AsoFinancialActivityStatement statement)
    {
        var text = ReadMetadata(worksheet, "Time Period");
        var match = PeriodRegex.Match(text);
        if (!match.Success)
        {
            return;
        }

        if (DateOnly.TryParse(match.Groups["start"].Value, CultureInfo.GetCultureInfo("en-CA"), out var start))
        {
            statement.PeriodStart = start;
        }

        if (DateOnly.TryParse(match.Groups["end"].Value, CultureInfo.GetCultureInfo("en-CA"), out var end))
        {
            statement.PeriodEnd = end;
        }
    }

    private static DateOnly? ParseMonth(IXLCell cell, string label)
    {
        if (cell.TryGetValue<DateTime>(out var date))
        {
            return new DateOnly(date.Year, date.Month, 1);
        }

        var formats = new[] { "MMMM yyyy", "MMM yyyy", "M/d/yyyy", "MM/dd/yyyy" };
        if (DateTime.TryParseExact(label.Trim(), formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out date) ||
            DateTime.TryParse(label, CultureInfo.GetCultureInfo("en-CA"), DateTimeStyles.None, out date))
        {
            return new DateOnly(date.Year, date.Month, 1);
        }

        return null;
    }

    private static decimal? Number(
        IXLWorksheet worksheet,
        int row,
        IReadOnlyDictionary<string, int> columns,
        string key)
    {
        if (!columns.TryGetValue(key, out var column))
        {
            return null;
        }

        var cell = worksheet.Cell(row, column);
        if (cell.IsEmpty())
        {
            return null;
        }

        if (cell.TryGetValue<decimal>(out var value))
        {
            return value;
        }

        var text = ReadText(cell)
            .Replace("$", "", StringComparison.Ordinal)
            .Replace(",", "", StringComparison.Ordinal)
            .Replace("%", "", StringComparison.Ordinal)
            .Trim();
        var negative = text.StartsWith('(') && text.EndsWith(')');
        text = text.Trim('(', ')');
        if (!decimal.TryParse(text, NumberStyles.Number | NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out value))
        {
            return null;
        }

        if (negative)
        {
            value = -value;
        }

        if (key.Equals("lossratio", StringComparison.OrdinalIgnoreCase) &&
            ReadText(cell).Contains('%') &&
            Math.Abs(value) > 1m)
        {
            value /= 100m;
        }

        return value;
    }

    private static string Text(
        IXLWorksheet worksheet,
        int row,
        IReadOnlyDictionary<string, int> columns,
        string key)
    {
        return columns.TryGetValue(key, out var column)
            ? ReadText(worksheet.Cell(row, column))
            : "";
    }

    private static int Column(IReadOnlyDictionary<string, int> columns, string key)
    {
        return columns.TryGetValue(key, out var column) ? column : 1;
    }

    private static string NormalizeHeader(string value)
    {
        return new string(value
            .ToLowerInvariant()
            .Where(char.IsLetterOrDigit)
            .ToArray());
    }

    private static string ReadText(IXLCell cell)
    {
        return cell.GetFormattedString().Trim();
    }

    private static bool IsWorkbook(string filePath)
    {
        return Path.GetExtension(filePath).Equals(".xlsx", StringComparison.OrdinalIgnoreCase) ||
               Path.GetExtension(filePath).Equals(".xlsm", StringComparison.OrdinalIgnoreCase);
    }

    private static string FindCarrier(IXLWorksheet worksheet)
    {
        var sample = worksheet.RangeUsed()?.Cells()
            .Take(500)
            .Select(ReadText)
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .ToList() ?? [];

        var carriers = new[]
        {
            "Manulife",
            "Canada Life",
            "Sun Life",
            "Equitable Life",
            "RBC",
            "GreenShield",
            "GroupHEALTH",
            "GroupSource",
            "Victor"
        };
        return carriers.FirstOrDefault(carrier =>
            sample.Any(value => value.Contains(carrier, StringComparison.OrdinalIgnoreCase))) ?? "";
    }
}

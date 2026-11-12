using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using RenewalReviewBuilder.App.Models;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

namespace RenewalReviewBuilder.App.Services;

/// <summary>
/// Extracts a Sun Life "Summary of All Operating Accounts" statement embedded in a renewal PDF.
/// The source PDF can expose its custom font as /UNICxxxx glyph names, so glyph decoding happens
/// before signature matching and coordinate-based table extraction.
/// </summary>
public sealed class SunLifeOperatingAccountsPdfParser
{
    private const double RowTolerance = 3.0;
    private const double LabelColumnRight = 150.0;
    private static readonly CultureInfo CanadianCulture = CultureInfo.GetCultureInfo("en-CA");
    private static readonly Regex UnicodeGlyphRegex = new(
        @"/(?:UNIC|uni)(?<hex>[0-9A-Fa-f]{4,6})",
        RegexOptions.Compiled);
    private static readonly Regex NonAlphaNumericRegex = new(@"[^A-Za-z0-9]+", RegexOptions.Compiled);
    private static readonly Regex ContractRegex = new(@"Contract\s*:\s*(?<value>[A-Za-z0-9-]+)", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex PeriodRegex = new(
        @"For\s+the\s+Period\s*:\s*(?<start>.+?)\s+to\s+(?<end>.+)$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex StatementDateRegex = new(
        @"Statement\s+Date\s*:\s*(?<value>.+)$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex FundingStatusRegex = new(
        @"Current\s+Funding\s+Status\s*:\s*(?<value>.+)$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly HashSet<string> MonthLabels = new(StringComparer.OrdinalIgnoreCase)
    {
        "JAN", "FEB", "MAR", "APR", "MAY", "JUN",
        "JUL", "AUG", "SEP", "OCT", "NOV", "DEC", "TOTAL"
    };

    public AsoFinancialActivityStatement? Extract(string pdfPath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(pdfPath);
        if (!File.Exists(pdfPath))
        {
            throw new FileNotFoundException("The renewal PDF was not found.", pdfPath);
        }

        using var document = PdfDocument.Open(pdfPath);
        var pages = document.GetPages().Select(ParsePage).ToList();
        var firstPageIndex = pages.FindIndex(page =>
            page.NormalizedText.Contains("SUMMARYOFALLOPERATINGACCOUNTS", StringComparison.Ordinal) &&
            page.NormalizedText.Contains("PAGE1OF4", StringComparison.Ordinal));
        if (firstPageIndex < 0)
        {
            return null;
        }

        var statementPages = pages.Skip(firstPageIndex).Take(4).ToList();
        if (statementPages.Count != 4 ||
            statementPages.Where((page, index) =>
                !page.NormalizedText.Contains($"PAGE{index + 1}OF4", StringComparison.Ordinal)).Any())
        {
            throw new InvalidDataException("The four-page Summary of All Operating Accounts section is incomplete.");
        }

        var statement = ReadHeader(statementPages[0], pdfPath);
        statement.SourcePageStart = statementPages[0].PageNumber;
        statement.SourcePageEnd = statementPages[^1].PageNumber;
        statement.SourceSheet = $"PDF pages {statement.SourcePageStart}-{statement.SourcePageEnd}: Summary of All Operating Accounts";

        ReadActivityPage(statementPages[0], statement);
        ReadActivityPage(statementPages[1], statement);
        ReadTaxPage(statementPages[2], statement);
        ReadTaxPage(statementPages[3], statement);
        if (statement.Total is not null && !statement.Total.Gst.HasValue &&
            statement.Periods.Count > 0 && statement.Periods.All(period => period.Gst.HasValue))
        {
            statement.Total.Gst = statement.Periods.Sum(period => period.Gst!.Value);
        }

        statement.Periods.Sort((left, right) => Nullable.Compare(left.ActivityMonth, right.ActivityMonth));
        for (var index = 0; index < statement.Periods.Count; index++)
        {
            statement.Periods[index].DisplayOrder = index;
            FinalizePeriod(statement.Periods[index]);
        }

        if (statement.Total is not null)
        {
            statement.Total.DisplayOrder = statement.Periods.Count;
            FinalizePeriod(statement.Total);
        }

        Validate(statement);
        statement.ImportStatus = statement.ValidationIssues.Count == 0
            ? AsoFinancialImportStatuses.Complete
            : AsoFinancialImportStatuses.NeedsReview;
        statement.ValidationSummary = string.Join(" | ", statement.ValidationIssues);
        return statement;
    }

    private static ParsedPage ParsePage(Page page)
    {
        var tokens = page.GetWords()
            .Select(word => new PageToken(
                DecodeGlyphNames(word.Text),
                word.BoundingBox.Left,
                word.BoundingBox.Right,
                word.BoundingBox.Bottom))
            .Where(token => !string.IsNullOrWhiteSpace(token.Text))
            .OrderByDescending(token => token.Bottom)
            .ThenBy(token => token.Left)
            .ToList();

        var rows = new List<PageRow>();
        foreach (var token in tokens)
        {
            var row = rows.FirstOrDefault(candidate => Math.Abs(candidate.Baseline - token.Bottom) < RowTolerance);
            if (row is null)
            {
                row = new PageRow(token.Bottom);
                rows.Add(row);
            }

            row.Tokens.Add(token);
        }

        rows = rows
            .OrderByDescending(row => row.Baseline)
            .Select(row =>
            {
                row.Tokens.Sort((left, right) => left.Left.CompareTo(right.Left));
                return row;
            })
            .ToList();

        return new ParsedPage(
            page.Number,
            rows,
            Normalize(string.Concat(tokens.Select(token => token.Text))));
    }

    private static AsoFinancialActivityStatement ReadHeader(ParsedPage page, string pdfPath)
    {
        var headerRows = page.Rows.TakeWhile(row => !IsMonthHeader(row)).ToList();
        var clientName = headerRows
            .Select(row => row.Text)
            .FirstOrDefault(text => !string.IsNullOrWhiteSpace(text) &&
                                    !text.StartsWith("Contract", StringComparison.OrdinalIgnoreCase) &&
                                    !text.StartsWith("Summary", StringComparison.OrdinalIgnoreCase)) ?? "";
        var contract = MatchValue(headerRows, ContractRegex);
        var periodText = Match(headerRows, PeriodRegex);
        var statementDate = MatchValue(headerRows, StatementDateRegex);
        var fundingStatus = MatchValue(headerRows, FundingStatusRegex);

        var statement = new AsoFinancialActivityStatement
        {
            ClientName = clientName,
            ContractNumber = contract,
            Carrier = "Sun Life",
            BillingArrangement = fundingStatus,
            SourceFundingStatus = fundingStatus,
            Benefits = "Dental",
            SourceFilePath = Path.GetFullPath(pdfPath),
            SourceFileName = Path.GetFileName(pdfPath),
            ImportedAt = DateTime.Now,
            StatementDate = ParseDate(statementDate)
        };

        if (periodText is not null)
        {
            statement.PeriodStart = ParseDate(periodText.Groups["start"].Value);
            statement.PeriodEnd = ParseDate(periodText.Groups["end"].Value);
        }

        return statement;
    }

    private static void ReadActivityPage(ParsedPage page, AsoFinancialActivityStatement statement)
    {
        var columns = ReadColumns(page);
        if (columns.Count == 0)
        {
            statement.ValidationIssues.Add($"Page {page.PageNumber}: month columns were not detected.");
            return;
        }

        var opening = ReadField(page, columns, "OPENINGBALANCE");
        var payments = ReadField(page, columns, "PAYMENTRECEIVEDNOTALLOCATEDBY");
        var claims = ReadField(page, columns, "TOTALCLAIMSPAID");
        var handling = ReadField(page, columns, "CLAIMSHANDLINGCHARGES");
        var generalAdmin = ReadField(page, columns, "GENERALADMINCHARGE");
        var profit = ReadField(page, columns, "PROFITCHARGE");
        var serviceFees = ReadField(page, columns, "TOTALSERVICEFEES");
        var consulting = ReadField(page, columns, "CONSULTINGFEES");
        var taxes = ReadField(page, columns, "TOTALTAXES");
        var interest = ReadField(page, columns, "TOTALINTEREST");
        var closing = ReadField(page, columns, "CLOSINGBALANCE");

        foreach (var column in columns)
        {
            var period = column.Label.Equals("TOTAL", StringComparison.OrdinalIgnoreCase)
                ? statement.Total ??= new AsoFinancialActivityPeriod { PeriodLabel = "Total", IsTotal = true }
                : GetOrCreateMonth(statement, column.Label);
            if (period is null)
            {
                continue;
            }

            period.OpeningBalance = Value(opening, column.Label);
            period.PaymentReceivedNotAllocated = Value(payments, column.Label);
            period.DentalClaimsPaid = Value(claims, column.Label);
            period.ClaimsHandlingCharges = Value(handling, column.Label);
            period.GeneralAdminCharge = Value(generalAdmin, column.Label);
            period.ProfitCharge = Value(profit, column.Label);
            period.TotalServiceFees = Value(serviceFees, column.Label);
            period.ConsultingFees = Value(consulting, column.Label);
            period.TotalTaxes = Value(taxes, column.Label);
            period.InterestCreditCharge = Value(interest, column.Label);
            period.CurrentBalance = Value(closing, column.Label);
        }
    }

    private static void ReadTaxPage(ParsedPage page, AsoFinancialActivityStatement statement)
    {
        var columns = ReadColumns(page);
        var gst = ReadField(page, columns, "GST");
        foreach (var column in columns)
        {
            var period = column.Label.Equals("TOTAL", StringComparison.OrdinalIgnoreCase)
                ? statement.Total
                : FindMonth(statement, column.Label);
            if (period is not null)
            {
                period.Gst = Value(gst, column.Label);
            }
        }
    }

    private static List<PageColumn> ReadColumns(ParsedPage page)
    {
        var expected = ExpectedColumns(page);
        var header = page.Rows.FirstOrDefault(IsMonthHeader);
        if (header is null)
        {
            return expected;
        }

        var columns = header.Tokens
            .Select(token => new PageColumn(Normalize(token.Text), (token.Left + token.Right) / 2.0))
            .Where(column => MonthLabels.Contains(column.Label))
            .ToList();
        if (!columns.Any(column => column.Label.Equals("TOTAL", StringComparison.OrdinalIgnoreCase)) &&
            (Normalize(header.Text).Contains("TOTAL", StringComparison.Ordinal) ||
             page.NormalizedText.Contains("PAGE2OF4", StringComparison.Ordinal) ||
             page.NormalizedText.Contains("PAGE4OF4", StringComparison.Ordinal)))
        {
            var totalTokens = header.Tokens
                .Where(token => token.Left > columns.Max(column => column.Center))
                .ToList();
            var center = totalTokens.Count > 0
                ? (totalTokens.Min(token => token.Left) + totalTokens.Max(token => token.Right)) / 2.0
                : columns[^1].Center + 58.0;
            columns.Add(new PageColumn("TOTAL", center));
        }

        return columns.Count == expected.Count ? columns : expected;
    }

    private static List<PageColumn> ExpectedColumns(ParsedPage page)
    {
        if (page.NormalizedText.Contains("PAGE1OF4", StringComparison.Ordinal) ||
            page.NormalizedText.Contains("PAGE3OF4", StringComparison.Ordinal))
        {
            return
            [
                new("JUN", 178), new("JUL", 236), new("AUG", 293), new("SEP", 351),
                new("OCT", 408), new("NOV", 466), new("DEC", 524)
            ];
        }

        if (page.NormalizedText.Contains("PAGE2OF4", StringComparison.Ordinal) ||
            page.NormalizedText.Contains("PAGE4OF4", StringComparison.Ordinal))
        {
            return
            [
                new("JAN", 178), new("FEB", 236), new("MAR", 293),
                new("APR", 351), new("MAY", 408), new("TOTAL", 466)
            ];
        }

        return [];
    }

    private static Dictionary<string, decimal?> ReadField(
        ParsedPage page,
        IReadOnlyList<PageColumn> columns,
        string normalizedLabel)
    {
        var rowIndex = page.Rows.FindIndex(row => Normalize(row.LabelText).Equals(normalizedLabel, StringComparison.Ordinal));
        if (rowIndex < 0)
        {
            rowIndex = page.Rows.FindIndex(row => Normalize(row.LabelText).Contains(normalizedLabel, StringComparison.Ordinal));
        }
        if (rowIndex < 0)
        {
            return [];
        }

        for (var index = rowIndex; index < Math.Min(page.Rows.Count, rowIndex + 3); index++)
        {
            var values = MapAmounts(page.Rows[index], columns);
            if (values.Count > 0)
            {
                return values;
            }
        }

        return [];
    }

    private static Dictionary<string, decimal?> MapAmounts(PageRow row, IReadOnlyList<PageColumn> columns)
    {
        var result = new Dictionary<string, decimal?>(StringComparer.OrdinalIgnoreCase);
        var amounts = row.Tokens
            .Where(token => token.Left >= LabelColumnRight)
            .Select(token => TryParseAmount(token.Text, out var amount)
                ? new ParsedAmount(token, amount)
                : null)
            .Where(amount => amount is not null)
            .Cast<ParsedAmount>()
            .OrderBy(amount => amount.Token.Left)
            .ToList();
        if (amounts.Count == columns.Count)
        {
            for (var index = 0; index < columns.Count; index++)
            {
                result[columns[index].Label] = amounts[index].Value;
            }

            return result;
        }

        foreach (var parsed in amounts)
        {
            var center = (parsed.Token.Left + parsed.Token.Right) / 2.0;
            var column = columns.OrderBy(candidate => Math.Abs(candidate.Center - center)).FirstOrDefault();
            if (column is not null)
            {
                result[column.Label] = parsed.Value;
            }
        }

        return result;
    }

    private static AsoFinancialActivityPeriod? GetOrCreateMonth(
        AsoFinancialActivityStatement statement,
        string monthLabel)
    {
        var existing = FindMonth(statement, monthLabel);
        if (existing is not null)
        {
            return existing;
        }

        if (!statement.PeriodStart.HasValue ||
            !DateTime.TryParseExact(monthLabel, "MMM", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedMonth))
        {
            return null;
        }

        var year = parsedMonth.Month >= statement.PeriodStart.Value.Month
            ? statement.PeriodStart.Value.Year
            : statement.PeriodStart.Value.Year + 1;
        var activityMonth = new DateOnly(year, parsedMonth.Month, 1);
        var period = new AsoFinancialActivityPeriod
        {
            ActivityMonth = activityMonth,
            PeriodLabel = activityMonth.ToString("MMMM yyyy", CultureInfo.InvariantCulture)
        };
        statement.Periods.Add(period);
        return period;
    }

    private static AsoFinancialActivityPeriod? FindMonth(
        AsoFinancialActivityStatement statement,
        string monthLabel)
    {
        if (!DateTime.TryParseExact(monthLabel, "MMM", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedMonth))
        {
            return null;
        }

        return statement.Periods.FirstOrDefault(period => period.ActivityMonth?.Month == parsedMonth.Month);
    }

    private static void FinalizePeriod(AsoFinancialActivityPeriod period)
    {
        period.TotalDeposit = period.PaymentReceivedNotAllocated;
        period.UnpooledBenefitPayments = period.DentalClaimsPaid;
        period.TotalBenefitPayment = period.DentalClaimsPaid;
        period.AdminFees = period.TotalServiceFees;
        period.OtherCharges = period.ConsultingFees;
        period.TotalCharges = SumWhenAvailable(
            period.TotalBenefitPayment,
            period.TotalServiceFees,
            period.ConsultingFees,
            period.TotalTaxes,
            period.InterestCreditCharge);
        period.BalanceStatus = period.CurrentBalance switch
        {
            < 0m => "Deficit",
            > 0m => "Funds Available",
            0m => "Balanced",
            _ => ""
        };
    }

    private static void Validate(AsoFinancialActivityStatement statement)
    {
        Require(statement, statement.ClientName, "Client name was not detected.");
        Require(statement, statement.ContractNumber, "Contract number was not detected.");
        Require(statement, statement.SourceFundingStatus, "Current Funding Status was not detected.");
        if (!statement.StatementDate.HasValue) statement.ValidationIssues.Add("Statement Date was not detected.");
        if (!statement.PeriodStart.HasValue || !statement.PeriodEnd.HasValue)
            statement.ValidationIssues.Add("The operating-account period was not detected.");
        if (statement.Periods.Count != 12)
            statement.ValidationIssues.Add($"Expected 12 monthly operating-account rows but found {statement.Periods.Count}.");
        if (statement.Total is null)
        {
            statement.ValidationIssues.Add("The reported total column was not detected.");
            return;
        }

        var ordered = statement.Periods.OrderBy(period => period.ActivityMonth).ToList();
        for (var index = 0; index < ordered.Count; index++)
        {
            var period = ordered[index];
            var previousMonth = index > 0 ? ordered[index - 1].ActivityMonth : null;
            if (period.ActivityMonth is { } activityMonth && previousMonth is { } priorMonth &&
                activityMonth != priorMonth.AddMonths(1))
            {
                statement.ValidationIssues.Add($"{period.PeriodLabel}: activity months are not consecutive.");
            }

            ValidateServiceFees(statement, period);
            ValidateBalance(statement, period);
            ValidateTax(statement, period);
            var previousBalance = index > 0 ? ordered[index - 1].CurrentBalance : null;
            if (period.OpeningBalance is { } openingBalance && previousBalance is { } priorBalance &&
                !Close(openingBalance, priorBalance))
            {
                statement.ValidationIssues.Add($"{period.PeriodLabel}: opening balance does not match the prior closing balance.");
            }
        }

        ValidateServiceFees(statement, statement.Total);
        ValidateTax(statement, statement.Total);
        ValidateTotal(statement, period => period.PaymentReceivedNotAllocated, statement.Total.PaymentReceivedNotAllocated, "payments received");
        ValidateTotal(statement, period => period.DentalClaimsPaid, statement.Total.DentalClaimsPaid, "Dental claims paid");
        ValidateTotal(statement, period => period.ClaimsHandlingCharges, statement.Total.ClaimsHandlingCharges, "claims handling charges");
        ValidateTotal(statement, period => period.GeneralAdminCharge, statement.Total.GeneralAdminCharge, "general administration charges");
        ValidateTotal(statement, period => period.ProfitCharge, statement.Total.ProfitCharge, "profit charges");
        ValidateTotal(statement, period => period.TotalServiceFees, statement.Total.TotalServiceFees, "total service fees");
        ValidateTotal(statement, period => period.ConsultingFees, statement.Total.ConsultingFees, "consulting fees");
        ValidateTotal(statement, period => period.TotalTaxes, statement.Total.TotalTaxes, "total taxes");
        ValidateTotal(statement, period => period.InterestCreditCharge, statement.Total.InterestCreditCharge, "total interest");

        if (statement.PeriodStart.HasValue && ordered.FirstOrDefault()?.ActivityMonth is { } firstMonth &&
            firstMonth != new DateOnly(statement.PeriodStart.Value.Year, statement.PeriodStart.Value.Month, 1))
        {
            statement.ValidationIssues.Add("The first activity month does not match the reported period start.");
        }

        if (statement.PeriodEnd.HasValue && ordered.LastOrDefault()?.ActivityMonth is { } lastMonth &&
            lastMonth != new DateOnly(statement.PeriodEnd.Value.Year, statement.PeriodEnd.Value.Month, 1))
        {
            statement.ValidationIssues.Add("The last activity month does not match the reported period end.");
        }
    }

    private static void ValidateServiceFees(AsoFinancialActivityStatement statement, AsoFinancialActivityPeriod period)
    {
        if (period.ClaimsHandlingCharges.HasValue && period.GeneralAdminCharge.HasValue &&
            period.ProfitCharge.HasValue && period.TotalServiceFees.HasValue &&
            !Close(period.ClaimsHandlingCharges.Value + period.GeneralAdminCharge.Value + period.ProfitCharge.Value,
                period.TotalServiceFees.Value))
        {
            statement.ValidationIssues.Add($"{period.PeriodLabel}: service-fee components do not reconcile.");
        }
    }

    private static void ValidateBalance(AsoFinancialActivityStatement statement, AsoFinancialActivityPeriod period)
    {
        if (!period.OpeningBalance.HasValue || !period.PaymentReceivedNotAllocated.HasValue ||
            !period.DentalClaimsPaid.HasValue || !period.TotalServiceFees.HasValue ||
            !period.ConsultingFees.HasValue || !period.TotalTaxes.HasValue ||
            !period.InterestCreditCharge.HasValue || !period.CurrentBalance.HasValue)
        {
            statement.ValidationIssues.Add($"{period.PeriodLabel}: the operating-account balance equation is incomplete.");
            return;
        }

        var calculated = period.OpeningBalance.Value + period.PaymentReceivedNotAllocated.Value -
                         period.DentalClaimsPaid.Value - period.TotalServiceFees.Value -
                         period.ConsultingFees.Value - period.TotalTaxes.Value -
                         period.InterestCreditCharge.Value;
        if (!Close(calculated, period.CurrentBalance.Value))
        {
            statement.ValidationIssues.Add(
                $"{period.PeriodLabel}: calculated closing balance {Money(calculated)} does not match {Money(period.CurrentBalance.Value)}.");
        }
    }

    private static void ValidateTax(AsoFinancialActivityStatement statement, AsoFinancialActivityPeriod period)
    {
        if (period.TotalTaxes.HasValue && !period.Gst.HasValue)
        {
            statement.ValidationIssues.Add($"{period.PeriodLabel}: the GST breakdown was not detected.");
            return;
        }

        if (period.TotalTaxes.HasValue && period.Gst.HasValue && !Close(period.TotalTaxes.Value, period.Gst.Value))
        {
            statement.ValidationIssues.Add($"{period.PeriodLabel}: GST does not match total taxes.");
        }
    }

    private static void ValidateTotal(
        AsoFinancialActivityStatement statement,
        Func<AsoFinancialActivityPeriod, decimal?> selector,
        decimal? reported,
        string label)
    {
        if (!reported.HasValue || statement.Periods.Any(period => !selector(period).HasValue))
        {
            statement.ValidationIssues.Add($"The {label} total could not be fully reconciled.");
            return;
        }

        var calculated = statement.Periods.Sum(period => selector(period)!.Value);
        if (!Close(calculated, reported.Value))
        {
            statement.ValidationIssues.Add(
                $"The monthly {label} sum {Money(calculated)} does not match the reported total {Money(reported.Value)}.");
        }
    }

    private static void Require(AsoFinancialActivityStatement statement, string value, string issue)
    {
        if (string.IsNullOrWhiteSpace(value)) statement.ValidationIssues.Add(issue);
    }

    private static decimal? Value(IReadOnlyDictionary<string, decimal?> values, string label) =>
        values.TryGetValue(label, out var value) ? value : null;

    private static decimal? SumWhenAvailable(params decimal?[] values) =>
        values.All(value => !value.HasValue) ? null : values.Sum(value => value ?? 0m);

    private static bool IsMonthHeader(PageRow row) =>
        row.Tokens.Count(token => MonthLabels.Contains(Normalize(token.Text))) >= 5;

    private static string MatchValue(IReadOnlyList<PageRow> rows, Regex regex) =>
        Match(rows, regex)?.Groups["value"].Value.Trim() ?? "";

    private static Match? Match(IReadOnlyList<PageRow> rows, Regex regex)
    {
        foreach (var row in rows)
        {
            var match = regex.Match(row.Text);
            if (match.Success) return match;
        }

        return null;
    }

    private static DateOnly? ParseDate(string value)
    {
        var cleaned = Regex.Replace(value.Replace(".", ""), @"\s+", " ").Trim();
        return DateTime.TryParse(cleaned, CanadianCulture, DateTimeStyles.AllowWhiteSpaces, out var parsed)
            ? DateOnly.FromDateTime(parsed)
            : null;
    }

    private static bool TryParseAmount(string value, out decimal amount)
    {
        var cleaned = value.Trim().Replace("$", "").Replace(",", "").Replace(" ", "");
        var negative = cleaned.StartsWith('(') && cleaned.EndsWith(')');
        cleaned = cleaned.Trim('(', ')');
        if (!decimal.TryParse(cleaned, NumberStyles.Number | NumberStyles.AllowLeadingSign,
                CultureInfo.InvariantCulture, out amount))
        {
            return false;
        }

        if (negative) amount = -amount;
        return true;
    }

    private static string DecodeGlyphNames(string value)
    {
        if (string.IsNullOrEmpty(value) || !value.Contains("/UNI", StringComparison.OrdinalIgnoreCase))
        {
            return value;
        }

        var builder = new StringBuilder();
        var lastIndex = 0;
        foreach (Match match in UnicodeGlyphRegex.Matches(value))
        {
            builder.Append(value, lastIndex, match.Index - lastIndex);
            if (int.TryParse(match.Groups["hex"].Value, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var codePoint) &&
                codePoint is >= 0 and <= 0x10FFFF)
            {
                builder.Append(char.ConvertFromUtf32(codePoint));
            }
            lastIndex = match.Index + match.Length;
        }

        builder.Append(value, lastIndex, value.Length - lastIndex);
        return builder.ToString();
    }

    private static string Normalize(string value) => NonAlphaNumericRegex.Replace(value ?? "", "").ToUpperInvariant();
    private static bool Close(decimal left, decimal right) => Math.Abs(left - right) <= 0.05m;
    private static string Money(decimal value) => value.ToString("$#,0.00;($#,0.00);$0.00", CanadianCulture);

    private sealed record PageToken(string Text, double Left, double Right, double Bottom);
    private sealed record ParsedAmount(PageToken Token, decimal Value);
    private sealed record PageColumn(string Label, double Center);
    private sealed record ParsedPage(int PageNumber, List<PageRow> Rows, string NormalizedText);

    private sealed class PageRow(double baseline)
    {
        public double Baseline { get; } = baseline;
        public List<PageToken> Tokens { get; } = [];
        public string Text => string.Join(" ", Tokens.Select(token => token.Text));
        public string LabelText => string.Join(" ", Tokens.Where(token => token.Left < LabelColumnRight).Select(token => token.Text));
    }
}

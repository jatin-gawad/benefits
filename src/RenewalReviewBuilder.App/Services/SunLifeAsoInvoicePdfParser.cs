using System.Globalization;
using System.Text.RegularExpressions;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

/// <summary>
/// Reads Sun Life monthly ASO Invoice PDFs. The statement values are kept as
/// carrier-observed facts; billed fees and GST are never replaced by estimates.
/// </summary>
public sealed class SunLifeAsoInvoicePdfParser
{
    private static readonly CultureInfo CanadianCulture = CultureInfo.GetCultureInfo("en-CA");
    private static readonly Regex AmountRegex = new(
        @"(?<value>\(\s*\$?\s*[\d,]+(?:\.\d{2})?\s*\)|-?\s*\$?\s*[\d,]+(?:\.\d{2})?)",
        RegexOptions.Compiled);
    private static readonly Regex NonAlphaNumericRegex = new(@"[^A-Za-z0-9]+", RegexOptions.Compiled);
    private static readonly Regex PeriodRegex = new(
        @"^For\s+the\s+Period\s*:\s*(?<start>.+?)\s+-\s+(?<end>.+)$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex RequiredFloatRegex = new(
        @"float\s+established\s+at\s+\$?(?<amount>[\d,]+(?:\.\d{2})?)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private readonly PdfTextExtractor _extractor;

    public SunLifeAsoInvoicePdfParser(PdfTextExtractor? extractor = null)
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

        foreach (var path in Directory.EnumerateFiles(folderPath, "*.*", SearchOption.AllDirectories)
                     .Where(path => path.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
                     .OrderBy(path => path, StringComparer.OrdinalIgnoreCase))
        {
            try
            {
                var statement = Extract(path);
                if (statement is null)
                {
                    continue;
                }

                result.FinancialStatements.Add(statement);
                var period = statement.Periods.SingleOrDefault();
                if (period?.ActivityMonth is null ||
                    period.DentalClaimsPaid is null ||
                    !statement.ImportStatus.Equals(AsoFinancialImportStatuses.Complete, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                result.Records.Add(new AsoBillingActivityRecord
                {
                    ServiceMonth = period.ActivityMonth.Value,
                    PaidDentalClaims = period.DentalClaimsPaid.Value,
                    AdministrationFee = SumWhenAvailable(period.TotalServiceFees, period.ConsultingFees),
                    Gst = period.Gst,
                    InterestCreditCharge = period.InterestCreditCharge,
                    TotalAsoCost = period.TotalCharges,
                    SourcePath = path
                });
            }
            catch (Exception exception)
            {
                result.Warnings.Add($"{Path.GetFileName(path)} could not be read as a Sun Life ASO invoice: {exception.Message}");
            }
        }

        ReplaceWithDistinctMonths(result.Records);
        return result;
    }

    public AsoFinancialActivityStatement? Extract(string pdfPath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(pdfPath);
        if (!File.Exists(pdfPath))
        {
            throw new FileNotFoundException("The Sun Life ASO invoice was not found.", pdfPath);
        }

        var pages = _extractor.Extract(pdfPath);
        if (pages.Count == 0)
        {
            return null;
        }

        var firstPage = pages[0];
        var firstPageText = firstPage.Text;
        if (!firstPageText.Contains("ASO Invoice", StringComparison.OrdinalIgnoreCase) ||
            !firstPageText.Contains("Plan Sponsor", StringComparison.OrdinalIgnoreCase) ||
            !firstPageText.Contains("Sun Life", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        var allLines = pages.SelectMany(page => page.Lines).ToList();
        var allText = string.Join(" ", pages.Select(page => page.Text));
        var periodLine = firstPage.Lines.FirstOrDefault(line =>
            line.StartsWith("For the Period", StringComparison.OrdinalIgnoreCase));
        var periodMatch = periodLine is null ? Match.Empty : PeriodRegex.Match(periodLine);
        var periodStart = periodMatch.Success ? ParseDate(periodMatch.Groups["start"].Value) : null;
        var periodEnd = periodMatch.Success ? ParseDate(periodMatch.Groups["end"].Value) : null;
        var fundingStatus = ReadTextValue(allLines, "Funding Status") ??
                            (allText.Contains("unfunded arrangement", StringComparison.OrdinalIgnoreCase)
                                ? "Unfunded"
                                : "");

        var period = new AsoFinancialActivityPeriod
        {
            ActivityMonth = periodStart.HasValue ? new DateOnly(periodStart.Value.Year, periodStart.Value.Month, 1) : null,
            PeriodLabel = periodStart?.ToString("MMMM yyyy", CultureInfo.InvariantCulture) ?? "Monthly ASO invoice",
            DisplayOrder = 0,
            PaymentReceivedNotAllocated = ReadLabeledAmount(firstPage.Lines, "Total Payments and Credits"),
            DentalClaimsPaid = ReadLabeledAmount(firstPage.Lines, "Claims Paid"),
            ClaimsHandlingCharges = ReadLabeledAmount(firstPage.Lines, "Claims Handling Charges"),
            GeneralAdminCharge = ReadLabeledAmount(firstPage.Lines, "General Administration Charge"),
            ProfitCharge = ReadLabeledAmount(firstPage.Lines, "Profit Charge"),
            ConsultingFees = ReadLabeledAmount(firstPage.Lines, "Payment made to your consultant - Fees"),
            TotalTaxes = ReadLabeledAmount(firstPage.Lines, "Subtotal Taxes"),
            Gst = ReadLabeledAmount(firstPage.Lines, "GST"),
            InterestCreditCharge = ReadLabeledAmount(firstPage.Lines, "Total Interest"),
            TotalCharges = ReadLabeledAmount(firstPage.Lines, "Total Charges"),
            SubTotal = ReadLabeledAmount(firstPage.Lines, "Subtotal Claims Paid and Service Fees"),
            RequiredFloat = ReadRequiredFloat(allText)
        };
        ReadBalanceSummary(firstPage.Lines, period);
        period.TotalServiceFees = SumWhenAvailable(
            period.ClaimsHandlingCharges,
            period.GeneralAdminCharge,
            period.ProfitCharge);
        period.TotalDeposit = period.PaymentReceivedNotAllocated;
        period.NetDeposit = period.PaymentReceivedNotAllocated;
        period.UnpooledBenefitPayments = period.DentalClaimsPaid;
        period.TotalBenefitPayment = period.DentalClaimsPaid;
        period.AdminFees = period.TotalServiceFees;
        period.OtherCharges = period.ConsultingFees;
        period.BalanceStatus = BalanceStatus(period.CurrentBalance);

        var statement = new AsoFinancialActivityStatement
        {
            ClientName = ReadTextValue(firstPage.Lines, "Plan Sponsor") ?? "",
            ContractNumber = ReadTextValue(firstPage.Lines, "Contract") ?? "",
            Carrier = "Sun Life",
            BillingArrangement = fundingStatus,
            Divisions = ReadTextValue(firstPage.Lines, "Operating Account") ?? "",
            Benefits = "Dental",
            PeriodStart = periodStart,
            PeriodEnd = periodEnd,
            StatementDate = ParseDate(ReadTextValue(firstPage.Lines, "Statement Date")),
            SourceFundingStatus = fundingStatus,
            SourcePageStart = 1,
            SourcePageEnd = pages.Count,
            SourceFilePath = Path.GetFullPath(pdfPath),
            SourceFileName = Path.GetFileName(pdfPath),
            SourceSheet = "PDF page 1: ASO Invoice statement summary",
            ImportedAt = DateTime.Now
        };
        statement.Periods.Add(period);
        statement.Total = CopyAsTotal(period);
        Validate(statement, period);
        statement.ImportStatus = statement.ValidationIssues.Count == 0
            ? AsoFinancialImportStatuses.Complete
            : AsoFinancialImportStatuses.NeedsReview;
        statement.ValidationSummary = string.Join(" | ", statement.ValidationIssues);
        return statement;
    }

    private static void ReadBalanceSummary(IReadOnlyList<string> lines, AsoFinancialActivityPeriod period)
    {
        var headerIndex = lines.ToList().FindIndex(line =>
            Normalize(line).StartsWith("OPENINGBALANCETOTALCREDITSTOTALCHARGESCLOSINGBALANCEAMOUNTPAYABLE", StringComparison.Ordinal));
        if (headerIndex < 0 || headerIndex + 1 >= lines.Count)
        {
            return;
        }

        var amounts = ReadAmounts(lines[headerIndex + 1]);
        if (amounts.Count >= 5)
        {
            period.OpeningBalance = amounts[0];
            period.CurrentBalance = amounts[3];
        }
    }

    private static decimal? ReadLabeledAmount(IEnumerable<string> lines, string label)
    {
        var normalizedLabel = Normalize(label);
        foreach (var line in lines)
        {
            if (!Normalize(line).StartsWith(normalizedLabel, StringComparison.Ordinal))
            {
                continue;
            }

            var amounts = ReadAmounts(line);
            if (amounts.Count > 0)
            {
                return amounts[^1];
            }
        }

        return null;
    }

    private static string? ReadTextValue(IEnumerable<string> lines, string label)
    {
        var pattern = new Regex($@"^{Regex.Escape(label)}\s*:\s*(?<value>.+)$", RegexOptions.IgnoreCase);
        foreach (var line in lines)
        {
            var match = pattern.Match(line);
            if (match.Success)
            {
                return match.Groups["value"].Value.Trim();
            }
        }

        return null;
    }

    private static decimal? ReadRequiredFloat(string text)
    {
        var match = RequiredFloatRegex.Match(text);
        return match.Success ? ParseAmount(match.Groups["amount"].Value) : null;
    }

    private static List<decimal> ReadAmounts(string text) =>
        AmountRegex.Matches(text)
            .Select(match => ParseAmount(match.Groups["value"].Value))
            .Where(value => value.HasValue)
            .Select(value => value!.Value)
            .ToList();

    private static decimal? ParseAmount(string text)
    {
        var cleaned = text.Trim();
        var negative = cleaned.StartsWith('(') && cleaned.EndsWith(')');
        cleaned = cleaned.Trim('(', ')').Replace("$", "").Replace(",", "").Replace(" ", "");
        if (!decimal.TryParse(cleaned, NumberStyles.Number | NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out var amount))
        {
            return null;
        }

        return negative ? -amount : amount;
    }

    private static DateOnly? ParseDate(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return null;
        }

        if (DateTime.TryParse(text, CanadianCulture, DateTimeStyles.AllowWhiteSpaces, out var parsed) ||
            DateTime.TryParse(text.Replace(".", ""), CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out parsed))
        {
            return DateOnly.FromDateTime(parsed);
        }

        return null;
    }

    private static string Normalize(string value) => NonAlphaNumericRegex.Replace(value, "").ToUpperInvariant();

    private static decimal? SumWhenAvailable(params decimal?[] values)
    {
        var available = values.Where(value => value.HasValue).Select(value => value!.Value).ToList();
        return available.Count == 0 ? null : available.Sum();
    }

    private static string BalanceStatus(decimal? balance) => balance switch
    {
        < 0m => "Deficit",
        > 0m => "Funds Available",
        0m => "Balanced",
        _ => ""
    };

    private static AsoFinancialActivityPeriod CopyAsTotal(AsoFinancialActivityPeriod period) => new()
    {
        PeriodLabel = "Total",
        IsTotal = true,
        DisplayOrder = 1,
        OpeningBalance = period.OpeningBalance,
        PaymentReceivedNotAllocated = period.PaymentReceivedNotAllocated,
        DentalClaimsPaid = period.DentalClaimsPaid,
        ClaimsHandlingCharges = period.ClaimsHandlingCharges,
        GeneralAdminCharge = period.GeneralAdminCharge,
        ProfitCharge = period.ProfitCharge,
        TotalServiceFees = period.TotalServiceFees,
        ConsultingFees = period.ConsultingFees,
        TotalDeposit = period.TotalDeposit,
        NetDeposit = period.NetDeposit,
        UnpooledBenefitPayments = period.UnpooledBenefitPayments,
        TotalBenefitPayment = period.TotalBenefitPayment,
        AdminFees = period.AdminFees,
        OtherCharges = period.OtherCharges,
        TotalTaxes = period.TotalTaxes,
        Gst = period.Gst,
        InterestCreditCharge = period.InterestCreditCharge,
        TotalCharges = period.TotalCharges,
        SubTotal = period.SubTotal,
        RequiredFloat = period.RequiredFloat,
        CurrentBalance = period.CurrentBalance,
        BalanceStatus = period.BalanceStatus
    };

    private static void Validate(AsoFinancialActivityStatement statement, AsoFinancialActivityPeriod period)
    {
        Require(statement, statement.ClientName, "Plan sponsor was not detected.");
        Require(statement, statement.ContractNumber, "Contract number was not detected.");
        if (!statement.PeriodStart.HasValue || !statement.PeriodEnd.HasValue)
            statement.ValidationIssues.Add("The invoice period was not detected.");
        if (!statement.StatementDate.HasValue)
            statement.ValidationIssues.Add("The statement date was not detected.");
        if (!period.DentalClaimsPaid.HasValue)
            statement.ValidationIssues.Add("Dental claims paid were not detected.");
        if (!period.TotalServiceFees.HasValue)
            statement.ValidationIssues.Add("Billed service fees were not detected.");
        if (!period.Gst.HasValue)
            statement.ValidationIssues.Add("GST was not detected.");
        if (!period.TotalCharges.HasValue)
            statement.ValidationIssues.Add("Total charges were not detected.");

        var calculatedSubtotal = SumWhenAvailable(period.DentalClaimsPaid, period.TotalServiceFees, period.ConsultingFees);
        ValidateAmount(statement, calculatedSubtotal, period.SubTotal, "Claims and service-fee subtotal");
        var calculatedTotal = SumWhenAvailable(calculatedSubtotal, period.TotalTaxes, period.InterestCreditCharge);
        ValidateAmount(statement, calculatedTotal, period.TotalCharges, "Total charges");
    }

    private static void ValidateAmount(
        AsoFinancialActivityStatement statement,
        decimal? calculated,
        decimal? reported,
        string label)
    {
        if (calculated.HasValue && reported.HasValue && Math.Abs(calculated.Value - reported.Value) > 0.02m)
        {
            statement.ValidationIssues.Add($"{label} does not reconcile: calculated {calculated.Value:C2}, reported {reported.Value:C2}.");
        }
    }

    private static void Require(AsoFinancialActivityStatement statement, string value, string issue)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            statement.ValidationIssues.Add(issue);
        }
    }

    private static void ReplaceWithDistinctMonths(List<AsoBillingActivityRecord> records)
    {
        var distinct = records
            .GroupBy(record => record.ServiceMonth)
            .Select(group => group.OrderByDescending(record => record.SourcePath, StringComparer.OrdinalIgnoreCase).First())
            .OrderBy(record => record.ServiceMonth)
            .ToList();
        records.Clear();
        records.AddRange(distinct);
    }
}

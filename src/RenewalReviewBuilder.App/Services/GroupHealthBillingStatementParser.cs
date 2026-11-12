using System.Globalization;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed class GroupHealthBillingStatementParser
{
    private static readonly Regex PinLineRegex = new(@"^(?<pin>\d{3}-\d{3}-\d{3}[A-Z])\s+(?<rest>.+)$", RegexOptions.Compiled);
    private static readonly Regex CertificateLineRegex = new(@"^(?<name>.+?)\s+(?<certificate>\d{12,})\s+(?<values>.+)$", RegexOptions.Compiled);
    private static readonly Regex ChangeReportRegex = new(
        @"^(?<pin>\d{3}-\d{3}-\d{3}[A-Z])\s+(?<name>.+?)\s+(?<change>AGE BAND SHIFT|S/OCC|C-[A-Z0-9]+|ADRS|REPT|[ASCTDR])\s+(?<date>\d{4}-\d{2}-\d{2})\s+(?<annual>-?[\d,]+(?:\.\d{2})?)\s+(?<class>\S+)\s+(?<cost>\S+)\s+(?<tail>.*)$",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    private static readonly string[] DetailBenefits =
    [
        "Life",
        "AD&D",
        "LTD",
        "Health",
        "Dental",
        "Flex Admin",
        "Member AD&D",
        "Member AD&D SP",
        "Member AD&D CH",
        "EE Premium",
        "ER Premium",
        "Total Premium"
    ];

    private static readonly string[] SummaryBenefits =
    [
        "Life",
        "AD&D",
        "LTD",
        "Health",
        "EFAP",
        "Virtual Healthcare",
        "Dental",
        "Flex Admin",
        "Member AD&D",
        "Member AD&D SP",
        "Member AD&D CH",
        "EE Premium",
        "ER Premium",
        "Total Premium"
    ];

    private static readonly Dictionary<string, int> MonthNumbers = new(StringComparer.OrdinalIgnoreCase)
    {
        ["Jan"] = 1,
        ["January"] = 1,
        ["Feb"] = 2,
        ["February"] = 2,
        ["Mar"] = 3,
        ["March"] = 3,
        ["Apr"] = 4,
        ["April"] = 4,
        ["May"] = 5,
        ["Jun"] = 6,
        ["June"] = 6,
        ["Jul"] = 7,
        ["July"] = 7,
        ["Aug"] = 8,
        ["August"] = 8,
        ["Sep"] = 9,
        ["Sept"] = 9,
        ["September"] = 9,
        ["Oct"] = 10,
        ["October"] = 10,
        ["Nov"] = 11,
        ["November"] = 11,
        ["Dec"] = 12,
        ["December"] = 12
    };

    private readonly PdfTextExtractor _extractor = new();
    private readonly GroupHealthBillingValidationOptions _options;

    public GroupHealthBillingStatementParser(GroupHealthBillingValidationOptions? options = null)
    {
        _options = options ?? new GroupHealthBillingValidationOptions();
    }

    public GroupHealthBillingExtractionResult Extract(string pdfPath)
    {
        if (string.IsNullOrWhiteSpace(pdfPath) || !File.Exists(pdfPath))
        {
            throw new FileNotFoundException("GroupHEALTH billing statement was not found.", pdfPath);
        }

        var fullPath = Path.GetFullPath(pdfPath);
        var pages = _extractor.Extract(fullPath);
        var result = new GroupHealthBillingExtractionResult
        {
            Carrier = "GroupHEALTH",
            SourceFilePath = fullPath,
            SourceFileName = Path.GetFileName(fullPath),
            SourceFileHash = ComputeFileHash(fullPath),
            PageCount = pages.Count
        };

        ReadMetadata(result, pages);
        ReadChangeReport(result, pages);
        ReadEmployeeRows(result, pages);
        ReadClassAndStatementTotals(result, pages);

        new GroupHealthBillingValidator(_options).Validate(result);
        GroupHealthBillingRateCalculator.BuildRateLines(result);

        return result;
    }

    private static void ReadMetadata(GroupHealthBillingExtractionResult result, IReadOnlyList<PdfPageText> pages)
    {
        var firstPageLines = pages.FirstOrDefault()?.Lines ?? [];
        result.ClientName = ResolveClientName(firstPageLines);
        result.ClientNumber = ResolveClientNumberFromFileName(result.SourceFileName);

        var allText = string.Join(" ", pages.SelectMany(page => page.Lines));
        var periodMatch = Regex.Match(allText, @"BENEFITS?\s+BILLING\s+STATEMENT\s*-\s*(?<month>[A-Za-z]+),\s*(?<year>\d{4})", RegexOptions.IgnoreCase);
        if (periodMatch.Success)
        {
            result.StatementMonth = MonthNumber(periodMatch.Groups["month"].Value);
            result.StatementYear = ToInt(periodMatch.Groups["year"].Value) ?? DateTime.Today.Year;
            result.StatementPeriodStart = new DateOnly(result.StatementYear, result.StatementMonth, 1);
            result.StatementPeriodEnd = result.StatementPeriodStart.Value.AddMonths(1).AddDays(-1);
        }

        var printedMatch = Regex.Match(allText, @"Printed:\s*(?<date>[A-Za-z]+\s+\d{1,2},\s+\d{4})", RegexOptions.IgnoreCase);
        if (printedMatch.Success &&
            DateOnly.TryParseExact(printedMatch.Groups["date"].Value, "MMMM d, yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var printedDate))
        {
            result.StatementDate = printedDate;
        }

        result.PremiumGrandTotal = ParseLabeledAmount(allText, @"PREMIUM\s+GRAND\s+TOTAL\s+(?<amount>\(?-?[\d,]+\.\d{2}\)?)");
        result.GstHstGrandTotal = ParseLabeledAmount(allText, @"GST/HST\s+GRAND\s+TOTAL\s+(?<amount>\(?-?[\d,]+\.\d{2}\)?)");
        result.StatementGrandTotal = ParseLastLabeledAmount(allText, @"(?<!PREMIUM\s)(?<!GST/HST\s)GRAND\s+TOTAL\s+(?<amount>\(?-?[\d,]+\.\d{2}\)?)")
                                     ?? ParseLabeledAmount(allText, @"Balance\s+Due\s+\$?\s*(?<amount>\(?-?[\d,]+\.\d{2}\)?)");
    }

    private static string ResolveClientName(IReadOnlyList<string> firstPageLines)
    {
        foreach (var line in firstPageLines)
        {
            if (line.Contains("Benefits Billing Statement", StringComparison.OrdinalIgnoreCase) ||
                line.Contains("Printed:", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (line.StartsWith("STATEMENT RECONCILIATION", StringComparison.OrdinalIgnoreCase))
            {
                break;
            }

            if (!string.IsNullOrWhiteSpace(line))
            {
                return line.Trim();
            }
        }

        return "Unknown Client";
    }

    private static string ResolveClientNumberFromFileName(string fileName)
    {
        var match = Regex.Match(fileName, @"^(?<code>[A-Za-z0-9]+)_bill_", RegexOptions.IgnoreCase);
        return match.Success ? match.Groups["code"].Value.ToUpperInvariant() : "";
    }

    private static void ReadEmployeeRows(GroupHealthBillingExtractionResult result, IReadOnlyList<PdfPageText> pages)
    {
        foreach (var page in pages)
        {
            if (!page.Lines.Any(line => line.Contains("PIN ANNUAL CLASS/", StringComparison.OrdinalIgnoreCase)))
            {
                continue;
            }

            for (var index = 0; index < page.Lines.Count; index++)
            {
                var line = page.Lines[index];
                var match = PinLineRegex.Match(line);
                if (!match.Success)
                {
                    if (line.StartsWith("PAGE TOTALS", StringComparison.OrdinalIgnoreCase))
                    {
                        AddDetailTotals(result, "Page", page.PageNumber.ToString(CultureInfo.InvariantCulture), "Page Total", line, page.PageNumber);
                    }

                    continue;
                }

                var nextLine = index + 1 < page.Lines.Count ? page.Lines[index + 1] : "";
                if (!LooksLikePremiumLine(nextLine))
                {
                    continue;
                }

                var gstLine = index + 2 < page.Lines.Count && page.Lines[index + 2].StartsWith("GST/HST", StringComparison.OrdinalIgnoreCase)
                    ? page.Lines[index + 2]
                    : "";
                var totalLineIndex = string.IsNullOrWhiteSpace(gstLine) ? index + 2 : index + 3;
                var totalLine = totalLineIndex < page.Lines.Count && page.Lines[totalLineIndex].StartsWith("Total", StringComparison.OrdinalIgnoreCase)
                    ? page.Lines[totalLineIndex]
                    : "";

                var row = ParseEmployeeRow(match.Groups["pin"].Value, match.Groups["rest"].Value, nextLine, gstLine, totalLine, page.PageNumber, index + 1);
                if (row is not null)
                {
                    result.Rows.Add(row);
                }
            }
        }
    }

    private static bool LooksLikePremiumLine(string line)
    {
        if (string.IsNullOrWhiteSpace(line) ||
            line.StartsWith("PAGE TOTALS", StringComparison.OrdinalIgnoreCase) ||
            line.StartsWith("PIN ", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        return Regex.IsMatch(line, @"\d+\.\d{2}|waived", RegexOptions.IgnoreCase);
    }

    private static GroupHealthBillingRow? ParseEmployeeRow(
        string employeeId,
        string headerRest,
        string premiumLine,
        string gstLine,
        string totalLine,
        int pageNumber,
        int sourceLine)
    {
        var headerTokens = SplitTokens(headerRest);
        if (headerTokens.Count < 3)
        {
            return null;
        }

        var row = new GroupHealthBillingRow
        {
            EmployeeId = employeeId,
            AnnualEarnings = ParseDecimal(headerTokens[0]),
            ClassCode = headerTokens[1],
            EmploymentClass = headerTokens[1],
            CostDivision = headerTokens[2],
            SourcePage = pageNumber,
            SourceLine = sourceLine,
            SourceText = $"{employeeId} {headerRest} | {premiumLine} | {gstLine} | {totalLine}".Trim()
        };

        var headerIndex = 3;
        if (headerTokens.Count > headerIndex + 1 &&
            MonthNumbers.ContainsKey(headerTokens[headerIndex]) &&
            Regex.IsMatch(headerTokens[headerIndex + 1], @"^\d{2}$"))
        {
            row.RetroMonth = $"{headerTokens[headerIndex]} {headerTokens[headerIndex + 1]}";
            headerIndex += 2;
        }

        var coverage = headerTokens.Skip(headerIndex).ToList();
        var coverageIndex = 0;
        row.LifeVolume = TakeDecimal(coverage, ref coverageIndex);
        row.AdAndDVolume = TakeDecimal(coverage, ref coverageIndex);
        if (coverageIndex < coverage.Count && IsNumericToken(coverage[coverageIndex]))
        {
            row.LtdVolume = TakeDecimal(coverage, ref coverageIndex);
        }

        if (coverageIndex < coverage.Count)
        {
            row.HealthTier = coverage[coverageIndex++];
        }

        if (coverageIndex < coverage.Count)
        {
            row.DentalTier = coverage[coverageIndex++];
        }

        row.MemberAdAndDVolume = TakeDecimal(coverage, ref coverageIndex);
        row.MemberAdAndDSpVolume = TakeDecimal(coverage, ref coverageIndex);
        row.MemberAdAndDChVolume = TakeDecimal(coverage, ref coverageIndex);

        var premiumValues = ParsePremiumLine(premiumLine, out var employeeName, out var certificateId);
        row.EmployeeName = employeeName;
        row.CertificateId = string.IsNullOrWhiteSpace(certificateId) ? employeeId : certificateId;
        ApplyDetailAmounts(row, premiumValues);

        var gstValues = ParseNumberTokens(gstLine);
        row.GstHst = gstValues.LastOrDefault();

        var totalValues = ParseNumberTokens(totalLine);
        if (totalValues.Count == 12)
        {
            row.RowTotal = totalValues[11];
            row.LifePremium = totalValues[0];
            row.AdAndDPremium = totalValues[1];
            row.LtdPremium = totalValues[2];
            row.HealthPremium = totalValues[3];
            row.DentalPremium = totalValues[4];
            row.FlexAdminFee = totalValues[5];
            row.MemberAdAndDPremium = totalValues[6];
            row.MemberAdAndDSpPremium = totalValues[7];
            row.MemberAdAndDChPremium = totalValues[8];
            row.EmployeePremium = totalValues[9];
            row.EmployerPremium = totalValues[10] - (row.GstHst ?? 0m);
        }
        else
        {
            row.RowTotal = row.TotalPremium;
        }

        if (row.GstHst is null && row.TotalPremium is not null && row.RowTotal is not null)
        {
            row.GstHst = row.RowTotal.Value - row.TotalPremium.Value;
        }

        if (!string.IsNullOrWhiteSpace(row.RetroMonth))
        {
            row.RetroPremium = row.TotalPremium;
            row.RetroGstHst = row.GstHst;
        }

        return row;
    }

    private static List<decimal> ParsePremiumLine(string premiumLine, out string employeeName, out string certificateId)
    {
        employeeName = "";
        certificateId = "";
        var match = CertificateLineRegex.Match(premiumLine);
        if (match.Success)
        {
            employeeName = match.Groups["name"].Value.Trim();
            certificateId = match.Groups["certificate"].Value.Trim();
            return ParseNumberTokens(match.Groups["values"].Value);
        }

        var tokens = SplitTokens(premiumLine);
        var firstAmountIndex = tokens.FindIndex(token => IsNumericToken(token) || token.Equals("waived", StringComparison.OrdinalIgnoreCase));
        if (firstAmountIndex < 0)
        {
            employeeName = premiumLine.Trim();
            return [];
        }

        employeeName = string.Join(" ", tokens.Take(firstAmountIndex));
        return tokens
            .Skip(firstAmountIndex)
            .Select(token => token.Equals("waived", StringComparison.OrdinalIgnoreCase) ? 0m : ParseDecimal(token))
            .Where(value => value.HasValue)
            .Select(value => value!.Value)
            .ToList();
    }

    private static void ApplyDetailAmounts(GroupHealthBillingRow row, IReadOnlyList<decimal> values)
    {
        if (values.Count != 12)
        {
            return;
        }

        row.LifePremium = values[0];
        row.AdAndDPremium = values[1];
        row.LtdPremium = values[2];
        row.HealthPremium = values[3];
        row.DentalPremium = values[4];
        row.FlexAdminFee = values[5];
        row.MemberAdAndDPremium = values[6];
        row.MemberAdAndDSpPremium = values[7];
        row.MemberAdAndDChPremium = values[8];
        row.EmployeePremium = values[9];
        row.EmployerPremium = values[10];
        row.TotalPremium = values[11];
    }

    private static void AddDetailTotals(
        GroupHealthBillingExtractionResult result,
        string totalLevel,
        string classCode,
        string totalType,
        string line,
        int sourcePage)
    {
        var amounts = ParseNumberTokens(line);
        for (var index = 0; index < amounts.Count && index < DetailBenefits.Length; index++)
        {
            result.Totals.Add(new GroupHealthBillingStatementTotal
            {
                TotalLevel = totalLevel,
                ClassCode = classCode,
                Benefit = DetailBenefits[index],
                TotalType = totalType,
                Premium = amounts[index],
                TotalPremium = amounts[index],
                SourcePage = sourcePage
            });
        }
    }

    private static void ReadClassAndStatementTotals(GroupHealthBillingExtractionResult result, IReadOnlyList<PdfPageText> pages)
    {
        foreach (var page in pages.Where(page => page.Lines.Any(line => line.Contains("Curr Lvls", StringComparison.OrdinalIgnoreCase))))
        {
            string currentClass = "";
            foreach (var line in page.Lines)
            {
                var classMatch = Regex.Match(line, @"^(?<class>TOTAL|\S+)\s+Curr\s+Lvls\s*:", RegexOptions.IgnoreCase);
                if (classMatch.Success)
                {
                    currentClass = classMatch.Groups["class"].Value;
                    continue;
                }

                var totalType = SummaryTotalType(line);
                if (string.IsNullOrWhiteSpace(currentClass) || string.IsNullOrWhiteSpace(totalType))
                {
                    continue;
                }

                var amounts = ParseNumberTokens(TextAfterLastColon(line));
                for (var index = 0; index < amounts.Count && index < SummaryBenefits.Length; index++)
                {
                    result.Totals.Add(new GroupHealthBillingStatementTotal
                    {
                        TotalLevel = currentClass.Equals("TOTAL", StringComparison.OrdinalIgnoreCase) ? "Grand Total" : "Class",
                        ClassCode = currentClass,
                        Benefit = SummaryBenefits[index],
                        TotalType = totalType,
                        Premium = totalType.Contains("GST/HST", StringComparison.OrdinalIgnoreCase) ? null : amounts[index],
                        GstHst = totalType.Contains("GST/HST", StringComparison.OrdinalIgnoreCase) ? amounts[index] : null,
                        TotalPremium = amounts[index],
                        SourcePage = page.PageNumber
                    });
                }
            }
        }
    }

    private static string SummaryTotalType(string line)
    {
        if (line.Contains("Curr Prem:", StringComparison.OrdinalIgnoreCase))
        {
            return "Current Premium";
        }

        if (line.Contains("Curr GST/HST:", StringComparison.OrdinalIgnoreCase))
        {
            return "Current GST/HST";
        }

        if (line.Contains("Retro Prem:", StringComparison.OrdinalIgnoreCase))
        {
            return "Retro Premium";
        }

        if (line.Contains("Retr GST/HST:", StringComparison.OrdinalIgnoreCase))
        {
            return "Retro GST/HST";
        }

        if (line.Contains("Totl Prem:", StringComparison.OrdinalIgnoreCase))
        {
            return "Total Premium";
        }

        return "";
    }

    private static string TextAfterLastColon(string line)
    {
        var index = line.LastIndexOf(':');
        return index >= 0 ? line[(index + 1)..] : line;
    }

    private static void ReadChangeReport(GroupHealthBillingExtractionResult result, IReadOnlyList<PdfPageText> pages)
    {
        foreach (var page in pages.Where(page => page.Lines.Any(line => line.StartsWith("SUMMARY CHANGE REPORT", StringComparison.OrdinalIgnoreCase))))
        {
            foreach (var line in page.Lines)
            {
                var match = ChangeReportRegex.Match(line);
                if (!match.Success)
                {
                    continue;
                }

                result.EmployeeChangeLog.Add(new GroupHealthBillingEmployeeChange
                {
                    Carrier = result.Carrier,
                    StatementMonth = result.StatementMonth,
                    StatementYear = result.StatementYear,
                    EmployeeId = match.Groups["pin"].Value.Trim(),
                    EmployeeName = match.Groups["name"].Value.Trim(),
                    ChangeType = ChangeTypeLabel(match.Groups["change"].Value.Trim()),
                    FieldName = "Summary Change Report",
                    CurrentValue = line.Trim(),
                    Severity = GroupHealthBillingValidationStatuses.Information,
                    Source = $"PDF page {page.PageNumber}"
                });
            }
        }
    }

    private static string ChangeTypeLabel(string code)
    {
        var normalized = code.Trim().ToUpperInvariant();
        if (normalized.StartsWith("C-", StringComparison.OrdinalIgnoreCase))
        {
            return "Class Change";
        }

        return normalized switch
        {
            "A" => "New Employee",
            "S" => "Salary Change",
            "C" => "Class Change",
            "T" => "Termination",
            "D" => "Delete",
            "R" => "Reinstatement",
            "ADRS" => "Address Change",
            "REPT" => "Report Change",
            "S/OCC" => "Occupation Change",
            "AGE BAND SHIFT" => "Age Band Change",
            _ => code
        };
    }

    private static List<string> SplitTokens(string value)
    {
        return value.Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).ToList();
    }

    private static decimal? TakeDecimal(IReadOnlyList<string> tokens, ref int index)
    {
        if (index >= tokens.Count || !IsNumericToken(tokens[index]))
        {
            return null;
        }

        return ParseDecimal(tokens[index++]);
    }

    private static bool IsNumericToken(string token)
    {
        return Regex.IsMatch(token.Trim().TrimEnd('*'), @"^-?\(?[\d,]+(?:\.\d+)?\)?$");
    }

    private static List<decimal> ParseNumberTokens(string value)
    {
        return SplitTokens(value)
            .Select(token => token.TrimEnd('*'))
            .Select(ParseDecimal)
            .Where(amount => amount.HasValue)
            .Select(amount => amount!.Value)
            .ToList();
    }

    private static decimal? ParseDecimal(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        var cleaned = value.Trim().TrimEnd('*').Replace("$", "", StringComparison.OrdinalIgnoreCase);
        var negative = cleaned.StartsWith("(", StringComparison.OrdinalIgnoreCase) && cleaned.EndsWith(")", StringComparison.OrdinalIgnoreCase);
        cleaned = cleaned.Trim('(', ')').Replace(",", "", StringComparison.OrdinalIgnoreCase);
        return decimal.TryParse(cleaned, NumberStyles.Number | NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out var parsed)
            ? negative ? -parsed : parsed
            : null;
    }

    private static int? ToInt(string value)
    {
        return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed) ? parsed : null;
    }

    private static int MonthNumber(string month)
    {
        return MonthNumbers.TryGetValue(month, out var number) ? number : 1;
    }

    private static decimal? ParseLabeledAmount(string text, string pattern)
    {
        var match = Regex.Match(text, pattern, RegexOptions.IgnoreCase);
        return match.Success ? ParseDecimal(match.Groups["amount"].Value) : null;
    }

    private static decimal? ParseLastLabeledAmount(string text, string pattern)
    {
        var matches = Regex.Matches(text, pattern, RegexOptions.IgnoreCase);
        return matches.Count > 0 ? ParseDecimal(matches[^1].Groups["amount"].Value) : null;
    }

    private static string ComputeFileHash(string filePath)
    {
        using var stream = File.OpenRead(filePath);
        return Convert.ToHexString(SHA256.HashData(stream));
    }
}

public sealed class GroupHealthBillingValidator
{
    private readonly GroupHealthBillingValidationOptions _options;

    public GroupHealthBillingValidator(GroupHealthBillingValidationOptions? options = null)
    {
        _options = options ?? new GroupHealthBillingValidationOptions();
    }

    public void Validate(GroupHealthBillingExtractionResult result)
    {
        result.ValidationResults.Clear();

        AddCheck(result, "Statement", "Rows extracted", result.Rows.Count > 0 ? GroupHealthBillingValidationStatuses.Passed : GroupHealthBillingValidationStatuses.Failed,
            result.Rows.Count > 0 ? "Employee billing rows were extracted." : "No employee billing rows were extracted.");

        ValidateRows(result);
        ValidatePageTotals(result);
        ValidateStatementTotals(result);
        ValidateDuplicateRows(result);

        BuildIntegrityReport(result);
    }

    private void ValidateRows(GroupHealthBillingExtractionResult result)
    {
        foreach (var row in result.Rows)
        {
            var rowStatus = GroupHealthBillingValidationStatuses.Passed;
            var flags = new List<string>();

            if (string.IsNullOrWhiteSpace(row.EmployeeName) || string.IsNullOrWhiteSpace(row.CertificateId))
            {
                AddRowCheck(result, row, "Missing key fields", GroupHealthBillingValidationStatuses.Warning, "Employee name or certificate id is missing.");
                rowStatus = WorseStatus(rowStatus, GroupHealthBillingValidationStatuses.Warning);
                flags.Add("Missing value");
            }

            var eeErExpected = (row.EmployeePremium ?? 0m) + (row.EmployerPremium ?? 0m);
            AddAmountCheck(result, row, "Row", "EE + ER premium", row.TotalPremium, eeErExpected, "Employee premium plus employer premium should equal total premium.");
            if (Abs((row.TotalPremium ?? 0m) - eeErExpected) > _options.MoneyTolerance)
            {
                rowStatus = WorseStatus(rowStatus, GroupHealthBillingValidationStatuses.Failed);
                flags.Add("EE/ER mismatch");
            }

            var totalExpected = (row.TotalPremium ?? 0m) + (row.GstHst ?? 0m);
            AddAmountCheck(result, row, "Row", "Premium + GST/HST", row.RowTotal, totalExpected, "Premium plus GST/HST should equal row total.");
            if (Abs((row.RowTotal ?? 0m) - totalExpected) > _options.MoneyTolerance)
            {
                rowStatus = WorseStatus(rowStatus, GroupHealthBillingValidationStatuses.Failed);
                flags.Add("Tax mismatch");
            }

            if ((row.TotalPremium ?? 0m) < 0m || (row.RowTotal ?? 0m) < 0m)
            {
                var hasRetro = !string.IsNullOrWhiteSpace(row.RetroMonth);
                AddRowCheck(
                    result,
                    row,
                    "Negative premium",
                    hasRetro ? GroupHealthBillingValidationStatuses.Information : GroupHealthBillingValidationStatuses.Warning,
                    hasRetro ? "Negative premium is tied to a retro adjustment." : "Negative premium is not tied to a retro month.");
                if (!hasRetro)
                {
                    rowStatus = WorseStatus(rowStatus, GroupHealthBillingValidationStatuses.Warning);
                    flags.Add("Negative without retro");
                }
            }

            if ((row.MemberAdAndDPremium ?? 0m) > 0m && (row.MemberAdAndDVolume ?? 0m) <= 0m)
            {
                rowStatus = WorseStatus(rowStatus, GroupHealthBillingValidationStatuses.Warning);
                flags.Add("Member AD&D volume missing");
                AddRowCheck(result, row, "Member AD&D volume", GroupHealthBillingValidationStatuses.Warning, "Member AD&D premium exists without a matching coverage volume.");
            }

            row.ValidationStatus = rowStatus;
            row.ValidationFlags = string.Join("; ", flags);
        }
    }

    private void ValidatePageTotals(GroupHealthBillingExtractionResult result)
    {
        var pageTotalRows = result.Totals
            .Where(total => total.TotalLevel.Equals("Page", StringComparison.OrdinalIgnoreCase) &&
                            total.Benefit.Equals("Total Premium", StringComparison.OrdinalIgnoreCase))
            .ToList();

        foreach (var pageTotal in pageTotalRows)
        {
            if (!int.TryParse(pageTotal.ClassCode, NumberStyles.Integer, CultureInfo.InvariantCulture, out var pageNumber))
            {
                continue;
            }

            var calculated = result.Rows
                .Where(row => row.SourcePage == pageNumber)
                .Sum(row => row.RowTotal ?? 0m);
            AddAmountCheck(
                result,
                null,
                "Page",
                $"Page {pageNumber} total premium",
                pageTotal.TotalPremium,
                calculated,
                "Sum of employee row totals should equal the PDF page total.",
                sourcePage: pageNumber);
        }
    }

    private void ValidateStatementTotals(GroupHealthBillingExtractionResult result)
    {
        var totalPremium = result.Rows.Sum(row => row.TotalPremium ?? 0m);
        var totalTax = result.Rows.Sum(row => row.GstHst ?? 0m);
        var rowTotal = result.Rows.Sum(row => row.RowTotal ?? 0m);
        var totalEe = result.Rows.Sum(row => row.EmployeePremium ?? 0m);
        var totalErWithTax = result.Rows.Sum(row => (row.EmployerPremium ?? 0m) + (row.GstHst ?? 0m));

        AddAmountCheck(result, null, "Statement", "Premium grand total", result.PremiumGrandTotal, totalPremium, "Employee row premiums should equal the statement premium grand total.");
        AddAmountCheck(result, null, "Statement", "GST/HST grand total", result.GstHstGrandTotal, totalTax, "Employee row GST/HST should equal the statement GST/HST grand total.");
        AddAmountCheck(result, null, "Statement", "Grand total", result.StatementGrandTotal, rowTotal, "Employee row totals should equal the statement grand total.");

        var pdfEe = FindGrandTotal(result, "EE Premium");
        var pdfEr = FindGrandTotal(result, "ER Premium");
        AddAmountCheck(result, null, "Statement", "EE premium grand total", pdfEe, totalEe, "Employee paid premiums should equal the statement EE premium total.");
        AddAmountCheck(result, null, "Statement", "ER premium grand total", pdfEr, totalErWithTax, "Employer premiums plus GST/HST should equal the statement ER premium total.");

        var healthSummary = (FindGrandTotal(result, "Health") ?? 0m) +
                            (FindGrandTotal(result, "EFAP") ?? 0m) +
                            (FindGrandTotal(result, "Virtual Healthcare") ?? 0m);
        var healthRows = result.Rows.Sum(row => row.HealthPremium ?? 0m);
        AddAmountCheck(result, null, "Statement", "Health detail to split totals", healthSummary, healthRows, "The visible employee Health column should reconcile to Health + EFAP + Virtual Healthcare statement totals.");
    }

    private static decimal? FindGrandTotal(GroupHealthBillingExtractionResult result, string benefit)
    {
        return result.Totals
            .Where(total => total.TotalLevel.Equals("Grand Total", StringComparison.OrdinalIgnoreCase) &&
                            total.TotalType.Equals("Total Premium", StringComparison.OrdinalIgnoreCase) &&
                            total.Benefit.Equals(benefit, StringComparison.OrdinalIgnoreCase))
            .Select(total => total.TotalPremium)
            .FirstOrDefault();
    }

    private void ValidateDuplicateRows(GroupHealthBillingExtractionResult result)
    {
        var duplicateGroups = result.Rows
            .Where(row => !string.IsNullOrWhiteSpace(row.CertificateId))
            .GroupBy(row => row.CertificateId, StringComparer.OrdinalIgnoreCase)
            .Where(group => group.Count() > 1)
            .ToList();

        foreach (var group in duplicateGroups)
        {
            var allRetro = group.All(row => !string.IsNullOrWhiteSpace(row.RetroMonth)) ||
                           group.Any(row => !string.IsNullOrWhiteSpace(row.RetroMonth));
            AddCheck(
                result,
                "Employee",
                "Duplicate certificate rows",
                allRetro ? GroupHealthBillingValidationStatuses.Information : GroupHealthBillingValidationStatuses.Warning,
                $"{group.Key} appears {group.Count():#,0} times in the statement{(allRetro ? " because of retro/change rows." : ".")}",
                certificateId: group.Key);
        }
    }

    private void AddAmountCheck(
        GroupHealthBillingExtractionResult result,
        GroupHealthBillingRow? row,
        string level,
        string checkName,
        decimal? expected,
        decimal? actual,
        string message,
        int? sourcePage = null)
    {
        var difference = expected.HasValue && actual.HasValue ? actual.Value - expected.Value : (decimal?)null;
        var status = !expected.HasValue || !actual.HasValue
            ? GroupHealthBillingValidationStatuses.NotAvailable
            : Abs(difference!.Value) <= _options.MoneyTolerance
                ? GroupHealthBillingValidationStatuses.Passed
                : GroupHealthBillingValidationStatuses.Failed;

        result.ValidationResults.Add(new GroupHealthBillingValidationResult
        {
            BillingStatementRowId = row?.BillingStatementRowId,
            ValidationLevel = level,
            CheckName = checkName,
            Status = status,
            Message = status == GroupHealthBillingValidationStatuses.Passed ? message : $"{message} Difference: {difference:0.00}.",
            ExpectedAmount = expected,
            ActualAmount = actual,
            DifferenceAmount = difference,
            EmployeeId = row?.EmployeeId ?? "",
            CertificateId = row?.CertificateId ?? "",
            ClassCode = row?.ClassCode ?? "",
            SourcePage = sourcePage ?? row?.SourcePage
        });
    }

    private static void AddRowCheck(GroupHealthBillingExtractionResult result, GroupHealthBillingRow row, string checkName, string status, string message)
    {
        result.ValidationResults.Add(new GroupHealthBillingValidationResult
        {
            BillingStatementRowId = row.BillingStatementRowId,
            ValidationLevel = "Row",
            CheckName = checkName,
            Status = status,
            Message = message,
            EmployeeId = row.EmployeeId,
            CertificateId = row.CertificateId,
            ClassCode = row.ClassCode,
            SourcePage = row.SourcePage
        });
    }

    private static void AddCheck(
        GroupHealthBillingExtractionResult result,
        string level,
        string checkName,
        string status,
        string message,
        string certificateId = "")
    {
        result.ValidationResults.Add(new GroupHealthBillingValidationResult
        {
            ValidationLevel = level,
            CheckName = checkName,
            Status = status,
            Message = message,
            CertificateId = certificateId
        });
    }

    private void BuildIntegrityReport(GroupHealthBillingExtractionResult result)
    {
        var failed = result.ValidationResults.Count(row => row.Status.Equals(GroupHealthBillingValidationStatuses.Failed, StringComparison.OrdinalIgnoreCase));
        var warnings = result.ValidationResults.Count(row => row.Status.Equals(GroupHealthBillingValidationStatuses.Warning, StringComparison.OrdinalIgnoreCase));
        var passed = result.ValidationResults.Count(row => row.Status.Equals(GroupHealthBillingValidationStatuses.Passed, StringComparison.OrdinalIgnoreCase));
        var missingValues = result.ValidationResults.Count(row => row.CheckName.Contains("Missing", StringComparison.OrdinalIgnoreCase));
        var duplicateRows = result.ValidationResults.Count(row => row.CheckName.Contains("Duplicate", StringComparison.OrdinalIgnoreCase));
        var statementDifference = result.StatementGrandTotal.HasValue
            ? result.Rows.Sum(row => row.RowTotal ?? 0m) - result.StatementGrandTotal.Value
            : (decimal?)null;

        var score = 100m -
                    failed * _options.FailedDeduction -
                    warnings * _options.WarningDeduction -
                    missingValues * _options.MissingKeyFieldDeduction;
        score = score < 0m ? 0m : score > 100m ? 100m : score;

        var unreconciledStatementTotal = statementDifference.HasValue && Abs(statementDifference.Value) > _options.MoneyTolerance;
        var finalStatus = unreconciledStatementTotal || score < 90m || failed > 0
            ? GroupHealthBillingImportStatuses.Failed
            : score >= 98m
                ? GroupHealthBillingImportStatuses.Passed
                : GroupHealthBillingImportStatuses.Warning;

        result.BillingIntegrityScore = decimal.Round(score, 2);
        result.ImportStatus = finalStatus;
        result.IntegrityReport = new GroupHealthBillingIntegrityReport
        {
            TotalEmployeesImported = result.Rows.Select(row => string.IsNullOrWhiteSpace(row.CertificateId) ? row.EmployeeId : row.CertificateId).Distinct(StringComparer.OrdinalIgnoreCase).Count(),
            TotalPremiumsImported = result.Rows.Sum(row => row.RowTotal ?? 0m),
            StatementTotal = result.StatementGrandTotal,
            DifferenceAmount = statementDifference,
            PassedChecks = passed,
            Warnings = warnings,
            FailedChecks = failed,
            NewHires = result.EmployeeChangeLog.Count(row => row.ChangeType.Equals("New Employee", StringComparison.OrdinalIgnoreCase)),
            Terminations = result.EmployeeChangeLog.Count(row => row.ChangeType.Equals("Termination", StringComparison.OrdinalIgnoreCase)),
            SalaryChanges = result.EmployeeChangeLog.Count(row => row.ChangeType.Equals("Salary Change", StringComparison.OrdinalIgnoreCase)),
            CoverageChanges = result.EmployeeChangeLog.Count(row => row.ChangeType.Contains("Coverage", StringComparison.OrdinalIgnoreCase)),
            TierChanges = result.EmployeeChangeLog.Count(row => row.ChangeType.Contains("Tier", StringComparison.OrdinalIgnoreCase)),
            DuplicateRows = duplicateRows,
            MissingValues = missingValues,
            BillingIntegrityScore = result.BillingIntegrityScore,
            FinalStatus = finalStatus
        };
    }

    private static string WorseStatus(string current, string candidate)
    {
        return Rank(candidate) > Rank(current) ? candidate : current;
    }

    private static int Rank(string status)
    {
        return status switch
        {
            GroupHealthBillingValidationStatuses.Failed => 3,
            GroupHealthBillingValidationStatuses.Warning => 2,
            GroupHealthBillingValidationStatuses.Information => 1,
            _ => 0
        };
    }

    private static decimal Abs(decimal value)
    {
        return value < 0m ? -value : value;
    }
}

public static class GroupHealthBillingRateCalculator
{
    public static void BuildRateLines(GroupHealthBillingExtractionResult result)
    {
        result.RateLines.Clear();
        AddVolumeRate(result, "Life", row => row.LifeVolume, row => row.LifePremium, 1000m);
        AddVolumeRate(result, "AD&D", row => row.AdAndDVolume, row => row.AdAndDPremium, 1000m);
        AddVolumeRate(result, "LTD", row => row.LtdVolume, row => row.LtdPremium, 100m);
        AddTierRate(result, "Health", row => row.HealthTier, row => row.HealthPremium);
        AddTierRate(result, "Dental", row => row.DentalTier, row => row.DentalPremium);
        AddCountRate(result, "Flex Admin", row => row.FlexAdminFee);
        AddVolumeRate(result, "Member AD&D", row => row.MemberAdAndDVolume, row => row.MemberAdAndDPremium, 1000m);
        AddVolumeRate(result, "Member AD&D SP", row => row.MemberAdAndDSpVolume, row => row.MemberAdAndDSpPremium, 1000m);
        AddVolumeRate(result, "Member AD&D CH", row => row.MemberAdAndDChVolume, row => row.MemberAdAndDChPremium, 1000m);

        AddStatementSplitRate(result, "EFAP");
        AddStatementSplitRate(result, "Virtual Healthcare");
    }

    private static void AddVolumeRate(
        GroupHealthBillingExtractionResult result,
        string benefit,
        Func<GroupHealthBillingRow, decimal?> volumeSelector,
        Func<GroupHealthBillingRow, decimal?> premiumSelector,
        decimal volumeUnit)
    {
        var groups = result.Rows
            .GroupBy(row => row.ClassCode, StringComparer.OrdinalIgnoreCase)
            .Where(group => group.Sum(row => volumeSelector(row) ?? 0m) > 0m)
            .ToList();

        foreach (var group in groups)
        {
            var volume = group.Sum(row => volumeSelector(row) ?? 0m);
            var premium = group.Sum(row => premiumSelector(row) ?? 0m);
            var rate = volume > 0m ? premium / (volume / volumeUnit) : (decimal?)null;
            AddRateLine(result, benefit, group.Key, "", $"per ${volumeUnit:#,0} volume", rate);
        }
    }

    private static void AddTierRate(
        GroupHealthBillingExtractionResult result,
        string benefit,
        Func<GroupHealthBillingRow, string> tierSelector,
        Func<GroupHealthBillingRow, decimal?> premiumSelector)
    {
        var groups = result.Rows
            .Where(row => !string.IsNullOrWhiteSpace(tierSelector(row)) && !tierSelector(row).Equals("0", StringComparison.OrdinalIgnoreCase))
            .GroupBy(row => new { ClassCode = row.ClassCode.ToUpperInvariant(), Tier = tierSelector(row).ToUpperInvariant() });

        foreach (var group in groups)
        {
            var count = group.Count();
            var premium = group.Sum(row => premiumSelector(row) ?? 0m);
            AddRateLine(result, benefit, group.Key.ClassCode, group.Key.Tier, "per employee by tier", count > 0 ? premium / count : null);
        }
    }

    private static void AddCountRate(
        GroupHealthBillingExtractionResult result,
        string benefit,
        Func<GroupHealthBillingRow, decimal?> premiumSelector)
    {
        var groups = result.Rows.GroupBy(row => row.ClassCode, StringComparer.OrdinalIgnoreCase);
        foreach (var group in groups)
        {
            var count = group.Count(row => (premiumSelector(row) ?? 0m) != 0m);
            if (count <= 0)
            {
                continue;
            }

            AddRateLine(result, benefit, group.Key, "", "per billed employee", group.Sum(row => premiumSelector(row) ?? 0m) / count);
        }
    }

    private static void AddStatementSplitRate(GroupHealthBillingExtractionResult result, string benefit)
    {
        var premium = result.Totals
            .Where(row => row.TotalLevel.Equals("Grand Total", StringComparison.OrdinalIgnoreCase) &&
                          row.TotalType.Equals("Total Premium", StringComparison.OrdinalIgnoreCase) &&
                          row.Benefit.Equals(benefit, StringComparison.OrdinalIgnoreCase))
            .Select(row => row.TotalPremium)
            .FirstOrDefault();
        var employees = result.IntegrityReport.TotalEmployeesImported > 0 ? result.IntegrityReport.TotalEmployeesImported : result.Rows.Count;
        if (!premium.HasValue || employees <= 0)
        {
            return;
        }

        AddRateLine(result, benefit, "All", "", "statement total per employee", premium.Value / employees);
    }

    private static void AddRateLine(GroupHealthBillingExtractionResult result, string benefit, string classCode, string tier, string basis, decimal? rate)
    {
        result.RateLines.Add(new GroupHealthClientRateTableLine
        {
            Carrier = result.Carrier,
            EffectiveMonth = result.StatementMonth,
            EffectiveYear = result.StatementYear,
            Benefit = benefit,
            ClassCode = classCode,
            Tier = tier,
            VolumeBasis = basis,
            TotalRate = rate,
            CalculatedRate = rate,
            RateStatus = rate.HasValue ? GroupHealthBillingValidationStatuses.Passed : GroupHealthBillingValidationStatuses.NotAvailable,
            IsCalculated = true
        });
    }
}

public sealed class GroupHealthBillingChangeDetector
{
    private readonly GroupHealthBillingValidationOptions _options;

    public GroupHealthBillingChangeDetector(GroupHealthBillingValidationOptions? options = null)
    {
        _options = options ?? new GroupHealthBillingValidationOptions();
    }

    public List<GroupHealthBillingEmployeeChange> Detect(
        IReadOnlyCollection<GroupHealthBillingRow> previousRows,
        IReadOnlyCollection<GroupHealthBillingRow> currentRows,
        string carrier,
        int statementMonth,
        int statementYear)
    {
        var changes = new List<GroupHealthBillingEmployeeChange>();
        var previous = CollapseCurrentRows(previousRows).ToDictionary(KeyForRow, StringComparer.OrdinalIgnoreCase);
        var current = CollapseCurrentRows(currentRows).ToDictionary(KeyForRow, StringComparer.OrdinalIgnoreCase);

        foreach (var row in current.Values)
        {
            if (!previous.TryGetValue(KeyForRow(row), out var prior))
            {
                changes.Add(CreateChange(row, carrier, statementMonth, statementYear, "New Employee", "Employee", "", row.EmployeeName));
                continue;
            }

            Compare(changes, row, prior, carrier, statementMonth, statementYear, "Salary Change", "AnnualEarnings", prior.AnnualEarnings, row.AnnualEarnings);
            Compare(changes, row, prior, carrier, statementMonth, statementYear, "Class Change", "ClassCode", prior.ClassCode, row.ClassCode);
            Compare(changes, row, prior, carrier, statementMonth, statementYear, "Health Tier Change", "HealthTier", prior.HealthTier, row.HealthTier);
            Compare(changes, row, prior, carrier, statementMonth, statementYear, "Dental Tier Change", "DentalTier", prior.DentalTier, row.DentalTier);
            Compare(changes, row, prior, carrier, statementMonth, statementYear, "Coverage Change", "LifeVolume", prior.LifeVolume, row.LifeVolume);
            Compare(changes, row, prior, carrier, statementMonth, statementYear, "Coverage Change", "AD&DVolume", prior.AdAndDVolume, row.AdAndDVolume);
            Compare(changes, row, prior, carrier, statementMonth, statementYear, "Coverage Change", "LTDVolume", prior.LtdVolume, row.LtdVolume);
            Compare(changes, row, prior, carrier, statementMonth, statementYear, "Coverage Change", "MemberAD&DVolume", prior.MemberAdAndDVolume, row.MemberAdAndDVolume);

            var previousPremium = prior.RowTotal ?? 0m;
            var currentPremium = row.RowTotal ?? 0m;
            if (previousPremium != 0m)
            {
                var changePercent = (currentPremium - previousPremium) / previousPremium * 100m;
                if (Abs(changePercent) > _options.PremiumChangeThresholdPercent)
                {
                    changes.Add(CreateChange(row, carrier, statementMonth, statementYear, "Premium Change", "RowTotal", previousPremium.ToString("0.00", CultureInfo.InvariantCulture), currentPremium.ToString("0.00", CultureInfo.InvariantCulture), currentPremium - previousPremium, changePercent, GroupHealthBillingValidationStatuses.Warning));
                }
            }
        }

        foreach (var row in previous.Values.Where(row => !current.ContainsKey(KeyForRow(row))))
        {
            changes.Add(CreateChange(row, carrier, statementMonth, statementYear, "Termination", "Employee", row.EmployeeName, ""));
        }

        return changes;
    }

    private static IEnumerable<GroupHealthBillingRow> CollapseCurrentRows(IReadOnlyCollection<GroupHealthBillingRow> rows)
    {
        return rows
            .Where(row => string.IsNullOrWhiteSpace(row.RetroMonth))
            .GroupBy(KeyForRow, StringComparer.OrdinalIgnoreCase)
            .Select(group => group.OrderByDescending(row => row.RowTotal ?? 0m).First());
    }

    private static string KeyForRow(GroupHealthBillingRow row)
    {
        return !string.IsNullOrWhiteSpace(row.CertificateId)
            ? row.CertificateId
            : !string.IsNullOrWhiteSpace(row.EmployeeId)
                ? row.EmployeeId
                : row.EmployeeName;
    }

    private static void Compare(
        List<GroupHealthBillingEmployeeChange> changes,
        GroupHealthBillingRow current,
        GroupHealthBillingRow previous,
        string carrier,
        int statementMonth,
        int statementYear,
        string changeType,
        string fieldName,
        decimal? previousValue,
        decimal? currentValue)
    {
        if (previousValue == currentValue)
        {
            return;
        }

        changes.Add(CreateChange(
            current,
            carrier,
            statementMonth,
            statementYear,
            changeType,
            fieldName,
            previousValue?.ToString("0.##", CultureInfo.InvariantCulture) ?? "",
            currentValue?.ToString("0.##", CultureInfo.InvariantCulture) ?? "",
            currentValue - previousValue));
    }

    private static void Compare(
        List<GroupHealthBillingEmployeeChange> changes,
        GroupHealthBillingRow current,
        GroupHealthBillingRow previous,
        string carrier,
        int statementMonth,
        int statementYear,
        string changeType,
        string fieldName,
        string previousValue,
        string currentValue)
    {
        if (string.Equals(previousValue, currentValue, StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        changes.Add(CreateChange(current, carrier, statementMonth, statementYear, changeType, fieldName, previousValue, currentValue));
    }

    private static GroupHealthBillingEmployeeChange CreateChange(
        GroupHealthBillingRow row,
        string carrier,
        int statementMonth,
        int statementYear,
        string changeType,
        string fieldName,
        string previousValue,
        string currentValue,
        decimal? difference = null,
        decimal? changePercent = null,
        string severity = GroupHealthBillingValidationStatuses.Information)
    {
        return new GroupHealthBillingEmployeeChange
        {
            Carrier = carrier,
            StatementMonth = statementMonth,
            StatementYear = statementYear,
            EmployeeId = row.EmployeeId,
            CertificateId = row.CertificateId,
            EmployeeName = row.EmployeeName,
            ChangeType = changeType,
            FieldName = fieldName,
            PreviousValue = previousValue,
            CurrentValue = currentValue,
            DifferenceAmount = difference,
            ChangePercent = changePercent,
            Severity = severity,
            Source = "Month-over-month billing comparison"
        };
    }

    private static decimal Abs(decimal value)
    {
        return value < 0m ? -value : value;
    }
}

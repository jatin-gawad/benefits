using System.Globalization;
using System.Text.RegularExpressions;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed class ManulifeBillingStatementParser : IManulifeBillingStatementParser
{
    private readonly PdfTextExtractor _extractor = new();
    private readonly ManulifeBillingValidationOptions _options;

    private static readonly string[] ProvinceNames =
    [
        "British Columbia",
        "Alberta",
        "Saskatchewan",
        "Manitoba",
        "Ontario",
        "Quebec",
        "Nova Scotia",
        "Prince Edward Island",
        "PEI",
        "New Brunswick",
        "Newfoundland & Labrador",
        "Newfoundland and Labrador",
        "Newfoundland"
    ];

    private static readonly Dictionary<string, int> MonthNumbers = new(StringComparer.OrdinalIgnoreCase)
    {
        ["JAN"] = 1,
        ["FEB"] = 2,
        ["MAR"] = 3,
        ["APR"] = 4,
        ["MAY"] = 5,
        ["JUN"] = 6,
        ["JUL"] = 7,
        ["AUG"] = 8,
        ["SEP"] = 9,
        ["SEPT"] = 9,
        ["OCT"] = 10,
        ["NOV"] = 11,
        ["DEC"] = 12
    };

    public ManulifeBillingStatementParser(ManulifeBillingValidationOptions? options = null)
    {
        _options = options ?? new ManulifeBillingValidationOptions();
    }

    public BillingExtractionResult Extract(string pdfPath)
    {
        if (string.IsNullOrWhiteSpace(pdfPath) || !File.Exists(pdfPath))
        {
            throw new FileNotFoundException("Billing statement PDF was not found.", pdfPath);
        }

        var pages = _extractor.Extract(pdfPath);
        var result = new BillingExtractionResult
        {
            SourceFile = Path.GetFullPath(pdfPath),
            PageCount = pages.Count,
            ReadablePdfExtraction = pages.Count > 0 && pages.Any(page => page.Lines.Count > 0)
        };

        ReadStatementMetadata(result, pages);

        var sections = ReadBillingSections(result, pages);
        foreach (var detail in sections.SelectMany(section => section.Details))
        {
            detail.BillingPeriod = result.BillingPeriod;
            detail.SourceFile = result.SourceFile;
            result.Details.Add(detail);
        }

        BuildSummaries(result, sections);
        BuildClassMappings(result);
        BuildReconciliationResults(result, sections);
        result.ReconciliationExecuted = true;
        result.ImportStatus = BillingImportStatuses.NeedsReview;

        return result;
    }

    public static string OptionLabelForClassCode(string classCode)
    {
        var cleaned = (classCode ?? "").Trim();
        return cleaned.ToUpperInvariant() switch
        {
            "01" => "Option 1",
            "02" => "Option 2",
            "03" => "Option 3",
            "BB" or "B" => "Executive",
            "UU" or "U" => "Union",
            "RR" or "R" => "Retiree",
            "" => "Class",
            _ => $"Class {cleaned}"
        };
    }

    private static void ReadStatementMetadata(BillingExtractionResult result, IReadOnlyList<PdfPageText> pages)
    {
        var firstLines = pages.FirstOrDefault()?.Lines ?? [];
        result.BillingPeriod = FindBillingPeriod(firstLines);
        (result.PeriodStart, result.PeriodEnd) = ParseBillingPeriodDates(result.BillingPeriod);
        result.BillingId = FindFirstMatch(pages, @"\bBilling Id\s+(?<value>\S+)", "value");
        result.ClientNumber = result.BillingId;
        result.ClientName = FindClientName(firstLines, result.SourceFile, out var clientNameConfirmed);
        result.ClientIdentityConfirmed = clientNameConfirmed && !string.IsNullOrWhiteSpace(result.BillingId);
        result.CarrierIdentityConfirmed = pages
            .SelectMany(page => page.Lines)
            .Any(line => line.Contains("Manulife", StringComparison.OrdinalIgnoreCase));

        var firstPageText = string.Join(" ", firstLines);
        result.PdfGrandPremiumBeforeTax = ParseLabeledAmount(firstPageText, @"Your Premium Due\s+\$?\s*(?<amount>\(?[\d,]+\.\d{2}\)?)");
        result.PdfGrandTax = ParseLabeledAmount(firstPageText, @"Your Taxes Due\s+\$?\s*(?<amount>\(?[\d,]+\.\d{2}\)?)");
        result.PdfGrandTotalWithTax = ParseLabeledAmount(firstPageText, @"Your Current Month Amount Due\s+\$?\s*(?<amount>\(?[\d,]+\.\d{2}\)?)");
    }

    private static string FindBillingPeriod(IReadOnlyList<string> lines)
    {
        for (var index = 0; index < lines.Count; index++)
        {
            var line = Clean(lines[index]);
            if (line.StartsWith("Billing period:", StringComparison.OrdinalIgnoreCase))
            {
                var inline = line["Billing period:".Length..].Trim();
                if (!string.IsNullOrWhiteSpace(inline))
                {
                    return inline;
                }

                if (index + 1 < lines.Count)
                {
                    return Clean(lines[index + 1]);
                }
            }
        }

        return "";
    }

    private static (DateOnly? Start, DateOnly? End) ParseBillingPeriodDates(string billingPeriod)
    {
        var match = Regex.Match(
            billingPeriod ?? "",
            @"(?<m1>[A-Z]{3,4})\s+(?<d1>\d{1,2})\s+(?<y1>\d{4})\s+to\s+(?<m2>[A-Z]{3,4})\s+(?<d2>\d{1,2})\s+(?<y2>\d{4})",
            RegexOptions.IgnoreCase);
        if (!match.Success)
        {
            return (null, null);
        }

        var start = TryBuildDate(match.Groups["y1"].Value, match.Groups["m1"].Value, match.Groups["d1"].Value);
        var end = TryBuildDate(match.Groups["y2"].Value, match.Groups["m2"].Value, match.Groups["d2"].Value);
        return (start, end);
    }

    private static DateOnly? TryBuildDate(string yearText, string monthText, string dayText)
    {
        if (!int.TryParse(yearText, NumberStyles.Integer, CultureInfo.InvariantCulture, out var year) ||
            !MonthNumbers.TryGetValue(monthText, out var month) ||
            !int.TryParse(dayText, NumberStyles.Integer, CultureInfo.InvariantCulture, out var day))
        {
            return null;
        }

        return new DateOnly(year, month, day);
    }

    private static string FindClientName(IReadOnlyList<string> lines, string sourceFile, out bool confirmed)
    {
        confirmed = false;
        for (var index = 0; index < lines.Count; index++)
        {
            if (!lines[index].Contains("Payment due:", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            for (var candidateIndex = index + 1; candidateIndex < Math.Min(lines.Count, index + 8); candidateIndex++)
            {
                var candidate = Clean(lines[candidateIndex]);
                if (string.IsNullOrWhiteSpace(candidate) ||
                    candidate.Contains("Questions?", StringComparison.OrdinalIgnoreCase) ||
                    candidate.StartsWith("Attn:", StringComparison.OrdinalIgnoreCase) ||
                    candidate.StartsWith("Internet:", StringComparison.OrdinalIgnoreCase) ||
                    candidate.StartsWith("Customer Service:", StringComparison.OrdinalIgnoreCase) ||
                    candidate.StartsWith("Write:", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                confirmed = true;
                return candidate;
            }
        }

        return Path.GetFileNameWithoutExtension(sourceFile);
    }

    private static string FindFirstMatch(IEnumerable<PdfPageText> pages, string pattern, string groupName)
    {
        foreach (var line in pages.SelectMany(page => page.Lines))
        {
            var match = Regex.Match(line, pattern, RegexOptions.IgnoreCase);
            if (match.Success)
            {
                return Clean(match.Groups[groupName].Value);
            }
        }

        return "";
    }

    private static decimal? ParseLabeledAmount(string text, string pattern)
    {
        var match = Regex.Match(text ?? "", pattern, RegexOptions.IgnoreCase);
        return match.Success && TryParseAmount(match.Groups["amount"].Value, allowWholeNumber: false, out var amount)
            ? amount
            : null;
    }

    private static List<BillingSection> ReadBillingSections(BillingExtractionResult result, IReadOnlyList<PdfPageText> pages)
    {
        var sections = new List<BillingSection>();
        BillingSection? current = null;
        string previousLine = "";

        foreach (var page in pages)
        {
            foreach (var rawLine in page.Lines)
            {
                var line = Clean(rawLine);
                if (string.IsNullOrWhiteSpace(line))
                {
                    continue;
                }

                if (line.StartsWith("Class ", StringComparison.OrdinalIgnoreCase))
                {
                    if (current is not null &&
                        current.Header.Equals(line, StringComparison.OrdinalIgnoreCase) &&
                        current.HasMonthlyTotal &&
                        !current.HasTaxTotal)
                    {
                        previousLine = line;
                        continue;
                    }

                    if (current is not null)
                    {
                        sections.Add(current);
                    }

                    current = new BillingSection(line, ExtractProvince(line), page.PageNumber);
                    previousLine = line;
                    continue;
                }

                if (current is null)
                {
                    previousLine = line;
                    continue;
                }

                if (line.StartsWith("Total Monthly Summary", StringComparison.OrdinalIgnoreCase))
                {
                    current.ReadMonthlyTotal(line, previousLine);
                    previousLine = line;
                    continue;
                }

                if (current.PendingSplitTotal && line.StartsWith("$", StringComparison.Ordinal))
                {
                    current.TryCompleteSplitMonthlyTotal(line);
                    previousLine = line;
                    continue;
                }

                if (line.StartsWith("Total Tax Summary", StringComparison.OrdinalIgnoreCase))
                {
                    current.ReadTaxSummary(line);
                    previousLine = line;
                    continue;
                }

                if (!current.HasMonthlyTotal && TryParseDetailRow(line, current.Province, page.PageNumber, out var detail))
                {
                    current.Details.Add(detail);
                    previousLine = line;
                    continue;
                }

                if (!current.HasMonthlyTotal && current.Details.Count > 0 && IsBenefitContinuation(line))
                {
                    current.Details[^1].Benefit = Clean($"{current.Details[^1].Benefit} {line}");
                }

                previousLine = line;
            }
        }

        if (current is not null)
        {
            sections.Add(current);
        }

        if (sections.Count == 0)
        {
            result.Warnings.Add("No Manulife billing detail sections were found.");
        }

        return sections;
    }

    private static string ExtractProvince(string classHeader)
    {
        foreach (var province in ProvinceNames)
        {
            var match = Regex.Match(classHeader, $@"\b{Regex.Escape(province)}\b", RegexOptions.IgnoreCase);
            if (match.Success)
            {
                return match.Value;
            }
        }

        return "No province";
    }

    private static bool TryParseDetailRow(string line, string province, int sourcePage, out BillingOptionPremiumDetail detail)
    {
        detail = new BillingOptionPremiumDetail();

        var tokens = line
            .Split(' ', StringSplitOptions.RemoveEmptyEntries)
            .Where(token => token != "$")
            .ToList();
        if (tokens.Count < 10)
        {
            return false;
        }

        var valueTokens = tokens.Skip(tokens.Count - 5).Take(5).ToList();
        if (valueTokens.Any(token => !IsMoneyToken(token, allowWholeNumber: false)))
        {
            return false;
        }

        var prefix = tokens.Take(tokens.Count - 5).ToList();
        if (prefix.Count < 5)
        {
            return false;
        }

        var benefit = Clean(string.Join(" ", prefix.Take(prefix.Count - 4)));
        var dependentCode = Clean(prefix[^4]);
        var classCode = Clean(prefix[^3]);
        var livesText = prefix[^2];
        var volumeText = prefix[^1];

        if (string.IsNullOrWhiteSpace(benefit) ||
            benefit.Equals("Benefit", StringComparison.OrdinalIgnoreCase) ||
            !Regex.IsMatch(classCode, @"^[A-Za-z0-9]{1,8}$"))
        {
            return false;
        }

        detail = new BillingOptionPremiumDetail
        {
            Province = province,
            Benefit = benefit,
            DependentCode = dependentCode,
            ClassCode = classCode,
            OptionLabel = OptionLabelForClassCode(classCode),
            Lives = ParseNullableInt(livesText),
            Volume = ParseNullableAmount(volumeText, allowWholeNumber: true),
            MonthlyPremium = ParseAmountOrZero(valueTokens[0]),
            PremiumAdjustments = ParseAmountOrZero(valueTokens[1]),
            TotalPremiumAdjustments = ParseAmountOrZero(valueTokens[2]),
            SourcePage = sourcePage,
            Confidence = 0.90m
        };

        return true;
    }

    private static bool IsBenefitContinuation(string line)
    {
        if (!Regex.IsMatch(line, @"^[A-Za-z][A-Za-z &/+\-]*$"))
        {
            return false;
        }

        var blockedPrefixes = new[]
        {
            "Monthly",
            "Billing",
            "Benefit",
            "Adjustments",
            "Tax Summary",
            "GST/HST",
            "Billing Id",
            "Class",
            "Sales Tax",
            "Manitoba",
            "Ontario",
            "Quebec",
            "Saskatchewan",
            "Total"
        };

        return blockedPrefixes.All(prefix => !line.StartsWith(prefix, StringComparison.OrdinalIgnoreCase));
    }

    private static void BuildSummaries(BillingExtractionResult result, IReadOnlyList<BillingSection> sections)
    {
        foreach (var summary in result.Details
                     .GroupBy(row => new { row.Province, row.Benefit, row.ClassCode, row.OptionLabel })
                     .Select(group =>
                     {
                         var monthlyPremium = group.Sum(row => row.TotalPremiumAdjustments);
                         return new BillingOptionSummary
                         {
                             BillingPeriod = result.BillingPeriod,
                             Province = group.Key.Province,
                             Benefit = group.Key.Benefit,
                             ClassCode = group.Key.ClassCode,
                             OptionLabel = group.Key.OptionLabel,
                             Lives = group.Sum(row => row.Lives ?? 0),
                             MonthlyPremium = monthlyPremium,
                             EstimatedAnnualPremium = monthlyPremium * 12m,
                             SourceFile = result.SourceFile
                         };
                     })
                     .OrderBy(row => row.Province, StringComparer.OrdinalIgnoreCase)
                     .ThenBy(row => row.Benefit, StringComparer.OrdinalIgnoreCase)
                     .ThenBy(row => row.ClassCode, StringComparer.OrdinalIgnoreCase))
        {
            result.OptionSummaries.Add(summary);
        }

        foreach (var summary in result.Details
                     .GroupBy(row => row.Province, StringComparer.OrdinalIgnoreCase)
                     .Select(group =>
                     {
                         var totalMonthlyPremium = group.Sum(row => row.TotalPremiumAdjustments);
                         var totalTax = sections
                             .Where(section => section.Province.Equals(group.Key, StringComparison.OrdinalIgnoreCase))
                             .Sum(section => section.PdfTax ?? 0m);
                         return new BillingProvinceSummary
                         {
                             BillingPeriod = result.BillingPeriod,
                             Province = group.Key,
                             TotalLives = group.Sum(row => row.Lives ?? 0),
                             TotalMonthlyPremium = totalMonthlyPremium,
                             EstimatedAnnualPremium = totalMonthlyPremium * 12m,
                             TotalTax = totalTax,
                             TotalPremiumWithTax = totalMonthlyPremium + totalTax,
                             SourceFile = result.SourceFile
                         };
                     })
                     .OrderBy(row => row.Province, StringComparer.OrdinalIgnoreCase))
        {
            result.ProvinceSummaries.Add(summary);
        }
    }

    private static void BuildClassMappings(BillingExtractionResult result)
    {
        foreach (var mapping in result.Details
                     .GroupBy(row => row.ClassCode, StringComparer.OrdinalIgnoreCase)
                     .Select(group => new BillingClassMapping
                     {
                         Carrier = result.Carrier,
                         ClassCode = group.Key,
                         OptionLabel = OptionLabelForClassCode(group.Key),
                         FirstSeenBillingPeriod = result.BillingPeriod,
                         LastSeenBillingPeriod = result.BillingPeriod,
                         SourceFile = result.SourceFile
                     })
                     .OrderBy(row => row.ClassCode, StringComparer.OrdinalIgnoreCase))
        {
            result.ClassMappings.Add(mapping);
        }
    }

    private void BuildReconciliationResults(BillingExtractionResult result, IReadOnlyList<BillingSection> sections)
    {
        foreach (var section in sections)
        {
            var calculatedPremium = section.Details.Sum(row => row.TotalPremiumAdjustments);
            var calculatedLives = section.Details.Sum(row => row.Lives ?? 0);
            result.ReconciliationResults.Add(CreateReconciliation(
                result,
                "Row-level section total",
                section.Province,
                "",
                "",
                null,
                calculatedLives,
                section.PdfTotalPremiumBeforeTax,
                calculatedPremium,
                section.PdfTax,
                section.PdfTax,
                section.PdfTotalWithTax,
                calculatedPremium + (section.PdfTax ?? 0m),
                section.SourcePage,
                section.HasMonthlyTotal ? "" : "PDF section total was not available."));
        }

        foreach (var group in result.Details.GroupBy(row => new { row.Province, row.Benefit }))
        {
            result.ReconciliationResults.Add(new BillingReconciliationResult
            {
                BillingPeriod = result.BillingPeriod,
                SourceFile = result.SourceFile,
                ReconciliationLevel = "Province benefit subtotal",
                Province = group.Key.Province,
                Benefit = group.Key.Benefit,
                CalculatedLives = group.Sum(row => row.Lives ?? 0),
                CalculatedPremium = group.Sum(row => row.TotalPremiumAdjustments),
                Status = BillingReconciliationStatuses.NotAvailable,
                WarningMessage = "PDF benefit subtotal was not available in the extracted section.",
                SourcePage = group.Min(row => row.SourcePage),
                CreatedAt = DateTime.Now
            });
        }

        foreach (var group in result.Details.GroupBy(row => new { row.Province, row.Benefit, row.ClassCode }))
        {
            result.ReconciliationResults.Add(new BillingReconciliationResult
            {
                BillingPeriod = result.BillingPeriod,
                SourceFile = result.SourceFile,
                ReconciliationLevel = "Class subtotal",
                Province = group.Key.Province,
                Benefit = group.Key.Benefit,
                ClassCode = group.Key.ClassCode,
                CalculatedLives = group.Sum(row => row.Lives ?? 0),
                CalculatedPremium = group.Sum(row => row.TotalPremiumAdjustments),
                Status = BillingReconciliationStatuses.NotAvailable,
                WarningMessage = "PDF class subtotal was not available in the extracted section.",
                SourcePage = group.Min(row => row.SourcePage),
                CreatedAt = DateTime.Now
            });
        }

        foreach (var province in result.ProvinceSummaries)
        {
            var provinceSections = sections
                .Where(section => section.Province.Equals(province.Province, StringComparison.OrdinalIgnoreCase))
                .ToList();
            var pdfPremium = SumNullable(provinceSections.Select(section => section.PdfTotalPremiumBeforeTax));
            var pdfTax = SumNullable(provinceSections.Select(section => section.PdfTax));
            var pdfTotal = SumNullable(provinceSections.Select(section => section.PdfTotalWithTax));

            result.ReconciliationResults.Add(CreateReconciliation(
                result,
                "Province total",
                province.Province,
                "",
                "",
                null,
                province.TotalLives,
                pdfPremium,
                province.TotalMonthlyPremium,
                pdfTax,
                province.TotalTax,
                pdfTotal,
                province.TotalPremiumWithTax,
                provinceSections.Count > 0 ? provinceSections.Min(section => section.SourcePage) : null,
                ""));

            result.ReconciliationResults.Add(CreateReconciliation(
                result,
                "Tax validation",
                province.Province,
                "",
                "",
                null,
                null,
                null,
                null,
                pdfTax,
                province.TotalTax,
                pdfTotal,
                province.TotalMonthlyPremium + province.TotalTax,
                provinceSections.Count > 0 ? provinceSections.Min(section => section.SourcePage) : null,
                ""));
        }

        var calculatedGrandPremium = result.ProvinceSummaries.Sum(row => row.TotalMonthlyPremium);
        var calculatedGrandTax = result.ProvinceSummaries.Sum(row => row.TotalTax);
        result.ReconciliationResults.Add(CreateReconciliation(
            result,
            "Grand total",
            "",
            "",
            "",
            null,
            result.ProvinceSummaries.Sum(row => row.TotalLives),
            result.PdfGrandPremiumBeforeTax,
            calculatedGrandPremium,
            result.PdfGrandTax,
            calculatedGrandTax,
            result.PdfGrandTotalWithTax,
            calculatedGrandPremium + calculatedGrandTax,
            1,
            ""));

        var ordered = result.ReconciliationResults
            .OrderBy(row => StatusSort(row.Status))
            .ThenBy(row => row.ReconciliationLevel, StringComparer.OrdinalIgnoreCase)
            .ThenBy(row => row.Province, StringComparer.OrdinalIgnoreCase)
            .ThenBy(row => row.Benefit, StringComparer.OrdinalIgnoreCase)
            .ThenBy(row => row.ClassCode, StringComparer.OrdinalIgnoreCase)
            .ToList();
        result.ReconciliationResults.Clear();
        result.ReconciliationResults.AddRange(ordered);
    }

    private BillingReconciliationResult CreateReconciliation(
        BillingExtractionResult result,
        string level,
        string province,
        string benefit,
        string classCode,
        int? pdfLives,
        int? calculatedLives,
        decimal? pdfPremium,
        decimal? calculatedPremium,
        decimal? pdfTax,
        decimal? calculatedTax,
        decimal? pdfTotalWithTax,
        decimal? calculatedTotalWithTax,
        int? sourcePage,
        string unavailableMessage)
    {
        var lives = ManulifeBillingLivesReconciliation.Evaluate(pdfLives, calculatedLives, _options);
        var livesDifference = lives.Difference;
        var premiumDifference = Difference(calculatedPremium, pdfPremium);
        var taxDifference = Difference(calculatedTax, pdfTax);
        var totalDifference = Difference(calculatedTotalWithTax, pdfTotalWithTax);
        var livesStatus = lives.Status;
        var livesMessage = lives.Message;
        var anyMoneyPdfValue = pdfPremium.HasValue || pdfTax.HasValue || pdfTotalWithTax.HasValue;
        var anyComparisonValue = anyMoneyPdfValue || livesStatus != BillingReconciliationStatuses.NotAvailable;

        var messages = new List<string>();
        if (!anyComparisonValue)
        {
            messages.Add(string.IsNullOrWhiteSpace(unavailableMessage) ? "PDF comparison value was not available." : unavailableMessage);
        }

        if (ExceedsTolerance(premiumDifference))
        {
            messages.Add($"Premium differs by {premiumDifference:0.00}.");
        }

        if (ExceedsTolerance(taxDifference))
        {
            messages.Add($"Tax differs by {taxDifference:0.00}.");
        }

        if (ExceedsTolerance(totalDifference))
        {
            messages.Add($"Total including tax differs by {totalDifference:0.00}.");
        }

        if (livesStatus is BillingReconciliationStatuses.Warning or BillingReconciliationStatuses.Failed)
        {
            messages.Add(livesMessage);
        }

        var status = !anyComparisonValue
            ? BillingReconciliationStatuses.NotAvailable
            : livesStatus == BillingReconciliationStatuses.Failed ||
              ExceedsTolerance(premiumDifference) ||
              ExceedsTolerance(taxDifference) ||
              ExceedsTolerance(totalDifference)
                ? BillingReconciliationStatuses.Failed
                : livesStatus == BillingReconciliationStatuses.Warning
                    ? BillingReconciliationStatuses.Warning
                    : BillingReconciliationStatuses.Passed;

        return new BillingReconciliationResult
        {
            BillingPeriod = result.BillingPeriod,
            SourceFile = result.SourceFile,
            ReconciliationLevel = level,
            Province = province,
            Benefit = benefit,
            ClassCode = classCode,
            PdfLives = pdfLives,
            CalculatedLives = calculatedLives,
            LivesDifference = livesDifference,
            LivesStatus = livesStatus,
            LivesWarningMessage = livesMessage,
            PdfPremium = pdfPremium,
            CalculatedPremium = calculatedPremium,
            PremiumDifference = premiumDifference,
            PdfTax = pdfTax,
            CalculatedTax = calculatedTax,
            TaxDifference = taxDifference,
            PdfTotalWithTax = pdfTotalWithTax,
            CalculatedTotalWithTax = calculatedTotalWithTax,
            TotalDifference = totalDifference,
            Status = status,
            WarningMessage = string.Join(" ", messages),
            SourcePage = sourcePage,
            CreatedAt = DateTime.Now
        };
    }

    private static decimal? Difference(decimal? calculated, decimal? pdf)
    {
        return calculated.HasValue && pdf.HasValue
            ? calculated.Value - pdf.Value
            : null;
    }

    private bool ExceedsTolerance(decimal? value)
    {
        return value.HasValue && Math.Abs(value.Value) > _options.MoneyTolerance;
    }

    private static decimal? SumNullable(IEnumerable<decimal?> values)
    {
        var items = values.Where(value => value.HasValue).Select(value => value!.Value).ToList();
        return items.Count == 0 ? null : items.Sum();
    }

    private static int StatusSort(string status)
    {
        return status switch
        {
            BillingReconciliationStatuses.Failed => 0,
            BillingReconciliationStatuses.Warning => 1,
            BillingReconciliationStatuses.Passed => 2,
            _ => 3
        };
    }

    private static int? ParseNullableInt(string value)
    {
        var cleaned = (value ?? "").Replace(",", "", StringComparison.Ordinal).Trim();
        return string.IsNullOrWhiteSpace(cleaned) || cleaned == "-"
            ? null
            : int.TryParse(cleaned, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed)
                ? parsed
                : null;
    }

    private static decimal? ParseNullableAmount(string value, bool allowWholeNumber)
    {
        return TryParseAmount(value, allowWholeNumber, out var parsed)
            ? parsed
            : null;
    }

    private static decimal ParseAmountOrZero(string value)
    {
        return TryParseAmount(value, allowWholeNumber: false, out var parsed) ? parsed : 0m;
    }

    private static bool TryParseAmount(string value, bool allowWholeNumber, out decimal parsed)
    {
        parsed = 0m;
        var cleaned = (value ?? "").Trim();
        if (cleaned == "-")
        {
            return true;
        }

        cleaned = cleaned.Replace("$", "", StringComparison.Ordinal).Replace(",", "", StringComparison.Ordinal);
        var isNegative = cleaned.StartsWith("(", StringComparison.Ordinal) && cleaned.EndsWith(")", StringComparison.Ordinal);
        cleaned = cleaned.Trim('(', ')');

        var style = NumberStyles.AllowDecimalPoint | NumberStyles.AllowLeadingSign;
        if (!allowWholeNumber && !cleaned.Contains('.', StringComparison.Ordinal))
        {
            return false;
        }

        if (!decimal.TryParse(cleaned, style, CultureInfo.InvariantCulture, out parsed))
        {
            return false;
        }

        if (isNegative)
        {
            parsed *= -1m;
        }

        return true;
    }

    private static bool IsMoneyToken(string token, bool allowWholeNumber)
    {
        return token == "-" ||
               Regex.IsMatch(
                   token ?? "",
                   allowWholeNumber
                       ? @"^\(?-?\d[\d,]*(?:\.\d+)?\)?$"
                       : @"^\(?-?\d[\d,]*\.\d+\)?$");
    }

    private static List<string> AmountTokens(string line, bool allowWholeNumber)
    {
        return (line ?? "")
            .Split(' ', StringSplitOptions.RemoveEmptyEntries)
            .Where(token => token != "$")
            .Where(token => IsMoneyToken(token, allowWholeNumber))
            .ToList();
    }

    private static string Clean(string value)
    {
        return Regex.Replace(value ?? "", @"\s+", " ").Trim();
    }

    private sealed class BillingSection
    {
        private List<decimal>? _pendingSplitTotal;

        public BillingSection(string header, string province, int sourcePage)
        {
            Header = header;
            Province = province;
            SourcePage = sourcePage;
        }

        public string Header { get; }
        public string Province { get; }
        public int SourcePage { get; }
        public List<BillingOptionPremiumDetail> Details { get; } = [];
        public decimal? PdfMonthlyPremium { get; private set; }
        public decimal? PdfPremiumAdjustments { get; private set; }
        public decimal? PdfTotalPremiumBeforeTax { get; private set; }
        public decimal? PdfTax { get; private set; }
        public decimal? PdfTotalWithTax { get; private set; }
        public bool HasMonthlyTotal => PdfMonthlyPremium.HasValue || _pendingSplitTotal is not null;
        public bool HasTaxTotal { get; private set; }
        public bool PendingSplitTotal => _pendingSplitTotal is { Count: 3 };

        public void ReadMonthlyTotal(string line, string previousLine)
        {
            var tokens = AmountTokens(line, allowWholeNumber: true);
            if (tokens.Count >= 5)
            {
                if (IsDecimalFragment(previousLine) && Regex.IsMatch(tokens[0], @"^\d{1,2}$"))
                {
                    var combined = $"{previousLine.Trim()}{tokens[0]}";
                    ApplyMonthlyTotal(
                        ParseAmountOrZero(combined),
                        ParseAmountOrZero(tokens[1]),
                        ParseAmountOrZero(tokens[2]),
                        ParseAmountOrZero(tokens[3]),
                        ParseAmountOrZero(tokens[4]));
                    return;
                }

                ApplyMonthlyTotal(
                    ParseAmountOrZero(tokens[0]),
                    ParseAmountOrZero(tokens[1]),
                    ParseAmountOrZero(tokens[2]),
                    ParseAmountOrZero(tokens[3]),
                    ParseAmountOrZero(tokens[4]));
                return;
            }

            if (tokens.Count == 3)
            {
                _pendingSplitTotal = tokens.Select(ParseAmountOrZero).ToList();
            }
        }

        public void TryCompleteSplitMonthlyTotal(string line)
        {
            if (_pendingSplitTotal is not { Count: 3 })
            {
                return;
            }

            var tokens = AmountTokens(line, allowWholeNumber: true);
            if (tokens.Count < 2)
            {
                return;
            }

            ApplyMonthlyTotal(
                _pendingSplitTotal[0],
                _pendingSplitTotal[1],
                ParseAmountOrZero(tokens[0]),
                _pendingSplitTotal[2],
                ParseAmountOrZero(tokens[1]));
            _pendingSplitTotal = null;
        }

        public void ReadTaxSummary(string line)
        {
            var tokens = AmountTokens(line, allowWholeNumber: true);
            if (tokens.Count > 0)
            {
                PdfTax = ParseAmountOrZero(tokens[^1]);
            }

            HasTaxTotal = true;
        }

        private void ApplyMonthlyTotal(
            decimal monthlyPremium,
            decimal premiumAdjustments,
            decimal totalPremiumBeforeTax,
            decimal tax,
            decimal totalWithTax)
        {
            PdfMonthlyPremium = monthlyPremium;
            PdfPremiumAdjustments = premiumAdjustments;
            PdfTotalPremiumBeforeTax = totalPremiumBeforeTax;
            PdfTax = tax;
            PdfTotalWithTax = totalWithTax;
        }

        private static bool IsDecimalFragment(string value)
        {
            return Regex.IsMatch((value ?? "").Trim(), @"^\d[\d,]*\.\d$");
        }
    }
}

using System.Globalization;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

namespace RenewalReviewBuilder.App.Services;

public sealed class RevisedRateScheduleImporter
{
    private static readonly string[] SupportedExtensions = [".pdf", ".xlsx", ".xlsm"];
    private static readonly CultureInfo CanadaCulture = CultureInfo.GetCultureInfo("en-CA");
    private readonly Func<string, CarrierPdfParseResult> _parsePdf;

    public RevisedRateScheduleImporter(Func<string, CarrierPdfParseResult>? parsePdf = null)
    {
        _parsePdf = parsePdf ?? (path => new CarrierPdfParser().Parse(path));
    }

    public RevisedRateScheduleImportResult Import(string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
        {
            return RevisedRateScheduleImportResult.Failure(["Choose an existing revised carrier rate sheet."]);
        }

        var extension = Path.GetExtension(filePath);
        if (!SupportedExtensions.Contains(extension, StringComparer.OrdinalIgnoreCase))
        {
            return RevisedRateScheduleImportResult.Failure(["Revised rate sheets must be a PDF, .xlsx, or .xlsm file."]);
        }

        return extension.Equals(".pdf", StringComparison.OrdinalIgnoreCase)
            ? ImportPdf(filePath)
            : ImportWorkbook(filePath);
    }

    private RevisedRateScheduleImportResult ImportPdf(string filePath)
    {
        var parsed = _parsePdf(filePath);
        var rows = parsed.Pdf.RateLines
            .Where(line => line.RenewalRate.HasValue || line.RenewalPremium.HasValue ||
                           line.CarrierProposedRate.HasValue || line.CarrierProposedPremium.HasValue)
            .Select(line =>
            {
                var identity = CanonicalIdentity(line.Benefit, line.TierOrClass);
                return new RevisedRateScheduleRow
                {
                    SourceBenefit = line.Benefit.Trim(),
                    CanonicalBenefit = identity.Benefit,
                    TierOrClass = identity.TierOrClass,
                    SourceClass = line.AccountClass.Trim(),
                    Volume = line.Volume,
                    Lives = line.Lives,
                    CurrentRate = line.CurrentRate,
                    CurrentPremium = line.CurrentPremium,
                    RevisedRate = line.RenewalRate ?? line.CarrierProposedRate,
                    RevisedPremium = line.RenewalPremium ?? line.CarrierProposedPremium,
                    SourceRow = SourcePage(line.SourceReference),
                    SourceReference = line.SourceReference
                };
            })
            .Where(row => !IsTotalRow(row.SourceBenefit))
            .ToList();

        if (rows.Count == 0)
        {
            return RevisedRateScheduleImportResult.Failure(
            [
                "No revised rate or premium rows were found in this PDF. Confirm that it is a supported carrier renewal or revised rate document."
            ]);
        }

        var errors = ValidateRows(rows, "PDF");
        if (errors.Count > 0)
        {
            return RevisedRateScheduleImportResult.Failure(errors);
        }

        return RevisedRateScheduleImportResult.Success(new RevisedRateSchedule
        {
            SourceFilePath = Path.GetFullPath(filePath),
            SourceFileName = Path.GetFileName(filePath),
            SourceKind = "PDF",
            SheetName = "PDF",
            Carrier = FirstNonBlank(parsed.Carrier, parsed.Pdf.Carrier),
            ClientName = parsed.Pdf.ClientName,
            PolicyNumber = parsed.Pdf.PolicyNumber,
            RenewalDate = parsed.Pdf.RenewalDate,
            Rows = rows,
            Warnings = parsed.Pdf.Warnings.ToList()
        });
    }

    private static RevisedRateScheduleImportResult ImportWorkbook(string filePath)
    {
        using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        using var workbook = new XLWorkbook(stream);
        var detectionErrors = new List<string>();

        foreach (var worksheet in workbook.Worksheets)
        {
            var result = TryReadWorksheet(filePath, worksheet);
            if (result.Succeeded)
            {
                return result;
            }

            if (result.IsDetected)
            {
                detectionErrors.AddRange(result.Errors);
            }
        }

        return detectionErrors.Count > 0
            ? RevisedRateScheduleImportResult.Failure(detectionErrors)
            : RevisedRateScheduleImportResult.Failure(
            [
                "No revised rate table was found. Expected a Benefit column plus a Revised, Negotiated, Final, Renewal, Proposed, or New Rate/Monthly Premium column."
            ]);
    }

    private static RevisedRateScheduleImportResult TryReadWorksheet(string filePath, IXLWorksheet worksheet)
    {
        var used = worksheet.RangeUsed();
        if (used is null)
        {
            return RevisedRateScheduleImportResult.NotDetected();
        }

        var firstRow = used.RangeAddress.FirstAddress.RowNumber;
        var lastRow = used.RangeAddress.LastAddress.RowNumber;
        var firstColumn = used.RangeAddress.FirstAddress.ColumnNumber;
        var lastColumn = used.RangeAddress.LastAddress.ColumnNumber;

        for (var headerRow = firstRow; headerRow <= Math.Min(lastRow, firstRow + 35); headerRow++)
        {
            var benefitColumn = FindColumn(worksheet, headerRow, firstColumn, lastColumn, IsBenefitHeader);
            if (benefitColumn == 0)
            {
                continue;
            }

            var detection = DetectValueColumns(worksheet, headerRow, firstColumn, lastColumn);
            if (!detection.HasValueColumns)
            {
                continue;
            }

            if (detection.HasNetAndGrossColumns)
            {
                return RevisedRateScheduleImportResult.Failure(
                [
                    $"The sheet '{worksheet.Name}' contains separate Net and Gross rate columns. Use 'Load Net/Gross Rate Schedule' for this file."
                ]);
            }

            if (detection.Errors.Count > 0)
            {
                return RevisedRateScheduleImportResult.Failure(detection.Errors);
            }

            var tierColumn = FindColumn(worksheet, headerRow, firstColumn, lastColumn, IsTierHeader);
            var classColumn = FindColumn(worksheet, headerRow, firstColumn, lastColumn, IsClassHeader);
            var rows = new List<RevisedRateScheduleRow>();
            var errors = new List<string>();
            var currentBenefit = "";

            for (var rowNumber = headerRow + 1; rowNumber <= lastRow; rowNumber++)
            {
                var rawBenefit = CellText(worksheet.Cell(rowNumber, benefitColumn));
                var rawTier = tierColumn > 0 ? CellText(worksheet.Cell(rowNumber, tierColumn)) : "";
                var rawClass = classColumn > 0 ? CellText(worksheet.Cell(rowNumber, classColumn)) : "";
                var rate = 0m;
                var premium = 0m;
                var hasRate = detection.RateColumn > 0 && TryReadDecimal(worksheet.Cell(rowNumber, detection.RateColumn), out rate);
                var hasPremium = detection.PremiumColumn > 0 && TryReadDecimal(worksheet.Cell(rowNumber, detection.PremiumColumn), out premium);

                if (!hasRate && !hasPremium)
                {
                    if (!string.IsNullOrWhiteSpace(rawBenefit) && !IsTotalRow(rawBenefit) && !IsTierLabel(rawBenefit))
                    {
                        currentBenefit = rawBenefit;
                    }

                    continue;
                }

                if (hasRate && rate < 0m)
                {
                    errors.Add($"{worksheet.Name} row {rowNumber}: revised rate cannot be negative.");
                    continue;
                }

                if (hasPremium && premium < 0m)
                {
                    errors.Add($"{worksheet.Name} row {rowNumber}: revised premium cannot be negative.");
                    continue;
                }

                var sourceBenefit = rawBenefit;
                if (IsTierLabel(rawBenefit) && !string.IsNullOrWhiteSpace(currentBenefit))
                {
                    rawTier = FirstNonBlank(rawTier, rawBenefit);
                    sourceBenefit = currentBenefit;
                }
                else if (string.IsNullOrWhiteSpace(sourceBenefit))
                {
                    sourceBenefit = currentBenefit;
                }
                else if (!IsTotalRow(sourceBenefit))
                {
                    currentBenefit = sourceBenefit;
                }

                if (string.IsNullOrWhiteSpace(sourceBenefit) || IsTotalRow(sourceBenefit))
                {
                    continue;
                }

                var identity = CanonicalIdentity(sourceBenefit, FirstNonBlank(rawTier, rawClass));
                rows.Add(new RevisedRateScheduleRow
                {
                    SourceBenefit = sourceBenefit,
                    CanonicalBenefit = identity.Benefit,
                    TierOrClass = identity.TierOrClass,
                    SourceClass = rawClass,
                    RevisedRate = hasRate ? rate : null,
                    RevisedPremium = hasPremium ? premium : null,
                    SourceRow = rowNumber,
                    SourceReference = $"{Path.GetFileName(filePath)} | {worksheet.Name} | row {rowNumber}"
                });
            }

            errors.AddRange(ValidateRows(rows, worksheet.Name));
            if (rows.Count == 0)
            {
                errors.Add($"The sheet '{worksheet.Name}' contains revised rate headers but no usable benefit rows.");
            }

            if (errors.Count > 0)
            {
                return RevisedRateScheduleImportResult.Failure(errors.Distinct().ToList());
            }

            var metadata = ReadMetadata(worksheet, firstRow, lastRow, firstColumn, lastColumn);
            return RevisedRateScheduleImportResult.Success(new RevisedRateSchedule
            {
                SourceFilePath = Path.GetFullPath(filePath),
                SourceFileName = Path.GetFileName(filePath),
                SourceKind = "Excel",
                SheetName = worksheet.Name,
                Carrier = metadata.Carrier,
                ClientName = metadata.ClientName,
                PolicyNumber = metadata.PolicyNumber,
                RenewalDate = metadata.RenewalDate,
                Rows = rows
            });
        }

        return RevisedRateScheduleImportResult.NotDetected();
    }

    private static ValueColumnDetection DetectValueColumns(
        IXLWorksheet worksheet,
        int headerRow,
        int firstColumn,
        int lastColumn)
    {
        var candidates = new List<ValueColumnCandidate>();
        var headerTexts = new List<string>();
        for (var column = firstColumn; column <= lastColumn; column++)
        {
            var leaf = CellText(worksheet.Cell(headerRow, column));
            var parent = headerRow > 1
                ? NearestStageHeader(worksheet, headerRow - 1, column, firstColumn)
                : "";
            var combined = FirstNonBlank(parent, leaf).Equals(leaf, StringComparison.OrdinalIgnoreCase)
                ? leaf
                : $"{parent} {leaf}";
            var normalized = Normalize(combined);
            headerTexts.Add(normalized);

            var score = StageScore(normalized);
            if (score <= 0)
            {
                continue;
            }

            if (IsRateValueHeader(normalized))
            {
                candidates.Add(new ValueColumnCandidate(column, ValueColumnKind.Rate, score, combined));
            }
            else if (IsPremiumValueHeader(normalized))
            {
                candidates.Add(new ValueColumnCandidate(column, ValueColumnKind.Premium, score, combined));
            }
        }

        var errors = new List<string>();
        var rateColumn = PickColumn(candidates, ValueColumnKind.Rate, errors);
        var premiumColumn = PickColumn(candidates, ValueColumnKind.Premium, errors);
        return new ValueColumnDetection
        {
            RateColumn = rateColumn,
            PremiumColumn = premiumColumn,
            HasNetAndGrossColumns = headerTexts.Any(text => text.Contains("netrate", StringComparison.Ordinal)) &&
                                         headerTexts.Any(text => text.Contains("grossrate", StringComparison.Ordinal)),
            Errors = errors
        };
    }

    private static int PickColumn(
        IReadOnlyList<ValueColumnCandidate> candidates,
        ValueColumnKind kind,
        ICollection<string> errors)
    {
        var matches = candidates.Where(candidate => candidate.Kind == kind).OrderByDescending(candidate => candidate.Score).ToList();
        if (matches.Count == 0)
        {
            return 0;
        }

        var top = matches.Where(candidate => candidate.Score == matches[0].Score).ToList();
        if (top.Count > 1)
        {
            errors.Add($"More than one equally ranked revised {kind.ToString().ToLowerInvariant()} column was found: {string.Join(", ", top.Select(item => item.Header))}.");
            return 0;
        }

        return top[0].Column;
    }

    private static string NearestStageHeader(IXLWorksheet worksheet, int row, int column, int firstColumn)
    {
        var mergedHeader = worksheet.MergedRanges.FirstOrDefault(range =>
            range.RangeAddress.FirstAddress.RowNumber <= row &&
            range.RangeAddress.LastAddress.RowNumber >= row &&
            range.RangeAddress.FirstAddress.ColumnNumber <= column &&
            range.RangeAddress.LastAddress.ColumnNumber >= column);
        if (mergedHeader is not null)
        {
            var mergedValue = CellText(mergedHeader.FirstCell());
            var normalizedMergedValue = Normalize(mergedValue);
            if (StageScore(normalizedMergedValue) > 0 || normalizedMergedValue.Contains("current", StringComparison.Ordinal))
            {
                return mergedValue;
            }
        }

        for (var candidate = column; candidate >= firstColumn; candidate--)
        {
            var value = CellText(worksheet.Cell(row, candidate));
            var normalized = Normalize(value);
            if (StageScore(normalized) > 0 || normalized.Contains("current", StringComparison.Ordinal))
            {
                return value;
            }
        }

        return "";
    }

    private static int StageScore(string normalized)
    {
        if (normalized.Contains("negotiated", StringComparison.Ordinal) ||
            normalized.Contains("revised", StringComparison.Ordinal) ||
            normalized.Contains("finalapproved", StringComparison.Ordinal) ||
            normalized.Contains("finalrate", StringComparison.Ordinal) ||
            normalized.Contains("finalpremium", StringComparison.Ordinal))
        {
            return 30;
        }

        if (normalized.Contains("renewal", StringComparison.Ordinal) ||
            normalized.Contains("proposed", StringComparison.Ordinal) ||
            normalized.Contains("newrate", StringComparison.Ordinal) ||
            normalized.Contains("newpremium", StringComparison.Ordinal))
        {
            return 20;
        }

        return normalized.Contains("current", StringComparison.Ordinal) ? 0 : 10;
    }

    private static bool IsRateValueHeader(string normalized) =>
        normalized.Contains("rate", StringComparison.Ordinal) &&
        !normalized.Contains("ratebasis", StringComparison.Ordinal) &&
        !normalized.Contains("change", StringComparison.Ordinal) &&
        !normalized.Contains("variance", StringComparison.Ordinal) &&
        !normalized.Contains("premium", StringComparison.Ordinal);

    private static bool IsPremiumValueHeader(string normalized) =>
        (normalized.Contains("premium", StringComparison.Ordinal) || normalized.Contains("monthlycost", StringComparison.Ordinal)) &&
        !normalized.Contains("annual", StringComparison.Ordinal) &&
        !normalized.Contains("change", StringComparison.Ordinal) &&
        !normalized.Contains("variance", StringComparison.Ordinal);

    private static RevisedRateScheduleMetadata ReadMetadata(
        IXLWorksheet worksheet,
        int firstRow,
        int lastRow,
        int firstColumn,
        int lastColumn)
    {
        var metadata = new RevisedRateScheduleMetadata();
        var scanLastRow = Math.Min(lastRow, firstRow + 24);
        var scanLastColumn = Math.Min(lastColumn, firstColumn + 14);
        var allText = new List<string>();

        for (var row = firstRow; row <= scanLastRow; row++)
        {
            for (var column = firstColumn; column <= scanLastColumn; column++)
            {
                var text = CellText(worksheet.Cell(row, column));
                if (string.IsNullOrWhiteSpace(text))
                {
                    continue;
                }

                allText.Add(text);
                var normalized = Normalize(text);
                var adjacent = column < scanLastColumn ? CellText(worksheet.Cell(row, column + 1)) : "";
                var inlineValue = ValueAfterLabel(text);
                var value = FirstNonBlank(inlineValue, adjacent);

                if (string.IsNullOrWhiteSpace(metadata.PolicyNumber) &&
                    (normalized.StartsWith("policynumber", StringComparison.Ordinal) || normalized.StartsWith("contractnumber", StringComparison.Ordinal)))
                {
                    metadata.PolicyNumber = value;
                }
                else if (string.IsNullOrWhiteSpace(metadata.ClientName) &&
                         (normalized.StartsWith("clientname", StringComparison.Ordinal) ||
                          normalized.StartsWith("contractholder", StringComparison.Ordinal) ||
                          normalized.StartsWith("policyname", StringComparison.Ordinal)))
                {
                    metadata.ClientName = value;
                }
                else if (!metadata.RenewalDate.HasValue &&
                         (normalized.StartsWith("renewaleffectivedate", StringComparison.Ordinal) ||
                          normalized.StartsWith("rateeffectivedate", StringComparison.Ordinal) ||
                          normalized.StartsWith("effectivedate", StringComparison.Ordinal) ||
                          normalized.StartsWith("renewaldate", StringComparison.Ordinal)) &&
                         TryReadDate(value, out var date))
                {
                    metadata.RenewalDate = date;
                }
            }
        }

        metadata.Carrier = DetectCarrier(string.Join(" ", allText));
        return metadata;
    }

    private static string DetectCarrier(string text)
    {
        var normalized = Normalize(text);
        if (normalized.Contains("canadalife", StringComparison.Ordinal) || normalized.Contains("greatwestlife", StringComparison.Ordinal)) return "Canada Life";
        if (normalized.Contains("manulife", StringComparison.Ordinal)) return "Manulife";
        if (normalized.Contains("sunlife", StringComparison.Ordinal)) return "Sun Life";
        if (normalized.Contains("equitable", StringComparison.Ordinal)) return "Equitable";
        if (normalized.Contains("grouphealth", StringComparison.Ordinal) || normalized.Contains("groupplus", StringComparison.Ordinal)) return "GroupHEALTH";
        if (normalized.Contains("victor", StringComparison.Ordinal)) return "Victor";
        if (normalized.Contains("desjardins", StringComparison.Ordinal)) return "Desjardins";
        if (normalized.Contains("beneva", StringComparison.Ordinal)) return "Beneva";
        if (normalized.Contains("empirelife", StringComparison.Ordinal)) return "Empire Life";
        if (normalized.Contains("greenshield", StringComparison.Ordinal)) return "GreenShield";
        return "";
    }

    private static List<string> ValidateRows(IReadOnlyList<RevisedRateScheduleRow> rows, string sourceName)
    {
        var errors = new List<string>();
        var duplicates = rows
            .GroupBy(row => $"{BenefitKey(row.CanonicalBenefit)}|{TierKey(row.TierOrClass)}|{ClassKey(row.SourceClass)}")
            .Where(group => group.Count() > 1)
            .ToList();
        foreach (var duplicate in duplicates)
        {
            errors.Add($"{sourceName} has duplicate revised mappings: {string.Join(", ", duplicate.Select(row => $"'{row.SourceBenefit}' (row/page {row.SourceRow})"))}.");
        }

        return errors;
    }

    private static int FindColumn(
        IXLWorksheet worksheet,
        int row,
        int firstColumn,
        int lastColumn,
        Func<string, bool> predicate)
    {
        for (var column = firstColumn; column <= lastColumn; column++)
        {
            if (predicate(Normalize(CellText(worksheet.Cell(row, column)))))
            {
                return column;
            }
        }

        return 0;
    }

    private static bool IsBenefitHeader(string value) =>
        value is "benefit" or "benefitline" or "coverage" or "benefitclass" || value.StartsWith("benefit", StringComparison.Ordinal);

    private static bool IsTierHeader(string value) =>
        value is "tier" or "tierclass" or "familycategory" or "coveragecategory" or "ratecategory";

    private static bool IsClassHeader(string value) =>
        value is "class" or "classname" or "classnumber" or "accountclass";

    private static bool TryReadDecimal(IXLCell cell, out decimal value)
    {
        if (cell.TryGetValue(out value))
        {
            return true;
        }

        var text = CellText(cell);
        if (string.IsNullOrWhiteSpace(text) || text is "-" or "N/A")
        {
            value = 0m;
            return false;
        }

        var negative = text.Contains('(') && text.Contains(')');
        var cleaned = text.Replace("$", "", StringComparison.Ordinal)
            .Replace(",", "", StringComparison.Ordinal)
            .Replace("%", "", StringComparison.Ordinal)
            .Replace("(", "", StringComparison.Ordinal)
            .Replace(")", "", StringComparison.Ordinal)
            .Trim();
        if (decimal.TryParse(cleaned, NumberStyles.Number | NumberStyles.AllowExponent, CultureInfo.InvariantCulture, out value) ||
            decimal.TryParse(cleaned, NumberStyles.Number | NumberStyles.AllowExponent, CultureInfo.CurrentCulture, out value))
        {
            if (negative)
            {
                value *= -1m;
            }

            return true;
        }

        value = 0m;
        return false;
    }

    private static bool TryReadDate(string value, out DateOnly date)
    {
        var formats = new[] { "MMMM d, yyyy", "MMMM dd, yyyy", "MMM d, yyyy", "yyyy-MM-dd", "M/d/yyyy", "MM/dd/yyyy" };
        foreach (var format in formats)
        {
            if (DateOnly.TryParseExact(value, format, CanadaCulture, DateTimeStyles.AllowWhiteSpaces, out date))
            {
                return true;
            }
        }

        return DateOnly.TryParse(value, CanadaCulture, DateTimeStyles.AllowWhiteSpaces, out date);
    }

    private static string ValueAfterLabel(string text)
    {
        var separator = text.IndexOf(':');
        return separator >= 0 && separator < text.Length - 1 ? text[(separator + 1)..].Trim() : "";
    }

    private static string CellText(IXLCell cell) => cell.GetFormattedString().Trim();

    private static int SourcePage(string sourceReference)
    {
        var match = Regex.Match(sourceReference ?? "", @"\bpage\s+(?<page>\d+)\b", RegexOptions.IgnoreCase);
        return match.Success && int.TryParse(match.Groups["page"].Value, out var page) ? page : 0;
    }

    private static bool IsTierLabel(string value) => !string.IsNullOrWhiteSpace(TierKey(value));

    private static bool IsTotalRow(string value)
    {
        var normalized = Normalize(value);
        return normalized.Contains("total", StringComparison.Ordinal) ||
               normalized.Contains("difference", StringComparison.Ordinal) ||
               normalized.Contains("variance", StringComparison.Ordinal);
    }

    private static (string Benefit, string TierOrClass) CanonicalIdentity(string sourceBenefit, string tierOrClass)
    {
        var normalized = Normalize(sourceBenefit);
        var tier = DisplayTier(FirstNonBlank(tierOrClass, sourceBenefit));

        if (normalized.Contains("dependentlife", StringComparison.Ordinal) || normalized.StartsWith("deplife", StringComparison.Ordinal)) return ("Dependent Life", tier);
        if (normalized.Contains("optionallife", StringComparison.Ordinal) || normalized.Contains("voluntarylife", StringComparison.Ordinal)) return ("Optional Life", tier);
        if (normalized.Contains("criticalillness", StringComparison.Ordinal)) return ("Critical Illness", tier);
        if (normalized is "add" or "asi" || normalized.Contains("adandd", StringComparison.Ordinal) || normalized.Contains("accidentaldeath", StringComparison.Ordinal) || normalized.Contains("accidentseriousillness", StringComparison.Ordinal)) return ("AD&D", tier);
        if (normalized is "ltd" || normalized.Contains("longtermdisability", StringComparison.Ordinal)) return ("LTD", tier);
        if (normalized is "std" || normalized.Contains("shorttermdisability", StringComparison.Ordinal)) return ("STD", tier);
        if (normalized.Contains("hcsa", StringComparison.Ordinal) || normalized.Contains("healthcarespendingaccount", StringComparison.Ordinal)) return ("HCSA", tier);
        if (normalized.Contains("virtualhealth", StringComparison.Ordinal) || normalized.Contains("onlineaccess", StringComparison.Ordinal)) return ("Virtual Health Care", tier);
        if (normalized is "contact" || normalized.Contains("employeeassistance", StringComparison.Ordinal) || normalized.Contains("familyassistance", StringComparison.Ordinal) || normalized is "eap") return ("EAP", tier);
        if (normalized.Contains("vision", StringComparison.Ordinal)) return ("Vision", tier);
        if (normalized.Contains("dental", StringComparison.Ordinal)) return ("Dental", tier);
        if (normalized.Contains("health", StringComparison.Ordinal) || normalized is "ehc" or "medical") return ("Extended Health Care", tier);
        if (normalized.Contains("life", StringComparison.Ordinal)) return ("Life", tier);
        return (sourceBenefit.Trim(), tier);
    }

    public static string BenefitKey(string value) => Normalize(CanonicalIdentity(value, "").Benefit);

    public static string TierKey(params string[] values)
    {
        var normalized = Normalize(string.Join(" ", values));
        if (normalized.Contains("singleparent", StringComparison.Ordinal)) return "singleparent";
        if (normalized.Contains("parentchild", StringComparison.Ordinal) || normalized.Contains("employeeandchild", StringComparison.Ordinal)) return "parentchild";
        if (normalized.Contains("single", StringComparison.Ordinal) || normalized.Contains("employeeonly", StringComparison.Ordinal)) return "single";
        if (normalized.Contains("couple", StringComparison.Ordinal) || normalized.Contains("employeeandspouse", StringComparison.Ordinal)) return "couple";
        if (normalized.Contains("family", StringComparison.Ordinal)) return "family";
        return "";
    }

    public static string ClassKey(string value)
    {
        var normalized = Normalize(value).Replace("class", "", StringComparison.Ordinal).Replace("account", "", StringComparison.Ordinal);
        return normalized is "all" or "allclasses" ? "" : normalized;
    }

    public static bool ClassMatches(string sourceClass, string targetClass)
    {
        var source = ClassTokens(sourceClass);
        var target = ClassTokens(targetClass);
        if (source.Count == 0 || target.Count == 0)
        {
            return source.Count == 0 && target.Count == 0;
        }

        return source.Overlaps(target);
    }

    public static bool ClientMatches(string sourceClient, string targetClient)
    {
        var source = ClientKey(sourceClient);
        var target = ClientKey(targetClient);
        if (string.IsNullOrWhiteSpace(source) || string.IsNullOrWhiteSpace(target))
        {
            return true;
        }

        return source.Equals(target, StringComparison.Ordinal) ||
               (Math.Min(source.Length, target.Length) >= 5 &&
                (source.Contains(target, StringComparison.Ordinal) || target.Contains(source, StringComparison.Ordinal)));
    }

    public static string PolicyKey(string value) => Normalize(value);

    public static bool CarrierMatches(string sourceCarrier, string targetCarrier)
    {
        var source = Normalize(sourceCarrier);
        var target = Normalize(targetCarrier);
        return string.IsNullOrWhiteSpace(source) || string.IsNullOrWhiteSpace(target) ||
               source.Equals(target, StringComparison.Ordinal) ||
               source.Contains(target, StringComparison.Ordinal) || target.Contains(source, StringComparison.Ordinal);
    }

    private static string ClientKey(string value)
    {
        var words = BenefitMatcher.NormalizeWords(value).Split(' ', StringSplitOptions.RemoveEmptyEntries)
            .Where(word => word is not "ltd" and not "limited" and not "inc" and not "incorporated" and not "company" and not "corp" and not "corporation")
            .ToArray();
        return Normalize(string.Join(" ", words));
    }

    private static HashSet<string> ClassTokens(string value)
    {
        var tokens = Regex.Split((value ?? "").ToUpperInvariant(), @"[^A-Z0-9]+")
            .Where(token => !string.IsNullOrWhiteSpace(token) && token is not "CLASS" and not "ALL")
            .ToHashSet(StringComparer.Ordinal);
        return tokens;
    }

    private static string DisplayTier(string value)
    {
        return TierKey(value) switch
        {
            "single" => "Single",
            "couple" => "Couple",
            "family" => "Family",
            "singleparent" => "Single Parent",
            "parentchild" => "Parent/Child",
            _ => ""
        };
    }

    private static string Normalize(string value) => new((value ?? "").Where(char.IsLetterOrDigit).Select(char.ToLowerInvariant).ToArray());

    private static string FirstNonBlank(params string[] values) => values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)) ?? "";

    private enum ValueColumnKind
    {
        Rate,
        Premium
    }

    private sealed record ValueColumnCandidate(int Column, ValueColumnKind Kind, int Score, string Header);

    private sealed class ValueColumnDetection
    {
        public int RateColumn { get; init; }
        public int PremiumColumn { get; init; }
        public bool HasNetAndGrossColumns { get; init; }
        public IReadOnlyList<string> Errors { get; init; } = [];
        public bool HasValueColumns => RateColumn > 0 || PremiumColumn > 0 || Errors.Count > 0;
    }

    private sealed class RevisedRateScheduleMetadata
    {
        public string Carrier { get; set; } = "";
        public string ClientName { get; set; } = "";
        public string PolicyNumber { get; set; } = "";
        public DateOnly? RenewalDate { get; set; }
    }
}

public sealed class RevisedRateScheduleImportResult
{
    public RevisedRateSchedule? Schedule { get; private init; }
    public IReadOnlyList<string> Errors { get; private init; } = [];
    public bool IsDetected => Schedule is not null || Errors.Count > 0;
    public bool Succeeded => Schedule is not null && Errors.Count == 0;

    public static RevisedRateScheduleImportResult Success(RevisedRateSchedule schedule) => new() { Schedule = schedule };
    public static RevisedRateScheduleImportResult Failure(IReadOnlyList<string> errors) => new() { Errors = errors };
    public static RevisedRateScheduleImportResult NotDetected() => new();
}

public sealed class RevisedRateSchedule
{
    public string SourceFilePath { get; init; } = "";
    public string SourceFileName { get; init; } = "";
    public string SourceKind { get; init; } = "";
    public string SheetName { get; init; } = "";
    public string Carrier { get; init; } = "";
    public string ClientName { get; init; } = "";
    public string PolicyNumber { get; init; } = "";
    public DateOnly? RenewalDate { get; init; }
    public IReadOnlyList<RevisedRateScheduleRow> Rows { get; init; } = [];
    public IReadOnlyList<string> Warnings { get; init; } = [];
}

public sealed class RevisedRateScheduleRow
{
    public string SourceBenefit { get; init; } = "";
    public string CanonicalBenefit { get; init; } = "";
    public string TierOrClass { get; init; } = "";
    public string SourceClass { get; init; } = "";
    public decimal? Volume { get; init; }
    public int? Lives { get; init; }
    public decimal? CurrentRate { get; init; }
    public decimal? CurrentPremium { get; init; }
    public decimal? RevisedRate { get; init; }
    public decimal? RevisedPremium { get; init; }
    public int SourceRow { get; init; }
    public string SourceReference { get; init; } = "";
}

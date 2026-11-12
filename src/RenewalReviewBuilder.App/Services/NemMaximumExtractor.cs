using System.Globalization;
using System.Text.RegularExpressions;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed class NemMaximumExtractor
{
    private const string LifeBenefit = "Life Insurance";
    private const string LtdBenefit = "Long-Term Disability";
    private const string MoneyOrNaPattern = @"(?:-|–|—|N/?A|not\s+available|\$?\s*\d{1,3}(?:,\d{3})*(?:\.\d{2})?|\$?\s*\d{4,}(?:\.\d{2})?)";
    private const string BenefitPattern = @"(?:Employee\s+Term\s+Life\s+Insurance|Employee\s+Life\s+Insurance|Basic\s+Life(?:\s+Insurance)?|Life(?:\s+Insurance)?|Long[-\s]+Term\s+Disability|Long[-\s]+Term\s+Disability\s+Benefit|LTD)";

    private static readonly Regex NemTableRowRegex = new(
        $@"^(?<benefit>{BenefitPattern})\s+(?<current>{MoneyOrNaPattern})\s+(?<available>{MoneyOrNaPattern})(?:\s|$)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex InlineNemRegex = new(
        $@"(?<benefit>{BenefitPattern}).{{0,90}}(?:Non[-\s]?Evidence\s+(?:Maximum|Limit|Amount)|NEM).{{0,90}}(?<current>{MoneyOrNaPattern})(?:.{{0,50}}(?<available>{MoneyOrNaPattern}))?",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex SplitMoneyTokenRegex = new(
        @"(?<![\d,])(?<prefix>\d{1,2})\s+(?<rest>\d{1,2},\d{3})(?![\d,])",
        RegexOptions.Compiled);

    private static readonly Regex SpacedCommaMoneyTokenRegex = new(
        @"(?<=\d)\s+,\s*(?=\d{3}(?:\D|$))",
        RegexOptions.Compiled);

    private static readonly Regex CurrentMaximumRegex = new(
        $@"(?<benefit>{BenefitPattern}).{{0,140}}(?:Overall\s+Maximum|Plan\s+Maximum|Maximum\s+Monthly\s+Benefit|Monthly\s+Benefit\s+Maximum|Maximum\s+Benefit|Maximum\s+Amount|Benefit\s+Maximum|maximum\s+of)\s*(?<maximum>{MoneyOrNaPattern})",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex ContextMaximumRegex = new(
        $@"(?:Overall\s+Maximum|Plan\s+Maximum|Maximum\s+Monthly\s+Benefit|Monthly\s+Benefit\s+Maximum|Maximum\s+Benefit|Maximum\s+Amount|Benefit\s+Maximum|maximum\s+of)\s*(?<maximum>{MoneyOrNaPattern})",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex SunLifeNemAndBenefitMaximumRegex = new(
        $@"NEM\s+and\s+benefit\s+maximum\s+are\s+(?<amount>{MoneyOrNaPattern})",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex SunLifeCurrentNemAndMaximumRegex = new(
        $@"current\s+non[-\s]?evidence\s+maximum\s+is\s+(?<nem>{MoneyOrNaPattern})(?:/month)?\s+and\s+overall\s+maximum\s+is\s+(?<maximum>{MoneyOrNaPattern})(?:/month)?",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex SunLifeAvailableNemRegex = new(
        $@"increase\s+your\s+plan(?:'|â€™|’)?s\s+NEM\s+to\s+(?<available>{MoneyOrNaPattern})(?:/month)?",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private readonly PdfTextExtractor _pdfTextExtractor;

    public NemMaximumExtractor(PdfTextExtractor? pdfTextExtractor = null)
    {
        _pdfTextExtractor = pdfTextExtractor ?? new PdfTextExtractor();
    }

    public List<NemMaximumRow> ExtractFromFiles(IEnumerable<string> pdfPaths)
    {
        var rows = new List<NemMaximumRow>();
        foreach (var path in pdfPaths.Where(File.Exists).Distinct(StringComparer.OrdinalIgnoreCase))
        {
            try
            {
                AddOrMerge(rows, ExtractFromPages(path, _pdfTextExtractor.Extract(path)));
            }
            catch (Exception ex)
            {
                AppLogger.Error(ex);
            }
        }

        return OrderRows(rows);
    }

    public static List<NemMaximumRow> ExtractFromPages(string sourcePath, IReadOnlyList<PdfPageText> pages)
    {
        var rows = new List<NemMaximumRow>();
        foreach (var page in pages)
        {
            ExtractNemRowsFromPage(sourcePath, page, rows);
            ExtractCurrentMaximumRowsFromPage(sourcePath, page, rows);
        }

        return OrderRows(rows);
    }

    public static string NormalizeBenefitName(string value)
    {
        var key = NormalizeKey(value);
        if (key.Contains("ltd", StringComparison.Ordinal) ||
            key.Contains("longtermdisability", StringComparison.Ordinal))
        {
            return LtdBenefit;
        }

        if (key.Contains("life", StringComparison.Ordinal) &&
            !key.Contains("dependent", StringComparison.Ordinal))
        {
            return LifeBenefit;
        }

        return ParsingHelpers.Clean(value);
    }

    private static void ExtractNemRowsFromPage(string sourcePath, PdfPageText page, List<NemMaximumRow> rows)
    {
        if (!ContainsNemClue(page.Text))
        {
            return;
        }

        var currentBenefit = "";
        foreach (var line in page.Lines.Select(NormalizeNemLine).Where(line => line.Length > 0))
        {
            if (TryMapBenefitHeading(line, out var headingBenefit))
            {
                currentBenefit = headingBenefit;
            }

            ExtractSunLifeNarrativeNemRows(sourcePath, page, line, currentBenefit, rows);

            var tableMatch = NemTableRowRegex.Match(line);
            if (tableMatch.Success)
            {
                AddOrMerge(rows, new NemMaximumRow
                {
                    Benefit = NormalizeBenefitName(tableMatch.Groups["benefit"].Value),
                    CurrentNem = NormalizeMoneyText(tableMatch.Groups["current"].Value),
                    AvailableNem = NormalizeMoneyText(tableMatch.Groups["available"].Value),
                    SourceFileName = Path.GetFileName(sourcePath),
                    SourcePage = page.PageNumber,
                    SourceReference = $"PDF page {page.PageNumber}: {line}",
                    ConfidenceScore = 0.85m
                });
                continue;
            }

            foreach (Match match in InlineNemRegex.Matches(line))
            {
                AddOrMerge(rows, new NemMaximumRow
                {
                    Benefit = NormalizeBenefitName(match.Groups["benefit"].Value),
                    CurrentNem = NormalizeMoneyText(match.Groups["current"].Value),
                    AvailableNem = NormalizeMoneyText(match.Groups["available"].Value),
                    SourceFileName = Path.GetFileName(sourcePath),
                    SourcePage = page.PageNumber,
                    SourceReference = $"PDF page {page.PageNumber}: {line}",
                    ConfidenceScore = 0.75m
                });
            }
        }

        foreach (Match match in InlineNemRegex.Matches(NormalizeNemLine(page.Text)))
        {
            AddOrMerge(rows, new NemMaximumRow
            {
                Benefit = NormalizeBenefitName(match.Groups["benefit"].Value),
                CurrentNem = NormalizeMoneyText(match.Groups["current"].Value),
                AvailableNem = NormalizeMoneyText(match.Groups["available"].Value),
                SourceFileName = Path.GetFileName(sourcePath),
                SourcePage = page.PageNumber,
                SourceReference = $"PDF page {page.PageNumber}: Non-Evidence Maximums",
                ConfidenceScore = 0.65m
            });
        }
    }

    private static void ExtractSunLifeNarrativeNemRows(
        string sourcePath,
        PdfPageText page,
        string line,
        string currentBenefit,
        List<NemMaximumRow> rows)
    {
        if (string.IsNullOrWhiteSpace(currentBenefit))
        {
            return;
        }

        var sameMaximumMatch = SunLifeNemAndBenefitMaximumRegex.Match(line);
        if (sameMaximumMatch.Success)
        {
            var amount = NormalizeMoneyText(sameMaximumMatch.Groups["amount"].Value);
            AddOrMerge(rows, new NemMaximumRow
            {
                Benefit = currentBenefit,
                CurrentMaximum = amount,
                CurrentNem = amount,
                AvailableNem = amount,
                SourceFileName = Path.GetFileName(sourcePath),
                SourcePage = page.PageNumber,
                SourceReference = $"PDF page {page.PageNumber}: {line}",
                ConfidenceScore = 0.88m
            });
            return;
        }

        var currentMatch = SunLifeCurrentNemAndMaximumRegex.Match(line);
        if (currentMatch.Success)
        {
            AddOrMerge(rows, new NemMaximumRow
            {
                Benefit = currentBenefit,
                CurrentMaximum = NormalizeMoneyText(currentMatch.Groups["maximum"].Value),
                CurrentNem = NormalizeMoneyText(currentMatch.Groups["nem"].Value),
                SourceFileName = Path.GetFileName(sourcePath),
                SourcePage = page.PageNumber,
                SourceReference = $"PDF page {page.PageNumber}: {line}",
                ConfidenceScore = 0.88m
            });
            return;
        }

        var availableMatch = SunLifeAvailableNemRegex.Match(line);
        if (availableMatch.Success)
        {
            AddOrMerge(rows, new NemMaximumRow
            {
                Benefit = currentBenefit,
                AvailableNem = NormalizeMoneyText(availableMatch.Groups["available"].Value),
                SourceFileName = Path.GetFileName(sourcePath),
                SourcePage = page.PageNumber,
                SourceReference = $"PDF page {page.PageNumber}: {line}",
                ConfidenceScore = 0.9m
            });
        }
    }

    private static void ExtractCurrentMaximumRowsFromPage(string sourcePath, PdfPageText page, List<NemMaximumRow> rows)
    {
        var currentBenefit = "";
        foreach (var line in page.Lines.Select(NormalizeNemLine).Where(line => line.Length > 0))
        {
            if (TryMapBenefitHeading(line, out var headingBenefit))
            {
                currentBenefit = headingBenefit;
            }

            if (!IsCurrentMaximumCandidate(line))
            {
                continue;
            }

            var lineMatch = CurrentMaximumRegex.Match(line);
            var benefit = lineMatch.Success
                ? NormalizeBenefitName(lineMatch.Groups["benefit"].Value)
                : currentBenefit;
            var maximum = lineMatch.Success
                ? lineMatch.Groups["maximum"].Value
                : ContextMaximumRegex.Match(line).Groups["maximum"].Value;

            if (string.IsNullOrWhiteSpace(benefit) || string.IsNullOrWhiteSpace(maximum))
            {
                continue;
            }

            AddOrMerge(rows, new NemMaximumRow
            {
                Benefit = benefit,
                CurrentMaximum = NormalizeMoneyText(maximum),
                SourceFileName = Path.GetFileName(sourcePath),
                SourcePage = page.PageNumber,
                SourceReference = $"PDF page {page.PageNumber}: {line}",
                ConfidenceScore = lineMatch.Success ? 0.7m : 0.55m
            });
        }
    }

    private static bool ContainsNemClue(string text)
    {
        return text.Contains("Non-Evidence", StringComparison.OrdinalIgnoreCase) ||
               text.Contains("Non Evidence", StringComparison.OrdinalIgnoreCase) ||
               text.Contains("Non-evidence", StringComparison.OrdinalIgnoreCase) ||
               Regex.IsMatch(text, @"\bNEMs?\b", RegexOptions.IgnoreCase);
    }

    private static bool IsCurrentMaximumCandidate(string line)
    {
        if (!line.Contains("maximum", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        if (ContainsNemClue(line) ||
            line.Contains("evidence of insurability", StringComparison.OrdinalIgnoreCase) ||
            line.Contains("age reduction", StringComparison.OrdinalIgnoreCase) ||
            line.Contains("reduction schedule", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        return ContextMaximumRegex.IsMatch(line);
    }

    private static bool TryMapBenefitHeading(string line, out string benefit)
    {
        benefit = "";
        if (line.Contains("Dependent", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        if (line.Contains("Long Term Disability", StringComparison.OrdinalIgnoreCase) ||
            line.Contains("Long-Term Disability", StringComparison.OrdinalIgnoreCase) ||
            Regex.IsMatch(line, @"\bLTD\b", RegexOptions.IgnoreCase))
        {
            benefit = LtdBenefit;
            return true;
        }

        if (line.Contains("Life Insurance", StringComparison.OrdinalIgnoreCase) ||
            line.Contains("Employee Life", StringComparison.OrdinalIgnoreCase) ||
            line.Contains("Employee Term Life", StringComparison.OrdinalIgnoreCase) ||
            line.Equals("Life", StringComparison.OrdinalIgnoreCase))
        {
            benefit = LifeBenefit;
            return true;
        }

        return false;
    }

    private static void AddOrMerge(List<NemMaximumRow> rows, IEnumerable<NemMaximumRow> newRows)
    {
        foreach (var row in newRows)
        {
            AddOrMerge(rows, row);
        }
    }

    private static void AddOrMerge(List<NemMaximumRow> rows, NemMaximumRow row)
    {
        if (string.IsNullOrWhiteSpace(row.Benefit))
        {
            return;
        }

        var benefit = NormalizeBenefitName(row.Benefit);
        var existing = rows.FirstOrDefault(current => NormalizeBenefitName(current.Benefit).Equals(benefit, StringComparison.OrdinalIgnoreCase));
        if (existing is null)
        {
            row.Benefit = benefit;
            rows.Add(row);
            return;
        }

        if (ShouldReplace(existing.CurrentMaximum, row.CurrentMaximum, existing.ConfidenceScore, row.ConfidenceScore))
        {
            existing.CurrentMaximum = row.CurrentMaximum;
            UpdateSource(existing, row);
        }

        if (ShouldReplace(existing.CurrentNem, row.CurrentNem, existing.ConfidenceScore, row.ConfidenceScore))
        {
            existing.CurrentNem = row.CurrentNem;
            UpdateSource(existing, row);
        }

        if (ShouldReplace(existing.AvailableNem, row.AvailableNem, existing.ConfidenceScore, row.ConfidenceScore))
        {
            existing.AvailableNem = row.AvailableNem;
            UpdateSource(existing, row);
        }

        existing.ConfidenceScore = Math.Max(existing.ConfidenceScore, row.ConfidenceScore);
    }

    private static bool ShouldReplace(string current, string proposed, decimal currentConfidence, decimal proposedConfidence)
    {
        if (string.IsNullOrWhiteSpace(proposed))
        {
            return false;
        }

        return string.IsNullOrWhiteSpace(current) || proposedConfidence >= currentConfidence;
    }

    private static void UpdateSource(NemMaximumRow target, NemMaximumRow source)
    {
        target.SourceFileName = FirstNonBlank(source.SourceFileName, target.SourceFileName);
        target.SourcePage = source.SourcePage ?? target.SourcePage;
        target.SourceReference = FirstNonBlank(source.SourceReference, target.SourceReference);
    }

    private static List<NemMaximumRow> OrderRows(IEnumerable<NemMaximumRow> rows)
    {
        return rows
            .OrderBy(row => NemSort(row.Benefit))
            .ThenBy(row => row.Benefit, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static int NemSort(string benefit)
    {
        return NormalizeBenefitName(benefit) switch
        {
            LifeBenefit => 0,
            LtdBenefit => 1,
            _ => 99
        };
    }

    private static string NormalizeMoneyText(string value)
    {
        value = ParsingHelpers.Clean(value);
        if (string.IsNullOrWhiteSpace(value))
        {
            return "";
        }

        if (value.Equals("N/A", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("NA", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("-", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("–", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("—", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("not available", StringComparison.OrdinalIgnoreCase))
        {
            return "N/A";
        }

        var numeric = value.Replace("$", "", StringComparison.Ordinal)
            .Replace(",", "", StringComparison.Ordinal)
            .Replace(" ", "", StringComparison.Ordinal);
        return decimal.TryParse(numeric, NumberStyles.Number, CultureInfo.InvariantCulture, out var amount)
            ? amount.ToString("#,##0", CultureInfo.CurrentCulture)
            : value;
    }

    private static string NormalizeNemLine(string value)
    {
        value = ParsingHelpers.Clean(value);
        value = SpacedCommaMoneyTokenRegex.Replace(value, ",");
        var previous = "";
        while (!previous.Equals(value, StringComparison.Ordinal))
        {
            previous = value;
            value = SplitMoneyTokenRegex.Replace(value, match =>
                match.Groups["prefix"].Value + match.Groups["rest"].Value);
        }

        return value;
    }

    private static string NormalizeKey(string value)
    {
        return new string(ParsingHelpers.Clean(value)
            .Where(char.IsLetterOrDigit)
            .Select(char.ToLowerInvariant)
            .ToArray());
    }

    private static string FirstNonBlank(params string[] values)
    {
        return values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)) ?? "";
    }
}

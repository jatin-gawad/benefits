using System.Globalization;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using RenewalReviewBuilder.App.Models;
using UglyToad.PdfPig;

namespace RenewalReviewBuilder.App.Services;

/// <summary>
/// Reads GroupHEALTH Claims Range chart exports using PDF word coordinates. The chart is not a PDF
/// table, so category and data-label positions are retained as extraction evidence.
/// </summary>
public sealed class GroupHealthClaimsRangePdfParser
{
    private const double LineTolerance = 3.0;
    private static readonly Regex WholeNumberRegex = new(
        @"^\d{1,3}(?:,\d{3})*$|^\d+$|^\d+(?:\.0+)$",
        RegexOptions.Compiled);
    private static readonly Regex AxisNumberRegex = new(
        @"^\d+(?:\.\d+)?$",
        RegexOptions.Compiled);
    private static readonly Regex GroupRegex = new(
        @"^Group\s*:\s*(?<value>.+)$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex PolicyRegex = new(
        @"^Policy\s*:\s*(?<value>\d+)$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex StartDateRegex = new(
        @"^Start\s+Date\s*:\s*(?<value>.+)$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex EndDateRegex = new(
        @"^End\s+Date\s*:\s*(?<value>.+)$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public ClaimsRangeReport Parse(string pdfPath)
    {
        if (string.IsNullOrWhiteSpace(pdfPath))
        {
            throw new ArgumentException("A Claims Range PDF path is required.", nameof(pdfPath));
        }

        if (!File.Exists(pdfPath))
        {
            throw new FileNotFoundException("Claims Range PDF was not found.", pdfPath);
        }

        var fullPath = Path.GetFullPath(pdfPath);
        var analyzedSourceSnapshot = CaptureSourceSnapshot(fullPath);
        var report = new ClaimsRangeReport
        {
            SourcePath = fullPath,
            SourceFileName = Path.GetFileName(pdfPath),
            SourceSnapshotIsAuthoritative = true,
            SourceFileLength = analyzedSourceSnapshot.Length,
            SourceLastWriteTimeUtc = analyzedSourceSnapshot.LastWriteTimeUtc,
            SourceContentSha256 = analyzedSourceSnapshot.ContentSha256
        };

        using (var document = PdfDocument.Open(fullPath))
        {
            foreach (var pdfPage in document.GetPages())
            {
            var words = pdfPage.GetWords()
                .Select(word => new WordToken(
                    Clean(word.Text),
                    word.BoundingBox.Left,
                    word.BoundingBox.Bottom,
                    word.BoundingBox.Right,
                    word.BoundingBox.Top))
                .Where(word => !string.IsNullOrWhiteSpace(word.Text))
                .ToList();
            var lines = BuildLines(words);
            var benefit = DetectBenefit(lines);
            if (benefit is null)
            {
                continue;
            }

            var sourceReference = $"{report.SourceFileName} page {pdfPage.Number}";
            var page = new ClaimsRangePage
            {
                Benefit = benefit.Value.Key,
                BenefitDisplayName = benefit.Value.DisplayName,
                SourcePage = pdfPage.Number,
                SourceReference = sourceReference
            };

            var pageClient = MatchLineValue(lines, GroupRegex);
            var pagePolicy = MatchLineValue(lines, PolicyRegex);
            var pageStart = ParseDate(MatchLineValue(lines, StartDateRegex));
            var pageEnd = ParseDate(MatchLineValue(lines, EndDateRegex));
            MergeHeader(report, pageClient, pagePolicy, pageStart, pageEnd, sourceReference);

            var categoryLine = FindCategoryLine(lines);
            if (categoryLine is null)
            {
                report.Errors.Add($"The nine Claims Range categories could not be located on {sourceReference}.");
                report.Pages.Add(page);
                continue;
            }

            var matchedBands = MatchAllBands(categoryLine.Words);
            if (matchedBands.Count != ClaimsRangeBandCatalog.All.Count)
            {
                report.Errors.Add($"Expected nine Claims Range categories but found {matchedBands.Count} on {sourceReference}.");
                report.Pages.Add(page);
                continue;
            }

            var plotTop = FindPlotTop(words, categoryLine, matchedBands, pdfPage.Height);
            var plotBottom = categoryLine.Top + 2.0;
            for (var index = 0; index < matchedBands.Count; index++)
            {
                var band = matchedBands[index];
                var leftBoundary = index == 0
                    ? band.Bounds.CenterX - ((matchedBands[index + 1].Bounds.CenterX - band.Bounds.CenterX) / 2.0)
                    : (matchedBands[index - 1].Bounds.CenterX + band.Bounds.CenterX) / 2.0;
                var rightBoundary = index == matchedBands.Count - 1
                    ? band.Bounds.CenterX + ((band.Bounds.CenterX - matchedBands[index - 1].Bounds.CenterX) / 2.0)
                    : (band.Bounds.CenterX + matchedBands[index + 1].Bounds.CenterX) / 2.0;

                var countCandidates = words
                    .Where(word => word.CenterX >= leftBoundary && word.CenterX < rightBoundary)
                    .Where(word => word.Bottom >= plotBottom && word.Top <= plotTop)
                    .Where(word => TryParseCount(word.Text, out _))
                    .OrderBy(word => Math.Abs(word.CenterX - band.Bounds.CenterX))
                    .ThenByDescending(word => word.Bottom)
                    .ToList();

                if (countCandidates.Count == 0)
                {
                    report.Errors.Add($"The count for {band.Definition.DisplayLabel} was not found on {sourceReference}.");
                    continue;
                }

                var countWord = countCandidates[0];
                _ = TryParseCount(countWord.Text, out var count);
                page.Rows.Add(new ClaimsRangeRow
                {
                    Benefit = page.Benefit,
                    RangeCode = band.Definition.RangeCode,
                    RangeLabel = band.Definition.DisplayLabel,
                    RawRangeLabel = band.RawLabel,
                    SortOrder = band.Definition.SortOrder,
                    LowerBoundInclusive = band.Definition.LowerBoundInclusive,
                    UpperBoundInclusive = band.Definition.UpperBoundInclusive,
                    CertificateCount = count,
                    SourcePage = pdfPage.Number,
                    SourceReference = sourceReference,
                    LabelBounds = ToModelBounds(band.Bounds),
                    CountBounds = ToModelBounds(countWord.Bounds)
                });
            }

            if (page.Rows.Count != ClaimsRangeBandCatalog.All.Count)
            {
                report.Errors.Add($"Expected nine Claims Range counts but extracted {page.Rows.Count} on {sourceReference}.");
            }

                report.Pages.Add(page);
            }
        }

        ValidateReport(report);
        var finalSourceSnapshot = CaptureSourceSnapshot(fullPath);
        if (analyzedSourceSnapshot != finalSourceSnapshot)
        {
            report.Errors.Add("The Claims Range source PDF changed while it was being analyzed. Re-run the analysis before review or export.");
        }
        return report;
    }

    private static SourceSnapshot CaptureSourceSnapshot(string pdfPath)
    {
        var before = new FileInfo(pdfPath);
        var length = before.Length;
        var lastWriteTimeUtc = before.LastWriteTimeUtc;
        using var stream = File.OpenRead(pdfPath);
        var contentSha256 = Convert.ToHexString(SHA256.HashData(stream));
        var after = new FileInfo(pdfPath);
        if (after.Length != length || after.LastWriteTimeUtc != lastWriteTimeUtc)
        {
            throw new IOException("The Claims Range source PDF changed while its fingerprint was being calculated.");
        }

        return new SourceSnapshot(length, lastWriteTimeUtc, contentSha256);
    }

    private static void ValidateReport(ClaimsRangeReport report)
    {
        if (report.Pages.Count == 0)
        {
            report.Errors.Add("No Extended Health Care or Dental Claims Range pages were found.");
            return;
        }

        if (report.Pages.Count != 2)
        {
            report.Errors.Add($"Expected two Claims Range pages but found {report.Pages.Count}.");
        }

        foreach (var benefit in new[] { ClaimsRangeBenefitKeys.Health, ClaimsRangeBenefitKeys.Dental })
        {
            var matchingPages = report.Pages.Count(page =>
                page.Benefit.Equals(benefit, StringComparison.OrdinalIgnoreCase));
            if (matchingPages != 1)
            {
                report.Errors.Add($"Expected one {benefit} Claims Range page but found {matchingPages}.");
            }
        }

        if (string.IsNullOrWhiteSpace(report.ClientName))
        {
            report.Errors.Add("Claims Range client name was not found.");
        }

        if (string.IsNullOrWhiteSpace(report.PolicyNumber))
        {
            report.Errors.Add("Claims Range policy number was not found.");
        }

        if (!report.PeriodStart.HasValue || !report.PeriodEnd.HasValue)
        {
            report.Errors.Add("Claims Range start or end date was not found.");
        }
        else if (report.PeriodStart > report.PeriodEnd)
        {
            report.Errors.Add("Claims Range start date is after its end date.");
        }

        if (report.RowCount != ClaimsRangeBandCatalog.All.Count * 2)
        {
            report.Errors.Add($"A complete Claims Range report requires 18 rows; {report.RowCount} were extracted.");
        }
    }

    private static void MergeHeader(
        ClaimsRangeReport report,
        string pageClient,
        string pagePolicy,
        DateOnly? pageStart,
        DateOnly? pageEnd,
        string sourceReference)
    {
        MergeTextHeader(
            report.ClientName,
            pageClient,
            value => report.ClientName = value,
            "client name",
            sourceReference,
            report.Errors);
        MergeTextHeader(
            report.PolicyNumber,
            pagePolicy,
            value => report.PolicyNumber = value,
            "policy number",
            sourceReference,
            report.Errors);
        MergeDateHeader(
            report.PeriodStart,
            pageStart,
            value => report.PeriodStart = value,
            "start date",
            sourceReference,
            report.Errors);
        MergeDateHeader(
            report.PeriodEnd,
            pageEnd,
            value => report.PeriodEnd = value,
            "end date",
            sourceReference,
            report.Errors);
    }

    private static void MergeTextHeader(
        string existing,
        string incoming,
        Action<string> assign,
        string fieldName,
        string sourceReference,
        ICollection<string> errors)
    {
        if (string.IsNullOrWhiteSpace(incoming))
        {
            errors.Add($"Claims Range {fieldName} was not found on {sourceReference}.");
            return;
        }

        if (string.IsNullOrWhiteSpace(existing))
        {
            assign(incoming.Trim());
            return;
        }

        if (!NormalizeIdentity(existing).Equals(NormalizeIdentity(incoming), StringComparison.OrdinalIgnoreCase))
        {
            errors.Add($"Claims Range {fieldName} differs between pages: '{existing}' and '{incoming}' on {sourceReference}.");
        }
    }

    private static void MergeDateHeader(
        DateOnly? existing,
        DateOnly? incoming,
        Action<DateOnly?> assign,
        string fieldName,
        string sourceReference,
        ICollection<string> errors)
    {
        if (!incoming.HasValue)
        {
            errors.Add($"Claims Range {fieldName} was not found on {sourceReference}.");
            return;
        }

        if (!existing.HasValue)
        {
            assign(incoming);
            return;
        }

        if (existing != incoming)
        {
            errors.Add($"Claims Range {fieldName} differs between pages on {sourceReference}.");
        }
    }

    private static double FindPlotTop(
        IReadOnlyList<WordToken> words,
        WordLine categoryLine,
        IReadOnlyList<MatchedBand> bands,
        double pageHeight)
    {
        var firstBandLeft = bands[0].Bounds.Left;
        var axisWords = words
            .Where(word => word.Right < firstBandLeft - 5.0)
            .Where(word => word.Bottom > categoryLine.Top + 2.0)
            .Where(word => AxisNumberRegex.IsMatch(word.Text.Replace(",", "", StringComparison.Ordinal)))
            .ToList();
        var clusters = new List<List<WordToken>>();
        foreach (var word in axisWords.OrderBy(word => word.Right))
        {
            var cluster = clusters.FirstOrDefault(existing => Math.Abs(existing[0].Right - word.Right) <= 3.0);
            if (cluster is null)
            {
                clusters.Add([word]);
            }
            else
            {
                cluster.Add(word);
            }
        }

        var valueAxis = clusters
            .Where(cluster => cluster.Count >= 3)
            .OrderByDescending(cluster => cluster.Count)
            .FirstOrDefault();
        if (valueAxis is not null)
        {
            return Math.Min(pageHeight, valueAxis.Max(word => word.Top) + 20.0);
        }

        // Fallback stays below the header area for the standard landscape Claims Range export.
        return Math.Min(pageHeight * 0.72, categoryLine.Top + (pageHeight * 0.36));
    }

    private static WordLine? FindCategoryLine(IEnumerable<WordLine> lines)
    {
        return lines
            .Select(line => new { Line = line, Matches = MatchAllBands(line.Words) })
            .Where(candidate => candidate.Matches.Count == ClaimsRangeBandCatalog.All.Count)
            .OrderBy(candidate => candidate.Line.Bottom)
            .Select(candidate => candidate.Line)
            .FirstOrDefault();
    }

    private static List<MatchedBand> MatchAllBands(IReadOnlyList<WordToken> words)
    {
        var results = new List<MatchedBand>();
        var cursor = 0;
        foreach (var definition in ClaimsRangeBandCatalog.All)
        {
            var match = FindBand(words, definition, cursor);
            if (match is null)
            {
                return results;
            }

            results.Add(match);
            cursor = match.EndWordIndex + 1;
        }

        return results;
    }

    private static MatchedBand? FindBand(
        IReadOnlyList<WordToken> words,
        ClaimsRangeBandDefinition definition,
        int startIndex)
    {
        for (var candidateStart = startIndex; candidateStart < Math.Min(words.Count, startIndex + 4); candidateStart++)
        {
            for (var length = 1; length <= 4 && candidateStart + length <= words.Count; length++)
            {
                var candidateWords = words.Skip(candidateStart).Take(length).ToList();
                var key = NormalizeRangeKey(string.Concat(candidateWords.Select(word => word.Text)));
                if (!key.Equals(definition.PdfMatchKey, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                return new MatchedBand(
                    definition,
                    string.Join(" ", candidateWords.Select(word => word.Text)),
                    WordBounds.Union(candidateWords),
                    candidateStart + length - 1);
            }
        }

        return null;
    }

    private static (string Key, string DisplayName)? DetectBenefit(IEnumerable<WordLine> lines)
    {
        foreach (var line in lines)
        {
            if (line.Text.Contains("Extended Health Care Claims Range by Certificate", StringComparison.OrdinalIgnoreCase))
            {
                return (ClaimsRangeBenefitKeys.Health, "Extended Health Care");
            }

            if (line.Text.Contains("Dental Care Claims Range by Certificate", StringComparison.OrdinalIgnoreCase))
            {
                return (ClaimsRangeBenefitKeys.Dental, "Dental Care");
            }
        }

        return null;
    }

    private static string MatchLineValue(IEnumerable<WordLine> lines, Regex regex)
    {
        foreach (var line in lines)
        {
            var match = regex.Match(line.Text);
            if (match.Success)
            {
                return match.Groups["value"].Value.Trim();
            }
        }

        return "";
    }

    private static DateOnly? ParseDate(string value)
    {
        if (DateTime.TryParse(
                value,
                CultureInfo.InvariantCulture,
                DateTimeStyles.AllowWhiteSpaces,
                out var parsed))
        {
            return DateOnly.FromDateTime(parsed);
        }

        return null;
    }

    private static bool TryParseCount(string value, out int count)
    {
        count = 0;
        var normalized = value.Trim();
        if (!WholeNumberRegex.IsMatch(normalized))
        {
            return false;
        }

        normalized = normalized.Replace(",", "", StringComparison.Ordinal);
        if (normalized.Contains('.'))
        {
            if (!decimal.TryParse(normalized, NumberStyles.Number, CultureInfo.InvariantCulture, out var decimalValue) ||
                decimal.Truncate(decimalValue) != decimalValue ||
                decimalValue > int.MaxValue)
            {
                return false;
            }

            count = decimal.ToInt32(decimalValue);
            return true;
        }

        return int.TryParse(normalized, NumberStyles.None, CultureInfo.InvariantCulture, out count) && count >= 0;
    }

    private static List<WordLine> BuildLines(IEnumerable<WordToken> words)
    {
        var rows = new List<List<WordToken>>();
        foreach (var word in words.OrderByDescending(word => word.Bottom).ThenBy(word => word.Left))
        {
            var row = rows.FirstOrDefault(existing => Math.Abs(existing[0].Bottom - word.Bottom) < LineTolerance);
            if (row is null)
            {
                rows.Add([word]);
            }
            else
            {
                row.Add(word);
            }
        }

        return rows
            .Select(row => new WordLine(row.OrderBy(word => word.Left).ToList()))
            .ToList();
    }

    private static string NormalizeRangeKey(string value)
    {
        return value
            .Replace("$", "", StringComparison.Ordinal)
            .Replace(",", "", StringComparison.Ordinal)
            .Replace(" ", "", StringComparison.Ordinal)
            .Replace("–", "-", StringComparison.Ordinal)
            .Replace("—", "-", StringComparison.Ordinal)
            .Replace("−", "-", StringComparison.Ordinal);
    }

    private static string NormalizeIdentity(string value)
    {
        return new string(value.Where(char.IsLetterOrDigit).Select(char.ToUpperInvariant).ToArray());
    }

    private static string Clean(string value) => Regex.Replace(value ?? "", @"\s+", " ").Trim();

    private static ClaimsRangeBoundingBox ToModelBounds(WordBounds bounds) => new()
    {
        Left = Convert.ToDecimal(bounds.Left),
        Bottom = Convert.ToDecimal(bounds.Bottom),
        Right = Convert.ToDecimal(bounds.Right),
        Top = Convert.ToDecimal(bounds.Top)
    };

    private sealed record WordToken(string Text, double Left, double Bottom, double Right, double Top)
    {
        public double CenterX => Left + ((Right - Left) / 2.0);
        public WordBounds Bounds => new(Left, Bottom, Right, Top);
    }

    private sealed class WordLine
    {
        public WordLine(IReadOnlyList<WordToken> words)
        {
            Words = words;
            Text = string.Join(" ", words.Select(word => word.Text));
            var bounds = WordBounds.Union(words);
            Bottom = bounds.Bottom;
            Top = bounds.Top;
        }

        public IReadOnlyList<WordToken> Words { get; }
        public string Text { get; }
        public double Bottom { get; }
        public double Top { get; }
    }

    private sealed record MatchedBand(
        ClaimsRangeBandDefinition Definition,
        string RawLabel,
        WordBounds Bounds,
        int EndWordIndex);

    private sealed record WordBounds(double Left, double Bottom, double Right, double Top)
    {
        public double CenterX => Left + ((Right - Left) / 2.0);

        public static WordBounds Union(IEnumerable<WordToken> words)
        {
            var list = words.ToList();
            return new WordBounds(
                list.Min(word => word.Left),
                list.Min(word => word.Bottom),
                list.Max(word => word.Right),
                list.Max(word => word.Top));
        }
    }

    private sealed record SourceSnapshot(long Length, DateTime LastWriteTimeUtc, string ContentSha256);
}

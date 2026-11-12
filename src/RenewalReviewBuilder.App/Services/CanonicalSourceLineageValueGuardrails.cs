using System.Globalization;
using System.Text.RegularExpressions;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed record CanonicalSourceLineageResolution(
    string ResolvedSourceFileName,
    string Status,
    string Message);

public sealed record CanonicalPercentageGuardrailResult(
    decimal? NormalizedValue,
    string Status,
    string Message);

public static partial class CanonicalSourceLineageValueGuardrails
{
    public static CanonicalSourceLineageResolution ResolveSourceFileName(
        long? sourceDocumentId,
        string? trackedSourceFileName,
        string? linkedDocumentFileName)
    {
        var tracked = FileName(trackedSourceFileName);
        var linked = FileName(linkedDocumentFileName);

        if (sourceDocumentId.HasValue)
        {
            if (string.IsNullOrWhiteSpace(linked))
            {
                return new CanonicalSourceLineageResolution(
                    tracked,
                    CanonicalSourceLineageStatuses.NeedsReview,
                    $"Source document {sourceDocumentId.Value.ToString(CultureInfo.InvariantCulture)} is linked, but its document filename is missing.");
            }

            if (!string.IsNullOrWhiteSpace(tracked) &&
                !tracked.Equals(linked, StringComparison.OrdinalIgnoreCase))
            {
                return new CanonicalSourceLineageResolution(
                    linked,
                    CanonicalSourceLineageStatuses.NeedsReview,
                    $"Tracked source filename '{tracked}' conflicts with linked document filename '{linked}'. The linked document filename is canonical.");
            }

            return new CanonicalSourceLineageResolution(
                linked,
                CanonicalSourceLineageStatuses.Verified,
                "Source filename matches the linked document record.");
        }

        if (!string.IsNullOrWhiteSpace(tracked))
        {
            return new CanonicalSourceLineageResolution(
                tracked,
                CanonicalSourceLineageStatuses.Unlinked,
                "Source filename is preserved, but no document record is linked.");
        }

        return new CanonicalSourceLineageResolution(
            linked,
            string.IsNullOrWhiteSpace(linked)
                ? CanonicalSourceLineageStatuses.Missing
                : CanonicalSourceLineageStatuses.Unlinked,
            string.IsNullOrWhiteSpace(linked)
                ? "No source filename or linked document was available."
                : "Document filename is preserved without a linked document identifier.");
    }

    public static decimal? ParseSinglePercentage(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        var matches = NumberRegex().Matches(value).ToList();
        if (matches.Count != 1 ||
            !decimal.TryParse(matches[0].Groups["value"].Value, NumberStyles.Number, CultureInfo.InvariantCulture, out var parsed))
        {
            return null;
        }

        if (!value.Contains('%') && parsed is > 0m and <= 1m)
        {
            parsed *= 100m;
        }

        return NormalizePercentage(parsed);
    }

    public static decimal? NormalizePercentage(decimal? value)
    {
        return value is >= 0m and <= 100m ? value : null;
    }

    public static bool IsPercentageOutOfRange(decimal? value)
    {
        return value.HasValue && value is < 0m or > 100m;
    }

    public static bool HasOutOfRangePercentageText(string? value)
    {
        return ExtractPercentageValues(value).Any(item => item is < 0m or > 100m);
    }

    public static bool HasMultiplePercentageValues(string? value)
    {
        return ExtractPercentageValues(value).Count > 1;
    }

    public static CanonicalPercentageGuardrailResult EvaluatePercentage(
        decimal? structuredValue,
        params string?[] sourceValues)
    {
        var percentages = sourceValues
            .SelectMany(ExtractPercentageValues)
            .Distinct()
            .ToList();
        if (IsPercentageOutOfRange(structuredValue) || percentages.Any(value => value is < 0m or > 100m))
        {
            return new CanonicalPercentageGuardrailResult(
                null,
                CanonicalValueGuardrailStatuses.Rejected,
                "An out-of-range reimbursement percentage was rejected; the source text remains available for review.");
        }

        if (percentages.Count > 1)
        {
            return new CanonicalPercentageGuardrailResult(
                NormalizePercentage(structuredValue),
                CanonicalValueGuardrailStatuses.PreservedForReview,
                "Multiple reimbursement percentages were preserved as source text and require structured review.");
        }

        var normalized = NormalizePercentage(structuredValue) ?? percentages.FirstOrDefault();
        if (structuredValue.HasValue || percentages.Count == 1)
        {
            return new CanonicalPercentageGuardrailResult(
                normalized,
                CanonicalValueGuardrailStatuses.Valid,
                "Reimbursement percentage is within the canonical 0-100 range.");
        }

        return new CanonicalPercentageGuardrailResult(
            null,
            CanonicalValueGuardrailStatuses.NotApplicable,
            "No reimbursement percentage was supplied for this fact.");
    }

    public static IReadOnlyList<decimal> ExtractPercentageValues(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return [];
        }

        return PercentageRegex()
            .Matches(value)
            .Select(match => decimal.TryParse(
                match.Groups["value"].Value,
                NumberStyles.Number,
                CultureInfo.InvariantCulture,
                out var parsed)
                    ? (decimal?)parsed
                    : null)
            .Where(parsed => parsed.HasValue)
            .Select(parsed => parsed!.Value)
            .ToList();
    }

    private static string FileName(string? value)
    {
        return string.IsNullOrWhiteSpace(value) ? "" : Path.GetFileName(value.Trim());
    }

    [GeneratedRegex(@"(?<![\d.])(?<value>-?\d+(?:\.\d+)?)(?![\d.])")]
    private static partial Regex NumberRegex();

    [GeneratedRegex(@"(?<![\d.])(?<value>-?\d+(?:\.\d+)?)\s*%(?!\d)")]
    private static partial Regex PercentageRegex();
}

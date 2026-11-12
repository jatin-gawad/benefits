using System.Globalization;
using System.Security.Cryptography;
using System.Text;

namespace RenewalReviewBuilder.App.Models;

public static class ClaimsRangeBenefitKeys
{
    public const string Health = "Health";
    public const string Dental = "Dental";
}

public static class ClaimsRangeComparisonStatuses
{
    public const string Ready = "Ready";
    public const string NeedsReview = "Needs Review";
}

public static class ClaimsRangeValidationSeverities
{
    public const string Warning = "Warning";
    public const string NeedsReview = "Needs Review";
    public const string Error = "Error";
}

public static class ClaimsRangeValidationCodes
{
    public const string CurrentReportMissing = "CLAIMS_RANGE_CURRENT_REPORT_MISSING";
    public const string PriorReportMissing = "CLAIMS_RANGE_PRIOR_REPORT_MISSING";
    public const string ReportIncomplete = "CLAIMS_RANGE_REPORT_INCOMPLETE";
    public const string PolicyMismatch = "CLAIMS_RANGE_POLICY_MISMATCH";
    public const string ClientNameMismatch = "CLAIMS_RANGE_CLIENT_NAME_MISMATCH";
    public const string PeriodNotContiguous = "CLAIMS_RANGE_PERIOD_NOT_CONTIGUOUS";
    public const string DuplicatePeriod = "CLAIMS_RANGE_DUPLICATE_PERIOD";
    public const string SourceChangedSinceAnalysis = "CLAIMS_RANGE_SOURCE_CHANGED_SINCE_ANALYSIS";
}

public static class ClaimsRangeCodes
{
    public const string From0To499 = "RANGE_0_499";
    public const string From500To999 = "RANGE_500_999";
    public const string From1000To1999 = "RANGE_1000_1999";
    public const string From2000To2999 = "RANGE_2000_2999";
    public const string From3000To3999 = "RANGE_3000_3999";
    public const string From4000To4999 = "RANGE_4000_4999";
    public const string From5000To7499 = "RANGE_5000_7499";
    public const string From7500To9999 = "RANGE_7500_9999";
    public const string Over10000 = "RANGE_OVER_10000";
}

public static class ClaimsRangeCanonicalFactContract
{
    public const string MeasurementType = "ClaimingCertificates";
    public const string DimensionType = "AnnualPaidClaimsRange";
    public const string Unit = "Certificates";
}

public sealed class ClaimsRangeBandDefinition
{
    public string RangeCode { get; init; } = "";
    public string DisplayLabel { get; init; } = "";
    public int SortOrder { get; init; }
    public int LowerBoundInclusive { get; init; }
    public int? UpperBoundInclusive { get; init; }
    public string PdfMatchKey { get; init; } = "";
}

public static class ClaimsRangeBandCatalog
{
    public static IReadOnlyList<ClaimsRangeBandDefinition> All { get; } =
    [
        Band(ClaimsRangeCodes.From0To499, "$0-$499", 0, 0, 499, "0-499"),
        Band(ClaimsRangeCodes.From500To999, "$500-$999", 1, 500, 999, "500-999"),
        Band(ClaimsRangeCodes.From1000To1999, "$1,000-$1,999", 2, 1000, 1999, "1000-1999"),
        Band(ClaimsRangeCodes.From2000To2999, "$2,000-$2,999", 3, 2000, 2999, "2000-2999"),
        Band(ClaimsRangeCodes.From3000To3999, "$3,000-$3,999", 4, 3000, 3999, "3000-3999"),
        Band(ClaimsRangeCodes.From4000To4999, "$4,000-$4,999", 5, 4000, 4999, "4000-4999"),
        Band(ClaimsRangeCodes.From5000To7499, "$5,000-$7,499", 6, 5000, 7499, "5000-7499"),
        Band(ClaimsRangeCodes.From7500To9999, "$7,500-$9,999", 7, 7500, 9999, "7500-9999"),
        Band(ClaimsRangeCodes.Over10000, "> $10,000", 8, 10001, null, ">10000")
    ];

    private static ClaimsRangeBandDefinition Band(
        string code,
        string label,
        int sortOrder,
        int lower,
        int? upper,
        string matchKey) =>
        new()
        {
            RangeCode = code,
            DisplayLabel = label,
            SortOrder = sortOrder,
            LowerBoundInclusive = lower,
            UpperBoundInclusive = upper,
            PdfMatchKey = matchKey
        };
}

public sealed class ClaimsRangeBoundingBox
{
    public decimal Left { get; init; }
    public decimal Bottom { get; init; }
    public decimal Right { get; init; }
    public decimal Top { get; init; }
    public decimal Width => Right - Left;
    public decimal Height => Top - Bottom;
    public decimal CenterX => Left + (Width / 2m);
}

public sealed class ClaimsRangeRow
{
    public string MeasurementType => ClaimsRangeCanonicalFactContract.MeasurementType;
    public string DimensionType => ClaimsRangeCanonicalFactContract.DimensionType;
    public string Unit => ClaimsRangeCanonicalFactContract.Unit;
    public string Benefit { get; set; } = "";
    public string RangeCode { get; set; } = "";
    public string RangeLabel { get; set; } = "";
    public string RawRangeLabel { get; set; } = "";
    public int SortOrder { get; set; }
    public int LowerBoundInclusive { get; set; }
    public int? UpperBoundInclusive { get; set; }
    public int CertificateCount { get; set; }
    public int SourcePage { get; set; }
    public string SourceReference { get; set; } = "";
    public ClaimsRangeBoundingBox LabelBounds { get; set; } = new();
    public ClaimsRangeBoundingBox CountBounds { get; set; } = new();
}

public sealed class ClaimsRangePage
{
    public string Benefit { get; set; } = "";
    public string BenefitDisplayName { get; set; } = "";
    public int SourcePage { get; set; }
    public string SourceReference { get; set; } = "";
    public List<ClaimsRangeRow> Rows { get; } = [];
    public int TotalCertificateCount => Rows.Sum(row => row.CertificateCount);
    public ClaimsRangeRow? HighestActiveRange => Rows
        .Where(row => row.CertificateCount > 0)
        .OrderByDescending(row => row.SortOrder)
        .FirstOrDefault();
}

public sealed class ClaimsRangeReport
{
    public string SourcePath { get; set; } = "";
    public string SourceFileName { get; set; } = "";
    public bool SourceSnapshotIsAuthoritative { get; set; }
    public long? SourceFileLength { get; set; }
    public DateTime? SourceLastWriteTimeUtc { get; set; }
    public string SourceContentSha256 { get; set; } = "";
    public string Carrier { get; set; } = "GroupHEALTH";
    public string ClientName { get; set; } = "";
    public string PolicyNumber { get; set; } = "";
    public DateOnly? PeriodStart { get; set; }
    public DateOnly? PeriodEnd { get; set; }
    public List<ClaimsRangePage> Pages { get; } = [];
    public List<string> Warnings { get; } = [];
    public List<string> Errors { get; } = [];
    public IEnumerable<ClaimsRangeRow> Rows => Pages.SelectMany(page => page.Rows);
    public int RowCount => Rows.Count();
    public bool IsComplete => IsGroupHealthReport()
        ? Errors.Count == 0 &&
          Pages.Count == 2 &&
          Pages.Select(page => page.Benefit).Distinct(StringComparer.OrdinalIgnoreCase).Count() == 2 &&
          Pages.All(page => page.Rows.Count == ClaimsRangeBandCatalog.All.Count) &&
          RowCount == ClaimsRangeBandCatalog.All.Count * 2
        : Errors.Count == 0 &&
          Pages.Count > 0 &&
          Pages.Select(page => page.Benefit).Distinct(StringComparer.OrdinalIgnoreCase).Count() == Pages.Count &&
          Pages.All(page => page.Rows.Count > 0 &&
                            page.Rows.Select(row => row.RangeCode)
                                .Distinct(StringComparer.OrdinalIgnoreCase).Count() == page.Rows.Count);
    public string PeriodLabel => PeriodStart.HasValue && PeriodEnd.HasValue
        ? $"{PeriodStart:MMM yyyy}-{PeriodEnd:MMM yyyy}"
        : "";

    private bool IsGroupHealthReport() =>
        Carrier.Equals("GroupHEALTH", StringComparison.OrdinalIgnoreCase);
}

public sealed class ClaimsRangeComparisonRow
{
    public string Benefit { get; set; } = "";
    public string RangeCode { get; set; } = "";
    public string RangeLabel { get; set; } = "";
    public int SortOrder { get; set; }
    public int LowerBoundInclusive { get; set; }
    public int? UpperBoundInclusive { get; set; }
    public int PriorCount { get; set; }
    public int CurrentCount { get; set; }
    public int? PriorSourcePage { get; set; }
    public int? CurrentSourcePage { get; set; }
    public string PriorSourceReference { get; set; } = "";
    public string CurrentSourceReference { get; set; } = "";
}

public sealed class ClaimsRangeValidationIssue
{
    public string Code { get; set; } = "";
    public string Severity { get; set; } = ClaimsRangeValidationSeverities.NeedsReview;
    public string Message { get; set; } = "";
    public bool CanConfirmForSession { get; set; }
}

public sealed class ClaimsRangeComparison
{
    private bool _identityConfirmedForSession;
    private string _identityConfirmationNote = "";
    private string _identityConfirmedBy = "";
    private DateTimeOffset? _identityConfirmedAt;
    private string _confirmedSourcePairFingerprint = "";

    public ClaimsRangeReport? PriorReport { get; set; }
    public ClaimsRangeReport? CurrentReport { get; set; }
    public List<ClaimsRangeComparisonRow> Rows { get; } = [];
    public List<ClaimsRangeValidationIssue> ValidationIssues { get; } = [];

    // This confirmation intentionally remains in-memory. Persistence requires a separate reviewed workflow.
    public bool IdentityConfirmedForSession => EnsureIdentityConfirmationIsCurrent(out _);
    public string IdentityConfirmationNote
    {
        get
        {
            _ = EnsureIdentityConfirmationIsCurrent(out _);
            return _identityConfirmationNote;
        }
    }
    public string IdentityConfirmedBy
    {
        get
        {
            _ = EnsureIdentityConfirmationIsCurrent(out _);
            return _identityConfirmedBy;
        }
    }
    public DateTimeOffset? IdentityConfirmedAt
    {
        get
        {
            _ = EnsureIdentityConfirmationIsCurrent(out _);
            return _identityConfirmedAt;
        }
    }
    public string ConfirmedSourcePairFingerprint
    {
        get
        {
            _ = EnsureIdentityConfirmationIsCurrent(out _);
            return _confirmedSourcePairFingerprint;
        }
    }
    public string SourcePairFingerprint => ClaimsRangeConfirmationFingerprint.Build(this).Fingerprint;
    public bool SourceSnapshotsMatchAnalysis =>
        ClaimsRangeConfirmationFingerprint.Build(this).SourcesMatchAnalyzedSnapshots;

    public string GeneratedTakeaway { get; set; } = "";
    public string? TakeawayOverride { get; set; }
    public string EffectiveTakeaway => TakeawayOverride is null
        ? GeneratedTakeaway
        : TakeawayOverride.Trim();
    public bool IsTakeawayOverridden => TakeawayOverride is not null;

    public bool IdentityConfirmationRequired => ValidationIssues.Any(issue =>
        issue.Code.Equals(ClaimsRangeValidationCodes.ClientNameMismatch, StringComparison.OrdinalIgnoreCase));

    public bool HasUnresolvedBlockingIssues
    {
        get
        {
            var identityConfirmed = EnsureIdentityConfirmationIsCurrent(out var sourcesMatchAnalyzedSnapshots);
            return !sourcesMatchAnalyzedSnapshots || ValidationIssues.Any(issue =>
                issue.Severity.Equals(ClaimsRangeValidationSeverities.Error, StringComparison.OrdinalIgnoreCase) ||
                (issue.Severity.Equals(ClaimsRangeValidationSeverities.NeedsReview, StringComparison.OrdinalIgnoreCase) &&
                 (!issue.CanConfirmForSession || !identityConfirmed)));
        }
    }

    public string Status => HasUnresolvedBlockingIssues
        ? ClaimsRangeComparisonStatuses.NeedsReview
        : ClaimsRangeComparisonStatuses.Ready;

    public bool CanUseForClientOutput => Status.Equals(
        ClaimsRangeComparisonStatuses.Ready,
        StringComparison.OrdinalIgnoreCase);

    public void ConfirmIdentityForSession(string confirmedBy, string note)
    {
        if (!IdentityConfirmationRequired)
        {
            return;
        }

        var state = ClaimsRangeConfirmationFingerprint.Build(this);
        if (!state.SourcesMatchAnalyzedSnapshots)
        {
            ClearIdentityConfirmationCore();
            return;
        }

        _identityConfirmedForSession = true;
        _identityConfirmedBy = confirmedBy?.Trim() ?? "";
        _identityConfirmationNote = note?.Trim() ?? "";
        _identityConfirmedAt = DateTimeOffset.Now;
        _confirmedSourcePairFingerprint = state.Fingerprint;
    }

    public void ClearIdentityConfirmation()
    {
        ClearIdentityConfirmationCore();
    }

    private bool EnsureIdentityConfirmationIsCurrent(out bool sourcesMatchAnalyzedSnapshots)
    {
        var state = ClaimsRangeConfirmationFingerprint.Build(this);
        sourcesMatchAnalyzedSnapshots = state.SourcesMatchAnalyzedSnapshots;
        if (!_identityConfirmedForSession)
        {
            return false;
        }

        if (!IdentityConfirmationRequired ||
            !sourcesMatchAnalyzedSnapshots ||
            !_confirmedSourcePairFingerprint.Equals(state.Fingerprint, StringComparison.Ordinal))
        {
            ClearIdentityConfirmationCore();
            return false;
        }

        return true;
    }

    private void ClearIdentityConfirmationCore()
    {
        _identityConfirmedForSession = false;
        _identityConfirmedBy = "";
        _identityConfirmationNote = "";
        _identityConfirmedAt = null;
        _confirmedSourcePairFingerprint = "";
    }
}

internal sealed record ClaimsRangeConfirmationFingerprintState(
    string Fingerprint,
    bool SourcesMatchAnalyzedSnapshots);

internal static class ClaimsRangeConfirmationFingerprint
{
    public static ClaimsRangeConfirmationFingerprintState Build(ClaimsRangeComparison comparison)
    {
        var text = new StringBuilder();
        var sourcesMatch = true;
        AppendReport(text, "PRIOR", comparison.PriorReport, ref sourcesMatch);
        AppendReport(text, "CURRENT", comparison.CurrentReport, ref sourcesMatch);

        foreach (var row in comparison.Rows
                     .OrderBy(row => row.Benefit, StringComparer.OrdinalIgnoreCase)
                     .ThenBy(row => row.SortOrder)
                     .ThenBy(row => row.RangeCode, StringComparer.OrdinalIgnoreCase))
        {
            Append(text,
                "COMPARISON_ROW",
                row.Benefit,
                row.RangeCode,
                row.RangeLabel,
                row.SortOrder,
                row.LowerBoundInclusive,
                row.UpperBoundInclusive,
                row.PriorCount,
                row.CurrentCount,
                row.PriorSourcePage,
                row.CurrentSourcePage,
                row.PriorSourceReference,
                row.CurrentSourceReference);
        }

        foreach (var issue in comparison.ValidationIssues
                     .OrderBy(issue => issue.Code, StringComparer.OrdinalIgnoreCase)
                     .ThenBy(issue => issue.Message, StringComparer.Ordinal))
        {
            Append(text,
                "VALIDATION",
                issue.Code,
                issue.Severity,
                issue.Message,
                issue.CanConfirmForSession);
        }

        Append(text, "GENERATED_TAKEAWAY", comparison.GeneratedTakeaway);
        var hash = SHA256.HashData(Encoding.UTF8.GetBytes(text.ToString()));
        return new ClaimsRangeConfirmationFingerprintState(
            Convert.ToHexString(hash),
            sourcesMatch);
    }

    private static void AppendReport(
        StringBuilder text,
        string role,
        ClaimsRangeReport? report,
        ref bool sourcesMatch)
    {
        if (report is null)
        {
            Append(text, role, "MISSING_REPORT");
            sourcesMatch = false;
            return;
        }

        var liveSource = LiveSourceFingerprint(report, out var sourceMatchesSnapshot);
        sourcesMatch &= sourceMatchesSnapshot;
        Append(text,
            role,
            FullPath(report.SourcePath),
            report.SourceFileName,
            report.SourceSnapshotIsAuthoritative,
            report.SourceFileLength,
            report.SourceLastWriteTimeUtc?.ToUniversalTime().Ticks,
            report.SourceContentSha256,
            liveSource,
            report.Carrier,
            report.ClientName,
            report.PolicyNumber,
            report.PeriodStart?.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
            report.PeriodEnd?.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));

        foreach (var page in report.Pages
                     .OrderBy(page => page.Benefit, StringComparer.OrdinalIgnoreCase)
                     .ThenBy(page => page.SourcePage))
        {
            Append(text,
                "PAGE",
                page.Benefit,
                page.BenefitDisplayName,
                page.SourcePage,
                page.SourceReference);
            foreach (var row in page.Rows
                         .OrderBy(row => row.SortOrder)
                         .ThenBy(row => row.RangeCode, StringComparer.OrdinalIgnoreCase))
            {
                Append(text,
                    "REPORT_ROW",
                    row.Benefit,
                    row.RangeCode,
                    row.RangeLabel,
                    row.RawRangeLabel,
                    row.SortOrder,
                    row.LowerBoundInclusive,
                    row.UpperBoundInclusive,
                    row.CertificateCount,
                    row.SourcePage,
                    row.SourceReference,
                    row.LabelBounds.Left,
                    row.LabelBounds.Bottom,
                    row.LabelBounds.Right,
                    row.LabelBounds.Top,
                    row.CountBounds.Left,
                    row.CountBounds.Bottom,
                    row.CountBounds.Right,
                    row.CountBounds.Top);
            }
        }

        foreach (var error in report.Errors.Order(StringComparer.Ordinal))
        {
            Append(text, "ERROR", error);
        }

        foreach (var warning in report.Warnings.Order(StringComparer.Ordinal))
        {
            Append(text, "WARNING", warning);
        }
    }

    private static string LiveSourceFingerprint(ClaimsRangeReport report, out bool matchesSnapshot)
    {
        if (!report.SourceSnapshotIsAuthoritative)
        {
            matchesSnapshot = true;
            return $"UNTRACKED|{FullPath(report.SourcePath)}";
        }

        try
        {
            if (string.IsNullOrWhiteSpace(report.SourcePath) || !File.Exists(report.SourcePath))
            {
                matchesSnapshot = false;
                return $"MISSING|{FullPath(report.SourcePath)}";
            }

            var file = new FileInfo(report.SourcePath);
            using var stream = File.OpenRead(report.SourcePath);
            var hash = Convert.ToHexString(SHA256.HashData(stream));
            matchesSnapshot = report.SourceFileLength == file.Length &&
                              report.SourceLastWriteTimeUtc?.ToUniversalTime().Ticks == file.LastWriteTimeUtc.Ticks &&
                              report.SourceContentSha256.Equals(hash, StringComparison.OrdinalIgnoreCase);
            return string.Join('|',
                FullPath(report.SourcePath),
                file.Length.ToString(CultureInfo.InvariantCulture),
                file.LastWriteTimeUtc.Ticks.ToString(CultureInfo.InvariantCulture),
                hash);
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
            matchesSnapshot = false;
            return $"UNREADABLE|{FullPath(report.SourcePath)}|{exception.GetType().Name}";
        }
    }

    private static void Append(StringBuilder text, params object?[] values)
    {
        foreach (var value in values)
        {
            var rendered = value switch
            {
                null => "<null>",
                IFormattable formattable => formattable.ToString(null, CultureInfo.InvariantCulture),
                _ => value.ToString() ?? ""
            };
            text.Append(rendered.Length.ToString(CultureInfo.InvariantCulture));
            text.Append(':');
            text.Append(rendered);
            text.Append('|');
        }
        text.AppendLine();
    }

    private static string FullPath(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return "";
        }

        try
        {
            return Path.GetFullPath(path);
        }
        catch
        {
            return path;
        }
    }
}

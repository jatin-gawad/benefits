using System.Text.RegularExpressions;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed class ClaimsRangeComparisonBuilder
{
    private readonly ClaimsRangeCommentaryBuilder _commentaryBuilder;

    public ClaimsRangeComparisonBuilder(ClaimsRangeCommentaryBuilder? commentaryBuilder = null)
    {
        _commentaryBuilder = commentaryBuilder ?? new ClaimsRangeCommentaryBuilder();
    }

    public ClaimsRangeComparison Build(
        IEnumerable<ClaimsRangeReport> reports,
        DateOnly? renewalDate = null)
    {
        ArgumentNullException.ThrowIfNull(reports);
        var candidates = reports
            .Where(report => report is not null)
            .OrderBy(report => report.PeriodEnd ?? DateOnly.MinValue)
            .ThenBy(report => report.SourcePath, StringComparer.OrdinalIgnoreCase)
            .ToList();
        var comparison = new ClaimsRangeComparison();

        comparison.CurrentReport = SelectCurrent(candidates, renewalDate);
        if (comparison.CurrentReport is null)
        {
            AddIssue(
                comparison,
                ClaimsRangeValidationCodes.CurrentReportMissing,
                ClaimsRangeValidationSeverities.Error,
                "A current Claims Range report could not be selected.");
            comparison.GeneratedTakeaway = _commentaryBuilder.Build(comparison);
            return comparison;
        }

        comparison.PriorReport = SelectPrior(candidates, comparison.CurrentReport);
        if (comparison.PriorReport is null)
        {
            AddIssue(
                comparison,
                ClaimsRangeValidationCodes.PriorReportMissing,
                ClaimsRangeValidationSeverities.Error,
                "A prior Claims Range report could not be selected.");
        }

        ValidateDuplicatePeriods(candidates, comparison);
        ValidateCompleteness(comparison.CurrentReport, "current", comparison);
        if (comparison.PriorReport is not null)
        {
            ValidateCompleteness(comparison.PriorReport, "prior", comparison);
            ValidatePair(comparison.PriorReport, comparison.CurrentReport, comparison);
            BuildRows(comparison);
        }

        comparison.GeneratedTakeaway = _commentaryBuilder.Build(comparison);
        return comparison;
    }

    private static ClaimsRangeReport? SelectCurrent(
        IReadOnlyList<ClaimsRangeReport> candidates,
        DateOnly? renewalDate)
    {
        var dated = candidates.Where(report => report.PeriodEnd.HasValue);
        if (renewalDate.HasValue)
        {
            var completedBeforeRenewal = dated
                .Where(report => report.PeriodEnd <= renewalDate)
                .OrderByDescending(report => report.PeriodEnd)
                .ThenBy(report => report.SourcePath, StringComparer.OrdinalIgnoreCase)
                .FirstOrDefault();
            if (completedBeforeRenewal is not null)
            {
                return completedBeforeRenewal;
            }
        }

        return dated
            .OrderByDescending(report => report.PeriodEnd)
            .ThenBy(report => report.SourcePath, StringComparer.OrdinalIgnoreCase)
            .FirstOrDefault();
    }

    private static ClaimsRangeReport? SelectPrior(
        IReadOnlyList<ClaimsRangeReport> candidates,
        ClaimsRangeReport current)
    {
        var earlier = candidates
            .Where(report => !ReferenceEquals(report, current))
            .Where(report => report.PeriodEnd.HasValue &&
                             (!current.PeriodStart.HasValue || report.PeriodEnd < current.PeriodStart))
            .ToList();
        var samePolicy = earlier
            .Where(report => Id(report.PolicyNumber).Equals(Id(current.PolicyNumber), StringComparison.OrdinalIgnoreCase))
            .ToList();
        var pool = samePolicy.Count > 0 ? samePolicy : earlier;
        return pool
            .OrderByDescending(report =>
                current.PeriodStart.HasValue && report.PeriodEnd?.AddDays(1) == current.PeriodStart ? 1 : 0)
            .ThenByDescending(report => report.PeriodEnd)
            .ThenBy(report => report.SourcePath, StringComparer.OrdinalIgnoreCase)
            .FirstOrDefault();
    }

    private static void ValidateDuplicatePeriods(
        IReadOnlyList<ClaimsRangeReport> candidates,
        ClaimsRangeComparison comparison)
    {
        var duplicates = candidates
            .Where(report => report.PeriodStart.HasValue && report.PeriodEnd.HasValue)
            .GroupBy(report => $"{Id(report.PolicyNumber)}|{report.PeriodStart:yyyy-MM-dd}|{report.PeriodEnd:yyyy-MM-dd}")
            .Where(group => group.Count() > 1)
            .ToList();
        foreach (var duplicate in duplicates)
        {
            AddIssue(
                comparison,
                ClaimsRangeValidationCodes.DuplicatePeriod,
                ClaimsRangeValidationSeverities.NeedsReview,
                $"Multiple Claims Range reports were found for the same policy and period: {string.Join(", ", duplicate.Select(report => report.SourceFileName))}.");
        }
    }

    private static void ValidateCompleteness(
        ClaimsRangeReport report,
        string role,
        ClaimsRangeComparison comparison)
    {
        if (report.IsComplete)
        {
            return;
        }

        var detail = report.Errors.Count > 0
            ? string.Join(" ", report.Errors)
            : $"Only {report.RowCount} of 18 expected rows were extracted.";
        AddIssue(
            comparison,
            ClaimsRangeValidationCodes.ReportIncomplete,
            ClaimsRangeValidationSeverities.Error,
            $"The {role} Claims Range report is incomplete. {detail}");
    }

    private static void ValidatePair(
        ClaimsRangeReport prior,
        ClaimsRangeReport current,
        ClaimsRangeComparison comparison)
    {
        if (!Id(prior.PolicyNumber).Equals(Id(current.PolicyNumber), StringComparison.OrdinalIgnoreCase))
        {
            AddIssue(
                comparison,
                ClaimsRangeValidationCodes.PolicyMismatch,
                ClaimsRangeValidationSeverities.Error,
                $"Claims Range policy numbers do not match: prior '{prior.PolicyNumber}' and current '{current.PolicyNumber}'.");
        }

        if (!ClientId(prior.ClientName).Equals(ClientId(current.ClientName), StringComparison.OrdinalIgnoreCase))
        {
            AddIssue(
                comparison,
                ClaimsRangeValidationCodes.ClientNameMismatch,
                ClaimsRangeValidationSeverities.NeedsReview,
                $"Claims Range client names require confirmation: prior '{prior.ClientName}' and current '{current.ClientName}'.",
                canConfirmForSession: true);
        }

        if (prior.PeriodEnd.HasValue && current.PeriodStart.HasValue &&
            prior.PeriodEnd.Value.AddDays(1) != current.PeriodStart.Value)
        {
            AddIssue(
                comparison,
                ClaimsRangeValidationCodes.PeriodNotContiguous,
                ClaimsRangeValidationSeverities.NeedsReview,
                $"Claims Range periods are not contiguous: prior ends {prior.PeriodEnd:MMMM d, yyyy} and current starts {current.PeriodStart:MMMM d, yyyy}.");
        }
    }

    private static void BuildRows(ClaimsRangeComparison comparison)
    {
        var benefits = comparison.PriorReport!.Rows
            .Concat(comparison.CurrentReport!.Rows)
            .Select(row => row.Benefit)
            .Where(benefit => !string.IsNullOrWhiteSpace(benefit))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(benefit => benefit.Equals(ClaimsRangeBenefitKeys.Health, StringComparison.OrdinalIgnoreCase) ? 0 : 1)
            .ToList();
        var useCanonicalGroupHealthBands =
            comparison.PriorReport.Carrier.Equals("GroupHEALTH", StringComparison.OrdinalIgnoreCase) &&
            comparison.CurrentReport.Carrier.Equals("GroupHEALTH", StringComparison.OrdinalIgnoreCase);

        foreach (var benefit in benefits)
        {
            var priorRows = comparison.PriorReport!.Rows
                .Where(row => row.Benefit.Equals(benefit, StringComparison.OrdinalIgnoreCase))
                .GroupBy(row => row.RangeCode, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.OrdinalIgnoreCase);
            var currentRows = comparison.CurrentReport!.Rows
                .Where(row => row.Benefit.Equals(benefit, StringComparison.OrdinalIgnoreCase))
                .GroupBy(row => row.RangeCode, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.OrdinalIgnoreCase);

            var definitions = useCanonicalGroupHealthBands
                ? ClaimsRangeBandCatalog.All
                : priorRows.Values
                    .Concat(currentRows.Values)
                    .GroupBy(row => row.RangeCode, StringComparer.OrdinalIgnoreCase)
                    .Select(group => group.OrderByDescending(row => currentRows.ContainsKey(row.RangeCode)).First())
                    .OrderBy(row => row.LowerBoundInclusive)
                    .ThenBy(row => row.SortOrder)
                    .Select(row => new ClaimsRangeBandDefinition
                    {
                        RangeCode = row.RangeCode,
                        DisplayLabel = row.RangeLabel,
                        SortOrder = row.SortOrder,
                        LowerBoundInclusive = row.LowerBoundInclusive,
                        UpperBoundInclusive = row.UpperBoundInclusive,
                        PdfMatchKey = row.RawRangeLabel
                    })
                    .ToList();

            foreach (var definition in definitions)
            {
                priorRows.TryGetValue(definition.RangeCode, out var prior);
                currentRows.TryGetValue(definition.RangeCode, out var current);
                comparison.Rows.Add(new ClaimsRangeComparisonRow
                {
                    Benefit = benefit,
                    RangeCode = definition.RangeCode,
                    RangeLabel = definition.DisplayLabel,
                    SortOrder = definition.SortOrder,
                    LowerBoundInclusive = definition.LowerBoundInclusive,
                    UpperBoundInclusive = definition.UpperBoundInclusive,
                    PriorCount = prior?.CertificateCount ?? 0,
                    CurrentCount = current?.CertificateCount ?? 0,
                    PriorSourcePage = prior?.SourcePage,
                    CurrentSourcePage = current?.SourcePage,
                    PriorSourceReference = prior?.SourceReference ?? "",
                    CurrentSourceReference = current?.SourceReference ?? ""
                });
            }
        }
    }

    private static void AddIssue(
        ClaimsRangeComparison comparison,
        string code,
        string severity,
        string message,
        bool canConfirmForSession = false)
    {
        comparison.ValidationIssues.Add(new ClaimsRangeValidationIssue
        {
            Code = code,
            Severity = severity,
            Message = message,
            CanConfirmForSession = canConfirmForSession
        });
    }

    private static string Id(string value) => new(
        (value ?? "").Where(char.IsLetterOrDigit).Select(char.ToUpperInvariant).ToArray());

    private static string ClientId(string value)
    {
        var tokens = Regex.Matches((value ?? "").ToUpperInvariant(), @"[A-Z0-9]+")
            .Select(match => match.Value)
            .Where(token => token is not "INC" and not "INCORPORATED" and not "LTD" and not "LIMITED" and
                            not "CORP" and not "CORPORATION" and not "CO" and not "COMPANY")
            .ToList();
        return string.Concat(tokens);
    }
}

using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Models.History;
using RenewalReviewBuilder.App.Models.Querying;
using RenewalReviewBuilder.App.Services.CanonicalClaims;

namespace RenewalReviewBuilder.App.Services.ClientHistory;

public sealed class ClaimsHistoryContributor : IClientHistoryContributor
{
    private readonly Func<long, CancellationToken, Task<IReadOnlyList<CanonicalClaimsReportReadModel>>> _loader;

    public ClaimsHistoryContributor() : this((_, _) => Task.FromResult<IReadOnlyList<CanonicalClaimsReportReadModel>>([])) { }

    public ClaimsHistoryContributor(DatabaseService database) : this(async (clientId, cancellationToken) =>
    {
        cancellationToken.ThrowIfCancellationRequested();
        var reports = await new CanonicalClaimsQueryService(database.DatabasePath).GetAllForClientAsync(clientId);
        cancellationToken.ThrowIfCancellationRequested();
        return reports;
    }) { }

    internal ClaimsHistoryContributor(
        Func<long, CancellationToken, Task<IReadOnlyList<CanonicalClaimsReportReadModel>>> loader) => _loader = loader;

    public string SourceModule => "Claims";

    public async Task<IReadOnlyList<ClientHistoryEventReadModel>> ContributeAsync(
        long clientId,
        CancellationToken cancellationToken = default) =>
        (await LoadAndBuildAsync(clientId, null, cancellationToken)).Events;

    public async Task<ClientHistoryContribution> ContributeAsync(AsOfQueryContext context)
    {
        var result = await LoadAndBuildAsync(context.ClientId, context.AsOfDate, context.CancellationToken);
        return new ClientHistoryContribution
        {
            Events = result.Events,
            HistoricalSupport = new HistoricalQuerySupportStatus
            {
                SourceModule = SourceModule,
                Support = HistoricalQuerySupport.Limited,
                Detail = result.SourceCount == 0
                    ? "Claims replay has limited support because no canonical claims experience reports are saved."
                    : "Claims replay uses validated multi-month experience periods recorded by the replay date. Review-required, invalid, undated, short monthly reports, and later-recorded corrections are excluded; overwritten report versions cannot be reconstructed."
            }
        };
    }

    private async Task<BuildResult> LoadAndBuildAsync(
        long clientId,
        DateOnly? asOfDate,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var source = await _loader(clientId, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();

        var eligible = source
            .Where(report => report.ClientId == clientId &&
                             report.ValidationStatus.Equals(CanonicalClaimsValidationStatus.Valid.ToString(), StringComparison.OrdinalIgnoreCase) &&
                             IsMeaningfulExperiencePeriod(report.Period) &&
                             IsKnownBy(report, asOfDate) &&
                             HasUsableFacts(report))
            .ToList();
        var events = new List<ClientHistoryEventReadModel>();
        foreach (var group in eligible.GroupBy(NaturalPeriodKey, StringComparer.Ordinal))
        {
            var reports = group.OrderBy(item => item.CanonicalReportId).ToList();
            var representative = reports.OrderByDescending(item => item.UpdatedAt).ThenByDescending(item => item.CanonicalReportId).First();
            var reportIds = reports.Select(item => item.CanonicalReportId).Distinct().OrderBy(value => value).ToList();
            var related = reportIds.Select(id => $"CanonicalClaimsReport:{id}").ToList();
            foreach (var documentId in reports.Where(item => item.SourceDocumentId.HasValue).Select(item => item.SourceDocumentId!.Value).Distinct())
            {
                related.Add($"Document:{documentId}");
            }

            var evidence = reports
                .Where(item => item.SourceDocumentId.HasValue || item.SourceDocumentVersionId.HasValue)
                .GroupBy(item => new { item.SourceDocumentId, item.SourceDocumentVersionId })
                .Select(item => item.First())
                .Select(item => CanonicalHistoryEventFactory.Evidence(
                    item.CanonicalReportId.ToString(CultureInfo.InvariantCulture),
                    HistoryValidationStatus.Valid,
                    .95m,
                    item.SourceDocumentId,
                    item.SourceDocumentVersionId,
                    sourceFactId: $"CanonicalClaimsReport:{item.CanonicalReportId}"))
                .ToList();
            var sourceEntityId = $"{clientId.ToString(CultureInfo.InvariantCulture)}|{NaturalPeriodKey(representative)}";
            var stableKey = ClientHistoryProjectionBuilder.StableKey(SourceModule, "CanonicalClaimsReportPeriod", sourceEntityId, "experience-period");
            events.Add(CanonicalHistoryEventFactory.Create(
                SourceModule,
                "CanonicalClaimsReportPeriod",
                sourceEntityId,
                "experience-period",
                clientId,
                HistoryDateReference.FromExact(representative.Period.End!.Value),
                RecordedAt(representative),
                "Claims experience period",
                "Claims",
                $"Claims experience through {representative.Period.End:MMM d, yyyy}",
                PeriodSummary(representative),
                HistoryValidationStatus.Valid,
                .95m,
                evidence,
                related,
                stableKey: stableKey));
        }

        return new BuildResult(events, source.Count);
    }

    /// <summary>
    /// A report-level History milestone must describe more than one month of experience.
    /// Monthly child facts can support an otherwise qualified multi-month report, but a routine
    /// monthly report is never promoted into its own strategic History event.
    /// </summary>
    private static bool IsMeaningfulExperiencePeriod(CanonicalClaimsPeriod period)
    {
        if (!period.Start.HasValue || !period.End.HasValue) return false;
        if (period.End.Value <= period.Start.Value.AddMonths(1)) return false;
        var periodType = period.PeriodType.Trim();
        return !periodType.Equals("Monthly", StringComparison.OrdinalIgnoreCase) &&
               !periodType.Equals("Month", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsKnownBy(CanonicalClaimsReportReadModel report, DateOnly? asOfDate)
    {
        if (!asOfDate.HasValue) return true;
        var recordedAt = RecordedAt(report);
        return recordedAt.HasValue && DateOnly.FromDateTime(recordedAt.Value) <= asOfDate.Value;
    }

    private static DateTime? RecordedAt(CanonicalClaimsReportReadModel report) =>
        report.UpdatedAt != default
            ? report.UpdatedAt
            : report.CreatedAt != default
                ? report.CreatedAt
                : null;

    private static bool HasUsableFacts(CanonicalClaimsReportReadModel report) =>
        report.FinancialFacts.Any(Usable) ||
        report.MonthlyFacts.Any(Usable) ||
        report.CategoryFacts.Any(Usable) ||
        report.DistributionFacts.Any(Usable) ||
        report.BenchmarkFacts.Any(Usable) ||
        report.UtilizationFacts.Any(item =>
            item.ValidationStatus == CanonicalClaimsValidationStatus.Valid &&
            item.RecordStatus.Equals("Active", StringComparison.OrdinalIgnoreCase) &&
            (item.UniqueClaimants.HasValue || item.Occurrences.HasValue || item.SubmittedAmount.HasValue || item.EligibleAmount.HasValue || item.PaidAmount.HasValue));

    private static bool Usable(CanonicalClaimFact fact) =>
        fact.ValidationStatus == CanonicalClaimsValidationStatus.Valid &&
        (fact.Availability is CanonicalDataAvailability.Present or CanonicalDataAvailability.ReportedZero);

    private static string NaturalPeriodKey(CanonicalClaimsReportReadModel report) => string.Join("|",
        report.Carrier.Trim().ToUpperInvariant(),
        report.PolicyNumber.Trim().ToUpperInvariant(),
        report.Period.Start?.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) ?? "",
        report.Period.End?.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) ?? "",
        report.Period.PeriodType.Trim().ToUpperInvariant());

    private static string PeriodSummary(CanonicalClaimsReportReadModel report)
    {
        var parts = new List<string>();
        if (!string.IsNullOrWhiteSpace(report.Carrier)) parts.Add(report.Carrier.Trim());
        if (!string.IsNullOrWhiteSpace(report.PolicyNumber)) parts.Add($"policy {report.PolicyNumber.Trim()}");
        if (report.Period.Start.HasValue && report.Period.End.HasValue)
        {
            parts.Add($"{report.Period.Start:MMM d, yyyy} to {report.Period.End:MMM d, yyyy}");
        }
        if (!string.IsNullOrWhiteSpace(report.Period.PeriodType)) parts.Add(report.Period.PeriodType.Trim());
        return parts.Count == 0 ? "Validated canonical claims experience period." : string.Join(" · ", parts) + ".";
    }

    private sealed record BuildResult(IReadOnlyList<ClientHistoryEventReadModel> Events, int SourceCount);
}

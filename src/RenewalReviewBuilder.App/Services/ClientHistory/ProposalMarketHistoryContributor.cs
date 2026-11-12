using System.Globalization;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Models.History;
using RenewalReviewBuilder.App.Models.Querying;

namespace RenewalReviewBuilder.App.Services.ClientHistory;

public sealed class ProposalMarketHistoryContributor : IClientHistoryContributor
{
    private static readonly Regex YearOnly = new("^\\d{4}$", RegexOptions.CultureInvariant);
    private static readonly Regex MonthOnly = new("^(?<year>\\d{4})-(?<month>\\d{2})$", RegexOptions.CultureInvariant);
    private readonly Func<long, CancellationToken, Task<ProposalCanonicalProfile>> _loader;

    public ProposalMarketHistoryContributor() : this((clientId, _) => Task.FromResult(new ProposalCanonicalProfile { ClientId = clientId })) { }

    public ProposalMarketHistoryContributor(DatabaseService database) : this(async (clientId, cancellationToken) =>
    {
        cancellationToken.ThrowIfCancellationRequested();
        var profile = await new ProposalCanonicalProfileService(database.DatabasePath).BuildAsync(clientId);
        cancellationToken.ThrowIfCancellationRequested();
        return profile;
    }) { }

    internal ProposalMarketHistoryContributor(Func<long, CancellationToken, Task<ProposalCanonicalProfile>> loader) => _loader = loader;

    public string SourceModule => "ProposalMarket";

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
                    ? "Proposal replay has limited support because no canonical proposal records are saved."
                    : result.RecordedTimeExcludedCount > 0
                        ? "Proposal replay excludes proposals saved or updated after the selected date, and proposals without a canonical recorded time. Proposal status-effective history is not recorded."
                        : "Proposal replay includes factual received dates from usable canonical proposal profiles. Review-required, undated, and inferred lifecycle stages are excluded; status-effective history is not recorded."
            }
        };
    }

    private async Task<BuildResult> LoadAndBuildAsync(
        long clientId,
        DateOnly? asOfDate,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var profile = await _loader(clientId, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        var events = new List<ClientHistoryEventReadModel>();
        var recordedTimeExcludedCount = 0;

        if (profile.ClientId != clientId)
        {
            return new BuildResult([], profile.Proposals.Count, profile.Proposals.Count);
        }

        foreach (var proposal in profile.Proposals.OrderBy(item => item.ProposalId))
        {
            var recordedAt = Recorded(proposal.UpdatedAt, proposal.CreatedAt);
            if (!WasKnownBy(recordedAt, asOfDate))
            {
                recordedTimeExcludedCount++;
                continue;
            }

            if (proposal.ReadinessStatus is ProposalCanonicalReadinessStatuses.NeedsReview or ProposalCanonicalReadinessStatuses.NotReady)
            {
                continue;
            }

            var date = ProposalDate(proposal);
            if (date.Precision == HistoryDatePrecision.Unknown)
            {
                continue;
            }

            var carrier = proposal.CarrierName.Equals("Unknown Carrier", StringComparison.OrdinalIgnoreCase)
                ? ""
                : proposal.CarrierName.Trim();
            var confidence = CanonicalHistoryEventFactory.ClampConfidence(
                proposal.Evidence.ConfidenceScore ?? proposal.ExtractionConfidence);
            var related = new List<string> { $"Proposal:{proposal.ProposalId}" };
            if (proposal.CarrierId.HasValue) related.Add($"Carrier:{proposal.CarrierId.Value}");
            var summary = new List<string>();
            if (!string.IsNullOrWhiteSpace(carrier)) summary.Add($"Carrier: {carrier}");
            var quotedEffectiveDate = DisplayDate(proposal.EffectiveDateSourceText, proposal.EffectiveDate);
            if (!string.IsNullOrWhiteSpace(quotedEffectiveDate)) summary.Add($"quoted effective date: {quotedEffectiveDate}");

            events.Add(CanonicalHistoryEventFactory.Create(
                SourceModule,
                "Proposal",
                proposal.ProposalId.ToString(CultureInfo.InvariantCulture),
                "received",
                clientId,
                date,
                recordedAt,
                "Proposal received",
                "Market",
                string.IsNullOrWhiteSpace(carrier) ? "Proposal received" : $"{carrier} proposal received",
                summary.Count == 0 ? "Canonical proposal received." : string.Join(" · ", summary) + ".",
                HistoryValidationStatus.Valid,
                confidence,
                Evidence(proposal, confidence),
                related));
        }

        return new BuildResult(events, profile.Proposals.Count, recordedTimeExcludedCount);
    }

    private static IReadOnlyList<HistoryEvidenceReference> Evidence(ProposalCanonicalQuote proposal, decimal confidence)
    {
        var evidence = proposal.Evidence;
        if (string.IsNullOrWhiteSpace(evidence.SourceTable) &&
            !evidence.SourceId.HasValue &&
            string.IsNullOrWhiteSpace(evidence.SourceFileName) &&
            string.IsNullOrWhiteSpace(proposal.SourceFileName))
        {
            return [];
        }

        var sourceFact = !string.IsNullOrWhiteSpace(evidence.SourceTable)
            ? $"{evidence.SourceTable}:{evidence.SourceId?.ToString(CultureInfo.InvariantCulture) ?? proposal.ProposalId.ToString(CultureInfo.InvariantCulture)}"
            : proposal.SourceFileHash;
        return
        [
            CanonicalHistoryEventFactory.Evidence(
                proposal.ProposalId.ToString(CultureInfo.InvariantCulture),
                HistoryValidationStatus.Valid,
                confidence,
                sourceFactId: sourceFact,
                page: evidence.SourcePage,
                sheet: evidence.SourceSheet,
                section: FirstNonBlank(evidence.SourceLocation, evidence.SourceFileName, proposal.SourceFileName),
                extractedText: evidence.RawPayloadJson)
        ];
    }

    private static HistoryDateReference ProposalDate(ProposalCanonicalQuote proposal)
    {
        var raw = proposal.ProposalDateSourceText.Trim();
        if (YearOnly.IsMatch(raw) && int.TryParse(raw, NumberStyles.None, CultureInfo.InvariantCulture, out var year))
        {
            return HistoryDateReference.FromYear(year);
        }

        var month = MonthOnly.Match(raw);
        if (month.Success &&
            int.TryParse(month.Groups["year"].Value, NumberStyles.None, CultureInfo.InvariantCulture, out year) &&
            int.TryParse(month.Groups["month"].Value, NumberStyles.None, CultureInfo.InvariantCulture, out var monthNumber) &&
            monthNumber is >= 1 and <= 12)
        {
            return HistoryDateReference.FromMonth(year, monthNumber);
        }

        return proposal.ProposalDate.HasValue
            ? HistoryDateReference.FromExact(proposal.ProposalDate.Value)
            : new HistoryDateReference();
    }

    private static DateTime? Recorded(params DateTime?[] values)
    {
        var value = values.Where(item => item.HasValue).Select(item => item!.Value).DefaultIfEmpty().Max();
        return value == default ? null : value;
    }

    private static bool WasKnownBy(DateTime? recordedAt, DateOnly? asOfDate) =>
        !asOfDate.HasValue ||
        recordedAt.HasValue && DateOnly.FromDateTime(recordedAt.Value) <= asOfDate.Value;

    private static string DisplayDate(string sourceText, DateOnly? parsedDate)
    {
        var raw = sourceText.Trim();
        if (YearOnly.IsMatch(raw)) return raw;
        var month = MonthOnly.Match(raw);
        if (month.Success &&
            int.TryParse(month.Groups["year"].Value, NumberStyles.None, CultureInfo.InvariantCulture, out var year) &&
            int.TryParse(month.Groups["month"].Value, NumberStyles.None, CultureInfo.InvariantCulture, out var monthNumber) &&
            monthNumber is >= 1 and <= 12)
        {
            return new DateOnly(year, monthNumber, 1).ToString("MMMM yyyy", CultureInfo.InvariantCulture);
        }

        return parsedDate?.ToString("MMM d, yyyy", CultureInfo.InvariantCulture) ?? "";
    }

    private static string FirstNonBlank(params string[] values) =>
        values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value))?.Trim() ?? "";

    private sealed record BuildResult(
        IReadOnlyList<ClientHistoryEventReadModel> Events,
        int SourceCount,
        int RecordedTimeExcludedCount);
}

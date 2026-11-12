using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Models.History;

namespace RenewalReviewBuilder.App.Services.ClientHistory;

public sealed class HistoryRelationshipTimelineService
{
    public IReadOnlyList<HistoryRelationshipMilestone> BuildMilestones(IEnumerable<ClientHistoryEventReadModel> events)
    {
        var representatives = events
            .Where(Qualifies)
            .GroupBy(item => DateReference(item).Year)
            .Select(group => group
                .OrderByDescending(Priority)
                .ThenByDescending(item => DateReference(item).SortMonth)
                .ThenByDescending(item => DateReference(item).SortDay)
                .ThenBy(item => item.EventId, StringComparer.Ordinal)
                .First())
            .OrderBy(item => DateReference(item).SortYear)
            .ThenBy(item => DateReference(item).SortMonth)
            .ThenBy(item => DateReference(item).SortDay)
            .ToList();

        var visible = representatives.Count <= 6
            ? representatives
            : new[] { representatives[0] }.Concat(representatives.TakeLast(5)).ToList();

        return visible.Select(item =>
        {
            var date = DateReference(item);
            return new HistoryRelationshipMilestone
            {
                EventId = item.EventId,
                EffectiveDate = item.EffectiveDate,
                DateReference = date,
                Label = item.Title,
                IconGlyph = Icon(item),
                AccentColor = AccentColor(item),
                AccentBackground = AccentBackground(item),
                ScreenReaderDescription = string.IsNullOrWhiteSpace(date.DisplayText) || date.Precision == HistoryDatePrecision.Unknown
                    ? item.Title
                    : $"{item.Title} {date.DisplayText}"
            };
        }).ToList();
    }
    public bool Qualifies(ClientHistoryEventReadModel historyEvent) =>
        !historyEvent.IsExcluded &&
        DateReference(historyEvent).Precision != HistoryDatePrecision.Unknown &&
        (
            historyEvent.EventType.Contains("onboard", StringComparison.OrdinalIgnoreCase) ||
            historyEvent.Title.Contains("onboard", StringComparison.OrdinalIgnoreCase) ||
            historyEvent.EventType.Contains("benchmark", StringComparison.OrdinalIgnoreCase) ||
            historyEvent.Title.Contains("benchmark", StringComparison.OrdinalIgnoreCase) ||
            historyEvent.Category.Contains("Carrier", StringComparison.OrdinalIgnoreCase) ||
            historyEvent.EventType.Contains("Carrier", StringComparison.OrdinalIgnoreCase) ||
            historyEvent.EventType.Contains("market", StringComparison.OrdinalIgnoreCase) ||
            historyEvent.EventType.Contains("training", StringComparison.OrdinalIgnoreCase) ||
            historyEvent.Title.Contains("training", StringComparison.OrdinalIgnoreCase) ||
            (historyEvent.EventType.Contains("renewal", StringComparison.OrdinalIgnoreCase) &&
             historyEvent.EventType.Contains("current", StringComparison.OrdinalIgnoreCase)) ||
            (historyEvent.Category.Contains("Plan", StringComparison.OrdinalIgnoreCase) &&
             historyEvent.ChangeSet.Changes.Count >= 2)
        );

    private static HistoryDateReference DateReference(ClientHistoryEventReadModel historyEvent) =>
        historyEvent.EffectiveDateReference.Precision != HistoryDatePrecision.Unknown
            ? historyEvent.EffectiveDateReference
            : historyEvent.EffectiveDate.HasValue
                ? HistoryDateReference.FromExact(historyEvent.EffectiveDate.Value)
                : new HistoryDateReference();

    private static int Priority(ClientHistoryEventReadModel historyEvent) =>
        historyEvent.Category.Contains("Carrier", StringComparison.OrdinalIgnoreCase) || historyEvent.EventType.Contains("Carrier", StringComparison.OrdinalIgnoreCase) ? 50 :
        historyEvent.Title.Contains("training", StringComparison.OrdinalIgnoreCase) ? 40 :
        historyEvent.EventType.Contains("market", StringComparison.OrdinalIgnoreCase) || historyEvent.Title.Contains("benchmark", StringComparison.OrdinalIgnoreCase) ? 30 :
        historyEvent.Category.Contains("Plan", StringComparison.OrdinalIgnoreCase) ? 20 : 10;

    private static string AccentColor(ClientHistoryEventReadModel historyEvent) =>
        historyEvent.Category.Contains("Carrier", StringComparison.OrdinalIgnoreCase) ? "#0B5DE8" :
        historyEvent.Title.Contains("training", StringComparison.OrdinalIgnoreCase) || historyEvent.Category.Contains("Administration", StringComparison.OrdinalIgnoreCase) ? "#6D28D9" :
        historyEvent.EventType.Contains("market", StringComparison.OrdinalIgnoreCase) || historyEvent.Title.Contains("benchmark", StringComparison.OrdinalIgnoreCase) ? "#F47A13" : "#18A55B";

    private static string AccentBackground(ClientHistoryEventReadModel historyEvent) =>
        historyEvent.Category.Contains("Carrier", StringComparison.OrdinalIgnoreCase) ? "#EEF4FF" :
        historyEvent.Title.Contains("training", StringComparison.OrdinalIgnoreCase) || historyEvent.Category.Contains("Administration", StringComparison.OrdinalIgnoreCase) ? "#F3EDFF" :
        historyEvent.EventType.Contains("market", StringComparison.OrdinalIgnoreCase) || historyEvent.Title.Contains("benchmark", StringComparison.OrdinalIgnoreCase) ? "#FFF1E6" : "#EAF8F0";

    private static string Icon(ClientHistoryEventReadModel historyEvent) =>
        historyEvent.Category.Contains("Carrier", StringComparison.OrdinalIgnoreCase) ? "\uE7F4" :
        historyEvent.Title.Contains("training", StringComparison.OrdinalIgnoreCase) ? "\uE716" :
        historyEvent.EventType.Contains("market", StringComparison.OrdinalIgnoreCase) ? "\uE9D2" : "\uE81C";
}

public sealed class HistoryRenewalEvolutionService
{
    public IReadOnlyList<HistoryRenewalEvolutionPoint> Build(IEnumerable<RenewalVersionRecord> renewalVersions)
    {
        var ordered = renewalVersions
            .GroupBy(RenewalIdentityKey)
            .Select(group => group
                .OrderByDescending(item => item.IsSelected)
                .ThenByDescending(item => item.RenewalVersionNumber ?? int.MinValue)
                .ThenByDescending(item => item.DateCreated)
                .ThenByDescending(item => item.VersionId)
                .First())
            .OrderBy(item => DateReference(item).SortYear)
            .ThenBy(item => DateReference(item).SortMonth)
            .ThenBy(item => DateReference(item).SortDay)
            .ThenBy(item => item.VersionId)
            .ToList();

        var points = new List<HistoryRenewalEvolutionPoint>();
        string? previousCarrier = null;
        foreach (var version in ordered)
        {
            var carrierChanged = !string.IsNullOrWhiteSpace(previousCarrier) &&
                                 !string.IsNullOrWhiteSpace(version.Carrier) &&
                                 !version.Carrier.Equals(previousCarrier, StringComparison.OrdinalIgnoreCase);
            if (!string.IsNullOrWhiteSpace(version.Carrier))
            {
                previousCarrier = version.Carrier;
            }

            points.Add(new HistoryRenewalEvolutionPoint
            {
                EventId = ClientHistoryProjectionBuilder.StableKey("Renewal", "RenewalVersion", version.VersionId, "version"),
                RenewalId = version.RenewalId,
                DateReference = DateReference(version),
                RequestedIncrease = version.RequestedIncreasePercent,
                NegotiatedIncrease = version.NegotiatedIncreasePercent,
                FinalIncrease = version.FinalAcceptedIncreasePercent,
                IndustryComparison = version.IndustryBenchmarkPercent,
                NegotiatedSavingsAnnual = version.NegotiatedSavingsAnnual,
                IsCarrierChange = carrierChanged,
                IsCurrent = version.IsSelected,
                RenewalStatus = version.Status,
                DetailDescription = version.Status
            });
        }

        return points;
    }

    private static string RenewalIdentityKey(RenewalVersionRecord version) =>
        version.RenewalId.HasValue
            ? $"renewal:{version.RenewalId.Value}"
            : DateKey(version);

    private static string DateKey(RenewalVersionRecord version)
    {
        var date = DateReference(version);
        return date.Precision switch
        {
            HistoryDatePrecision.ExactDate when date.ExactDate.HasValue => $"exact:{date.ExactDate:yyyy-MM-dd}",
            HistoryDatePrecision.YearOnly when date.Year.HasValue => $"year:{date.Year.Value}",
            _ => $"unknown:{version.VersionId}"
        };
    }

    private static HistoryDateReference DateReference(RenewalVersionRecord version) =>
        version.EffectiveDate.HasValue
            ? HistoryDateReference.FromExact(version.EffectiveDate.Value)
            : version.RenewalYear.HasValue
                ? HistoryDateReference.FromYear(version.RenewalYear.Value)
                : new HistoryDateReference();
}

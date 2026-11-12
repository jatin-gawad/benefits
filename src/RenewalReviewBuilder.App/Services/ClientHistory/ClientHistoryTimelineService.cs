using RenewalReviewBuilder.App.Models.History;

namespace RenewalReviewBuilder.App.Services.ClientHistory;

public sealed class ClientHistoryTimelineService : IClientHistoryFilterService
{
    public IReadOnlyList<ClientHistoryEventReadModel> Filter(IEnumerable<ClientHistoryEventReadModel> events, HistorySelectionState selection) =>
        events.Where(historyEvent => !historyEvent.IsExcluded).ToList();

    public IReadOnlyList<ClientHistoryTimelineRow> BuildRows(IEnumerable<ClientHistoryEventReadModel> events, ClientHistoryFilter filter)
    {
        var query = events.Where(historyEvent => !historyEvent.IsExcluded);
        if (!string.IsNullOrWhiteSpace(filter.Search))
        {
            query = query.Where(historyEvent =>
                $"{historyEvent.Title} {historyEvent.Summary} {historyEvent.EventType} {historyEvent.Category}"
                    .Contains(filter.Search, StringComparison.OrdinalIgnoreCase));
        }

        if (!filter.Category.Equals("All Events", StringComparison.OrdinalIgnoreCase) &&
            !filter.Category.Equals("All benefit areas", StringComparison.OrdinalIgnoreCase))
        {
            query = query.Where(historyEvent =>
                historyEvent.Category.Equals(filter.Category, StringComparison.OrdinalIgnoreCase) ||
                historyEvent.EventType.Equals(filter.Category, StringComparison.OrdinalIgnoreCase));
        }

        if (filter.From.HasValue)
        {
            query = query.Where(historyEvent => DateReference(historyEvent).IsOnOrAfter(filter.From.Value));
        }
        if (filter.To.HasValue)
        {
            query = query.Where(historyEvent => DateReference(historyEvent).IsOnOrBefore(filter.To.Value));
        }

        var ordered = filter.OldestFirst
            ? query.OrderBy(historyEvent => DateReference(historyEvent).SortYear)
                .ThenBy(historyEvent => DateReference(historyEvent).SortMonth)
                .ThenBy(historyEvent => DateReference(historyEvent).SortDay)
                .ThenBy(historyEvent => historyEvent.RecordedDate)
            : query.OrderByDescending(historyEvent => DateReference(historyEvent).SortYear)
                .ThenByDescending(historyEvent => DateReference(historyEvent).SortMonth)
                .ThenByDescending(historyEvent => DateReference(historyEvent).SortDay)
                .ThenByDescending(historyEvent => historyEvent.RecordedDate);

        return ordered.Select(historyEvent => new ClientHistoryTimelineRow
        {
            EventId = historyEvent.EventId,
            EffectiveDate = historyEvent.EffectiveDate,
            EffectiveDateDisplay = DateReference(historyEvent).DisplayText,
            IconGlyph = Icon(historyEvent.Category),
            AccentColor = AccentColor(historyEvent.Category),
            AccentBackground = AccentBackground(historyEvent.Category),
            EventType = historyEvent.EventType,
            Title = historyEvent.Title,
            Summary = OneLine(historyEvent.Summary),
            PrimaryImpact = historyEvent.FinancialImpact,
            EvidenceCount = historyEvent.Evidence.Count,
            Confidence = historyEvent.Confidence
        }).ToList();
    }

    public ClientHistoryKpis CalculateKpis(IEnumerable<ClientHistoryEventReadModel> events, DateOnly today)
    {
        var items = events.Where(historyEvent => !historyEvent.IsExcluded).ToList();
        var first = items.Where(historyEvent => historyEvent.EffectiveDate.HasValue)
            .Select(historyEvent => historyEvent.EffectiveDate!.Value)
            .DefaultIfEmpty()
            .Min();
        var approved = items
            .Where(historyEvent => historyEvent.ValidationStatus == HistoryValidationStatus.Valid && historyEvent.FinancialImpact.StartsWith("$", StringComparison.Ordinal))
            .GroupBy(historyEvent => historyEvent.StableSourceKey)
            .Select(group => group.First())
            .ToList();
        var values = approved.Select(historyEvent =>
            decimal.TryParse(
                historyEvent.FinancialImpact.TrimStart('$'),
                System.Globalization.NumberStyles.Currency,
                System.Globalization.CultureInfo.GetCultureInfo("en-CA"),
                out var value)
                    ? value
                    : 0m).ToList();
        var carrier = ResolveCurrentCarrier(items);

        return new ClientHistoryKpis
        {
            YearsAsClient = first == default ? null : today.Year - first.Year,
            TotalValueCreated = values.Count == 0 ? null : values.Sum(),
            BenefitImprovementCount = items.Count(historyEvent => historyEvent.Category.Contains("Plan", StringComparison.OrdinalIgnoreCase)),
            SuccessfulNegotiationCount = items.Count(historyEvent => historyEvent.EventType.Contains("Negotiat", StringComparison.OrdinalIgnoreCase)),
            CurrentCarrier = carrier?.Title ?? "Not available",
            CarrierEffectiveDate = carrier?.EffectiveDate
        };
    }

    private static ClientHistoryEventReadModel? ResolveCurrentCarrier(IReadOnlyList<ClientHistoryEventReadModel> items)
    {
        var canonical = items
            .Where(item => item.SourceModule.Equals("CarrierVendor", StringComparison.OrdinalIgnoreCase) &&
                           item.ValidationStatus == HistoryValidationStatus.Valid &&
                           item.ChangeSet.Changes.Any(change => change.FieldKey.Equals("Carrier", StringComparison.OrdinalIgnoreCase)))
            .OrderBy(item => DateReference(item).SortYear)
            .ThenBy(item => DateReference(item).SortMonth)
            .ThenBy(item => DateReference(item).SortDay)
            .ThenBy(item => item.RecordedDate)
            .ToList();
        var activeByBenefit = new Dictionary<string, CanonicalCarrierState>(StringComparer.OrdinalIgnoreCase);
        foreach (var item in canonical)
        {
            var change = item.ChangeSet.Changes.First(value =>
                value.FieldKey.Equals("Carrier", StringComparison.OrdinalIgnoreCase));
            foreach (var benefitKey in CarrierBenefitKeys(item))
            {
                if (string.IsNullOrWhiteSpace(change.NewValue))
                {
                    if (activeByBenefit.TryGetValue(benefitKey, out var active) &&
                        TerminationMatches(active, item, change))
                    {
                        activeByBenefit.Remove(benefitKey);
                    }
                    continue;
                }

                activeByBenefit[benefitKey] = new CanonicalCarrierState(
                    item,
                    change.NewReferenceId,
                    change.NewValue.Trim());
            }
        }

        if (activeByBenefit.Count > 0)
        {
            var carrierNames = activeByBenefit.Values
                .Select(value => value.Name)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (carrierNames.Count > 1)
            {
                return new ClientHistoryEventReadModel { Title = "Multiple carriers" };
            }

            var carrierName = carrierNames[0];
            var carrierStates = activeByBenefit.Values
                .Where(value => value.Name.Equals(carrierName, StringComparison.OrdinalIgnoreCase))
                .ToList();
            var representative = carrierStates
                .OrderByDescending(value => DateReference(value.Event).SortYear)
                .ThenByDescending(value => DateReference(value.Event).SortMonth)
                .ThenByDescending(value => DateReference(value.Event).SortDay)
                .ThenByDescending(value => value.Event.RecordedDate)
                .First().Event;
            var exactDates = carrierStates
                .Select(value => value.Event.EffectiveDate)
                .Where(value => value.HasValue)
                .Select(value => value!.Value)
                .ToList();
            DateOnly? effectiveDate = exactDates.Count == 0 ? null : exactDates.Min();
            return new ClientHistoryEventReadModel
            {
                EventId = representative.EventId,
                StableSourceKey = representative.StableSourceKey,
                SourceModule = representative.SourceModule,
                SourceEntityType = representative.SourceEntityType,
                SourceEntityId = representative.SourceEntityId,
                EffectiveDate = effectiveDate,
                EffectiveDateReference = effectiveDate.HasValue
                    ? HistoryDateReference.FromExact(effectiveDate.Value)
                    : representative.EffectiveDateReference,
                Title = carrierName
            };
        }

        if (canonical.Count > 0) return null;
        return items
            .Where(item => item.Category.Contains("Carrier", StringComparison.OrdinalIgnoreCase) &&
                           item.ValidationStatus == HistoryValidationStatus.Valid &&
                           !item.EventType.Contains("termination", StringComparison.OrdinalIgnoreCase) &&
                           !item.EventType.Contains("provider", StringComparison.OrdinalIgnoreCase))
            .OrderByDescending(item => item.EventType.Contains("Current carrier", StringComparison.OrdinalIgnoreCase))
            .ThenByDescending(item => item.EffectiveDate)
            .FirstOrDefault();
    }

    private static IReadOnlyList<string> CarrierBenefitKeys(ClientHistoryEventReadModel item)
    {
        var keys = item.RelatedRecords
            .Where(value => value.StartsWith("BenefitType:", StringComparison.OrdinalIgnoreCase))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        return keys.Count == 0 ? ["CarrierScope:default"] : keys;
    }

    private static bool TerminationMatches(
        CanonicalCarrierState active,
        ClientHistoryEventReadModel termination,
        HistoryChangeItem change)
    {
        if (!string.IsNullOrWhiteSpace(change.PreviousReferenceId))
        {
            return active.ReferenceId.Equals(change.PreviousReferenceId, StringComparison.OrdinalIgnoreCase);
        }

        if (!string.IsNullOrWhiteSpace(change.PreviousValue))
        {
            return active.Name.Equals(change.PreviousValue, StringComparison.OrdinalIgnoreCase);
        }

        return active.Event.SourceEntityId.Equals(termination.SourceEntityId, StringComparison.OrdinalIgnoreCase);
    }

    private sealed record CanonicalCarrierState(
        ClientHistoryEventReadModel Event,
        string ReferenceId,
        string Name);

    private static HistoryDateReference DateReference(ClientHistoryEventReadModel historyEvent) =>
        historyEvent.EffectiveDateReference.Precision != HistoryDatePrecision.Unknown
            ? historyEvent.EffectiveDateReference
            : historyEvent.EffectiveDate.HasValue
                ? HistoryDateReference.FromExact(historyEvent.EffectiveDate.Value)
                : new HistoryDateReference();

    private static string OneLine(string value) =>
        string.Join(" ", value.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries));

    private static string AccentColor(string category) =>
        category.Contains("Carrier", StringComparison.OrdinalIgnoreCase) ? "#0B5DE8" :
        category.Contains("Financial", StringComparison.OrdinalIgnoreCase) || category.Contains("Market", StringComparison.OrdinalIgnoreCase) ? "#F47A13" :
        category.Contains("Renewal", StringComparison.OrdinalIgnoreCase) || category.Contains("Administration", StringComparison.OrdinalIgnoreCase) ? "#6D28D9" : "#18A55B";

    private static string AccentBackground(string category) =>
        category.Contains("Carrier", StringComparison.OrdinalIgnoreCase) ? "#EEF4FF" :
        category.Contains("Financial", StringComparison.OrdinalIgnoreCase) || category.Contains("Market", StringComparison.OrdinalIgnoreCase) ? "#FFF1E6" :
        category.Contains("Renewal", StringComparison.OrdinalIgnoreCase) || category.Contains("Administration", StringComparison.OrdinalIgnoreCase) ? "#F3EDFF" : "#EAF8F0";

    private static string Icon(string category) =>
        category.Contains("Plan", StringComparison.OrdinalIgnoreCase) ? "\uE8F1" :
        category.Contains("Carrier", StringComparison.OrdinalIgnoreCase) ? "\uE7F4" :
        category.Contains("Financial", StringComparison.OrdinalIgnoreCase) ? "\uEAFD" :
        "\uE81C";
}

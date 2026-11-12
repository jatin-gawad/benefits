using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Models.History;
using RenewalReviewBuilder.App.Models.Querying;

namespace RenewalReviewBuilder.App.Services.ClientHistory;

public sealed class PlanDesignHistoryContributor : IClientHistoryContributor
{
    private readonly Func<long, CancellationToken, Task<IReadOnlyList<PlanBenefitRule>>> _loader;

    public PlanDesignHistoryContributor(DatabaseService database) : this(async (clientId, cancellationToken) =>
    {
        cancellationToken.ThrowIfCancellationRequested();
        var rules = await database.GetPlanBenefitRulesAsync(clientId);
        cancellationToken.ThrowIfCancellationRequested();
        return rules;
    }) { }

    internal PlanDesignHistoryContributor(
        Func<long, CancellationToken, Task<IReadOnlyList<PlanBenefitRule>>> loader) => _loader = loader;

    public string SourceModule => "PlanDesign";

    public async Task<IReadOnlyList<ClientHistoryEventReadModel>> ContributeAsync(
        long clientId,
        CancellationToken cancellationToken = default) =>
        (await LoadAndBuildAsync(clientId, null, cancellationToken)).Events;

    public async Task<ClientHistoryContribution> ContributeAsync(AsOfQueryContext context)
    {
        var result = await LoadAndBuildAsync(context.ClientId, context.AsOfDate, context.CancellationToken);
        var detail = result.SourceCount == 0
            ? "Plan Design replay is limited because no effective-dated normalized benefit rules are saved."
            : result.RecordedTimeExcludedCount > 0
                ? "Plan Design replay excludes rules saved or updated after the selected date, and rules without a canonical recorded time. Controlled reprocessing does not preserve an immutable rule-version chain."
                : result.MissingFeatureCount > 0
                    ? "Plan Design replay is limited because some rules are not linked to a canonical plan-design feature. Those rules remain standalone facts and are not compared into a fabricated change chain."
                    : result.ExcludedCount > 0
                        ? "Plan Design replay uses effective-dated normalized benefit rules; undated, incomplete, and review-required rules are excluded."
                        : "Plan Design replay uses effective-dated normalized benefit rules. Historical reconstruction remains limited because controlled reprocessing does not preserve an immutable rule-version chain.";
        return new ClientHistoryContribution
        {
            Events = result.Events,
            HistoricalSupport = new HistoricalQuerySupportStatus
            {
                SourceModule = SourceModule,
                Support = HistoricalQuerySupport.Limited,
                Detail = detail
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

        var canonicalRules = source
            .Where(rule => rule.ClientId == clientId &&
                           rule.PlanBenefitRuleId.HasValue &&
                           rule.NormalizationStatus == BenefitRuleNormalizationStatus.Normalized &&
                           rule.IsComplete &&
                           rule.EffectiveDate.HasValue)
            .ToList();
        var recordedTimeExcludedCount = asOfDate.HasValue
            ? canonicalRules.Count(rule => !WasKnownBy(Recorded(rule.UpdatedAtUtc, rule.CreatedAtUtc), asOfDate))
            : 0;
        var rules = canonicalRules
            .Where(rule => WasKnownBy(Recorded(rule.UpdatedAtUtc, rule.CreatedAtUtc), asOfDate))
            .OrderBy(rule => rule.EffectiveDate)
            .ThenBy(rule => rule.PlanDesignFeatureId)
            .ThenBy(rule => rule.DisplayOrder)
            .ThenBy(rule => rule.PlanBenefitRuleId)
            .ToList();

        var events = new List<ClientHistoryEventReadModel>();
        foreach (var group in rules.GroupBy(ComparableKey, StringComparer.Ordinal))
        {
            PlanBenefitRule? previous = null;
            foreach (var current in group)
            {
                if (previous is null)
                {
                    events.Add(CreateEffectiveEvent(clientId, current));
                    previous = current;
                    continue;
                }

                if (HasValidityGap(previous, current))
                {
                    AddTerminationEvent(events, clientId, previous);
                    events.Add(CreateEffectiveEvent(clientId, current));
                }
                else if (!SameCanonicalValues(previous, current))
                {
                    events.Add(CreateRuleEvent(
                        clientId,
                        current,
                        "changed",
                        "Plan design rule changed",
                        "Tiered reimbursement rule changed",
                        RuleSummary(current),
                        BuildChanges(previous, current),
                        current.EffectiveDate!.Value));
                }

                // Always advance the validity cursor, including for a contiguous
                // successor whose canonical values are unchanged. Its evidence and
                // exact termination date are the current facts for this rule chain.
                previous = current;
            }

            AddTerminationEvent(events, clientId, previous);
        }

        return new BuildResult(
            events,
            source.Count,
            source.Count - rules.Count,
            recordedTimeExcludedCount,
            rules.Count(rule => !rule.PlanDesignFeatureId.HasValue));
    }

    private ClientHistoryEventReadModel CreateEffectiveEvent(long clientId, PlanBenefitRule rule) =>
        CreateRuleEvent(
            clientId,
            rule,
            "effective",
            "Plan design rule effective",
            "Tiered reimbursement rule effective",
            RuleSummary(rule),
            BuildInitialValues(rule),
            rule.EffectiveDate!.Value);

    private void AddTerminationEvent(
        ICollection<ClientHistoryEventReadModel> events,
        long clientId,
        PlanBenefitRule? rule)
    {
        if (rule?.TerminationDate is not DateOnly termination ||
            termination < rule.EffectiveDate!.Value)
        {
            return;
        }

        events.Add(CreateRuleEvent(
            clientId,
            rule,
            "terminated",
            "Plan design rule termination",
            "Tiered reimbursement rule ended",
            RuleSummary(rule),
            new HistoryChangeSet
            {
                Changes =
                [
                    new HistoryChangeItem
                    {
                        FieldKey = "RuleStatus",
                        DisplayLabel = "Rule status",
                        ValueType = HistoryValueType.Text,
                        PreviousValue = "Effective",
                        NewValue = "Ended"
                    }
                ]
            },
            termination));
    }

    private static bool HasValidityGap(PlanBenefitRule previous, PlanBenefitRule current) =>
        previous.TerminationDate is DateOnly termination &&
        termination.DayNumber + 1 < current.EffectiveDate!.Value.DayNumber;

    private ClientHistoryEventReadModel CreateRuleEvent(
        long clientId,
        PlanBenefitRule rule,
        string subtype,
        string eventType,
        string title,
        string summary,
        HistoryChangeSet changes,
        DateOnly effectiveDate)
    {
        var ruleId = rule.PlanBenefitRuleId!.Value;
        var related = new List<string> { $"PlanBenefitRule:{ruleId}" };
        if (rule.PlanDesignFeatureId.HasValue) related.Add($"PlanDesignFeature:{rule.PlanDesignFeatureId.Value}");
        if (rule.SourceDocumentId.HasValue) related.Add($"Document:{rule.SourceDocumentId.Value}");
        var stableKey = ClientHistoryProjectionBuilder.StableKey(
            SourceModule,
            "PlanBenefitRule",
            ruleId.ToString(CultureInfo.InvariantCulture),
            subtype);
        var confidence = CanonicalHistoryEventFactory.ClampConfidence(rule.ConfidenceScore);
        return CanonicalHistoryEventFactory.Create(
            SourceModule,
            "PlanBenefitRule",
            ruleId.ToString(CultureInfo.InvariantCulture),
            subtype,
            clientId,
            HistoryDateReference.FromExact(effectiveDate),
            Recorded(rule.UpdatedAtUtc, rule.CreatedAtUtc),
            eventType,
            "Plan Design",
            title,
            summary,
            HistoryValidationStatus.Valid,
            confidence,
            [CanonicalHistoryEventFactory.Evidence(
                ruleId.ToString(CultureInfo.InvariantCulture),
                HistoryValidationStatus.Valid,
                confidence,
                rule.SourceDocumentId,
                section: rule.SourceLocation,
                extractedText: rule.SourceText)],
            related,
            changes,
            stableKey: stableKey);
    }

    private static HistoryChangeSet BuildInitialValues(PlanBenefitRule rule) => new()
    {
        Changes =
        [
            InitialValue("RuleType", "Rule type", HistoryValueType.Text, rule.RuleType.ToString()),
            InitialValue("Tier1ReimbursementPercent", "Tier 1 reimbursement", HistoryValueType.Percentage,
                CanonicalHistoryEventFactory.Number(rule.Tier1ReimbursementPercent), "%"),
            InitialValue("Tier1ThresholdAmount", "Tier 1 threshold", HistoryValueType.Money,
                CanonicalHistoryEventFactory.Number(rule.Tier1ThresholdAmount), rule.CurrencyCode),
            InitialValue("Tier2ReimbursementPercent", "Tier 2 reimbursement", HistoryValueType.Percentage,
                CanonicalHistoryEventFactory.Number(rule.Tier2ReimbursementPercent), "%"),
            InitialValue("ThresholdBasis", "Threshold basis", HistoryValueType.Text, rule.ThresholdBasis?.ToString() ?? ""),
            InitialValue("ResetPeriod", "Reset period", HistoryValueType.Text, rule.ResetPeriod?.ToString() ?? ""),
            InitialValue("CurrencyCode", "Currency", HistoryValueType.Text, rule.CurrencyCode)
        ]
    };

    private static HistoryChangeItem InitialValue(
        string key,
        string label,
        HistoryValueType type,
        string value,
        string unit = "") => new()
    {
        FieldKey = key,
        DisplayLabel = label,
        ValueType = type,
        NewValue = value,
        Unit = unit
    };

    private static HistoryChangeSet BuildChanges(PlanBenefitRule previous, PlanBenefitRule current)
    {
        var changes = new List<HistoryChangeItem>();
        Add(changes, "Tier1ReimbursementPercent", "Tier 1 reimbursement", HistoryValueType.Percentage,
            CanonicalHistoryEventFactory.Number(previous.Tier1ReimbursementPercent), CanonicalHistoryEventFactory.Number(current.Tier1ReimbursementPercent), "%");
        Add(changes, "Tier1ThresholdAmount", "Tier 1 threshold", HistoryValueType.Money,
            CanonicalHistoryEventFactory.Number(previous.Tier1ThresholdAmount), CanonicalHistoryEventFactory.Number(current.Tier1ThresholdAmount), current.CurrencyCode);
        Add(changes, "Tier2ReimbursementPercent", "Tier 2 reimbursement", HistoryValueType.Percentage,
            CanonicalHistoryEventFactory.Number(previous.Tier2ReimbursementPercent), CanonicalHistoryEventFactory.Number(current.Tier2ReimbursementPercent), "%");
        Add(changes, "ThresholdBasis", "Threshold basis", HistoryValueType.Text, previous.ThresholdBasis?.ToString() ?? "", current.ThresholdBasis?.ToString() ?? "");
        Add(changes, "ResetPeriod", "Reset period", HistoryValueType.Text, previous.ResetPeriod?.ToString() ?? "", current.ResetPeriod?.ToString() ?? "");
        Add(changes, "CurrencyCode", "Currency", HistoryValueType.Text, previous.CurrencyCode, current.CurrencyCode);
        return new HistoryChangeSet { Changes = changes };
    }

    private static void Add(
        ICollection<HistoryChangeItem> changes,
        string key,
        string label,
        HistoryValueType type,
        string previous,
        string current,
        string unit = "")
    {
        if (previous.Equals(current, StringComparison.Ordinal)) return;
        changes.Add(new HistoryChangeItem
        {
            FieldKey = key,
            DisplayLabel = label,
            ValueType = type,
            PreviousValue = previous,
            NewValue = current,
            Unit = unit
        });
    }

    private static bool SameCanonicalValues(PlanBenefitRule left, PlanBenefitRule right) =>
        left.Tier1ReimbursementPercent == right.Tier1ReimbursementPercent &&
        left.Tier1ThresholdAmount == right.Tier1ThresholdAmount &&
        left.Tier2ReimbursementPercent == right.Tier2ReimbursementPercent &&
        left.ThresholdBasis == right.ThresholdBasis &&
        left.ResetPeriod == right.ResetPeriod &&
        left.CurrencyCode.Equals(right.CurrencyCode, StringComparison.OrdinalIgnoreCase);

    private static string ComparableKey(PlanBenefitRule rule) =>
        $"{rule.PlanDesignFeatureId?.ToString(CultureInfo.InvariantCulture) ?? $"rule-{rule.PlanBenefitRuleId}"}|{rule.RuleType}|{rule.DisplayOrder}";

    private static string RuleSummary(PlanBenefitRule rule) =>
        $"{CanonicalHistoryEventFactory.Number(rule.Tier1ReimbursementPercent)}% to " +
        $"{CanonicalHistoryEventFactory.Number(rule.Tier1ThresholdAmount)} {rule.CurrencyCode}; then " +
        $"{CanonicalHistoryEventFactory.Number(rule.Tier2ReimbursementPercent)}%; " +
        $"{rule.ThresholdBasis}; {rule.ResetPeriod}.";

    private static DateTime? Recorded(params DateTime[] values)
    {
        var value = values.Where(item => item != default).DefaultIfEmpty().Max();
        return value == default ? null : value;
    }

    private static bool WasKnownBy(DateTime? recordedAt, DateOnly? asOfDate) =>
        !asOfDate.HasValue ||
        recordedAt.HasValue && DateOnly.FromDateTime(recordedAt.Value) <= asOfDate.Value;

    private sealed record BuildResult(
        IReadOnlyList<ClientHistoryEventReadModel> Events,
        int SourceCount,
        int ExcludedCount,
        int RecordedTimeExcludedCount,
        int MissingFeatureCount);
}

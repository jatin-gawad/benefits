using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Models.History;
using RenewalReviewBuilder.App.Models.Querying;

namespace RenewalReviewBuilder.App.Services.ClientHistory;

internal sealed record FinancialHistorySourceData(
    IReadOnlyList<FinancialArrangementHistoryVersion> Versions,
    IReadOnlyList<FinancialArrangementAuditEvent> AuditEvents,
    IReadOnlyList<AsoFinancialActivityStatement> Statements);

public sealed class FinancialHistoryContributor : IClientHistoryContributor
{
    private readonly Func<long, CancellationToken, Task<FinancialHistorySourceData>> _loader;

    public FinancialHistoryContributor() : this((_, _) => Task.FromResult(new FinancialHistorySourceData([], [], []))) { }

    public FinancialHistoryContributor(DatabaseService database) : this(async (clientId, cancellationToken) =>
    {
        cancellationToken.ThrowIfCancellationRequested();
        var service = new FinancialArrangementService(database.DatabasePath);
        var versionsTask = service.GetHistoryVersionsForClientAsync(clientId);
        var auditsTask = service.GetAuditEventsAsync(clientId);
        var statementsTask = database.GetAsoFinancialActivityHistoryAsync(clientId);
        await Task.WhenAll(versionsTask, auditsTask, statementsTask);
        cancellationToken.ThrowIfCancellationRequested();
        return new FinancialHistorySourceData(
            await versionsTask,
            await auditsTask,
            await statementsTask);
    }) { }

    internal FinancialHistoryContributor(Func<long, CancellationToken, Task<FinancialHistorySourceData>> loader) => _loader = loader;

    public string SourceModule => "Financial";

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
                    ? "Financial history has limited replay support; no confirmed arrangement versions or complete ASO statements are saved."
                    : "Financial replay includes confirmed effective-dated arrangement versions and complete ASO statements recorded by the replay date. Validity-end chains, review-required rows, and undated legacy facts are not reconstructed."
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
        var events = new List<ClientHistoryEventReadModel>();
        events.AddRange(BuildVersionEvents(clientId, source.Versions, source.AuditEvents, asOfDate));
        events.AddRange(BuildAsoEvents(clientId, source.Statements, asOfDate));
        return new BuildResult(events, source.Versions.Count + source.Statements.Count);
    }

    private IEnumerable<ClientHistoryEventReadModel> BuildVersionEvents(
        long clientId,
        IReadOnlyList<FinancialArrangementHistoryVersion> source,
        IReadOnlyList<FinancialArrangementAuditEvent> audits,
        DateOnly? asOfDate)
    {
        var confirmed = source
            .Where(version => version.ClientId == clientId &&
                              version.ArrangementStatus.Equals("Confirmed", StringComparison.OrdinalIgnoreCase) &&
                              IsKnownBy(version, audits, asOfDate))
            .OrderBy(version => version.EffectiveDate)
            .ThenBy(version => version.FinancialArrangementVersionId)
            .ToList();

        foreach (var group in confirmed.GroupBy(version => $"{version.Carrier.Trim().ToUpperInvariant()}|{version.Scenario}"))
        {
            FinancialArrangementHistoryVersion? previous = null;
            foreach (var version in group)
            {
                var audit = ConfirmationAudit(version, audits);
                var confidence = CanonicalHistoryEventFactory.ClampConfidence(version.TargetLossRatioConfidence);
                var poolingConfirmed = IsPoolingConfirmed(version, audits, asOfDate);
                var previousPoolingConfirmed = previous is not null && IsPoolingConfirmed(previous, audits, asOfDate);
                var changes = VersionChanges(previous, version, previousPoolingConfirmed, poolingConfirmed);
                var evidence = VersionEvidence(version, poolingConfirmed, confidence);
                var related = new List<string>
                {
                    $"FinancialArrangement:{version.FinancialArrangementId}",
                    $"FinancialArrangementVersion:{version.FinancialArrangementVersionId}"
                };
                if (version.SourceDocumentId.HasValue) related.Add($"Document:{version.SourceDocumentId.Value}");
                if (IsCommissionConfirmed(version) && version.CommissionArrangementId.HasValue)
                {
                    related.Add($"CommissionArrangement:{version.CommissionArrangementId.Value}");
                    if (version.CommissionSourceDocumentId.HasValue) related.Add($"Document:{version.CommissionSourceDocumentId.Value}");
                }
                if (poolingConfirmed && version.PoolingArrangementId.HasValue)
                {
                    related.Add($"PoolingArrangement:{version.PoolingArrangementId.Value}");
                    if (version.PoolingSourceDocumentId.HasValue) related.Add($"Document:{version.PoolingSourceDocumentId.Value}");
                }
                foreach (var charge in version.Charges)
                {
                    related.Add($"FinancialArrangementCharge:{charge.FinancialArrangementChargeId}");
                }

                yield return CanonicalHistoryEventFactory.Create(
                    SourceModule,
                    "FinancialArrangementVersion",
                    version.FinancialArrangementVersionId.ToString(CultureInfo.InvariantCulture),
                    "version",
                    clientId,
                    HistoryDateReference.FromExact(version.EffectiveDate),
                    KnownAt(version, audit),
                    "Financial arrangement version",
                    "Financial",
                    $"{version.Carrier} financial arrangement confirmed".Trim(),
                    $"{version.Scenario} arrangement effective {version.EffectiveDate:MMM d, yyyy}.",
                    HistoryValidationStatus.Valid,
                    confidence,
                    evidence,
                    related.Distinct(StringComparer.Ordinal).ToList(),
                    changes,
                    audit?.Reason ?? "",
                    audit?.User ?? "");
                previous = version;
            }
        }
    }

    private IEnumerable<ClientHistoryEventReadModel> BuildAsoEvents(
        long clientId,
        IReadOnlyList<AsoFinancialActivityStatement> source,
        DateOnly? asOfDate)
    {
        var current = source
            .Where(statement => statement.ClientId == clientId &&
                                statement.ImportStatus.Equals(AsoFinancialImportStatuses.Complete, StringComparison.OrdinalIgnoreCase))
            .Where(statement => asOfDate.HasValue
                ? statement.ImportedAt != default && DateOnly.FromDateTime(statement.ImportedAt) <= asOfDate.Value
                : statement.IsCurrent)
            .GroupBy(StatementNaturalKey, StringComparer.Ordinal)
            .Select(group => group.OrderByDescending(statement => statement.VersionNumber)
                .ThenByDescending(statement => statement.ImportedAt)
                .First())
            .OrderBy(statement => StatementDate(statement).SortYear)
            .ThenBy(statement => StatementDate(statement).SortMonth)
            .ThenBy(statement => StatementDate(statement).SortDay)
            .ThenBy(statement => statement.StatementId)
            .ToList();

        foreach (var statement in current)
        {
            var date = StatementDate(statement);
            if (date.Precision == HistoryDatePrecision.Unknown) continue;
            var statementIdentity = StatementNaturalKey(statement);
            var confidence = .95m;
            var evidence = StatementEvidence(statement, HistoryValidationStatus.Valid, confidence);
            var related = new List<string> { $"AsoFinancialStatement:{statement.StatementId}" };
            if (statement.SourceDocumentId.HasValue) related.Add($"Document:{statement.SourceDocumentId.Value}");

            yield return CanonicalHistoryEventFactory.Create(
                SourceModule,
                "AsoFinancialStatement",
                statement.StatementId.ToString(CultureInfo.InvariantCulture),
                "statement",
                clientId,
                date,
                statement.ImportedAt == default ? null : statement.ImportedAt,
                "ASO financial statement",
                "Financial",
                "ASO financial statement recorded",
                StatementSummary(statement),
                HistoryValidationStatus.Valid,
                confidence,
                evidence,
                related,
                stableKey: ClientHistoryProjectionBuilder.StableKey(
                    SourceModule,
                    "AsoFinancialStatement",
                    statementIdentity,
                    "statement"));

            var closing = statement.Total ?? statement.Periods
                .Where(period => !period.IsTotal)
                .OrderBy(period => period.ActivityMonth)
                .ThenBy(period => period.DisplayOrder)
                .LastOrDefault();
            if (closing is null) continue;

            if (closing.CurrentBalance.HasValue && !string.IsNullOrWhiteSpace(closing.BalanceStatus))
            {
                yield return CanonicalHistoryEventFactory.Create(
                    SourceModule,
                    "AsoFinancialPeriod",
                    closing.PeriodId.ToString(CultureInfo.InvariantCulture),
                    "funding-position",
                    clientId,
                    date,
                    statement.ImportedAt == default ? null : statement.ImportedAt,
                    "ASO funding position",
                    "Financial",
                    $"ASO funding status: {closing.BalanceStatus.Trim()}",
                    "The source statement records a closing funding position.",
                    HistoryValidationStatus.Valid,
                    confidence,
                    evidence,
                    related,
                    new HistoryChangeSet
                    {
                        Changes =
                        [
                            new HistoryChangeItem
                            {
                                FieldKey = "BalanceStatus",
                                DisplayLabel = "Source balance status",
                                ValueType = HistoryValueType.Text,
                                NewValue = closing.BalanceStatus.Trim()
                            },
                            new HistoryChangeItem
                            {
                                FieldKey = "CurrentBalance",
                                DisplayLabel = "Current balance",
                                ValueType = HistoryValueType.Money,
                                NewValue = CanonicalHistoryEventFactory.Number(closing.CurrentBalance),
                                Unit = "CAD"
                            }
                        ]
                    },
                    stableKey: ClientHistoryProjectionBuilder.StableKey(
                        SourceModule,
                        "AsoFinancialStatement",
                        statementIdentity,
                        "funding-position"));
            }

            if (closing.DeficitPaymentRefund is decimal refund && refund != 0m)
            {
                yield return CanonicalHistoryEventFactory.Create(
                    SourceModule,
                    "AsoFinancialPeriod",
                    closing.PeriodId.ToString(CultureInfo.InvariantCulture),
                    "deficit-payment-refund",
                    clientId,
                    date,
                    statement.ImportedAt == default ? null : statement.ImportedAt,
                    "ASO deficit payment / refund",
                    "Financial",
                    "ASO deficit payment / refund recorded",
                    "The source statement contains an explicit deficit payment / refund amount; direction is not inferred.",
                    HistoryValidationStatus.Valid,
                    confidence,
                    evidence,
                    related,
                    new HistoryChangeSet
                    {
                        Changes =
                        [
                            new HistoryChangeItem
                            {
                                FieldKey = "DeficitPaymentRefund",
                                DisplayLabel = "Deficit payment / refund",
                                ValueType = HistoryValueType.Money,
                                NewValue = CanonicalHistoryEventFactory.Number(refund),
                                Unit = "CAD"
                            }
                        ]
                    },
                    stableKey: ClientHistoryProjectionBuilder.StableKey(
                        SourceModule,
                        "AsoFinancialStatement",
                        statementIdentity,
                        "deficit-payment-refund"));
            }
        }
    }

    private static HistoryChangeSet VersionChanges(
        FinancialArrangementHistoryVersion? previous,
        FinancialArrangementHistoryVersion current,
        bool previousPoolingConfirmed,
        bool currentPoolingConfirmed)
    {
        var currentTerms = VersionTerms(current, currentPoolingConfirmed)
            .ToDictionary(term => term.FieldKey, StringComparer.Ordinal);
        if (previous is null)
        {
            return new HistoryChangeSet
            {
                Changes = currentTerms.Values.Select(term => Change(term, "", term.Value)).ToList()
            };
        }

        var previousTerms = VersionTerms(previous, previousPoolingConfirmed)
            .ToDictionary(term => term.FieldKey, StringComparer.Ordinal);
        var changes = new List<HistoryChangeItem>();
        foreach (var term in currentTerms.Values)
        {
            if (previousTerms.TryGetValue(term.FieldKey, out var prior) &&
                prior.Value.Equals(term.Value, StringComparison.Ordinal)) continue;
            changes.Add(Change(term, prior?.Value ?? "", term.Value));
        }
        foreach (var prior in previousTerms.Values)
        {
            if (currentTerms.ContainsKey(prior.FieldKey) || !IsCanonicalRemoval(prior.FieldKey, current)) continue;
            changes.Add(Change(prior, prior.Value, ""));
        }

        return new HistoryChangeSet { Changes = changes };
    }

    private static HistoryChangeItem Change(CanonicalFinancialTerm term, string previousValue, string newValue) => new()
    {
        FieldKey = term.FieldKey,
        DisplayLabel = term.DisplayLabel,
        ValueType = term.ValueType,
        PreviousValue = previousValue,
        NewValue = newValue,
        Unit = term.Unit
    };

    private static bool IsCanonicalRemoval(string fieldKey, FinancialArrangementHistoryVersion current) =>
        fieldKey.StartsWith("Charge:", StringComparison.Ordinal) ||
        (fieldKey.StartsWith("Commission", StringComparison.Ordinal) && !current.CommissionArrangementId.HasValue) ||
        (fieldKey.StartsWith("Pooling", StringComparison.Ordinal) && !current.PoolingArrangementId.HasValue);

    private static IReadOnlyList<CanonicalFinancialTerm> VersionTerms(
        FinancialArrangementHistoryVersion version,
        bool poolingConfirmed)
    {
        var terms = new List<CanonicalFinancialTerm>();
        AddNumber(terms, "TargetLossRatio", "Target loss ratio", HistoryValueType.Percentage, version.TargetLossRatio, "%");
        if (version.TargetLossRatio.HasValue)
        {
            AddText(terms, "TargetLossRatioBasis", "Target loss ratio basis", version.TargetLossRatioBasis);
        }

        if (IsCommissionConfirmed(version))
        {
            AddNumber(terms, "CommissionRate", "Commission rate", HistoryValueType.Percentage, version.CommissionRate, "%");
            AddNumber(terms, "CommissionAmount", "Commission amount", HistoryValueType.Money, version.CommissionAmount, "CAD");
            if (version.CommissionRate.HasValue || version.CommissionAmount.HasValue)
            {
                AddText(terms, "CommissionBasis", "Commission basis", version.CommissionBasis);
            }
        }

        if (poolingConfirmed)
        {
            AddNumber(terms, "PoolingRate", "Pooling rate", HistoryValueType.Percentage, version.PoolingRate, "%");
            AddNumber(terms, "PoolingThreshold", "Pooling threshold", HistoryValueType.Money, version.PoolingThreshold, "CAD");
            if (version.PoolingRate.HasValue || version.PoolingThreshold.HasValue)
            {
                AddText(terms, "PoolingBasis", "Pooling basis", version.PoolingBasis);
                AddText(terms, "PoolingBenefit", "Pooling benefit", version.PoolingBenefit);
                AddText(terms, "PoolingType", "Pooling type", version.PoolingType);
            }
        }

        var chargeKeyCounts = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (var charge in version.Charges
                     .OrderBy(item => item.ChargeType)
                     .ThenBy(item => item.Benefit, StringComparer.OrdinalIgnoreCase)
                     .ThenBy(item => item.FinancialArrangementChargeId))
        {
            if (!charge.Rate.HasValue && !charge.Amount.HasValue) continue;
            var baseKey = $"Charge:{charge.ChargeType}:{KeyToken(charge.Benefit)}";
            var ordinal = chargeKeyCounts.TryGetValue(baseKey, out var existing) ? existing + 1 : 1;
            chargeKeyCounts[baseKey] = ordinal;
            var key = $"{baseKey}:{ordinal}";
            var label = string.IsNullOrWhiteSpace(charge.Benefit)
                ? charge.ChargeType.ToString()
                : $"{charge.ChargeType} ({charge.Benefit.Trim()})";
            AddNumber(terms, key + ":Rate", label + " rate", HistoryValueType.Percentage, charge.Rate, "%");
            AddNumber(terms, key + ":Amount", label + " amount", HistoryValueType.Money, charge.Amount, "CAD");
            AddText(terms, key + ":Basis", label + " basis", charge.Basis);
        }

        return terms;
    }

    private static IReadOnlyList<HistoryEvidenceReference> VersionEvidence(
        FinancialArrangementHistoryVersion version,
        bool poolingConfirmed,
        decimal confidence)
    {
        var evidence = new List<HistoryEvidenceReference>();
        if (version.SourceDocumentId.HasValue ||
            version.TargetLossRatioSourcePage.HasValue ||
            !string.IsNullOrWhiteSpace(version.TargetLossRatioSourceText))
        {
            evidence.Add(CanonicalHistoryEventFactory.Evidence(
                version.FinancialArrangementVersionId.ToString(CultureInfo.InvariantCulture),
                HistoryValidationStatus.Valid,
                confidence,
                version.SourceDocumentId,
                page: version.TargetLossRatioSourcePage,
                section: version.TargetLossRatioBasis,
                extractedText: version.TargetLossRatioSourceText));
        }

        if (IsCommissionConfirmed(version) &&
            version.CommissionArrangementId.HasValue &&
            version.CommissionSourceDocumentId.HasValue)
        {
            evidence.Add(CanonicalHistoryEventFactory.Evidence(
                version.CommissionArrangementId.Value.ToString(CultureInfo.InvariantCulture),
                HistoryValidationStatus.Valid,
                .90m,
                version.CommissionSourceDocumentId,
                section: string.IsNullOrWhiteSpace(version.CommissionSource)
                    ? "Confirmed commission"
                    : version.CommissionSource));
        }

        if (poolingConfirmed &&
            version.PoolingArrangementId.HasValue &&
            (version.PoolingSourceDocumentId.HasValue ||
             version.PoolingSourcePage.HasValue ||
             !string.IsNullOrWhiteSpace(version.PoolingSourceText)))
        {
            evidence.Add(CanonicalHistoryEventFactory.Evidence(
                version.PoolingArrangementId.Value.ToString(CultureInfo.InvariantCulture),
                HistoryValidationStatus.Valid,
                CanonicalHistoryEventFactory.ClampConfidence(version.PoolingConfidence),
                version.PoolingSourceDocumentId,
                page: version.PoolingSourcePage,
                section: version.PoolingBasis,
                extractedText: version.PoolingSourceText));
        }

        foreach (var charge in version.Charges.Where(item => item.Rate.HasValue || item.Amount.HasValue))
        {
            if (!version.SourceDocumentId.HasValue &&
                !charge.SourcePage.HasValue &&
                string.IsNullOrWhiteSpace(charge.SourceText)) continue;
            evidence.Add(CanonicalHistoryEventFactory.Evidence(
                charge.FinancialArrangementChargeId.ToString(CultureInfo.InvariantCulture),
                HistoryValidationStatus.Valid,
                CanonicalHistoryEventFactory.ClampConfidence(charge.ExtractionConfidence),
                version.SourceDocumentId,
                page: charge.SourcePage,
                section: $"{charge.ChargeType} · {charge.Benefit}".Trim(' ', '·'),
                extractedText: charge.SourceText));
        }

        return evidence;
    }

    private static bool IsCommissionConfirmed(FinancialArrangementHistoryVersion version) =>
        version.CommissionArrangementId.HasValue &&
        version.CommissionConfirmationStatus.Equals("Confirmed", StringComparison.OrdinalIgnoreCase);

    private static bool IsPoolingConfirmed(
        FinancialArrangementHistoryVersion version,
        IReadOnlyList<FinancialArrangementAuditEvent> audits,
        DateOnly? asOfDate) =>
        version.PoolingArrangementId.HasValue && audits.Any(item =>
            item.ClientId == version.ClientId &&
            item.EntityType.Equals("PoolingArrangement", StringComparison.OrdinalIgnoreCase) &&
            item.EntityId == version.PoolingArrangementId.Value &&
            item.EventType.Contains("Confirmed", StringComparison.OrdinalIgnoreCase) &&
            (!asOfDate.HasValue ||
             (item.OccurredAtUtc != default &&
              DateOnly.FromDateTime(item.OccurredAtUtc) <= asOfDate.Value)));

    private static FinancialArrangementAuditEvent? ConfirmationAudit(
        FinancialArrangementHistoryVersion version,
        IReadOnlyList<FinancialArrangementAuditEvent> audits) => audits
        .Where(item => item.ClientId == version.ClientId &&
                       item.EntityType.Equals("FinancialArrangement", StringComparison.OrdinalIgnoreCase) &&
                       item.EntityId == version.FinancialArrangementId &&
                       item.EventType.Contains("Confirmed", StringComparison.OrdinalIgnoreCase))
        .OrderByDescending(item => item.OccurredAtUtc)
        .FirstOrDefault();

    private static bool IsKnownBy(
        FinancialArrangementHistoryVersion version,
        IReadOnlyList<FinancialArrangementAuditEvent> audits,
        DateOnly? asOfDate)
    {
        if (!asOfDate.HasValue) return true;
        var knownAt = KnownAt(version, ConfirmationAudit(version, audits));
        return knownAt.HasValue && DateOnly.FromDateTime(knownAt.Value) <= asOfDate.Value;
    }

    private static DateTime? KnownAt(
        FinancialArrangementHistoryVersion version,
        FinancialArrangementAuditEvent? confirmationAudit)
    {
        DateTime? createdAt = version.CreatedAtUtc == default ? null : version.CreatedAtUtc;
        DateTime? confirmedAt = confirmationAudit?.OccurredAtUtc == default
            ? null
            : confirmationAudit?.OccurredAtUtc;
        if (!createdAt.HasValue) return confirmedAt;
        if (!confirmedAt.HasValue) return createdAt;
        return createdAt.Value >= confirmedAt.Value ? createdAt : confirmedAt;
    }

    private static void AddNumber(
        ICollection<CanonicalFinancialTerm> terms,
        string fieldKey,
        string displayLabel,
        HistoryValueType valueType,
        decimal? value,
        string unit)
    {
        if (!value.HasValue) return;
        var displayValue = valueType == HistoryValueType.Percentage
            ? value.Value * 100m
            : value.Value;
        terms.Add(new CanonicalFinancialTerm(
            fieldKey,
            displayLabel,
            valueType,
            CanonicalHistoryEventFactory.Number(displayValue),
            unit));
    }

    private static void AddText(
        ICollection<CanonicalFinancialTerm> terms,
        string fieldKey,
        string displayLabel,
        string value)
    {
        if (string.IsNullOrWhiteSpace(value)) return;
        terms.Add(new CanonicalFinancialTerm(fieldKey, displayLabel, HistoryValueType.Text, value.Trim(), ""));
    }

    private static string KeyToken(string value)
    {
        var token = string.Concat(value.Where(char.IsLetterOrDigit));
        return string.IsNullOrWhiteSpace(token) ? "All" : token;
    }

    private static IReadOnlyList<HistoryEvidenceReference> StatementEvidence(
        AsoFinancialActivityStatement statement,
        HistoryValidationStatus validationStatus,
        decimal confidence) =>
        statement.SourceDocumentId.HasValue ||
        !string.IsNullOrWhiteSpace(statement.SourceFileName) ||
        !string.IsNullOrWhiteSpace(statement.SourceSheet)
            ?
            [
                CanonicalHistoryEventFactory.Evidence(
                    statement.StatementId.ToString(CultureInfo.InvariantCulture),
                    validationStatus,
                    confidence,
                    statement.SourceDocumentId,
                    sourceFactId: statement.SourceHash,
                    page: statement.SourcePageStart,
                    sheet: statement.SourceSheet,
                    section: statement.SourceFileName,
                    extractedText: statement.ValidationSummary)
            ]
            : [];

    private static HistoryDateReference StatementDate(AsoFinancialActivityStatement statement)
    {
        if (statement.StatementDate.HasValue) return HistoryDateReference.FromExact(statement.StatementDate.Value);
        if (statement.PeriodEnd.HasValue) return HistoryDateReference.FromExact(statement.PeriodEnd.Value);
        var month = statement.Periods
            .Where(period => !period.IsTotal && period.ActivityMonth.HasValue)
            .Select(period => period.ActivityMonth!.Value)
            .OrderBy(value => value)
            .LastOrDefault();
        return month == default ? new HistoryDateReference() : HistoryDateReference.FromMonth(month.Year, month.Month);
    }

    private static string StatementNaturalKey(AsoFinancialActivityStatement statement)
    {
        var date = StatementDate(statement);
        var dateKey = date.Precision switch
        {
            HistoryDatePrecision.ExactDate => date.ExactDate?.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) ?? "",
            HistoryDatePrecision.MonthAndYear => $"{date.Year:0000}-{date.Month:00}",
            HistoryDatePrecision.YearOnly => $"{date.Year:0000}",
            _ => ""
        };
        var periodStart = statement.PeriodStart?.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) ?? "";
        var periodEnd = statement.PeriodEnd?.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) ?? "";
        var fallbackDateKey = string.IsNullOrWhiteSpace(periodStart) && string.IsNullOrWhiteSpace(periodEnd)
            ? dateKey
            : "";

        return string.Join("|",
            statement.ContractNumber.Trim().ToUpperInvariant(),
            periodStart,
            periodEnd,
            fallbackDateKey);
    }

    private static string StatementSummary(AsoFinancialActivityStatement statement)
    {
        var parts = new List<string>();
        if (!string.IsNullOrWhiteSpace(statement.Carrier)) parts.Add(statement.Carrier.Trim());
        if (!string.IsNullOrWhiteSpace(statement.ContractNumber)) parts.Add($"contract {statement.ContractNumber.Trim()}");
        if (statement.PeriodStart.HasValue && statement.PeriodEnd.HasValue)
        {
            parts.Add($"period {statement.PeriodStart:MMM d, yyyy} to {statement.PeriodEnd:MMM d, yyyy}");
        }
        return parts.Count == 0 ? "Saved canonical ASO statement period." : string.Join(" · ", parts) + ".";
    }

    private sealed record CanonicalFinancialTerm(
        string FieldKey,
        string DisplayLabel,
        HistoryValueType ValueType,
        string Value,
        string Unit);

    private sealed record BuildResult(IReadOnlyList<ClientHistoryEventReadModel> Events, int SourceCount);
}

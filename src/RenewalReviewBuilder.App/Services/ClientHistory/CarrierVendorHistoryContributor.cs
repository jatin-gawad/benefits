using System.Threading;
using System.Threading.Tasks;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Models.History;
using RenewalReviewBuilder.App.Models.Querying;
using RenewalReviewBuilder.App.Repositories;

namespace RenewalReviewBuilder.App.Services.ClientHistory;

public sealed class CarrierVendorHistoryContributor : IClientHistoryContributor
{
    private readonly Func<long, CancellationToken, Task<IReadOnlyList<ClientBenefitAdministrationRecord>>> _loader;

    public CarrierVendorHistoryContributor() : this((_, _) => Task.FromResult<IReadOnlyList<ClientBenefitAdministrationRecord>>([])) { }

    public CarrierVendorHistoryContributor(DatabaseService database) : this(async (clientId, cancellationToken) =>
    {
        cancellationToken.ThrowIfCancellationRequested();
        var records = await new BenefitAdministrationRepository(database.DatabasePath).GetByClientIdAsync(clientId);
        cancellationToken.ThrowIfCancellationRequested();
        return records;
    }) { }

    internal CarrierVendorHistoryContributor(
        Func<long, CancellationToken, Task<IReadOnlyList<ClientBenefitAdministrationRecord>>> loader) => _loader = loader;

    public string SourceModule => "CarrierVendor";

    public async Task<IReadOnlyList<ClientHistoryEventReadModel>> ContributeAsync(
        long clientId,
        CancellationToken cancellationToken = default) =>
        (await LoadAndBuildAsync(clientId, null, cancellationToken)).Events;

    public async Task<ClientHistoryContribution> ContributeAsync(AsOfQueryContext context)
    {
        var result = await LoadAndBuildAsync(context.ClientId, context.AsOfDate, context.CancellationToken);
        var support = result.SourceCount == 0
            ? HistoricalQuerySupport.Unavailable
            : context.AsOfDate.HasValue || result.HasIncompleteHistory
                ? HistoricalQuerySupport.Limited
                : HistoricalQuerySupport.Supported;
        var detail = result.RecordedTimeExcludedCount > 0
            ? "Carrier and service-provider replay excludes relationships saved or updated after the selected date, and relationships without a canonical recorded time. Benefit-administration rows do not preserve immutable prior states."
            : support switch
        {
            HistoricalQuerySupport.Unavailable => "Carrier and service-provider history is unavailable because no benefit-administration records are saved.",
            HistoricalQuerySupport.Limited when context.AsOfDate.HasValue => "Carrier and service-provider replay includes confirmed effective-dated assignments known by the selected date. Reconstruction remains limited because benefit-administration rows do not preserve immutable prior states.",
            HistoricalQuerySupport.Limited => "Carrier and service-provider history includes confirmed effective-dated assignments; undated, rejected, or review-required relationships are omitted or flagged.",
            _ => "Carrier and service-provider replay uses confirmed effective and termination dates."
        };

        return new ClientHistoryContribution
        {
            Events = result.Events,
            HistoricalSupport = new HistoricalQuerySupportStatus
            {
                SourceModule = SourceModule,
                Support = support,
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
        var records = await _loader(clientId, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();

        var snapshots = new List<RelationshipSnapshot>();
        var incomplete = false;
        var recordedTimeExcludedCount = 0;
        foreach (var record in records)
        {
            if (record.ClientId != clientId)
            {
                incomplete = true;
                continue;
            }

            var recordKnownAt = Recorded(record.UpdatedAtUtc, record.CreatedAtUtc);
            if (!WasKnownBy(recordKnownAt, asOfDate))
            {
                incomplete = true;
                recordedTimeExcludedCount++;
                continue;
            }

            if (IsRejected(record.ReviewStatusCode))
            {
                incomplete = true;
                continue;
            }

            var parentStatus = ValidationStatus(record.ReviewStatusCode);
            if (record.InsurerOrganizationId is > 0 &&
                !string.IsNullOrWhiteSpace(record.InsurerName) &&
                IsInsurerType(record.InsurerOrganizationTypeCode))
            {
                snapshots.Add(new RelationshipSnapshot(
                    GroupKey: $"carrier|{record.BenefitTypeId}",
                    EntityType: "ClientBenefitAdministration",
                    EntityId: record.ClientBenefitAdministrationId.ToString(),
                    IdentityId: record.InsurerOrganizationId.Value,
                    IdentityName: record.InsurerName.Trim(),
                    RelationshipKind: RelationshipKind.Carrier,
                    Role: "Insurer",
                    Benefit: BenefitLabel(record),
                    BenefitTypeId: record.BenefitTypeId,
                    EffectiveDate: ToDate(record.EffectiveDate),
                    TerminationDate: ToDate(record.TerminationDate),
                    RecordedAt: recordKnownAt,
                    ValidationStatus: parentStatus,
                    Confidence: CanonicalHistoryEventFactory.ClampConfidence(record.ExtractionConfidence),
                    DocumentId: record.SourceDocumentId,
                    Page: record.SourcePageNumber,
                    SourceText: record.SourceText,
                    RelatedRecords: Related(record.PlanDesignVersionId, record.BenefitTypeId, "Carrier", record.InsurerOrganizationId.Value, record.ClientBenefitAdministrationId)));
            }
            else if (!string.IsNullOrWhiteSpace(record.InsurerName) || record.InsurerOrganizationId.HasValue)
            {
                incomplete = true;
            }

            foreach (var provider in record.ServiceProviders)
            {
                var providerKnownAt = Recorded(
                    provider.UpdatedAtUtc,
                    provider.CreatedAtUtc);
                if (!WasKnownBy(providerKnownAt, asOfDate))
                {
                    incomplete = true;
                    recordedTimeExcludedCount++;
                    continue;
                }

                if (IsRejected(provider.ReviewStatusCode))
                {
                    incomplete = true;
                    continue;
                }

                if (provider.ProviderOrganizationId <= 0 || string.IsNullOrWhiteSpace(provider.ProviderName))
                {
                    incomplete = true;
                    continue;
                }

                // Provider facts have their own review state. A confirmed parent cannot
                // promote an unconfirmed attached provider into canonical history.
                var providerStatus = parentStatus == HistoryValidationStatus.Valid &&
                                     ValidationStatus(provider.ReviewStatusCode) == HistoryValidationStatus.Valid
                    ? HistoryValidationStatus.Valid
                    : HistoryValidationStatus.NeedsReview;
                snapshots.Add(new RelationshipSnapshot(
                    GroupKey: $"provider|{record.BenefitTypeId}|{provider.ProviderRoleCode.Trim().ToUpperInvariant()}",
                    EntityType: "BenefitServiceProviderAssignment",
                    EntityId: provider.BenefitServiceProviderAssignmentId.ToString(),
                    IdentityId: provider.ProviderOrganizationId,
                    IdentityName: provider.ProviderName.Trim(),
                    RelationshipKind: RelationshipKind.Provider,
                    Role: RoleLabel(provider.ProviderRoleCode),
                    Benefit: BenefitLabel(record),
                    BenefitTypeId: record.BenefitTypeId,
                    EffectiveDate: ToDate(provider.EffectiveDate),
                    TerminationDate: ToDate(provider.TerminationDate),
                    RecordedAt: providerKnownAt,
                    ValidationStatus: providerStatus,
                    Confidence: CanonicalHistoryEventFactory.ClampConfidence(record.ExtractionConfidence),
                    DocumentId: provider.SourceDocumentId ?? record.SourceDocumentId,
                    Page: provider.SourcePageNumber ?? record.SourcePageNumber,
                    SourceText: record.SourceText,
                    RelatedRecords: Related(record.PlanDesignVersionId, record.BenefitTypeId, "Organization", provider.ProviderOrganizationId, record.ClientBenefitAdministrationId)));
            }
        }

        if (snapshots.Any(item => !item.EffectiveDate.HasValue || item.ValidationStatus != HistoryValidationStatus.Valid))
        {
            incomplete = true;
        }

        var events = CollapseDuplicateCarrierTransitions(clientId, snapshots
            .GroupBy(item => item.GroupKey, StringComparer.Ordinal)
            .SelectMany(group => BuildRelationshipEvents(clientId, group))
            .ToList())
            .GroupBy(item => item.StableSourceKey, StringComparer.Ordinal)
            .Select(group => group.First())
            .ToList();
        return new BuildResult(events, records.Count, incomplete, recordedTimeExcludedCount);
    }

    private IEnumerable<ClientHistoryEventReadModel> BuildRelationshipEvents(
        long clientId,
        IEnumerable<RelationshipSnapshot> source)
    {
        var confirmedSnapshots = source
            .Where(item => item.ValidationStatus == HistoryValidationStatus.Valid)
            .ToList();
        var snapshots = confirmedSnapshots
            .Where(item => item.EffectiveDate.HasValue)
            .OrderBy(item => item.EffectiveDate)
            .ThenBy(item => item.EntityId, StringComparer.Ordinal)
            .ToList();
        var transitionSnapshots = snapshots
            .GroupBy(item => new { item.EffectiveDate, item.IdentityId, item.ValidationStatus })
            .Select(group => group.OrderBy(item => item.EntityId, StringComparer.Ordinal).First())
            .OrderBy(item => item.EffectiveDate)
            .ThenBy(item => item.EntityId, StringComparer.Ordinal)
            .ToList();

        RelationshipSnapshot? previous = null;
        foreach (var current in transitionSnapshots)
        {
            var previousEnded = previous?.TerminationDate is DateOnly previousTermination &&
                                previousTermination <= current.EffectiveDate!.Value;
            if (previous is not null && previous.IdentityId == current.IdentityId && !previousEnded)
            {
                previous = current;
                continue;
            }

            // An explicit termination is its own canonical fact. The next relationship is
            // therefore a new assignment, not a second representation of the replacement.
            var changed = previous is not null && !previousEnded;
            var relationLabel = current.RelationshipKind == RelationshipKind.Carrier ? "Carrier" : "Service provider";
            var title = changed
                ? $"{relationLabel} changed to {current.IdentityName}"
                : $"{current.IdentityName} assigned";
            var summary = changed
                ? $"{current.Benefit}: {previous!.IdentityName} to {current.IdentityName} ({current.Role})."
                : $"{current.Role} for {current.Benefit}.";
            var changes = new HistoryChangeSet
            {
                Changes =
                [
                    new HistoryChangeItem
                    {
                        FieldKey = current.RelationshipKind == RelationshipKind.Carrier ? "Carrier" : "ServiceProvider",
                        DisplayLabel = relationLabel,
                        ValueType = current.RelationshipKind == RelationshipKind.Carrier ? HistoryValueType.CarrierReference : HistoryValueType.VendorReference,
                        PreviousValue = changed ? previous!.IdentityName : "",
                        NewValue = current.IdentityName,
                        PreviousReferenceId = changed ? previous!.IdentityId.ToString() : "",
                        NewReferenceId = current.IdentityId.ToString()
                    }
                ]
            };

            yield return CreateEvent(
                clientId,
                current,
                changed ? "changed" : "assigned",
                changed ? $"{relationLabel} change" : $"{relationLabel} assignment",
                title,
                summary,
                changes);
            previous = current;
        }

        foreach (var ended in confirmedSnapshots
                     .Where(item => item.TerminationDate.HasValue)
                     .OrderBy(item => item.TerminationDate)
                     .ThenBy(item => item.EntityId, StringComparer.Ordinal))
        {
            var termination = ended.TerminationDate!.Value;
            var relationLabel = ended.RelationshipKind == RelationshipKind.Carrier ? "Carrier" : "Service provider";
            var relationshipField = ended.RelationshipKind == RelationshipKind.Carrier ? "Carrier" : "ServiceProvider";
            yield return CreateEvent(
                clientId,
                ended with { EffectiveDate = termination },
                "terminated",
                $"{relationLabel} termination",
                $"{ended.IdentityName} assignment ended",
                $"{ended.Role} for {ended.Benefit} ended.",
                new HistoryChangeSet
                {
                    Changes =
                    [
                        new HistoryChangeItem
                        {
                            FieldKey = relationshipField,
                            DisplayLabel = relationLabel,
                            ValueType = ended.RelationshipKind == RelationshipKind.Carrier
                                ? HistoryValueType.CarrierReference
                                : HistoryValueType.VendorReference,
                            PreviousValue = ended.IdentityName,
                            NewValue = "",
                            PreviousReferenceId = ended.IdentityId.ToString(),
                            NewReferenceId = ""
                        },
                        new HistoryChangeItem
                        {
                            FieldKey = "AssignmentStatus",
                            DisplayLabel = "Assignment status",
                            ValueType = HistoryValueType.Text,
                            PreviousValue = "Assigned",
                            NewValue = "Ended"
                        }
                    ]
                });
        }
    }

    private static IReadOnlyList<ClientHistoryEventReadModel> CollapseDuplicateCarrierTransitions(
        long clientId,
        IReadOnlyList<ClientHistoryEventReadModel> events)
    {
        var collapsed = new List<ClientHistoryEventReadModel>();
        var carrierGroups = events
            .Where(item => TryCarrierTransitionKey(item, out _))
            .GroupBy(item =>
            {
                TryCarrierTransitionKey(item, out var key);
                return key;
            }, StringComparer.Ordinal);

        collapsed.AddRange(events.Where(item => !TryCarrierTransitionKey(item, out _)));
        foreach (var group in carrierGroups)
        {
            var ordered = group.OrderBy(item => item.SourceEntityId, StringComparer.Ordinal)
                .ThenBy(item => item.StableSourceKey, StringComparer.Ordinal)
                .ToList();
            var representative = ordered[0];
            var change = representative.ChangeSet.Changes.First(item =>
                item.FieldKey.Equals("Carrier", StringComparison.OrdinalIgnoreCase));
            var transitionEntityId = $"client-{clientId}|{representative.EffectiveDate:yyyyMMdd}|{IdentityKey(change.PreviousReferenceId, change.PreviousValue)}|{IdentityKey(change.NewReferenceId, change.NewValue)}";
            var subtype = string.IsNullOrWhiteSpace(change.NewValue) ? "terminated" :
                string.IsNullOrWhiteSpace(change.PreviousValue) ? "assigned" : "changed";
            var stableKey = ClientHistoryProjectionBuilder.StableKey(SourceModuleName, "CarrierRelationship", transitionEntityId, subtype);
            collapsed.Add(CloneMerged(
                representative,
                stableKey,
                transitionEntityId,
                ordered.SelectMany(item => item.Evidence),
                ordered.SelectMany(item => item.RelatedRecords),
                ordered.Select(item => item.Summary),
                ordered.Select(item => item.RecordedDate)));
        }

        return collapsed
            .OrderBy(item => item.EffectiveDateReference.SortYear)
            .ThenBy(item => item.EffectiveDateReference.SortMonth)
            .ThenBy(item => item.EffectiveDateReference.SortDay)
            .ThenBy(item => item.StableSourceKey, StringComparer.Ordinal)
            .ToList();
    }

    private static bool TryCarrierTransitionKey(ClientHistoryEventReadModel item, out string key)
    {
        var change = item.ChangeSet.Changes.FirstOrDefault(value =>
            value.FieldKey.Equals("Carrier", StringComparison.OrdinalIgnoreCase));
        if (change is null || item.EffectiveDate is null)
        {
            key = "";
            return false;
        }

        key = $"{item.EffectiveDate:yyyyMMdd}|{IdentityKey(change.PreviousReferenceId, change.PreviousValue)}|{IdentityKey(change.NewReferenceId, change.NewValue)}";
        return true;
    }

    private static string IdentityKey(string referenceId, string name) =>
        !string.IsNullOrWhiteSpace(referenceId) ? $"id-{referenceId.Trim()}" :
        !string.IsNullOrWhiteSpace(name) ? $"name-{name.Trim().ToUpperInvariant()}" : "none";

    private static ClientHistoryEventReadModel CloneMerged(
        ClientHistoryEventReadModel source,
        string stableKey,
        string sourceEntityId,
        IEnumerable<HistoryEvidenceReference> evidence,
        IEnumerable<string> relatedRecords,
        IEnumerable<string> summaries,
        IEnumerable<DateTime?> recordedDates)
    {
        var recordedAt = recordedDates.Where(value => value.HasValue).Select(value => value!.Value).DefaultIfEmpty().Max();
        var ownedEvidence = evidence
            .GroupBy(EvidenceKey, StringComparer.Ordinal)
            .Select(group => CloneEvidence(group.First(), stableKey))
            .OrderBy(EvidenceKey, StringComparer.Ordinal)
            .ToList();
        return new ClientHistoryEventReadModel
        {
            EventId = stableKey,
            StableSourceKey = stableKey,
            ClientId = source.ClientId,
            SourceModule = source.SourceModule,
            SourceEntityType = "CarrierRelationship",
            SourceEntityId = sourceEntityId,
            EffectiveDate = source.EffectiveDate,
            EffectiveDateReference = source.EffectiveDateReference,
            RecordedDate = recordedAt == default ? null : recordedAt,
            EventType = source.EventType,
            Category = source.Category,
            Title = source.Title,
            Summary = string.Join(" ", summaries.Where(value => !string.IsNullOrWhiteSpace(value)).Distinct(StringComparer.Ordinal)),
            BusinessRationale = source.BusinessRationale,
            BusinessImpact = source.BusinessImpact,
            ChangeSet = source.ChangeSet,
            FinancialImpact = source.FinancialImpact,
            RecruitmentImpact = source.RecruitmentImpact,
            RetentionImpact = source.RetentionImpact,
            AdministrationImpact = source.AdministrationImpact,
            EmployeeExperienceImpact = source.EmployeeExperienceImpact,
            RequestedBy = source.RequestedBy,
            ApprovedBy = source.ApprovedBy,
            Notes = source.Notes,
            RelatedRecords = relatedRecords.Distinct(StringComparer.Ordinal).OrderBy(value => value, StringComparer.Ordinal).ToList(),
            Evidence = ownedEvidence,
            Confidence = source.Confidence,
            ValidationStatus = source.ValidationStatus,
            Origin = source.Origin,
            IsExcluded = source.IsExcluded,
            IsManuallyEdited = source.IsManuallyEdited,
            CreatedAt = recordedAt,
            UpdatedAt = recordedAt
        };
    }

    private static HistoryEvidenceReference CloneEvidence(HistoryEvidenceReference source, string eventId) => new()
    {
        EventId = eventId,
        DocumentId = source.DocumentId,
        DocumentVersionId = source.DocumentVersionId,
        SourceFactId = source.SourceFactId,
        SourceEntityId = source.SourceEntityId,
        PageNumber = source.PageNumber,
        SheetName = source.SheetName,
        CellRange = source.CellRange,
        SectionName = source.SectionName,
        ExtractedTextReference = source.ExtractedTextReference,
        Confidence = source.Confidence,
        ValidationStatus = source.ValidationStatus
    };

    private static string EvidenceKey(HistoryEvidenceReference evidence) =>
        $"{evidence.DocumentId}|{evidence.DocumentVersionId}|{evidence.SourceFactId}|{evidence.SourceEntityId}|{evidence.PageNumber}|{evidence.SheetName}|{evidence.CellRange}|{evidence.SectionName}|{evidence.ExtractedTextReference}";

    private ClientHistoryEventReadModel CreateEvent(
        long clientId,
        RelationshipSnapshot snapshot,
        string subtype,
        string eventType,
        string title,
        string summary,
        HistoryChangeSet changes) => CanonicalHistoryEventFactory.Create(
            SourceModule,
            snapshot.EntityType,
            snapshot.EntityId,
            subtype,
            clientId,
            HistoryDateReference.FromExact(snapshot.EffectiveDate!.Value),
            snapshot.RecordedAt,
            eventType,
            "Carrier / Vendor",
            title,
            summary,
            snapshot.ValidationStatus,
            snapshot.Confidence,
            [CanonicalHistoryEventFactory.Evidence(
                snapshot.EntityId,
                snapshot.ValidationStatus,
                snapshot.Confidence,
                snapshot.DocumentId,
                page: snapshot.Page,
                extractedText: snapshot.SourceText)],
            snapshot.RelatedRecords,
            changes);

    private static IReadOnlyList<string> Related(
        long? planDesignVersionId,
        long benefitTypeId,
        string identityType,
        long identityId,
        long? parentId = null)
    {
        var related = new List<string> { $"{identityType}:{identityId}", $"BenefitType:{benefitTypeId}" };
        if (planDesignVersionId.HasValue) related.Add($"PlanDesignVersion:{planDesignVersionId.Value}");
        if (parentId.HasValue) related.Add($"BenefitAdministration:{parentId.Value}");
        return related;
    }

    private static HistoryValidationStatus ValidationStatus(string status) =>
        status.Equals(BenefitAdministrationReviewStatuses.Confirmed, StringComparison.OrdinalIgnoreCase)
            ? HistoryValidationStatus.Valid
            : HistoryValidationStatus.NeedsReview;

    private static bool IsRejected(string status) =>
        status.Equals(BenefitAdministrationReviewStatuses.Rejected, StringComparison.OrdinalIgnoreCase);

    private static bool IsInsurerType(string organizationType) =>
        organizationType.Equals(BenefitAdministrationOrganizationTypes.Insurer, StringComparison.OrdinalIgnoreCase) ||
        organizationType.Equals(BenefitAdministrationOrganizationTypes.Carrier, StringComparison.OrdinalIgnoreCase);

    private static string BenefitLabel(ClientBenefitAdministrationRecord record) =>
        !string.IsNullOrWhiteSpace(record.BenefitDisplayName) ? record.BenefitDisplayName.Trim() :
        !string.IsNullOrWhiteSpace(record.BenefitTypeCode) ? record.BenefitTypeCode.Trim() :
        $"benefit {record.BenefitTypeId}";

    private static string RoleLabel(string role) => string.IsNullOrWhiteSpace(role)
        ? "Service provider"
        : string.Join(" ", role.Split('_', StringSplitOptions.RemoveEmptyEntries).Select(word =>
            char.ToUpperInvariant(word[0]) + word[1..].ToLowerInvariant()));

    private static DateOnly? ToDate(DateTime? value) => value.HasValue ? DateOnly.FromDateTime(value.Value) : null;

    private static DateTime? Recorded(params DateTime[] values)
    {
        var recorded = values.Where(value => value != default).DefaultIfEmpty().Max();
        return recorded == default ? null : recorded;
    }

    private static bool WasKnownBy(DateTime? recordedAt, DateOnly? asOfDate) =>
        !asOfDate.HasValue ||
        recordedAt.HasValue && DateOnly.FromDateTime(recordedAt.Value) <= asOfDate.Value;

    private sealed record BuildResult(
        IReadOnlyList<ClientHistoryEventReadModel> Events,
        int SourceCount,
        bool HasIncompleteHistory,
        int RecordedTimeExcludedCount);
    private const string SourceModuleName = "CarrierVendor";
    private enum RelationshipKind { Carrier, Provider }
    private sealed record RelationshipSnapshot(
        string GroupKey,
        string EntityType,
        string EntityId,
        long IdentityId,
        string IdentityName,
        RelationshipKind RelationshipKind,
        string Role,
        string Benefit,
        long BenefitTypeId,
        DateOnly? EffectiveDate,
        DateOnly? TerminationDate,
        DateTime? RecordedAt,
        HistoryValidationStatus ValidationStatus,
        decimal Confidence,
        long? DocumentId,
        int? Page,
        string SourceText,
        IReadOnlyList<string> RelatedRecords);
}

using System.Globalization;
using RenewalReviewBuilder.App.Models.History;

namespace RenewalReviewBuilder.App.Services.ClientHistory;

internal static class CanonicalHistoryEventFactory
{
    internal static ClientHistoryEventReadModel Create(
        string sourceModule,
        string sourceEntityType,
        string sourceEntityId,
        string subtype,
        long clientId,
        HistoryDateReference date,
        DateTime? recordedAt,
        string eventType,
        string category,
        string title,
        string summary,
        HistoryValidationStatus validationStatus,
        decimal confidence,
        IReadOnlyList<HistoryEvidenceReference>? evidence = null,
        IReadOnlyList<string>? relatedRecords = null,
        HistoryChangeSet? changeSet = null,
        string businessRationale = "",
        string approvedBy = "",
        string? stableKey = null)
    {
        stableKey ??= ClientHistoryProjectionBuilder.StableKey(sourceModule, sourceEntityType, sourceEntityId, subtype);
        var createdAt = recordedAt ?? default;
        var ownedEvidence = evidence?.Select(item => CloneEvidence(item, stableKey)).ToList()
            ?? [];
        return new ClientHistoryEventReadModel
        {
            EventId = stableKey,
            StableSourceKey = stableKey,
            ClientId = clientId,
            SourceModule = sourceModule,
            SourceEntityType = sourceEntityType,
            SourceEntityId = sourceEntityId,
            EffectiveDate = date.Precision == HistoryDatePrecision.ExactDate ? date.ExactDate : null,
            EffectiveDateReference = date,
            RecordedDate = recordedAt,
            EventType = eventType,
            Category = category,
            Title = title,
            Summary = summary,
            BusinessRationale = businessRationale,
            ApprovedBy = approvedBy,
            ChangeSet = changeSet ?? new HistoryChangeSet(),
            RelatedRecords = relatedRecords ?? [],
            Evidence = ownedEvidence,
            Confidence = ClampConfidence(confidence),
            ValidationStatus = validationStatus,
            Origin = HistoryOrigin.CanonicalProjection,
            CreatedAt = createdAt,
            UpdatedAt = createdAt
        };
    }

    private static HistoryEvidenceReference CloneEvidence(
        HistoryEvidenceReference source,
        string eventId) => new()
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

    internal static HistoryEvidenceReference Evidence(
        string sourceEntityId,
        HistoryValidationStatus validationStatus,
        decimal confidence,
        long? documentId = null,
        long? documentVersionId = null,
        string sourceFactId = "",
        int? page = null,
        string sheet = "",
        string cellRange = "",
        string section = "",
        string extractedText = "") => new()
    {
        DocumentId = documentId,
        DocumentVersionId = documentVersionId,
        SourceFactId = sourceFactId,
        SourceEntityId = sourceEntityId,
        PageNumber = page,
        SheetName = sheet,
        CellRange = cellRange,
        SectionName = section,
        ExtractedTextReference = extractedText,
        Confidence = ClampConfidence(confidence),
        ValidationStatus = validationStatus
    };

    internal static string Number(decimal? value) =>
        value?.ToString("0.####", CultureInfo.InvariantCulture) ?? "";

    internal static decimal ClampConfidence(decimal? confidence, decimal fallback = .90m) =>
        Math.Clamp(confidence ?? fallback, 0m, 1m);
}

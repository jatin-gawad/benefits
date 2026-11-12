using System.Text.Json;
using Microsoft.Data.Sqlite;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed partial class DatabaseService
{
    private static async Task InsertFlexibleBenefitStagingAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long clientId,
        long sourceDocumentId,
        long sourceDocumentVersionId,
        FlexibleBenefitImportRequest request,
        FlexibleBenefitExtractionResult extraction,
        string finalStatus)
    {
        var snapshot = new
        {
            extraction.Classification,
            extraction.Client,
            extraction.Plan,
            extraction.Classes,
            extraction.Categories,
            extraction.EligibilityRules,
            extraction.FundingConfigurations,
            extraction.Fees,
            extraction.Agreements,
            extraction.Contacts,
            extraction.SourceFields,
            extraction.ValidationIssues,
            extraction.SourcePath,
            extraction.SourceHash,
            extraction.Confidence,
            extraction.Status,
            OcrSummary = extraction.OcrDocument is null
                ? null
                : new
                {
                    extraction.OcrDocument.ExtractionMethod,
                    extraction.OcrDocument.Confidence,
                    PageCount = extraction.OcrDocument.Pages.Count
                }
        };

        using var command = CreateCommand(connection, transaction, @"
INSERT INTO FlexibleBenefitExtractionStaging (
    client_id, source_document_id, source_document_version_id, document_hash, carrier,
    document_type, raw_extraction_json, extraction_status, review_status, created_at, created_by
) VALUES (
    $client_id, $source_document_id, $source_document_version_id, $document_hash, $carrier,
    $document_type, $raw_extraction_json, $extraction_status, $review_status, $created_at, $created_by
);");
        AddParameter(command, "$client_id", clientId);
        AddParameter(command, "$source_document_id", sourceDocumentId);
        AddParameter(command, "$source_document_version_id", sourceDocumentVersionId);
        AddParameter(command, "$document_hash", extraction.SourceHash.ToUpperInvariant());
        AddParameter(command, "$carrier", Clean(extraction.Classification.Administrator));
        AddParameter(command, "$document_type", extraction.Classification.DocumentSubtype.ToString());
        AddParameter(command, "$raw_extraction_json", JsonSerializer.Serialize(snapshot));
        AddParameter(command, "$extraction_status", finalStatus);
        AddParameter(command, "$review_status", ToFlexibleReviewStatus(finalStatus));
        AddParameter(command, "$created_at", ToDateTimeText(DateTime.Now));
        AddParameter(command, "$created_by", Clean(request.ImportedBy));
        await command.ExecuteNonQueryAsync();
    }

    private static async Task InsertFlexibleBenefitSourceFieldsAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long sourceDocumentId,
        long sourceDocumentVersionId,
        FlexibleBenefitExtractionResult extraction,
        IReadOnlyDictionary<string, long> entityIds)
    {
        foreach (var sourceField in extraction.SourceFields)
        {
            var entityId = ResolveFlexibleBenefitSourceEntityId(sourceField, entityIds);
            using var command = CreateCommand(connection, transaction, @"
INSERT INTO FlexibleBenefitPlanSourceField (
    source_document_id, source_document_version_id, entity_type, entity_id, field_name,
    raw_value, normalized_value, source_page_number, bounding_box_json, extraction_method,
    extraction_confidence, was_manually_corrected, corrected_by, corrected_at, review_status
) VALUES (
    $source_document_id, $source_document_version_id, $entity_type, $entity_id, $field_name,
    $raw_value, $normalized_value, $source_page_number, $bounding_box_json, $extraction_method,
    $extraction_confidence, $was_manually_corrected, $corrected_by, $corrected_at, $review_status
);");
            AddParameter(command, "$source_document_id", sourceDocumentId);
            AddParameter(command, "$source_document_version_id", sourceDocumentVersionId);
            AddParameter(command, "$entity_type", Clean(sourceField.EntityType));
            AddParameter(command, "$entity_id", entityId);
            AddParameter(command, "$field_name", Clean(sourceField.FieldName));
            AddParameter(command, "$raw_value", sourceField.RawValue);
            AddParameter(command, "$normalized_value", sourceField.NormalizedValue);
            AddParameter(command, "$source_page_number", sourceField.SourcePageNumber);
            AddParameter(command, "$bounding_box_json", sourceField.BoundingBoxJson);
            AddParameter(
                command,
                "$extraction_method",
                FirstNonBlank(sourceField.ExtractionMethod, extraction.OcrDocument?.ExtractionMethod, PdfExtractionMethods.WindowsLocalOcr));
            AddParameter(command, "$extraction_confidence", sourceField.ExtractionConfidence);
            AddParameter(command, "$was_manually_corrected", sourceField.WasManuallyCorrected ? 1 : 0);
            AddParameter(command, "$corrected_by", Clean(sourceField.CorrectedBy));
            AddParameter(
                command,
                "$corrected_at",
                sourceField.CorrectedAt.HasValue ? ToDateTimeText(sourceField.CorrectedAt.Value) : null);
            AddParameter(command, "$review_status", FirstNonBlank(sourceField.ReviewStatus, FlexibleBenefitReviewStatuses.Pending));
            await command.ExecuteNonQueryAsync();
        }
    }

    private static long? ResolveFlexibleBenefitSourceEntityId(
        FlexibleBenefitPlanSourceField sourceField,
        IReadOnlyDictionary<string, long> entityIds)
    {
        if (!string.IsNullOrWhiteSpace(sourceField.EntityKey) &&
            entityIds.TryGetValue(sourceField.EntityKey, out var keyedEntityId))
        {
            return keyedEntityId;
        }

        var standardKey = sourceField.EntityType switch
        {
            "Client" => "client",
            "FlexibleBenefitPlan" => "plan",
            _ => ""
        };
        return !string.IsNullOrWhiteSpace(standardKey) && entityIds.TryGetValue(standardKey, out var standardEntityId)
            ? standardEntityId
            : null;
    }
}

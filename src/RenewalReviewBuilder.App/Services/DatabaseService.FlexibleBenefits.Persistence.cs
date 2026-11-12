using System.Globalization;
using System.Text.Json;
using Microsoft.Data.Sqlite;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed partial class DatabaseService
{
    public async Task<FlexibleBenefitImportResult> SaveFlexibleBenefitExtractionAsync(
        FlexibleBenefitImportRequest request,
        FlexibleBenefitExtractionResult extraction,
        FlexibleBenefitValidationResult validation)
    {
        ArgumentNullException.ThrowIfNull(request);
        ArgumentNullException.ThrowIfNull(extraction);
        ArgumentNullException.ThrowIfNull(validation);
        if (!request.ClientId.HasValue || request.ClientId <= 0)
        {
            return FlexibleBenefitFailure("Confirm the client before saving flexible-benefits data.", extraction, validation.Issues);
        }

        try
        {
            await InitializeCoreAsync();
            using var connection = await OpenConnectionAsync();
            using var transaction = connection.BeginTransaction();
            var clientId = request.ClientId.Value;
            if (!await FlexibleBenefitClientExistsAsync(connection, transaction, clientId))
            {
                return FlexibleBenefitFailure("The selected client no longer exists in the database.", extraction, validation.Issues, clientId);
            }

            var duplicateDocumentId = await FindDocumentIdByHashAsync(
                connection,
                transaction,
                clientId,
                DocumentTypes.FlexibleBenefitsPlan,
                extraction.SourceHash.ToUpperInvariant());
            if (duplicateDocumentId.HasValue)
            {
                return new FlexibleBenefitImportResult
                {
                    Outcome = FlexibleBenefitImportOutcome.Duplicate,
                    Message = "This exact flexible-benefits document is already saved for the selected client. No records were changed.",
                    ClientId = clientId,
                    DocumentId = duplicateDocumentId,
                    Status = ImportStatuses.NeedsReview,
                    Extraction = extraction
                };
            }

            var currentForPeriod = await FindCurrentFlexibleBenefitPlanAsync(
                connection,
                transaction,
                clientId,
                extraction.Plan.EffectiveDate);
            if (currentForPeriod is not null && request.ImportMode == FlexibleBenefitImportMode.Cancel)
            {
                return new FlexibleBenefitImportResult
                {
                    Outcome = FlexibleBenefitImportOutcome.ExistingVersion,
                    Message = "A different flexible-benefits document already exists for this effective period. The default action is Cancel; choose Import as New Version or Replace Existing Version explicitly.",
                    ClientId = clientId,
                    DocumentVersionId = currentForPeriod.DocumentVersionId,
                    FlexibleBenefitPlanId = currentForPeriod.PlanId,
                    VersionNumber = currentForPeriod.VersionNumber,
                    Status = ImportStatuses.NeedsReview,
                    Extraction = extraction
                };
            }

            var sameHashOtherClient = await SameFlexibleBenefitHashExistsForAnotherClientAsync(
                connection,
                transaction,
                clientId,
                extraction.SourceHash.ToUpperInvariant());
            if (sameHashOtherClient)
            {
                AddFlexibleBenefitValidationIssue(
                    validation,
                    "same_hash_different_client",
                    "Warning",
                    "SourceDocument",
                    "The same file hash is already associated with another client. This import is preserved separately and requires review.");
            }

            var priorCurrent = currentForPeriod ?? await FindCurrentFlexibleBenefitPlanAsync(connection, transaction, clientId, null);
            var nextVersion = await NextFlexibleBenefitVersionNumberAsync(connection, transaction, clientId);
            var finalStatus = validation.Issues.Any(issue => issue.IsFailure)
                ? ImportStatuses.Failed
                : validation.Issues.Count > 0
                    ? ImportStatuses.NeedsReview
                    : ImportStatuses.Completed;
            if (finalStatus == ImportStatuses.Failed)
            {
                return FlexibleBenefitFailure("Validation failed. No normalized flexible-benefits records were saved.", extraction, validation.Issues, clientId);
            }

            await UpdateFlexibleBenefitClientIdentityAsync(connection, transaction, clientId, extraction.Client);
            var carrierId = await FindCarrierIdByNameAsync(connection, transaction, extraction.Classification.Carrier);
            var administratorId = await FindCarrierIdByNameAsync(connection, transaction, extraction.Classification.Administrator);
            var document = await UpsertDocumentAsync(
                connection,
                transaction,
                clientId,
                DocumentTypes.FlexibleBenefitsPlan,
                BenefitsIntelligenceModules.FlexibleBenefitsIntelligence,
                request.FilePath,
                extraction.Classification.Administrator,
                extraction.Plan.EffectiveDate,
                extraction.Plan.EffectiveDate,
                extraction.Plan.TerminationDate);
            if (!document.DocumentId.HasValue)
            {
                return FlexibleBenefitFailure("The signed source document could not be registered. No plan records were saved.", extraction, validation.Issues, clientId);
            }

            if (priorCurrent is not null)
            {
                await MarkFlexibleBenefitVersionHistoricalAsync(
                    connection,
                    transaction,
                    priorCurrent,
                    request.ImportMode == FlexibleBenefitImportMode.ReplaceExistingVersion);
            }

            var documentVersionId = await InsertFlexibleBenefitDocumentVersionAsync(
                connection,
                transaction,
                document.DocumentId.Value,
                clientId,
                extraction,
                nextVersion,
                priorCurrent?.DocumentVersionId,
                request,
                finalStatus);
            var planId = await InsertFlexibleBenefitPlanAsync(
                connection,
                transaction,
                clientId,
                carrierId,
                administratorId,
                document.DocumentId.Value,
                documentVersionId,
                priorCurrent?.PlanId,
                nextVersion,
                request,
                extraction,
                finalStatus);

            var entityIds = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase)
            {
                ["client"] = clientId,
                ["plan"] = planId
            };
            await InsertFlexibleBenefitChildrenAsync(connection, transaction, planId, document.DocumentId.Value, extraction, entityIds);
            await InsertFlexibleBenefitStagingAsync(connection, transaction, clientId, document.DocumentId.Value, documentVersionId, request, extraction, finalStatus);
            await InsertFlexibleBenefitSourceFieldsAsync(connection, transaction, document.DocumentId.Value, documentVersionId, extraction, entityIds);

            await InsertPdfExtractionLogAsync(
                connection,
                transaction,
                document.DocumentId.Value,
                PdfExtractionMethods.WindowsLocalOcr,
                extraction.Confidence,
                "",
                extraction.OcrDocument?.Pages.Count,
                null,
                null,
                "",
                finalStatus == ImportStatuses.NeedsReview);
            var source = await StartImportTrackingAsync(
                connection,
                transaction,
                document,
                SourceTypes.Pdf,
                extraction.Classification.Administrator,
                DocumentTypes.FlexibleBenefitsPlan,
                PdfExtractionMethods.WindowsLocalOcr,
                extraction.Confidence);
            var payload = BuildFlexibleBenefitNormalizedPayload(request, extraction, validation, finalStatus);
            finalStatus = await CompleteImportTrackingAsync(
                connection,
                transaction,
                document.DocumentId.Value,
                source,
                payload,
                finalStatus);
            await ExecuteNonQueryAsync(
                connection,
                transaction,
                "UPDATE Documents SET processing_status = $status WHERE document_id = $document_id;",
                command =>
                {
                    AddParameter(command, "$status", finalStatus);
                    AddParameter(command, "$document_id", document.DocumentId.Value);
                });

            transaction.Commit();
            var result = new FlexibleBenefitImportResult
            {
                Outcome = finalStatus == ImportStatuses.Completed
                    ? FlexibleBenefitImportOutcome.Completed
                    : FlexibleBenefitImportOutcome.NeedsReview,
                Message = finalStatus == ImportStatuses.Completed
                    ? $"Flexible-benefits plan version {nextVersion} was saved and validated."
                    : $"Flexible-benefits plan version {nextVersion} was saved with review items.",
                ClientId = clientId,
                DocumentId = document.DocumentId,
                DocumentVersionId = documentVersionId,
                FlexibleBenefitPlanId = planId,
                VersionNumber = nextVersion,
                Status = finalStatus,
                Extraction = extraction
            };
            result.ValidationIssues.AddRange(validation.Issues);
            return result;
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            return FlexibleBenefitFailure("The flexible-benefits database transaction failed. No partial plan version was committed.", extraction, validation.Issues, request.ClientId);
        }
    }

    private static NormalizedImportPayload BuildFlexibleBenefitNormalizedPayload(
        FlexibleBenefitImportRequest request,
        FlexibleBenefitExtractionResult extraction,
        FlexibleBenefitValidationResult validation,
        string finalStatus)
    {
        var payload = new NormalizedImportPayload
        {
            DocumentType = DocumentTypes.FlexibleBenefitsPlan,
            Client = new NormalizedClient
            {
                ClientName = FirstNonBlank(request.ClientName, extraction.Client.DbaName, extraction.Client.LegalName),
                ClientNumber = request.ClientId?.ToString(CultureInfo.InvariantCulture) ?? "",
                Province = extraction.Client.Province
            },
            Carrier = new NormalizedCarrier { CarrierName = extraction.Classification.Administrator },
            Policy = new NormalizedPolicy
            {
                PolicyNumber = extraction.Plan.PolicyNumber,
                ClientName = FirstNonBlank(request.ClientName, extraction.Client.DbaName, extraction.Client.LegalName),
                CarrierName = extraction.Classification.Administrator,
                EffectiveDate = extraction.Plan.EffectiveDate
            }
        };
        payload.FlexibleBenefitPlans.Add(new NormalizedFlexibleBenefitPlan
        {
            AdministratorName = extraction.Classification.Administrator,
            PlanName = extraction.Plan.PlanName,
            PlanType = extraction.Plan.PlanType.ToString(),
            EffectiveDate = extraction.Plan.EffectiveDate,
            ClassCount = extraction.Classes.Count,
            EnabledCategoryCount = extraction.Categories.Count(category => category.Enabled == true),
            ReviewStatus = finalStatus,
            Extraction = extraction
        });
        foreach (var issue in validation.Issues)
        {
            payload.ValidationIssues.Add(new ValidationIssueRecord
            {
                IssueCode = issue.Code,
                Severity = issue.Severity,
                TableName = "FlexibleBenefitPlan",
                FieldName = issue.FieldName,
                Message = issue.Message,
                RawSourceReference = extraction.SourcePath,
                CreatedAt = DateTime.Now
            });
        }

        return payload;
    }

    private static FlexibleBenefitImportResult FlexibleBenefitFailure(
        string message,
        FlexibleBenefitExtractionResult extraction,
        IEnumerable<FlexibleBenefitValidationIssue> issues,
        long? clientId = null)
    {
        var result = new FlexibleBenefitImportResult
        {
            Outcome = FlexibleBenefitImportOutcome.Failed,
            Message = message,
            ClientId = clientId,
            Status = ImportStatuses.Failed,
            Extraction = extraction
        };
        result.ValidationIssues.AddRange(issues);
        return result;
    }

    private static void AddFlexibleBenefitValidationIssue(
        FlexibleBenefitValidationResult validation,
        string code,
        string severity,
        string fieldName,
        string message)
    {
        if (validation.Issues.Any(issue => issue.Code == code && issue.FieldName == fieldName))
        {
            return;
        }

        validation.Issues.Add(new FlexibleBenefitValidationIssue
        {
            Code = code,
            Severity = severity,
            FieldName = fieldName,
            Message = message
        });
        validation.Status = FlexibleBenefitExtractionStatuses.NeedsReview;
    }

    private sealed record FlexibleBenefitCurrentPlan(
        long PlanId,
        long DocumentVersionId,
        int VersionNumber,
        DateOnly? EffectiveDate);
}

using System.Globalization;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using Microsoft.Data.Sqlite;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services.CanonicalClaims;
using RenewalReviewBuilder.App.Services.Normalization;

namespace RenewalReviewBuilder.App.Services;

public sealed partial class DatabaseService : IManulifeBillingPersistence
{
    private const string DefaultClientStatus = "Active";
    private const string UnknownClientName = "Unknown Client";
    private ExecutiveReportingSnapshot? _executiveReportingCache;
    private DateTime _executiveReportingCacheLoadedAt = DateTime.MinValue;
    private readonly TimeSpan _executiveReportingCacheDuration = TimeSpan.FromSeconds(45);

    public DatabaseService(string? databasePath = null)
    {
        DatabasePath = string.IsNullOrWhiteSpace(databasePath)
            ? AppPaths.BenefitsIntelligenceDatabasePath
            : databasePath;
    }

    public string DatabasePath { get; }

    public void InvalidateExecutiveReportingCache()
    {
        _executiveReportingCache = null;
        _executiveReportingCacheLoadedAt = DateTime.MinValue;
    }

    public DatabaseOperationResult Initialize()
    {
        try
        {
            InitializeCoreAsync().GetAwaiter().GetResult();
            return DatabaseOperationResult.Success($"Benefits Intelligence database is ready: {DatabasePath}", DatabasePath);
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            return DatabaseOperationResult.Failure($"Benefits Intelligence database is unavailable: {ex.Message}", DatabasePath);
        }
    }

    public async Task<IReadOnlyList<PlanBenefitRule>> GetPlanBenefitRulesAsync(long clientId)
    {
        await InitializeCoreAsync();
        using var connection = await OpenConnectionAsync();
        using var command = connection.CreateCommand();
        command.CommandText = """
            SELECT *
            FROM PlanBenefitRule
            WHERE ClientId = $client_id
            ORDER BY PlanDesignFeatureId, DisplayOrder, PlanBenefitRuleId;
            """;
        AddParameter(command, "$client_id", clientId);

        var rules = new List<PlanBenefitRule>();
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            rules.Add(ReadPlanBenefitRule(reader));
        }

        return rules;
    }

    public async Task<DatabaseSaveResult> SaveRenewalAnalysisAsync(
        AnalysisResult analysis,
        bool saveRenewalVersions = true)
    {
        try
        {
            await InitializeCoreAsync();
            using var connection = await OpenConnectionAsync();
            using var transaction = connection.BeginTransaction();

            var effectiveDate = analysis.RenewalCase.RenewalDate ?? analysis.Pdf.RenewalDate;
            var carrier = FirstNonBlank(analysis.Pdf.Carrier, "Manulife");
            var renewalExperiencePeriod = analysis.Pdf.ExperienceLines
                .Select(line => ParseExperiencePeriod(line.Period))
                .Where(period => period.Start.HasValue && period.End.HasValue)
                .OrderByDescending(period => period.End)
                .FirstOrDefault();
            var resolvedClientName = FirstNonBlank(
                ClientNameResolver.Resolve(
                    analysis.RenewalCase.ClientName,
                    analysis.Pdf.ClientName,
                    analysis.RenewalCase.PdfPath),
                UnknownClientName);
            var primaryPolicyNumber = FirstNonBlank(analysis.RenewalCase.PolicyNumber, analysis.Pdf.PolicyNumber);
            var clientId = await UpsertClientAsync(
                connection,
                transaction,
                new ClientRecord
                {
                    ClientName = resolvedClientName,
                    ClientNumber = primaryPolicyNumber,
                    Carrier = carrier,
                    EffectiveDate = effectiveDate,
                    Status = DefaultClientStatus
                });

            var document = await UpsertDocumentAsync(
                connection,
                transaction,
                clientId,
                DocumentTypes.Renewal,
                BenefitsIntelligenceModules.RenewalIntelligence,
                analysis.RenewalCase.PdfPath,
                carrier,
                effectiveDate,
                renewalExperiencePeriod.Start,
                renewalExperiencePeriod.End);
            NormalizedSourceMetadata? renewalImportSource = null;

            if (document.DocumentId is not null)
            {
                await InsertPdfExtractionLogAsync(
                    connection,
                    transaction,
                    document.DocumentId.Value,
                    PdfExtractionMethods.PdfPig,
                    analysis.Pdf.RateLines.Count > 0 ? 0.90m : 0.50m,
                    "",
                    null,
                    analysis.Pdf.RateLines.Count,
                    null,
                    string.Join(" ", analysis.Pdf.Warnings),
                    analysis.Pdf.Warnings.Count > 0);

                await DeleteRenewalsForDocumentAsync(connection, transaction, document.DocumentId.Value);
                await DeleteRatesForDocumentAsync(connection, transaction, document.DocumentId.Value);
                await DeleteRenewalVersionsForDocumentAsync(connection, transaction, document.DocumentId.Value);
                await DeleteClaimsDistributionForDocumentAsync(connection, transaction, document.DocumentId.Value);
                await DeleteClaimsForDocumentAsync(connection, transaction, document.DocumentId.Value);
                renewalImportSource = await StartImportTrackingAsync(
                    connection,
                    transaction,
                    document,
                    SourceTypes.Pdf,
                    carrier,
                    DocumentTypes.Renewal,
                    PdfExtractionMethods.PdfPig,
                    analysis.Pdf.RateLines.Count > 0 ? 0.90m : 0.50m);
            }

            var renewalRows = BuildRenewalRows(analysis, clientId, document.DocumentId);
            foreach (var renewal in renewalRows)
            {
                await InsertRenewalAsync(connection, transaction, renewal);
            }

            var rateRows = BuildRateRows(analysis, clientId, document.DocumentId);
            foreach (var rate in rateRows)
            {
                await InsertRateAsync(connection, transaction, rate);
            }

            var versionRows = 0;
            if (saveRenewalVersions && analysis.RenewalVersions.Count == 0)
            {
                new RenewalEngine().BuildRenewalVersions(analysis);
            }

            if (saveRenewalVersions)
            {
                versionRows = await InsertRenewalVersionsAsync(
                    connection,
                    transaction,
                    analysis,
                    clientId,
                    document.DocumentId,
                    effectiveDate);
            }
            var claimsDistributionRows = await InsertClaimsDistributionReportsAsync(
                connection,
                transaction,
                analysis,
                clientId,
                document.DocumentId);
            var claimsExperienceRows = await InsertRenewalClaimsExperienceAsync(
                connection,
                transaction,
                analysis,
                clientId,
                document.DocumentId);
            var parsedRenewalTermsRows = await UpsertParsedRenewalExperienceTermsAsync(
                connection,
                transaction,
                analysis,
                clientId,
                carrier,
                effectiveDate);

            if (renewalImportSource is not null && document.DocumentId is not null)
            {
                var normalized = NormalizeRenewalData.FromAnalysis(analysis, renewalImportSource);
                await CompleteImportTrackingAsync(connection, transaction, document.DocumentId.Value, renewalImportSource, normalized);
            }

            if (document.DocumentId is not null)
            {
                await UpsertRenewalAnalysisSnapshotAsync(
                    connection,
                    transaction,
                    analysis,
                    clientId,
                    document.DocumentId.Value,
                    effectiveDate,
                    carrier,
                    resolvedClientName,
                    primaryPolicyNumber);
            }

            transaction.Commit();
            analysis.DatabaseClientId = clientId;
            var message = $"Saved {renewalRows.Count} renewal row(s), {rateRows.Count} rate row(s), {versionRows} renewal version row(s), {claimsDistributionRows} claims distribution row(s), {claimsExperienceRows} claims experience row(s), and {parsedRenewalTermsRows} parsed renewal experience term row(s) to Benefits_Intelligence.db.";
            var canonicalResult = await new CanonicalClaimsImportService(this).ImportExistingClaimsAsync(
                new CarrierPdfParseResult
                {
                    Carrier = carrier,
                    Pdf = analysis.Pdf,
                    WorkbookFromPdf = analysis.Workbook
                },
                clientId,
                document.DocumentId,
                renewalImportSource?.ImportBatchId);
            if (canonicalResult is not null)
            {
                message += canonicalResult.Succeeded
                    ? $" {carrier} canonical claims: {canonicalResult.CompletionStatus}; {canonicalResult.PersistedFactCount} fact(s)."
                    : $" {carrier} canonical claims: Failed; legacy data remains saved. {canonicalResult.Message}";
            }

            var asoFinancialRows = 0;
            foreach (var statement in analysis.AsoFinancialActivityStatements)
            {
                if (!string.IsNullOrWhiteSpace(statement.ContractNumber) &&
                    !statement.ContractNumber.Equals(primaryPolicyNumber, StringComparison.OrdinalIgnoreCase))
                {
                    var policyLink = await LinkVerifiedAsoPolicyAsync(
                        clientId,
                        statement.ContractNumber,
                        statement.ClientName,
                        statement.Carrier,
                        statement.SourceFileName,
                        "Embedded renewal operating accounts");
                    if (!policyLink.Succeeded)
                    {
                        message += $" ASO operating accounts require review: {policyLink.Message}";
                        continue;
                    }
                }

                var existingAsoSourceDocumentId = SameFilePath(
                    statement.SourceFilePath,
                    analysis.RenewalCase.PdfPath)
                    ? document.DocumentId
                    : null;
                var asoSave = await SaveAsoFinancialActivityAsync(
                    statement,
                    "Renewal workflow",
                    clientId,
                    existingAsoSourceDocumentId);
                if (asoSave.Succeeded)
                {
                    asoFinancialRows += asoSave.InsertedRows;
                    message += asoSave.WasDuplicate
                        ? " Embedded ASO operating accounts were already saved."
                        : $" Embedded ASO operating accounts: {asoSave.InsertedRows} row(s), {statement.ImportStatus}.";
                }
                else
                {
                    message += $" ASO operating accounts require review: {asoSave.Message}";
                }
            }

            return DatabaseSaveResult.Success(
                message,
                DatabasePath,
                clientId,
                document.DocumentId,
                renewalRows.Count + rateRows.Count + versionRows + claimsDistributionRows + claimsExperienceRows + parsedRenewalTermsRows + asoFinancialRows,
                document.WasReprocessed ? 1 : 0);
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            return DatabaseSaveResult.Failure($"Benefits Intelligence database could not save renewal data: {ex.Message}", DatabasePath);
        }
    }

    /// <summary>
    /// Preserves a recognized policy schedule and a reviewable import state before a carrier extractor exists.
    /// This intentionally writes no Plan Design facts.
    /// </summary>
    public async Task<DatabaseSaveResult> RecordPolicyScheduleReviewAsync(
        ClientRecord client,
        DiscoveredDocument schedule,
        string issueCode,
        string message)
    {
        if (string.IsNullOrWhiteSpace(schedule.FilePath) || !File.Exists(schedule.FilePath))
        {
            return DatabaseSaveResult.Failure("The group policy schedule could not be preserved because its source file was not found.", DatabasePath);
        }

        try
        {
            await InitializeCoreAsync();
            using var connection = await OpenConnectionAsync();
            using var transaction = connection.BeginTransaction();
            var carrier = FirstNonBlank(schedule.Carrier, client.Carrier, "Unknown");
            var clientId = await UpsertClientAsync(connection, transaction, new ClientRecord
            {
                ClientName = FirstNonBlank(client.ClientName, UnknownClientName),
                Carrier = carrier,
                ClientNumber = client.ClientNumber,
                EffectiveDate = client.EffectiveDate,
                Status = FirstNonBlank(client.Status, DefaultClientStatus)
            });
            var document = await UpsertDocumentAsync(
                connection,
                transaction,
                clientId,
                DocumentTypes.GroupPolicySchedule,
                BenefitsIntelligenceModules.BenefitBookletIntelligence,
                schedule.FilePath,
                carrier,
                schedule.RateEffectiveDate,
                null,
                null);
            if (!document.DocumentId.HasValue)
            {
                throw new InvalidOperationException("The group policy schedule document ID was not created.");
            }

            var source = await StartImportTrackingAsync(
                connection,
                transaction,
                document,
                SourceTypes.Pdf,
                carrier,
                DocumentTypes.GroupPolicySchedule,
                "Policy schedule discovery and connector routing",
                schedule.DocumentConfidenceScore);
            var payload = new NormalizedImportPayload
            {
                DocumentType = DocumentTypes.GroupPolicySchedule,
                Client = new NormalizedClient { ClientName = client.ClientName, ClientNumber = client.ClientNumber },
                Carrier = new NormalizedCarrier { CarrierName = carrier },
                Source = source.Copy()
            };
            payload.ValidationIssues.Add(new ValidationIssueRecord
            {
                IssueCode = issueCode,
                Severity = "Warning",
                TableName = "Documents",
                FieldName = "processor_registration",
                Message = message,
                RawSourceReference = schedule.FilePath,
                CreatedAt = DateTime.Now
            });
            var status = await CompleteImportTrackingAsync(
                connection,
                transaction,
                document.DocumentId.Value,
                source,
                payload,
                ImportStatuses.NeedsReview);

            foreach (var reason in schedule.MatchedSignals)
            {
                await ExecuteNonQueryAsync(connection, transaction, @"
INSERT INTO extraction_audit_log (
    source_document_id, source_file, source_page, section_type, field_name,
    extracted_value, source_text, confidence_score, validation_status, validation_message, created_at
)
VALUES (
    $source_document_id, $source_file, NULL, 'Policy Schedule Discovery', 'classification_signal',
    $signal, $signal, $confidence_score, 'Needs review', $message, $created_at
);", command =>
                {
                    AddParameter(command, "$source_document_id", document.DocumentId.Value);
                    AddParameter(command, "$source_file", schedule.FileName);
                    AddParameter(command, "$signal", reason);
                    AddParameter(command, "$confidence_score", schedule.DocumentConfidenceScore);
                    AddParameter(command, "$message", message);
                    AddParameter(command, "$created_at", ToDateTimeText(DateTime.Now));
                });
            }

            transaction.Commit();
            return DatabaseSaveResult.Success(
                $"Group policy schedule preserved with status {status}: {message}",
                DatabasePath,
                clientId,
                document.DocumentId,
                schedule.MatchedSignals.Count + 1,
                document.WasReprocessed ? 1 : 0);
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            return DatabaseSaveResult.Failure($"Group policy schedule review state could not be saved: {ex.Message}", DatabasePath);
        }
    }

    /// <summary>Persists cross-document verification outcomes using the existing reviewable ValidationIssue store.</summary>
    public async Task<DatabaseSaveResult> SavePlanDesignVerificationAsync(
        long clientId,
        IEnumerable<PlanDesignVerificationResult> results)
    {
        try
        {
            await InitializeCoreAsync();
            using var connection = await OpenConnectionAsync();
            using var transaction = connection.BeginTransaction();
            await ExecuteNonQueryAsync(connection, transaction, """
                DELETE FROM ValidationIssue
                WHERE issue_code LIKE 'plan_design_verification_%'
                  AND source_document_id IN (SELECT document_id FROM Documents WHERE client_id = $client_id);
                """, command => AddParameter(command, "$client_id", clientId));
            var saved = 0;
            foreach (var result in results)
            {
                var source = result.Evidence?.Fact ?? result.Approved?.Fact;
                if (source is null) continue;
                await InsertValidationIssueAsync(connection, transaction, new ValidationIssueRecord
                {
                    SourceDocumentId = source.SourceDocumentId,
                    IssueCode = $"plan_design_verification_{ToVerificationIssueCode(result.Status)}",
                    Severity = result.Status is PlanDesignVerificationStatuses.MaterialConflict or PlanDesignVerificationStatuses.ManualReviewRequired ? "Warning" : "Information",
                    TableName = "PlanDesign",
                    FieldName = source.FeatureName,
                    Message = result.Message,
                    RawSourceReference = source.RawSourceReference,
                    CreatedAt = DateTime.Now
                });
                saved++;
            }
            transaction.Commit();
            return DatabaseSaveResult.Success($"Saved {saved} Plan Design cross-document verification result(s).", DatabasePath, clientId, null, saved);
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            return DatabaseSaveResult.Failure($"Plan Design verification results could not be saved: {ex.Message}", DatabasePath);
        }
    }

    public async Task<DatabaseSaveResult> RecordPlanDesignVerificationActionAsync(
        long clientId,
        string className,
        string fieldName,
        string action,
        string note,
        string sourceDocument,
        string reviewer)
    {
        try
        {
            await InitializeCoreAsync();
            using var connection = await OpenConnectionAsync();
            using var transaction = connection.BeginTransaction();
            await ExecuteNonQueryAsync(connection, transaction, """
                INSERT INTO class_comparison_audit (client_id, class_name, field_name, previous_value, new_value, reason, source_document, created_at, created_by)
                VALUES ($client_id, $class_name, $field_name, '', $action, $reason, $source_document, $created_at, $reviewer);
                """, command =>
            {
                AddParameter(command, "$client_id", clientId);
                AddParameter(command, "$class_name", className);
                AddParameter(command, "$field_name", fieldName);
                AddParameter(command, "$action", action);
                AddParameter(command, "$reason", FirstNonBlank(note, "Plan Design verification action."));
                AddParameter(command, "$source_document", sourceDocument);
                AddParameter(command, "$created_at", ToDateTimeText(DateTime.Now));
                AddParameter(command, "$reviewer", FirstNonBlank(reviewer, "Desktop user"));
            });
            transaction.Commit();
            return DatabaseSaveResult.Success("Plan Design verification action recorded; approved values were not changed.", DatabasePath, clientId, null, 1);
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            return DatabaseSaveResult.Failure($"Plan Design verification action could not be recorded: {ex.Message}", DatabasePath);
        }
    }

    private static string ToVerificationIssueCode(string status) =>
        Regex.Replace(status.Trim().ToLowerInvariant(), "[^a-z0-9]+", "_").Trim('_');

    public async Task<DatabaseSaveResult> SaveGenericRenewalArtifactsAsync(
        AnalysisResult analysis,
        long? clientId,
        long? documentId,
        IReadOnlyList<ExtractionAuditEntry> auditEntries,
        IReadOnlyList<PlanDesignFeatureExtraction>? planDesignFeatures = null)
    {
        if (!clientId.HasValue)
        {
            return DatabaseSaveResult.Failure("Generic renewal artifacts were not saved because the client was not saved first.", DatabasePath);
        }

        try
        {
            await InitializeCoreAsync();
            using var connection = await OpenConnectionAsync();
            using var transaction = connection.BeginTransaction();

            if (documentId.HasValue)
            {
                await ExecuteNonQueryAsync(connection, transaction, "DELETE FROM benefit_experience_history WHERE source_document_id = $document_id;", command => AddParameter(command, "$document_id", documentId));
                await ExecuteNonQueryAsync(connection, transaction, "DELETE FROM claims_distribution WHERE source_document_id = $document_id;", command => AddParameter(command, "$document_id", documentId));
                await DeleteClaimsDistributionForDocumentAsync(connection, transaction, documentId.Value);
                await ExecuteNonQueryAsync(connection, transaction, "DELETE FROM demographic_summary WHERE source_document_id = $document_id;", command => AddParameter(command, "$document_id", documentId));
                await ExecuteNonQueryAsync(connection, transaction, "DELETE FROM demographics_by_age_gender WHERE source_document_id = $document_id;", command => AddParameter(command, "$document_id", documentId));
                await DeletePlanDesignFeaturesForDocumentAsync(connection, transaction, documentId.Value);
                await ExecuteNonQueryAsync(connection, transaction, "DELETE FROM extraction_audit_log WHERE source_document_id = $document_id;", command => AddParameter(command, "$document_id", documentId));
            }

            var insertedRows = 0;
            foreach (var row in analysis.Pdf.ExperienceLines)
            {
                using var command = CreateCommand(connection, transaction, @"
INSERT INTO benefit_experience_history (
    client_id,
    source_document_id,
    benefit,
    experience_period,
    premium,
    paid_claims,
    incurred_claims,
    incurred_loss_ratio,
    target_loss_ratio,
    source_reference,
    date_loaded
)
VALUES (
    $client_id,
    $source_document_id,
    $benefit,
    $experience_period,
    $premium,
    $paid_claims,
    $incurred_claims,
    $incurred_loss_ratio,
    $target_loss_ratio,
    $source_reference,
    $date_loaded
);");
                AddParameter(command, "$client_id", clientId.Value);
                AddParameter(command, "$source_document_id", documentId);
                AddParameter(command, "$benefit", row.Benefit);
                AddParameter(command, "$experience_period", row.Period);
                AddParameter(command, "$premium", row.BilledPremium);
                AddParameter(command, "$paid_claims", row.PaidClaims);
                AddParameter(command, "$incurred_claims", row.IncurredClaims);
                AddParameter(command, "$incurred_loss_ratio", row.IncurredLossRatio);
                AddParameter(command, "$target_loss_ratio", NormalizePercentDecimal(row.TargetLossRatio));
                AddParameter(command, "$source_reference", row.SourceReference);
                AddParameter(command, "$date_loaded", ToDateTimeText(DateTime.Now));
                await command.ExecuteNonQueryAsync();
                insertedRows++;
            }

            foreach (var row in analysis.Workbook.ClaimTrends)
            {
                using var command = CreateCommand(connection, transaction, @"
INSERT INTO claims_distribution (
    client_id,
    source_document_id,
    benefit,
    service_category,
    period_label,
    amount_paid,
    percent_of_total,
    norm_percent_of_total,
    source_reference,
    date_loaded
)
VALUES (
    $client_id,
    $source_document_id,
    $benefit,
    $service_category,
    $period_label,
    $amount_paid,
    $percent_of_total,
    $norm_percent_of_total,
    $source_reference,
    $date_loaded
);");
                AddParameter(command, "$client_id", clientId.Value);
                AddParameter(command, "$source_document_id", documentId);
                AddParameter(command, "$benefit", row.Benefit);
                AddParameter(command, "$service_category", row.Category);
                AddParameter(command, "$period_label", row.Period);
                AddParameter(command, "$amount_paid", row.Amount);
                AddParameter(command, "$percent_of_total", row.PercentOfTotal);
                AddParameter(command, "$norm_percent_of_total", row.BenchmarkPercent);
                AddParameter(command, "$source_reference", row.SourceReference);
                AddParameter(command, "$date_loaded", ToDateTimeText(DateTime.Now));
                await command.ExecuteNonQueryAsync();
                insertedRows++;
            }

            insertedRows += await InsertClaimsDistributionReportsAsync(
                connection,
                transaction,
                analysis,
                clientId.Value,
                documentId);

            foreach (var group in analysis.Pdf.LifeDemographics.Concat(analysis.Pdf.LtdDemographics).GroupBy(row => row.Benefit))
            {
                var currentLives = group.Sum(row => (row.CurrentMaleLives ?? 0) + (row.CurrentFemaleLives ?? 0));
                var currentVolume = group.Sum(row => (row.CurrentMaleVolume ?? 0m) + (row.CurrentFemaleVolume ?? 0m));
                using var command = CreateCommand(connection, transaction, @"
INSERT INTO demographic_summary (
    client_id,
    source_document_id,
    benefit,
    period_label,
    lives,
    volume,
    date_loaded
)
VALUES (
    $client_id,
    $source_document_id,
    $benefit,
    'Current',
    $lives,
    $volume,
    $date_loaded
);");
                AddParameter(command, "$client_id", clientId.Value);
                AddParameter(command, "$source_document_id", documentId);
                AddParameter(command, "$benefit", group.Key);
                AddParameter(command, "$lives", currentLives);
                AddParameter(command, "$volume", currentVolume);
                AddParameter(command, "$date_loaded", ToDateTimeText(DateTime.Now));
                await command.ExecuteNonQueryAsync();
                insertedRows++;
            }

            foreach (var row in analysis.Pdf.LifeDemographics.Concat(analysis.Pdf.LtdDemographics))
            {
                insertedRows += await InsertDemographicAgeGenderRowsAsync(connection, transaction, clientId.Value, documentId, row);
            }

            var featuresToSave = (planDesignFeatures ?? []).ToList();
            var featureRules = TieredReimbursementRuleNormalizer.MapPlanDesignFeatures(featuresToSave)
                .ToDictionary(result => result.SourceIndex, result => result.Rule);
            for (var featureIndex = 0; featureIndex < featuresToSave.Count; featureIndex++)
            {
                var featureId = await InsertPlanDesignFeatureAsync(
                    connection,
                    transaction,
                    clientId.Value,
                    documentId,
                    featuresToSave[featureIndex]);
                if (featureRules.TryGetValue(featureIndex, out var rule))
                {
                    rule.PlanDesignFeatureId = featureId;
                    rule.ClientId = clientId.Value;
                    rule.SourceDocumentId = documentId;
                    rule.EffectiveDate ??= analysis.RenewalCase.RenewalDate ?? analysis.Pdf.RenewalDate;
                    await InsertPlanBenefitRuleAsync(connection, transaction, rule);
                }
                insertedRows++;
            }

            foreach (var audit in auditEntries)
            {
                using var command = CreateCommand(connection, transaction, @"
INSERT INTO extraction_audit_log (
    source_document_id,
    source_file,
    source_page,
    section_type,
    field_name,
    extracted_value,
    source_text,
    confidence_score,
    validation_status,
    validation_message,
    created_at
)
VALUES (
    $source_document_id,
    $source_file,
    $source_page,
    $section_type,
    $field_name,
    $extracted_value,
    $source_text,
    $confidence_score,
    $validation_status,
    $validation_message,
    $created_at
);");
                AddParameter(command, "$source_document_id", documentId);
                AddParameter(command, "$source_file", audit.SourceFile);
                AddParameter(command, "$source_page", audit.SourcePage);
                AddParameter(command, "$section_type", audit.SectionType);
                AddParameter(command, "$field_name", audit.FieldName);
                AddParameter(command, "$extracted_value", audit.ExtractedValue);
                AddParameter(command, "$source_text", audit.SourceText);
                AddParameter(command, "$confidence_score", audit.ConfidenceScore);
                AddParameter(command, "$validation_status", audit.ValidationStatus);
                AddParameter(command, "$validation_message", audit.ValidationMessage);
                AddParameter(command, "$created_at", ToDateTimeText(DateTime.Now));
                await command.ExecuteNonQueryAsync();
                insertedRows++;
            }

            transaction.Commit();
            return DatabaseSaveResult.Success(
                $"Saved {insertedRows} generic renewal artifact row(s) to Benefits_Intelligence.db.",
                DatabasePath,
                clientId,
                documentId,
                insertedRows);
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            return DatabaseSaveResult.Failure($"Benefits Intelligence database could not save generic renewal artifacts: {ex.Message}", DatabasePath);
        }
    }

    public async Task<DatabaseSaveResult> SaveClaimsExperienceAsync(
        ClaimsExperienceDataModelResult result,
        IEnumerable<string> sourcePaths,
        string sourceFolder,
        bool monthlyOnly = false)
    {
        try
        {
            await InitializeCoreAsync();
            using var connection = await OpenConnectionAsync();
            using var transaction = connection.BeginTransaction();

            var clientId = await UpsertClientAsync(
                connection,
                transaction,
                new ClientRecord
                {
                    ClientName = FirstNonBlank(result.ClientName, UnknownClientName),
                    ClientNumber = result.ClientNumber,
                    Carrier = FirstNonBlank(result.Carrier, "Unknown"),
                    EffectiveDate = null,
                    Status = DefaultClientStatus
                });

            var documentPaths = ResolveExistingSourceFiles(sourcePaths)
                .Concat(ResolveExistingSourceFiles(result.SourceDocumentPaths))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            var documents = new List<DocumentSaveResult>();
            foreach (var sourcePath in documentPaths)
            {
                var document = await UpsertDocumentAsync(
                    connection,
                    transaction,
                    clientId,
                    DocumentTypes.ClaimsExperience,
                    BenefitsIntelligenceModules.ClaimsExperienceIntelligence,
                    sourcePath,
                    FirstNonBlank(result.Carrier, "Unknown"),
                    result.ExperiencePeriodEnd,
                    result.ExperiencePeriodStart,
                    result.ExperiencePeriodEnd);

                if (document.DocumentId is not null)
                {
                    documents.Add(document);
                    if (Path.GetExtension(sourcePath).Equals(".pdf", StringComparison.OrdinalIgnoreCase))
                    {
                        await InsertPdfExtractionLogAsync(
                            connection,
                            transaction,
                            document.DocumentId.Value,
                            PdfExtractionMethods.PdfPig,
                            result.Summaries.Count > 0 ? 0.85m : 0.45m,
                            "",
                            null,
                            result.TableCount,
                            null,
                            string.Join(" ", result.Warnings),
                            result.Warnings.Count > 0);
                    }
                }
            }

            var sourceDocument = documents.FirstOrDefault(document => document.FilePath.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase)) ??
                                 documents.FirstOrDefault();
            var sourceDocumentId = sourceDocument?.DocumentId;
            NormalizedSourceMetadata? claimsImportSource = null;

            if (sourceDocumentId is not null)
            {
                if (!monthlyOnly)
                {
                    await DeleteClaimsForDocumentAsync(connection, transaction, sourceDocumentId.Value);
                }

                await DeleteMonthlyClaimsForDocumentAsync(connection, transaction, sourceDocumentId.Value);
                if (!monthlyOnly)
                {
                    claimsImportSource = await StartImportTrackingAsync(
                        connection,
                        transaction,
                        sourceDocument!,
                        SourceTypeFromFilePath(sourceDocument!.FilePath),
                        FirstNonBlank(result.Carrier, "Unknown"),
                        DocumentTypes.ClaimsExperience,
                        sourceDocument.FilePath.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) ? PdfExtractionMethods.PdfPig : "Claims workbook parser",
                        result.Summaries.Count > 0 ? 0.85m : 0.45m);
                }
            }

            var insertedRows = 0;
            if (!monthlyOnly)
            {
                foreach (var summary in result.Summaries)
                {
                    await InsertClaimsExperienceAsync(
                        connection,
                        transaction,
                        new ClaimsExperienceRecord
                        {
                            ClientId = clientId,
                            ExperiencePeriodStart = result.ExperiencePeriodStart,
                            ExperiencePeriodEnd = result.ExperiencePeriodEnd,
                            BenefitLine = summary.BenefitLine,
                            Premium = summary.Premium,
                            PaidClaims = summary.PaidClaims,
                            IncurredClaims = summary.IncurredClaims,
                            LossRatio = summary.LossRatio,
                            Lives = summary.Lives,
                            SourceDocumentId = sourceDocumentId,
                            DateLoaded = DateTime.Now
                        });
                    insertedRows++;
                }
            }

            var insertedMonthlyRows = 0;
            var monthlyClaimsToSave = result.AllPeriodMonthlyClaims.Count > 0
                ? result.AllPeriodMonthlyClaims
                : result.MonthlyClaims;
            foreach (var monthlyClaim in monthlyClaimsToSave)
            {
                monthlyClaim.ClientId = clientId;
                monthlyClaim.Carrier = FirstNonBlank(monthlyClaim.Carrier, result.Carrier, "Unknown");
                monthlyClaim.SourceDocumentId = sourceDocumentId;
                monthlyClaim.DateLoaded = DateTime.Now;
                await InsertMonthlyClaimsAsync(connection, transaction, monthlyClaim);
                insertedMonthlyRows++;
            }

            if (claimsImportSource is not null && sourceDocumentId is not null)
            {
                var normalized = NormalizeClaimsData.FromClaimsExperienceResult(result, claimsImportSource);
                await CompleteImportTrackingAsync(connection, transaction, sourceDocumentId.Value, claimsImportSource, normalized);
            }

            transaction.Commit();
            var canonicalMessage = " Annual claims history and canonical facts were preserved.";
            if (!monthlyOnly)
            {
                var canonicalImport = await new CanonicalClaimsImportService(this).ImportExistingClaimsAsync(
                    result,
                    clientId,
                    sourceDocumentId,
                    claimsImportSource?.ImportBatchId);
                canonicalMessage = canonicalImport is null
                    ? " Canonical mapping was not enabled for this carrier in Phase 1."
                    : canonicalImport.Succeeded
                        ? $" Canonical report {canonicalImport.CanonicalReportId} stored with {canonicalImport.PersistedFactCount} fact(s)."
                        : $" Legacy data was saved, but canonical storage needs review: {canonicalImport.Message}";
            }

            return DatabaseSaveResult.Success(
                $"Saved {insertedRows} claims summary row(s) and {insertedMonthlyRows} monthly claim row(s) to Benefits_Intelligence.db.{canonicalMessage}",
                DatabasePath,
                clientId,
                sourceDocumentId,
                insertedRows + insertedMonthlyRows,
                documents.Count(document => document.WasReprocessed));
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            return DatabaseSaveResult.Failure($"Benefits Intelligence database could not save claims data: {ex.Message}", DatabasePath);
        }
    }

    public async Task<ManulifeBillingPersistenceResult> SaveBillingExtractionAsync(
        BillingExtractionResult result,
        ManulifeBillingSaveMode saveMode = ManulifeBillingSaveMode.Cancel,
        long? selectedClientId = null,
        string performedBy = "")
    {
        try
        {
            await InitializeCoreAsync();
            var completion = new ManulifeBillingCompletionPolicy().Evaluate(result, selectedClientId.HasValue);
            if (!completion.CanPersist)
            {
                result.ImportStatus = BillingImportStatuses.Failed;
                return new ManulifeBillingPersistenceResult
                {
                    Outcome = ManulifeBillingSaveOutcomes.Failed,
                    Message = "The Manulife billing statement failed required completion checks and was not saved.",
                    DatabasePath = DatabasePath
                };
            }

            result.ImportStatus = completion.RecommendedStatus;
            using var connection = await OpenConnectionAsync();
            using var transaction = connection.BeginTransaction();

            var clientId = await ResolveBillingClientAsync(
                connection,
                transaction,
                result,
                selectedClientId);
            var carrier = FirstNonBlank(result.Carrier, "Manulife");
            var fullPath = Path.GetFullPath(result.SourceFile);
            var fileHash = await ComputeFileHashAsync(fullPath);

            var exactStatement = await FindManulifeBillingStatementByHashAsync(
                connection,
                transaction,
                clientId,
                carrier,
                fileHash);
            var existingDocumentId = await FindDocumentIdByHashAsync(
                connection,
                transaction,
                clientId,
                DocumentTypes.Billing,
                fileHash);
            if (exactStatement is not null || existingDocumentId is not null)
            {
                transaction.Rollback();
                return new ManulifeBillingPersistenceResult
                {
                    Outcome = ManulifeBillingSaveOutcomes.Duplicate,
                    Message = "This Manulife billing file has already been imported for this client.",
                    DatabasePath = DatabasePath,
                    ClientId = clientId,
                    DocumentId = existingDocumentId,
                    ExistingStatementVersionId = exactStatement
                };
            }

            var existingVersions = await GetActiveManulifeBillingVersionsAsync(
                connection,
                transaction,
                clientId,
                carrier,
                result.BillingPeriod);
            if (existingVersions.Count > 0 && saveMode == ManulifeBillingSaveMode.Cancel)
            {
                transaction.Rollback();
                return new ManulifeBillingPersistenceResult
                {
                    Outcome = ManulifeBillingSaveOutcomes.ExistingVersion,
                    Message = "An existing Manulife billing version was found. The import was cancelled without changing history.",
                    DatabasePath = DatabasePath,
                    ClientId = clientId,
                    ExistingStatementVersionId = existingVersions[0]
                };
            }

            var crossClientHashMatch = await HasDocumentHashForAnotherClientAsync(
                connection,
                transaction,
                clientId,
                DocumentTypes.Billing,
                fileHash);
            if (crossClientHashMatch)
            {
                result.ImportStatus = BillingImportStatuses.NeedsReview;
            }

            var versionNumber = await GetNextManulifeBillingVersionNumberAsync(
                connection,
                transaction,
                clientId,
                carrier,
                result.BillingPeriod);

            var document = await UpsertDocumentAsync(
                connection,
                transaction,
                clientId,
                DocumentTypes.Billing,
                BenefitsIntelligenceModules.BillingPremiumIntelligence,
                result.SourceFile,
                carrier,
                result.PeriodStart,
                result.PeriodStart,
                result.PeriodEnd);
            NormalizedSourceMetadata? billingImportSource = null;

            if (document.DocumentId is not null)
            {
                await InsertPdfExtractionLogAsync(
                    connection,
                    transaction,
                    document.DocumentId.Value,
                    PdfExtractionMethods.PdfPig,
                    result.ReconciliationResults.Any(row => row.Status == BillingReconciliationStatuses.Failed) ? 0.75m : 0.92m,
                    "",
                    result.PageCount,
                    null,
                    null,
                    string.Join(" ", result.Warnings),
                    result.ImportStatus == BillingImportStatuses.NeedsReview);

                await UpdateDocumentProcessingStatusAsync(connection, transaction, document.DocumentId.Value, result.ImportStatus);
                billingImportSource = await StartImportTrackingAsync(
                    connection,
                    transaction,
                    document,
                    SourceTypeFromFilePath(result.SourceFile),
                    carrier,
                    DocumentTypes.Billing,
                    PdfExtractionMethods.PdfPig,
                    result.ImportStatus == BillingImportStatuses.Completed ? 0.92m : 0.75m);
            }

            if (document.DocumentId is null || billingImportSource?.ImportBatchId is null)
            {
                throw new InvalidOperationException("The billing document or import batch could not be created.");
            }
            var importBatchId = billingImportSource.ImportBatchId.Value;

            var replacesVersionId = saveMode == ManulifeBillingSaveMode.ReplaceExistingVersion && existingVersions.Count > 0
                ? existingVersions[0]
                : (long?)null;
            var statementVersionId = await InsertManulifeBillingStatementVersionAsync(
                connection,
                transaction,
                result,
                clientId,
                document.DocumentId.Value,
                importBatchId,
                fileHash,
                versionNumber,
                replacesVersionId,
                FirstNonBlank(performedBy, Environment.UserName));

            foreach (var detail in result.Details)
            {
                detail.ClientId = clientId;
                detail.SourceDocumentId = document.DocumentId;
                detail.StatementVersionId = statementVersionId;
                await InsertBillingOptionPremiumAsync(connection, transaction, detail);
            }

            foreach (var summary in result.ProvinceSummaries)
            {
                summary.ClientId = clientId;
                summary.SourceDocumentId = document.DocumentId;
                summary.StatementVersionId = statementVersionId;
                await InsertBillingProvinceSummaryAsync(connection, transaction, summary);
            }

            foreach (var summary in result.OptionSummaries)
            {
                summary.ClientId = clientId;
                summary.SourceDocumentId = document.DocumentId;
                summary.StatementVersionId = statementVersionId;
                await InsertBillingOptionSummaryAsync(connection, transaction, summary);
            }

            foreach (var mapping in result.ClassMappings)
            {
                mapping.ClientId = clientId;
                await UpsertBillingClassMappingAsync(connection, transaction, mapping);
            }

            foreach (var reconciliation in result.ReconciliationResults)
            {
                reconciliation.ClientId = clientId;
                reconciliation.SourceDocumentId = document.DocumentId;
                reconciliation.StatementVersionId = statementVersionId;
                await InsertBillingReconciliationResultAsync(connection, transaction, reconciliation);
            }

            if (billingImportSource is not null && document.DocumentId is not null)
            {
                var normalized = NormalizeBillingData.FromBilling(result, billingImportSource);
                var trackedStatus = await CompleteImportTrackingAsync(
                    connection,
                    transaction,
                    document.DocumentId.Value,
                    billingImportSource,
                    normalized,
                    result.ImportStatus);
                if (!trackedStatus.Equals(result.ImportStatus, StringComparison.OrdinalIgnoreCase))
                {
                    result.ImportStatus = trackedStatus;
                    await UpdateDocumentProcessingStatusAsync(connection, transaction, document.DocumentId.Value, trackedStatus);
                    await UpdateManulifeBillingStatementStatusAsync(connection, transaction, statementVersionId, trackedStatus);
                }
            }

            if (saveMode == ManulifeBillingSaveMode.ReplaceExistingVersion && existingVersions.Count > 0)
            {
                await MarkManulifeBillingVersionsReplacedAsync(
                    connection,
                    transaction,
                    existingVersions,
                    statementVersionId,
                    clientId,
                    FirstNonBlank(performedBy, Environment.UserName));
            }

            transaction.Commit();

            var rowCount = result.Details.Count +
                           result.ProvinceSummaries.Count +
                           result.OptionSummaries.Count +
                           result.ClassMappings.Count +
                           result.ReconciliationResults.Count;
            return new ManulifeBillingPersistenceResult
            {
                Succeeded = true,
                Outcome = ManulifeBillingSaveOutcomes.Saved,
                Message = $"Saved Manulife billing version {versionNumber} for {result.BillingPeriod}. Status: {result.ImportStatus}.",
                DatabasePath = DatabasePath,
                ClientId = clientId,
                DocumentId = document.DocumentId,
                ImportBatchId = importBatchId,
                StatementVersionId = statementVersionId,
                VersionNumber = versionNumber,
                SavedRowCount = rowCount,
                CrossClientHashMatch = crossClientHashMatch
            };
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            result.ImportStatus = BillingImportStatuses.Failed;
            return new ManulifeBillingPersistenceResult
            {
                Outcome = ManulifeBillingSaveOutcomes.Failed,
                Message = $"Benefits Intelligence database could not save billing data: {ex.Message}",
                DatabasePath = DatabasePath
            };
        }
    }

    public async Task<DatabaseSaveResult> SaveGroupHealthBillingStatementAsync(
        GroupHealthBillingExtractionResult result,
        GroupHealthBillingSaveMode saveMode = GroupHealthBillingSaveMode.Cancel)
    {
        try
        {
            await InitializeCoreAsync();
            using var connection = await OpenConnectionAsync();
            using var transaction = connection.BeginTransaction();

            var clientId = result.ClientId > 0
                ? result.ClientId
                : await UpsertClientAsync(
                    connection,
                    transaction,
                    new ClientRecord
                    {
                        ClientName = FirstNonBlank(result.ClientName, UnknownClientName),
                        ClientNumber = FirstNonBlank(result.ClientNumber, ""),
                        Carrier = FirstNonBlank(result.Carrier, "GroupHEALTH"),
                        EffectiveDate = result.StatementPeriodStart,
                        Status = DefaultClientStatus
                    });

            result.ClientId = clientId;
            var carrier = FirstNonBlank(result.Carrier, "GroupHEALTH");
            var exactDuplicate = await FindGroupHealthBillingStatementAsync(
                connection,
                transaction,
                clientId,
                carrier,
                result.StatementMonth,
                result.StatementYear,
                result.SourceFileHash);
            if (exactDuplicate is not null)
            {
                transaction.Rollback();
                return DatabaseSaveResult.Failure(
                    "This GroupHEALTH billing statement has already been imported for this client, carrier, and statement month.",
                    DatabasePath);
            }

            var existingMonthStatements = await GetActiveGroupHealthStatementIdsForMonthAsync(
                connection,
                transaction,
                clientId,
                carrier,
                result.StatementMonth,
                result.StatementYear);
            var maxVersion = await GetMaxGroupHealthImportVersionAsync(
                connection,
                transaction,
                clientId,
                carrier,
                result.StatementMonth,
                result.StatementYear);

            if (existingMonthStatements.Count > 0 && saveMode == GroupHealthBillingSaveMode.Cancel)
            {
                transaction.Rollback();
                return DatabaseSaveResult.Failure(
                    "Existing billing found. Replace, save as new version, or cancel.",
                    DatabasePath);
            }

            if (existingMonthStatements.Count > 0 && saveMode == GroupHealthBillingSaveMode.Replace)
            {
                await MarkGroupHealthStatementsReplacedAsync(connection, transaction, existingMonthStatements);
            }

            result.ImportVersion = maxVersion + 1;
            var previousRows = await GetPreviousGroupHealthBillingRowsAsync(
                connection,
                transaction,
                clientId,
                carrier,
                result.StatementMonth,
                result.StatementYear);
            var detectedChanges = new GroupHealthBillingChangeDetector().Detect(
                previousRows,
                result.Rows,
                carrier,
                result.StatementMonth,
                result.StatementYear);
            result.EmployeeChangeLog.AddRange(detectedChanges);

            var document = await UpsertDocumentAsync(
                connection,
                transaction,
                clientId,
                DocumentTypes.Billing,
                BenefitsIntelligenceModules.BillingPremiumIntelligence,
                result.SourceFilePath,
                carrier,
                result.StatementPeriodStart,
                result.StatementPeriodStart,
                result.StatementPeriodEnd);

            if (document.DocumentId is not null)
            {
                await InsertPdfExtractionLogAsync(
                    connection,
                    transaction,
                    document.DocumentId.Value,
                    PdfExtractionMethods.PdfPig,
                    result.BillingIntegrityScore / 100m,
                    "",
                    result.PageCount,
                    null,
                    null,
                    result.ImportStatus,
                    !result.ImportStatus.Equals(GroupHealthBillingImportStatuses.Passed, StringComparison.OrdinalIgnoreCase));
                await UpdateDocumentProcessingStatusAsync(connection, transaction, document.DocumentId.Value, result.ImportStatus);
            }

            var statementId = await InsertGroupHealthBillingStatementAsync(connection, transaction, result, clientId, document.DocumentId);
            result.BillingStatementId = statementId;

            foreach (var row in result.Rows)
            {
                row.ClientId = clientId;
                row.BillingStatementId = statementId;
                row.BillingStatementRowId = await InsertGroupHealthBillingRowAsync(connection, transaction, row);
            }

            foreach (var total in result.Totals)
            {
                total.ClientId = clientId;
                total.BillingStatementId = statementId;
                await InsertGroupHealthBillingTotalAsync(connection, transaction, total);
            }

            var rateTableId = await InsertGroupHealthRateTableAsync(connection, transaction, result, statementId, clientId);
            foreach (var line in result.RateLines)
            {
                line.ClientId = clientId;
                line.BillingStatementId = statementId;
                line.ClientRateTableId = rateTableId;
                await InsertGroupHealthRateLineAsync(connection, transaction, line);
            }

            foreach (var validation in result.ValidationResults)
            {
                validation.ClientId = clientId;
                validation.BillingStatementId = statementId;
                await InsertGroupHealthValidationResultAsync(connection, transaction, validation);
            }

            foreach (var change in result.EmployeeChangeLog)
            {
                change.ClientId = clientId;
                change.BillingStatementId = statementId;
                change.Carrier = carrier;
                change.StatementMonth = result.StatementMonth;
                change.StatementYear = result.StatementYear;
                await InsertGroupHealthEmployeeChangeAsync(connection, transaction, change);
            }

            transaction.Commit();

            var rowsSaved = 1 + result.Rows.Count + result.Totals.Count + result.RateLines.Count + result.ValidationResults.Count + result.EmployeeChangeLog.Count;
            return DatabaseSaveResult.Success(
                $"Saved GroupHEALTH billing statement version {result.ImportVersion} for {result.StatementMonth:00}/{result.StatementYear}. Status: {result.ImportStatus}.",
                DatabasePath,
                clientId,
                document.DocumentId,
                rowsSaved,
                existingMonthStatements.Count);
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            return DatabaseSaveResult.Failure($"Benefits Intelligence database could not save GroupHEALTH billing data: {ex.Message}", DatabasePath);
        }
    }

    public async Task<IReadOnlyList<GroupHealthBillingHistoryRecord>> GetGroupHealthBillingHistoryAsync(long clientId)
    {
        await InitializeCoreAsync();
        using var connection = await OpenConnectionAsync();
        using var command = CreateCommand(
            connection,
            null,
            @"
SELECT *
FROM BillingStatements
WHERE client_id = $client_id
  AND is_deleted = 0
ORDER BY statement_year DESC, statement_month DESC, import_version DESC;");
        AddParameter(command, "$client_id", clientId);

        var rows = new List<GroupHealthBillingHistoryRecord>();
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            rows.Add(ReadGroupHealthBillingHistory(reader));
        }

        return rows;
    }

    public async Task<IReadOnlyList<GroupHealthBillingRow>> GetGroupHealthBillingRowsAsync(long billingStatementId)
    {
        await InitializeCoreAsync();
        using var connection = await OpenConnectionAsync();
        using var command = CreateCommand(
            connection,
            null,
            @"
SELECT *
FROM BillingStatementRows
WHERE billing_statement_id = $billing_statement_id
ORDER BY source_page, source_line, employee_name;");
        AddParameter(command, "$billing_statement_id", billingStatementId);

        var rows = new List<GroupHealthBillingRow>();
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            rows.Add(ReadGroupHealthBillingRow(reader));
        }

        return rows;
    }

    public async Task<IReadOnlyList<GroupHealthClientRateTableLine>> GetGroupHealthRateTableLinesAsync(long billingStatementId)
    {
        await InitializeCoreAsync();
        using var connection = await OpenConnectionAsync();
        using var command = CreateCommand(
            connection,
            null,
            @"
SELECT *
FROM ClientRateTableLines
WHERE billing_statement_id = $billing_statement_id
ORDER BY benefit, class_code, tier;");
        AddParameter(command, "$billing_statement_id", billingStatementId);

        var rows = new List<GroupHealthClientRateTableLine>();
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            rows.Add(ReadGroupHealthRateLine(reader));
        }

        return rows;
    }

    public async Task<IReadOnlyList<GroupHealthBillingValidationResult>> GetGroupHealthValidationResultsAsync(long billingStatementId)
    {
        await InitializeCoreAsync();
        using var connection = await OpenConnectionAsync();
        using var command = CreateCommand(
            connection,
            null,
            @"
SELECT *
FROM BillingValidationResults
WHERE billing_statement_id = $billing_statement_id
ORDER BY
    CASE status WHEN 'Failed' THEN 0 WHEN 'Warning' THEN 1 WHEN 'Information' THEN 2 ELSE 3 END,
    validation_level,
    check_name;");
        AddParameter(command, "$billing_statement_id", billingStatementId);

        var rows = new List<GroupHealthBillingValidationResult>();
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            rows.Add(ReadGroupHealthValidationResult(reader));
        }

        return rows;
    }

    public async Task<IReadOnlyList<GroupHealthBillingEmployeeChange>> GetGroupHealthEmployeeChangesAsync(long billingStatementId)
    {
        await InitializeCoreAsync();
        using var connection = await OpenConnectionAsync();
        using var command = CreateCommand(
            connection,
            null,
            @"
SELECT *
FROM BillingEmployeeChangeLog
WHERE billing_statement_id = $billing_statement_id
ORDER BY
    CASE severity WHEN 'Failed' THEN 0 WHEN 'Warning' THEN 1 ELSE 2 END,
    change_type,
    employee_name;");
        AddParameter(command, "$billing_statement_id", billingStatementId);

        var rows = new List<GroupHealthBillingEmployeeChange>();
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            rows.Add(ReadGroupHealthEmployeeChange(reader));
        }

        return rows;
    }

    public async Task<DatabaseOperationResult> DeleteGroupHealthBillingStatementAsync(long billingStatementId)
    {
        try
        {
            await InitializeCoreAsync();
            using var connection = await OpenConnectionAsync();
            await ExecuteNonQueryAsync(
                connection,
                null,
                @"
UPDATE BillingStatements
SET is_deleted = 1,
    deleted_at = $deleted_at,
    updated_at = $updated_at
WHERE billing_statement_id = $billing_statement_id;",
                command =>
                {
                    AddParameter(command, "$deleted_at", ToDateTimeText(DateTime.Now));
                    AddParameter(command, "$updated_at", ToDateTimeText(DateTime.Now));
                    AddParameter(command, "$billing_statement_id", billingStatementId);
                });

            return DatabaseOperationResult.Success("Billing statement was removed from active Billing History.", DatabasePath);
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            return DatabaseOperationResult.Failure($"Billing statement could not be deleted: {ex.Message}", DatabasePath);
        }
    }

    public async Task<DatabaseSaveResult> SaveRenewalAssumptionsAsync(IEnumerable<RenewalAssumptionRecord> assumptions)
    {
        try
        {
            await InitializeCoreAsync();
            using var connection = await OpenConnectionAsync();
            using var transaction = connection.BeginTransaction();

            var insertedRows = 0;
            foreach (var assumption in assumptions.Where(item => item.ClientId > 0))
            {
                await UpsertRenewalAssumptionAsync(connection, transaction, assumption);
                insertedRows++;
            }

            transaction.Commit();
            return DatabaseSaveResult.Success(
                $"Saved {insertedRows} renewal assumption row(s) to Benefits_Intelligence.db.",
                DatabasePath,
                insertedRows: insertedRows);
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            return DatabaseSaveResult.Failure($"Benefits Intelligence database could not save renewal assumptions: {ex.Message}", DatabasePath);
        }
    }

    public async Task<DatabaseSaveResult> SaveRenewalProjectionsAsync(IEnumerable<RenewalProjectionRecord> projections)
    {
        try
        {
            await InitializeCoreAsync();
            using var connection = await OpenConnectionAsync();
            using var transaction = connection.BeginTransaction();

            var insertedRows = 0;
            foreach (var projection in projections.Where(item => item.ClientId > 0))
            {
                await UpsertRenewalProjectionAsync(connection, transaction, projection);
                insertedRows++;
            }

            transaction.Commit();
            return DatabaseSaveResult.Success(
                $"Saved {insertedRows} renewal projection row(s) to Benefits_Intelligence.db.",
                DatabasePath,
                insertedRows: insertedRows);
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            return DatabaseSaveResult.Failure($"Benefits Intelligence database could not save renewal projections: {ex.Message}", DatabasePath);
        }
    }

    public async Task<DatabaseSaveResult> SaveNegotiationWorkbenchAsync(NegotiationWorkbenchSaveRequest request)
    {
        try
        {
            await InitializeCoreAsync();
            using var connection = await OpenConnectionAsync();
            using var transaction = connection.BeginTransaction();

            var clientId = await UpsertClientAsync(
                connection,
                transaction,
                new ClientRecord
                {
                    ClientName = FirstNonBlank(request.ClientName, UnknownClientName),
                    Carrier = request.Carrier,
                    EffectiveDate = request.EffectiveDate,
                    Status = DefaultClientStatus
                });

            await DeleteNegotiationVersionAsync(connection, transaction, clientId, request.EffectiveDate, request.VersionName);
            var versionId = await InsertNegotiationVersionAsync(connection, transaction, clientId, request);

            var insertedRows = 1;
            foreach (var rate in request.Rates)
            {
                await InsertNegotiationRateAsync(connection, transaction, versionId, clientId, request.RenewalId, rate);
                insertedRows++;
            }

            foreach (var calculation in request.Calculations)
            {
                await InsertNegotiationCalculationAsync(connection, transaction, versionId, clientId, request.RenewalId, calculation);
                insertedRows++;
            }

            await UpsertNegotiationAssumptionAsync(connection, transaction, versionId, clientId, request);
            insertedRows++;

            if (request.Note is not null)
            {
                await InsertNegotiationNoteAsync(connection, transaction, versionId, clientId, request.RenewalId, request.Note);
                insertedRows++;
            }

            transaction.Commit();
            return DatabaseSaveResult.Success(
                $"Saved {request.VersionName} negotiation version with {request.Rates.Count} rate row(s) to Benefits_Intelligence.db.",
                DatabasePath,
                clientId,
                insertedRows: insertedRows);
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            return DatabaseSaveResult.Failure($"Benefits Intelligence database could not save negotiation workbench data: {ex.Message}", DatabasePath);
        }
    }

    public async Task<DatabaseSaveResult> SavePlanDesignFromBookletExtractionAsync(
        BookletExtractionResult result,
        IEnumerable<string> inputPaths,
        long? targetClientId = null,
        ClientRecord? targetClient = null)
    {
        var sourcePdfPaths = ResolvePdfInputPaths(inputPaths);
        var planDesignRows = ReadPlanDesignRowsForBookletExtraction(result, sourcePdfPaths);
        if (targetClientId is not > 0 && targetClient is null)
        {
            return DatabaseSaveResult.Failure(
                "Choose an explicit destination client before saving booklet Plan Design data. No Plan Design rows were saved.",
                DatabasePath);
        }

        var client = targetClient ?? BuildPlanDesignClient(planDesignRows);
        ClientRecord? destinationClient = targetClient;
        if (targetClientId is > 0)
        {
            destinationClient = await GetClientAsync(targetClientId.Value);
            if (destinationClient is null)
            {
                return DatabaseSaveResult.Failure(
                    $"The selected destination client {targetClientId.Value.ToString(CultureInfo.InvariantCulture)} no longer exists. No Plan Design rows were saved.",
                    DatabasePath);
            }
        }

        if (destinationClient is null || string.IsNullOrWhiteSpace(destinationClient.ClientName))
        {
            return DatabaseSaveResult.Failure(
                "The explicit destination client does not have a valid client name. No Plan Design rows were saved.",
                DatabasePath);
        }

        var extractedClientNames = planDesignRows
            .Select(row => ClientNameResolver.ToDisplayName(row.ClientName))
            .Where(name => !ClientNameResolver.IsInvalidClientName(name))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        if (extractedClientNames.Count == 0)
        {
            return DatabaseSaveResult.Failure(
                $"The booklet extraction did not provide a reliable client name to verify against '{destinationClient.ClientName}'. No Plan Design rows were saved.",
                DatabasePath);
        }

        var mismatchedClient = extractedClientNames.FirstOrDefault(name =>
            !BookletClientIdentityGuardrail.ClientNamesLikelyMatch(destinationClient.ClientName, name));
        if (!string.IsNullOrWhiteSpace(mismatchedClient))
        {
            return DatabaseSaveResult.Failure(
                $"The selected destination client '{destinationClient.ClientName}' does not match the extracted booklet client '{mismatchedClient}'. No Plan Design rows were saved.",
                DatabasePath);
        }

        return await SavePlanDesignAsync(
            client,
            sourcePdfPaths,
            planDesignRows,
            targetClientId,
            result.BenefitAdministrationRecords);
    }

    public async Task<DatabaseSaveResult> SavePlanDesignAsync(
        ClientRecord client,
        IEnumerable<string> sourcePdfPaths,
        IReadOnlyList<PlanDesignRecord> planDesignRows,
        long? targetClientId = null,
        IReadOnlyList<BenefitAdministrationExtractionCandidate>? benefitAdministrationRecords = null,
        long? expectedSourceDocumentId = null,
        string? expectedSourceFileHash = null,
        string documentType = DocumentTypes.BenefitBooklet,
        string extractionMethod = PdfExtractionMethods.ExternalBookletExtractor,
        string planDesignSource = PlanDesignSources.Booklet,
        IReadOnlyList<ValidationIssueRecord>? validationIssues = null)
    {
        try
        {
            var canonicalPlanDesignRows = PlanDesignCanonicalMapper.NormalizeProvisionRows(planDesignRows);
            await InitializeCoreAsync();
            using var connection = await OpenConnectionAsync();
            using var transaction = connection.BeginTransaction();

            long clientId;
            if (targetClientId.HasValue)
            {
                using var selectedClient = CreateCommand(
                    connection,
                    transaction,
                    "SELECT COUNT(*) FROM Clients WHERE client_id = $client_id;");
                AddParameter(selectedClient, "$client_id", targetClientId.Value);
                if (Convert.ToInt32(await selectedClient.ExecuteScalarAsync(), CultureInfo.InvariantCulture) == 0)
                {
                    throw new InvalidOperationException("The selected client no longer exists.");
                }

                clientId = targetClientId.Value;
            }
            else
            {
                clientId = await UpsertClientAsync(
                    connection,
                    transaction,
                    new ClientRecord
                    {
                        ClientName = FirstNonBlank(client.ClientName, UnknownClientName),
                        ClientNumber = client.ClientNumber,
                        Carrier = client.Carrier,
                        Industry = client.Industry,
                        Province = client.Province,
                        Advisor = client.Advisor,
                        EffectiveDate = client.EffectiveDate,
                        RenewalMonth = client.RenewalMonth,
                        Status = FirstNonBlank(client.Status, DefaultClientStatus)
                    });
            }

            var documents = new List<DocumentSaveResult>();
            foreach (var sourcePath in ResolveExistingSourceFiles(sourcePdfPaths))
            {
                var document = await UpsertDocumentAsync(
                    connection,
                    transaction,
                    clientId,
                    documentType,
                    BenefitsIntelligenceModules.BenefitBookletIntelligence,
                    sourcePath,
                    client.Carrier,
                    client.EffectiveDate,
                    null,
                    null);

                if (document.DocumentId is not null)
                {
                    documents.Add(document);
                    await InsertPdfExtractionLogAsync(
                        connection,
                        transaction,
                        document.DocumentId.Value,
                        extractionMethod,
                        canonicalPlanDesignRows.Count > 0 ? 0.80m : 0.40m,
                        "",
                        null,
                        null,
                        null,
                        "",
                        canonicalPlanDesignRows.Count == 0);
                }
            }

            if (expectedSourceDocumentId.HasValue)
            {
                if (documents.Count != 1 || documents[0].DocumentId != expectedSourceDocumentId.Value)
                {
                    throw new InvalidOperationException(
                        $"Controlled reprocessing expected document {expectedSourceDocumentId.Value}, but the selected source resolved to a different document. No Plan Design rows were replaced.");
                }

                if (!string.IsNullOrWhiteSpace(expectedSourceFileHash) &&
                    !documents[0].FileHash.Equals(expectedSourceFileHash.Trim(), StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException(
                        "The booklet file changed after it was staged. No Plan Design rows were replaced.");
                }
            }

            foreach (var documentId in documents.Select(document => document.DocumentId!.Value).Distinct())
            {
                await DeletePlanDesignForDocumentAsync(connection, transaction, documentId);
                await DeletePlanDesignFeaturesForDocumentAsync(connection, transaction, documentId);
            }

            var preparedPlanFeatures = canonicalPlanDesignRows.Select(planRow =>
            {
                var reimbursementPercent = CanonicalSourceLineageValueGuardrails.NormalizePercentage(planRow.ReimbursementPercent);
                return new
                {
                    PlanRow = planRow,
                    SourceDocumentId = ResolvePlanDesignDocumentId(planRow, documents),
                    ReimbursementPercent = reimbursementPercent,
                    Feature = BuildPlanDesignFeature(planRow, reimbursementPercent, planDesignSource)
                };
            }).ToList();
            var sourceDocuments = documents
                .Where(document => document.DocumentId.HasValue)
                .GroupBy(document => document.DocumentId!.Value)
                .ToDictionary(group => group.Key, group => group.Last());
            var planDesignImportSources = new Dictionary<long, NormalizedSourceMetadata>();
            foreach (var documentId in preparedPlanFeatures
                         .Where(item => item.SourceDocumentId.HasValue)
                         .Select(item => item.SourceDocumentId!.Value)
                         .Distinct())
            {
                if (!sourceDocuments.TryGetValue(documentId, out var sourceDocument))
                {
                    continue;
                }

                planDesignImportSources[documentId] = await StartImportTrackingAsync(
                    connection,
                    transaction,
                    sourceDocument,
                    SourceTypeFromFilePath(sourceDocument.FilePath),
                    client.Carrier,
                    documentType,
                    extractionMethod,
                    0.80m);
            }
            var planDesignSourceDocument = sourceDocuments.Values.FirstOrDefault();
            var planFeatureRules = TieredReimbursementRuleNormalizer
                .MapPlanDesignFeatures(preparedPlanFeatures.Select(item => item.Feature).ToList())
                .ToDictionary(result => result.SourceIndex, result => result.Rule);

            var insertedRows = 0;
            for (var featureIndex = 0; featureIndex < preparedPlanFeatures.Count; featureIndex++)
            {
                var prepared = preparedPlanFeatures[featureIndex];
                var planRow = prepared.PlanRow;
                var sourceDocumentId = prepared.SourceDocumentId;
                await InsertPlanDesignAsync(
                    connection,
                    transaction,
                    new PlanDesignRecord
                    {
                        ClientId = clientId,
                        ClassName = planRow.ClassName,
                        BenefitCategory = planRow.BenefitCategory,
                        BenefitName = planRow.BenefitName,
                        CoverageDescription = planRow.CoverageDescription,
                        ReimbursementPercent = prepared.ReimbursementPercent,
                        ReimbursementText = planRow.ReimbursementText,
                        MaximumAmount = planRow.MaximumAmount,
                        Frequency = planRow.Frequency,
                        WaitingPeriod = planRow.WaitingPeriod,
                        TerminationAge = planRow.TerminationAge,
                        SourceDocumentId = sourceDocumentId,
                        SourceFileName = planRow.SourceFileName,
                        SourceSection = planRow.SourceSection,
                        Notes = planRow.Notes,
                        SourcePage = planRow.SourcePage,
                        DateLoaded = DateTime.Now
                    });
                var planDesignFeatureId = await InsertPlanDesignFeatureAsync(
                    connection,
                    transaction,
                    clientId,
                    sourceDocumentId,
                    prepared.Feature);
                if (planFeatureRules.TryGetValue(featureIndex, out var rule))
                {
                    rule.PlanDesignFeatureId = planDesignFeatureId;
                    rule.ClientId = clientId;
                    rule.SourceDocumentId = sourceDocumentId;
                    rule.EffectiveDate ??= client.EffectiveDate;
                    await InsertPlanBenefitRuleAsync(connection, transaction, rule);
                }
                insertedRows++;
            }

            BenefitAdministrationImportResult? benefitAdministrationImport = null;
            if (benefitAdministrationRecords is { Count: > 0 })
            {
                foreach (var source in benefitAdministrationRecords)
                {
                    source.SourceDocumentId = ResolvePlanDesignDocumentId(
                        new PlanDesignRecord { SourceFileName = source.SourceFile },
                        documents);
                }

                benefitAdministrationImport = await new BenefitAdministrationService(DatabasePath)
                    .ImportInTransactionAsync(
                        connection,
                        transaction,
                        new BenefitAdministrationImportRequest
                        {
                            ClientId = clientId,
                            EffectiveDate = client.EffectiveDate?.ToDateTime(TimeOnly.MinValue),
                            Records = benefitAdministrationRecords
                        });
                if (!benefitAdministrationImport.Succeeded)
                {
                    throw new InvalidOperationException(benefitAdministrationImport.Message);
                }
            }

            foreach (var importSource in planDesignImportSources)
            {
                var sourceRows = preparedPlanFeatures
                    .Where(item => item.SourceDocumentId == importSource.Key)
                    .Select(item => item.PlanRow)
                    .ToList();
                var normalized = NormalizePlanDesignData.FromRows(sourceRows, importSource.Value);
                normalized.DocumentType = documentType;
                if (validationIssues is { Count: > 0 }) normalized.ValidationIssues.AddRange(validationIssues);
                await CompleteImportTrackingAsync(
                    connection,
                    transaction,
                    importSource.Key,
                    importSource.Value,
                    normalized);
            }

            transaction.Commit();
            var benefitAdministrationMessage = benefitAdministrationImport is null
                ? ""
                : $" {benefitAdministrationImport.Message}";
            return DatabaseSaveResult.Success(
                $"Saved {insertedRows} plan design row(s) to Benefits_Intelligence.db.{benefitAdministrationMessage}",
                DatabasePath,
                clientId,
                documents.FirstOrDefault()?.DocumentId,
                insertedRows,
                documents.Count(document => document.WasReprocessed));
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            return DatabaseSaveResult.Failure($"Benefits Intelligence database could not save plan design data: {ex.Message}", DatabasePath);
        }
    }

    public async Task<RenewalReviewBuilder.App.Models.ClientHistory> GetClientHistoryAsync(long clientId)
    {
        var history = new RenewalReviewBuilder.App.Models.ClientHistory { ClientId = clientId };
        history.Documents.AddRange(await GetDocumentsAsync(clientId));
        history.Renewals.AddRange(await GetRenewalHistoryAsync(clientId));
        history.Rates.AddRange(await GetRatesAsync(clientId));
        history.RenewalAssumptions.AddRange(await GetRenewalAssumptionsAsync(clientId));
        history.RenewalProjections.AddRange(await GetRenewalProjectionsAsync(clientId));
        history.RenewalVersions.AddRange(await GetRenewalVersionsAsync(clientId));
        history.ClaimsHistory.AddRange(await GetClaimsHistoryAsync(clientId));
        history.MonthlyClaims.AddRange(await GetMonthlyClaimsAsync(clientId));
        history.PlanDesign.AddRange(await GetPlanDesignAsync(clientId));
        history.Events.AddRange(await GetClientHistoryEventsAsync(clientId));
        return history;
    }

    public async Task<ClientRecord?> GetClientAsync(long clientId)
    {
        await InitializeCoreAsync();
        using var connection = await OpenConnectionAsync();
        using var command = CreateCommand(
            connection,
            null,
            "SELECT * FROM Clients WHERE client_id = $client_id LIMIT 1;");
        AddParameter(command, "$client_id", clientId);

        using var reader = await command.ExecuteReaderAsync();
        return await reader.ReadAsync()
            ? ReadClient(reader)
            : null;
    }

    public async Task<List<ClientRecord>> GetClientsAsync()
    {
        await InitializeCoreAsync();
        using var connection = await OpenConnectionAsync();
        using var command = CreateCommand(connection, null, "SELECT * FROM Clients ORDER BY client_name;");
        var records = new List<ClientRecord>();
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            records.Add(ReadClient(reader));
        }

        return records;
    }

    public async Task<DatabaseSaveResult> UpdateClientIndustryAsync(long clientId, string industry)
    {
        if (clientId <= 0)
        {
            return DatabaseSaveResult.Failure("A saved client is required before an industry can be updated.", DatabasePath);
        }

        if (string.IsNullOrWhiteSpace(industry))
        {
            return DatabaseSaveResult.Failure("Choose an industry before saving.", DatabasePath);
        }

        try
        {
            await InitializeCoreAsync();
            using var connection = await OpenConnectionAsync();
            using var command = CreateCommand(
                connection,
                null,
                @"UPDATE Clients
SET industry = $industry,
    date_updated = $date_updated
WHERE client_id = $client_id;");
            AddParameter(command, "$industry", Clean(industry));
            AddParameter(command, "$date_updated", ToDateTimeText(DateTime.Now));
            AddParameter(command, "$client_id", clientId);
            var updatedRows = await command.ExecuteNonQueryAsync();

            return updatedRows == 1
                ? DatabaseSaveResult.Success("Client industry saved.", DatabasePath, clientId, updatedRows: updatedRows)
                : DatabaseSaveResult.Failure("The selected client was not found, so its industry was not changed.", DatabasePath);
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            return DatabaseSaveResult.Failure($"Client industry could not be saved: {ex.Message}", DatabasePath);
        }
    }

    public async Task<IReadOnlyList<UnderwritingReferenceRate>> GetActiveUnderwritingReferenceRatesAsync()
    {
        await InitializeCoreAsync();
        using var connection = await OpenConnectionAsync();
        using var command = connection.CreateCommand();
        command.CommandText = """
            SELECT
                v.rate_type,
                v.version_name,
                r.age_band,
                r.gender,
                r.rate,
                v.unit_basis,
                v.source_name,
                v.source_sheet,
                r.display_order
            FROM underwriting_reference_rate_versions v
            INNER JOIN underwriting_reference_rates r
                ON r.rate_version_id = v.rate_version_id
            WHERE v.is_active = 1
            ORDER BY v.rate_type, r.display_order, r.gender;
            """;

        var rates = new List<UnderwritingReferenceRate>();
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            rates.Add(new UnderwritingReferenceRate
            {
                RateType = ReadString(reader, "rate_type"),
                VersionName = ReadString(reader, "version_name"),
                AgeBand = ReadString(reader, "age_band"),
                Gender = ReadString(reader, "gender"),
                Rate = ReadDecimal(reader, "rate") ?? 0m,
                UnitBasis = ReadString(reader, "unit_basis"),
                SourceName = ReadString(reader, "source_name"),
                SourceSheet = ReadString(reader, "source_sheet"),
                DisplayOrder = ReadInt(reader, "display_order") ?? 0
            });
        }

        return rates;
    }

    /// <summary>
    /// Permanently removes a client and every database row that is related to that client.
    /// This is deliberately a hard delete because a removed client must no longer appear in
    /// the desktop client list, dashboards, or historical data queries.
    /// </summary>
    public async Task<DatabaseOperationResult> DeleteClientAsync(long clientId)
    {
        if (clientId <= 0)
        {
            return DatabaseOperationResult.Failure("Select a saved client before deleting it.", DatabasePath);
        }

        try
        {
            await InitializeCoreAsync();

            // Foreign keys are enabled for normal application work.  They are temporarily
            // disabled on this private connection so that the dependency graph below can
            // delete all related rows in a deterministic child-before-parent order.
            using var connection = new SqliteConnection($"Data Source={DatabasePath}");
            await connection.OpenAsync();
            await ExecuteNonQueryAsync(connection, null, "PRAGMA foreign_keys = OFF;");
            var existingForeignKeyViolations = await ReadForeignKeyViolationsAsync(connection, null);

            var clientName = await GetClientNameForDeletionAsync(connection, clientId);
            if (clientName is null)
            {
                return DatabaseOperationResult.Failure("The selected client was not found, so nothing was deleted.", DatabasePath);
            }

            var tables = await ReadClientDeletionTablesAsync(connection);
            var ownedTables = tables.Values
                .Select(table => new
                {
                    Table = table,
                    Predicate = BuildClientOwnershipPredicate(table, tables, "target", new HashSet<string>(StringComparer.OrdinalIgnoreCase)),
                    Depth = GetClientOwnershipDepth(table, tables, new HashSet<string>(StringComparer.OrdinalIgnoreCase))
                })
                .Where(item => item.Predicate is not null && item.Depth is not null)
                .OrderByDescending(item => item.Depth)
                .ThenBy(item => item.Table.Name, StringComparer.OrdinalIgnoreCase)
                .ToList();

            using var transaction = connection.BeginTransaction();
            var deletedRows = 0;
            foreach (var table in ownedTables)
            {
                using var command = CreateCommand(
                    connection,
                    transaction,
                    $"DELETE FROM {QuoteIdentifier(table.Table.Name)} AS target WHERE {table.Predicate};");
                AddParameter(command, "$client_id", clientId);
                deletedRows += await command.ExecuteNonQueryAsync();
            }

            using (var command = CreateCommand(
                       connection,
                       transaction,
                       "DELETE FROM Clients WHERE client_id = $client_id;"))
            {
                AddParameter(command, "$client_id", clientId);
                deletedRows += await command.ExecuteNonQueryAsync();
            }

            await VerifyClientDeletionAsync(connection, transaction, clientId, tables.Values);
            var foreignKeyViolationsAfterDelete = await ReadForeignKeyViolationsAsync(connection, transaction);
            var newForeignKeyViolations = foreignKeyViolationsAfterDelete
                .Except(existingForeignKeyViolations, StringComparer.Ordinal)
                .ToList();
            if (newForeignKeyViolations.Count > 0)
            {
                throw new InvalidOperationException($"Client deletion would introduce a foreign-key integrity error: {newForeignKeyViolations[0]}.");
            }

            transaction.Commit();

            await ExecuteNonQueryAsync(connection, null, "PRAGMA foreign_keys = ON;");
            InvalidateExecutiveReportingCache();

            return DatabaseOperationResult.Success(
                $"Deleted {clientName} and {deletedRows.ToString("#,0", CultureInfo.InvariantCulture)} related database record(s).",
                DatabasePath);
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            return DatabaseOperationResult.Failure($"Client deletion could not be completed: {ex.Message}", DatabasePath);
        }
    }

    /// <summary>
    /// Moves all records that belong to one saved client to another saved client and then
    /// removes the source client. The operation is transactional: a conflicting unique
    /// record or an integrity failure leaves both client records unchanged.
    /// </summary>
    public async Task<DatabaseOperationResult> MergeClientAsync(long sourceClientId, long targetClientId)
    {
        if (sourceClientId <= 0 || targetClientId <= 0 || sourceClientId == targetClientId)
        {
            return DatabaseOperationResult.Failure("Choose two different saved clients before merging.", DatabasePath);
        }

        try
        {
            await InitializeCoreAsync();

            using var connection = new SqliteConnection($"Data Source={DatabasePath}");
            await connection.OpenAsync();
            await ExecuteNonQueryAsync(connection, null, "PRAGMA foreign_keys = OFF;");
            var existingForeignKeyViolations = await ReadForeignKeyViolationsAsync(connection, null);

            var sourceName = await GetClientNameForDeletionAsync(connection, sourceClientId);
            var targetName = await GetClientNameForDeletionAsync(connection, targetClientId);
            if (sourceName is null || targetName is null)
            {
                return DatabaseOperationResult.Failure("Both the duplicate and canonical saved clients must exist before merging.", DatabasePath);
            }

            var tables = await ReadClientDeletionTablesAsync(connection);
            using var transaction = connection.BeginTransaction();
            var movedRows = 0;
            foreach (var table in tables.Values
                         .Where(table => table.Columns.Contains("client_id"))
                         .OrderBy(table => table.Name, StringComparer.OrdinalIgnoreCase))
            {
                using var command = CreateCommand(
                    connection,
                    transaction,
                    $"UPDATE {QuoteIdentifier(table.Name)} SET {QuoteIdentifier("client_id")} = $target_client_id WHERE {QuoteIdentifier("client_id")} = $source_client_id;");
                AddParameter(command, "$target_client_id", targetClientId);
                AddParameter(command, "$source_client_id", sourceClientId);
                movedRows += await command.ExecuteNonQueryAsync();
            }

            using (var command = CreateCommand(
                       connection,
                       transaction,
                       "DELETE FROM Clients WHERE client_id = $source_client_id;"))
            {
                AddParameter(command, "$source_client_id", sourceClientId);
                if (await command.ExecuteNonQueryAsync() != 1)
                {
                    throw new InvalidOperationException("The duplicate client could not be removed after its records were transferred.");
                }
            }

            await VerifyClientDeletionAsync(connection, transaction, sourceClientId, tables.Values);
            var foreignKeyViolationsAfterMerge = await ReadForeignKeyViolationsAsync(connection, transaction);
            var newForeignKeyViolations = foreignKeyViolationsAfterMerge
                .Except(existingForeignKeyViolations, StringComparer.Ordinal)
                .ToList();
            if (newForeignKeyViolations.Count > 0)
            {
                throw new InvalidOperationException($"Client merge would introduce a foreign-key integrity error: {newForeignKeyViolations[0]}.");
            }

            transaction.Commit();
            await ExecuteNonQueryAsync(connection, null, "PRAGMA foreign_keys = ON;");
            InvalidateExecutiveReportingCache();

            return DatabaseOperationResult.Success(
                $"Merged {sourceName} into {targetName} and transferred {movedRows.ToString("#,0", CultureInfo.InvariantCulture)} related database record(s).",
                DatabasePath);
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            return DatabaseOperationResult.Failure($"Client merge could not be completed: {ex.Message}", DatabasePath);
        }
    }

    private static async Task<string?> GetClientNameForDeletionAsync(SqliteConnection connection, long clientId)
    {
        using var command = CreateCommand(
            connection,
            null,
            "SELECT client_name FROM Clients WHERE client_id = $client_id LIMIT 1;");
        AddParameter(command, "$client_id", clientId);
        return await command.ExecuteScalarAsync() as string;
    }

    private static async Task<Dictionary<string, ClientDeletionTable>> ReadClientDeletionTablesAsync(SqliteConnection connection)
    {
        var tableNames = new List<string>();
        using (var command = CreateCommand(
                   connection,
                   null,
                   "SELECT name FROM sqlite_master WHERE type = 'table' AND name NOT LIKE 'sqlite_%' AND name <> 'Clients';"))
        using (var reader = await command.ExecuteReaderAsync())
        {
            while (await reader.ReadAsync())
            {
                tableNames.Add(reader.GetString(0));
            }
        }

        var tables = new Dictionary<string, ClientDeletionTable>(StringComparer.OrdinalIgnoreCase);
        foreach (var tableName in tableNames)
        {
            var columns = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            using (var command = CreateCommand(connection, null, $"PRAGMA table_info({QuoteIdentifier(tableName)});"))
            using (var reader = await command.ExecuteReaderAsync())
            {
                while (await reader.ReadAsync())
                {
                    columns.Add(reader.GetString(reader.GetOrdinal("name")));
                }
            }

            var foreignKeys = new List<ClientDeletionForeignKey>();
            using (var command = CreateCommand(connection, null, $"PRAGMA foreign_key_list({QuoteIdentifier(tableName)});"))
            using (var reader = await command.ExecuteReaderAsync())
            {
                while (await reader.ReadAsync())
                {
                    foreignKeys.Add(new ClientDeletionForeignKey(
                        reader.GetString(reader.GetOrdinal("from")),
                        reader.GetString(reader.GetOrdinal("table")),
                        reader.IsDBNull(reader.GetOrdinal("to")) ? "" : reader.GetString(reader.GetOrdinal("to"))));
                }
            }

            tables[tableName] = new ClientDeletionTable(tableName, columns, foreignKeys);
        }

        return tables;
    }

    private static string? BuildClientOwnershipPredicate(
        ClientDeletionTable table,
        IReadOnlyDictionary<string, ClientDeletionTable> tables,
        string tableAlias,
        ISet<string> resolvingTables)
    {
        if (!resolvingTables.Add(table.Name))
        {
            return null;
        }

        var predicates = new List<string>();
        if (table.Columns.Contains("client_id"))
        {
            // Some older rows predate a client_id backfill.  Keep the direct
            // match, but also inspect the table's parent links below so those
            // legacy rows are deleted through their owning record.
            predicates.Add($"{QuoteIdentifier(tableAlias)}.{QuoteIdentifier("client_id")} = $client_id");
        }

        foreach (var foreignKey in table.ForeignKeys)
        {
            if (!tables.TryGetValue(foreignKey.ParentTable, out var parent))
            {
                continue;
            }

            var parentAlias = $"owner_{parent.Name}";
            var parentPredicate = BuildClientOwnershipPredicate(parent, tables, parentAlias, resolvingTables);
            if (parentPredicate is null)
            {
                continue;
            }

            var parentColumn = string.IsNullOrWhiteSpace(foreignKey.ParentColumn) ? "rowid" : foreignKey.ParentColumn;
            predicates.Add(
                $"EXISTS (SELECT 1 FROM {QuoteIdentifier(parent.Name)} AS {QuoteIdentifier(parentAlias)} " +
                $"WHERE {QuoteIdentifier(parentAlias)}.{QuoteIdentifier(parentColumn)} = {QuoteIdentifier(tableAlias)}.{QuoteIdentifier(foreignKey.ChildColumn)} " +
                $"AND {parentPredicate})");
        }

        resolvingTables.Remove(table.Name);
        return predicates.Count switch
        {
            0 => null,
            1 => predicates[0],
            _ => "(" + string.Join(" OR ", predicates) + ")"
        };
    }

    private static int? GetClientOwnershipDepth(
        ClientDeletionTable table,
        IReadOnlyDictionary<string, ClientDeletionTable> tables,
        ISet<string> resolvingTables)
    {
        if (table.Columns.Contains("client_id"))
        {
            return 0;
        }

        if (!resolvingTables.Add(table.Name))
        {
            return null;
        }

        var parentDepths = table.ForeignKeys
            .Select(foreignKey => tables.TryGetValue(foreignKey.ParentTable, out var parent)
                ? GetClientOwnershipDepth(parent, tables, resolvingTables)
                : null)
            .Where(depth => depth.HasValue)
            .Select(depth => depth!.Value)
            .ToList();

        resolvingTables.Remove(table.Name);
        return parentDepths.Count == 0 ? null : parentDepths.Min() + 1;
    }

    private static async Task VerifyClientDeletionAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long clientId,
        IEnumerable<ClientDeletionTable> tables)
    {
        foreach (var table in tables.Where(table => table.Columns.Contains("client_id")))
        {
            using var command = CreateCommand(
                connection,
                transaction,
                $"SELECT COUNT(*) FROM {QuoteIdentifier(table.Name)} WHERE {QuoteIdentifier("client_id")} = $client_id;");
            AddParameter(command, "$client_id", clientId);
            if (Convert.ToInt64(await command.ExecuteScalarAsync(), CultureInfo.InvariantCulture) != 0)
            {
                throw new InvalidOperationException($"Client-linked records remain in {table.Name}.");
            }
        }
    }

    private static async Task<HashSet<string>> ReadForeignKeyViolationsAsync(
        SqliteConnection connection,
        SqliteTransaction? transaction)
    {
        var violations = new HashSet<string>(StringComparer.Ordinal);
        using var command = CreateCommand(connection, transaction, "PRAGMA foreign_key_check;");
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            var table = reader.IsDBNull(0) ? "" : reader.GetString(0);
            var rowId = reader.IsDBNull(1) ? "" : reader.GetInt64(1).ToString(CultureInfo.InvariantCulture);
            var parent = reader.IsDBNull(2) ? "" : reader.GetString(2);
            var foreignKeyId = reader.IsDBNull(3) ? "" : reader.GetInt64(3).ToString(CultureInfo.InvariantCulture);
            violations.Add($"{table}|{rowId}|{parent}|{foreignKeyId}");
        }

        return violations;
    }

    private static string QuoteIdentifier(string identifier)
    {
        return $"\"{identifier.Replace("\"", "\"\"")}\"";
    }

    private sealed record ClientDeletionTable(
        string Name,
        IReadOnlySet<string> Columns,
        IReadOnlyList<ClientDeletionForeignKey> ForeignKeys);

    private sealed record ClientDeletionForeignKey(string ChildColumn, string ParentTable, string ParentColumn);

    public async Task<int> GetActiveClientCountAsync(string? clientName = null)
    {
        await InitializeCoreAsync();
        using var connection = await OpenConnectionAsync();
        var hasClientFilter = !string.IsNullOrWhiteSpace(clientName) &&
                              !clientName.Equals("All Clients", StringComparison.OrdinalIgnoreCase);
        using var command = CreateCommand(
            connection,
            null,
            hasClientFilter
                ? """
                  SELECT COUNT(*)
                  FROM Clients
                  WHERE (TRIM(status) = '' OR LOWER(status) = 'active')
                    AND client_name = $client_name;
                  """
                : """
                  SELECT COUNT(*)
                  FROM Clients
                  WHERE (TRIM(status) = '' OR LOWER(status) = 'active');
                  """);

        if (hasClientFilter)
        {
            AddParameter(command, "$client_name", clientName!.Trim());
        }

        var result = await command.ExecuteScalarAsync();
        return Convert.ToInt32(result, CultureInfo.InvariantCulture);
    }

    public async Task<List<RenewalQueueRecord>> GetRenewalQueueAsync()
    {
        await InitializeCoreAsync();
        using var connection = await OpenConnectionAsync();
        using var command = CreateCommand(
            connection,
            null,
            """
            WITH proposal_versions_raw AS (
                SELECT
                    rv.version_id,
                    rv.client_id,
                    rv.source_document_id,
                    rv.effective_date,
                    rv.renewal_year,
                    rv.date_created,
                    rv.carrier,
                    SUM(rr.current_premium) AS current_monthly_premium,
                    SUM(COALESCE(rr.proposed_premium, rr.version_premium)) AS renewal_monthly_premium
                FROM renewal_versions rv
                LEFT JOIN renewal_rates rr ON rr.version_id = rv.version_id
                LEFT JOIN Documents pd ON pd.document_id = rv.source_document_id
                WHERE rv.version_name = 'Carrier Proposed'
                  AND rv.source_document_id IS NOT NULL
                  AND LOWER(IFNULL(pd.file_name, '')) NOT LIKE '%group information report%'
                GROUP BY rv.version_id, rv.client_id, rv.source_document_id, rv.effective_date, rv.renewal_year, rv.date_created, rv.carrier
            ),
            proposal_versions AS (
                SELECT *
                FROM (
                    SELECT
                        pvr.*,
                        ROW_NUMBER() OVER (
                            PARTITION BY
                                pvr.client_id,
                                IFNULL(pvr.effective_date, ''),
                                IFNULL(pvr.renewal_year, 0),
                                IFNULL(pvr.carrier, '')
                            ORDER BY
                                CASE
                                    WHEN COALESCE(pvr.current_monthly_premium, 0) > 0
                                     AND COALESCE(pvr.renewal_monthly_premium, 0) > 0 THEN 0
                                    WHEN COALESCE(pvr.current_monthly_premium, 0) > 0
                                      OR COALESCE(pvr.renewal_monthly_premium, 0) > 0 THEN 1
                                    ELSE 2
                                END,
                                pvr.date_created DESC,
                                pvr.version_id DESC
                        ) AS queue_row_number
                    FROM proposal_versions_raw pvr
                )
                WHERE queue_row_number = 1
            ),
            renewal_metrics AS (
                SELECT
                    client_id,
                    source_document_id,
                    MIN(renewal_id) AS renewal_id,
                    MAX(CASE
                        WHEN LOWER(benefit_line) LIKE '%extended%'
                          OR LOWER(benefit_line) LIKE '%ehc%'
                          OR LOWER(benefit_line) LIKE '%health%'
                        THEN loss_ratio
                    END) AS ehc_loss_ratio,
                    MAX(CASE
                        WHEN LOWER(benefit_line) = 'dental'
                          OR LOWER(benefit_line) LIKE 'dental %'
                          OR LOWER(benefit_line) LIKE '% dental%'
                        THEN loss_ratio
                    END) AS dental_loss_ratio,
                    MAX(lives) AS lives,
                    MAX(date_loaded) AS date_loaded
                FROM Renewals
                GROUP BY client_id, source_document_id
            ),
            pooling_assumptions AS (
                SELECT
                    client_id,
                    effective_date_key,
                    pooling_threshold,
                    pooling_charge
                FROM (
                    SELECT
                        ra.client_id,
                        IFNULL(ra.effective_date, '') AS effective_date_key,
                        ra.pooling_threshold,
                        ra.pooling_charge,
                        ROW_NUMBER() OVER (
                            PARTITION BY ra.client_id, IFNULL(ra.effective_date, '')
                            ORDER BY
                                CASE
                                    WHEN LOWER(IFNULL(ra.benefit_line, '')) = 'all' THEN 0
                                    WHEN LOWER(IFNULL(ra.benefit_line, '')) LIKE '%extended%'
                                      OR LOWER(IFNULL(ra.benefit_line, '')) LIKE '%ehc%'
                                      OR LOWER(IFNULL(ra.benefit_line, '')) LIKE '%health%' THEN 1
                                    ELSE 2
                                END,
                                ra.updated_at DESC,
                                ra.assumption_id DESC
                        ) AS row_number
                    FROM renewal_assumptions ra
                    WHERE ra.pooling_threshold IS NOT NULL
                       OR ra.pooling_charge IS NOT NULL
                )
                WHERE row_number = 1
            ),
            saved_status AS (
                SELECT
                    rv.client_id,
                    IFNULL(rv.effective_date, '') AS effective_date_key,
                    rv.version_name,
                    rv.status,
                    rv.carrier,
                    ROW_NUMBER() OVER (
                        PARTITION BY rv.client_id, IFNULL(rv.effective_date, '')
                        ORDER BY rv.date_created DESC, rv.version_id DESC
                    ) AS row_number
                FROM renewal_versions rv
                WHERE rv.status IN (
                    'Healthy',
                    'Watch',
                    'High Risk',
                    'Missing Data',
                    'Ready to Review',
                    'Negotiated',
                    'Final Approved',
                    'Completed'
                )
            )
            SELECT
                c.client_id,
                c.client_name,
                c.client_number AS policy_number,
                COALESCE(NULLIF(ss.carrier, ''), NULLIF(pv.carrier, ''), NULLIF(c.carrier, ''), NULLIF(d.carrier, ''), 'Carrier') AS carrier,
                COALESCE(pv.effective_date, d.effective_date, c.effective_date) AS effective_date,
                COALESCE(pv.renewal_year, CAST(strftime('%Y', COALESCE(pv.effective_date, d.effective_date, c.effective_date)) AS INTEGER)) AS renewal_year,
                pv.source_document_id,
                d.file_path AS source_file_path,
                rm.renewal_id,
                CASE WHEN pv.current_monthly_premium IS NULL THEN NULL ELSE ROUND(pv.current_monthly_premium * 12, 2) END AS current_premium,
                CASE WHEN pv.renewal_monthly_premium IS NULL THEN NULL ELSE ROUND(pv.renewal_monthly_premium * 12, 2) END AS renewal_premium,
                CASE
                    WHEN pv.current_monthly_premium > 0 AND pv.renewal_monthly_premium IS NOT NULL
                    THEN ROUND((pv.renewal_monthly_premium / pv.current_monthly_premium - 1) * 100, 1)
                    ELSE NULL
                END AS premium_change_percent,
                rm.ehc_loss_ratio,
                pa.pooling_threshold,
                pa.pooling_charge,
                rm.dental_loss_ratio,
                rm.lives,
                COALESCE(NULLIF(ss.status, ''), '') AS status,
                COALESCE(NULLIF(ss.version_name, ''), 'Carrier Proposed') AS selected_renewal_version,
                COALESCE(rm.date_loaded, pv.date_created) AS date_loaded
            FROM proposal_versions pv
            INNER JOIN Clients c ON c.client_id = pv.client_id
            LEFT JOIN Documents d ON d.document_id = pv.source_document_id
            LEFT JOIN renewal_metrics rm ON rm.client_id = pv.client_id
                AND IFNULL(rm.source_document_id, -1) = IFNULL(pv.source_document_id, -1)
            LEFT JOIN pooling_assumptions pa ON pa.client_id = pv.client_id
                AND pa.effective_date_key = IFNULL(COALESCE(pv.effective_date, d.effective_date, c.effective_date), '')
            LEFT JOIN saved_status ss ON ss.client_id = pv.client_id
                AND ss.effective_date_key = IFNULL(COALESCE(pv.effective_date, d.effective_date, c.effective_date), '')
                AND ss.row_number = 1
            ORDER BY date_loaded DESC, c.client_name;
            """);

        using var reader = await command.ExecuteReaderAsync();
        var records = new List<RenewalQueueRecord>();
        while (await reader.ReadAsync())
        {
            var sourceFilePath = ReadString(reader, "source_file_path");
            var clientName = ClientNameResolver.Resolve(
                ReadString(reader, "client_name"),
                null,
                sourceFilePath);
            if (ClientNameResolver.IsInvalidClientName(clientName))
            {
                continue;
            }

            records.Add(new RenewalQueueRecord
            {
                ClientId = ReadLong(reader, "client_id"),
                RenewalId = ReadNullableLong(reader, "renewal_id") ?? 0,
                SourceDocumentId = ReadNullableLong(reader, "source_document_id"),
                SourceFilePath = sourceFilePath,
                ClientName = clientName,
                PolicyNumber = ReadString(reader, "policy_number"),
                Carrier = ReadString(reader, "carrier"),
                RenewalDate = ReadDateOnly(reader, "effective_date"),
                RenewalYear = ReadInt(reader, "renewal_year") ?? DateTime.Today.Year,
                CurrentPremium = ReadDecimal(reader, "current_premium"),
                RenewalPremium = ReadDecimal(reader, "renewal_premium"),
                PremiumChangePercent = ReadDecimal(reader, "premium_change_percent"),
                EhcLossRatio = ReadDecimal(reader, "ehc_loss_ratio"),
                PoolingThreshold = ReadDecimal(reader, "pooling_threshold"),
                PoolingCharge = ReadDecimal(reader, "pooling_charge"),
                DentalLossRatio = ReadDecimal(reader, "dental_loss_ratio"),
                Lives = ReadInt(reader, "lives"),
                Status = ReadString(reader, "status"),
                SelectedRenewalVersion = ReadString(reader, "selected_renewal_version"),
                DateLoaded = ReadDateTime(reader, "date_loaded") ?? DateTime.MinValue
            });
        }

        return records;
    }

    public async Task<ExecutiveReportingSnapshot> GetExecutiveReportingSnapshotAsync(bool forceRefresh = false)
    {
        if (!forceRefresh &&
            _executiveReportingCache is not null &&
            DateTime.Now - _executiveReportingCacheLoadedAt < _executiveReportingCacheDuration)
        {
            return _executiveReportingCache;
        }

        await InitializeCoreAsync();
        var snapshot = new ExecutiveReportingSnapshot { LoadedAt = DateTime.Now };
        using var connection = await OpenConnectionAsync();

        snapshot.Clients.AddRange(await GetExecutiveClientsAsync(connection));
        snapshot.Renewals.AddRange(await GetRenewalQueueAsync());
        snapshot.RenewalVersions.AddRange(await GetExecutiveRenewalVersionsAsync(connection));
        snapshot.RenewalAssumptions.AddRange(await GetExecutiveRenewalAssumptionsAsync(connection));
        snapshot.Claims.AddRange(await GetExecutiveClaimsAsync(connection));
        snapshot.MonthlyClaims.AddRange(await GetExecutiveMonthlyClaimsAsync(connection));
        snapshot.PlanDesign.AddRange(await GetExecutivePlanDesignAsync(connection));
        snapshot.Proposals.AddRange(await GetExecutiveProposalRecordsAsync(connection));
        snapshot.QuoteScenarios.AddRange(await GetExecutiveQuoteScenariosAsync(connection));
        snapshot.DataCategories.AddRange(await GetExecutiveDataCategoriesAsync(connection));

        _executiveReportingCache = snapshot;
        _executiveReportingCacheLoadedAt = snapshot.LoadedAt;
        return snapshot;
    }

    private static async Task<List<ClientRecord>> GetExecutiveClientsAsync(SqliteConnection connection)
    {
        using var command = CreateCommand(connection, null, "SELECT * FROM Clients ORDER BY client_name");
        var records = new List<ClientRecord>();
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            records.Add(ReadClient(reader));
        }

        return records;
    }

    private static async Task<List<RenewalVersionRecord>> GetExecutiveRenewalVersionsAsync(SqliteConnection connection)
    {
        using var command = CreateCommand(connection, null, "SELECT * FROM renewal_versions ORDER BY date_created DESC, version_id");
        var records = new List<RenewalVersionRecord>();
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            records.Add(ReadRenewalVersion(reader));
        }

        foreach (var version in records)
        {
            version.Rates.AddRange(await GetRenewalVersionRatesAsync(connection, version.VersionId));
            version.Calculations.AddRange(await GetRenewalVersionCalculationsAsync(connection, version.VersionId));
        }

        return records;
    }

    private static async Task<List<RenewalAssumptionRecord>> GetExecutiveRenewalAssumptionsAsync(SqliteConnection connection)
    {
        using var command = CreateCommand(connection, null, "SELECT * FROM renewal_assumptions ORDER BY effective_date DESC, benefit_line");
        var records = new List<RenewalAssumptionRecord>();
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            records.Add(ReadRenewalAssumption(reader));
        }

        return records;
    }

    private static async Task<List<ClaimsExperienceRecord>> GetExecutiveClaimsAsync(SqliteConnection connection)
    {
        using var command = CreateCommand(connection, null, "SELECT * FROM ClaimsExperience ORDER BY experience_period_end DESC, benefit_line");
        var records = new List<ClaimsExperienceRecord>();
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            records.Add(ReadClaimsExperience(reader));
        }

        return records;
    }

    private static async Task<List<MonthlyClaimsRecord>> GetExecutiveMonthlyClaimsAsync(SqliteConnection connection)
    {
        using var command = CreateCommand(connection, null, "SELECT * FROM MonthlyClaims ORDER BY month DESC, benefit_line");
        var records = new List<MonthlyClaimsRecord>();
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            records.Add(ReadMonthlyClaims(reader));
        }

        return records;
    }

    private static async Task<List<PlanDesignRecord>> GetExecutivePlanDesignAsync(SqliteConnection connection)
    {
        using var command = CreateCommand(connection, null, "SELECT * FROM PlanDesign ORDER BY benefit_category, benefit_name");
        var records = new List<PlanDesignRecord>();
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            records.Add(ReadPlanDesign(reader));
        }

        return records;
    }

    private static async Task<List<ExecutiveProposalRecord>> GetExecutiveProposalRecordsAsync(SqliteConnection connection)
    {
        if (!await TableExistsAsync(connection, "proposal_files"))
        {
            return [];
        }

        try
        {
            var sql = await TableExistsAsync(connection, "proposal_summary")
                ? """
                  SELECT
                      pf.id AS proposal_id,
                      IFNULL(ps.client_name, pf.client_id) AS client_name,
                      IFNULL(ps.carrier_name, pf.detected_carrier) AS carrier,
                      pf.status AS status,
                      pf.created_at AS date_loaded,
                      SUM(pp.annual_premium) AS proposed_premium,
                      SUM(pp.employer_cost) AS employer_cost,
                      SUM(pp.employee_cost) AS employee_cost
                  FROM proposal_files pf
                  LEFT JOIN proposal_summary ps ON ps.proposal_id = pf.proposal_id
                  LEFT JOIN proposal_premiums pp ON pp.proposal_id = pf.proposal_id
                  GROUP BY pf.id, ps.client_name, ps.carrier_name, pf.client_id, pf.detected_carrier, pf.status, pf.created_at
                  ORDER BY pf.created_at DESC
                  """
                : """
                  SELECT
                      pf.id AS proposal_id,
                      pf.client_id AS client_name,
                      pf.detected_carrier AS carrier,
                      pf.status AS status,
                      pf.created_at AS date_loaded,
                      NULL AS proposed_premium,
                      NULL AS employer_cost,
                      NULL AS employee_cost
                  FROM proposal_files pf
                  ORDER BY pf.created_at DESC
                  """;

            using var command = CreateCommand(connection, null, sql);
            var records = new List<ExecutiveProposalRecord>();
            using var reader = await command.ExecuteReaderAsync();
            while (await reader.ReadAsync())
            {
                records.Add(new ExecutiveProposalRecord
                {
                    ProposalId = ReadLong(reader, "proposal_id"),
                    ClientName = ReadString(reader, "client_name"),
                    Carrier = ReadString(reader, "carrier"),
                    Status = FirstNonBlank(ReadString(reader, "status"), "Quotes received"),
                    ProposedPremium = ReadDecimal(reader, "proposed_premium"),
                    EmployerCost = ReadDecimal(reader, "employer_cost"),
                    EmployeeCost = ReadDecimal(reader, "employee_cost"),
                    DateLoaded = ReadDateTime(reader, "date_loaded") ?? DateTime.MinValue
                });
            }

            return records;
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            return [];
        }
    }

    private static async Task<List<ExecutiveQuoteScenarioRecord>> GetExecutiveQuoteScenariosAsync(SqliteConnection connection)
    {
        if (!await TableExistsAsync(connection, "quote_scenarios"))
        {
            return [];
        }

        using var command = CreateCommand(
            connection,
            null,
            "SELECT * FROM quote_scenarios ORDER BY date_created DESC, quote_scenario_id DESC");
        var records = new List<ExecutiveQuoteScenarioRecord>();
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            records.Add(new ExecutiveQuoteScenarioRecord
            {
                ScenarioId = ReadLong(reader, "quote_scenario_id"),
                ClientId = ReadLong(reader, "client_id"),
                ScenarioName = ReadString(reader, "scenario_name"),
                Carrier = ReadString(reader, "carrier"),
                Status = FirstNonBlank(ReadString(reader, "status"), "Request received"),
                CreatedAt = ReadDateTime(reader, "date_created") ?? DateTime.MinValue
            });
        }

        return records;
    }

    private static async Task<List<ExecutiveDataCategoryRecord>> GetExecutiveDataCategoriesAsync(SqliteConnection connection)
    {
        using var tableCommand = CreateCommand(
            connection,
            null,
            "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY name;");
        var tableNames = new List<string>();
        using (var reader = await tableCommand.ExecuteReaderAsync())
        {
            while (await reader.ReadAsync())
            {
                tableNames.Add(reader.GetString(0));
            }
        }

        var records = new List<ExecutiveDataCategoryRecord>();
        foreach (var tableName in tableNames)
        {
            try
            {
                using var countCommand = CreateCommand(connection, null, $"SELECT COUNT(*) FROM [{tableName}]");
                var count = Convert.ToInt32(await countCommand.ExecuteScalarAsync(), CultureInfo.InvariantCulture);
                records.Add(new ExecutiveDataCategoryRecord
                {
                    Category = tableName,
                    Records = count,
                    Status = count > 0 ? "Active" : "Ready"
                });
            }
            catch (Exception ex)
            {
                AppLogger.Error(ex);
            }
        }

        return records;
    }

    public async Task<RenewalRecord?> GetLatestRenewalAsync(long clientId)
    {
        var renewals = await QueryRenewalsAsync(
            "SELECT * FROM Renewals WHERE client_id = $client_id ORDER BY renewal_year DESC, date_loaded DESC LIMIT 1",
            command => AddParameter(command, "$client_id", clientId));
        return renewals.FirstOrDefault();
    }

    public async Task<List<RenewalRecord>> GetRenewalHistoryAsync(long clientId)
    {
        return await QueryRenewalsAsync(
            "SELECT * FROM Renewals WHERE client_id = $client_id ORDER BY renewal_year DESC, benefit_line",
            command => AddParameter(command, "$client_id", clientId));
    }

    public async Task<List<RateRecord>> GetRatesAsync(long clientId)
    {
        await InitializeCoreAsync();
        using var connection = await OpenConnectionAsync();
        using var command = CreateCommand(
            connection,
            null,
            "SELECT * FROM Rates WHERE client_id = $client_id ORDER BY effective_date DESC, benefit_line, class");
        AddParameter(command, "$client_id", clientId);

        var records = new List<RateRecord>();
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            records.Add(ReadRate(reader));
        }

        return records;
    }

    public async Task<List<RenewalAssumptionRecord>> GetRenewalAssumptionsAsync(long clientId)
    {
        await InitializeCoreAsync();
        using var connection = await OpenConnectionAsync();
        using var command = CreateCommand(
            connection,
            null,
            "SELECT * FROM renewal_assumptions WHERE client_id = $client_id ORDER BY effective_date DESC, benefit_line");
        AddParameter(command, "$client_id", clientId);

        var records = new List<RenewalAssumptionRecord>();
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            records.Add(ReadRenewalAssumption(reader));
        }

        return records;
    }

    public async Task<List<CarrierAssumptionRecord>> GetCarrierAssumptionsAsync()
    {
        await InitializeCoreAsync();
        using var connection = await OpenConnectionAsync();
        using var command = CreateCommand(
            connection,
            null,
            @"SELECT *
              FROM carrier_assumptions
              ORDER BY carrier, block_name, benefit_line, effective_date DESC");

        var records = new List<CarrierAssumptionRecord>();
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            records.Add(ReadCarrierAssumption(reader));
        }

        return records;
    }

    public async Task<DatabaseSaveResult> SaveCarrierAssumptionsAsync(IEnumerable<CarrierAssumptionRecord> assumptions)
    {
        try
        {
            await InitializeCoreAsync();
            using var connection = await OpenConnectionAsync();
            using var transaction = connection.BeginTransaction();

            var records = assumptions
                .Where(item => !string.IsNullOrWhiteSpace(item.Carrier) && !string.IsNullOrWhiteSpace(item.BenefitLine))
                .ToList();

            foreach (var record in records)
            {
                await UpsertCarrierAssumptionAsync(connection, transaction, record);
            }

            transaction.Commit();
            return DatabaseSaveResult.Success(
                $"Saved {records.Count} carrier underwriting assumption row(s).",
                DatabasePath,
                insertedRows: records.Count);
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            return DatabaseSaveResult.Failure($"Carrier underwriting assumptions could not be saved: {ex.Message}", DatabasePath);
        }
    }

    public async Task<List<PlanDesignFactorRecord>> GetPlanDesignFactorsAsync(PlanDesignFactorFilter? filter = null)
    {
        await InitializeCoreAsync();
        using var connection = await OpenConnectionAsync();
        using var command = CreateCommand(
            connection,
            null,
            @"SELECT *
              FROM plan_design_factors
              ORDER BY carrier_name COLLATE NOCASE,
                       benefit_category COLLATE NOCASE,
                       benefit_line COLLATE NOCASE,
                       COALESCE(effective_date, '') DESC,
                       id DESC");

        var records = new List<PlanDesignFactorRecord>();
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            records.Add(ReadPlanDesignFactor(reader));
        }

        return filter is null
            ? records
            : records.Where(record => MatchesPlanDesignFactorFilter(record, filter)).ToList();
    }

    public async Task<PlanDesignFactorSummary> GetPlanDesignFactorSummaryAsync()
    {
        var today = DateOnly.FromDateTime(DateTime.Today);
        var records = await GetPlanDesignFactorsAsync();
        return new PlanDesignFactorSummary
        {
            TotalFactors = records.Count,
            CarriersCovered = records
                .Select(record => record.CarrierName)
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Count(),
            ActiveFactors = records.Count(record => record.IsActiveOn(today)),
            EstimatedPlaceholderFactors = records.Count(record => record.IsEstimatedPlaceholder),
            ExpiringSoon = records.Count(record => record.IsExpiringSoon(today))
        };
    }

    public async Task<PlanDesignFactorRecord?> GetPlanDesignFactorByIdAsync(long id)
    {
        await InitializeCoreAsync();
        using var connection = await OpenConnectionAsync();
        using var command = CreateCommand(
            connection,
            null,
            "SELECT * FROM plan_design_factors WHERE id = $id;");
        AddParameter(command, "$id", id);

        using var reader = await command.ExecuteReaderAsync();
        return await reader.ReadAsync() ? ReadPlanDesignFactor(reader) : null;
    }

    public async Task<DatabaseSaveResult> SavePlanDesignFactorAsync(PlanDesignFactorRecord factor)
    {
        try
        {
            NormalizePlanDesignFactor(factor);
            var errors = ValidatePlanDesignFactor(factor);
            if (errors.Count > 0)
            {
                return DatabaseSaveResult.Failure(string.Join(" ", errors), DatabasePath);
            }

            await InitializeCoreAsync();
            using var connection = await OpenConnectionAsync();
            using var transaction = connection.BeginTransaction();

            var isNewFactor = factor.Id <= 0;
            if (!isNewFactor)
            {
                await UpdatePlanDesignFactorAsync(connection, transaction, factor);
            }
            else
            {
                factor.Id = await InsertPlanDesignFactorAsync(connection, transaction, factor);
            }

            transaction.Commit();
            return DatabaseSaveResult.Success(
                isNewFactor ? "Added plan design factor." : "Saved plan design factor.",
                DatabasePath,
                insertedRows: 1);
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            return DatabaseSaveResult.Failure($"Plan design factor could not be saved: {ex.Message}", DatabasePath);
        }
    }

    public async Task<DatabaseSaveResult> ArchivePlanDesignFactorAsync(long id)
    {
        try
        {
            await InitializeCoreAsync();
            using var connection = await OpenConnectionAsync();
            using var transaction = connection.BeginTransaction();

            var archiveDate = DateOnly.FromDateTime(DateTime.Today).AddDays(-1);
            await ExecuteNonQueryAsync(
                connection,
                transaction,
                @"UPDATE plan_design_factors
                  SET expiry_date = $expiry_date,
                      notes = CASE
                          WHEN notes IS NULL OR TRIM(notes) = '' THEN $archive_note
                          ELSE notes || ' ' || $archive_note
                      END,
                      updated_at = CURRENT_TIMESTAMP
                  WHERE id = $id;",
                command =>
                {
                    AddParameter(command, "$expiry_date", ToDateText(archiveDate));
                    AddParameter(command, "$archive_note", $"Archived from Plan Design Factors on {DateTime.Today:yyyy-MM-dd}.");
                    AddParameter(command, "$id", id);
                });

            transaction.Commit();
            return DatabaseSaveResult.Success("Archived plan design factor.", DatabasePath, updatedRows: 1);
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            return DatabaseSaveResult.Failure($"Plan design factor could not be archived: {ex.Message}", DatabasePath);
        }
    }

    public async Task<PlanDesignFactorImportResult> ImportPlanDesignFactorsFromExcelAsync(string workbookPath)
    {
        await InitializeCoreAsync();
        var result = new PlanDesignFactorImportResult();
        using var workbook = new XLWorkbook(workbookPath);
        var worksheet = workbook.Worksheets.FirstOrDefault();
        if (worksheet is null)
        {
            result.Warnings.Add("No usable factor rows were found in the workbook.");
            return result;
        }

        var headerRow = worksheet.FirstRowUsed();
        if (headerRow is null)
        {
            result.Warnings.Add("No usable factor rows were found in the workbook.");
            return result;
        }

        var headers = BuildPlanDesignFactorHeaderMap(headerRow);
        var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? headerRow.RowNumber();

        using var connection = await OpenConnectionAsync();
        using var transaction = connection.BeginTransaction();

        for (var rowNumber = headerRow.RowNumber() + 1; rowNumber <= lastRow; rowNumber++)
        {
            var row = worksheet.Row(rowNumber);
            if (row.IsEmpty())
            {
                continue;
            }

            var factor = ReadPlanDesignFactorImportRow(row, headers, Path.GetFileName(workbookPath));
            NormalizePlanDesignFactor(factor);
            var errors = ValidatePlanDesignFactor(factor);
            if (errors.Count > 0)
            {
                result.SkippedCount++;
                result.Warnings.Add($"Row {rowNumber} skipped: {string.Join(" ", errors)}");
                continue;
            }

            var existing = await FindPlanDesignFactorMatchAsync(connection, transaction, factor);
            if (existing.HasValue)
            {
                factor.Id = existing.Value;
                await UpdatePlanDesignFactorAsync(connection, transaction, factor);
                result.UpdatedCount++;
            }
            else
            {
                await InsertPlanDesignFactorAsync(connection, transaction, factor);
                result.InsertedCount++;
            }
        }

        transaction.Commit();
        return result;
    }

    public async Task ExportPlanDesignFactorsToExcelAsync(string workbookPath, PlanDesignFactorFilter? filter = null)
    {
        var records = await GetPlanDesignFactorsAsync(filter);
        Directory.CreateDirectory(Path.GetDirectoryName(workbookPath) ?? ".");

        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Plan Design Factors");
        for (var column = 0; column < PlanDesignFactorExportHeaders.Length; column++)
        {
            worksheet.Cell(1, column + 1).Value = PlanDesignFactorExportHeaders[column];
        }

        var row = 2;
        foreach (var factor in records)
        {
            worksheet.Cell(row, 1).Value = factor.CarrierName;
            worksheet.Cell(row, 2).Value = factor.BenefitCategory;
            worksheet.Cell(row, 3).Value = factor.BenefitLine;
            worksheet.Cell(row, 4).Value = factor.CurrentDesign;
            worksheet.Cell(row, 5).Value = factor.ProposedDesign;
            worksheet.Cell(row, 6).Value = factor.ChangeType;
            worksheet.Cell(row, 7).Value = factor.Factor;
            worksheet.Cell(row, 8).Value = factor.FactorDirection;
            worksheet.Cell(row, 9).Value = ToDateText(factor.EffectiveDate);
            worksheet.Cell(row, 10).Value = ToDateText(factor.ExpiryDate);
            worksheet.Cell(row, 11).Value = factor.SourceDocument;
            worksheet.Cell(row, 12).Value = factor.ConfidenceLevel;
            worksheet.Cell(row, 13).Value = factor.Notes;
            worksheet.Cell(row, 7).Style.NumberFormat.Format = "0.0000";
            row++;
        }

        var headerRange = worksheet.Range(1, 1, 1, PlanDesignFactorExportHeaders.Length);
        headerRange.Style.Font.Bold = true;
        headerRange.Style.Fill.BackgroundColor = XLColor.FromHtml("#003A8C");
        headerRange.Style.Font.FontColor = XLColor.White;
        worksheet.SheetView.FreezeRows(1);
        worksheet.Columns().AdjustToContents();
        workbook.SaveAs(workbookPath);
    }

    // Proposal Intelligence and Renewal Intelligence can call this helper later as:
    // get_factor(carrier_name, benefit_line, current_design, proposed_design, effective_date).
    public async Task<PlanDesignFactorRecord?> GetFactorAsync(
        string carrierName,
        string benefitLine,
        string currentDesign,
        string proposedDesign,
        DateOnly? effectiveDate)
    {
        var asOfDate = effectiveDate ?? DateOnly.FromDateTime(DateTime.Today);
        var carrierKey = NormalizeForMatch(carrierName);
        var benefitKey = NormalizeForMatch(benefitLine);
        var currentKey = NormalizeForMatch(currentDesign);
        var proposedKey = NormalizeForMatch(proposedDesign);
        var placeholderCarrierKey = NormalizeForMatch(PlanDesignFactorRecord.GeneralPlaceholderCarrier);
        var records = await GetPlanDesignFactorsAsync();

        return records
            .Where(record => record.IsActiveOn(asOfDate))
            .Where(record => NormalizeForMatch(record.BenefitLine) == benefitKey)
            .Where(record => string.IsNullOrWhiteSpace(currentDesign) || NormalizeForMatch(record.CurrentDesign) == currentKey)
            .Where(record => string.IsNullOrWhiteSpace(proposedDesign) || NormalizeForMatch(record.ProposedDesign) == proposedKey)
            .Where(record => NormalizeForMatch(record.CarrierName) == carrierKey ||
                             NormalizeForMatch(record.CarrierName) == placeholderCarrierKey)
            .OrderByDescending(record => NormalizeForMatch(record.CarrierName) == carrierKey)
            .ThenByDescending(record => record.EffectiveDate ?? DateOnly.MinValue)
            .ThenByDescending(record => record.UpdatedAt)
            .ThenByDescending(record => record.Id)
            .FirstOrDefault();
    }

    // Proposal Intelligence and Renewal Intelligence can call this helper later as apply_factor(current_premium, factor).
    public static decimal ApplyFactor(decimal currentPremium, decimal factor)
    {
        if (factor <= 0m)
        {
            throw new ArgumentOutOfRangeException(nameof(factor), "Factor must be greater than 0.");
        }

        return decimal.Round(currentPremium * factor, 2);
    }

    // Proposal Intelligence and Renewal Intelligence can call this helper later as estimate_alternate_premium(...).
    public async Task<PlanDesignFactorEstimate?> EstimateAlternatePremiumAsync(
        string carrierName,
        string benefitLine,
        string currentDesign,
        string proposedDesign,
        decimal currentPremium,
        DateOnly? effectiveDate)
    {
        var factor = await GetFactorAsync(carrierName, benefitLine, currentDesign, proposedDesign, effectiveDate);
        return factor is null ? null : BuildPlanDesignFactorEstimate(currentPremium, factor);
    }

    public static PlanDesignFactorEstimate BuildPlanDesignFactorEstimate(decimal currentPremium, PlanDesignFactorRecord factor)
    {
        var revisedPremium = ApplyFactor(currentPremium, factor.Factor);
        var dollarChange = revisedPremium - currentPremium;
        var percentageChange = currentPremium == 0m ? 0m : decimal.Round(dollarChange / currentPremium, 4);

        return new PlanDesignFactorEstimate
        {
            FactorRecord = factor,
            CurrentPremium = currentPremium,
            EstimatedRevisedPremium = revisedPremium,
            DollarChange = dollarChange,
            PercentageChange = percentageChange,
            FactorSource = string.IsNullOrWhiteSpace(factor.SourceDocument)
                ? $"Plan design factor #{factor.Id}"
                : factor.SourceDocument,
            ConfidenceLevel = factor.ConfidenceLevel
        };
    }

    public async Task<List<RenewalProjectionRecord>> GetRenewalProjectionsAsync(long clientId)
    {
        await InitializeCoreAsync();
        using var connection = await OpenConnectionAsync();
        using var command = CreateCommand(
            connection,
            null,
            "SELECT * FROM renewal_projections WHERE client_id = $client_id ORDER BY benefit_line");
        AddParameter(command, "$client_id", clientId);

        var records = new List<RenewalProjectionRecord>();
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            records.Add(ReadRenewalProjection(reader));
        }

        return records;
    }

    public async Task<List<RenewalVersionRecord>> GetRenewalVersionsAsync(long clientId)
    {
        await InitializeCoreAsync();
        using var connection = await OpenConnectionAsync();
        using var command = CreateCommand(
            connection,
            null,
            "SELECT * FROM renewal_versions WHERE client_id = $client_id ORDER BY date_created DESC, version_id");
        AddParameter(command, "$client_id", clientId);

        var records = new List<RenewalVersionRecord>();
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            records.Add(ReadRenewalVersion(reader));
        }

        foreach (var version in records)
        {
            version.Rates.AddRange(await GetRenewalVersionRatesAsync(connection, version.VersionId));
            version.Calculations.AddRange(await GetRenewalVersionCalculationsAsync(connection, version.VersionId));
        }

        return records;
    }

    private static async Task<List<RenewalRateVersionRecord>> GetRenewalVersionRatesAsync(SqliteConnection connection, long versionId)
    {
        using var command = CreateCommand(
            connection,
            null,
            "SELECT * FROM renewal_rates WHERE version_id = $version_id ORDER BY benefit_line, class");
        AddParameter(command, "$version_id", versionId);

        var records = new List<RenewalRateVersionRecord>();
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            records.Add(ReadRenewalRateVersion(reader));
        }

        return records;
    }

    private static async Task<List<RenewalCalculationRecord>> GetRenewalVersionCalculationsAsync(SqliteConnection connection, long versionId)
    {
        using var command = CreateCommand(
            connection,
            null,
            "SELECT * FROM renewal_calculations WHERE version_id = $version_id ORDER BY benefit_line");
        AddParameter(command, "$version_id", versionId);

        var records = new List<RenewalCalculationRecord>();
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            records.Add(ReadRenewalCalculation(reader));
        }

        return records;
    }

    public async Task<List<ClaimsExperienceRecord>> GetClaimsHistoryAsync(long clientId)
    {
        await InitializeCoreAsync();
        using var connection = await OpenConnectionAsync();
        using var command = CreateCommand(
            connection,
            null,
            "SELECT * FROM ClaimsExperience WHERE client_id = $client_id ORDER BY experience_period_end DESC, benefit_line");
        AddParameter(command, "$client_id", clientId);

        var records = new List<ClaimsExperienceRecord>();
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            records.Add(ReadClaimsExperience(reader));
        }

        return records;
    }

    public async Task<List<MonthlyClaimsRecord>> GetMonthlyClaimsAsync(long clientId)
    {
        await InitializeCoreAsync();
        using var connection = await OpenConnectionAsync();
        using var command = CreateCommand(
            connection,
            null,
            "SELECT * FROM MonthlyClaims WHERE client_id = $client_id ORDER BY month DESC, benefit_line");
        AddParameter(command, "$client_id", clientId);

        var records = new List<MonthlyClaimsRecord>();
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            records.Add(ReadMonthlyClaims(reader));
        }

        return records;
    }

    public async Task<List<PlanDesignRecord>> GetPlanDesignAsync(long clientId)
    {
        await InitializeCoreAsync();
        using var connection = await OpenConnectionAsync();
        using var command = CreateCommand(
            connection,
            null,
            "SELECT * FROM PlanDesign WHERE client_id = $client_id ORDER BY benefit_category, benefit_name");
        AddParameter(command, "$client_id", clientId);

        var records = new List<PlanDesignRecord>();
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            records.Add(ReadPlanDesign(reader));
        }

        return records;
    }

    public async Task<List<DocumentRecord>> GetDocumentsAsync(long clientId)
    {
        await InitializeCoreAsync();
        using var connection = await OpenConnectionAsync();
        using var command = CreateCommand(
            connection,
            null,
            "SELECT * FROM Documents WHERE client_id = $client_id ORDER BY processed_date DESC, file_name");
        AddParameter(command, "$client_id", clientId);

        var records = new List<DocumentRecord>();
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            records.Add(ReadDocument(reader));
        }

        return records;
    }

    private async Task InitializeCoreAsync()
    {
        AppPaths.EnsureCreated();
        var folder = Path.GetDirectoryName(DatabasePath);
        if (!string.IsNullOrWhiteSpace(folder))
        {
            Directory.CreateDirectory(folder);
        }

        using var connection = await OpenConnectionAsync();
        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS Clients (
    client_id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_name TEXT NOT NULL COLLATE NOCASE,
    client_number TEXT,
    carrier TEXT,
    industry TEXT,
    province TEXT,
    advisor TEXT,
    effective_date TEXT,
    renewal_month INTEGER,
    status TEXT NOT NULL DEFAULT 'Active',
    date_created TEXT NOT NULL,
    date_updated TEXT NOT NULL
);");

        await EnsureColumnExistsAsync(connection, "Clients", "province", "TEXT");
        await EnsureColumnExistsAsync(connection, "Clients", "advisor", "TEXT");
        await EnsureColumnExistsAsync(connection, "Clients", "renewal_month", "INTEGER");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE UNIQUE INDEX IF NOT EXISTS UX_Clients_ClientName
ON Clients(client_name COLLATE NOCASE);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE UNIQUE INDEX IF NOT EXISTS UX_Clients_ClientNumber
ON Clients(client_number)
WHERE client_number IS NOT NULL AND TRIM(client_number) <> '';");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS Documents (
    document_id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL,
    document_type TEXT NOT NULL,
    module_name TEXT NOT NULL,
    file_name TEXT NOT NULL,
    file_type TEXT,
    file_path TEXT NOT NULL,
    file_hash TEXT NOT NULL,
    carrier TEXT,
    effective_date TEXT,
    period_start TEXT,
    period_end TEXT,
    upload_date TEXT,
    processed_date TEXT NOT NULL,
    processing_status TEXT NOT NULL,
    FOREIGN KEY (client_id) REFERENCES Clients(client_id)
);");

        await EnsureColumnExistsAsync(connection, "Documents", "file_type", "TEXT");
        await EnsureColumnExistsAsync(connection, "Documents", "period_start", "TEXT");
        await EnsureColumnExistsAsync(connection, "Documents", "period_end", "TEXT");
        await EnsureColumnExistsAsync(connection, "Documents", "upload_date", "TEXT");

        // Document identity is scoped to its owning client and document type.
        await ExecuteNonQueryAsync(connection, null, "DROP INDEX IF EXISTS UX_Documents_FileHash;");
        await ExecuteNonQueryAsync(connection, null, @"
CREATE UNIQUE INDEX IF NOT EXISTS UX_Documents_ClientTypeHash
ON Documents(client_id, document_type, file_hash);");
        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_Documents_FileHash
ON Documents(file_hash);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_Documents_Client
ON Documents(client_id, document_type, processed_date);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS PdfExtractionLog (
    log_id INTEGER PRIMARY KEY AUTOINCREMENT,
    document_id INTEGER NOT NULL,
    extraction_method TEXT NOT NULL,
    confidence_score REAL,
    fallback_reason TEXT,
    page_count INTEGER,
    table_count INTEGER,
    processing_time REAL,
    error_message TEXT,
    manual_review_required INTEGER NOT NULL DEFAULT 0,
    date_processed TEXT NOT NULL,
    FOREIGN KEY (document_id) REFERENCES Documents(document_id)
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_PdfExtractionLog_Document
ON PdfExtractionLog(document_id, date_processed);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS Renewals (
    renewal_id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL,
    renewal_year INTEGER NOT NULL,
    carrier TEXT,
    benefit_line TEXT,
    current_rate REAL,
    renewal_rate REAL,
    rate_change_percent REAL,
    premium REAL,
    claims REAL,
    loss_ratio REAL,
    lives INTEGER,
    renewal_status TEXT,
    source_document_id INTEGER,
    date_loaded TEXT NOT NULL,
    FOREIGN KEY (client_id) REFERENCES Clients(client_id),
    FOREIGN KEY (source_document_id) REFERENCES Documents(document_id)
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_Renewals_Client
ON Renewals(client_id, renewal_year, benefit_line);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS Rates (
    rate_id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL,
    carrier TEXT,
    benefit_line TEXT,
    class TEXT,
    current_rate REAL,
    proposed_rate REAL,
    negotiated_rate REAL,
    rate_basis TEXT,
    effective_date TEXT,
    source_document_id INTEGER,
    date_loaded TEXT NOT NULL,
    FOREIGN KEY (client_id) REFERENCES Clients(client_id),
    FOREIGN KEY (source_document_id) REFERENCES Documents(document_id)
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_Rates_Client
ON Rates(client_id, effective_date, benefit_line, class);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS renewal_assumptions (
    assumption_id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL,
    carrier TEXT NOT NULL DEFAULT '',
    benefit_line TEXT NOT NULL DEFAULT '',
    trend REAL,
    target_loss_ratio REAL,
    pooling_charge REAL,
    pooling_threshold REAL,
    ibnr REAL,
    admin_fee REAL,
    risk_charge REAL,
    effective_date TEXT NOT NULL DEFAULT '',
    FOREIGN KEY (client_id) REFERENCES Clients(client_id),
    UNIQUE (client_id, carrier, benefit_line, effective_date)
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_RenewalAssumptions_Client
ON renewal_assumptions(client_id, effective_date, benefit_line);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS renewal_projections (
    projection_id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL,
    benefit_line TEXT NOT NULL DEFAULT '',
    current_premium REAL,
    projected_claims REAL,
    trend_adjustment REAL,
    required_premium REAL,
    proposed_premium REAL,
    negotiated_premium REAL,
    renewal_change_percent REAL,
    savings REAL,
    FOREIGN KEY (client_id) REFERENCES Clients(client_id),
    UNIQUE (client_id, benefit_line)
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_RenewalProjections_Client
ON renewal_projections(client_id, benefit_line);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS renewal_versions (
    version_id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL,
    version_name TEXT NOT NULL,
    source_module TEXT NOT NULL DEFAULT '',
    status TEXT NOT NULL DEFAULT '',
    is_selected INTEGER NOT NULL DEFAULT 0,
    effective_date TEXT,
    source_document_id INTEGER,
    date_created TEXT NOT NULL,
    FOREIGN KEY (client_id) REFERENCES Clients(client_id),
    FOREIGN KEY (source_document_id) REFERENCES Documents(document_id)
);");

        await EnsureColumnExistsAsync(connection, "renewal_versions", "renewal_id", "INTEGER");
        await EnsureColumnExistsAsync(connection, "renewal_versions", "renewal_year", "INTEGER");
        await EnsureColumnExistsAsync(connection, "renewal_versions", "carrier", "TEXT");
        await EnsureColumnExistsAsync(connection, "renewal_versions", "updated_at", "TEXT");
        await EnsureColumnExistsAsync(connection, "renewal_versions", "created_by", "TEXT");
        await EnsureColumnExistsAsync(connection, "renewal_versions", "notes", "TEXT");
        await EnsureRenewalCanonicalPlaceholderColumnsAsync(connection);

        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_RenewalVersions_Client
ON renewal_versions(client_id, effective_date, version_name);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS renewal_rates (
    rate_version_id INTEGER PRIMARY KEY AUTOINCREMENT,
    version_id INTEGER NOT NULL,
    benefit_line TEXT NOT NULL DEFAULT '',
    class TEXT NOT NULL DEFAULT '',
    account_class TEXT NOT NULL DEFAULT '',
    volume REAL,
    lives INTEGER,
    current_rate REAL,
    current_premium REAL,
    proposed_rate REAL,
    proposed_premium REAL,
    version_rate REAL,
    version_premium REAL,
    change_percent REAL,
    rate_basis TEXT NOT NULL DEFAULT '',
    rate_value_basis TEXT NOT NULL DEFAULT 'Carrier-stated',
    client_presentation_rate_basis TEXT NOT NULL DEFAULT 'Carrier-stated',
    current_net_rate REAL,
    current_gross_rate REAL,
    proposed_net_rate REAL,
    proposed_gross_rate REAL,
    version_net_rate REAL,
    version_gross_rate REAL,
    rate_schedule_source_file TEXT NOT NULL DEFAULT '',
    rate_schedule_source_sheet TEXT NOT NULL DEFAULT '',
    rate_schedule_source_row INTEGER,
    source_reference TEXT NOT NULL DEFAULT '',
    FOREIGN KEY (version_id) REFERENCES renewal_versions(version_id)
);");

        await EnsureColumnExistsAsync(connection, "renewal_rates", "renewal_id", "INTEGER");
        await EnsureColumnExistsAsync(connection, "renewal_rates", "client_id", "INTEGER");
        await EnsureColumnExistsAsync(connection, "renewal_rates", "class_name", "TEXT");
        await EnsureColumnExistsAsync(connection, "renewal_rates", "modelled_rate", "REAL");
        await EnsureColumnExistsAsync(connection, "renewal_rates", "negotiated_rate", "REAL");
        await EnsureColumnExistsAsync(connection, "renewal_rates", "modelled_premium", "REAL");
        await EnsureColumnExistsAsync(connection, "renewal_rates", "negotiated_premium", "REAL");
        await EnsureColumnExistsAsync(connection, "renewal_rates", "proposed_change_pct", "REAL");
        await EnsureColumnExistsAsync(connection, "renewal_rates", "modelled_change_pct", "REAL");
        await EnsureColumnExistsAsync(connection, "renewal_rates", "negotiated_change_pct", "REAL");
        await EnsureColumnExistsAsync(connection, "renewal_rates", "savings", "REAL");
        await EnsureColumnExistsAsync(connection, "renewal_rates", "status", "TEXT");
        await EnsureColumnExistsAsync(connection, "renewal_rates", "notes", "TEXT");
        await EnsureColumnExistsAsync(connection, "renewal_rates", "source", "TEXT");
        await EnsureColumnExistsAsync(connection, "renewal_rates", "updated_at", "TEXT");
        await EnsureColumnExistsAsync(connection, "renewal_rates", "rate_value_basis", "TEXT NOT NULL DEFAULT 'Carrier-stated'");
        await EnsureColumnExistsAsync(connection, "renewal_rates", "client_presentation_rate_basis", "TEXT NOT NULL DEFAULT 'Carrier-stated'");
        await EnsureColumnExistsAsync(connection, "renewal_rates", "current_net_rate", "REAL");
        await EnsureColumnExistsAsync(connection, "renewal_rates", "current_gross_rate", "REAL");
        await EnsureColumnExistsAsync(connection, "renewal_rates", "proposed_net_rate", "REAL");
        await EnsureColumnExistsAsync(connection, "renewal_rates", "proposed_gross_rate", "REAL");
        await EnsureColumnExistsAsync(connection, "renewal_rates", "version_net_rate", "REAL");
        await EnsureColumnExistsAsync(connection, "renewal_rates", "version_gross_rate", "REAL");
        await EnsureColumnExistsAsync(connection, "renewal_rates", "rate_schedule_source_file", "TEXT NOT NULL DEFAULT ''");
        await EnsureColumnExistsAsync(connection, "renewal_rates", "rate_schedule_source_sheet", "TEXT NOT NULL DEFAULT ''");
        await EnsureColumnExistsAsync(connection, "renewal_rates", "rate_schedule_source_row", "INTEGER");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_RenewalRates_Version
ON renewal_rates(version_id, benefit_line, class);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS renewal_calculations (
    calculation_id INTEGER PRIMARY KEY AUTOINCREMENT,
    version_id INTEGER NOT NULL,
    source_module TEXT NOT NULL DEFAULT '',
    benefit_line TEXT NOT NULL DEFAULT '',
    current_claims REAL,
    current_premium REAL,
    trended_claims REAL,
    pooling_adjustment REAL,
    ibnr_adjustment REAL,
    retention_admin_fee_adjustment REAL,
    manual_adjustment REAL,
    projected_claims REAL,
    required_premium REAL,
    proposed_premium REAL,
    negotiated_premium REAL,
    required_rate REAL,
    target_loss_ratio REAL,
    projected_loss_ratio REAL,
    renewal_change_percent REAL,
    negotiated_change_percent REAL,
    savings REAL,
    trend REAL,
    FOREIGN KEY (version_id) REFERENCES renewal_versions(version_id)
);");

        await EnsureColumnExistsAsync(connection, "renewal_calculations", "renewal_id", "INTEGER");
        await EnsureColumnExistsAsync(connection, "renewal_calculations", "client_id", "INTEGER");
        await EnsureColumnExistsAsync(connection, "renewal_calculations", "carrier_proposed_premium", "REAL");
        await EnsureColumnExistsAsync(connection, "renewal_calculations", "modelled_premium", "REAL");
        await EnsureColumnExistsAsync(connection, "renewal_calculations", "proposed_change_pct", "REAL");
        await EnsureColumnExistsAsync(connection, "renewal_calculations", "modelled_change_pct", "REAL");
        await EnsureColumnExistsAsync(connection, "renewal_calculations", "negotiated_change_pct", "REAL");
        await EnsureColumnExistsAsync(connection, "renewal_calculations", "calculation_notes", "TEXT");
        await EnsureColumnExistsAsync(connection, "renewal_calculations", "calculated_at", "TEXT");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_RenewalCalculations_Version
ON renewal_calculations(version_id, benefit_line);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS underwriting_assumptions (
    underwriting_assumption_id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL,
    version_id INTEGER,
    carrier TEXT NOT NULL DEFAULT '',
    benefit_line TEXT NOT NULL DEFAULT '',
    trend REAL,
    target_loss_ratio REAL,
    pooling_adjustment REAL,
    ibnr_adjustment REAL,
    retention_admin_fee REAL,
    manual_adjustment REAL,
    effective_date TEXT,
    date_loaded TEXT NOT NULL,
    FOREIGN KEY (client_id) REFERENCES Clients(client_id),
    FOREIGN KEY (version_id) REFERENCES renewal_versions(version_id)
);");

        await EnsureColumnExistsAsync(connection, "renewal_assumptions", "renewal_version_id", "INTEGER");
        await EnsureColumnExistsAsync(connection, "renewal_assumptions", "renewal_id", "INTEGER");
        await EnsureColumnExistsAsync(connection, "renewal_assumptions", "lives", "INTEGER");
        await EnsureColumnExistsAsync(connection, "renewal_assumptions", "volume", "REAL");
        await EnsureColumnExistsAsync(connection, "renewal_assumptions", "credibility_factor", "REAL");
        await EnsureColumnExistsAsync(connection, "renewal_assumptions", "manual_adjustment_pct", "REAL");
        await EnsureColumnExistsAsync(connection, "renewal_assumptions", "updated_at", "TEXT");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_UnderwritingAssumptions_Client
ON underwriting_assumptions(client_id, version_id, benefit_line);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS negotiation_notes (
    note_id INTEGER PRIMARY KEY AUTOINCREMENT,
    renewal_id INTEGER,
    client_id INTEGER NOT NULL,
    renewal_version_id INTEGER,
    round_number INTEGER,
    note_date TEXT,
    carrier_response TEXT,
    consultant_notes TEXT,
    next_action TEXT,
    status TEXT,
    created_at TEXT NOT NULL,
    created_by TEXT,
    FOREIGN KEY (client_id) REFERENCES Clients(client_id),
    FOREIGN KEY (renewal_version_id) REFERENCES renewal_versions(version_id)
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_NegotiationNotes_Renewal
ON negotiation_notes(client_id, renewal_id, renewal_version_id);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS carrier_assumptions (
    carrier_assumption_id INTEGER PRIMARY KEY AUTOINCREMENT,
    carrier TEXT NOT NULL DEFAULT '',
    benefit_line TEXT NOT NULL DEFAULT '',
    assumption_name TEXT NOT NULL DEFAULT '',
    assumption_value REAL,
    effective_date TEXT,
    date_loaded TEXT NOT NULL
);");

        await EnsureColumnExistsAsync(connection, "carrier_assumptions", "block_name", "TEXT NOT NULL DEFAULT ''");
        await EnsureColumnExistsAsync(connection, "carrier_assumptions", "funding_type", "TEXT NOT NULL DEFAULT ''");
        await EnsureColumnExistsAsync(connection, "carrier_assumptions", "target_loss_ratio", "REAL");
        await EnsureColumnExistsAsync(connection, "carrier_assumptions", "tlr_basis", "TEXT NOT NULL DEFAULT ''");
        await EnsureColumnExistsAsync(connection, "carrier_assumptions", "premium_tax", "REAL");
        await EnsureColumnExistsAsync(connection, "carrier_assumptions", "commission", "REAL");
        await EnsureColumnExistsAsync(connection, "carrier_assumptions", "admin_margin", "REAL");
        await EnsureColumnExistsAsync(connection, "carrier_assumptions", "risk_margin", "REAL");
        await EnsureColumnExistsAsync(connection, "carrier_assumptions", "trend", "REAL");
        await EnsureColumnExistsAsync(connection, "carrier_assumptions", "credibility_factor", "REAL");
        await EnsureColumnExistsAsync(connection, "carrier_assumptions", "pooling_threshold", "REAL");
        await EnsureColumnExistsAsync(connection, "carrier_assumptions", "pooling_charge", "REAL");
        await EnsureColumnExistsAsync(connection, "carrier_assumptions", "reserve_ibnr", "REAL");
        await EnsureColumnExistsAsync(connection, "carrier_assumptions", "status", "TEXT NOT NULL DEFAULT 'Active'");
        await EnsureColumnExistsAsync(connection, "carrier_assumptions", "notes", "TEXT NOT NULL DEFAULT ''");
        await EnsureColumnExistsAsync(connection, "carrier_assumptions", "updated_at", "TEXT");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_CarrierAssumptions_Library
ON carrier_assumptions(carrier, block_name, benefit_line, effective_date);");

        await EnsureCanadaLifeIndependentContractorConditionsAsync(connection);

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS plan_design_factors (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    carrier_name TEXT NOT NULL,
    benefit_category TEXT NOT NULL,
    benefit_line TEXT NOT NULL,
    current_design TEXT,
    proposed_design TEXT,
    change_type TEXT NOT NULL,
    factor REAL NOT NULL,
    factor_direction TEXT,
    effective_date TEXT,
    expiry_date TEXT,
    source_document TEXT,
    notes TEXT,
    confidence_level TEXT,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP,
    updated_at TEXT DEFAULT CURRENT_TIMESTAMP
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_PlanDesignFactors_Lookup
ON plan_design_factors(carrier_name, benefit_line, current_design, proposed_design, effective_date, expiry_date);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_PlanDesignFactors_Filters
ON plan_design_factors(benefit_category, change_type, confidence_level);");

        await SeedPlanDesignFactorPlaceholdersIfEmptyAsync(connection);

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS quote_scenarios (
    quote_scenario_id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL,
    version_id INTEGER,
    scenario_name TEXT NOT NULL DEFAULT '',
    carrier TEXT NOT NULL DEFAULT '',
    plan_design_summary TEXT NOT NULL DEFAULT '',
    status TEXT NOT NULL DEFAULT '',
    date_created TEXT NOT NULL,
    FOREIGN KEY (client_id) REFERENCES Clients(client_id),
    FOREIGN KEY (version_id) REFERENCES renewal_versions(version_id)
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_QuoteScenarios_Client
ON quote_scenarios(client_id, scenario_name);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS source_files (
    source_file_id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL,
    document_id INTEGER,
    module_name TEXT NOT NULL DEFAULT '',
    file_name TEXT NOT NULL DEFAULT '',
    file_path TEXT NOT NULL DEFAULT '',
    file_hash TEXT NOT NULL DEFAULT '',
    date_loaded TEXT NOT NULL,
    FOREIGN KEY (client_id) REFERENCES Clients(client_id),
    FOREIGN KEY (document_id) REFERENCES Documents(document_id)
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS demographics (
    demographic_id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL,
    benefit_line TEXT NOT NULL DEFAULT '',
    lives INTEGER,
    volume REAL,
    average_age REAL,
    effective_date TEXT,
    source_document_id INTEGER,
    date_loaded TEXT NOT NULL,
    FOREIGN KEY (client_id) REFERENCES Clients(client_id),
    FOREIGN KEY (source_document_id) REFERENCES Documents(document_id)
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS claims_categories (
    claims_category_id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL,
    benefit_line TEXT NOT NULL DEFAULT '',
    category TEXT NOT NULL DEFAULT '',
    period_label TEXT NOT NULL DEFAULT '',
    amount REAL,
    percent_of_total REAL,
    source_document_id INTEGER,
    date_loaded TEXT NOT NULL,
    FOREIGN KEY (client_id) REFERENCES Clients(client_id),
    FOREIGN KEY (source_document_id) REFERENCES Documents(document_id)
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS benefit_experience_history (
    benefit_experience_id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL,
    source_document_id INTEGER,
    benefit TEXT NOT NULL DEFAULT '',
    experience_period TEXT NOT NULL DEFAULT '',
    premium REAL,
    paid_claims REAL,
    incurred_claims REAL,
    incurred_loss_ratio REAL,
    target_loss_ratio REAL,
    annual_trend_factor REAL,
    credibility REAL,
    fee_guide_adjustment REAL,
    source_reference TEXT NOT NULL DEFAULT '',
    date_loaded TEXT NOT NULL,
    FOREIGN KEY (client_id) REFERENCES Clients(client_id),
    FOREIGN KEY (source_document_id) REFERENCES Documents(document_id)
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS renewal_factors (
    renewal_factor_id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL,
    source_document_id INTEGER,
    benefit TEXT NOT NULL DEFAULT '',
    target_loss_ratio REAL,
    trend_factor REAL,
    credibility REAL,
    fee_guide_adjustment REAL,
    pooling_charge REAL,
    pooling_level REAL,
    first_dollar_out_of_country TEXT,
    source_reference TEXT NOT NULL DEFAULT '',
    date_loaded TEXT NOT NULL,
    FOREIGN KEY (client_id) REFERENCES Clients(client_id),
    FOREIGN KEY (source_document_id) REFERENCES Documents(document_id)
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS pooling_arrangements (
    pooling_arrangement_id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL,
    source_document_id INTEGER,
    benefit TEXT NOT NULL DEFAULT '',
    pooling_level REAL,
    pooling_charge REAL,
    first_dollar_out_of_country TEXT,
    source_reference TEXT NOT NULL DEFAULT '',
    date_loaded TEXT NOT NULL,
    FOREIGN KEY (client_id) REFERENCES Clients(client_id),
    FOREIGN KEY (source_document_id) REFERENCES Documents(document_id)
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS demographic_summary (
    demographic_summary_id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL,
    source_document_id INTEGER,
    benefit TEXT NOT NULL DEFAULT '',
    period_label TEXT NOT NULL DEFAULT '',
    lives INTEGER,
    volume REAL,
    average_age REAL,
    combined_average_age REAL,
    date_loaded TEXT NOT NULL,
    FOREIGN KEY (client_id) REFERENCES Clients(client_id),
    FOREIGN KEY (source_document_id) REFERENCES Documents(document_id)
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS demographics_by_age_gender (
    demographic_age_gender_id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL,
    source_document_id INTEGER,
    benefit TEXT NOT NULL DEFAULT '',
    period_label TEXT NOT NULL DEFAULT '',
    age_band TEXT NOT NULL DEFAULT '',
    gender TEXT NOT NULL DEFAULT '',
    lives INTEGER,
    volume REAL,
    source_reference TEXT NOT NULL DEFAULT '',
    date_loaded TEXT NOT NULL,
    FOREIGN KEY (client_id) REFERENCES Clients(client_id),
    FOREIGN KEY (source_document_id) REFERENCES Documents(document_id)
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS underwriting_reference_rate_versions (
    rate_version_id INTEGER PRIMARY KEY AUTOINCREMENT,
    rate_type TEXT NOT NULL,
    version_name TEXT NOT NULL,
    unit_basis TEXT NOT NULL,
    source_name TEXT NOT NULL,
    source_sheet TEXT NOT NULL,
    effective_date TEXT,
    is_active INTEGER NOT NULL DEFAULT 1,
    date_created TEXT NOT NULL,
    UNIQUE(rate_type, version_name)
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS underwriting_reference_rates (
    reference_rate_id INTEGER PRIMARY KEY AUTOINCREMENT,
    rate_version_id INTEGER NOT NULL,
    age_band TEXT NOT NULL,
    gender TEXT NOT NULL,
    rate REAL NOT NULL,
    display_order INTEGER NOT NULL DEFAULT 0,
    FOREIGN KEY (rate_version_id) REFERENCES underwriting_reference_rate_versions(rate_version_id),
    UNIQUE(rate_version_id, age_band, gender)
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_UnderwritingReferenceRates_Active
ON underwriting_reference_rate_versions(rate_type, is_active, version_name);");

        await SeedUnderwritingReferenceRatesAsync(connection);

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS claims_distribution (
    claims_distribution_id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL,
    source_document_id INTEGER,
    benefit TEXT NOT NULL DEFAULT '',
    service_category TEXT NOT NULL DEFAULT '',
    period_label TEXT NOT NULL DEFAULT '',
    amount_paid REAL,
    percent_of_total REAL,
    norm_percent_of_total REAL,
    number_of_occurrences INTEGER,
    occurrence_percent_of_total REAL,
    occurrence_norm_percent_of_total REAL,
    average_amount_paid REAL,
    source_reference TEXT NOT NULL DEFAULT '',
    date_loaded TEXT NOT NULL,
    FOREIGN KEY (client_id) REFERENCES Clients(client_id),
    FOREIGN KEY (source_document_id) REFERENCES Documents(document_id)
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS claims_distribution_reports (
    claims_distribution_report_id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL,
    source_document_id INTEGER,
    carrier_name TEXT NOT NULL DEFAULT '',
    policy_number TEXT NOT NULL DEFAULT '',
    client_name TEXT NOT NULL DEFAULT '',
    report_period TEXT NOT NULL DEFAULT '',
    period_start TEXT NOT NULL DEFAULT '',
    period_end TEXT NOT NULL DEFAULT '',
    benefit_type TEXT NOT NULL DEFAULT '',
    division TEXT NOT NULL DEFAULT '',
    class TEXT NOT NULL DEFAULT '',
    report_created_date TEXT NOT NULL DEFAULT '',
    source_file TEXT NOT NULL DEFAULT '',
    table_found INTEGER NOT NULL DEFAULT 0,
    paramedical_found INTEGER NOT NULL DEFAULT 0,
    total_row_found INTEGER NOT NULL DEFAULT 0,
    total_amount_paid REAL,
    sum_of_service_amounts REAL,
    amounts_balance_to_total INTEGER NOT NULL DEFAULT 0,
    warnings TEXT NOT NULL DEFAULT '',
    date_loaded TEXT NOT NULL,
    FOREIGN KEY (client_id) REFERENCES Clients(client_id),
    FOREIGN KEY (source_document_id) REFERENCES Documents(document_id)
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS claims_distribution_rows (
    claims_distribution_row_id INTEGER PRIMARY KEY AUTOINCREMENT,
    claims_distribution_report_id INTEGER,
    client_id INTEGER NOT NULL,
    source_document_id INTEGER,
    benefit TEXT NOT NULL DEFAULT '',
    service_category TEXT NOT NULL DEFAULT '',
    service_description_original TEXT NOT NULL DEFAULT '',
    amount_paid REAL,
    percent_paid REAL,
    norm_percent_paid REAL,
    occurrences INTEGER,
    percent_occurrences REAL,
    norm_percent_occurrences REAL,
    avg_claim REAL,
    source_reference TEXT NOT NULL DEFAULT '',
    carrier_name TEXT NOT NULL DEFAULT '',
    policy_number TEXT NOT NULL DEFAULT '',
    report_period TEXT NOT NULL DEFAULT '',
    division TEXT NOT NULL DEFAULT '',
    class TEXT NOT NULL DEFAULT '',
    date_loaded TEXT NOT NULL,
    FOREIGN KEY (claims_distribution_report_id) REFERENCES claims_distribution_reports(claims_distribution_report_id),
    FOREIGN KEY (client_id) REFERENCES Clients(client_id),
    FOREIGN KEY (source_document_id) REFERENCES Documents(document_id)
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_ClaimsDistributionRows_Client
ON claims_distribution_rows(client_id, source_document_id, benefit, service_category);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS plan_design_features (
    plan_design_feature_id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL,
    source_document_id INTEGER,
    benefit TEXT NOT NULL DEFAULT '',
    category TEXT NOT NULL DEFAULT '',
    feature_name TEXT NOT NULL DEFAULT '',
    feature_value TEXT NOT NULL DEFAULT '',
    coinsurance TEXT,
    deductible TEXT,
    maximum TEXT,
    frequency TEXT,
    termination_age TEXT,
    class_number TEXT,
    class_name TEXT,
    source_page INTEGER,
    source_text TEXT,
    confidence_score REAL,
    date_loaded TEXT NOT NULL,
    FOREIGN KEY (client_id) REFERENCES Clients(client_id),
    FOREIGN KEY (source_document_id) REFERENCES Documents(document_id)
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS extraction_audit_log (
    audit_id INTEGER PRIMARY KEY AUTOINCREMENT,
    source_document_id INTEGER,
    source_file TEXT NOT NULL DEFAULT '',
    source_page INTEGER,
    section_type TEXT NOT NULL DEFAULT '',
    field_name TEXT NOT NULL DEFAULT '',
    extracted_value TEXT NOT NULL DEFAULT '',
    source_text TEXT NOT NULL DEFAULT '',
    confidence_score REAL,
    validation_status TEXT NOT NULL DEFAULT '',
    validation_message TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL,
    FOREIGN KEY (source_document_id) REFERENCES Documents(document_id)
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS ClaimsExperience (
    claims_id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL,
    experience_period_start TEXT,
    experience_period_end TEXT,
    benefit_line TEXT,
    premium REAL,
    paid_claims REAL,
    incurred_claims REAL,
    loss_ratio REAL,
    lives INTEGER,
    source_document_id INTEGER,
    date_loaded TEXT NOT NULL,
    FOREIGN KEY (client_id) REFERENCES Clients(client_id),
    FOREIGN KEY (source_document_id) REFERENCES Documents(document_id)
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_ClaimsExperience_Client
ON ClaimsExperience(client_id, experience_period_end, benefit_line);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS MonthlyClaims (
    claim_month_id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL,
    carrier TEXT,
    benefit_line TEXT,
    month TEXT,
    lives INTEGER,
    premium REAL,
    paid_claims REAL,
    incurred_claims REAL,
    claims_plus_fees REAL,
    loss_ratio REAL,
    source_document_id INTEGER,
    date_loaded TEXT NOT NULL,
    FOREIGN KEY (client_id) REFERENCES Clients(client_id),
    FOREIGN KEY (source_document_id) REFERENCES Documents(document_id)
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_MonthlyClaims_Client
ON MonthlyClaims(client_id, month, benefit_line);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS PlanDesign (
    plan_design_id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL,
    class_name TEXT NOT NULL DEFAULT '',
    benefit_category TEXT,
    benefit_name TEXT,
    coverage_description TEXT,
    reimbursement_percent REAL,
    maximum_amount TEXT,
    frequency TEXT,
    waiting_period TEXT,
    termination_age TEXT,
    booklet_source_file_name TEXT NOT NULL DEFAULT '',
    source_section TEXT NOT NULL DEFAULT '',
    notes TEXT NOT NULL DEFAULT '',
    source_document_id INTEGER,
    source_page INTEGER,
    date_loaded TEXT NOT NULL,
    FOREIGN KEY (client_id) REFERENCES Clients(client_id),
    FOREIGN KEY (source_document_id) REFERENCES Documents(document_id)
);");

        await EnsureColumnExistsAsync(connection, "PlanDesign", "class_name", "TEXT NOT NULL DEFAULT ''");
        await EnsureColumnExistsAsync(connection, "PlanDesign", "booklet_source_file_name", "TEXT NOT NULL DEFAULT ''");
        await EnsureColumnExistsAsync(connection, "PlanDesign", "source_section", "TEXT NOT NULL DEFAULT ''");
        await EnsureColumnExistsAsync(connection, "PlanDesign", "notes", "TEXT NOT NULL DEFAULT ''");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_PlanDesign_Client
ON PlanDesign(client_id, class_name, benefit_category, benefit_name);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS class_comparison_preferences (
    client_id INTEGER PRIMARY KEY,
    baseline_class_name TEXT NOT NULL DEFAULT '',
    selected_class_names TEXT NOT NULL DEFAULT '[]',
    active_tab TEXT NOT NULL DEFAULT 'Class Comparison',
    key_benefits_only INTEGER NOT NULL DEFAULT 0,
    highlight_differences INTEGER NOT NULL DEFAULT 1,
    collapsed_sections TEXT NOT NULL DEFAULT '[]',
    recommended_class_name TEXT NOT NULL DEFAULT '',
    updated_at TEXT NOT NULL,
    updated_by TEXT NOT NULL DEFAULT '',
    FOREIGN KEY (client_id) REFERENCES Clients(client_id)
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS class_comparison_notes (
    note_id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL,
    class_name TEXT NOT NULL DEFAULT '',
    note_text TEXT NOT NULL,
    created_at TEXT NOT NULL,
    created_by TEXT NOT NULL DEFAULT '',
    updated_at TEXT NOT NULL,
    updated_by TEXT NOT NULL DEFAULT '',
    FOREIGN KEY (client_id) REFERENCES Clients(client_id)
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_ClassComparisonNotes_Client
ON class_comparison_notes(client_id, created_at DESC);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS class_comparison_audit (
    audit_id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL,
    class_name TEXT NOT NULL DEFAULT '',
    field_name TEXT NOT NULL DEFAULT '',
    previous_value TEXT NOT NULL DEFAULT '',
    new_value TEXT NOT NULL DEFAULT '',
    reason TEXT NOT NULL DEFAULT '',
    source_document TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL,
    created_by TEXT NOT NULL DEFAULT '',
    FOREIGN KEY (client_id) REFERENCES Clients(client_id)
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_ClassComparisonAudit_Client
ON class_comparison_audit(client_id, created_at DESC);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS billing_option_premiums (
    billing_option_premium_id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL,
    source_document_id INTEGER,
    billing_period TEXT NOT NULL DEFAULT '',
    province TEXT NOT NULL DEFAULT '',
    benefit TEXT NOT NULL DEFAULT '',
    class_code TEXT NOT NULL DEFAULT '',
    option_label TEXT NOT NULL DEFAULT '',
    dependent_code TEXT NOT NULL DEFAULT '',
    lives INTEGER,
    volume REAL,
    monthly_premium REAL,
    premium_adjustments REAL,
    total_premium_adjustments REAL,
    source_page INTEGER,
    source_file TEXT NOT NULL DEFAULT '',
    confidence REAL,
    date_loaded TEXT NOT NULL,
    FOREIGN KEY (client_id) REFERENCES Clients(client_id),
    FOREIGN KEY (source_document_id) REFERENCES Documents(document_id)
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_BillingOptionPremiums_Client
ON billing_option_premiums(client_id, billing_period);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_BillingOptionPremiums_Detail
ON billing_option_premiums(client_id, province, benefit, class_code, billing_period);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS billing_province_summary (
    billing_province_summary_id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL,
    source_document_id INTEGER,
    billing_period TEXT NOT NULL DEFAULT '',
    province TEXT NOT NULL DEFAULT '',
    total_lives INTEGER,
    total_monthly_premium REAL,
    estimated_annual_premium REAL,
    total_tax REAL,
    total_premium_with_tax REAL,
    source_file TEXT NOT NULL DEFAULT '',
    date_loaded TEXT NOT NULL,
    FOREIGN KEY (client_id) REFERENCES Clients(client_id),
    FOREIGN KEY (source_document_id) REFERENCES Documents(document_id)
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_BillingProvinceSummary_Client
ON billing_province_summary(client_id, province, billing_period);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS billing_option_summary (
    billing_option_summary_id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL,
    source_document_id INTEGER,
    billing_period TEXT NOT NULL DEFAULT '',
    province TEXT NOT NULL DEFAULT '',
    benefit TEXT NOT NULL DEFAULT '',
    class_code TEXT NOT NULL DEFAULT '',
    option_label TEXT NOT NULL DEFAULT '',
    lives INTEGER,
    monthly_premium REAL,
    estimated_annual_premium REAL,
    source_file TEXT NOT NULL DEFAULT '',
    date_loaded TEXT NOT NULL,
    FOREIGN KEY (client_id) REFERENCES Clients(client_id),
    FOREIGN KEY (source_document_id) REFERENCES Documents(document_id)
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_BillingOptionSummary_Detail
ON billing_option_summary(client_id, province, benefit, class_code, billing_period);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS ClassMapping (
    class_mapping_id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL,
    carrier TEXT NOT NULL DEFAULT '',
    class_code TEXT NOT NULL DEFAULT '',
    option_label TEXT NOT NULL DEFAULT '',
    first_seen_billing_period TEXT NOT NULL DEFAULT '',
    last_seen_billing_period TEXT NOT NULL DEFAULT '',
    source_file TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    FOREIGN KEY (client_id) REFERENCES Clients(client_id),
    UNIQUE (client_id, carrier, class_code)
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_ClassMapping_Client
ON ClassMapping(client_id, carrier, class_code);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS BillingReconciliationResults (
    billing_reconciliation_result_id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL,
    source_document_id INTEGER,
    billing_period TEXT NOT NULL DEFAULT '',
    source_file TEXT NOT NULL DEFAULT '',
    reconciliation_level TEXT NOT NULL DEFAULT '',
    province TEXT NOT NULL DEFAULT '',
    benefit TEXT NOT NULL DEFAULT '',
    class_code TEXT NOT NULL DEFAULT '',
    pdf_lives INTEGER,
    calculated_lives INTEGER,
    lives_difference INTEGER,
    pdf_premium REAL,
    calculated_premium REAL,
    premium_difference REAL,
    pdf_tax REAL,
    calculated_tax REAL,
    tax_difference REAL,
    pdf_total_with_tax REAL,
    calculated_total_with_tax REAL,
    total_difference REAL,
    status TEXT NOT NULL DEFAULT '',
    warning_message TEXT NOT NULL DEFAULT '',
    source_page INTEGER,
    created_at TEXT NOT NULL,
    FOREIGN KEY (client_id) REFERENCES Clients(client_id),
    FOREIGN KEY (source_document_id) REFERENCES Documents(document_id)
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_BillingReconciliationResults_Client
ON BillingReconciliationResults(client_id, billing_period, province, benefit, class_code);");

        await CarrierDatabaseMigration.ApplyAsync(connection);
        await BonusProgramDatabaseMigration.ApplyAsync(connection);
        await EnsureManulifeBillingSchemaAsync(connection);
        await EnsureManulifeExcelBillingSchemaAsync(connection);
        await EnsureGroupHealthBillingSchemaAsync(connection);
        await EnsureSourceTrackingSchemaAsync(connection);
        await BenefitAdministrationDatabaseMigration.ApplyAsync(connection);
        await PlanBenefitRuleDatabaseMigration.ApplyAsync(connection);
        await CanonicalClaimsDatabaseMigration.ApplyAsync(connection);
        await EnsureFlexibleBenefitSchemaAsync(connection);
        await EnsureControlledBookletReprocessingSchemaAsync(connection);
        await EnsureClientPolicySchemaAsync(connection);
        await EnsureAsoFinancialActivitySchemaAsync(connection);
        await FinancialArrangementDatabaseMigration.ApplyAsync(connection);
    }

    private static async Task EnsureCanadaLifeIndependentContractorConditionsAsync(SqliteConnection connection)
    {
        const string condition = "Canada Life underwriting condition for independent contractors (50% of group): the client must keep enrollment records current and accurate, and each covered independent contractor must have a minimum one-year contract with the client.";

        await ExecuteNonQueryAsync(
            connection,
            null,
            @"
UPDATE carrier_assumptions
SET notes = CASE
        WHEN TRIM(COALESCE(notes, '')) = '' THEN $condition
        ELSE TRIM(notes) || ' ' || $condition
    END,
    updated_at = $updated_at
WHERE carrier COLLATE NOCASE = 'Canada Life'
  AND COALESCE(notes, '') NOT LIKE '%' || $condition_marker || '%';",
            command =>
            {
                AddParameter(command, "$condition", condition);
                AddParameter(command, "$condition_marker", "Canada Life underwriting condition for independent contractors");
                AddParameter(command, "$updated_at", DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture));
            });
    }

    private static async Task EnsureManulifeBillingSchemaAsync(SqliteConnection connection)
    {
        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS manulife_billing_statement_versions (
    statement_version_id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL,
    source_document_id INTEGER NOT NULL,
    import_batch_id INTEGER,
    carrier TEXT NOT NULL DEFAULT 'Manulife',
    billing_period TEXT NOT NULL DEFAULT '',
    period_start TEXT,
    period_end TEXT,
    source_file_name TEXT NOT NULL DEFAULT '',
    source_file_path TEXT NOT NULL DEFAULT '',
    source_file_hash TEXT NOT NULL DEFAULT '',
    version_number INTEGER NOT NULL DEFAULT 1,
    import_status TEXT NOT NULL DEFAULT 'Needs Review',
    is_replaced INTEGER NOT NULL DEFAULT 0,
    replaces_statement_version_id INTEGER,
    replaced_by_statement_version_id INTEGER,
    created_at TEXT NOT NULL,
    created_by TEXT NOT NULL DEFAULT '',
    FOREIGN KEY (client_id) REFERENCES Clients(client_id),
    FOREIGN KEY (source_document_id) REFERENCES Documents(document_id),
    FOREIGN KEY (replaces_statement_version_id) REFERENCES manulife_billing_statement_versions(statement_version_id),
    FOREIGN KEY (replaced_by_statement_version_id) REFERENCES manulife_billing_statement_versions(statement_version_id),
    UNIQUE (client_id, carrier, billing_period, version_number)
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_ManulifeBillingStatementVersions_ClientPeriod
ON manulife_billing_statement_versions(client_id, carrier, billing_period, is_replaced, version_number);");
        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_ManulifeBillingStatementVersions_Hash
ON manulife_billing_statement_versions(client_id, carrier, source_file_hash);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS manulife_billing_version_audit (
    audit_id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL,
    replaced_statement_version_id INTEGER NOT NULL,
    replacement_statement_version_id INTEGER NOT NULL,
    action TEXT NOT NULL DEFAULT 'ReplaceExistingVersion',
    performed_by TEXT NOT NULL DEFAULT '',
    performed_at TEXT NOT NULL,
    FOREIGN KEY (client_id) REFERENCES Clients(client_id),
    FOREIGN KEY (replaced_statement_version_id) REFERENCES manulife_billing_statement_versions(statement_version_id),
    FOREIGN KEY (replacement_statement_version_id) REFERENCES manulife_billing_statement_versions(statement_version_id)
);");

        await EnsureColumnExistsAsync(connection, "billing_option_premiums", "statement_version_id", "INTEGER");
        await EnsureColumnExistsAsync(connection, "billing_province_summary", "statement_version_id", "INTEGER");
        await EnsureColumnExistsAsync(connection, "billing_option_summary", "statement_version_id", "INTEGER");
        await EnsureColumnExistsAsync(connection, "BillingReconciliationResults", "statement_version_id", "INTEGER");
        await EnsureColumnExistsAsync(connection, "BillingReconciliationResults", "lives_status", "TEXT NOT NULL DEFAULT 'Not Available'");
        await EnsureColumnExistsAsync(connection, "BillingReconciliationResults", "lives_warning_message", "TEXT NOT NULL DEFAULT ''");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_BillingOptionPremiums_StatementVersion
ON billing_option_premiums(statement_version_id);");
        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_BillingProvinceSummary_StatementVersion
ON billing_province_summary(statement_version_id);");
        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_BillingOptionSummary_StatementVersion
ON billing_option_summary(statement_version_id);");
        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_BillingReconciliationResults_StatementVersion
ON BillingReconciliationResults(statement_version_id);");
    }

    private static async Task EnsureGroupHealthBillingSchemaAsync(SqliteConnection connection)
    {
        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS BillingStatements (
    billing_statement_id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL,
    source_document_id INTEGER,
    carrier TEXT NOT NULL DEFAULT 'GroupHEALTH',
    statement_month INTEGER NOT NULL,
    statement_year INTEGER NOT NULL,
    statement_date TEXT,
    import_date TEXT NOT NULL,
    source_file_name TEXT NOT NULL DEFAULT '',
    source_file_path TEXT NOT NULL DEFAULT '',
    source_file_hash TEXT NOT NULL DEFAULT '',
    import_version INTEGER NOT NULL DEFAULT 1,
    import_status TEXT NOT NULL DEFAULT 'Draft',
    billing_integrity_score REAL NOT NULL DEFAULT 0,
    employee_count INTEGER NOT NULL DEFAULT 0,
    total_employee_premium REAL,
    total_employer_premium REAL,
    total_premium REAL,
    total_gst_hst REAL,
    statement_total REAL,
    is_replaced INTEGER NOT NULL DEFAULT 0,
    replaced_at TEXT,
    is_deleted INTEGER NOT NULL DEFAULT 0,
    deleted_at TEXT,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    FOREIGN KEY (client_id) REFERENCES Clients(client_id),
    FOREIGN KEY (source_document_id) REFERENCES Documents(document_id),
    UNIQUE (client_id, carrier, statement_month, statement_year, source_file_hash)
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_BillingStatements_ClientMonth
ON BillingStatements(client_id, carrier, statement_year, statement_month, import_version);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS BillingStatementRows (
    billing_statement_row_id INTEGER PRIMARY KEY AUTOINCREMENT,
    billing_statement_id INTEGER NOT NULL,
    client_id INTEGER NOT NULL,
    class_code TEXT NOT NULL DEFAULT '',
    employee_id TEXT NOT NULL DEFAULT '',
    certificate_id TEXT NOT NULL DEFAULT '',
    employee_name TEXT NOT NULL DEFAULT '',
    annual_earnings REAL,
    cost_division TEXT NOT NULL DEFAULT '',
    employment_class TEXT NOT NULL DEFAULT '',
    retro_month TEXT NOT NULL DEFAULT '',
    life_volume REAL,
    life_premium REAL,
    adnd_volume REAL,
    adnd_premium REAL,
    ltd_volume REAL,
    ltd_premium REAL,
    health_tier TEXT NOT NULL DEFAULT '',
    health_premium REAL,
    dental_tier TEXT NOT NULL DEFAULT '',
    dental_premium REAL,
    efap_premium REAL,
    virtual_healthcare_premium REAL,
    flex_admin_fee REAL,
    member_adnd_volume REAL,
    member_adnd_premium REAL,
    member_adnd_sp_volume REAL,
    member_adnd_sp_premium REAL,
    member_adnd_ch_volume REAL,
    member_adnd_ch_premium REAL,
    employee_premium REAL,
    employer_premium REAL,
    total_premium REAL,
    gst_hst REAL,
    retro_premium REAL,
    retro_gst_hst REAL,
    row_total REAL,
    source_page INTEGER,
    source_line INTEGER,
    validation_status TEXT NOT NULL DEFAULT '',
    validation_flags TEXT NOT NULL DEFAULT '',
    source_text TEXT NOT NULL DEFAULT '',
    date_loaded TEXT NOT NULL,
    FOREIGN KEY (billing_statement_id) REFERENCES BillingStatements(billing_statement_id),
    FOREIGN KEY (client_id) REFERENCES Clients(client_id)
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_BillingStatementRows_Statement
ON BillingStatementRows(billing_statement_id, class_code, certificate_id);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS BillingStatementTotals (
    billing_statement_total_id INTEGER PRIMARY KEY AUTOINCREMENT,
    billing_statement_id INTEGER NOT NULL,
    client_id INTEGER NOT NULL,
    total_level TEXT NOT NULL DEFAULT '',
    class_code TEXT NOT NULL DEFAULT '',
    benefit TEXT NOT NULL DEFAULT '',
    total_type TEXT NOT NULL DEFAULT '',
    lives INTEGER,
    volume REAL,
    premium REAL,
    gst_hst REAL,
    total_premium REAL,
    source_page INTEGER,
    date_loaded TEXT NOT NULL,
    FOREIGN KEY (billing_statement_id) REFERENCES BillingStatements(billing_statement_id),
    FOREIGN KEY (client_id) REFERENCES Clients(client_id)
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_BillingStatementTotals_Statement
ON BillingStatementTotals(billing_statement_id, total_level, class_code, benefit);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS ClientRateTables (
    client_rate_table_id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL,
    billing_statement_id INTEGER,
    carrier TEXT NOT NULL DEFAULT '',
    effective_month INTEGER NOT NULL,
    effective_year INTEGER NOT NULL,
    source TEXT NOT NULL DEFAULT '',
    status TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    FOREIGN KEY (client_id) REFERENCES Clients(client_id),
    FOREIGN KEY (billing_statement_id) REFERENCES BillingStatements(billing_statement_id)
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_ClientRateTables_Client
ON ClientRateTables(client_id, carrier, effective_year, effective_month);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS ClientRateTableLines (
    client_rate_table_line_id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_rate_table_id INTEGER NOT NULL,
    billing_statement_id INTEGER NOT NULL,
    client_id INTEGER NOT NULL,
    carrier TEXT NOT NULL DEFAULT '',
    effective_month INTEGER NOT NULL,
    effective_year INTEGER NOT NULL,
    benefit TEXT NOT NULL DEFAULT '',
    class_code TEXT NOT NULL DEFAULT '',
    tier TEXT NOT NULL DEFAULT '',
    volume_basis TEXT NOT NULL DEFAULT '',
    employee_rate REAL,
    employer_rate REAL,
    total_rate REAL,
    calculated_rate REAL,
    expected_rate REAL,
    rate_status TEXT NOT NULL DEFAULT '',
    is_calculated INTEGER NOT NULL DEFAULT 1,
    source TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL,
    FOREIGN KEY (client_rate_table_id) REFERENCES ClientRateTables(client_rate_table_id),
    FOREIGN KEY (billing_statement_id) REFERENCES BillingStatements(billing_statement_id),
    FOREIGN KEY (client_id) REFERENCES Clients(client_id)
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_ClientRateTableLines_Statement
ON ClientRateTableLines(billing_statement_id, benefit, class_code, tier);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS BillingValidationResults (
    billing_validation_result_id INTEGER PRIMARY KEY AUTOINCREMENT,
    billing_statement_id INTEGER NOT NULL,
    billing_statement_row_id INTEGER,
    client_id INTEGER NOT NULL,
    validation_level TEXT NOT NULL DEFAULT '',
    check_name TEXT NOT NULL DEFAULT '',
    status TEXT NOT NULL DEFAULT '',
    message TEXT NOT NULL DEFAULT '',
    expected_amount REAL,
    actual_amount REAL,
    difference_amount REAL,
    benefit TEXT NOT NULL DEFAULT '',
    class_code TEXT NOT NULL DEFAULT '',
    employee_id TEXT NOT NULL DEFAULT '',
    certificate_id TEXT NOT NULL DEFAULT '',
    source_page INTEGER,
    created_at TEXT NOT NULL,
    FOREIGN KEY (billing_statement_id) REFERENCES BillingStatements(billing_statement_id),
    FOREIGN KEY (billing_statement_row_id) REFERENCES BillingStatementRows(billing_statement_row_id),
    FOREIGN KEY (client_id) REFERENCES Clients(client_id)
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_BillingValidationResults_Statement
ON BillingValidationResults(billing_statement_id, status, validation_level);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS BillingEmployeeChangeLog (
    billing_employee_change_log_id INTEGER PRIMARY KEY AUTOINCREMENT,
    billing_statement_id INTEGER NOT NULL,
    client_id INTEGER NOT NULL,
    carrier TEXT NOT NULL DEFAULT '',
    statement_month INTEGER NOT NULL,
    statement_year INTEGER NOT NULL,
    employee_id TEXT NOT NULL DEFAULT '',
    certificate_id TEXT NOT NULL DEFAULT '',
    employee_name TEXT NOT NULL DEFAULT '',
    change_type TEXT NOT NULL DEFAULT '',
    field_name TEXT NOT NULL DEFAULT '',
    previous_value TEXT NOT NULL DEFAULT '',
    current_value TEXT NOT NULL DEFAULT '',
    difference_amount REAL,
    change_percent REAL,
    severity TEXT NOT NULL DEFAULT '',
    source TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL,
    FOREIGN KEY (billing_statement_id) REFERENCES BillingStatements(billing_statement_id),
    FOREIGN KEY (client_id) REFERENCES Clients(client_id)
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_BillingEmployeeChangeLog_Statement
ON BillingEmployeeChangeLog(billing_statement_id, change_type, certificate_id);");
    }

    private static async Task SeedUnderwritingReferenceRatesAsync(SqliteConnection connection)
    {
        foreach (var version in UnderwritingReferenceRateCatalog.Defaults
                     .GroupBy(rate => new
                     {
                         rate.RateType,
                         rate.VersionName,
                         rate.UnitBasis,
                         rate.SourceName,
                         rate.SourceSheet
                     }))
        {
            using (var versionCommand = CreateCommand(connection, null, @"
INSERT OR IGNORE INTO underwriting_reference_rate_versions (
    rate_type,
    version_name,
    unit_basis,
    source_name,
    source_sheet,
    effective_date,
    is_active,
    date_created
)
VALUES (
    $rate_type,
    $version_name,
    $unit_basis,
    $source_name,
    $source_sheet,
    NULL,
    1,
    $date_created
);"))
            {
                AddParameter(versionCommand, "$rate_type", version.Key.RateType);
                AddParameter(versionCommand, "$version_name", version.Key.VersionName);
                AddParameter(versionCommand, "$unit_basis", version.Key.UnitBasis);
                AddParameter(versionCommand, "$source_name", version.Key.SourceName);
                AddParameter(versionCommand, "$source_sheet", version.Key.SourceSheet);
                AddParameter(versionCommand, "$date_created", ToDateTimeText(DateTime.Now));
                await versionCommand.ExecuteNonQueryAsync();
            }

            long rateVersionId;
            using (var idCommand = CreateCommand(connection, null, @"
SELECT rate_version_id
FROM underwriting_reference_rate_versions
WHERE rate_type = $rate_type
  AND version_name = $version_name;"))
            {
                AddParameter(idCommand, "$rate_type", version.Key.RateType);
                AddParameter(idCommand, "$version_name", version.Key.VersionName);
                rateVersionId = Convert.ToInt64(await idCommand.ExecuteScalarAsync(), CultureInfo.InvariantCulture);
            }

            using (var countCommand = CreateCommand(connection, null, @"
SELECT COUNT(*)
FROM underwriting_reference_rates
WHERE rate_version_id = $rate_version_id;"))
            {
                AddParameter(countCommand, "$rate_version_id", rateVersionId);
                var existingCount = Convert.ToInt32(await countCommand.ExecuteScalarAsync(), CultureInfo.InvariantCulture);
                if (existingCount >= version.Count())
                {
                    continue;
                }
            }

            foreach (var rate in version)
            {
                using var rateCommand = CreateCommand(connection, null, @"
INSERT OR IGNORE INTO underwriting_reference_rates (
    rate_version_id,
    age_band,
    gender,
    rate,
    display_order
)
VALUES (
    $rate_version_id,
    $age_band,
    $gender,
    $rate,
    $display_order
);");
                AddParameter(rateCommand, "$rate_version_id", rateVersionId);
                AddParameter(rateCommand, "$age_band", rate.AgeBand);
                AddParameter(rateCommand, "$gender", rate.Gender);
                AddParameter(rateCommand, "$rate", rate.Rate);
                AddParameter(rateCommand, "$display_order", rate.DisplayOrder);
                await rateCommand.ExecuteNonQueryAsync();
            }
        }
    }

    private async Task<SqliteConnection> OpenConnectionAsync()
    {
        var connection = new SqliteConnection($"Data Source={DatabasePath}");
        await connection.OpenAsync();
        await ExecuteNonQueryAsync(connection, null, "PRAGMA foreign_keys = ON;");
        return connection;
    }

    // Additive, idempotent schema preparation only. Existing records remain untouched and null.
    private static async Task EnsureRenewalCanonicalPlaceholderColumnsAsync(SqliteConnection connection)
    {
        await EnsureColumnExistsAsync(connection, "renewal_versions", "requested_increase_percent", "REAL");
        await EnsureColumnExistsAsync(connection, "renewal_versions", "negotiated_increase_percent", "REAL");
        await EnsureColumnExistsAsync(connection, "renewal_versions", "final_accepted_increase_percent", "REAL");
        await EnsureColumnExistsAsync(connection, "renewal_versions", "negotiation_status", "TEXT");
        await EnsureColumnExistsAsync(connection, "renewal_versions", "negotiation_stage", "TEXT");
        await EnsureColumnExistsAsync(connection, "renewal_versions", "negotiation_completed_date", "TEXT");
        await EnsureColumnExistsAsync(connection, "renewal_versions", "negotiated_savings_annual", "REAL");
        await EnsureColumnExistsAsync(connection, "renewal_versions", "negotiated_savings_monthly", "REAL");
        await EnsureColumnExistsAsync(connection, "renewal_versions", "negotiated_savings_type", "TEXT");
        await EnsureColumnExistsAsync(connection, "renewal_versions", "negotiated_by", "TEXT");
        await EnsureColumnExistsAsync(connection, "renewal_versions", "industry_benchmark_percent", "REAL");
        await EnsureColumnExistsAsync(connection, "renewal_versions", "industry_benchmark_source", "TEXT");
        await EnsureColumnExistsAsync(connection, "renewal_versions", "industry_benchmark_population", "TEXT");
        await EnsureColumnExistsAsync(connection, "renewal_versions", "industry_benchmark_year", "INTEGER");
        await EnsureColumnExistsAsync(connection, "renewal_versions", "industry_benchmark_confidence", "REAL");
        await EnsureColumnExistsAsync(connection, "renewal_versions", "renewal_recommendation_type", "TEXT");
        await EnsureColumnExistsAsync(connection, "renewal_versions", "renewal_decision_reason", "TEXT");
        await EnsureColumnExistsAsync(connection, "renewal_versions", "renewal_strategy_summary", "TEXT");
        await EnsureColumnExistsAsync(connection, "renewal_versions", "renewal_risk_level", "TEXT");
        await EnsureColumnExistsAsync(connection, "renewal_versions", "renewal_complexity", "TEXT");
        await EnsureColumnExistsAsync(connection, "renewal_versions", "renewal_result", "TEXT");
        await EnsureColumnExistsAsync(connection, "renewal_versions", "renewal_completed_date", "TEXT");
        await EnsureColumnExistsAsync(connection, "renewal_versions", "renewal_approved_by", "TEXT");
        await EnsureColumnExistsAsync(connection, "renewal_versions", "renewal_accepted_date", "TEXT");
        await EnsureColumnExistsAsync(connection, "renewal_versions", "renewal_version_number", "INTEGER");
        await EnsureColumnExistsAsync(connection, "renewal_versions", "renewal_scenario_id", "INTEGER");
        await EnsureColumnExistsAsync(connection, "renewal_versions", "previous_renewal_scenario_id", "INTEGER");
        await EnsureColumnExistsAsync(connection, "renewal_versions", "scenario_type", "TEXT");
        await EnsureColumnExistsAsync(connection, "renewal_versions", "history_event_id", "INTEGER");
        await EnsureColumnExistsAsync(connection, "renewal_versions", "relationship_milestone_id", "TEXT");
        await EnsureColumnExistsAsync(connection, "renewal_versions", "evolution_group_id", "TEXT");
    }

    private static async Task EnsureColumnExistsAsync(
        SqliteConnection connection,
        string tableName,
        string columnName,
        string columnDefinition)
    {
        using var command = CreateCommand(connection, null, $"PRAGMA table_info({tableName});");
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            var existingColumn = reader.GetString(reader.GetOrdinal("name"));
            if (existingColumn.Equals(columnName, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }
        }

        await ExecuteNonQueryAsync(connection, null, $"ALTER TABLE {tableName} ADD COLUMN {columnName} {columnDefinition};");
    }

    private static async Task EnsureSourceTrackingSchemaAsync(SqliteConnection connection)
    {
        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS Policy (
    policy_id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER,
    carrier_id INTEGER,
    policy_number TEXT NOT NULL DEFAULT '',
    effective_date TEXT,
    renewal_date TEXT,
    status TEXT NOT NULL DEFAULT '',
    source_type TEXT,
    source_carrier TEXT,
    source_file_name TEXT,
    source_document_id INTEGER,
    source_api_endpoint TEXT,
    import_batch_id INTEGER,
    imported_at TEXT,
    imported_by TEXT,
    confidence_score REAL,
    extraction_method TEXT,
    raw_source_reference TEXT
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS Benefit (
    benefit_id INTEGER PRIMARY KEY AUTOINCREMENT,
    benefit_category TEXT NOT NULL DEFAULT '',
    benefit_line TEXT NOT NULL DEFAULT '',
    normalized_name TEXT NOT NULL DEFAULT '',
    source_type TEXT,
    source_carrier TEXT,
    source_file_name TEXT,
    source_document_id INTEGER,
    source_api_endpoint TEXT,
    import_batch_id INTEGER,
    imported_at TEXT,
    imported_by TEXT,
    confidence_score REAL,
    extraction_method TEXT,
    raw_source_reference TEXT
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS BenefitClass (
    benefit_class_id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER,
    class_code TEXT NOT NULL DEFAULT '',
    class_name TEXT NOT NULL DEFAULT '',
    source_type TEXT,
    source_carrier TEXT,
    source_file_name TEXT,
    source_document_id INTEGER,
    source_api_endpoint TEXT,
    import_batch_id INTEGER,
    imported_at TEXT,
    imported_by TEXT,
    confidence_score REAL,
    extraction_method TEXT,
    raw_source_reference TEXT
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS BenefitOption (
    benefit_option_id INTEGER PRIMARY KEY AUTOINCREMENT,
    benefit_class_id INTEGER,
    benefit_id INTEGER,
    option_label TEXT NOT NULL DEFAULT '',
    source_type TEXT,
    source_carrier TEXT,
    source_file_name TEXT,
    source_document_id INTEGER,
    source_api_endpoint TEXT,
    import_batch_id INTEGER,
    imported_at TEXT,
    imported_by TEXT,
    confidence_score REAL,
    extraction_method TEXT,
    raw_source_reference TEXT
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS Premium (
    premium_id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER,
    benefit_line TEXT NOT NULL DEFAULT '',
    class_name TEXT NOT NULL DEFAULT '',
    period_label TEXT NOT NULL DEFAULT '',
    amount REAL,
    source_type TEXT,
    source_carrier TEXT,
    source_file_name TEXT,
    source_document_id INTEGER,
    source_api_endpoint TEXT,
    import_batch_id INTEGER,
    imported_at TEXT,
    imported_by TEXT,
    confidence_score REAL,
    extraction_method TEXT,
    raw_source_reference TEXT
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS BillingStatement (
    billing_statement_id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER,
    billing_period TEXT NOT NULL DEFAULT '',
    province TEXT NOT NULL DEFAULT '',
    benefit_line TEXT NOT NULL DEFAULT '',
    class_code TEXT NOT NULL DEFAULT '',
    option_label TEXT NOT NULL DEFAULT '',
    lives INTEGER,
    monthly_premium REAL,
    source_type TEXT,
    source_carrier TEXT,
    source_file_name TEXT,
    source_document_id INTEGER,
    source_api_endpoint TEXT,
    import_batch_id INTEGER,
    imported_at TEXT,
    imported_by TEXT,
    confidence_score REAL,
    extraction_method TEXT,
    raw_source_reference TEXT
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS Pooling (
    pooling_id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER,
    benefit_line TEXT NOT NULL DEFAULT '',
    pooling_threshold REAL,
    pooling_charge REAL,
    source_type TEXT,
    source_carrier TEXT,
    source_file_name TEXT,
    source_document_id INTEGER,
    source_api_endpoint TEXT,
    import_batch_id INTEGER,
    imported_at TEXT,
    imported_by TEXT,
    confidence_score REAL,
    extraction_method TEXT,
    raw_source_reference TEXT
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS Commission (
    commission_id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER,
    benefit_line TEXT NOT NULL DEFAULT '',
    commission_rate REAL,
    commission_amount REAL,
    source_type TEXT,
    source_carrier TEXT,
    source_file_name TEXT,
    source_document_id INTEGER,
    source_api_endpoint TEXT,
    import_batch_id INTEGER,
    imported_at TEXT,
    imported_by TEXT,
    confidence_score REAL,
    extraction_method TEXT,
    raw_source_reference TEXT
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS Quote (
    quote_id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER,
    carrier TEXT NOT NULL DEFAULT '',
    quote_name TEXT NOT NULL DEFAULT '',
    monthly_premium REAL,
    recommendation TEXT NOT NULL DEFAULT '',
    source_type TEXT,
    source_carrier TEXT,
    source_file_name TEXT,
    source_document_id INTEGER,
    source_api_endpoint TEXT,
    import_batch_id INTEGER,
    imported_at TEXT,
    imported_by TEXT,
    confidence_score REAL,
    extraction_method TEXT,
    raw_source_reference TEXT
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS Proposal (
    proposal_id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER,
    carrier TEXT NOT NULL DEFAULT '',
    proposal_name TEXT NOT NULL DEFAULT '',
    proposed_premium REAL,
    status TEXT NOT NULL DEFAULT '',
    source_type TEXT,
    source_carrier TEXT,
    source_file_name TEXT,
    source_document_id INTEGER,
    source_api_endpoint TEXT,
    import_batch_id INTEGER,
    imported_at TEXT,
    imported_by TEXT,
    confidence_score REAL,
    extraction_method TEXT,
    raw_source_reference TEXT
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS ImportBatch (
    import_batch_id INTEGER PRIMARY KEY AUTOINCREMENT,
    source_type TEXT NOT NULL DEFAULT '',
    source_carrier TEXT NOT NULL DEFAULT '',
    source_file_name TEXT NOT NULL DEFAULT '',
    source_document_id TEXT NOT NULL DEFAULT '',
    source_api_endpoint TEXT NOT NULL DEFAULT '',
    document_type TEXT NOT NULL DEFAULT '',
    status TEXT NOT NULL DEFAULT '',
    imported_at TEXT NOT NULL,
    imported_by TEXT NOT NULL DEFAULT '',
    confidence_score REAL,
    extraction_method TEXT NOT NULL DEFAULT '',
    raw_source_reference TEXT NOT NULL DEFAULT '',
    notes TEXT NOT NULL DEFAULT ''
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS DocumentSource (
    document_source_id INTEGER PRIMARY KEY AUTOINCREMENT,
    import_batch_id INTEGER,
    document_id INTEGER,
    source_type TEXT NOT NULL DEFAULT '',
    source_carrier TEXT NOT NULL DEFAULT '',
    source_file_name TEXT NOT NULL DEFAULT '',
    source_document_id TEXT NOT NULL DEFAULT '',
    source_api_endpoint TEXT NOT NULL DEFAULT '',
    document_type TEXT NOT NULL DEFAULT '',
    imported_at TEXT NOT NULL,
    imported_by TEXT NOT NULL DEFAULT '',
    confidence_score REAL,
    extraction_method TEXT NOT NULL DEFAULT '',
    raw_source_reference TEXT NOT NULL DEFAULT '',
    FOREIGN KEY (document_id) REFERENCES Documents(document_id),
    FOREIGN KEY (import_batch_id) REFERENCES ImportBatch(import_batch_id)
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS ImportLog (
    import_log_id INTEGER PRIMARY KEY AUTOINCREMENT,
    import_batch_id INTEGER,
    step_name TEXT NOT NULL DEFAULT '',
    status TEXT NOT NULL DEFAULT '',
    message TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL,
    FOREIGN KEY (import_batch_id) REFERENCES ImportBatch(import_batch_id)
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS ValidationIssue (
    validation_issue_id INTEGER PRIMARY KEY AUTOINCREMENT,
    import_batch_id INTEGER,
    source_document_id INTEGER,
    issue_code TEXT NOT NULL DEFAULT '',
    severity TEXT NOT NULL DEFAULT 'Warning',
    table_name TEXT NOT NULL DEFAULT '',
    field_name TEXT NOT NULL DEFAULT '',
    message TEXT NOT NULL DEFAULT '',
    raw_source_reference TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL,
    FOREIGN KEY (import_batch_id) REFERENCES ImportBatch(import_batch_id),
    FOREIGN KEY (source_document_id) REFERENCES Documents(document_id)
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_ImportBatch_Source
ON ImportBatch(source_type, source_carrier, document_type, imported_at);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS ClientHistoryEvent (
    history_event_id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL,
    effective_date TEXT,
    event_type TEXT NOT NULL DEFAULT '',
    benefit_category TEXT NOT NULL DEFAULT '',
    change_type TEXT NOT NULL DEFAULT '',
    title TEXT NOT NULL DEFAULT '',
    description TEXT NOT NULL DEFAULT '',
    previous_value TEXT NOT NULL DEFAULT '',
    new_value TEXT NOT NULL DEFAULT '',
    source_document_id INTEGER,
    source_sheet TEXT NOT NULL DEFAULT '',
    source_row INTEGER NOT NULL DEFAULT 0,
    original_text TEXT NOT NULL DEFAULT '',
    extraction_method TEXT NOT NULL DEFAULT '',
    confidence REAL,
    review_status TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL,
    created_by TEXT NOT NULL DEFAULT '',
    updated_at TEXT NOT NULL,
    updated_by TEXT NOT NULL DEFAULT '',
    import_batch_id INTEGER,
    source_fingerprint TEXT NOT NULL DEFAULT '',
    FOREIGN KEY (client_id) REFERENCES Clients(client_id),
    FOREIGN KEY (source_document_id) REFERENCES Documents(document_id),
    FOREIGN KEY (import_batch_id) REFERENCES ImportBatch(import_batch_id)
);
CREATE UNIQUE INDEX IF NOT EXISTS UX_ClientHistoryEvent_Fingerprint
ON ClientHistoryEvent(client_id, source_fingerprint)
WHERE TRIM(source_fingerprint) <> '';
CREATE INDEX IF NOT EXISTS IX_ClientHistoryEvent_Duplicate
ON ClientHistoryEvent(client_id, effective_date, event_type, description);
CREATE INDEX IF NOT EXISTS IX_ClientHistoryEvent_Evidence
ON ClientHistoryEvent(source_document_id, source_sheet, source_row);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_ValidationIssue_Batch
ON ValidationIssue(import_batch_id, issue_code);");

        await EnsureRenewalAnalysisSnapshotSchemaAsync(connection);

        foreach (var tableName in SourceTrackedTables())
        {
            await EnsureSourceMetadataColumnsAsync(connection, tableName);
        }
    }

    private static IEnumerable<string> SourceTrackedTables()
    {
        return
        [
            "Renewals",
            "Rates",
            "renewal_versions",
            "renewal_rates",
            "renewal_calculations",
            "renewal_assumptions",
            "renewal_projections",
            "underwriting_assumptions",
            "ClaimsExperience",
            "MonthlyClaims",
            "PlanDesign",
            "billing_option_premiums",
            "billing_province_summary",
            "billing_option_summary",
            "BillingReconciliationResults",
            "benefit_experience_history",
            "renewal_factors",
            "pooling_arrangements",
            "demographic_summary",
            "demographics_by_age_gender",
            "claims_distribution",
            "claims_distribution_reports",
            "claims_distribution_rows",
            "plan_design_features",
            "extraction_audit_log",
            "Policy",
            "Benefit",
            "BenefitClass",
            "BenefitOption",
            "Premium",
            "BillingStatement",
            "Pooling",
            "Commission",
            "Quote",
            "Proposal"
        ];
    }

    private static async Task EnsureSourceMetadataColumnsAsync(SqliteConnection connection, string tableName)
    {
        await EnsureColumnExistsAsync(connection, tableName, "source_type", "TEXT");
        await EnsureColumnExistsAsync(connection, tableName, "source_carrier", "TEXT");
        await EnsureColumnExistsAsync(connection, tableName, "source_file_name", "TEXT");
        await EnsureColumnExistsAsync(connection, tableName, "source_document_id", "INTEGER");
        await EnsureColumnExistsAsync(connection, tableName, "source_api_endpoint", "TEXT");
        await EnsureColumnExistsAsync(connection, tableName, "import_batch_id", "INTEGER");
        await EnsureColumnExistsAsync(connection, tableName, "imported_at", "TEXT");
        await EnsureColumnExistsAsync(connection, tableName, "imported_by", "TEXT");
        await EnsureColumnExistsAsync(connection, tableName, "confidence_score", "REAL");
        await EnsureColumnExistsAsync(connection, tableName, "extraction_method", "TEXT");
        await EnsureColumnExistsAsync(connection, tableName, "raw_source_reference", "TEXT");
    }

    private static async Task<long> UpsertClientAsync(SqliteConnection connection, SqliteTransaction transaction, ClientRecord client)
    {
        var clientName = ClientNameResolver.ToDisplayName(FirstNonBlank(client.ClientName, UnknownClientName));
        var clientNumber = Clean(client.ClientNumber);
        var now = DateTime.Now;
        var clientId = await FindClientIdByNumberAsync(connection, transaction, clientNumber) ??
                       await FindClientIdByNameAsync(connection, transaction, clientName);

        if (clientId is not null)
        {
            using var update = CreateCommand(connection, transaction, @"
UPDATE Clients
SET client_name = COALESCE(NULLIF($client_name, ''), client_name),
    client_number = COALESCE(NULLIF($client_number, ''), client_number),
    carrier = COALESCE(NULLIF($carrier, ''), carrier),
    industry = COALESCE(NULLIF($industry, ''), industry),
    province = COALESCE(NULLIF($province, ''), province),
    advisor = COALESCE(NULLIF($advisor, ''), advisor),
    effective_date = COALESCE(NULLIF($effective_date, ''), effective_date),
    renewal_month = COALESCE($renewal_month, renewal_month),
    status = COALESCE(NULLIF($status, ''), status),
    date_updated = $date_updated
WHERE client_id = $client_id;");
            AddParameter(update, "$client_name", clientName);
            AddParameter(update, "$client_number", clientNumber);
            AddParameter(update, "$carrier", Clean(client.Carrier));
            AddParameter(update, "$industry", Clean(client.Industry));
            AddParameter(update, "$province", Clean(client.Province));
            AddParameter(update, "$advisor", Clean(client.Advisor));
            AddParameter(update, "$effective_date", ToDateText(client.EffectiveDate));
            AddParameter(update, "$renewal_month", NormalizeRenewalMonth(client.RenewalMonth, client.EffectiveDate));
            AddParameter(update, "$status", FirstNonBlank(client.Status, DefaultClientStatus));
            AddParameter(update, "$date_updated", ToDateTimeText(now));
            AddParameter(update, "$client_id", clientId.Value);
            await update.ExecuteNonQueryAsync();
            return clientId.Value;
        }

        using var insert = CreateCommand(connection, transaction, @"
INSERT INTO Clients (
    client_name,
    client_number,
    carrier,
    industry,
    province,
    advisor,
    effective_date,
    renewal_month,
    status,
    date_created,
    date_updated
)
VALUES (
    $client_name,
    NULLIF($client_number, ''),
    NULLIF($carrier, ''),
    NULLIF($industry, ''),
    NULLIF($province, ''),
    NULLIF($advisor, ''),
    NULLIF($effective_date, ''),
    $renewal_month,
    $status,
    $date_created,
    $date_updated
);");
        AddParameter(insert, "$client_name", clientName);
        AddParameter(insert, "$client_number", clientNumber);
        AddParameter(insert, "$carrier", Clean(client.Carrier));
        AddParameter(insert, "$industry", Clean(client.Industry));
        AddParameter(insert, "$province", Clean(client.Province));
        AddParameter(insert, "$advisor", Clean(client.Advisor));
        AddParameter(insert, "$effective_date", ToDateText(client.EffectiveDate));
        AddParameter(insert, "$renewal_month", NormalizeRenewalMonth(client.RenewalMonth, client.EffectiveDate));
        AddParameter(insert, "$status", FirstNonBlank(client.Status, DefaultClientStatus));
        AddParameter(insert, "$date_created", ToDateTimeText(now));
        AddParameter(insert, "$date_updated", ToDateTimeText(now));
        await insert.ExecuteNonQueryAsync();

        using var idCommand = CreateCommand(connection, transaction, "SELECT last_insert_rowid();");
        return ToLong(await idCommand.ExecuteScalarAsync());
    }

    private static async Task<long?> FindClientIdByNumberAsync(SqliteConnection connection, SqliteTransaction transaction, string clientNumber)
    {
        if (string.IsNullOrWhiteSpace(clientNumber))
        {
            return null;
        }

        using var command = CreateCommand(
            connection,
            transaction,
            "SELECT client_id FROM Clients WHERE client_number = $client_number LIMIT 1;");
        AddParameter(command, "$client_number", clientNumber);
        return ToNullableLong(await command.ExecuteScalarAsync());
    }

    private static async Task<long?> FindClientIdByNameAsync(SqliteConnection connection, SqliteTransaction transaction, string clientName)
    {
        using var command = CreateCommand(
            connection,
            transaction,
            "SELECT client_id FROM Clients WHERE client_name = $client_name COLLATE NOCASE LIMIT 1;");
        AddParameter(command, "$client_name", clientName);
        return ToNullableLong(await command.ExecuteScalarAsync());
    }

    private static async Task<long> ResolveBillingClientAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        BillingExtractionResult result,
        long? selectedClientId)
    {
        if (selectedClientId.HasValue)
        {
            using var selected = CreateCommand(
                connection,
                transaction,
                "SELECT client_id, client_name, client_number FROM Clients WHERE client_id = $client_id LIMIT 1;");
            AddParameter(selected, "$client_id", selectedClientId.Value);
            using var reader = await selected.ExecuteReaderAsync();
            if (!await reader.ReadAsync())
            {
                throw new InvalidOperationException("The selected client no longer exists.");
            }

            var selectedName = ReadString(reader, "client_name");
            var selectedNumber = ReadString(reader, "client_number");
            var numberMismatch = result.ClientIdentityConfirmed &&
                                 !string.IsNullOrWhiteSpace(result.ClientNumber) &&
                                 !string.IsNullOrWhiteSpace(selectedNumber) &&
                                 !result.ClientNumber.Equals(selectedNumber, StringComparison.OrdinalIgnoreCase);
            var nameMismatch = result.ClientIdentityConfirmed &&
                               string.IsNullOrWhiteSpace(result.ClientNumber) &&
                               !string.IsNullOrWhiteSpace(result.ClientName) &&
                               !result.ClientName.Equals(selectedName, StringComparison.OrdinalIgnoreCase);
            if (numberMismatch || nameMismatch)
            {
                throw new InvalidOperationException("The selected client does not match the client identity on the billing statement.");
            }

            return selectedClientId.Value;
        }

        var clientName = FirstNonBlank(result.ClientName, UnknownClientName);
        var clientNumber = FirstNonBlank(result.ClientNumber, result.BillingId);
        var existing = await FindClientIdByNumberAsync(connection, transaction, clientNumber) ??
                       await FindClientIdByNameAsync(connection, transaction, clientName);
        if (existing.HasValue)
        {
            // Billing imports must never replace the client's lifecycle status.
            return existing.Value;
        }

        return await UpsertClientAsync(
            connection,
            transaction,
            new ClientRecord
            {
                ClientName = clientName,
                ClientNumber = clientNumber,
                Carrier = FirstNonBlank(result.Carrier, "Manulife"),
                EffectiveDate = result.PeriodStart,
                Status = DefaultClientStatus
            });
    }

    private static async Task<long?> FindManulifeBillingStatementByHashAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long clientId,
        string carrier,
        string fileHash)
    {
        using var command = CreateCommand(connection, transaction, @"
SELECT statement_version_id
FROM manulife_billing_statement_versions
WHERE client_id = $client_id
  AND carrier = $carrier COLLATE NOCASE
  AND source_file_hash = $source_file_hash
LIMIT 1;");
        AddParameter(command, "$client_id", clientId);
        AddParameter(command, "$carrier", Clean(carrier));
        AddParameter(command, "$source_file_hash", Clean(fileHash));
        return ToNullableLong(await command.ExecuteScalarAsync());
    }

    private static async Task<List<long>> GetActiveManulifeBillingVersionsAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long clientId,
        string carrier,
        string billingPeriod)
    {
        using var command = CreateCommand(connection, transaction, @"
SELECT statement_version_id
FROM manulife_billing_statement_versions
WHERE client_id = $client_id
  AND carrier = $carrier COLLATE NOCASE
  AND billing_period = $billing_period
  AND is_replaced = 0
ORDER BY version_number DESC;");
        AddParameter(command, "$client_id", clientId);
        AddParameter(command, "$carrier", Clean(carrier));
        AddParameter(command, "$billing_period", Clean(billingPeriod));
        var ids = new List<long>();
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            ids.Add(reader.GetInt64(0));
        }

        return ids;
    }

    private static async Task<int> GetNextManulifeBillingVersionNumberAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long clientId,
        string carrier,
        string billingPeriod)
    {
        using var command = CreateCommand(connection, transaction, @"
SELECT COALESCE(MAX(version_number), 0) + 1
FROM manulife_billing_statement_versions
WHERE client_id = $client_id
  AND carrier = $carrier COLLATE NOCASE
  AND billing_period = $billing_period;");
        AddParameter(command, "$client_id", clientId);
        AddParameter(command, "$carrier", Clean(carrier));
        AddParameter(command, "$billing_period", Clean(billingPeriod));
        return Convert.ToInt32(await command.ExecuteScalarAsync(), CultureInfo.InvariantCulture);
    }

    private static async Task<bool> HasDocumentHashForAnotherClientAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long clientId,
        string documentType,
        string fileHash)
    {
        using var command = CreateCommand(connection, transaction, @"
SELECT COUNT(*)
FROM Documents
WHERE client_id <> $client_id
  AND document_type = $document_type
  AND file_hash = $file_hash;");
        AddParameter(command, "$client_id", clientId);
        AddParameter(command, "$document_type", Clean(documentType));
        AddParameter(command, "$file_hash", Clean(fileHash));
        return Convert.ToInt32(await command.ExecuteScalarAsync(), CultureInfo.InvariantCulture) > 0;
    }

    private static async Task<long> InsertManulifeBillingStatementVersionAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        BillingExtractionResult result,
        long clientId,
        long documentId,
        long importBatchId,
        string fileHash,
        int versionNumber,
        long? replacesVersionId,
        string performedBy)
    {
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO manulife_billing_statement_versions (
    client_id, source_document_id, import_batch_id, carrier, billing_period,
    period_start, period_end, source_file_name, source_file_path, source_file_hash,
    version_number, import_status, is_replaced, replaces_statement_version_id,
    created_at, created_by
)
VALUES (
    $client_id, $source_document_id, $import_batch_id, $carrier, $billing_period,
    $period_start, $period_end, $source_file_name, $source_file_path, $source_file_hash,
    $version_number, $import_status, 0, $replaces_statement_version_id,
    $created_at, $created_by
);");
        AddParameter(command, "$client_id", clientId);
        AddParameter(command, "$source_document_id", documentId);
        AddParameter(command, "$import_batch_id", importBatchId);
        AddParameter(command, "$carrier", Clean(result.Carrier));
        AddParameter(command, "$billing_period", Clean(result.BillingPeriod));
        AddParameter(command, "$period_start", ToDateText(result.PeriodStart));
        AddParameter(command, "$period_end", ToDateText(result.PeriodEnd));
        AddParameter(command, "$source_file_name", Path.GetFileName(result.SourceFile));
        AddParameter(command, "$source_file_path", Path.GetFullPath(result.SourceFile));
        AddParameter(command, "$source_file_hash", Clean(fileHash));
        AddParameter(command, "$version_number", versionNumber);
        AddParameter(command, "$import_status", Clean(result.ImportStatus));
        AddParameter(command, "$replaces_statement_version_id", replacesVersionId);
        AddParameter(command, "$created_at", ToDateTimeText(DateTime.Now));
        AddParameter(command, "$created_by", Clean(performedBy));
        await command.ExecuteNonQueryAsync();

        using var idCommand = CreateCommand(connection, transaction, "SELECT last_insert_rowid();");
        return ToLong(await idCommand.ExecuteScalarAsync());
    }

    private static async Task UpdateManulifeBillingStatementStatusAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long statementVersionId,
        string status)
    {
        await ExecuteNonQueryAsync(
            connection,
            transaction,
            @"UPDATE manulife_billing_statement_versions
              SET import_status = $import_status
              WHERE statement_version_id = $statement_version_id;",
            command =>
            {
                AddParameter(command, "$import_status", Clean(status));
                AddParameter(command, "$statement_version_id", statementVersionId);
            });
    }

    private static async Task MarkManulifeBillingVersionsReplacedAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        IEnumerable<long> replacedVersionIds,
        long replacementVersionId,
        long clientId,
        string performedBy)
    {
        foreach (var replacedVersionId in replacedVersionIds)
        {
            await ExecuteNonQueryAsync(
                connection,
                transaction,
                @"UPDATE manulife_billing_statement_versions
                  SET is_replaced = 1,
                      replaced_by_statement_version_id = $replacement_statement_version_id
                  WHERE statement_version_id = $replaced_statement_version_id;",
                command =>
                {
                    AddParameter(command, "$replacement_statement_version_id", replacementVersionId);
                    AddParameter(command, "$replaced_statement_version_id", replacedVersionId);
                });

            await ExecuteNonQueryAsync(
                connection,
                transaction,
                @"INSERT INTO manulife_billing_version_audit (
                      client_id, replaced_statement_version_id, replacement_statement_version_id,
                      action, performed_by, performed_at
                  )
                  VALUES (
                      $client_id, $replaced_statement_version_id, $replacement_statement_version_id,
                      'ReplaceExistingVersion', $performed_by, $performed_at
                  );",
                command =>
                {
                    AddParameter(command, "$client_id", clientId);
                    AddParameter(command, "$replaced_statement_version_id", replacedVersionId);
                    AddParameter(command, "$replacement_statement_version_id", replacementVersionId);
                    AddParameter(command, "$performed_by", Clean(performedBy));
                    AddParameter(command, "$performed_at", ToDateTimeText(DateTime.Now));
                });
        }
    }

    private static async Task<DocumentSaveResult> UpsertDocumentAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long clientId,
        string documentType,
        string moduleName,
        string filePath,
        string carrier,
        DateOnly? effectiveDate,
        DateOnly? periodStart,
        DateOnly? periodEnd)
    {
        if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
        {
            return new DocumentSaveResult();
        }

        var fullPath = Path.GetFullPath(filePath);
        var fileHash = await ComputeFileHashAsync(fullPath);
        var existingDocumentId = await FindDocumentIdByHashAsync(
            connection,
            transaction,
            clientId,
            documentType,
            fileHash);
        var now = DateTime.Now;

        if (existingDocumentId is not null)
        {
            using var update = CreateCommand(connection, transaction, @"
UPDATE Documents
SET client_id = $client_id,
    document_type = $document_type,
    module_name = $module_name,
    file_name = $file_name,
    file_type = $file_type,
    file_path = $file_path,
    carrier = COALESCE(NULLIF($carrier, ''), carrier),
    effective_date = COALESCE(NULLIF($effective_date, ''), effective_date),
    period_start = COALESCE(NULLIF($period_start, ''), period_start),
    period_end = COALESCE(NULLIF($period_end, ''), period_end),
    upload_date = COALESCE(upload_date, $upload_date),
    processed_date = $processed_date,
    processing_status = 'reprocessed'
WHERE document_id = $document_id;");
            AddDocumentParameters(update, clientId, documentType, moduleName, fullPath, fileHash, carrier, effectiveDate, periodStart, periodEnd, now);
            AddParameter(update, "$document_id", existingDocumentId.Value);
            await update.ExecuteNonQueryAsync();
            return new DocumentSaveResult
            {
                DocumentId = existingDocumentId.Value,
                FilePath = fullPath,
                FileHash = fileHash,
                WasReprocessed = true
            };
        }

        using var insert = CreateCommand(connection, transaction, @"
INSERT INTO Documents (
    client_id,
    document_type,
    module_name,
    file_name,
    file_type,
    file_path,
    file_hash,
    carrier,
    effective_date,
    period_start,
    period_end,
    upload_date,
    processed_date,
    processing_status
)
VALUES (
    $client_id,
    $document_type,
    $module_name,
    $file_name,
    $file_type,
    $file_path,
    $file_hash,
    NULLIF($carrier, ''),
    NULLIF($effective_date, ''),
    NULLIF($period_start, ''),
    NULLIF($period_end, ''),
    $upload_date,
    $processed_date,
    'processed'
);");
        AddDocumentParameters(insert, clientId, documentType, moduleName, fullPath, fileHash, carrier, effectiveDate, periodStart, periodEnd, now);
        await insert.ExecuteNonQueryAsync();

        using var idCommand = CreateCommand(connection, transaction, "SELECT last_insert_rowid();");
        return new DocumentSaveResult
        {
            DocumentId = ToLong(await idCommand.ExecuteScalarAsync()),
            FilePath = fullPath,
            FileHash = fileHash,
            WasReprocessed = false
        };
    }

    private static void AddDocumentParameters(
        SqliteCommand command,
        long clientId,
        string documentType,
        string moduleName,
        string fullPath,
        string fileHash,
        string carrier,
        DateOnly? effectiveDate,
        DateOnly? periodStart,
        DateOnly? periodEnd,
        DateTime processedDate)
    {
        AddParameter(command, "$client_id", clientId);
        AddParameter(command, "$document_type", documentType);
        AddParameter(command, "$module_name", moduleName);
        AddParameter(command, "$file_name", Path.GetFileName(fullPath));
        AddParameter(command, "$file_type", FileTypeFromPath(fullPath));
        AddParameter(command, "$file_path", fullPath);
        AddParameter(command, "$file_hash", fileHash);
        AddParameter(command, "$carrier", Clean(carrier));
        AddParameter(command, "$effective_date", ToDateText(effectiveDate));
        AddParameter(command, "$period_start", ToDateText(periodStart));
        AddParameter(command, "$period_end", ToDateText(periodEnd));
        AddParameter(command, "$upload_date", ToDateTimeText(processedDate));
        AddParameter(command, "$processed_date", ToDateTimeText(processedDate));
    }

    private static async Task<long?> FindDocumentIdByHashAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long clientId,
        string documentType,
        string fileHash)
    {
        using var command = CreateCommand(
            connection,
            transaction,
            @"SELECT document_id
              FROM Documents
              WHERE client_id = $client_id
                AND document_type = $document_type
                AND file_hash = $file_hash
              LIMIT 1;");
        AddParameter(command, "$client_id", clientId);
        AddParameter(command, "$document_type", Clean(documentType));
        AddParameter(command, "$file_hash", fileHash);
        return ToNullableLong(await command.ExecuteScalarAsync());
    }

    private static async Task InsertPdfExtractionLogAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long documentId,
        string extractionMethod,
        decimal? confidenceScore,
        string fallbackReason,
        int? pageCount,
        int? tableCount,
        decimal? processingTime,
        string errorMessage,
        bool manualReviewRequired)
    {
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO PdfExtractionLog (
    document_id,
    extraction_method,
    confidence_score,
    fallback_reason,
    page_count,
    table_count,
    processing_time,
    error_message,
    manual_review_required,
    date_processed
)
VALUES (
    $document_id,
    $extraction_method,
    $confidence_score,
    NULLIF($fallback_reason, ''),
    $page_count,
    $table_count,
    $processing_time,
    NULLIF($error_message, ''),
    $manual_review_required,
    $date_processed
);");
        AddParameter(command, "$document_id", documentId);
        AddParameter(command, "$extraction_method", extractionMethod);
        AddParameter(command, "$confidence_score", confidenceScore);
        AddParameter(command, "$fallback_reason", Clean(fallbackReason));
        AddParameter(command, "$page_count", pageCount);
        AddParameter(command, "$table_count", tableCount);
        AddParameter(command, "$processing_time", processingTime);
        AddParameter(command, "$error_message", Clean(errorMessage));
        AddParameter(command, "$manual_review_required", manualReviewRequired ? 1 : 0);
        AddParameter(command, "$date_processed", ToDateTimeText(DateTime.Now));
        await command.ExecuteNonQueryAsync();
    }

    private static async Task<NormalizedSourceMetadata> StartImportTrackingAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        DocumentSaveResult document,
        string sourceType,
        string carrier,
        string documentType,
        string extractionMethod,
        decimal? confidenceScore)
    {
        var metadata = NormalizedSourceMetadata.ForFile(
            sourceType,
            carrier,
            document.FilePath,
            extractionMethod,
            confidenceScore);
        metadata.SourceDocumentId = document.DocumentId?.ToString(CultureInfo.InvariantCulture) ?? "";

        var importBatchId = await InsertImportBatchAsync(connection, transaction, metadata, documentType, ImportStatuses.Created);
        metadata.ImportBatchId = importBatchId;
        await InsertDocumentSourceAsync(connection, transaction, document.DocumentId, metadata, documentType);
        await InsertImportLogAsync(connection, transaction, importBatchId, "Create ImportBatch", ImportStatuses.Created, $"Import batch created for {metadata.SourceFileName}.");
        return metadata;
    }

    private static async Task<string> CompleteImportTrackingAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long documentId,
        NormalizedSourceMetadata source,
        NormalizedImportPayload payload,
        string? requiredStatus = null)
    {
        if (!source.ImportBatchId.HasValue)
        {
            return FirstNonBlank(requiredStatus, ImportStatuses.NeedsReview);
        }

        payload.Source = source.Copy();
        payload.Source.ImportBatchId = source.ImportBatchId;
        payload.Source.SourceDocumentId = documentId.ToString(CultureInfo.InvariantCulture);
        payload.ApplySourceToRecords();

        var issues = new NormalizedDataValidator().Validate(payload);
        issues.AddRange(payload.ValidationIssues);
        await DeleteValidationIssuesForBatchAsync(connection, transaction, source.ImportBatchId.Value);
        foreach (var issue in issues)
        {
            issue.ImportBatchId = source.ImportBatchId;
            issue.SourceDocumentId = documentId;
            await InsertValidationIssueAsync(connection, transaction, issue);
        }

        await ApplySourceMetadataForDocumentAsync(connection, transaction, documentId, source);

        var finalStatus = issues.Count == 0 &&
                          (string.IsNullOrWhiteSpace(requiredStatus) ||
                           requiredStatus.Equals(ImportStatuses.Completed, StringComparison.OrdinalIgnoreCase))
            ? ImportStatuses.Completed
            : ImportStatuses.NeedsReview;
        await InsertImportLogAsync(
            connection,
            transaction,
            source.ImportBatchId.Value,
            "Validate data",
            finalStatus,
            issues.Count == 0
                ? "Validation passed."
                : $"Validation saved {issues.Count} issue(s) for review.");
        await UpdateImportBatchStatusAsync(connection, transaction, source.ImportBatchId.Value, finalStatus, issues.Count == 0 ? "" : $"{issues.Count} validation issue(s).");
        return finalStatus;
    }

    private static async Task<long> InsertImportBatchAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        NormalizedSourceMetadata source,
        string documentType,
        string status)
    {
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO ImportBatch (
    source_type,
    source_carrier,
    source_file_name,
    source_document_id,
    source_api_endpoint,
    document_type,
    status,
    imported_at,
    imported_by,
    confidence_score,
    extraction_method,
    raw_source_reference,
    notes
)
VALUES (
    $source_type,
    $source_carrier,
    $source_file_name,
    $source_document_id,
    $source_api_endpoint,
    $document_type,
    $status,
    $imported_at,
    $imported_by,
    $confidence_score,
    $extraction_method,
    $raw_source_reference,
    ''
);
SELECT last_insert_rowid();");
        AddSourceParameters(command, source);
        AddParameter(command, "$document_type", Clean(documentType));
        AddParameter(command, "$status", Clean(status));
        return ToLong(await command.ExecuteScalarAsync());
    }

    private static async Task InsertDocumentSourceAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long? documentId,
        NormalizedSourceMetadata source,
        string documentType)
    {
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO DocumentSource (
    import_batch_id,
    document_id,
    source_type,
    source_carrier,
    source_file_name,
    source_document_id,
    source_api_endpoint,
    document_type,
    imported_at,
    imported_by,
    confidence_score,
    extraction_method,
    raw_source_reference
)
VALUES (
    $import_batch_id,
    $document_id,
    $source_type,
    $source_carrier,
    $source_file_name,
    $source_document_id,
    $source_api_endpoint,
    $document_type,
    $imported_at,
    $imported_by,
    $confidence_score,
    $extraction_method,
    $raw_source_reference
);");
        AddSourceParameters(command, source);
        AddParameter(command, "$document_id", documentId);
        AddParameter(command, "$document_type", Clean(documentType));
        await command.ExecuteNonQueryAsync();
    }

    private static async Task InsertImportLogAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long importBatchId,
        string stepName,
        string status,
        string message)
    {
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO ImportLog (
    import_batch_id,
    step_name,
    status,
    message,
    created_at
)
VALUES (
    $import_batch_id,
    $step_name,
    $status,
    $message,
    $created_at
);");
        AddParameter(command, "$import_batch_id", importBatchId);
        AddParameter(command, "$step_name", Clean(stepName));
        AddParameter(command, "$status", Clean(status));
        AddParameter(command, "$message", Clean(message));
        AddParameter(command, "$created_at", ToDateTimeText(DateTime.Now));
        await command.ExecuteNonQueryAsync();
    }

    private static async Task UpdateImportBatchStatusAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long importBatchId,
        string status,
        string notes)
    {
        using var command = CreateCommand(connection, transaction, @"
UPDATE ImportBatch
SET status = $status,
    notes = $notes
WHERE import_batch_id = $import_batch_id;");
        AddParameter(command, "$status", Clean(status));
        AddParameter(command, "$notes", Clean(notes));
        AddParameter(command, "$import_batch_id", importBatchId);
        await command.ExecuteNonQueryAsync();
    }

    private static async Task DeleteValidationIssuesForBatchAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long importBatchId)
    {
        using var command = CreateCommand(connection, transaction, "DELETE FROM ValidationIssue WHERE import_batch_id = $import_batch_id;");
        AddParameter(command, "$import_batch_id", importBatchId);
        await command.ExecuteNonQueryAsync();
    }

    private static async Task InsertValidationIssueAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        ValidationIssueRecord issue)
    {
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO ValidationIssue (
    import_batch_id,
    source_document_id,
    issue_code,
    severity,
    table_name,
    field_name,
    message,
    raw_source_reference,
    created_at
)
VALUES (
    $import_batch_id,
    $source_document_id,
    $issue_code,
    $severity,
    $table_name,
    $field_name,
    $message,
    $raw_source_reference,
    $created_at
);");
        AddParameter(command, "$import_batch_id", issue.ImportBatchId);
        AddParameter(command, "$source_document_id", issue.SourceDocumentId);
        AddParameter(command, "$issue_code", Clean(issue.IssueCode));
        AddParameter(command, "$severity", FirstNonBlank(issue.Severity, "Warning"));
        AddParameter(command, "$table_name", Clean(issue.TableName));
        AddParameter(command, "$field_name", Clean(issue.FieldName));
        AddParameter(command, "$message", Clean(issue.Message));
        AddParameter(command, "$raw_source_reference", Clean(issue.RawSourceReference));
        AddParameter(command, "$created_at", ToDateTimeText(issue.CreatedAt == default ? DateTime.Now : issue.CreatedAt));
        await command.ExecuteNonQueryAsync();
    }

    private static async Task ApplySourceMetadataForDocumentAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long documentId,
        NormalizedSourceMetadata source)
    {
        foreach (var tableName in DirectSourceDocumentTables())
        {
            using var command = CreateCommand(connection, transaction, $@"
UPDATE {tableName}
SET {SourceMetadataUpdateSet()}
WHERE source_document_id = $document_id;");
            AddSourceParameters(command, source);
            AddParameter(command, "$document_id", documentId);
            await command.ExecuteNonQueryAsync();
        }

        foreach (var tableName in RenewalVersionChildTables())
        {
            using var command = CreateCommand(connection, transaction, $@"
UPDATE {tableName}
SET {SourceMetadataUpdateSet()},
    source_document_id = $document_id
WHERE version_id IN (
    SELECT version_id
    FROM renewal_versions
    WHERE source_document_id = $document_id
);");
            AddSourceParameters(command, source);
            AddParameter(command, "$document_id", documentId);
            await command.ExecuteNonQueryAsync();
        }
    }

    private static IEnumerable<string> DirectSourceDocumentTables()
    {
        return
        [
            "Renewals",
            "Rates",
            "renewal_versions",
            "ClaimsExperience",
            "MonthlyClaims",
            "PlanDesign",
            "billing_option_premiums",
            "billing_province_summary",
            "billing_option_summary",
            "BillingReconciliationResults",
            "benefit_experience_history",
            "renewal_factors",
            "pooling_arrangements",
            "demographic_summary",
            "demographics_by_age_gender",
            "claims_distribution",
            "claims_distribution_reports",
            "claims_distribution_rows",
            "plan_design_features",
            "extraction_audit_log"
        ];
    }

    private static IEnumerable<string> RenewalVersionChildTables()
    {
        return
        [
            "renewal_rates",
            "renewal_calculations",
            "underwriting_assumptions"
        ];
    }

    private static string SourceMetadataUpdateSet()
    {
        return @"
    source_type = $source_type,
    source_carrier = $source_carrier,
    source_file_name = $source_file_name,
    source_api_endpoint = $source_api_endpoint,
    import_batch_id = $import_batch_id,
    imported_at = $imported_at,
    imported_by = $imported_by,
    confidence_score = $confidence_score,
    extraction_method = $extraction_method,
    raw_source_reference = $raw_source_reference";
    }

    private static void AddSourceParameters(SqliteCommand command, NormalizedSourceMetadata source)
    {
        AddParameter(command, "$source_type", Clean(source.SourceType));
        AddParameter(command, "$source_carrier", Clean(source.SourceCarrier));
        AddParameter(command, "$source_file_name", Clean(source.SourceFileName));
        AddParameter(command, "$source_document_id", Clean(source.SourceDocumentId));
        AddParameter(command, "$source_api_endpoint", Clean(source.SourceApiEndpoint));
        AddParameter(command, "$import_batch_id", source.ImportBatchId);
        AddParameter(command, "$imported_at", ToDateTimeText(source.ImportedAt == default ? DateTime.Now : source.ImportedAt));
        AddParameter(command, "$imported_by", Clean(source.ImportedBy));
        AddParameter(command, "$confidence_score", source.ConfidenceScore);
        AddParameter(command, "$extraction_method", Clean(source.ExtractionMethod));
        AddParameter(command, "$raw_source_reference", Clean(source.RawSourceReference));
    }

    private static async Task InsertRenewalAsync(SqliteConnection connection, SqliteTransaction transaction, RenewalRecord renewal)
    {
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO Renewals (
    client_id,
    renewal_year,
    carrier,
    benefit_line,
    current_rate,
    renewal_rate,
    rate_change_percent,
    premium,
    claims,
    loss_ratio,
    lives,
    renewal_status,
    source_document_id,
    date_loaded
)
VALUES (
    $client_id,
    $renewal_year,
    NULLIF($carrier, ''),
    NULLIF($benefit_line, ''),
    $current_rate,
    $renewal_rate,
    $rate_change_percent,
    $premium,
    $claims,
    $loss_ratio,
    $lives,
    $renewal_status,
    $source_document_id,
    $date_loaded
);");
        AddParameter(command, "$client_id", renewal.ClientId);
        AddParameter(command, "$renewal_year", renewal.RenewalYear);
        AddParameter(command, "$carrier", Clean(renewal.Carrier));
        AddParameter(command, "$benefit_line", Clean(renewal.BenefitLine));
        AddParameter(command, "$current_rate", renewal.CurrentRate);
        AddParameter(command, "$renewal_rate", renewal.RenewalRate);
        AddParameter(command, "$rate_change_percent", renewal.RateChangePercent);
        AddParameter(command, "$premium", renewal.Premium);
        AddParameter(command, "$claims", renewal.Claims);
        AddParameter(command, "$loss_ratio", renewal.LossRatio);
        AddParameter(command, "$lives", renewal.Lives);
        AddParameter(command, "$renewal_status", FirstNonBlank(renewal.RenewalStatus, "Processed"));
        AddParameter(command, "$source_document_id", renewal.SourceDocumentId);
        AddParameter(command, "$date_loaded", ToDateTimeText(renewal.DateLoaded));
        await command.ExecuteNonQueryAsync();
    }

    private static async Task InsertRateAsync(SqliteConnection connection, SqliteTransaction transaction, RateRecord rate)
    {
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO Rates (
    client_id,
    carrier,
    benefit_line,
    class,
    current_rate,
    proposed_rate,
    negotiated_rate,
    rate_basis,
    effective_date,
    source_document_id,
    date_loaded
)
VALUES (
    $client_id,
    NULLIF($carrier, ''),
    NULLIF($benefit_line, ''),
    NULLIF($class, ''),
    $current_rate,
    $proposed_rate,
    $negotiated_rate,
    NULLIF($rate_basis, ''),
    NULLIF($effective_date, ''),
    $source_document_id,
    $date_loaded
);");
        AddParameter(command, "$client_id", rate.ClientId);
        AddParameter(command, "$carrier", Clean(rate.Carrier));
        AddParameter(command, "$benefit_line", Clean(rate.BenefitLine));
        AddParameter(command, "$class", Clean(rate.ClassName));
        AddParameter(command, "$current_rate", rate.CurrentRate);
        AddParameter(command, "$proposed_rate", rate.ProposedRate);
        AddParameter(command, "$negotiated_rate", rate.NegotiatedRate);
        AddParameter(command, "$rate_basis", Clean(rate.RateBasis));
        AddParameter(command, "$effective_date", ToDateText(rate.EffectiveDate));
        AddParameter(command, "$source_document_id", rate.SourceDocumentId);
        AddParameter(command, "$date_loaded", ToDateTimeText(rate.DateLoaded));
        await command.ExecuteNonQueryAsync();
    }

    private static async Task<int> InsertRenewalVersionsAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        AnalysisResult analysis,
        long clientId,
        long? sourceDocumentId,
        DateOnly? effectiveDate)
    {
        var versions = analysis.RenewalVersions.Count > 0
            ? analysis.RenewalVersions
            : [];
        var insertedVersions = 0;

        foreach (var version in versions)
        {
            var versionId = await InsertRenewalVersionAsync(
                connection,
                transaction,
                version,
                clientId,
                sourceDocumentId,
                effectiveDate);
            insertedVersions++;

            foreach (var rateLine in version.RateLines)
            {
                await InsertRenewalRateVersionAsync(connection, transaction, versionId, rateLine);
            }

            foreach (var calculation in version.Calculations)
            {
                await InsertRenewalCalculationAsync(connection, transaction, versionId, calculation);
                await InsertUnderwritingAssumptionAsync(connection, transaction, clientId, versionId, calculation, effectiveDate);
            }
        }

        return insertedVersions;
    }

    private static async Task<long> InsertRenewalVersionAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        RenewalVersion version,
        long clientId,
        long? sourceDocumentId,
        DateOnly? effectiveDate)
    {
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO renewal_versions (
    client_id,
    version_name,
    source_module,
    status,
    is_selected,
    effective_date,
    source_document_id,
    date_created,
    requested_increase_percent,
    negotiated_increase_percent,
    final_accepted_increase_percent,
    negotiation_status,
    negotiation_stage,
    negotiation_completed_date,
    negotiated_savings_annual,
    negotiated_savings_monthly,
    negotiated_savings_type,
    negotiated_by,
    industry_benchmark_percent,
    industry_benchmark_source,
    industry_benchmark_population,
    industry_benchmark_year,
    industry_benchmark_confidence,
    renewal_recommendation_type,
    renewal_decision_reason,
    renewal_strategy_summary,
    renewal_risk_level,
    renewal_complexity,
    renewal_result,
    renewal_completed_date,
    renewal_approved_by,
    renewal_accepted_date,
    renewal_version_number,
    renewal_scenario_id,
    previous_renewal_scenario_id,
    scenario_type,
    history_event_id,
    relationship_milestone_id,
    evolution_group_id
)
VALUES (
    $client_id,
    $version_name,
    $source_module,
    $status,
    $is_selected,
    $effective_date,
    $source_document_id,
    $date_created,
    $requested_increase_percent,
    $negotiated_increase_percent,
    $final_accepted_increase_percent,
    $negotiation_status,
    $negotiation_stage,
    $negotiation_completed_date,
    $negotiated_savings_annual,
    $negotiated_savings_monthly,
    $negotiated_savings_type,
    $negotiated_by,
    $industry_benchmark_percent,
    $industry_benchmark_source,
    $industry_benchmark_population,
    $industry_benchmark_year,
    $industry_benchmark_confidence,
    $renewal_recommendation_type,
    $renewal_decision_reason,
    $renewal_strategy_summary,
    $renewal_risk_level,
    $renewal_complexity,
    $renewal_result,
    $renewal_completed_date,
    $renewal_approved_by,
    $renewal_accepted_date,
    $renewal_version_number,
    $renewal_scenario_id,
    $previous_renewal_scenario_id,
    $scenario_type,
    $history_event_id,
    $relationship_milestone_id,
    $evolution_group_id
);
SELECT last_insert_rowid();");
        AddParameter(command, "$client_id", clientId);
        AddParameter(command, "$version_name", Clean(version.VersionName));
        AddParameter(command, "$source_module", Clean(version.SourceModule));
        AddParameter(command, "$status", Clean(version.Status));
        AddParameter(command, "$is_selected", version.IsSelected ? 1 : 0);
        AddParameter(command, "$effective_date", ToDateText(effectiveDate));
        AddParameter(command, "$source_document_id", sourceDocumentId);
        AddParameter(command, "$date_created", ToDateTimeText(version.CreatedAt));
        AddParameter(command, "$requested_increase_percent", version.RequestedIncreasePercent);
        AddParameter(command, "$negotiated_increase_percent", version.NegotiatedIncreasePercent);
        AddParameter(command, "$final_accepted_increase_percent", version.FinalAcceptedIncreasePercent);
        AddParameter(command, "$negotiation_status", version.NegotiationStatus?.ToString());
        AddParameter(command, "$negotiation_stage", OptionalText(version.NegotiationStage ?? ""));
        AddParameter(command, "$negotiation_completed_date", OptionalText(ToDateText(version.NegotiationCompletedDate)));
        AddParameter(command, "$negotiated_savings_annual", version.NegotiatedSavingsAnnual);
        AddParameter(command, "$negotiated_savings_monthly", version.NegotiatedSavingsMonthly);
        AddParameter(command, "$negotiated_savings_type", version.NegotiatedSavingsType?.ToString());
        AddParameter(command, "$negotiated_by", OptionalText(version.NegotiatedBy ?? ""));
        AddParameter(command, "$industry_benchmark_percent", version.IndustryBenchmarkPercent);
        AddParameter(command, "$industry_benchmark_source", OptionalText(version.IndustryBenchmarkSource ?? ""));
        AddParameter(command, "$industry_benchmark_population", OptionalText(version.IndustryBenchmarkPopulation ?? ""));
        AddParameter(command, "$industry_benchmark_year", version.IndustryBenchmarkYear);
        AddParameter(command, "$industry_benchmark_confidence", version.IndustryBenchmarkConfidence);
        AddParameter(command, "$renewal_recommendation_type", version.RenewalRecommendationType?.ToString());
        AddParameter(command, "$renewal_decision_reason", OptionalText(version.RenewalDecisionReason ?? ""));
        AddParameter(command, "$renewal_strategy_summary", OptionalText(version.RenewalStrategySummary ?? ""));
        AddParameter(command, "$renewal_risk_level", version.RenewalRiskLevel?.ToString());
        AddParameter(command, "$renewal_complexity", OptionalText(version.RenewalComplexity ?? ""));
        AddParameter(command, "$renewal_result", OptionalText(version.RenewalResult ?? ""));
        AddParameter(command, "$renewal_completed_date", OptionalText(ToDateText(version.RenewalCompletedDate)));
        AddParameter(command, "$renewal_approved_by", OptionalText(version.RenewalApprovedBy ?? ""));
        AddParameter(command, "$renewal_accepted_date", OptionalText(ToDateText(version.RenewalAcceptedDate)));
        AddParameter(command, "$renewal_version_number", version.RenewalVersionNumber);
        AddParameter(command, "$renewal_scenario_id", version.RenewalScenarioId);
        AddParameter(command, "$previous_renewal_scenario_id", version.PreviousRenewalScenarioId);
        AddParameter(command, "$scenario_type", version.ScenarioType?.ToString());
        AddParameter(command, "$history_event_id", version.HistoryEventId);
        AddParameter(command, "$relationship_milestone_id", OptionalText(version.RelationshipMilestoneId ?? ""));
        AddParameter(command, "$evolution_group_id", OptionalText(version.EvolutionGroupId ?? ""));
        return ToLong(await command.ExecuteScalarAsync());
    }

    private static async Task InsertRenewalRateVersionAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long versionId,
        RenewalVersionRateLine rateLine)
    {
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO renewal_rates (
    version_id,
    benefit_line,
    class,
    account_class,
    volume,
    lives,
    current_rate,
    current_premium,
    proposed_rate,
    proposed_premium,
    version_rate,
    version_premium,
    change_percent,
    rate_basis,
    rate_value_basis,
    client_presentation_rate_basis,
    current_net_rate,
    current_gross_rate,
    proposed_net_rate,
    proposed_gross_rate,
    version_net_rate,
    version_gross_rate,
    rate_schedule_source_file,
    rate_schedule_source_sheet,
    rate_schedule_source_row,
    source_reference
)
VALUES (
    $version_id,
    $benefit_line,
    $class,
    $account_class,
    $volume,
    $lives,
    $current_rate,
    $current_premium,
    $proposed_rate,
    $proposed_premium,
    $version_rate,
    $version_premium,
    $change_percent,
    $rate_basis,
    $rate_value_basis,
    $client_presentation_rate_basis,
    $current_net_rate,
    $current_gross_rate,
    $proposed_net_rate,
    $proposed_gross_rate,
    $version_net_rate,
    $version_gross_rate,
    $rate_schedule_source_file,
    $rate_schedule_source_sheet,
    $rate_schedule_source_row,
    $source_reference
);");

        AddParameter(command, "$version_id", versionId);
        AddParameter(command, "$benefit_line", Clean(rateLine.Benefit));
        AddParameter(command, "$class", Clean(rateLine.TierOrClass));
        AddParameter(command, "$account_class", Clean(rateLine.AccountClass));
        AddParameter(command, "$volume", rateLine.Volume);
        AddParameter(command, "$lives", rateLine.Lives);
        AddParameter(command, "$current_rate", rateLine.CurrentRate);
        AddParameter(command, "$current_premium", rateLine.CurrentPremium);
        AddParameter(command, "$proposed_rate", rateLine.ProposedRate);
        AddParameter(command, "$proposed_premium", rateLine.ProposedPremium);
        AddParameter(command, "$version_rate", rateLine.VersionRate);
        AddParameter(command, "$version_premium", rateLine.VersionPremium);
        AddParameter(command, "$change_percent", rateLine.ChangePercent);
        AddParameter(command, "$rate_basis", Clean(rateLine.RateBasis));
        AddParameter(command, "$rate_value_basis", Clean(rateLine.RateValueBasis));
        AddParameter(command, "$client_presentation_rate_basis", Clean(rateLine.ClientPresentationRateBasis));
        AddParameter(command, "$current_net_rate", rateLine.CurrentNetRate);
        AddParameter(command, "$current_gross_rate", rateLine.CurrentGrossRate);
        AddParameter(command, "$proposed_net_rate", rateLine.ProposedNetRate);
        AddParameter(command, "$proposed_gross_rate", rateLine.ProposedGrossRate);
        AddParameter(command, "$version_net_rate", rateLine.VersionNetRate);
        AddParameter(command, "$version_gross_rate", rateLine.VersionGrossRate);
        AddParameter(command, "$rate_schedule_source_file", Clean(rateLine.RateScheduleSourceFile));
        AddParameter(command, "$rate_schedule_source_sheet", Clean(rateLine.RateScheduleSourceSheet));
        AddParameter(command, "$rate_schedule_source_row", rateLine.RateScheduleSourceRow);
        AddParameter(command, "$source_reference", Clean(rateLine.SourceReference));
        await command.ExecuteNonQueryAsync();
    }

    private static async Task InsertRenewalCalculationAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long versionId,
        UnderwritingCalculationResult calculation)
    {
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO renewal_calculations (
    version_id,
    source_module,
    benefit_line,
    current_claims,
    current_premium,
    trended_claims,
    pooling_adjustment,
    ibnr_adjustment,
    retention_admin_fee_adjustment,
    manual_adjustment,
    projected_claims,
    required_premium,
    proposed_premium,
    negotiated_premium,
    required_rate,
    target_loss_ratio,
    projected_loss_ratio,
    renewal_change_percent,
    negotiated_change_percent,
    savings,
    trend
)
VALUES (
    $version_id,
    $source_module,
    $benefit_line,
    $current_claims,
    $current_premium,
    $trended_claims,
    $pooling_adjustment,
    $ibnr_adjustment,
    $retention_admin_fee_adjustment,
    $manual_adjustment,
    $projected_claims,
    $required_premium,
    $proposed_premium,
    $negotiated_premium,
    $required_rate,
    $target_loss_ratio,
    $projected_loss_ratio,
    $renewal_change_percent,
    $negotiated_change_percent,
    $savings,
    $trend
);");
        AddUnderwritingCalculationParameters(command, versionId, calculation);
        await command.ExecuteNonQueryAsync();
    }

    private static async Task InsertUnderwritingAssumptionAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long clientId,
        long versionId,
        UnderwritingCalculationResult calculation,
        DateOnly? effectiveDate)
    {
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO underwriting_assumptions (
    client_id,
    version_id,
    carrier,
    benefit_line,
    trend,
    target_loss_ratio,
    pooling_adjustment,
    ibnr_adjustment,
    retention_admin_fee,
    manual_adjustment,
    effective_date,
    date_loaded
)
VALUES (
    $client_id,
    $version_id,
    $carrier,
    $benefit_line,
    $trend,
    $target_loss_ratio,
    $pooling_adjustment,
    $ibnr_adjustment,
    $retention_admin_fee,
    $manual_adjustment,
    $effective_date,
    $date_loaded
);");
        AddParameter(command, "$client_id", clientId);
        AddParameter(command, "$version_id", versionId);
        AddParameter(command, "$carrier", "Manulife");
        AddParameter(command, "$benefit_line", Clean(calculation.BenefitLine));
        AddParameter(command, "$trend", calculation.Trend);
        AddParameter(command, "$target_loss_ratio", calculation.TargetLossRatio);
        AddParameter(command, "$pooling_adjustment", calculation.PoolingAdjustment);
        AddParameter(command, "$ibnr_adjustment", calculation.IbnrAdjustment);
        AddParameter(command, "$retention_admin_fee", calculation.RetentionAdminFeeAdjustment);
        AddParameter(command, "$manual_adjustment", calculation.ManualAdjustment);
        AddParameter(command, "$effective_date", ToDateText(effectiveDate));
        AddParameter(command, "$date_loaded", ToDateTimeText(DateTime.Now));
        await command.ExecuteNonQueryAsync();
    }

    private static async Task DeleteNegotiationVersionAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long clientId,
        DateOnly? effectiveDate,
        string versionName)
    {
        var effectiveDateText = ToDateText(effectiveDate);
        var versionFilter = @"
SELECT version_id
FROM renewal_versions
WHERE client_id = $client_id
  AND version_name = $version_name
  AND IFNULL(effective_date, '') = IFNULL($effective_date, '')";

        await ExecuteNonQueryAsync(
            connection,
            transaction,
            $"DELETE FROM negotiation_notes WHERE renewal_version_id IN ({versionFilter});",
            command =>
            {
                AddParameter(command, "$client_id", clientId);
                AddParameter(command, "$version_name", Clean(versionName));
                AddParameter(command, "$effective_date", effectiveDateText);
            });
        await ExecuteNonQueryAsync(
            connection,
            transaction,
            $"DELETE FROM renewal_calculations WHERE version_id IN ({versionFilter});",
            command =>
            {
                AddParameter(command, "$client_id", clientId);
                AddParameter(command, "$version_name", Clean(versionName));
                AddParameter(command, "$effective_date", effectiveDateText);
            });
        await ExecuteNonQueryAsync(
            connection,
            transaction,
            $"DELETE FROM renewal_rates WHERE version_id IN ({versionFilter});",
            command =>
            {
                AddParameter(command, "$client_id", clientId);
                AddParameter(command, "$version_name", Clean(versionName));
                AddParameter(command, "$effective_date", effectiveDateText);
            });
        await ExecuteNonQueryAsync(
            connection,
            transaction,
            $"DELETE FROM underwriting_assumptions WHERE version_id IN ({versionFilter});",
            command =>
            {
                AddParameter(command, "$client_id", clientId);
                AddParameter(command, "$version_name", Clean(versionName));
                AddParameter(command, "$effective_date", effectiveDateText);
            });
        await ExecuteNonQueryAsync(
            connection,
            transaction,
            "DELETE FROM renewal_versions WHERE client_id = $client_id AND version_name = $version_name AND IFNULL(effective_date, '') = IFNULL($effective_date, '');",
            command =>
            {
                AddParameter(command, "$client_id", clientId);
                AddParameter(command, "$version_name", Clean(versionName));
                AddParameter(command, "$effective_date", effectiveDateText);
            });
    }

    private static async Task<long> InsertNegotiationVersionAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long clientId,
        NegotiationWorkbenchSaveRequest request)
    {
        var now = DateTime.Now;
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO renewal_versions (
    client_id,
    version_name,
    source_module,
    status,
    is_selected,
    effective_date,
    source_document_id,
    date_created,
    renewal_id,
    renewal_year,
    carrier,
    updated_at,
    created_by,
    notes,
    requested_increase_percent,
    negotiated_increase_percent,
    final_accepted_increase_percent,
    negotiation_status,
    negotiation_stage,
    negotiation_completed_date,
    negotiated_savings_annual,
    negotiated_savings_monthly,
    negotiated_savings_type,
    negotiated_by,
    industry_benchmark_percent,
    industry_benchmark_source,
    industry_benchmark_population,
    industry_benchmark_year,
    industry_benchmark_confidence,
    renewal_recommendation_type,
    renewal_decision_reason,
    renewal_strategy_summary,
    renewal_risk_level,
    renewal_complexity,
    renewal_result,
    renewal_completed_date,
    renewal_approved_by,
    renewal_accepted_date,
    renewal_version_number,
    renewal_scenario_id,
    previous_renewal_scenario_id,
    scenario_type,
    history_event_id,
    relationship_milestone_id,
    evolution_group_id
)
VALUES (
    $client_id,
    $version_name,
    $source_module,
    $status,
    $is_selected,
    $effective_date,
    $source_document_id,
    $date_created,
    $renewal_id,
    $renewal_year,
    $carrier,
    $updated_at,
    $created_by,
    $notes,
    $requested_increase_percent,
    $negotiated_increase_percent,
    $final_accepted_increase_percent,
    $negotiation_status,
    $negotiation_stage,
    $negotiation_completed_date,
    $negotiated_savings_annual,
    $negotiated_savings_monthly,
    $negotiated_savings_type,
    $negotiated_by,
    $industry_benchmark_percent,
    $industry_benchmark_source,
    $industry_benchmark_population,
    $industry_benchmark_year,
    $industry_benchmark_confidence,
    $renewal_recommendation_type,
    $renewal_decision_reason,
    $renewal_strategy_summary,
    $renewal_risk_level,
    $renewal_complexity,
    $renewal_result,
    $renewal_completed_date,
    $renewal_approved_by,
    $renewal_accepted_date,
    $renewal_version_number,
    $renewal_scenario_id,
    $previous_renewal_scenario_id,
    $scenario_type,
    $history_event_id,
    $relationship_milestone_id,
    $evolution_group_id
);
SELECT last_insert_rowid();");
        AddParameter(command, "$client_id", clientId);
        AddParameter(command, "$version_name", Clean(request.VersionName));
        AddParameter(command, "$source_module", BenefitsIntelligenceModules.RenewalIntelligence);
        AddParameter(command, "$status", Clean(request.Status));
        AddParameter(command, "$is_selected", 1);
        AddParameter(command, "$effective_date", ToDateText(request.EffectiveDate));
        AddParameter(command, "$source_document_id", null);
        AddParameter(command, "$date_created", ToDateTimeText(now));
        AddParameter(command, "$renewal_id", request.RenewalId > 0 ? request.RenewalId : null);
        AddParameter(command, "$renewal_year", request.RenewalYear);
        AddParameter(command, "$carrier", Clean(request.Carrier));
        AddParameter(command, "$updated_at", ToDateTimeText(now));
        AddParameter(command, "$created_by", Clean(request.CreatedBy));
        AddParameter(command, "$notes", Clean(request.Notes));
        AddParameter(command, "$requested_increase_percent", request.RequestedIncreasePercent);
        AddParameter(command, "$negotiated_increase_percent", request.NegotiatedIncreasePercent);
        AddParameter(command, "$final_accepted_increase_percent", request.FinalAcceptedIncreasePercent);
        AddParameter(command, "$negotiation_status", request.NegotiationStatus?.ToString());
        AddParameter(command, "$negotiation_stage", OptionalText(request.NegotiationStage ?? ""));
        AddParameter(command, "$negotiation_completed_date", OptionalText(ToDateText(request.NegotiationCompletedDate)));
        AddParameter(command, "$negotiated_savings_annual", request.NegotiatedSavingsAnnual);
        AddParameter(command, "$negotiated_savings_monthly", request.NegotiatedSavingsMonthly);
        AddParameter(command, "$negotiated_savings_type", request.NegotiatedSavingsType?.ToString());
        AddParameter(command, "$negotiated_by", OptionalText(request.NegotiatedBy ?? ""));
        AddParameter(command, "$industry_benchmark_percent", request.IndustryBenchmarkPercent);
        AddParameter(command, "$industry_benchmark_source", OptionalText(request.IndustryBenchmarkSource ?? ""));
        AddParameter(command, "$industry_benchmark_population", OptionalText(request.IndustryBenchmarkPopulation ?? ""));
        AddParameter(command, "$industry_benchmark_year", request.IndustryBenchmarkYear);
        AddParameter(command, "$industry_benchmark_confidence", request.IndustryBenchmarkConfidence);
        AddParameter(command, "$renewal_recommendation_type", request.RenewalRecommendationType?.ToString());
        AddParameter(command, "$renewal_decision_reason", OptionalText(request.RenewalDecisionReason ?? ""));
        AddParameter(command, "$renewal_strategy_summary", OptionalText(request.RenewalStrategySummary ?? ""));
        AddParameter(command, "$renewal_risk_level", request.RenewalRiskLevel?.ToString());
        AddParameter(command, "$renewal_complexity", OptionalText(request.RenewalComplexity ?? ""));
        AddParameter(command, "$renewal_result", OptionalText(request.RenewalResult ?? ""));
        AddParameter(command, "$renewal_completed_date", OptionalText(ToDateText(request.RenewalCompletedDate)));
        AddParameter(command, "$renewal_approved_by", OptionalText(request.RenewalApprovedBy ?? ""));
        AddParameter(command, "$renewal_accepted_date", OptionalText(ToDateText(request.RenewalAcceptedDate)));
        AddParameter(command, "$renewal_version_number", request.RenewalVersionNumber);
        AddParameter(command, "$renewal_scenario_id", request.RenewalScenarioId);
        AddParameter(command, "$previous_renewal_scenario_id", request.PreviousRenewalScenarioId);
        AddParameter(command, "$scenario_type", request.ScenarioType?.ToString());
        AddParameter(command, "$history_event_id", request.HistoryEventId);
        AddParameter(command, "$relationship_milestone_id", OptionalText(request.RelationshipMilestoneId ?? ""));
        AddParameter(command, "$evolution_group_id", OptionalText(request.EvolutionGroupId ?? ""));
        return ToLong(await command.ExecuteScalarAsync());
    }

    private static async Task InsertNegotiationRateAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long versionId,
        long clientId,
        long renewalId,
        NegotiationRateSaveRecord rate)
    {
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO renewal_rates (
    version_id,
    benefit_line,
    class,
    account_class,
    volume,
    lives,
    current_rate,
    current_premium,
    proposed_rate,
    proposed_premium,
    version_rate,
    version_premium,
    change_percent,
    rate_basis,
    rate_value_basis,
    client_presentation_rate_basis,
    current_net_rate,
    current_gross_rate,
    proposed_net_rate,
    proposed_gross_rate,
    version_net_rate,
    version_gross_rate,
    rate_schedule_source_file,
    rate_schedule_source_sheet,
    rate_schedule_source_row,
    source_reference,
    renewal_id,
    client_id,
    class_name,
    modelled_rate,
    negotiated_rate,
    modelled_premium,
    negotiated_premium,
    proposed_change_pct,
    modelled_change_pct,
    negotiated_change_pct,
    savings,
    status,
    notes,
    source,
    updated_at
)
VALUES (
    $version_id,
    $benefit_line,
    $class,
    $account_class,
    $volume,
    $lives,
    $current_rate,
    $current_premium,
    $proposed_rate,
    $proposed_premium,
    $version_rate,
    $version_premium,
    $change_percent,
    $rate_basis,
    $rate_value_basis,
    $client_presentation_rate_basis,
    $current_net_rate,
    $current_gross_rate,
    $proposed_net_rate,
    $proposed_gross_rate,
    $version_net_rate,
    $version_gross_rate,
    $rate_schedule_source_file,
    $rate_schedule_source_sheet,
    $rate_schedule_source_row,
    $source_reference,
    $renewal_id,
    $client_id,
    $class_name,
    $modelled_rate,
    $negotiated_rate,
    $modelled_premium,
    $negotiated_premium,
    $proposed_change_pct,
    $modelled_change_pct,
    $negotiated_change_pct,
    $savings,
    $status,
    $notes,
    $source,
    $updated_at
);");
        AddParameter(command, "$version_id", versionId);
        AddParameter(command, "$benefit_line", Clean(rate.BenefitLine));
        AddParameter(command, "$class", Clean(rate.ClassName));
        AddParameter(command, "$account_class", Clean(rate.AccountClass));
        AddParameter(command, "$volume", rate.Volume);
        AddParameter(command, "$lives", rate.Lives);
        AddParameter(command, "$current_rate", rate.CurrentRate);
        AddParameter(command, "$current_premium", rate.CurrentPremium);
        AddParameter(command, "$proposed_rate", rate.ProposedRate);
        AddParameter(command, "$proposed_premium", rate.ProposedPremium);
        AddParameter(command, "$version_rate", rate.NegotiatedRate ?? rate.ModelledRate ?? rate.ProposedRate);
        AddParameter(command, "$version_premium", rate.NegotiatedPremium ?? rate.ModelledPremium ?? rate.ProposedPremium);
        AddParameter(command, "$change_percent", PercentDecimalToPercent(rate.NegotiatedChangePct ?? rate.ModelledChangePct ?? rate.ProposedChangePct));
        AddParameter(command, "$rate_basis", Clean(rate.RateBasis));
        AddParameter(command, "$rate_value_basis", Clean(rate.RateValueBasis));
        AddParameter(command, "$client_presentation_rate_basis", Clean(rate.ClientPresentationRateBasis));
        AddParameter(command, "$current_net_rate", rate.CurrentNetRate);
        AddParameter(command, "$current_gross_rate", rate.CurrentGrossRate);
        AddParameter(command, "$proposed_net_rate", rate.ProposedNetRate);
        AddParameter(command, "$proposed_gross_rate", rate.ProposedGrossRate);
        AddParameter(command, "$version_net_rate", rate.VersionNetRate);
        AddParameter(command, "$version_gross_rate", rate.VersionGrossRate);
        AddParameter(command, "$rate_schedule_source_file", Clean(rate.RateScheduleSourceFile));
        AddParameter(command, "$rate_schedule_source_sheet", Clean(rate.RateScheduleSourceSheet));
        AddParameter(command, "$rate_schedule_source_row", rate.RateScheduleSourceRow);
        AddParameter(command, "$source_reference", Clean(rate.Source));
        AddParameter(command, "$renewal_id", renewalId > 0 ? renewalId : null);
        AddParameter(command, "$client_id", clientId);
        AddParameter(command, "$class_name", Clean(rate.ClassName));
        AddParameter(command, "$modelled_rate", rate.ModelledRate);
        AddParameter(command, "$negotiated_rate", rate.NegotiatedRate);
        AddParameter(command, "$modelled_premium", rate.ModelledPremium);
        AddParameter(command, "$negotiated_premium", rate.NegotiatedPremium);
        AddParameter(command, "$proposed_change_pct", rate.ProposedChangePct);
        AddParameter(command, "$modelled_change_pct", rate.ModelledChangePct);
        AddParameter(command, "$negotiated_change_pct", rate.NegotiatedChangePct);
        AddParameter(command, "$savings", rate.Savings);
        AddParameter(command, "$status", Clean(rate.Status));
        AddParameter(command, "$notes", Clean(rate.Notes));
        AddParameter(command, "$source", Clean(rate.Source));
        AddParameter(command, "$updated_at", ToDateTimeText(DateTime.Now));
        await command.ExecuteNonQueryAsync();
    }

    private static async Task InsertNegotiationCalculationAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long versionId,
        long clientId,
        long renewalId,
        UnderwritingCalculationResult calculation)
    {
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO renewal_calculations (
    version_id,
    source_module,
    benefit_line,
    current_claims,
    current_premium,
    trended_claims,
    pooling_adjustment,
    ibnr_adjustment,
    retention_admin_fee_adjustment,
    manual_adjustment,
    projected_claims,
    required_premium,
    proposed_premium,
    negotiated_premium,
    required_rate,
    target_loss_ratio,
    projected_loss_ratio,
    renewal_change_percent,
    negotiated_change_percent,
    savings,
    trend,
    renewal_id,
    client_id,
    carrier_proposed_premium,
    modelled_premium,
    proposed_change_pct,
    modelled_change_pct,
    negotiated_change_pct,
    calculation_notes,
    calculated_at
)
VALUES (
    $version_id,
    $source_module,
    $benefit_line,
    $current_claims,
    $current_premium,
    $trended_claims,
    $pooling_adjustment,
    $ibnr_adjustment,
    $retention_admin_fee_adjustment,
    $manual_adjustment,
    $projected_claims,
    $required_premium,
    $proposed_premium,
    $negotiated_premium,
    $required_rate,
    $target_loss_ratio,
    $projected_loss_ratio,
    $renewal_change_percent,
    $negotiated_change_percent,
    $savings,
    $trend,
    $renewal_id,
    $client_id,
    $carrier_proposed_premium,
    $modelled_premium,
    $proposed_change_pct,
    $modelled_change_pct,
    $negotiated_change_pct,
    $calculation_notes,
    $calculated_at
);");
        AddUnderwritingCalculationParameters(command, versionId, calculation);
        AddParameter(command, "$renewal_id", renewalId > 0 ? renewalId : null);
        AddParameter(command, "$client_id", clientId);
        AddParameter(command, "$carrier_proposed_premium", calculation.ProposedPremium);
        AddParameter(command, "$modelled_premium", calculation.RequiredPremium);
        AddParameter(command, "$proposed_change_pct", PercentToDecimal(calculation.RenewalChangePercent));
        AddParameter(command, "$modelled_change_pct", PercentToDecimal(UnderwritingEngine.CalculatePercentChange(calculation.CurrentPremium, calculation.RequiredPremium)));
        AddParameter(command, "$negotiated_change_pct", PercentToDecimal(calculation.NegotiatedChangePercent));
        AddParameter(command, "$calculation_notes", "Renewal Negotiation Workbench");
        AddParameter(command, "$calculated_at", ToDateTimeText(DateTime.Now));
        await command.ExecuteNonQueryAsync();
    }

    private static async Task UpsertNegotiationAssumptionAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long versionId,
        long clientId,
        NegotiationWorkbenchSaveRequest request)
    {
        var assumption = request.Assumptions;
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO renewal_assumptions (
    client_id,
    carrier,
    benefit_line,
    trend,
    target_loss_ratio,
    pooling_charge,
    pooling_threshold,
    ibnr,
    admin_fee,
    risk_charge,
    effective_date,
    renewal_version_id,
    renewal_id,
    lives,
    volume,
    credibility_factor,
    manual_adjustment_pct,
    updated_at
)
VALUES (
    $client_id,
    $carrier,
    $benefit_line,
    $trend,
    $target_loss_ratio,
    $pooling_charge,
    $pooling_threshold,
    $ibnr,
    $admin_fee,
    $risk_charge,
    $effective_date,
    $renewal_version_id,
    $renewal_id,
    $lives,
    $volume,
    $credibility_factor,
    $manual_adjustment_pct,
    $updated_at
)
ON CONFLICT(client_id, carrier, benefit_line, effective_date)
DO UPDATE SET
    trend = excluded.trend,
    target_loss_ratio = excluded.target_loss_ratio,
    pooling_charge = excluded.pooling_charge,
    pooling_threshold = excluded.pooling_threshold,
    ibnr = excluded.ibnr,
    admin_fee = excluded.admin_fee,
    risk_charge = excluded.risk_charge,
    renewal_version_id = excluded.renewal_version_id,
    renewal_id = excluded.renewal_id,
    lives = excluded.lives,
    volume = excluded.volume,
    credibility_factor = excluded.credibility_factor,
    manual_adjustment_pct = excluded.manual_adjustment_pct,
    updated_at = excluded.updated_at;");
        AddParameter(command, "$client_id", clientId);
        AddParameter(command, "$carrier", Clean(request.Carrier));
        AddParameter(command, "$benefit_line", Clean(assumption.BenefitLine));
        AddParameter(command, "$trend", assumption.TrendPct);
        AddParameter(command, "$target_loss_ratio", assumption.TargetLossRatioPct);
        AddParameter(command, "$pooling_charge", assumption.PoolingChargePct);
        AddParameter(command, "$pooling_threshold", assumption.PoolingThreshold);
        AddParameter(command, "$ibnr", assumption.IbnrPct);
        AddParameter(command, "$admin_fee", assumption.AdminFeePct);
        AddParameter(command, "$risk_charge", assumption.RiskChargePct);
        AddParameter(command, "$effective_date", ToDateText(request.EffectiveDate));
        AddParameter(command, "$renewal_version_id", versionId);
        AddParameter(command, "$renewal_id", request.RenewalId > 0 ? request.RenewalId : null);
        AddParameter(command, "$lives", assumption.Lives);
        AddParameter(command, "$volume", assumption.Volume);
        AddParameter(command, "$credibility_factor", assumption.CredibilityFactor);
        AddParameter(command, "$manual_adjustment_pct", assumption.ManualAdjustmentPct);
        AddParameter(command, "$updated_at", ToDateTimeText(DateTime.Now));
        await command.ExecuteNonQueryAsync();
    }

    private static async Task InsertNegotiationNoteAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long versionId,
        long clientId,
        long renewalId,
        NegotiationNoteRecord note)
    {
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO negotiation_notes (
    renewal_id,
    client_id,
    renewal_version_id,
    round_number,
    note_date,
    carrier_response,
    consultant_notes,
    next_action,
    status,
    created_at,
    created_by
)
VALUES (
    $renewal_id,
    $client_id,
    $renewal_version_id,
    $round_number,
    $note_date,
    $carrier_response,
    $consultant_notes,
    $next_action,
    $status,
    $created_at,
    $created_by
);");
        AddParameter(command, "$renewal_id", renewalId > 0 ? renewalId : null);
        AddParameter(command, "$client_id", clientId);
        AddParameter(command, "$renewal_version_id", versionId);
        AddParameter(command, "$round_number", note.RoundNumber);
        AddParameter(command, "$note_date", ToDateText(note.NoteDate));
        AddParameter(command, "$carrier_response", Clean(note.CarrierResponse));
        AddParameter(command, "$consultant_notes", Clean(note.ConsultantNotes));
        AddParameter(command, "$next_action", Clean(note.NextAction));
        AddParameter(command, "$status", Clean(note.Status));
        AddParameter(command, "$created_at", ToDateTimeText(DateTime.Now));
        AddParameter(command, "$created_by", Clean(note.CreatedBy));
        await command.ExecuteNonQueryAsync();
    }

    private static decimal? PercentDecimalToPercent(decimal? value)
    {
        return value.HasValue ? decimal.Round(value.Value * 100m, 1) : null;
    }

    private static decimal? PercentToDecimal(decimal? value)
    {
        return value.HasValue ? decimal.Round(value.Value / 100m, 6) : null;
    }

    private static decimal? NormalizePercentDecimal(decimal? value)
    {
        if (!value.HasValue)
        {
            return null;
        }

        return Math.Abs(value.Value) > 1m
            ? decimal.Round(value.Value / 100m, 6)
            : value.Value;
    }

    private static void AddUnderwritingCalculationParameters(
        SqliteCommand command,
        long versionId,
        UnderwritingCalculationResult calculation)
    {
        AddParameter(command, "$version_id", versionId);
        AddParameter(command, "$source_module", Clean(calculation.SourceModule));
        AddParameter(command, "$benefit_line", Clean(calculation.BenefitLine));
        AddParameter(command, "$current_claims", calculation.CurrentClaims);
        AddParameter(command, "$current_premium", calculation.CurrentPremium);
        AddParameter(command, "$trended_claims", calculation.TrendedClaims);
        AddParameter(command, "$pooling_adjustment", calculation.PoolingAdjustment);
        AddParameter(command, "$ibnr_adjustment", calculation.IbnrAdjustment);
        AddParameter(command, "$retention_admin_fee_adjustment", calculation.RetentionAdminFeeAdjustment);
        AddParameter(command, "$manual_adjustment", calculation.ManualAdjustment);
        AddParameter(command, "$projected_claims", calculation.ProjectedClaims);
        AddParameter(command, "$required_premium", calculation.RequiredPremium);
        AddParameter(command, "$proposed_premium", calculation.ProposedPremium);
        AddParameter(command, "$negotiated_premium", calculation.NegotiatedPremium);
        AddParameter(command, "$required_rate", calculation.RequiredRate);
        AddParameter(command, "$target_loss_ratio", calculation.TargetLossRatio);
        AddParameter(command, "$projected_loss_ratio", calculation.ProjectedLossRatio);
        AddParameter(command, "$renewal_change_percent", calculation.RenewalChangePercent);
        AddParameter(command, "$negotiated_change_percent", calculation.NegotiatedChangePercent);
        AddParameter(command, "$savings", calculation.Savings);
        AddParameter(command, "$trend", calculation.Trend);
    }

    private static async Task UpsertRenewalAssumptionAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        RenewalAssumptionRecord assumption)
    {
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO renewal_assumptions (
    client_id,
    carrier,
    benefit_line,
    trend,
    target_loss_ratio,
    pooling_charge,
    pooling_threshold,
    ibnr,
    admin_fee,
    risk_charge,
    effective_date
)
VALUES (
    $client_id,
    $carrier,
    $benefit_line,
    $trend,
    $target_loss_ratio,
    $pooling_charge,
    $pooling_threshold,
    $ibnr,
    $admin_fee,
    $risk_charge,
    $effective_date
)
ON CONFLICT(client_id, carrier, benefit_line, effective_date)
DO UPDATE SET
    trend = excluded.trend,
    target_loss_ratio = excluded.target_loss_ratio,
    pooling_charge = excluded.pooling_charge,
    pooling_threshold = excluded.pooling_threshold,
    ibnr = excluded.ibnr,
    admin_fee = excluded.admin_fee,
    risk_charge = excluded.risk_charge;");
        AddParameter(command, "$client_id", assumption.ClientId);
        AddParameter(command, "$carrier", Clean(assumption.Carrier));
        AddParameter(command, "$benefit_line", Clean(assumption.BenefitLine));
        AddParameter(command, "$trend", assumption.Trend);
        AddParameter(command, "$target_loss_ratio", assumption.TargetLossRatio);
        AddParameter(command, "$pooling_charge", assumption.PoolingCharge);
        AddParameter(command, "$pooling_threshold", assumption.PoolingThreshold);
        AddParameter(command, "$ibnr", assumption.Ibnr);
        AddParameter(command, "$admin_fee", assumption.AdminFee);
        AddParameter(command, "$risk_charge", assumption.RiskCharge);
        AddParameter(command, "$effective_date", ToDateText(assumption.EffectiveDate));
        await command.ExecuteNonQueryAsync();
    }

    private static async Task UpsertCarrierAssumptionAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        CarrierAssumptionRecord assumption)
    {
        var carrier = Clean(assumption.Carrier);
        var blockName = Clean(assumption.BlockName);
        var benefitLine = Clean(assumption.BenefitLine);
        var effectiveDate = ToDateText(assumption.EffectiveDate);

        await ExecuteNonQueryAsync(
            connection,
            transaction,
            @"DELETE FROM carrier_assumptions
              WHERE carrier = $carrier
                AND block_name = $block_name
                AND benefit_line = $benefit_line
                AND effective_date = $effective_date;",
            command =>
            {
                AddParameter(command, "$carrier", carrier);
                AddParameter(command, "$block_name", blockName);
                AddParameter(command, "$benefit_line", benefitLine);
                AddParameter(command, "$effective_date", effectiveDate);
            });

        using var insert = CreateCommand(connection, transaction, @"
INSERT INTO carrier_assumptions (
    carrier,
    block_name,
    benefit_line,
    funding_type,
    assumption_name,
    assumption_value,
    target_loss_ratio,
    tlr_basis,
    premium_tax,
    commission,
    admin_margin,
    risk_margin,
    trend,
    credibility_factor,
    pooling_threshold,
    pooling_charge,
    reserve_ibnr,
    effective_date,
    status,
    notes,
    date_loaded,
    updated_at
)
VALUES (
    $carrier,
    $block_name,
    $benefit_line,
    $funding_type,
    $assumption_name,
    $assumption_value,
    $target_loss_ratio,
    $tlr_basis,
    $premium_tax,
    $commission,
    $admin_margin,
    $risk_margin,
    $trend,
    $credibility_factor,
    $pooling_threshold,
    $pooling_charge,
    $reserve_ibnr,
    $effective_date,
    $status,
    $notes,
    $date_loaded,
    $updated_at
);");
        AddParameter(insert, "$carrier", carrier);
        AddParameter(insert, "$block_name", blockName);
        AddParameter(insert, "$benefit_line", benefitLine);
        AddParameter(insert, "$funding_type", Clean(assumption.FundingType));
        AddParameter(insert, "$assumption_name", "Target Loss Ratio");
        AddParameter(insert, "$assumption_value", assumption.TargetLossRatio);
        AddParameter(insert, "$target_loss_ratio", assumption.TargetLossRatio);
        AddParameter(insert, "$tlr_basis", TlrBasisOptions.Normalize(assumption.TlrBasis));
        AddParameter(insert, "$premium_tax", assumption.PremiumTax);
        AddParameter(insert, "$commission", assumption.Commission);
        AddParameter(insert, "$admin_margin", assumption.AdminMargin);
        AddParameter(insert, "$risk_margin", assumption.RiskMargin);
        AddParameter(insert, "$trend", assumption.Trend);
        AddParameter(insert, "$credibility_factor", assumption.CredibilityFactor);
        AddParameter(insert, "$pooling_threshold", assumption.PoolingThreshold);
        AddParameter(insert, "$pooling_charge", assumption.PoolingCharge);
        AddParameter(insert, "$reserve_ibnr", assumption.ReserveIbnr);
        AddParameter(insert, "$effective_date", effectiveDate);
        AddParameter(insert, "$status", FirstNonBlank(assumption.Status, "Active"));
        AddParameter(insert, "$notes", Clean(assumption.Notes));
        AddParameter(insert, "$date_loaded", ToDateTimeText(assumption.DateLoaded == DateTime.MinValue ? DateTime.Now : assumption.DateLoaded));
        AddParameter(insert, "$updated_at", ToDateTimeText(DateTime.Now));
        await insert.ExecuteNonQueryAsync();
    }

    private static async Task UpsertRenewalProjectionAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        RenewalProjectionRecord projection)
    {
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO renewal_projections (
    client_id,
    benefit_line,
    current_premium,
    projected_claims,
    trend_adjustment,
    required_premium,
    proposed_premium,
    negotiated_premium,
    renewal_change_percent,
    savings
)
VALUES (
    $client_id,
    $benefit_line,
    $current_premium,
    $projected_claims,
    $trend_adjustment,
    $required_premium,
    $proposed_premium,
    $negotiated_premium,
    $renewal_change_percent,
    $savings
)
ON CONFLICT(client_id, benefit_line)
DO UPDATE SET
    current_premium = excluded.current_premium,
    projected_claims = excluded.projected_claims,
    trend_adjustment = excluded.trend_adjustment,
    required_premium = excluded.required_premium,
    proposed_premium = excluded.proposed_premium,
    negotiated_premium = excluded.negotiated_premium,
    renewal_change_percent = excluded.renewal_change_percent,
    savings = excluded.savings;");
        AddParameter(command, "$client_id", projection.ClientId);
        AddParameter(command, "$benefit_line", Clean(projection.BenefitLine));
        AddParameter(command, "$current_premium", projection.CurrentPremium);
        AddParameter(command, "$projected_claims", projection.ProjectedClaims);
        AddParameter(command, "$trend_adjustment", projection.TrendAdjustment);
        AddParameter(command, "$required_premium", projection.RequiredPremium);
        AddParameter(command, "$proposed_premium", projection.ProposedPremium);
        AddParameter(command, "$negotiated_premium", projection.NegotiatedPremium);
        AddParameter(command, "$renewal_change_percent", projection.RenewalChangePercent);
        AddParameter(command, "$savings", projection.Savings);
        await command.ExecuteNonQueryAsync();
    }

    private static async Task InsertClaimsExperienceAsync(SqliteConnection connection, SqliteTransaction transaction, ClaimsExperienceRecord claims)
    {
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO ClaimsExperience (
    client_id,
    experience_period_start,
    experience_period_end,
    benefit_line,
    premium,
    paid_claims,
    incurred_claims,
    loss_ratio,
    lives,
    source_document_id,
    date_loaded
)
VALUES (
    $client_id,
    NULLIF($experience_period_start, ''),
    NULLIF($experience_period_end, ''),
    NULLIF($benefit_line, ''),
    $premium,
    $paid_claims,
    $incurred_claims,
    $loss_ratio,
    $lives,
    $source_document_id,
    $date_loaded
);");
        AddParameter(command, "$client_id", claims.ClientId);
        AddParameter(command, "$experience_period_start", ToDateText(claims.ExperiencePeriodStart));
        AddParameter(command, "$experience_period_end", ToDateText(claims.ExperiencePeriodEnd));
        AddParameter(command, "$benefit_line", Clean(claims.BenefitLine));
        AddParameter(command, "$premium", claims.Premium);
        AddParameter(command, "$paid_claims", claims.PaidClaims);
        AddParameter(command, "$incurred_claims", claims.IncurredClaims);
        AddParameter(command, "$loss_ratio", claims.LossRatio);
        AddParameter(command, "$lives", claims.Lives);
        AddParameter(command, "$source_document_id", claims.SourceDocumentId);
        AddParameter(command, "$date_loaded", ToDateTimeText(claims.DateLoaded));
        await command.ExecuteNonQueryAsync();
    }

    private static async Task InsertMonthlyClaimsAsync(SqliteConnection connection, SqliteTransaction transaction, MonthlyClaimsRecord monthlyClaim)
    {
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO MonthlyClaims (
    client_id,
    carrier,
    benefit_line,
    month,
    lives,
    premium,
    paid_claims,
    incurred_claims,
    claims_plus_fees,
    loss_ratio,
    source_document_id,
    date_loaded
)
VALUES (
    $client_id,
    NULLIF($carrier, ''),
    NULLIF($benefit_line, ''),
    NULLIF($month, ''),
    $lives,
    $premium,
    $paid_claims,
    $incurred_claims,
    $claims_plus_fees,
    $loss_ratio,
    $source_document_id,
    $date_loaded
);");
        AddParameter(command, "$client_id", monthlyClaim.ClientId);
        AddParameter(command, "$carrier", Clean(monthlyClaim.Carrier));
        AddParameter(command, "$benefit_line", Clean(monthlyClaim.BenefitLine));
        AddParameter(command, "$month", ToDateText(monthlyClaim.Month));
        AddParameter(command, "$lives", monthlyClaim.Lives);
        AddParameter(command, "$premium", monthlyClaim.Premium);
        AddParameter(command, "$paid_claims", monthlyClaim.PaidClaims);
        AddParameter(command, "$incurred_claims", monthlyClaim.IncurredClaims);
        AddParameter(command, "$claims_plus_fees", monthlyClaim.ClaimsPlusFees);
        AddParameter(command, "$loss_ratio", monthlyClaim.LossRatio);
        AddParameter(command, "$source_document_id", monthlyClaim.SourceDocumentId);
        AddParameter(command, "$date_loaded", ToDateTimeText(monthlyClaim.DateLoaded));
        await command.ExecuteNonQueryAsync();
    }

    private static async Task InsertPlanDesignAsync(SqliteConnection connection, SqliteTransaction transaction, PlanDesignRecord planDesign)
    {
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO PlanDesign (
    client_id,
    class_name,
    benefit_category,
    benefit_name,
    coverage_description,
    reimbursement_percent,
    maximum_amount,
    frequency,
    waiting_period,
    termination_age,
    booklet_source_file_name,
    source_section,
    notes,
    source_document_id,
    source_page,
    date_loaded
)
VALUES (
    $client_id,
    $class_name,
    NULLIF($benefit_category, ''),
    NULLIF($benefit_name, ''),
    NULLIF($coverage_description, ''),
    $reimbursement_percent,
    NULLIF($maximum_amount, ''),
    NULLIF($frequency, ''),
    NULLIF($waiting_period, ''),
    NULLIF($termination_age, ''),
    $booklet_source_file_name,
    $source_section,
    $notes,
    $source_document_id,
    $source_page,
    $date_loaded
);");
        AddParameter(command, "$client_id", planDesign.ClientId);
        AddParameter(command, "$class_name", Clean(planDesign.ClassName));
        AddParameter(command, "$benefit_category", Clean(planDesign.BenefitCategory));
        AddParameter(command, "$benefit_name", Clean(planDesign.BenefitName));
        AddParameter(command, "$coverage_description", Clean(planDesign.CoverageDescription));
        AddParameter(command, "$reimbursement_percent", planDesign.ReimbursementPercent);
        AddParameter(command, "$maximum_amount", Clean(planDesign.MaximumAmount));
        AddParameter(command, "$frequency", Clean(planDesign.Frequency));
        AddParameter(command, "$waiting_period", Clean(planDesign.WaitingPeriod));
        AddParameter(command, "$termination_age", Clean(planDesign.TerminationAge));
        AddParameter(command, "$booklet_source_file_name", Clean(planDesign.SourceFileName));
        AddParameter(command, "$source_section", Clean(planDesign.SourceSection));
        AddParameter(command, "$notes", Clean(planDesign.Notes));
        AddParameter(command, "$source_document_id", planDesign.SourceDocumentId);
        AddParameter(command, "$source_page", planDesign.SourcePage);
        AddParameter(command, "$date_loaded", ToDateTimeText(planDesign.DateLoaded));
        await command.ExecuteNonQueryAsync();
    }

    private static async Task<long> InsertPlanDesignFeatureAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long clientId,
        long? sourceDocumentId,
        PlanDesignFeatureExtraction feature)
    {
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO plan_design_features (
    client_id,
    source_document_id,
    benefit,
    category,
    feature_name,
    feature_value,
    coinsurance,
    deductible,
    maximum,
    frequency,
    termination_age,
    class_number,
    class_name,
    source_page,
    source_text,
    confidence_score,
    date_loaded
)
VALUES (
    $client_id,
    $source_document_id,
    $benefit,
    $category,
    $feature_name,
    $feature_value,
    $coinsurance,
    $deductible,
    $maximum,
    $frequency,
    $termination_age,
    $class_number,
    $class_name,
    $source_page,
    $source_text,
    $confidence_score,
    $date_loaded
);");
        AddParameter(command, "$client_id", clientId);
        AddParameter(command, "$source_document_id", sourceDocumentId);
        AddParameter(command, "$benefit", feature.Benefit);
        AddParameter(command, "$category", feature.Category);
        AddParameter(command, "$feature_name", feature.FeatureName);
        AddParameter(command, "$feature_value", feature.FeatureValue);
        AddParameter(command, "$coinsurance", feature.Coinsurance);
        AddParameter(command, "$deductible", feature.Deductible);
        AddParameter(command, "$maximum", feature.Maximum);
        AddParameter(command, "$frequency", feature.Frequency);
        AddParameter(command, "$termination_age", feature.TerminationAge);
        AddParameter(command, "$class_number", feature.ClassNumber);
        AddParameter(command, "$class_name", feature.ClassName);
        AddParameter(command, "$source_page", feature.SourcePage);
        AddParameter(command, "$source_text", feature.SourceText);
        AddParameter(command, "$confidence_score", feature.ConfidenceScore);
        AddParameter(command, "$date_loaded", ToDateTimeText(DateTime.Now));
        await command.ExecuteNonQueryAsync();

        using var identityCommand = CreateCommand(connection, transaction, "SELECT last_insert_rowid();");
        return Convert.ToInt64(await identityCommand.ExecuteScalarAsync(), CultureInfo.InvariantCulture);
    }

    private static async Task InsertPlanBenefitRuleAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        PlanBenefitRule rule)
    {
        if (!rule.PlanDesignFeatureId.HasValue || !rule.ClientId.HasValue)
        {
            throw new InvalidOperationException("A plan benefit rule requires its parent feature and client before it can be saved.");
        }

        using var command = CreateCommand(connection, transaction, """
            INSERT INTO PlanBenefitRule (
                PlanDesignFeatureId,
                ClientId,
                SourceDocumentId,
                RuleType,
                Tier1ReimbursementPercent,
                Tier1ThresholdAmount,
                Tier2ReimbursementPercent,
                ThresholdBasis,
                ResetPeriod,
                CurrencyCode,
                SourceText,
                SourceLocation,
                NormalizationStatus,
                ConfidenceScore,
                DisplayOrder,
                IsActive,
                EffectiveDate,
                TerminationDate,
                CreatedAtUtc,
                UpdatedAtUtc
            )
            VALUES (
                $plan_design_feature_id,
                $client_id,
                $source_document_id,
                $rule_type,
                $tier1_percent,
                $tier1_threshold,
                $tier2_percent,
                $threshold_basis,
                $reset_period,
                $currency_code,
                $source_text,
                $source_location,
                $normalization_status,
                $confidence_score,
                $display_order,
                $is_active,
                $effective_date,
                $termination_date,
                $created_at,
                $updated_at
            );
            """);
        AddParameter(command, "$plan_design_feature_id", rule.PlanDesignFeatureId.Value);
        AddParameter(command, "$client_id", rule.ClientId.Value);
        AddParameter(command, "$source_document_id", rule.SourceDocumentId);
        AddParameter(command, "$rule_type", rule.RuleType.ToString());
        AddParameter(command, "$tier1_percent", rule.Tier1ReimbursementPercent);
        AddParameter(command, "$tier1_threshold", rule.Tier1ThresholdAmount);
        AddParameter(command, "$tier2_percent", rule.Tier2ReimbursementPercent);
        AddParameter(command, "$threshold_basis", rule.ThresholdBasis?.ToString());
        AddParameter(command, "$reset_period", rule.ResetPeriod?.ToString());
        AddParameter(command, "$currency_code", Clean(rule.CurrencyCode));
        AddParameter(command, "$source_text", Clean(rule.SourceText));
        AddParameter(command, "$source_location", Clean(rule.SourceLocation));
        AddParameter(command, "$normalization_status", rule.NormalizationStatus.ToString());
        AddParameter(command, "$confidence_score", rule.ConfidenceScore);
        AddParameter(command, "$display_order", rule.DisplayOrder);
        AddParameter(command, "$is_active", rule.IsActive ? 1 : 0);
        AddParameter(command, "$effective_date", rule.EffectiveDate?.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));
        AddParameter(command, "$termination_date", rule.TerminationDate?.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));
        AddParameter(command, "$created_at", rule.CreatedAtUtc.ToString("O", CultureInfo.InvariantCulture));
        AddParameter(command, "$updated_at", rule.UpdatedAtUtc.ToString("O", CultureInfo.InvariantCulture));
        await command.ExecuteNonQueryAsync();
    }

    private static PlanDesignFeatureExtraction BuildPlanDesignFeature(
        PlanDesignRecord row,
        decimal? guardedReimbursementPercent,
        string planDesignSource = PlanDesignSources.Booklet)
    {
        var notes = Clean(row.Notes);
        var (featureName, featureValue) = ParsePlanDesignFeature(row);
        var benefit = FirstNonBlank(row.BenefitName, row.BenefitCategory);
        var category = FirstNonBlank(row.BenefitCategory, row.BenefitName);
        var isDeductible = featureName.Equals("Deductible", StringComparison.OrdinalIgnoreCase);
        return new PlanDesignFeatureExtraction
        {
            Benefit = benefit,
            Category = category,
            FeatureName = FirstNonBlank(row.ProvisionLabel, featureName),
            FeatureValue = featureValue,
            Coinsurance = FirstNonBlank(
                row.ReimbursementText,
                guardedReimbursementPercent.HasValue ? $"{guardedReimbursementPercent.Value:0.##}%" : ""),
            Deductible = isDeductible ? featureValue : "",
            Maximum = row.MaximumAmount,
            Frequency = row.Frequency,
            TerminationAge = row.TerminationAge,
            ClassName = row.ClassName,
            SourcePage = row.SourcePage,
            SourceText = Clean($"{row.CoverageDescription} {notes}"),
            SourceFileName = row.SourceFileName,
            PlanDesignSource = planDesignSource,
            ConfidenceScore = notes.Contains("low confidence", StringComparison.OrdinalIgnoreCase) ? 0.55m : 0.80m
        };
    }

    private static (string FeatureName, string FeatureValue) ParsePlanDesignFeature(PlanDesignRecord row)
    {
        var description = Clean(row.CoverageDescription);
        var labelledValue = Regex.Match(
            description,
            @"^(?<label>[A-Za-z][A-Za-z0-9 &/()'.+\-]{1,100}):\s*(?<value>[^\r\n]+)",
            RegexOptions.IgnoreCase);
        if (labelledValue.Success)
        {
            return (
                Clean(labelledValue.Groups["label"].Value),
                Clean(labelledValue.Groups["value"].Value));
        }

        return (
            FirstNonBlank(row.ProvisionLabel, row.BenefitName, row.BenefitCategory),
            FirstNonBlank(description, row.MaximumAmount, row.Frequency, row.WaitingPeriod, row.TerminationAge));
    }

    private static async Task InsertBillingOptionPremiumAsync(SqliteConnection connection, SqliteTransaction transaction, BillingOptionPremiumDetail row)
    {
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO billing_option_premiums (
    statement_version_id,
    client_id,
    source_document_id,
    billing_period,
    province,
    benefit,
    class_code,
    option_label,
    dependent_code,
    lives,
    volume,
    monthly_premium,
    premium_adjustments,
    total_premium_adjustments,
    source_page,
    source_file,
    confidence,
    date_loaded
)
VALUES (
    $statement_version_id,
    $client_id,
    $source_document_id,
    $billing_period,
    $province,
    $benefit,
    $class_code,
    $option_label,
    $dependent_code,
    $lives,
    $volume,
    $monthly_premium,
    $premium_adjustments,
    $total_premium_adjustments,
    $source_page,
    $source_file,
    $confidence,
    $date_loaded
);");
        AddParameter(command, "$statement_version_id", row.StatementVersionId);
        AddParameter(command, "$client_id", row.ClientId);
        AddParameter(command, "$source_document_id", row.SourceDocumentId);
        AddParameter(command, "$billing_period", Clean(row.BillingPeriod));
        AddParameter(command, "$province", Clean(row.Province));
        AddParameter(command, "$benefit", Clean(row.Benefit));
        AddParameter(command, "$class_code", Clean(row.ClassCode));
        AddParameter(command, "$option_label", Clean(row.OptionLabel));
        AddParameter(command, "$dependent_code", Clean(row.DependentCode));
        AddParameter(command, "$lives", row.Lives);
        AddParameter(command, "$volume", row.Volume);
        AddParameter(command, "$monthly_premium", row.MonthlyPremium);
        AddParameter(command, "$premium_adjustments", row.PremiumAdjustments);
        AddParameter(command, "$total_premium_adjustments", row.TotalPremiumAdjustments);
        AddParameter(command, "$source_page", row.SourcePage);
        AddParameter(command, "$source_file", Clean(row.SourceFile));
        AddParameter(command, "$confidence", row.Confidence);
        AddParameter(command, "$date_loaded", ToDateTimeText(DateTime.Now));
        await command.ExecuteNonQueryAsync();
    }

    private static async Task InsertBillingProvinceSummaryAsync(SqliteConnection connection, SqliteTransaction transaction, BillingProvinceSummary row)
    {
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO billing_province_summary (
    statement_version_id,
    client_id,
    source_document_id,
    billing_period,
    province,
    total_lives,
    total_monthly_premium,
    estimated_annual_premium,
    total_tax,
    total_premium_with_tax,
    source_file,
    date_loaded
)
VALUES (
    $statement_version_id,
    $client_id,
    $source_document_id,
    $billing_period,
    $province,
    $total_lives,
    $total_monthly_premium,
    $estimated_annual_premium,
    $total_tax,
    $total_premium_with_tax,
    $source_file,
    $date_loaded
);");
        AddParameter(command, "$statement_version_id", row.StatementVersionId);
        AddParameter(command, "$client_id", row.ClientId);
        AddParameter(command, "$source_document_id", row.SourceDocumentId);
        AddParameter(command, "$billing_period", Clean(row.BillingPeriod));
        AddParameter(command, "$province", Clean(row.Province));
        AddParameter(command, "$total_lives", row.TotalLives);
        AddParameter(command, "$total_monthly_premium", row.TotalMonthlyPremium);
        AddParameter(command, "$estimated_annual_premium", row.EstimatedAnnualPremium);
        AddParameter(command, "$total_tax", row.TotalTax);
        AddParameter(command, "$total_premium_with_tax", row.TotalPremiumWithTax);
        AddParameter(command, "$source_file", Clean(row.SourceFile));
        AddParameter(command, "$date_loaded", ToDateTimeText(DateTime.Now));
        await command.ExecuteNonQueryAsync();
    }

    private static async Task InsertBillingOptionSummaryAsync(SqliteConnection connection, SqliteTransaction transaction, BillingOptionSummary row)
    {
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO billing_option_summary (
    statement_version_id,
    client_id,
    source_document_id,
    billing_period,
    province,
    benefit,
    class_code,
    option_label,
    lives,
    monthly_premium,
    estimated_annual_premium,
    source_file,
    date_loaded
)
VALUES (
    $statement_version_id,
    $client_id,
    $source_document_id,
    $billing_period,
    $province,
    $benefit,
    $class_code,
    $option_label,
    $lives,
    $monthly_premium,
    $estimated_annual_premium,
    $source_file,
    $date_loaded
);");
        AddParameter(command, "$statement_version_id", row.StatementVersionId);
        AddParameter(command, "$client_id", row.ClientId);
        AddParameter(command, "$source_document_id", row.SourceDocumentId);
        AddParameter(command, "$billing_period", Clean(row.BillingPeriod));
        AddParameter(command, "$province", Clean(row.Province));
        AddParameter(command, "$benefit", Clean(row.Benefit));
        AddParameter(command, "$class_code", Clean(row.ClassCode));
        AddParameter(command, "$option_label", Clean(row.OptionLabel));
        AddParameter(command, "$lives", row.Lives);
        AddParameter(command, "$monthly_premium", row.MonthlyPremium);
        AddParameter(command, "$estimated_annual_premium", row.EstimatedAnnualPremium);
        AddParameter(command, "$source_file", Clean(row.SourceFile));
        AddParameter(command, "$date_loaded", ToDateTimeText(DateTime.Now));
        await command.ExecuteNonQueryAsync();
    }

    private static async Task UpsertBillingClassMappingAsync(SqliteConnection connection, SqliteTransaction transaction, BillingClassMapping row)
    {
        var now = DateTime.Now;
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO ClassMapping (
    client_id,
    carrier,
    class_code,
    option_label,
    first_seen_billing_period,
    last_seen_billing_period,
    source_file,
    created_at,
    updated_at
)
VALUES (
    $client_id,
    $carrier,
    $class_code,
    $option_label,
    $first_seen_billing_period,
    $last_seen_billing_period,
    $source_file,
    $created_at,
    $updated_at
)
ON CONFLICT(client_id, carrier, class_code) DO UPDATE SET
    option_label = excluded.option_label,
    last_seen_billing_period = excluded.last_seen_billing_period,
    source_file = excluded.source_file,
    updated_at = excluded.updated_at;");
        AddParameter(command, "$client_id", row.ClientId);
        AddParameter(command, "$carrier", Clean(row.Carrier));
        AddParameter(command, "$class_code", Clean(row.ClassCode));
        AddParameter(command, "$option_label", Clean(row.OptionLabel));
        AddParameter(command, "$first_seen_billing_period", Clean(row.FirstSeenBillingPeriod));
        AddParameter(command, "$last_seen_billing_period", Clean(row.LastSeenBillingPeriod));
        AddParameter(command, "$source_file", Clean(row.SourceFile));
        AddParameter(command, "$created_at", ToDateTimeText(row.CreatedAt == default ? now : row.CreatedAt));
        AddParameter(command, "$updated_at", ToDateTimeText(now));
        await command.ExecuteNonQueryAsync();
    }

    private static async Task InsertBillingReconciliationResultAsync(SqliteConnection connection, SqliteTransaction transaction, BillingReconciliationResult row)
    {
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO BillingReconciliationResults (
    statement_version_id,
    client_id,
    source_document_id,
    billing_period,
    source_file,
    reconciliation_level,
    province,
    benefit,
    class_code,
    pdf_lives,
    calculated_lives,
    lives_difference,
    lives_status,
    lives_warning_message,
    pdf_premium,
    calculated_premium,
    premium_difference,
    pdf_tax,
    calculated_tax,
    tax_difference,
    pdf_total_with_tax,
    calculated_total_with_tax,
    total_difference,
    status,
    warning_message,
    source_page,
    created_at
)
VALUES (
    $statement_version_id,
    $client_id,
    $source_document_id,
    $billing_period,
    $source_file,
    $reconciliation_level,
    $province,
    $benefit,
    $class_code,
    $pdf_lives,
    $calculated_lives,
    $lives_difference,
    $lives_status,
    $lives_warning_message,
    $pdf_premium,
    $calculated_premium,
    $premium_difference,
    $pdf_tax,
    $calculated_tax,
    $tax_difference,
    $pdf_total_with_tax,
    $calculated_total_with_tax,
    $total_difference,
    $status,
    $warning_message,
    $source_page,
    $created_at
);");
        AddParameter(command, "$statement_version_id", row.StatementVersionId);
        AddParameter(command, "$client_id", row.ClientId);
        AddParameter(command, "$source_document_id", row.SourceDocumentId);
        AddParameter(command, "$billing_period", Clean(row.BillingPeriod));
        AddParameter(command, "$source_file", Clean(row.SourceFile));
        AddParameter(command, "$reconciliation_level", Clean(row.ReconciliationLevel));
        AddParameter(command, "$province", Clean(row.Province));
        AddParameter(command, "$benefit", Clean(row.Benefit));
        AddParameter(command, "$class_code", Clean(row.ClassCode));
        AddParameter(command, "$pdf_lives", row.PdfLives);
        AddParameter(command, "$calculated_lives", row.CalculatedLives);
        AddParameter(command, "$lives_difference", row.LivesDifference);
        AddParameter(command, "$lives_status", Clean(row.LivesStatus));
        AddParameter(command, "$lives_warning_message", Clean(row.LivesWarningMessage));
        AddParameter(command, "$pdf_premium", row.PdfPremium);
        AddParameter(command, "$calculated_premium", row.CalculatedPremium);
        AddParameter(command, "$premium_difference", row.PremiumDifference);
        AddParameter(command, "$pdf_tax", row.PdfTax);
        AddParameter(command, "$calculated_tax", row.CalculatedTax);
        AddParameter(command, "$tax_difference", row.TaxDifference);
        AddParameter(command, "$pdf_total_with_tax", row.PdfTotalWithTax);
        AddParameter(command, "$calculated_total_with_tax", row.CalculatedTotalWithTax);
        AddParameter(command, "$total_difference", row.TotalDifference);
        AddParameter(command, "$status", Clean(row.Status));
        AddParameter(command, "$warning_message", Clean(row.WarningMessage));
        AddParameter(command, "$source_page", row.SourcePage);
        AddParameter(command, "$created_at", ToDateTimeText(row.CreatedAt == default ? DateTime.Now : row.CreatedAt));
        await command.ExecuteNonQueryAsync();
    }

    private static async Task<long?> FindGroupHealthBillingStatementAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long clientId,
        string carrier,
        int statementMonth,
        int statementYear,
        string sourceFileHash)
    {
        using var command = CreateCommand(
            connection,
            transaction,
            @"
SELECT billing_statement_id
FROM BillingStatements
WHERE client_id = $client_id
  AND carrier = $carrier
  AND statement_month = $statement_month
  AND statement_year = $statement_year
  AND source_file_hash = $source_file_hash
  AND is_deleted = 0
LIMIT 1;");
        AddParameter(command, "$client_id", clientId);
        AddParameter(command, "$carrier", Clean(carrier));
        AddParameter(command, "$statement_month", statementMonth);
        AddParameter(command, "$statement_year", statementYear);
        AddParameter(command, "$source_file_hash", Clean(sourceFileHash));
        return ToNullableLong(await command.ExecuteScalarAsync());
    }

    private static async Task<List<long>> GetActiveGroupHealthStatementIdsForMonthAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long clientId,
        string carrier,
        int statementMonth,
        int statementYear)
    {
        using var command = CreateCommand(
            connection,
            transaction,
            @"
SELECT billing_statement_id
FROM BillingStatements
WHERE client_id = $client_id
  AND carrier = $carrier
  AND statement_month = $statement_month
  AND statement_year = $statement_year
  AND is_replaced = 0
  AND is_deleted = 0;");
        AddParameter(command, "$client_id", clientId);
        AddParameter(command, "$carrier", Clean(carrier));
        AddParameter(command, "$statement_month", statementMonth);
        AddParameter(command, "$statement_year", statementYear);

        var ids = new List<long>();
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            ids.Add(reader.GetInt64(0));
        }

        return ids;
    }

    private static async Task<int> GetMaxGroupHealthImportVersionAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long clientId,
        string carrier,
        int statementMonth,
        int statementYear)
    {
        using var command = CreateCommand(
            connection,
            transaction,
            @"
SELECT COALESCE(MAX(import_version), 0)
FROM BillingStatements
WHERE client_id = $client_id
  AND carrier = $carrier
  AND statement_month = $statement_month
  AND statement_year = $statement_year
  AND is_deleted = 0;");
        AddParameter(command, "$client_id", clientId);
        AddParameter(command, "$carrier", Clean(carrier));
        AddParameter(command, "$statement_month", statementMonth);
        AddParameter(command, "$statement_year", statementYear);
        return Convert.ToInt32(await command.ExecuteScalarAsync(), CultureInfo.InvariantCulture);
    }

    private static async Task MarkGroupHealthStatementsReplacedAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        IReadOnlyCollection<long> statementIds)
    {
        foreach (var statementId in statementIds)
        {
            await ExecuteNonQueryAsync(
                connection,
                transaction,
                @"
UPDATE BillingStatements
SET is_replaced = 1,
    replaced_at = $replaced_at,
    updated_at = $updated_at
WHERE billing_statement_id = $billing_statement_id;",
                command =>
                {
                    AddParameter(command, "$replaced_at", ToDateTimeText(DateTime.Now));
                    AddParameter(command, "$updated_at", ToDateTimeText(DateTime.Now));
                    AddParameter(command, "$billing_statement_id", statementId);
                });
        }
    }

    private static async Task<List<GroupHealthBillingRow>> GetPreviousGroupHealthBillingRowsAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long clientId,
        string carrier,
        int statementMonth,
        int statementYear)
    {
        using var statementCommand = CreateCommand(
            connection,
            transaction,
            @"
SELECT billing_statement_id
FROM BillingStatements
WHERE client_id = $client_id
  AND carrier = $carrier
  AND is_replaced = 0
  AND is_deleted = 0
  AND ((statement_year * 12) + statement_month) < (($statement_year * 12) + $statement_month)
ORDER BY statement_year DESC, statement_month DESC, import_version DESC
LIMIT 1;");
        AddParameter(statementCommand, "$client_id", clientId);
        AddParameter(statementCommand, "$carrier", Clean(carrier));
        AddParameter(statementCommand, "$statement_month", statementMonth);
        AddParameter(statementCommand, "$statement_year", statementYear);
        var previousStatementId = ToNullableLong(await statementCommand.ExecuteScalarAsync());
        if (previousStatementId is null)
        {
            return [];
        }

        using var rowsCommand = CreateCommand(
            connection,
            transaction,
            "SELECT * FROM BillingStatementRows WHERE billing_statement_id = $billing_statement_id;");
        AddParameter(rowsCommand, "$billing_statement_id", previousStatementId.Value);
        var rows = new List<GroupHealthBillingRow>();
        using var reader = await rowsCommand.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            rows.Add(ReadGroupHealthBillingRow(reader));
        }

        return rows;
    }

    private static async Task<long> InsertGroupHealthBillingStatementAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        GroupHealthBillingExtractionResult result,
        long clientId,
        long? sourceDocumentId)
    {
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO BillingStatements (
    client_id,
    source_document_id,
    carrier,
    statement_month,
    statement_year,
    statement_date,
    import_date,
    source_file_name,
    source_file_path,
    source_file_hash,
    import_version,
    import_status,
    billing_integrity_score,
    employee_count,
    total_employee_premium,
    total_employer_premium,
    total_premium,
    total_gst_hst,
    statement_total,
    created_at,
    updated_at
)
VALUES (
    $client_id,
    $source_document_id,
    $carrier,
    $statement_month,
    $statement_year,
    $statement_date,
    $import_date,
    $source_file_name,
    $source_file_path,
    $source_file_hash,
    $import_version,
    $import_status,
    $billing_integrity_score,
    $employee_count,
    $total_employee_premium,
    $total_employer_premium,
    $total_premium,
    $total_gst_hst,
    $statement_total,
    $created_at,
    $updated_at
);");
        var now = DateTime.Now;
        AddParameter(command, "$client_id", clientId);
        AddParameter(command, "$source_document_id", sourceDocumentId);
        AddParameter(command, "$carrier", Clean(result.Carrier));
        AddParameter(command, "$statement_month", result.StatementMonth);
        AddParameter(command, "$statement_year", result.StatementYear);
        AddParameter(command, "$statement_date", ToDateText(result.StatementDate));
        AddParameter(command, "$import_date", ToDateTimeText(result.ImportDate == default ? now : result.ImportDate));
        AddParameter(command, "$source_file_name", Clean(result.SourceFileName));
        AddParameter(command, "$source_file_path", Clean(result.SourceFilePath));
        AddParameter(command, "$source_file_hash", Clean(result.SourceFileHash));
        AddParameter(command, "$import_version", result.ImportVersion);
        AddParameter(command, "$import_status", Clean(result.ImportStatus));
        AddParameter(command, "$billing_integrity_score", result.BillingIntegrityScore);
        AddParameter(command, "$employee_count", result.IntegrityReport.TotalEmployeesImported);
        AddParameter(command, "$total_employee_premium", result.Rows.Sum(row => row.EmployeePremium ?? 0m));
        AddParameter(command, "$total_employer_premium", result.Rows.Sum(row => (row.EmployerPremium ?? 0m) + (row.GstHst ?? 0m)));
        AddParameter(command, "$total_premium", result.PremiumGrandTotal ?? result.Rows.Sum(row => row.TotalPremium ?? 0m));
        AddParameter(command, "$total_gst_hst", result.GstHstGrandTotal ?? result.Rows.Sum(row => row.GstHst ?? 0m));
        AddParameter(command, "$statement_total", result.StatementGrandTotal ?? result.Rows.Sum(row => row.RowTotal ?? 0m));
        AddParameter(command, "$created_at", ToDateTimeText(now));
        AddParameter(command, "$updated_at", ToDateTimeText(now));
        await command.ExecuteNonQueryAsync();

        using var idCommand = CreateCommand(connection, transaction, "SELECT last_insert_rowid();");
        return ToLong(await idCommand.ExecuteScalarAsync());
    }

    private static async Task<long> InsertGroupHealthBillingRowAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        GroupHealthBillingRow row)
    {
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO BillingStatementRows (
    billing_statement_id, client_id, class_code, employee_id, certificate_id, employee_name,
    annual_earnings, cost_division, employment_class, retro_month,
    life_volume, life_premium, adnd_volume, adnd_premium, ltd_volume, ltd_premium,
    health_tier, health_premium, dental_tier, dental_premium, efap_premium, virtual_healthcare_premium,
    flex_admin_fee, member_adnd_volume, member_adnd_premium, member_adnd_sp_volume, member_adnd_sp_premium,
    member_adnd_ch_volume, member_adnd_ch_premium, employee_premium, employer_premium, total_premium,
    gst_hst, retro_premium, retro_gst_hst, row_total, source_page, source_line,
    validation_status, validation_flags, source_text, date_loaded
)
VALUES (
    $billing_statement_id, $client_id, $class_code, $employee_id, $certificate_id, $employee_name,
    $annual_earnings, $cost_division, $employment_class, $retro_month,
    $life_volume, $life_premium, $adnd_volume, $adnd_premium, $ltd_volume, $ltd_premium,
    $health_tier, $health_premium, $dental_tier, $dental_premium, $efap_premium, $virtual_healthcare_premium,
    $flex_admin_fee, $member_adnd_volume, $member_adnd_premium, $member_adnd_sp_volume, $member_adnd_sp_premium,
    $member_adnd_ch_volume, $member_adnd_ch_premium, $employee_premium, $employer_premium, $total_premium,
    $gst_hst, $retro_premium, $retro_gst_hst, $row_total, $source_page, $source_line,
    $validation_status, $validation_flags, $source_text, $date_loaded
);");
        AddParameter(command, "$billing_statement_id", row.BillingStatementId);
        AddParameter(command, "$client_id", row.ClientId);
        AddParameter(command, "$class_code", Clean(row.ClassCode));
        AddParameter(command, "$employee_id", Clean(row.EmployeeId));
        AddParameter(command, "$certificate_id", Clean(row.CertificateId));
        AddParameter(command, "$employee_name", Clean(row.EmployeeName));
        AddParameter(command, "$annual_earnings", row.AnnualEarnings);
        AddParameter(command, "$cost_division", Clean(row.CostDivision));
        AddParameter(command, "$employment_class", Clean(row.EmploymentClass));
        AddParameter(command, "$retro_month", Clean(row.RetroMonth));
        AddParameter(command, "$life_volume", row.LifeVolume);
        AddParameter(command, "$life_premium", row.LifePremium);
        AddParameter(command, "$adnd_volume", row.AdAndDVolume);
        AddParameter(command, "$adnd_premium", row.AdAndDPremium);
        AddParameter(command, "$ltd_volume", row.LtdVolume);
        AddParameter(command, "$ltd_premium", row.LtdPremium);
        AddParameter(command, "$health_tier", Clean(row.HealthTier));
        AddParameter(command, "$health_premium", row.HealthPremium);
        AddParameter(command, "$dental_tier", Clean(row.DentalTier));
        AddParameter(command, "$dental_premium", row.DentalPremium);
        AddParameter(command, "$efap_premium", row.EfapPremium);
        AddParameter(command, "$virtual_healthcare_premium", row.VirtualHealthcarePremium);
        AddParameter(command, "$flex_admin_fee", row.FlexAdminFee);
        AddParameter(command, "$member_adnd_volume", row.MemberAdAndDVolume);
        AddParameter(command, "$member_adnd_premium", row.MemberAdAndDPremium);
        AddParameter(command, "$member_adnd_sp_volume", row.MemberAdAndDSpVolume);
        AddParameter(command, "$member_adnd_sp_premium", row.MemberAdAndDSpPremium);
        AddParameter(command, "$member_adnd_ch_volume", row.MemberAdAndDChVolume);
        AddParameter(command, "$member_adnd_ch_premium", row.MemberAdAndDChPremium);
        AddParameter(command, "$employee_premium", row.EmployeePremium);
        AddParameter(command, "$employer_premium", row.EmployerPremium);
        AddParameter(command, "$total_premium", row.TotalPremium);
        AddParameter(command, "$gst_hst", row.GstHst);
        AddParameter(command, "$retro_premium", row.RetroPremium);
        AddParameter(command, "$retro_gst_hst", row.RetroGstHst);
        AddParameter(command, "$row_total", row.RowTotal);
        AddParameter(command, "$source_page", row.SourcePage);
        AddParameter(command, "$source_line", row.SourceLine);
        AddParameter(command, "$validation_status", Clean(row.ValidationStatus));
        AddParameter(command, "$validation_flags", Clean(row.ValidationFlags));
        AddParameter(command, "$source_text", Clean(row.SourceText));
        AddParameter(command, "$date_loaded", ToDateTimeText(DateTime.Now));
        await command.ExecuteNonQueryAsync();

        using var idCommand = CreateCommand(connection, transaction, "SELECT last_insert_rowid();");
        return ToLong(await idCommand.ExecuteScalarAsync());
    }

    private static async Task InsertGroupHealthBillingTotalAsync(SqliteConnection connection, SqliteTransaction transaction, GroupHealthBillingStatementTotal row)
    {
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO BillingStatementTotals (
    billing_statement_id, client_id, total_level, class_code, benefit, total_type,
    lives, volume, premium, gst_hst, total_premium, source_page, date_loaded
)
VALUES (
    $billing_statement_id, $client_id, $total_level, $class_code, $benefit, $total_type,
    $lives, $volume, $premium, $gst_hst, $total_premium, $source_page, $date_loaded
);");
        AddParameter(command, "$billing_statement_id", row.BillingStatementId);
        AddParameter(command, "$client_id", row.ClientId);
        AddParameter(command, "$total_level", Clean(row.TotalLevel));
        AddParameter(command, "$class_code", Clean(row.ClassCode));
        AddParameter(command, "$benefit", Clean(row.Benefit));
        AddParameter(command, "$total_type", Clean(row.TotalType));
        AddParameter(command, "$lives", row.Lives);
        AddParameter(command, "$volume", row.Volume);
        AddParameter(command, "$premium", row.Premium);
        AddParameter(command, "$gst_hst", row.GstHst);
        AddParameter(command, "$total_premium", row.TotalPremium);
        AddParameter(command, "$source_page", row.SourcePage);
        AddParameter(command, "$date_loaded", ToDateTimeText(DateTime.Now));
        await command.ExecuteNonQueryAsync();
    }

    private static async Task<long> InsertGroupHealthRateTableAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        GroupHealthBillingExtractionResult result,
        long statementId,
        long clientId)
    {
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO ClientRateTables (
    client_id, billing_statement_id, carrier, effective_month, effective_year, source, status, created_at, updated_at
)
VALUES (
    $client_id, $billing_statement_id, $carrier, $effective_month, $effective_year, $source, $status, $created_at, $updated_at
);");
        var now = DateTime.Now;
        AddParameter(command, "$client_id", clientId);
        AddParameter(command, "$billing_statement_id", statementId);
        AddParameter(command, "$carrier", Clean(result.Carrier));
        AddParameter(command, "$effective_month", result.StatementMonth);
        AddParameter(command, "$effective_year", result.StatementYear);
        AddParameter(command, "$source", "GroupHEALTH billing import");
        AddParameter(command, "$status", Clean(result.ImportStatus));
        AddParameter(command, "$created_at", ToDateTimeText(now));
        AddParameter(command, "$updated_at", ToDateTimeText(now));
        await command.ExecuteNonQueryAsync();

        using var idCommand = CreateCommand(connection, transaction, "SELECT last_insert_rowid();");
        return ToLong(await idCommand.ExecuteScalarAsync());
    }

    private static async Task InsertGroupHealthRateLineAsync(SqliteConnection connection, SqliteTransaction transaction, GroupHealthClientRateTableLine row)
    {
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO ClientRateTableLines (
    client_rate_table_id, billing_statement_id, client_id, carrier, effective_month, effective_year,
    benefit, class_code, tier, volume_basis, employee_rate, employer_rate, total_rate,
    calculated_rate, expected_rate, rate_status, is_calculated, source, created_at
)
VALUES (
    $client_rate_table_id, $billing_statement_id, $client_id, $carrier, $effective_month, $effective_year,
    $benefit, $class_code, $tier, $volume_basis, $employee_rate, $employer_rate, $total_rate,
    $calculated_rate, $expected_rate, $rate_status, $is_calculated, $source, $created_at
);");
        AddParameter(command, "$client_rate_table_id", row.ClientRateTableId);
        AddParameter(command, "$billing_statement_id", row.BillingStatementId);
        AddParameter(command, "$client_id", row.ClientId);
        AddParameter(command, "$carrier", Clean(row.Carrier));
        AddParameter(command, "$effective_month", row.EffectiveMonth);
        AddParameter(command, "$effective_year", row.EffectiveYear);
        AddParameter(command, "$benefit", Clean(row.Benefit));
        AddParameter(command, "$class_code", Clean(row.ClassCode));
        AddParameter(command, "$tier", Clean(row.Tier));
        AddParameter(command, "$volume_basis", Clean(row.VolumeBasis));
        AddParameter(command, "$employee_rate", row.EmployeeRate);
        AddParameter(command, "$employer_rate", row.EmployerRate);
        AddParameter(command, "$total_rate", row.TotalRate);
        AddParameter(command, "$calculated_rate", row.CalculatedRate);
        AddParameter(command, "$expected_rate", row.ExpectedRate);
        AddParameter(command, "$rate_status", Clean(row.RateStatus));
        AddParameter(command, "$is_calculated", row.IsCalculated ? 1 : 0);
        AddParameter(command, "$source", Clean(row.Source));
        AddParameter(command, "$created_at", ToDateTimeText(DateTime.Now));
        await command.ExecuteNonQueryAsync();
    }

    private static async Task InsertGroupHealthValidationResultAsync(SqliteConnection connection, SqliteTransaction transaction, GroupHealthBillingValidationResult row)
    {
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO BillingValidationResults (
    billing_statement_id, billing_statement_row_id, client_id, validation_level, check_name, status, message,
    expected_amount, actual_amount, difference_amount, benefit, class_code, employee_id, certificate_id, source_page, created_at
)
VALUES (
    $billing_statement_id, $billing_statement_row_id, $client_id, $validation_level, $check_name, $status, $message,
    $expected_amount, $actual_amount, $difference_amount, $benefit, $class_code, $employee_id, $certificate_id, $source_page, $created_at
);");
        AddParameter(command, "$billing_statement_id", row.BillingStatementId);
        AddParameter(command, "$billing_statement_row_id", row.BillingStatementRowId > 0 ? row.BillingStatementRowId : null);
        AddParameter(command, "$client_id", row.ClientId);
        AddParameter(command, "$validation_level", Clean(row.ValidationLevel));
        AddParameter(command, "$check_name", Clean(row.CheckName));
        AddParameter(command, "$status", Clean(row.Status));
        AddParameter(command, "$message", Clean(row.Message));
        AddParameter(command, "$expected_amount", row.ExpectedAmount);
        AddParameter(command, "$actual_amount", row.ActualAmount);
        AddParameter(command, "$difference_amount", row.DifferenceAmount);
        AddParameter(command, "$benefit", Clean(row.Benefit));
        AddParameter(command, "$class_code", Clean(row.ClassCode));
        AddParameter(command, "$employee_id", Clean(row.EmployeeId));
        AddParameter(command, "$certificate_id", Clean(row.CertificateId));
        AddParameter(command, "$source_page", row.SourcePage);
        AddParameter(command, "$created_at", ToDateTimeText(row.CreatedAt == default ? DateTime.Now : row.CreatedAt));
        await command.ExecuteNonQueryAsync();
    }

    private static async Task InsertGroupHealthEmployeeChangeAsync(SqliteConnection connection, SqliteTransaction transaction, GroupHealthBillingEmployeeChange row)
    {
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO BillingEmployeeChangeLog (
    billing_statement_id, client_id, carrier, statement_month, statement_year, employee_id, certificate_id, employee_name,
    change_type, field_name, previous_value, current_value, difference_amount, change_percent, severity, source, created_at
)
VALUES (
    $billing_statement_id, $client_id, $carrier, $statement_month, $statement_year, $employee_id, $certificate_id, $employee_name,
    $change_type, $field_name, $previous_value, $current_value, $difference_amount, $change_percent, $severity, $source, $created_at
);");
        AddParameter(command, "$billing_statement_id", row.BillingStatementId);
        AddParameter(command, "$client_id", row.ClientId);
        AddParameter(command, "$carrier", Clean(row.Carrier));
        AddParameter(command, "$statement_month", row.StatementMonth);
        AddParameter(command, "$statement_year", row.StatementYear);
        AddParameter(command, "$employee_id", Clean(row.EmployeeId));
        AddParameter(command, "$certificate_id", Clean(row.CertificateId));
        AddParameter(command, "$employee_name", Clean(row.EmployeeName));
        AddParameter(command, "$change_type", Clean(row.ChangeType));
        AddParameter(command, "$field_name", Clean(row.FieldName));
        AddParameter(command, "$previous_value", Clean(row.PreviousValue));
        AddParameter(command, "$current_value", Clean(row.CurrentValue));
        AddParameter(command, "$difference_amount", row.DifferenceAmount);
        AddParameter(command, "$change_percent", row.ChangePercent);
        AddParameter(command, "$severity", Clean(row.Severity));
        AddParameter(command, "$source", Clean(row.Source));
        AddParameter(command, "$created_at", ToDateTimeText(row.CreatedAt == default ? DateTime.Now : row.CreatedAt));
        await command.ExecuteNonQueryAsync();
    }

    private static async Task<int> InsertClaimsDistributionReportsAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        AnalysisResult analysis,
        long clientId,
        long? documentId)
    {
        var distributionRows = ExpandClaimsDistributionPeriods(analysis.Workbook.ClaimsDistributionRows);
        if (distributionRows.Count == 0)
        {
            return 0;
        }

        var insertedRows = 0;
        foreach (var group in distributionRows.GroupBy(row => new
                 {
                     row.SourceFile,
                     row.ReportPeriod,
                     row.Benefit,
                     row.Division,
                     row.Class
                 }))
        {
            var rows = group.ToList();
            var first = rows.First();
            var validation = ClaimsDistributionExtractor.BuildValidation(rows);
            var reportId = await InsertClaimsDistributionReportAsync(
                connection,
                transaction,
                clientId,
                documentId,
                first,
                validation);
            insertedRows++;

            foreach (var row in rows)
            {
                await InsertClaimsDistributionRowAsync(connection, transaction, clientId, documentId, reportId, row);
                insertedRows++;
            }
        }

        return insertedRows;
    }

    private static async Task<int> InsertRenewalClaimsExperienceAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        AnalysisResult analysis,
        long clientId,
        long? documentId)
    {
        if (!documentId.HasValue)
        {
            return 0;
        }

        var distributionTotals = ExpandClaimsDistributionPeriods(analysis.Workbook.ClaimsDistributionRows)
            .Where(row => row.ServiceCategory.Equals("Total", StringComparison.OrdinalIgnoreCase))
            .Select(row => new
            {
                Row = row,
                Benefit = NormalizeClaimsExperienceBenefit(row.Benefit),
                Start = ParsingHelpers.ParseDateOnly(row.PeriodStart),
                End = ParsingHelpers.ParseDateOnly(row.PeriodEnd)
            })
            .Where(item => !string.IsNullOrWhiteSpace(item.Benefit) && item.End.HasValue)
            .GroupBy(item => new { item.Benefit, item.Start, item.End })
            .Select(group => group
                .OrderByDescending(item => Path.GetExtension(item.Row.SourceFile).Equals(".xls", StringComparison.OrdinalIgnoreCase))
                .ThenByDescending(item => item.Row.AmountPaid ?? 0m)
                .First())
            .ToList();

        // A renewal-experience report is authoritative for complete periods. Prefer it over
        // companion service-category reports for premiums and incurred claims. When a dated
        // service report covers the same complete period, its total is the paid-claims source.
        var renewalExperienceRows = analysis.Pdf.SunLifeExperienceRows
            .Select(row => new { Row = row, Dates = ParseExperiencePeriod(row.Period) })
            .Where(item => item.Dates.Start.HasValue && item.Dates.End.HasValue)
            .Where(item => !string.IsNullOrWhiteSpace(NormalizeClaimsExperienceBenefit(item.Row.Benefit)))
            .GroupBy(item => new
            {
                Benefit = NormalizeClaimsExperienceBenefit(item.Row.Benefit),
                Start = item.Dates.Start!.Value,
                End = item.Dates.End!.Value
            })
            .Select(group => group.First())
            .ToList();

        if (renewalExperienceRows.Count > 0)
        {
            foreach (var experience in renewalExperienceRows)
            {
                var premium = experience.Row.BilledPremium;
                var benefit = NormalizeClaimsExperienceBenefit(experience.Row.Benefit);
                var serviceTotal = distributionTotals.FirstOrDefault(total =>
                    total.Benefit.Equals(benefit, StringComparison.OrdinalIgnoreCase) &&
                    total.Start == experience.Dates.Start &&
                    total.End == experience.Dates.End);
                var paidClaims = serviceTotal?.Row.AmountPaid ?? experience.Row.PaidClaims;
                await InsertClaimsExperienceAsync(
                    connection,
                    transaction,
                    new ClaimsExperienceRecord
                    {
                        ClientId = clientId,
                        ExperiencePeriodStart = experience.Dates.Start,
                        ExperiencePeriodEnd = experience.Dates.End,
                        BenefitLine = benefit,
                        Premium = premium,
                        PaidClaims = paidClaims,
                        IncurredClaims = experience.Row.IncurredClaims,
                        LossRatio = premium > 0m && paidClaims.HasValue ? paidClaims.Value / premium.Value * 100m : null,
                        Lives = experience.Row.AverageCoveredEmployees,
                        SourceDocumentId = documentId,
                        DateLoaded = DateTime.Now
                    });
            }

            return renewalExperienceRows.Count;
        }

        var inserted = 0;
        foreach (var total in distributionTotals)
        {
            var experience = analysis.Pdf.ExperienceLines
                .Where(row => NormalizeClaimsExperienceBenefit(row.Benefit).Equals(total.Benefit, StringComparison.OrdinalIgnoreCase))
                .Select(row => new { Row = row, Dates = ParseExperiencePeriod(row.Period) })
                .OrderByDescending(item => item.Dates.End == total.End)
                .ThenByDescending(item => item.Dates.End ?? DateOnly.MinValue)
                .FirstOrDefault();
            var premium = experience?.Row.BilledPremium;
            var paidClaims = total.Row.AmountPaid;

            await InsertClaimsExperienceAsync(
                connection,
                transaction,
                new ClaimsExperienceRecord
                {
                    ClientId = clientId,
                    ExperiencePeriodStart = total.Start,
                    ExperiencePeriodEnd = total.End,
                    BenefitLine = total.Benefit,
                    Premium = premium,
                    PaidClaims = paidClaims,
                    IncurredClaims = experience?.Row.IncurredClaims,
                    LossRatio = premium > 0m && paidClaims.HasValue ? paidClaims.Value / premium.Value * 100m : null,
                    Lives = experience?.Row.AverageLives,
                    SourceDocumentId = documentId,
                    DateLoaded = DateTime.Now
                });
            inserted++;
        }

        return inserted;
    }

    private static List<ClaimsDistributionRow> ExpandClaimsDistributionPeriods(
        IEnumerable<ClaimsDistributionRow> sourceRows)
    {
        var expanded = new List<ClaimsDistributionRow>();
        foreach (var row in sourceRows)
        {
            expanded.Add(row);
            if (string.IsNullOrWhiteSpace(row.PriorReportPeriod) ||
                (!row.PriorAmountPaid.HasValue && !row.PriorPercentPaid.HasValue))
            {
                continue;
            }

            expanded.Add(new ClaimsDistributionRow
            {
                CarrierName = row.CarrierName,
                PolicyNumber = row.PolicyNumber,
                ClientName = row.ClientName,
                ReportPeriod = row.PriorReportPeriod,
                PeriodStart = row.PriorPeriodStart,
                PeriodEnd = row.PriorPeriodEnd,
                Benefit = row.Benefit,
                Division = row.Division,
                Class = row.Class,
                ReportCreatedDate = row.ReportCreatedDate,
                SourceFile = row.SourceFile,
                ServiceCategory = row.ServiceCategory,
                ServiceDescriptionOriginal = row.ServiceDescriptionOriginal,
                AmountPaid = row.PriorAmountPaid,
                PercentPaid = row.PriorPercentPaid,
                NormPercentPaid = row.NormPercentPaid,
                Occurrences = row.PriorOccurrences,
                SourceReference = string.IsNullOrWhiteSpace(row.SourceReference)
                    ? "Prior period"
                    : $"{row.SourceReference} (prior period)"
            });
        }

        return expanded;
    }

    private static string NormalizeClaimsExperienceBenefit(string benefit)
    {
        var normalized = NormalizeForMatch(benefit);
        if (normalized.Contains("dental", StringComparison.OrdinalIgnoreCase))
        {
            return "Dental";
        }

        return normalized.Contains("extendedhealth", StringComparison.OrdinalIgnoreCase) ||
               normalized.Equals("medical", StringComparison.OrdinalIgnoreCase) ||
               normalized.Equals("health", StringComparison.OrdinalIgnoreCase) ||
               normalized.Equals("ehc", StringComparison.OrdinalIgnoreCase)
            ? "Extended Health Care"
            : "";
    }

    private static (DateOnly? Start, DateOnly? End) ParseExperiencePeriod(string period)
    {
        var dates = Regex.Matches(
                period ?? "",
                @"\d{1,2}-[A-Za-z]{3,9}-\d{4}|[A-Za-z]{3,9}\s+\d{1,2},\s*\d{4}|\d{1,2}/\d{1,2}/\d{4}")
            .Select(match => ParsingHelpers.ParseDateOnly(match.Value))
            .Where(date => date.HasValue)
            .Select(date => date!.Value)
            .ToList();
        return dates.Count >= 2 ? (dates[0], dates[1]) : (null, null);
    }

    private static async Task<long> InsertClaimsDistributionReportAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long clientId,
        long? documentId,
        ClaimsDistributionRow row,
        ClaimsDistributionValidation validation)
    {
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO claims_distribution_reports (
    client_id,
    source_document_id,
    carrier_name,
    policy_number,
    client_name,
    report_period,
    period_start,
    period_end,
    benefit_type,
    division,
    class,
    report_created_date,
    source_file,
    table_found,
    paramedical_found,
    total_row_found,
    total_amount_paid,
    sum_of_service_amounts,
    amounts_balance_to_total,
    warnings,
    date_loaded
)
VALUES (
    $client_id,
    $source_document_id,
    $carrier_name,
    $policy_number,
    $client_name,
    $report_period,
    $period_start,
    $period_end,
    $benefit_type,
    $division,
    $class,
    $report_created_date,
    $source_file,
    $table_found,
    $paramedical_found,
    $total_row_found,
    $total_amount_paid,
    $sum_of_service_amounts,
    $amounts_balance_to_total,
    $warnings,
    $date_loaded
);
SELECT last_insert_rowid();");
        AddParameter(command, "$client_id", clientId);
        AddParameter(command, "$source_document_id", documentId);
        AddParameter(command, "$carrier_name", row.CarrierName);
        AddParameter(command, "$policy_number", row.PolicyNumber);
        AddParameter(command, "$client_name", row.ClientName);
        AddParameter(command, "$report_period", row.ReportPeriod);
        AddParameter(command, "$period_start", row.PeriodStart);
        AddParameter(command, "$period_end", row.PeriodEnd);
        AddParameter(command, "$benefit_type", row.Benefit);
        AddParameter(command, "$division", row.Division);
        AddParameter(command, "$class", row.Class);
        AddParameter(command, "$report_created_date", row.ReportCreatedDate);
        AddParameter(command, "$source_file", row.SourceFile);
        AddParameter(command, "$table_found", validation.TableFound ? 1 : 0);
        AddParameter(command, "$paramedical_found", validation.ParamedicalFound ? 1 : 0);
        AddParameter(command, "$total_row_found", validation.TotalRowFound ? 1 : 0);
        AddParameter(command, "$total_amount_paid", validation.TotalAmountPaid);
        AddParameter(command, "$sum_of_service_amounts", validation.SumOfServiceAmounts);
        AddParameter(command, "$amounts_balance_to_total", validation.AmountsBalanceToTotal ? 1 : 0);
        AddParameter(command, "$warnings", string.Join("; ", validation.Warnings));
        AddParameter(command, "$date_loaded", ToDateTimeText(DateTime.Now));
        var reportId = await command.ExecuteScalarAsync();
        return Convert.ToInt64(reportId, CultureInfo.InvariantCulture);
    }

    private static async Task InsertClaimsDistributionRowAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long clientId,
        long? documentId,
        long reportId,
        ClaimsDistributionRow row)
    {
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO claims_distribution_rows (
    claims_distribution_report_id,
    client_id,
    source_document_id,
    benefit,
    service_category,
    service_description_original,
    amount_paid,
    percent_paid,
    norm_percent_paid,
    occurrences,
    percent_occurrences,
    norm_percent_occurrences,
    avg_claim,
    source_reference,
    carrier_name,
    policy_number,
    report_period,
    division,
    class,
    date_loaded
)
VALUES (
    $claims_distribution_report_id,
    $client_id,
    $source_document_id,
    $benefit,
    $service_category,
    $service_description_original,
    $amount_paid,
    $percent_paid,
    $norm_percent_paid,
    $occurrences,
    $percent_occurrences,
    $norm_percent_occurrences,
    $avg_claim,
    $source_reference,
    $carrier_name,
    $policy_number,
    $report_period,
    $division,
    $class,
    $date_loaded
);");
        AddParameter(command, "$claims_distribution_report_id", reportId);
        AddParameter(command, "$client_id", clientId);
        AddParameter(command, "$source_document_id", documentId);
        AddParameter(command, "$benefit", row.Benefit);
        AddParameter(command, "$service_category", row.ServiceCategory);
        AddParameter(command, "$service_description_original", row.ServiceDescriptionOriginal);
        AddParameter(command, "$amount_paid", row.AmountPaid);
        AddParameter(command, "$percent_paid", row.PercentPaid);
        AddParameter(command, "$norm_percent_paid", row.NormPercentPaid);
        AddParameter(command, "$occurrences", row.Occurrences);
        AddParameter(command, "$percent_occurrences", row.PercentOccurrences);
        AddParameter(command, "$norm_percent_occurrences", row.NormPercentOccurrences);
        AddParameter(command, "$avg_claim", row.AverageClaim);
        AddParameter(command, "$source_reference", row.SourceReference);
        AddParameter(command, "$carrier_name", row.CarrierName);
        AddParameter(command, "$policy_number", row.PolicyNumber);
        AddParameter(command, "$report_period", row.ReportPeriod);
        AddParameter(command, "$division", row.Division);
        AddParameter(command, "$class", row.Class);
        AddParameter(command, "$date_loaded", ToDateTimeText(DateTime.Now));
        await command.ExecuteNonQueryAsync();
    }

    private static async Task DeleteRenewalsForDocumentAsync(SqliteConnection connection, SqliteTransaction transaction, long documentId)
    {
        using var command = CreateCommand(connection, transaction, "DELETE FROM Renewals WHERE source_document_id = $document_id;");
        AddParameter(command, "$document_id", documentId);
        await command.ExecuteNonQueryAsync();
    }

    private static async Task DeleteRatesForDocumentAsync(SqliteConnection connection, SqliteTransaction transaction, long documentId)
    {
        using var command = CreateCommand(connection, transaction, "DELETE FROM Rates WHERE source_document_id = $document_id;");
        AddParameter(command, "$document_id", documentId);
        await command.ExecuteNonQueryAsync();
    }

    private static async Task DeleteRenewalVersionsForDocumentAsync(SqliteConnection connection, SqliteTransaction transaction, long documentId)
    {
        await ExecuteNonQueryAsync(
            connection,
            transaction,
            "DELETE FROM renewal_calculations WHERE version_id IN (SELECT version_id FROM renewal_versions WHERE source_document_id = $document_id);",
            command => AddParameter(command, "$document_id", documentId));
        await ExecuteNonQueryAsync(
            connection,
            transaction,
            "DELETE FROM renewal_rates WHERE version_id IN (SELECT version_id FROM renewal_versions WHERE source_document_id = $document_id);",
            command => AddParameter(command, "$document_id", documentId));
        await ExecuteNonQueryAsync(
            connection,
            transaction,
            "DELETE FROM underwriting_assumptions WHERE version_id IN (SELECT version_id FROM renewal_versions WHERE source_document_id = $document_id);",
            command => AddParameter(command, "$document_id", documentId));
        await ExecuteNonQueryAsync(
            connection,
            transaction,
            "DELETE FROM renewal_versions WHERE source_document_id = $document_id;",
            command => AddParameter(command, "$document_id", documentId));
    }

    private static async Task DeleteClaimsDistributionForDocumentAsync(SqliteConnection connection, SqliteTransaction transaction, long documentId)
    {
        await ExecuteNonQueryAsync(
            connection,
            transaction,
            "DELETE FROM claims_distribution_rows WHERE source_document_id = $document_id;",
            command => AddParameter(command, "$document_id", documentId));
        await ExecuteNonQueryAsync(
            connection,
            transaction,
            "DELETE FROM claims_distribution_reports WHERE source_document_id = $document_id;",
            command => AddParameter(command, "$document_id", documentId));
        await ExecuteNonQueryAsync(
            connection,
            transaction,
            "DELETE FROM claims_distribution WHERE source_document_id = $document_id;",
            command => AddParameter(command, "$document_id", documentId));
    }

    private static async Task DeleteClaimsForDocumentAsync(SqliteConnection connection, SqliteTransaction transaction, long documentId)
    {
        using var command = CreateCommand(connection, transaction, "DELETE FROM ClaimsExperience WHERE source_document_id = $document_id;");
        AddParameter(command, "$document_id", documentId);
        await command.ExecuteNonQueryAsync();
    }

    private static async Task DeleteMonthlyClaimsForDocumentAsync(SqliteConnection connection, SqliteTransaction transaction, long documentId)
    {
        using var command = CreateCommand(connection, transaction, "DELETE FROM MonthlyClaims WHERE source_document_id = $document_id;");
        AddParameter(command, "$document_id", documentId);
        await command.ExecuteNonQueryAsync();
    }

    private static async Task DeletePlanDesignForDocumentAsync(SqliteConnection connection, SqliteTransaction transaction, long documentId)
    {
        using var command = CreateCommand(connection, transaction, "DELETE FROM PlanDesign WHERE source_document_id = $document_id;");
        AddParameter(command, "$document_id", documentId);
        await command.ExecuteNonQueryAsync();
    }

    private static async Task DeletePlanDesignFeaturesForDocumentAsync(SqliteConnection connection, SqliteTransaction transaction, long documentId)
    {
        using (var ruleCommand = CreateCommand(connection, transaction, "DELETE FROM PlanBenefitRule WHERE SourceDocumentId = $document_id;"))
        {
            AddParameter(ruleCommand, "$document_id", documentId);
            await ruleCommand.ExecuteNonQueryAsync();
        }

        using var command = CreateCommand(connection, transaction, "DELETE FROM plan_design_features WHERE source_document_id = $document_id;");
        AddParameter(command, "$document_id", documentId);
        await command.ExecuteNonQueryAsync();
    }

    private static async Task UpdateDocumentProcessingStatusAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long documentId,
        string status)
    {
        await ExecuteNonQueryAsync(
            connection,
            transaction,
            @"
UPDATE Documents
SET processing_status = $status,
    processed_date = $processed_date
WHERE document_id = $document_id;",
            command =>
            {
                AddParameter(command, "$status", Clean(status));
                AddParameter(command, "$processed_date", ToDateTimeText(DateTime.Now));
                AddParameter(command, "$document_id", documentId);
            });
    }

    private static async Task<int> InsertDemographicAgeGenderRowsAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long clientId,
        long? documentId,
        DemographicLine row)
    {
        var insertedRows = 0;
        insertedRows += await InsertDemographicAgeGenderRowAsync(connection, transaction, clientId, documentId, row, "Prior", "Male", row.PriorMaleLives, row.PriorMaleVolume);
        insertedRows += await InsertDemographicAgeGenderRowAsync(connection, transaction, clientId, documentId, row, "Prior", "Female", row.PriorFemaleLives, row.PriorFemaleVolume);
        insertedRows += await InsertDemographicAgeGenderRowAsync(connection, transaction, clientId, documentId, row, "Current", "Male", row.CurrentMaleLives, row.CurrentMaleVolume);
        insertedRows += await InsertDemographicAgeGenderRowAsync(connection, transaction, clientId, documentId, row, "Current", "Female", row.CurrentFemaleLives, row.CurrentFemaleVolume);
        return insertedRows;
    }

    private static async Task<int> InsertDemographicAgeGenderRowAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long clientId,
        long? documentId,
        DemographicLine row,
        string periodLabel,
        string gender,
        int? lives,
        decimal? volume)
    {
        if (!lives.HasValue && !volume.HasValue)
        {
            return 0;
        }

        using var command = CreateCommand(connection, transaction, @"
INSERT INTO demographics_by_age_gender (
    client_id,
    source_document_id,
    benefit,
    period_label,
    age_band,
    gender,
    lives,
    volume,
    source_reference,
    date_loaded
)
VALUES (
    $client_id,
    $source_document_id,
    $benefit,
    $period_label,
    $age_band,
    $gender,
    $lives,
    $volume,
    $source_reference,
    $date_loaded
);");
        AddParameter(command, "$client_id", clientId);
        AddParameter(command, "$source_document_id", documentId);
        AddParameter(command, "$benefit", row.Benefit);
        AddParameter(command, "$period_label", periodLabel);
        AddParameter(command, "$age_band", row.AgeBand);
        AddParameter(command, "$gender", gender);
        AddParameter(command, "$lives", lives);
        AddParameter(command, "$volume", volume);
        AddParameter(command, "$source_reference", row.SourceReference);
        AddParameter(command, "$date_loaded", ToDateTimeText(DateTime.Now));
        await command.ExecuteNonQueryAsync();
        return 1;
    }

    private async Task<List<RenewalRecord>> QueryRenewalsAsync(string sql, Action<SqliteCommand> addParameters)
    {
        await InitializeCoreAsync();
        using var connection = await OpenConnectionAsync();
        using var command = CreateCommand(connection, null, sql);
        addParameters(command);

        var records = new List<RenewalRecord>();
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            records.Add(ReadRenewal(reader));
        }

        return records;
    }

    private static List<RenewalRecord> BuildRenewalRows(AnalysisResult analysis, long clientId, long? sourceDocumentId)
    {
        var renewalDate = analysis.RenewalCase.RenewalDate ?? analysis.Pdf.RenewalDate;
        var renewalYear = renewalDate?.Year ?? DateTime.Today.Year;
        var carrier = FirstNonBlank(analysis.Pdf.Carrier, "Manulife");
        var rows = new List<RenewalRecord>();

        foreach (var rateLine in analysis.Pdf.RateLines)
        {
            var experience = FindMatchingExperience(analysis.Pdf.ExperienceLines, rateLine.Benefit);
            var claims = experience?.PaidClaims ?? experience?.IncurredClaims;
            var lossRatio = experience?.PaidLossRatio ?? experience?.IncurredLossRatio ?? ComputeLossRatio(claims, rateLine.CurrentPremium ?? rateLine.RenewalPremium);
            rows.Add(new RenewalRecord
            {
                ClientId = clientId,
                RenewalYear = renewalYear,
                Carrier = carrier,
                BenefitLine = FirstNonBlank(rateLine.Benefit, "Overall"),
                CurrentRate = rateLine.CurrentRate,
                RenewalRate = rateLine.RenewalRate,
                RateChangePercent = rateLine.PercentChange,
                Premium = rateLine.RenewalPremium ?? rateLine.CurrentPremium,
                Claims = claims,
                LossRatio = lossRatio,
                Lives = rateLine.Lives,
                RenewalStatus = "Processed",
                SourceDocumentId = sourceDocumentId,
                DateLoaded = DateTime.Now
            });
        }

        if (rows.Count == 0)
        {
            var claims = analysis.Pdf.ExperienceLines.Sum(line => line.PaidClaims ?? line.IncurredClaims ?? 0m);
            var premium = analysis.RenewalMonthlyPremium ?? analysis.CurrentMonthlyPremium;
            rows.Add(new RenewalRecord
            {
                ClientId = clientId,
                RenewalYear = renewalYear,
                Carrier = carrier,
                BenefitLine = "Overall",
                RateChangePercent = analysis.MonthlyPremiumChangePercent,
                Premium = premium,
                Claims = claims == 0m ? null : claims,
                LossRatio = ComputeLossRatio(claims == 0m ? null : claims, premium),
                RenewalStatus = "Processed",
                SourceDocumentId = sourceDocumentId,
                DateLoaded = DateTime.Now
            });
        }

        return rows;
    }

    private static List<RateRecord> BuildRateRows(AnalysisResult analysis, long clientId, long? sourceDocumentId)
    {
        var effectiveDate = analysis.RenewalCase.RenewalDate ?? analysis.Pdf.RenewalDate;
        var loadedAt = DateTime.Now;
        var carrier = FirstNonBlank(analysis.Pdf.Carrier, "Manulife");
        var rows = new List<RateRecord>();

        foreach (var rateLine in analysis.Pdf.RateLines)
        {
            rows.Add(new RateRecord
            {
                ClientId = clientId,
                Carrier = carrier,
                BenefitLine = FirstNonBlank(rateLine.Benefit, "Overall"),
                ClassName = FirstNonBlank(rateLine.TierOrClass, rateLine.AccountClass),
                CurrentRate = rateLine.CurrentRate,
                ProposedRate = rateLine.RenewalRate,
                NegotiatedRate = null,
                RateBasis = InferRateBasis(rateLine),
                EffectiveDate = effectiveDate,
                SourceDocumentId = sourceDocumentId,
                DateLoaded = loadedAt
            });
        }

        return rows;
    }

    private static async Task<int> UpsertParsedRenewalExperienceTermsAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        AnalysisResult analysis,
        long clientId,
        string carrier,
        DateOnly? effectiveDate)
    {
        var targetLossRatio = FindExperienceTargetLossRatio(analysis.Pdf.ExperienceLines);
        if (!targetLossRatio.HasValue &&
            !analysis.Pdf.PoolingThreshold.HasValue &&
            !analysis.Pdf.PoolingCharge.HasValue)
        {
            return 0;
        }

        await UpsertRenewalAssumptionAsync(
            connection,
            transaction,
            new RenewalAssumptionRecord
            {
                ClientId = clientId,
                Carrier = carrier,
                BenefitLine = "All",
                TargetLossRatio = NormalizePercentDecimal(targetLossRatio),
                PoolingCharge = NormalizePercentDecimal(analysis.Pdf.PoolingCharge),
                PoolingThreshold = analysis.Pdf.PoolingThreshold,
                EffectiveDate = effectiveDate
            });

        return 1;
    }

    private static decimal? FindExperienceTargetLossRatio(IEnumerable<ExperienceLine> experienceLines)
    {
        var healthTarget = experienceLines
            .Where(row => row.Benefit.Contains("extended health", StringComparison.OrdinalIgnoreCase) ||
                          row.Benefit.Contains("health care", StringComparison.OrdinalIgnoreCase) ||
                          row.Benefit.Equals("EHC", StringComparison.OrdinalIgnoreCase))
            .Select(row => row.TargetLossRatio)
            .FirstOrDefault(value => value.HasValue);

        return healthTarget ?? experienceLines
            .Select(row => row.TargetLossRatio)
            .FirstOrDefault(value => value.HasValue);
    }

    private static string InferRateBasis(RateLine rateLine)
    {
        var normalized = NormalizeForMatch($"{rateLine.Benefit} {rateLine.TierOrClass} {rateLine.AccountClass}");
        if (normalized.Contains("single", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("employeeonly", StringComparison.OrdinalIgnoreCase))
        {
            return "single";
        }

        if (normalized.Contains("family", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("couple", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("spouse", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("dependent", StringComparison.OrdinalIgnoreCase))
        {
            return "family";
        }

        if (rateLine.Volume.HasValue ||
            normalized.Contains("volume", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("life", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("adandd", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("std", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("ltd", StringComparison.OrdinalIgnoreCase))
        {
            return "volume";
        }

        return "composite";
    }

    private static ExperienceLine? FindMatchingExperience(IEnumerable<ExperienceLine> experienceLines, string benefit)
    {
        return ExperienceLineSelector.FindBestMatch(experienceLines, benefit);
    }

    private static decimal? ComputeLossRatio(decimal? claims, decimal? premium)
    {
        return claims.HasValue && premium is > 0m
            ? decimal.Round(claims.Value / premium.Value * 100m, 1)
            : null;
    }

    private static ClientRecord BuildPlanDesignClient(IReadOnlyList<PlanDesignRecord> rows)
    {
        var row = rows.FirstOrDefault(item => !string.IsNullOrWhiteSpace(item.ClientName) || !string.IsNullOrWhiteSpace(item.ClientNumber));
        return new ClientRecord
        {
            ClientName = FirstNonBlank(row?.ClientName, UnknownClientName),
            ClientNumber = row?.ClientNumber ?? "",
            Carrier = row?.Carrier ?? "",
            Status = DefaultClientStatus
        };
    }

    internal static List<PlanDesignRecord> ReadPlanDesignRowsForBookletExtraction(
        BookletExtractionResult result,
        IEnumerable<string> inputPaths)
    {
        var sourcePdfPaths = ResolvePdfInputPaths(inputPaths);
        var workbookRows = ReadPlanDesignRows(result.MasterWorkbookPath);
        var sourceFileNames = sourcePdfPaths
            .Select(Path.GetFileName)
            .Where(fileName => !string.IsNullOrWhiteSpace(fileName))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        var sourceRows = workbookRows
            .Where(row => sourceFileNames.Contains(Path.GetFileName(row.SourceFileName)))
            .ToList();
        return sourceRows.Count > 0 ? sourceRows : workbookRows;
    }

    private static List<PlanDesignRecord> ReadPlanDesignRows(string workbookPath)
    {
        var records = new List<PlanDesignRecord>();
        if (string.IsNullOrWhiteSpace(workbookPath) || !File.Exists(workbookPath))
        {
            return records;
        }

        try
        {
            using var workbook = new XLWorkbook(workbookPath);
            var planDesignSheets = workbook.Worksheets
                .Where(sheet => sheet.Name.Equals("PlanDesign", StringComparison.OrdinalIgnoreCase))
                .ToList();
            var sheetsToRead = planDesignSheets.Count > 0
                ? planDesignSheets
                : workbook.Worksheets.Where(IsPlanDesignImportSheet).ToList();

            foreach (var sheet in sheetsToRead)
            {
                var headerRow = FindPlanDesignHeaderRow(sheet);
                if (headerRow == 0)
                {
                    continue;
                }

                var headers = BuildHeaderMap(sheet.Row(headerRow));
                var clientColumn = FindColumn(headers, "clientname", "client", "employer", "groupname");
                var clientNumberColumn = FindColumn(headers, "clientnumber", "policynumber", "groupnumber", "policy", "group");
                var carrierColumn = FindColumn(headers, "carrier", "insurer");
                var classColumn = FindColumn(headers, "classname", "class", "classcode", "classdescription", "employeegroup", "eligibilityclass");
                var categoryColumn = FindColumn(headers, "benefitcategory", "category", "benefittype");
                var nameColumn = FindColumn(headers, "benefitname", "benefit", "coverage", "service");
                var descriptionColumn = FindColumn(headers, "coveragedescription", "description", "coveragedetails", "plandescription", "planprovision", "details");
                var reimbursementColumn = FindColumn(headers, "reimbursementpercent", "reimbursement", "coinsurance", "coinsurancepercent", "percent");
                var maximumColumn = FindColumn(headers, "maximumamount", "maximum", "max", "annualmaximum");
                var frequencyColumn = FindColumn(headers, "frequency");
                var waitingPeriodColumn = FindColumn(headers, "waitingperiod", "waiting");
                var terminationAgeColumn = FindColumn(headers, "terminationage", "termination", "terminates");
                var sourceFileColumn = FindColumn(headers, "sourcefile", "sourcefilename", "filename", "pdf", "document");
                var sourceSectionColumn = FindColumn(headers, "sourcesection", "sourcelocation", "section", "sectionname", "benefitsection");
                var notesColumn = FindColumn(headers, "notes", "note", "remarks", "comment", "comments");
                var sourcePageColumn = FindColumn(headers, "sourcepage", "page", "pagenumber");

                foreach (var row in sheet.RowsUsed().Where(row => row.RowNumber() > headerRow))
                {
                    var benefitName = ReadCell(row, nameColumn);
                    var category = ReadCell(row, categoryColumn);
                    var description = ReadCell(row, descriptionColumn);
                    if (string.IsNullOrWhiteSpace(benefitName) &&
                        string.IsNullOrWhiteSpace(category) &&
                        string.IsNullOrWhiteSpace(description))
                    {
                        continue;
                    }

                    var reimbursementText = ReadCell(row, reimbursementColumn);
                    records.Add(new PlanDesignRecord
                    {
                        ClientName = ReadCell(row, clientColumn),
                        ClientNumber = ReadCell(row, clientNumberColumn),
                        Carrier = ReadCell(row, carrierColumn),
                        ClassName = ReadCell(row, classColumn),
                        BenefitCategory = category,
                        BenefitName = benefitName,
                        CoverageDescription = description,
                        ReimbursementPercent = ParsePercent(reimbursementText),
                        ReimbursementText = reimbursementText,
                        MaximumAmount = ReadCell(row, maximumColumn),
                        Frequency = ReadCell(row, frequencyColumn),
                        WaitingPeriod = ReadCell(row, waitingPeriodColumn),
                        TerminationAge = ReadCell(row, terminationAgeColumn),
                        SourceFileName = ReadCell(row, sourceFileColumn),
                        SourceSection = ReadCell(row, sourceSectionColumn),
                        Notes = ReadCell(row, notesColumn),
                        SourcePage = ParseInt(ReadCell(row, sourcePageColumn)),
                        DateLoaded = DateTime.Now
                    });
                }
            }
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
        }

        return records;
    }

    private static bool IsPlanDesignImportSheet(IXLWorksheet sheet)
    {
        var name = sheet.Name;
        return !name.StartsWith("Summary_", StringComparison.OrdinalIgnoreCase) &&
               !name.Equals("Booklet_Extraction_Review", StringComparison.OrdinalIgnoreCase) &&
               !name.Equals("Extraction_Review", StringComparison.OrdinalIgnoreCase) &&
               !name.Equals("Schedule_of_Benefits", StringComparison.OrdinalIgnoreCase) &&
               !name.Equals("Schedule_PPT_Export", StringComparison.OrdinalIgnoreCase) &&
               !name.Equals("ProcessingLog", StringComparison.OrdinalIgnoreCase);
    }

    private static int FindPlanDesignHeaderRow(IXLWorksheet sheet)
    {
        var lastRow = sheet.LastRowUsed()?.RowNumber() ?? 0;
        var scanLimit = Math.Min(lastRow, 20);
        for (var rowNumber = 1; rowNumber <= scanLimit; rowNumber++)
        {
            var headers = BuildHeaderMap(sheet.Row(rowNumber));
            var hasBenefit = FindColumn(headers, "benefitname", "benefit", "coverage", "service") > 0;
            var hasDescription = FindColumn(headers, "coveragedescription", "description", "coveragedetails", "plandescription", "planprovision", "details") > 0;
            if (hasBenefit && hasDescription)
            {
                return rowNumber;
            }
        }

        return 0;
    }

    private static Dictionary<string, int> BuildHeaderMap(IXLRow row)
    {
        var headers = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (var cell in row.CellsUsed())
        {
            var key = NormalizeHeader(cell.GetString());
            if (!string.IsNullOrWhiteSpace(key) && !headers.ContainsKey(key))
            {
                headers[key] = cell.Address.ColumnNumber;
            }
        }

        return headers;
    }

    private static int FindColumn(IReadOnlyDictionary<string, int> headers, params string[] aliases)
    {
        foreach (var alias in aliases.Select(NormalizeHeader))
        {
            if (headers.TryGetValue(alias, out var column))
            {
                return column;
            }
        }

        return 0;
    }

    private static string ReadCell(IXLRow row, int column)
    {
        return column <= 0 ? "" : row.Cell(column).GetFormattedString().Trim();
    }

    private static long? ResolvePlanDesignDocumentId(PlanDesignRecord row, IReadOnlyList<DocumentSaveResult> documents)
    {
        if (row.SourceDocumentId is not null)
        {
            return row.SourceDocumentId;
        }

        if (!string.IsNullOrWhiteSpace(row.SourceFileName))
        {
            var sourceFileName = Path.GetFileName(row.SourceFileName);
            var exactMatch = documents.FirstOrDefault(document =>
                Path.GetFileName(document.FilePath).Equals(sourceFileName, StringComparison.OrdinalIgnoreCase));
            if (exactMatch?.DocumentId is not null)
            {
                return exactMatch.DocumentId.Value;
            }

            var normalized = NormalizeForMatch(Path.GetFileNameWithoutExtension(row.SourceFileName));
            var partialMatches = documents
                .Where(document =>
                {
                    var documentName = NormalizeForMatch(Path.GetFileNameWithoutExtension(document.FilePath));
                    return documentName.Contains(normalized, StringComparison.OrdinalIgnoreCase) ||
                           normalized.Contains(documentName, StringComparison.OrdinalIgnoreCase);
                })
                .Where(document => document.DocumentId.HasValue)
                .ToList();
            if (partialMatches.Count == 1)
            {
                return partialMatches[0].DocumentId;
            }
        }

        return documents.Count == 1 ? documents[0].DocumentId : null;
    }

    private static List<string> ResolvePdfInputPaths(IEnumerable<string> inputPaths)
    {
        var paths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var inputPath in inputPaths.Where(path => !string.IsNullOrWhiteSpace(path)))
        {
            if (File.Exists(inputPath) && Path.GetExtension(inputPath).Equals(".pdf", StringComparison.OrdinalIgnoreCase))
            {
                paths.Add(Path.GetFullPath(inputPath));
                continue;
            }

            if (Directory.Exists(inputPath))
            {
                foreach (var pdfPath in Directory.EnumerateFiles(inputPath, "*.pdf", SearchOption.AllDirectories))
                {
                    paths.Add(Path.GetFullPath(pdfPath));
                }
            }
        }

        return paths.ToList();
    }

    private static List<string> ResolveExistingSourceFiles(IEnumerable<string> sourcePaths)
    {
        return sourcePaths
            .Where(path => !string.IsNullOrWhiteSpace(path))
            .Where(File.Exists)
            .Select(Path.GetFullPath)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static async Task ExecuteNonQueryAsync(
        SqliteConnection connection,
        SqliteTransaction? transaction,
        string sql,
        Action<SqliteCommand>? addParameters = null)
    {
        using var command = CreateCommand(connection, transaction, sql);
        addParameters?.Invoke(command);
        await command.ExecuteNonQueryAsync();
    }

    private static async Task<bool> TableExistsAsync(SqliteConnection connection, string tableName)
    {
        using var command = CreateCommand(
            connection,
            null,
            "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND LOWER(name) = LOWER($table_name);");
        AddParameter(command, "$table_name", tableName);
        return Convert.ToInt32(await command.ExecuteScalarAsync(), CultureInfo.InvariantCulture) > 0;
    }

    private static SqliteCommand CreateCommand(SqliteConnection connection, SqliteTransaction? transaction, string sql)
    {
        var command = connection.CreateCommand();
        command.CommandText = sql;
        if (transaction is not null)
        {
            command.Transaction = transaction;
        }

        return command;
    }

    private static void AddParameter(SqliteCommand command, string name, object? value)
    {
        command.Parameters.AddWithValue(name, value ?? DBNull.Value);
    }

    private static async Task<string> ComputeFileHashAsync(string filePath)
    {
        await using var stream = File.OpenRead(filePath);
        var hash = await SHA256.HashDataAsync(stream);
        return Convert.ToHexString(hash);
    }

    private static GroupHealthBillingHistoryRecord ReadGroupHealthBillingHistory(SqliteDataReader reader)
    {
        return new GroupHealthBillingHistoryRecord
        {
            BillingStatementId = ReadLong(reader, "billing_statement_id"),
            ClientId = ReadLong(reader, "client_id"),
            Carrier = ReadString(reader, "carrier"),
            StatementMonth = ReadInt(reader, "statement_month") ?? 0,
            StatementYear = ReadInt(reader, "statement_year") ?? 0,
            StatementDate = ReadDateOnly(reader, "statement_date"),
            ImportDate = ReadDateTime(reader, "import_date") ?? DateTime.MinValue,
            SourceFileName = ReadString(reader, "source_file_name"),
            ImportVersion = ReadInt(reader, "import_version") ?? 1,
            ImportStatus = ReadString(reader, "import_status"),
            BillingIntegrityScore = ReadDecimal(reader, "billing_integrity_score") ?? 0m,
            TotalEmployeePremium = ReadDecimal(reader, "total_employee_premium"),
            TotalEmployerPremium = ReadDecimal(reader, "total_employer_premium"),
            TotalPremium = ReadDecimal(reader, "statement_total") ?? ReadDecimal(reader, "total_premium"),
            EmployeeCount = ReadInt(reader, "employee_count") ?? 0,
            IsReplaced = ReadInt(reader, "is_replaced") == 1
        };
    }

    private static GroupHealthBillingRow ReadGroupHealthBillingRow(SqliteDataReader reader)
    {
        return new GroupHealthBillingRow
        {
            BillingStatementRowId = ReadLong(reader, "billing_statement_row_id"),
            BillingStatementId = ReadLong(reader, "billing_statement_id"),
            ClientId = ReadLong(reader, "client_id"),
            ClassCode = ReadString(reader, "class_code"),
            EmployeeId = ReadString(reader, "employee_id"),
            CertificateId = ReadString(reader, "certificate_id"),
            EmployeeName = ReadString(reader, "employee_name"),
            AnnualEarnings = ReadDecimal(reader, "annual_earnings"),
            CostDivision = ReadString(reader, "cost_division"),
            EmploymentClass = ReadString(reader, "employment_class"),
            RetroMonth = ReadString(reader, "retro_month"),
            LifeVolume = ReadDecimal(reader, "life_volume"),
            LifePremium = ReadDecimal(reader, "life_premium"),
            AdAndDVolume = ReadDecimal(reader, "adnd_volume"),
            AdAndDPremium = ReadDecimal(reader, "adnd_premium"),
            LtdVolume = ReadDecimal(reader, "ltd_volume"),
            LtdPremium = ReadDecimal(reader, "ltd_premium"),
            HealthTier = ReadString(reader, "health_tier"),
            HealthPremium = ReadDecimal(reader, "health_premium"),
            DentalTier = ReadString(reader, "dental_tier"),
            DentalPremium = ReadDecimal(reader, "dental_premium"),
            EfapPremium = ReadDecimal(reader, "efap_premium"),
            VirtualHealthcarePremium = ReadDecimal(reader, "virtual_healthcare_premium"),
            FlexAdminFee = ReadDecimal(reader, "flex_admin_fee"),
            MemberAdAndDVolume = ReadDecimal(reader, "member_adnd_volume"),
            MemberAdAndDPremium = ReadDecimal(reader, "member_adnd_premium"),
            MemberAdAndDSpVolume = ReadDecimal(reader, "member_adnd_sp_volume"),
            MemberAdAndDSpPremium = ReadDecimal(reader, "member_adnd_sp_premium"),
            MemberAdAndDChVolume = ReadDecimal(reader, "member_adnd_ch_volume"),
            MemberAdAndDChPremium = ReadDecimal(reader, "member_adnd_ch_premium"),
            EmployeePremium = ReadDecimal(reader, "employee_premium"),
            EmployerPremium = ReadDecimal(reader, "employer_premium"),
            TotalPremium = ReadDecimal(reader, "total_premium"),
            GstHst = ReadDecimal(reader, "gst_hst"),
            RetroPremium = ReadDecimal(reader, "retro_premium"),
            RetroGstHst = ReadDecimal(reader, "retro_gst_hst"),
            RowTotal = ReadDecimal(reader, "row_total"),
            SourcePage = ReadInt(reader, "source_page") ?? 0,
            SourceLine = ReadInt(reader, "source_line") ?? 0,
            ValidationStatus = ReadString(reader, "validation_status"),
            ValidationFlags = ReadString(reader, "validation_flags"),
            SourceText = ReadString(reader, "source_text")
        };
    }

    private static GroupHealthClientRateTableLine ReadGroupHealthRateLine(SqliteDataReader reader)
    {
        return new GroupHealthClientRateTableLine
        {
            ClientRateTableLineId = ReadLong(reader, "client_rate_table_line_id"),
            ClientRateTableId = ReadLong(reader, "client_rate_table_id"),
            BillingStatementId = ReadLong(reader, "billing_statement_id"),
            ClientId = ReadLong(reader, "client_id"),
            Carrier = ReadString(reader, "carrier"),
            EffectiveMonth = ReadInt(reader, "effective_month") ?? 0,
            EffectiveYear = ReadInt(reader, "effective_year") ?? 0,
            Benefit = ReadString(reader, "benefit"),
            ClassCode = ReadString(reader, "class_code"),
            Tier = ReadString(reader, "tier"),
            VolumeBasis = ReadString(reader, "volume_basis"),
            EmployeeRate = ReadDecimal(reader, "employee_rate"),
            EmployerRate = ReadDecimal(reader, "employer_rate"),
            TotalRate = ReadDecimal(reader, "total_rate"),
            CalculatedRate = ReadDecimal(reader, "calculated_rate"),
            ExpectedRate = ReadDecimal(reader, "expected_rate"),
            RateStatus = ReadString(reader, "rate_status"),
            IsCalculated = ReadInt(reader, "is_calculated") != 0,
            Source = ReadString(reader, "source")
        };
    }

    private static GroupHealthBillingValidationResult ReadGroupHealthValidationResult(SqliteDataReader reader)
    {
        return new GroupHealthBillingValidationResult
        {
            BillingValidationResultId = ReadLong(reader, "billing_validation_result_id"),
            BillingStatementId = ReadLong(reader, "billing_statement_id"),
            BillingStatementRowId = ReadNullableLong(reader, "billing_statement_row_id"),
            ClientId = ReadLong(reader, "client_id"),
            ValidationLevel = ReadString(reader, "validation_level"),
            CheckName = ReadString(reader, "check_name"),
            Status = ReadString(reader, "status"),
            Message = ReadString(reader, "message"),
            ExpectedAmount = ReadDecimal(reader, "expected_amount"),
            ActualAmount = ReadDecimal(reader, "actual_amount"),
            DifferenceAmount = ReadDecimal(reader, "difference_amount"),
            Benefit = ReadString(reader, "benefit"),
            ClassCode = ReadString(reader, "class_code"),
            EmployeeId = ReadString(reader, "employee_id"),
            CertificateId = ReadString(reader, "certificate_id"),
            SourcePage = ReadInt(reader, "source_page"),
            CreatedAt = ReadDateTime(reader, "created_at") ?? DateTime.MinValue
        };
    }

    private static GroupHealthBillingEmployeeChange ReadGroupHealthEmployeeChange(SqliteDataReader reader)
    {
        return new GroupHealthBillingEmployeeChange
        {
            BillingEmployeeChangeLogId = ReadLong(reader, "billing_employee_change_log_id"),
            BillingStatementId = ReadLong(reader, "billing_statement_id"),
            ClientId = ReadLong(reader, "client_id"),
            Carrier = ReadString(reader, "carrier"),
            StatementMonth = ReadInt(reader, "statement_month") ?? 0,
            StatementYear = ReadInt(reader, "statement_year") ?? 0,
            EmployeeId = ReadString(reader, "employee_id"),
            CertificateId = ReadString(reader, "certificate_id"),
            EmployeeName = ReadString(reader, "employee_name"),
            ChangeType = ReadString(reader, "change_type"),
            FieldName = ReadString(reader, "field_name"),
            PreviousValue = ReadString(reader, "previous_value"),
            CurrentValue = ReadString(reader, "current_value"),
            DifferenceAmount = ReadDecimal(reader, "difference_amount"),
            ChangePercent = ReadDecimal(reader, "change_percent"),
            Severity = ReadString(reader, "severity"),
            Source = ReadString(reader, "source"),
            CreatedAt = ReadDateTime(reader, "created_at") ?? DateTime.MinValue
        };
    }

    private static RenewalRecord ReadRenewal(SqliteDataReader reader)
    {
        return new RenewalRecord
        {
            RenewalId = ReadLong(reader, "renewal_id"),
            ClientId = ReadLong(reader, "client_id"),
            RenewalYear = ReadInt(reader, "renewal_year") ?? 0,
            Carrier = ReadString(reader, "carrier"),
            BenefitLine = ReadString(reader, "benefit_line"),
            CurrentRate = ReadDecimal(reader, "current_rate"),
            RenewalRate = ReadDecimal(reader, "renewal_rate"),
            RateChangePercent = ReadDecimal(reader, "rate_change_percent"),
            Premium = ReadDecimal(reader, "premium"),
            Claims = ReadDecimal(reader, "claims"),
            LossRatio = ReadDecimal(reader, "loss_ratio"),
            Lives = ReadInt(reader, "lives"),
            RenewalStatus = ReadString(reader, "renewal_status"),
            SourceDocumentId = ReadNullableLong(reader, "source_document_id"),
            DateLoaded = ReadDateTime(reader, "date_loaded") ?? DateTime.MinValue
        };
    }

    private static RateRecord ReadRate(SqliteDataReader reader)
    {
        return new RateRecord
        {
            RateId = ReadLong(reader, "rate_id"),
            ClientId = ReadLong(reader, "client_id"),
            Carrier = ReadString(reader, "carrier"),
            BenefitLine = ReadString(reader, "benefit_line"),
            ClassName = ReadString(reader, "class"),
            CurrentRate = ReadDecimal(reader, "current_rate"),
            ProposedRate = ReadDecimal(reader, "proposed_rate"),
            NegotiatedRate = ReadDecimal(reader, "negotiated_rate"),
            RateBasis = ReadString(reader, "rate_basis"),
            EffectiveDate = ReadDateOnly(reader, "effective_date"),
            SourceDocumentId = ReadNullableLong(reader, "source_document_id"),
            DateLoaded = ReadDateTime(reader, "date_loaded") ?? DateTime.MinValue
        };
    }

    private static RenewalAssumptionRecord ReadRenewalAssumption(SqliteDataReader reader)
    {
        return new RenewalAssumptionRecord
        {
            AssumptionId = ReadLong(reader, "assumption_id"),
            ClientId = ReadLong(reader, "client_id"),
            Carrier = ReadString(reader, "carrier"),
            BenefitLine = ReadString(reader, "benefit_line"),
            Trend = ReadDecimal(reader, "trend"),
            TargetLossRatio = ReadDecimal(reader, "target_loss_ratio"),
            PoolingCharge = ReadDecimal(reader, "pooling_charge"),
            PoolingThreshold = ReadDecimal(reader, "pooling_threshold"),
            Ibnr = ReadDecimal(reader, "ibnr"),
            AdminFee = ReadDecimal(reader, "admin_fee"),
            RiskCharge = ReadDecimal(reader, "risk_charge"),
            EffectiveDate = ReadDateOnly(reader, "effective_date")
        };
    }

    private static CarrierAssumptionRecord ReadCarrierAssumption(SqliteDataReader reader)
    {
        var targetLossRatio = ReadDecimal(reader, "target_loss_ratio") ?? ReadDecimal(reader, "assumption_value");
        return new CarrierAssumptionRecord
        {
            CarrierAssumptionId = ReadLong(reader, "carrier_assumption_id"),
            Carrier = ReadString(reader, "carrier"),
            BlockName = ReadString(reader, "block_name"),
            BenefitLine = ReadString(reader, "benefit_line"),
            FundingType = ReadString(reader, "funding_type"),
            TargetLossRatio = targetLossRatio,
            TlrBasis = ReadString(reader, "tlr_basis"),
            PremiumTax = ReadDecimal(reader, "premium_tax"),
            Commission = ReadDecimal(reader, "commission"),
            AdminMargin = ReadDecimal(reader, "admin_margin"),
            RiskMargin = ReadDecimal(reader, "risk_margin"),
            Trend = ReadDecimal(reader, "trend"),
            CredibilityFactor = ReadDecimal(reader, "credibility_factor"),
            PoolingThreshold = ReadDecimal(reader, "pooling_threshold"),
            PoolingCharge = ReadDecimal(reader, "pooling_charge"),
            ReserveIbnr = ReadDecimal(reader, "reserve_ibnr"),
            EffectiveDate = ReadDateOnly(reader, "effective_date"),
            Status = FirstNonBlank(ReadString(reader, "status"), "Active"),
            Notes = ReadString(reader, "notes"),
            DateLoaded = ReadDateTime(reader, "date_loaded") ?? DateTime.MinValue
        };
    }

    private static RenewalProjectionRecord ReadRenewalProjection(SqliteDataReader reader)
    {
        return new RenewalProjectionRecord
        {
            ProjectionId = ReadLong(reader, "projection_id"),
            ClientId = ReadLong(reader, "client_id"),
            BenefitLine = ReadString(reader, "benefit_line"),
            CurrentPremium = ReadDecimal(reader, "current_premium"),
            ProjectedClaims = ReadDecimal(reader, "projected_claims"),
            TrendAdjustment = ReadDecimal(reader, "trend_adjustment"),
            RequiredPremium = ReadDecimal(reader, "required_premium"),
            ProposedPremium = ReadDecimal(reader, "proposed_premium"),
            NegotiatedPremium = ReadDecimal(reader, "negotiated_premium"),
            RenewalChangePercent = ReadDecimal(reader, "renewal_change_percent"),
            Savings = ReadDecimal(reader, "savings")
        };
    }

    private static RenewalVersionRecord ReadRenewalVersion(SqliteDataReader reader)
    {
        return new RenewalVersionRecord
        {
            VersionId = ReadLong(reader, "version_id"),
            ClientId = ReadLong(reader, "client_id"),
            RenewalId = ReadNullableLong(reader, "renewal_id"),
            RenewalYear = ReadInt(reader, "renewal_year"),
            Carrier = ReadString(reader, "carrier"),
            VersionName = ReadString(reader, "version_name"),
            SourceModule = ReadString(reader, "source_module"),
            Status = ReadString(reader, "status"),
            IsSelected = ReadInt(reader, "is_selected") == 1,
            EffectiveDate = ReadDateOnly(reader, "effective_date"),
            SourceDocumentId = ReadNullableLong(reader, "source_document_id"),
            DateCreated = ReadDateTime(reader, "date_created") ?? DateTime.MinValue,
            RequestedIncreasePercent = ReadDecimal(reader, "requested_increase_percent"),
            NegotiatedIncreasePercent = ReadDecimal(reader, "negotiated_increase_percent"),
            FinalAcceptedIncreasePercent = ReadDecimal(reader, "final_accepted_increase_percent"),
            NegotiationStatus = ParseNullableEnum<NegotiationStatus>(ReadString(reader, "negotiation_status")),
            NegotiationStage = OptionalText(ReadString(reader, "negotiation_stage"))?.ToString(),
            NegotiationCompletedDate = ReadDateOnly(reader, "negotiation_completed_date"),
            NegotiatedSavingsAnnual = ReadDecimal(reader, "negotiated_savings_annual"),
            NegotiatedSavingsMonthly = ReadDecimal(reader, "negotiated_savings_monthly"),
            NegotiatedSavingsType = ParseNullableEnum<NegotiatedSavingsType>(ReadString(reader, "negotiated_savings_type")),
            NegotiatedBy = OptionalText(ReadString(reader, "negotiated_by"))?.ToString(),
            IndustryBenchmarkPercent = ReadDecimal(reader, "industry_benchmark_percent"),
            IndustryBenchmarkSource = OptionalText(ReadString(reader, "industry_benchmark_source"))?.ToString(),
            IndustryBenchmarkPopulation = OptionalText(ReadString(reader, "industry_benchmark_population"))?.ToString(),
            IndustryBenchmarkYear = ReadInt(reader, "industry_benchmark_year"),
            IndustryBenchmarkConfidence = ReadDecimal(reader, "industry_benchmark_confidence"),
            RenewalRecommendationType = ParseNullableEnum<RenewalRecommendationType>(ReadString(reader, "renewal_recommendation_type")),
            RenewalDecisionReason = OptionalText(ReadString(reader, "renewal_decision_reason"))?.ToString(),
            RenewalStrategySummary = OptionalText(ReadString(reader, "renewal_strategy_summary"))?.ToString(),
            RenewalRiskLevel = ParseNullableEnum<RenewalRiskLevel>(ReadString(reader, "renewal_risk_level")),
            RenewalComplexity = OptionalText(ReadString(reader, "renewal_complexity"))?.ToString(),
            RenewalResult = OptionalText(ReadString(reader, "renewal_result"))?.ToString(),
            RenewalCompletedDate = ReadDateOnly(reader, "renewal_completed_date"),
            RenewalApprovedBy = OptionalText(ReadString(reader, "renewal_approved_by"))?.ToString(),
            RenewalAcceptedDate = ReadDateOnly(reader, "renewal_accepted_date"),
            RenewalVersionNumber = ReadInt(reader, "renewal_version_number"),
            RenewalScenarioId = ReadNullableLong(reader, "renewal_scenario_id"),
            PreviousRenewalScenarioId = ReadNullableLong(reader, "previous_renewal_scenario_id"),
            ScenarioType = ParseNullableEnum<ScenarioType>(ReadString(reader, "scenario_type")),
            HistoryEventId = ReadNullableLong(reader, "history_event_id"),
            RelationshipMilestoneId = OptionalText(ReadString(reader, "relationship_milestone_id"))?.ToString(),
            EvolutionGroupId = OptionalText(ReadString(reader, "evolution_group_id"))?.ToString()
        };
    }

    private static RenewalRateVersionRecord ReadRenewalRateVersion(SqliteDataReader reader)
    {
        return new RenewalRateVersionRecord
        {
            RateVersionId = ReadLong(reader, "rate_version_id"),
            VersionId = ReadLong(reader, "version_id"),
            BenefitLine = ReadString(reader, "benefit_line"),
            ClassName = ReadString(reader, "class"),
            AccountClass = ReadString(reader, "account_class"),
            Volume = ReadDecimal(reader, "volume"),
            Lives = ReadInt(reader, "lives"),
            CurrentRate = ReadDecimal(reader, "current_rate"),
            CurrentPremium = ReadDecimal(reader, "current_premium"),
            ProposedRate = ReadDecimal(reader, "proposed_rate"),
            ProposedPremium = ReadDecimal(reader, "proposed_premium"),
            VersionRate = ReadDecimal(reader, "version_rate"),
            VersionPremium = ReadDecimal(reader, "version_premium"),
            ChangePercent = ReadDecimal(reader, "change_percent"),
            RateBasis = ReadString(reader, "rate_basis"),
            RateValueBasis = FirstNonBlank(ReadString(reader, "rate_value_basis"), RateValueBasisNames.CarrierStated),
            ClientPresentationRateBasis = FirstNonBlank(ReadString(reader, "client_presentation_rate_basis"), RateValueBasisNames.CarrierStated),
            CurrentNetRate = ReadDecimal(reader, "current_net_rate"),
            CurrentGrossRate = ReadDecimal(reader, "current_gross_rate"),
            ProposedNetRate = ReadDecimal(reader, "proposed_net_rate"),
            ProposedGrossRate = ReadDecimal(reader, "proposed_gross_rate"),
            VersionNetRate = ReadDecimal(reader, "version_net_rate"),
            VersionGrossRate = ReadDecimal(reader, "version_gross_rate"),
            RateScheduleSourceFile = ReadString(reader, "rate_schedule_source_file"),
            RateScheduleSourceSheet = ReadString(reader, "rate_schedule_source_sheet"),
            RateScheduleSourceRow = ReadInt(reader, "rate_schedule_source_row"),
            SourceReference = ReadString(reader, "source_reference")
        };
    }

    private static RenewalCalculationRecord ReadRenewalCalculation(SqliteDataReader reader)
    {
        return new RenewalCalculationRecord
        {
            CalculationId = ReadLong(reader, "calculation_id"),
            VersionId = ReadLong(reader, "version_id"),
            BenefitLine = ReadString(reader, "benefit_line"),
            TrendedClaims = ReadDecimal(reader, "trended_claims"),
            ProjectedClaims = ReadDecimal(reader, "projected_claims"),
            RequiredPremium = ReadDecimal(reader, "required_premium"),
            ProjectedLossRatio = ReadDecimal(reader, "projected_loss_ratio"),
            RenewalChangePercent = ReadDecimal(reader, "renewal_change_percent"),
            NegotiatedChangePercent = ReadDecimal(reader, "negotiated_change_percent"),
            Savings = ReadDecimal(reader, "savings")
        };
    }

    private static ClaimsExperienceRecord ReadClaimsExperience(SqliteDataReader reader)
    {
        return new ClaimsExperienceRecord
        {
            ClaimsId = ReadLong(reader, "claims_id"),
            ClientId = ReadLong(reader, "client_id"),
            ExperiencePeriodStart = ReadDateOnly(reader, "experience_period_start"),
            ExperiencePeriodEnd = ReadDateOnly(reader, "experience_period_end"),
            BenefitLine = ReadString(reader, "benefit_line"),
            Premium = ReadDecimal(reader, "premium"),
            PaidClaims = ReadDecimal(reader, "paid_claims"),
            IncurredClaims = ReadDecimal(reader, "incurred_claims"),
            LossRatio = ReadDecimal(reader, "loss_ratio"),
            Lives = ReadInt(reader, "lives"),
            SourceDocumentId = ReadNullableLong(reader, "source_document_id"),
            DateLoaded = ReadDateTime(reader, "date_loaded") ?? DateTime.MinValue
        };
    }

    private static MonthlyClaimsRecord ReadMonthlyClaims(SqliteDataReader reader)
    {
        return new MonthlyClaimsRecord
        {
            ClaimMonthId = ReadLong(reader, "claim_month_id"),
            ClientId = ReadLong(reader, "client_id"),
            Carrier = ReadString(reader, "carrier"),
            BenefitLine = ReadString(reader, "benefit_line"),
            Month = ReadDateOnly(reader, "month"),
            Lives = ReadInt(reader, "lives"),
            Premium = ReadDecimal(reader, "premium"),
            PaidClaims = ReadDecimal(reader, "paid_claims"),
            IncurredClaims = ReadDecimal(reader, "incurred_claims"),
            ClaimsPlusFees = ReadDecimal(reader, "claims_plus_fees"),
            LossRatio = ReadDecimal(reader, "loss_ratio"),
            SourceDocumentId = ReadNullableLong(reader, "source_document_id"),
            DateLoaded = ReadDateTime(reader, "date_loaded") ?? DateTime.MinValue
        };
    }

    private static PlanDesignRecord ReadPlanDesign(SqliteDataReader reader)
    {
        return new PlanDesignRecord
        {
            PlanDesignId = ReadLong(reader, "plan_design_id"),
            ClientId = ReadLong(reader, "client_id"),
            ClassName = ReadString(reader, "class_name"),
            BenefitCategory = ReadString(reader, "benefit_category"),
            BenefitName = ReadString(reader, "benefit_name"),
            CoverageDescription = ReadString(reader, "coverage_description"),
            ReimbursementPercent = ReadDecimal(reader, "reimbursement_percent"),
            MaximumAmount = ReadString(reader, "maximum_amount"),
            Frequency = ReadString(reader, "frequency"),
            WaitingPeriod = ReadString(reader, "waiting_period"),
            TerminationAge = ReadString(reader, "termination_age"),
            SourceDocumentId = ReadNullableLong(reader, "source_document_id"),
            SourceFileName = FirstNonBlank(ReadString(reader, "booklet_source_file_name"), ReadString(reader, "source_file_name")),
            SourceSection = ReadString(reader, "source_section"),
            Notes = ReadString(reader, "notes"),
            SourcePage = ReadInt(reader, "source_page"),
            DateLoaded = ReadDateTime(reader, "date_loaded") ?? DateTime.MinValue
        };
    }

    private static DocumentRecord ReadDocument(SqliteDataReader reader)
    {
        return new DocumentRecord
        {
            DocumentId = ReadLong(reader, "document_id"),
            ClientId = ReadLong(reader, "client_id"),
            DocumentType = ReadString(reader, "document_type"),
            ModuleName = ReadString(reader, "module_name"),
            FileName = ReadString(reader, "file_name"),
            FileType = ReadString(reader, "file_type"),
            FilePath = ReadString(reader, "file_path"),
            FileHash = ReadString(reader, "file_hash"),
            Carrier = ReadString(reader, "carrier"),
            EffectiveDate = ReadDateOnly(reader, "effective_date"),
            PeriodStart = ReadDateOnly(reader, "period_start"),
            PeriodEnd = ReadDateOnly(reader, "period_end"),
            UploadDate = ReadDateTime(reader, "upload_date") ?? DateTime.MinValue,
            ProcessedDate = ReadDateTime(reader, "processed_date") ?? DateTime.MinValue,
            ProcessingStatus = ReadString(reader, "processing_status")
        };
    }

    private static ClientRecord ReadClient(SqliteDataReader reader)
    {
        return new ClientRecord
        {
            ClientId = ReadLong(reader, "client_id"),
            ClientName = ReadString(reader, "client_name"),
            ClientNumber = ReadString(reader, "client_number"),
            Carrier = ReadString(reader, "carrier"),
            Industry = ReadString(reader, "industry"),
            Province = ReadString(reader, "province"),
            Advisor = ReadString(reader, "advisor"),
            EffectiveDate = ReadDateOnly(reader, "effective_date"),
            RenewalMonth = ReadInt(reader, "renewal_month"),
            Status = ReadString(reader, "status"),
            DateCreated = ReadDateTime(reader, "date_created") ?? DateTime.MinValue,
            DateUpdated = ReadDateTime(reader, "date_updated") ?? DateTime.MinValue
        };
    }

    private static PlanBenefitRule ReadPlanBenefitRule(SqliteDataReader reader)
    {
        return new PlanBenefitRule
        {
            PlanBenefitRuleId = ReadLong(reader, "PlanBenefitRuleId"),
            PlanDesignFeatureId = ReadLong(reader, "PlanDesignFeatureId"),
            ClientId = ReadLong(reader, "ClientId"),
            SourceDocumentId = ReadNullableLong(reader, "SourceDocumentId"),
            RuleType = ParseEnum(ReadString(reader, "RuleType"), BenefitRuleType.TieredReimbursement),
            Tier1ReimbursementPercent = ReadDecimal(reader, "Tier1ReimbursementPercent"),
            Tier1ThresholdAmount = ReadDecimal(reader, "Tier1ThresholdAmount"),
            Tier2ReimbursementPercent = ReadDecimal(reader, "Tier2ReimbursementPercent"),
            ThresholdBasis = ParseNullableEnum<BenefitRuleThresholdBasis>(ReadString(reader, "ThresholdBasis")),
            ResetPeriod = ParseNullableEnum<BenefitRuleResetPeriod>(ReadString(reader, "ResetPeriod")),
            CurrencyCode = ReadString(reader, "CurrencyCode"),
            SourceText = ReadString(reader, "SourceText"),
            SourceLocation = ReadString(reader, "SourceLocation"),
            NormalizationStatus = ParseEnum(
                ReadString(reader, "NormalizationStatus"),
                BenefitRuleNormalizationStatus.NeedsReview),
            ConfidenceScore = ReadDecimal(reader, "ConfidenceScore"),
            DisplayOrder = ReadInt(reader, "DisplayOrder") ?? 10,
            IsActive = (ReadInt(reader, "IsActive") ?? 0) == 1,
            EffectiveDate = ReadDateOnly(reader, "EffectiveDate"),
            TerminationDate = ReadDateOnly(reader, "TerminationDate"),
            CreatedAtUtc = ReadDateTime(reader, "CreatedAtUtc") ?? DateTime.MinValue,
            UpdatedAtUtc = ReadDateTime(reader, "UpdatedAtUtc") ?? DateTime.MinValue
        };
    }

    private static T ParseEnum<T>(string value, T fallback) where T : struct, Enum
    {
        return Enum.TryParse<T>(value, true, out var parsed) ? parsed : fallback;
    }

    private static T? ParseNullableEnum<T>(string value) where T : struct, Enum
    {
        return Enum.TryParse<T>(value, true, out var parsed) ? parsed : null;
    }

    private static long ReadLong(SqliteDataReader reader, string name)
    {
        return reader.GetInt64(reader.GetOrdinal(name));
    }

    private static long? ReadNullableLong(SqliteDataReader reader, string name)
    {
        var ordinal = reader.GetOrdinal(name);
        return reader.IsDBNull(ordinal) ? null : reader.GetInt64(ordinal);
    }

    private static int? ReadInt(SqliteDataReader reader, string name)
    {
        var ordinal = reader.GetOrdinal(name);
        return reader.IsDBNull(ordinal) ? null : Convert.ToInt32(reader.GetInt64(ordinal), CultureInfo.InvariantCulture);
    }

    private static decimal? ReadDecimal(SqliteDataReader reader, string name)
    {
        var ordinal = reader.GetOrdinal(name);
        return reader.IsDBNull(ordinal) ? null : Convert.ToDecimal(reader.GetValue(ordinal), CultureInfo.InvariantCulture);
    }

    private static string ReadString(SqliteDataReader reader, string name)
    {
        var ordinal = reader.GetOrdinal(name);
        return reader.IsDBNull(ordinal) ? "" : reader.GetString(ordinal);
    }

    private static DateOnly? ReadDateOnly(SqliteDataReader reader, string name)
    {
        var value = ReadString(reader, name);
        return DateOnly.TryParseExact(value, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var date)
            ? date
            : null;
    }

    private static DateTime? ReadDateTime(SqliteDataReader reader, string name)
    {
        var value = ReadString(reader, name);
        return DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var date)
            ? date
            : null;
    }

    private static string ToDateText(DateOnly? date)
    {
        return date?.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) ?? "";
    }

    private static string ToDateTimeText(DateTime date)
    {
        return date.ToString("O", CultureInfo.InvariantCulture);
    }

    private static long ToLong(object? value)
    {
        return Convert.ToInt64(value, CultureInfo.InvariantCulture);
    }

    private static long? ToNullableLong(object? value)
    {
        return value is null || value == DBNull.Value
            ? null
            : Convert.ToInt64(value, CultureInfo.InvariantCulture);
    }

    private static decimal? ParsePercent(string value)
    {
        return CanonicalSourceLineageValueGuardrails.ParseSinglePercentage(value);
    }

    private static int? ParseInt(string value)
    {
        var match = Regex.Match(value ?? "", @"\d+");
        return match.Success && int.TryParse(match.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed)
            ? parsed
            : null;
    }

    private static int? NormalizeRenewalMonth(int? renewalMonth, DateOnly? effectiveDate)
    {
        if (renewalMonth is >= 1 and <= 12)
        {
            return renewalMonth;
        }

        return effectiveDate?.Month;
    }

    private static string NormalizeHeader(string value)
    {
        return Regex.Replace(value ?? "", @"[^A-Za-z0-9]", "").ToLowerInvariant();
    }

    private static string NormalizeForMatch(string value)
    {
        return Regex.Replace(value ?? "", @"[^A-Za-z0-9]", "").ToLowerInvariant();
    }

    private static string FileTypeFromPath(string path)
    {
        var extension = Path.GetExtension(path);
        return string.IsNullOrWhiteSpace(extension)
            ? ""
            : extension.TrimStart('.').ToUpperInvariant();
    }

    private static string SourceTypeFromFilePath(string path)
    {
        var extension = Path.GetExtension(path);
        return extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) ||
               extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase) ||
               extension.Equals(".xls", StringComparison.OrdinalIgnoreCase) ||
               extension.Equals(".csv", StringComparison.OrdinalIgnoreCase)
            ? SourceTypes.Excel
            : SourceTypes.Pdf;
    }

    private static string FirstNonBlank(params string?[] values)
    {
        return values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value))?.Trim() ?? "";
    }

    private static bool SameFilePath(string left, string right)
    {
        if (string.IsNullOrWhiteSpace(left) || string.IsNullOrWhiteSpace(right))
        {
            return false;
        }

        try
        {
            return Path.GetFullPath(left).Equals(
                Path.GetFullPath(right),
                StringComparison.OrdinalIgnoreCase);
        }
        catch (Exception exception) when (exception is ArgumentException or NotSupportedException or PathTooLongException)
        {
            return false;
        }
    }

    private static string Clean(string? value)
    {
        return value?.Trim() ?? "";
    }

    private static async Task SeedPlanDesignFactorPlaceholdersIfEmptyAsync(SqliteConnection connection)
    {
        using var countCommand = CreateCommand(connection, null, "SELECT COUNT(*) FROM plan_design_factors;");
        var existingRows = Convert.ToInt32(await countCommand.ExecuteScalarAsync(), CultureInfo.InvariantCulture);
        if (existingRows > 0)
        {
            return;
        }

        using var transaction = connection.BeginTransaction();
        foreach (var factor in BuildDefaultPlanDesignFactors())
        {
            await InsertPlanDesignFactorAsync(connection, transaction, factor);
        }

        transaction.Commit();
    }

    private static List<PlanDesignFactorRecord> BuildDefaultPlanDesignFactors()
    {
        var effectiveDate = new DateOnly(2026, 1, 1);
        const string source = "Starter seed assumptions";
        const string notes = "Estimated placeholder only. Replace with real carrier-provided factors before client reliance.";

        return
        [
            CreatePlanDesignFactorSeed("Drugs", "Drug Reimbursement", "80%", "100%", "Reimbursement Change", 1.10m, "Increase", effectiveDate, source, notes),
            CreatePlanDesignFactorSeed("Drugs", "Drug Reimbursement", "100%", "80%", "Reimbursement Change", 0.91m, "Decrease", effectiveDate, source, notes),
            CreatePlanDesignFactorSeed("Drugs", "Generic Drug Substitution", "Brand", "Mandatory Generic", "Mandatory Generic", 0.95m, "Decrease", effectiveDate, source, notes),
            CreatePlanDesignFactorSeed("Dental", "Dental Basic", "80%", "100%", "Reimbursement Change", 1.12m, "Increase", effectiveDate, source, notes),
            CreatePlanDesignFactorSeed("Dental", "Dental Basic", "100%", "80%", "Reimbursement Change", 0.89m, "Decrease", effectiveDate, source, notes),
            CreatePlanDesignFactorSeed("Dental", "Dental Recall", "9 months", "6 months", "Dental Recall Change", 1.05m, "Increase", effectiveDate, source, notes),
            CreatePlanDesignFactorSeed("Dental", "Dental Annual Maximum", "$1,500", "$2,000", "Annual Maximum Change", 1.03m, "Increase", effectiveDate, source, notes),
            CreatePlanDesignFactorSeed("Dental", "Orthodontics", "No Orthodontics", "Add Ortho", "Add Orthodontics", 1.08m, "Increase", effectiveDate, source, notes),
            CreatePlanDesignFactorSeed("Extended Health Care", "Vision", "$250", "$500", "Vision Maximum Change", 1.02m, "Increase", effectiveDate, source, notes),
            CreatePlanDesignFactorSeed("Extended Health Care", "Paramedical", "$500", "$1,000", "Paramedical Maximum Change", 1.04m, "Increase", effectiveDate, source, notes),
            CreatePlanDesignFactorSeed("Long Term Disability", "LTD", "Current monthly maximum", "Increase monthly maximum", "LTD Maximum Change", 1.03m, "Increase", effectiveDate, source, notes),
            CreatePlanDesignFactorSeed("Long Term Disability", "LTD", "Current elimination period", "Reduce elimination period", "LTD Elimination Period Change", 1.05m, "Increase", effectiveDate, source, notes),
            CreatePlanDesignFactorSeed("Basic Life", "Basic Life", "$25,000", "$50,000", "Life Volume Change", 2.00m, "Increase", effectiveDate, source, notes)
        ];
    }

    private static PlanDesignFactorRecord CreatePlanDesignFactorSeed(
        string benefitCategory,
        string benefitLine,
        string currentDesign,
        string proposedDesign,
        string changeType,
        decimal factor,
        string direction,
        DateOnly effectiveDate,
        string sourceDocument,
        string notes)
    {
        return new PlanDesignFactorRecord
        {
            CarrierName = PlanDesignFactorRecord.GeneralPlaceholderCarrier,
            BenefitCategory = benefitCategory,
            BenefitLine = benefitLine,
            CurrentDesign = currentDesign,
            ProposedDesign = proposedDesign,
            ChangeType = changeType,
            Factor = factor,
            FactorDirection = direction,
            EffectiveDate = effectiveDate,
            SourceDocument = sourceDocument,
            Notes = notes,
            ConfidenceLevel = PlanDesignFactorRecord.PlaceholderConfidence
        };
    }

    private static async Task<long> InsertPlanDesignFactorAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        PlanDesignFactorRecord factor)
    {
        NormalizePlanDesignFactor(factor);
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO plan_design_factors (
    carrier_name,
    benefit_category,
    benefit_line,
    current_design,
    proposed_design,
    change_type,
    factor,
    factor_direction,
    effective_date,
    expiry_date,
    source_document,
    notes,
    confidence_level
)
VALUES (
    $carrier_name,
    $benefit_category,
    $benefit_line,
    $current_design,
    $proposed_design,
    $change_type,
    $factor,
    $factor_direction,
    $effective_date,
    $expiry_date,
    $source_document,
    $notes,
    $confidence_level
);
SELECT last_insert_rowid();");
        AddPlanDesignFactorParameters(command, factor);
        return Convert.ToInt64(await command.ExecuteScalarAsync(), CultureInfo.InvariantCulture);
    }

    private static async Task UpdatePlanDesignFactorAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        PlanDesignFactorRecord factor)
    {
        NormalizePlanDesignFactor(factor);
        using var command = CreateCommand(connection, transaction, @"
UPDATE plan_design_factors
SET carrier_name = $carrier_name,
    benefit_category = $benefit_category,
    benefit_line = $benefit_line,
    current_design = $current_design,
    proposed_design = $proposed_design,
    change_type = $change_type,
    factor = $factor,
    factor_direction = $factor_direction,
    effective_date = $effective_date,
    expiry_date = $expiry_date,
    source_document = $source_document,
    notes = $notes,
    confidence_level = $confidence_level,
    updated_at = CURRENT_TIMESTAMP
WHERE id = $id;");
        AddPlanDesignFactorParameters(command, factor);
        AddParameter(command, "$id", factor.Id);
        await command.ExecuteNonQueryAsync();
    }

    private static void AddPlanDesignFactorParameters(SqliteCommand command, PlanDesignFactorRecord factor)
    {
        AddParameter(command, "$carrier_name", factor.CarrierName);
        AddParameter(command, "$benefit_category", factor.BenefitCategory);
        AddParameter(command, "$benefit_line", factor.BenefitLine);
        AddParameter(command, "$current_design", OptionalText(factor.CurrentDesign));
        AddParameter(command, "$proposed_design", OptionalText(factor.ProposedDesign));
        AddParameter(command, "$change_type", factor.ChangeType);
        AddParameter(command, "$factor", factor.Factor);
        AddParameter(command, "$factor_direction", OptionalText(factor.FactorDirection));
        AddParameter(command, "$effective_date", OptionalText(ToDateText(factor.EffectiveDate)));
        AddParameter(command, "$expiry_date", OptionalText(ToDateText(factor.ExpiryDate)));
        AddParameter(command, "$source_document", OptionalText(factor.SourceDocument));
        AddParameter(command, "$notes", OptionalText(factor.Notes));
        AddParameter(command, "$confidence_level", OptionalText(factor.ConfidenceLevel));
    }

    private static async Task<long?> FindPlanDesignFactorMatchAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        PlanDesignFactorRecord factor)
    {
        using var command = CreateCommand(connection, transaction, @"
SELECT id
FROM plan_design_factors
WHERE carrier_name COLLATE NOCASE = $carrier_name
  AND benefit_category COLLATE NOCASE = $benefit_category
  AND benefit_line COLLATE NOCASE = $benefit_line
  AND COALESCE(current_design, '') COLLATE NOCASE = $current_design
  AND COALESCE(proposed_design, '') COLLATE NOCASE = $proposed_design
  AND change_type COLLATE NOCASE = $change_type
  AND COALESCE(effective_date, '') = $effective_date
ORDER BY id DESC
LIMIT 1;");
        AddParameter(command, "$carrier_name", factor.CarrierName);
        AddParameter(command, "$benefit_category", factor.BenefitCategory);
        AddParameter(command, "$benefit_line", factor.BenefitLine);
        AddParameter(command, "$current_design", factor.CurrentDesign);
        AddParameter(command, "$proposed_design", factor.ProposedDesign);
        AddParameter(command, "$change_type", factor.ChangeType);
        AddParameter(command, "$effective_date", ToDateText(factor.EffectiveDate));
        var result = await command.ExecuteScalarAsync();
        return ToNullableLong(result);
    }

    private static List<string> ValidatePlanDesignFactor(PlanDesignFactorRecord factor)
    {
        var errors = new List<string>();
        if (string.IsNullOrWhiteSpace(factor.CarrierName))
        {
            errors.Add("Carrier name is required.");
        }

        if (string.IsNullOrWhiteSpace(factor.BenefitCategory))
        {
            errors.Add("Benefit category is required.");
        }

        if (string.IsNullOrWhiteSpace(factor.BenefitLine))
        {
            errors.Add("Benefit line is required.");
        }

        if (string.IsNullOrWhiteSpace(factor.ChangeType))
        {
            errors.Add("Change type is required.");
        }

        if (factor.Factor <= 0m)
        {
            errors.Add("Factor must be greater than 0.");
        }

        if (factor.EffectiveDate.HasValue &&
            factor.ExpiryDate.HasValue &&
            factor.ExpiryDate.Value < factor.EffectiveDate.Value)
        {
            errors.Add("Expiry date cannot be before effective date.");
        }

        return errors;
    }

    private static void NormalizePlanDesignFactor(PlanDesignFactorRecord factor)
    {
        factor.CarrierName = Clean(factor.CarrierName);
        factor.BenefitCategory = Clean(factor.BenefitCategory);
        factor.BenefitLine = Clean(factor.BenefitLine);
        factor.CurrentDesign = Clean(factor.CurrentDesign);
        factor.ProposedDesign = Clean(factor.ProposedDesign);
        factor.ChangeType = Clean(factor.ChangeType);
        factor.FactorDirection = Clean(factor.FactorDirection);
        factor.SourceDocument = Clean(factor.SourceDocument);
        factor.Notes = Clean(factor.Notes);
        factor.ConfidenceLevel = string.IsNullOrWhiteSpace(factor.ConfidenceLevel)
            ? "Carrier Provided"
            : Clean(factor.ConfidenceLevel);

        if (string.IsNullOrWhiteSpace(factor.FactorDirection) && factor.Factor > 0m)
        {
            factor.FactorDirection = factor.Factor > 1m ? "Increase" : factor.Factor < 1m ? "Decrease" : "Neutral";
        }
    }

    private static bool MatchesPlanDesignFactorFilter(PlanDesignFactorRecord factor, PlanDesignFactorFilter filter)
    {
        if (!string.IsNullOrWhiteSpace(filter.SearchTerm) &&
            !PlanDesignFactorContains(factor, filter.SearchTerm))
        {
            return false;
        }

        if (!FilterEquals(factor.CarrierName, filter.CarrierName) ||
            !FilterEquals(factor.BenefitCategory, filter.BenefitCategory) ||
            !FilterContains(factor.BenefitLine, filter.BenefitLine) ||
            !FilterEquals(factor.ChangeType, filter.ChangeType) ||
            !FilterEquals(factor.ConfidenceLevel, filter.ConfidenceLevel))
        {
            return false;
        }

        return !filter.EffectiveDate.HasValue || factor.IsActiveOn(filter.EffectiveDate.Value);
    }

    private static bool PlanDesignFactorContains(PlanDesignFactorRecord factor, string searchTerm)
    {
        return FilterContains(factor.CarrierName, searchTerm) ||
               FilterContains(factor.BenefitCategory, searchTerm) ||
               FilterContains(factor.BenefitLine, searchTerm) ||
               FilterContains(factor.CurrentDesign, searchTerm) ||
               FilterContains(factor.ProposedDesign, searchTerm) ||
               FilterContains(factor.ChangeType, searchTerm) ||
               FilterContains(factor.SourceDocument, searchTerm) ||
               FilterContains(factor.Notes, searchTerm);
    }

    private static bool FilterEquals(string actual, string expected)
    {
        return string.IsNullOrWhiteSpace(expected) ||
               actual.Equals(expected.Trim(), StringComparison.OrdinalIgnoreCase);
    }

    private static bool FilterContains(string actual, string expected)
    {
        return string.IsNullOrWhiteSpace(expected) ||
               actual.Contains(expected.Trim(), StringComparison.OrdinalIgnoreCase);
    }

    private static PlanDesignFactorRecord ReadPlanDesignFactor(SqliteDataReader reader)
    {
        return new PlanDesignFactorRecord
        {
            Id = ReadLong(reader, "id"),
            CarrierName = ReadString(reader, "carrier_name"),
            BenefitCategory = ReadString(reader, "benefit_category"),
            BenefitLine = ReadString(reader, "benefit_line"),
            CurrentDesign = ReadString(reader, "current_design"),
            ProposedDesign = ReadString(reader, "proposed_design"),
            ChangeType = ReadString(reader, "change_type"),
            Factor = ReadDecimal(reader, "factor") ?? 0m,
            FactorDirection = ReadString(reader, "factor_direction"),
            EffectiveDate = ReadDateOnly(reader, "effective_date"),
            ExpiryDate = ReadDateOnly(reader, "expiry_date"),
            SourceDocument = ReadString(reader, "source_document"),
            Notes = ReadString(reader, "notes"),
            ConfidenceLevel = ReadString(reader, "confidence_level"),
            CreatedAt = ReadDateTime(reader, "created_at") ?? DateTime.MinValue,
            UpdatedAt = ReadDateTime(reader, "updated_at") ?? DateTime.MinValue
        };
    }

    private static Dictionary<string, int> BuildPlanDesignFactorHeaderMap(IXLRow headerRow)
    {
        var headers = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (var cell in headerRow.CellsUsed())
        {
            var key = NormalizeHeader(cell.GetString());
            if (!string.IsNullOrWhiteSpace(key))
            {
                headers[key] = cell.Address.ColumnNumber;
            }
        }

        return headers;
    }

    private static PlanDesignFactorRecord ReadPlanDesignFactorImportRow(
        IXLRow row,
        IReadOnlyDictionary<string, int> headers,
        string sourceFileName)
    {
        var sourceDocument = ReadPlanDesignFactorCellText(row, headers, "sourcedocument");
        return new PlanDesignFactorRecord
        {
            CarrierName = ReadPlanDesignFactorCellText(row, headers, "carrier"),
            BenefitCategory = ReadPlanDesignFactorCellText(row, headers, "benefitcategory"),
            BenefitLine = ReadPlanDesignFactorCellText(row, headers, "benefitline"),
            CurrentDesign = ReadPlanDesignFactorCellText(row, headers, "currentdesign"),
            ProposedDesign = ReadPlanDesignFactorCellText(row, headers, "proposeddesign"),
            ChangeType = ReadPlanDesignFactorCellText(row, headers, "changetype"),
            Factor = ReadPlanDesignFactorDecimal(row, headers, "factor") ?? 0m,
            FactorDirection = ReadPlanDesignFactorCellText(row, headers, "direction"),
            EffectiveDate = ReadPlanDesignFactorDate(row, headers, "effectivedate"),
            ExpiryDate = ReadPlanDesignFactorDate(row, headers, "expirydate"),
            SourceDocument = string.IsNullOrWhiteSpace(sourceDocument) ? sourceFileName : sourceDocument,
            ConfidenceLevel = ReadPlanDesignFactorCellText(row, headers, "confidencelevel"),
            Notes = ReadPlanDesignFactorCellText(row, headers, "notes")
        };
    }

    private static string ReadPlanDesignFactorCellText(IXLRow row, IReadOnlyDictionary<string, int> headers, string headerKey)
    {
        return headers.TryGetValue(headerKey, out var column)
            ? row.Cell(column).GetString().Trim()
            : "";
    }

    private static decimal? ReadPlanDesignFactorDecimal(IXLRow row, IReadOnlyDictionary<string, int> headers, string headerKey)
    {
        if (!headers.TryGetValue(headerKey, out var column))
        {
            return null;
        }

        var cell = row.Cell(column);
        if (cell.TryGetValue<decimal>(out var value))
        {
            return value;
        }

        var text = cell.GetString().Trim().Replace(",", "", StringComparison.Ordinal);
        return decimal.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out value) ||
               decimal.TryParse(text, NumberStyles.Any, CultureInfo.CurrentCulture, out value)
            ? value
            : null;
    }

    private static DateOnly? ReadPlanDesignFactorDate(IXLRow row, IReadOnlyDictionary<string, int> headers, string headerKey)
    {
        if (!headers.TryGetValue(headerKey, out var column))
        {
            return null;
        }

        var cell = row.Cell(column);
        if (cell.TryGetValue<DateTime>(out var dateTime))
        {
            return DateOnly.FromDateTime(dateTime);
        }

        var text = cell.GetString().Trim();
        if (DateOnly.TryParseExact(text, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var date))
        {
            return date;
        }

        return DateTime.TryParse(text, CultureInfo.CurrentCulture, DateTimeStyles.None, out dateTime) ||
               DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out dateTime)
            ? DateOnly.FromDateTime(dateTime)
            : null;
    }

    private static object? OptionalText(string value)
    {
        return string.IsNullOrWhiteSpace(value) ? null : value.Trim();
    }

    private static readonly string[] PlanDesignFactorExportHeaders =
    [
        "Carrier",
        "Benefit Category",
        "Benefit Line",
        "Current Design",
        "Proposed Design",
        "Change Type",
        "Factor",
        "Direction",
        "Effective Date",
        "Expiry Date",
        "Source Document",
        "Confidence Level",
        "Notes"
    ];

    private sealed class DocumentSaveResult
    {
        public long? DocumentId { get; init; }
        public string FilePath { get; init; } = "";
        public string FileHash { get; init; } = "";
        public bool WasReprocessed { get; init; }
    }
}

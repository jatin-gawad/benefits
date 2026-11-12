using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using Microsoft.Data.Sqlite;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed partial class DatabaseService
{
    internal async Task<CanonicalClaimsImportResult> PersistCanonicalClaimsAsync(
        CanonicalClaimsImportEnvelope envelope,
        IReadOnlyCollection<CanonicalClaimsValidationResult> validation,
        IReadOnlyCollection<CanonicalClaimsParityResult> parity)
    {
        await InitializeCoreAsync();
        using var connection = await OpenConnectionAsync();
        using var transaction = connection.BeginTransaction();
        try
        {
            var key = BuildCanonicalIdempotencyKey(envelope);
            var existingReportId = await FindCanonicalReportIdAsync(connection, transaction, key);
            var reportId = existingReportId ?? await InsertCanonicalReportAsync(connection, transaction, envelope, key, validation);
            if (existingReportId.HasValue)
            {
                await UpdateCanonicalReportAsync(connection, transaction, reportId, envelope, validation);
                await DeleteCanonicalReportChildrenAsync(connection, transaction, reportId);
            }

            await InsertCanonicalPeriodAsync(connection, transaction, reportId, envelope.Report.Period);
            foreach (var fact in envelope.FinancialFacts)
            {
                await InsertCanonicalFinancialFactAsync(connection, transaction, reportId, fact);
                await InsertCanonicalEvidenceAsync(connection, transaction, reportId, "Financial", fact);
            }
            foreach (var fact in envelope.MonthlyFacts)
            {
                await InsertCanonicalMonthlyFactAsync(connection, transaction, reportId, fact);
                await InsertCanonicalEvidenceAsync(connection, transaction, reportId, "Monthly", fact);
            }
            foreach (var fact in envelope.CategoryFacts)
            {
                await InsertCanonicalCategoryFactAsync(connection, transaction, reportId, fact);
                await InsertCanonicalEvidenceAsync(connection, transaction, reportId, "Category", fact);
            }
            foreach (var fact in envelope.DistributionFacts)
            {
                await InsertCanonicalDistributionFactAsync(connection, transaction, reportId, fact);
                await InsertCanonicalEvidenceAsync(connection, transaction, reportId, "Distribution", fact);
            }
            foreach (var fact in envelope.BenchmarkFacts)
            {
                await InsertCanonicalBenchmarkFactAsync(connection, transaction, reportId, fact);
                await InsertCanonicalEvidenceAsync(connection, transaction, reportId, "Benchmark", fact);
            }
            foreach (var fact in envelope.UtilizationFacts)
            {
                await InsertCanonicalUtilizationFactAsync(connection, transaction, reportId, fact);
                await InsertCanonicalUtilizationEvidenceAsync(connection, transaction, reportId, fact);
            }
            foreach (var revision in envelope.FactRevisions)
            {
                await InsertCanonicalFactRevisionAsync(connection, transaction, reportId, revision);
            }

            foreach (var item in validation)
            {
                await InsertCanonicalValidationAsync(connection, transaction, reportId, item);
                if (envelope.Report.ImportBatchId.HasValue)
                {
                    await InsertValidationIssueAsync(connection, transaction, new ValidationIssueRecord
                    {
                        ImportBatchId = envelope.Report.ImportBatchId,
                        SourceDocumentId = envelope.Report.SourceDocumentId,
                        IssueCode = item.Code,
                        Severity = item.Severity,
                        TableName = "CanonicalClaims",
                        FieldName = item.FieldName,
                        Message = item.Message,
                        RawSourceReference = string.IsNullOrWhiteSpace(item.SourceReference) ? item.FactId ?? "" : item.SourceReference,
                        CreatedAt = DateTime.Now
                    });
                }
            }
            foreach (var item in parity)
                await InsertCanonicalParityAsync(connection, transaction, reportId, item);

            if (envelope.Report.ImportBatchId.HasValue)
            {
                var completionStatus = HasCanonicalWarnings(validation) ? ImportStatuses.NeedsReview : ImportStatuses.Completed;
                await InsertImportLogAsync(
                    connection,
                    transaction,
                    envelope.Report.ImportBatchId.Value,
                    "Canonical Claims Import",
                    completionStatus,
                    $"Canonical report {reportId} stored with {FactCount(envelope)} fact(s); retry={existingReportId.HasValue}.");
            }

            transaction.Commit();
            var result = new CanonicalClaimsImportResult
            {
                Succeeded = true,
                WasIdempotentRetry = existingReportId.HasValue,
                CanonicalReportId = reportId,
                PersistedFactCount = FactCount(envelope),
                CompletionStatus = HasCanonicalWarnings(validation) ? ImportStatuses.NeedsReview : ImportStatuses.Completed,
                Message = existingReportId.HasValue ? "Canonical claims import was safely refreshed without duplicates." : "Canonical claims import was stored."
            };
            result.ValidationResults.AddRange(validation);
            result.ParityResults.AddRange(parity);
            return result;
        }
        catch (Exception ex)
        {
            transaction.Rollback();
            AppLogger.Error(ex);
            return new CanonicalClaimsImportResult { Succeeded = false, CompletionStatus = ImportStatuses.Failed, Message = $"Canonical claims transaction was rolled back: {ex.Message}" };
        }
    }

    private static int FactCount(CanonicalClaimsImportEnvelope envelope) => envelope.FinancialFacts.Count + envelope.MonthlyFacts.Count + envelope.CategoryFacts.Count + envelope.DistributionFacts.Count + envelope.BenchmarkFacts.Count + envelope.UtilizationFacts.Count + envelope.FactRevisions.Count;

    private static bool HasCanonicalWarnings(IEnumerable<CanonicalClaimsValidationResult> validation) =>
        validation.Any(item => item.Severity.Equals("Warning", StringComparison.OrdinalIgnoreCase));

    private static string BuildCanonicalIdempotencyKey(CanonicalClaimsImportEnvelope envelope)
    {
        var report = envelope.Report;
        var raw = $"{report.ClientId}|{report.Carrier.Trim().ToUpperInvariant()}|{report.PolicyNumber.Trim().ToUpperInvariant()}|{DateText(report.Period.Start)}|{DateText(report.Period.End)}|{report.SourceDocumentId}|{envelope.MappingVersion}";
        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(raw))).ToLowerInvariant();
    }

    private static async Task<long?> FindCanonicalReportIdAsync(SqliteConnection connection, SqliteTransaction transaction, string key)
    {
        using var command = CreateCommand(connection, transaction, "SELECT canonical_claims_report_id FROM CanonicalClaimsReport WHERE idempotency_key = $key LIMIT 1;");
        AddParameter(command, "$key", key);
        var value = await command.ExecuteScalarAsync();
        return value is null || value == DBNull.Value ? null : Convert.ToInt64(value, CultureInfo.InvariantCulture);
    }

    private static async Task<long> InsertCanonicalReportAsync(SqliteConnection connection, SqliteTransaction transaction, CanonicalClaimsImportEnvelope envelope, string key, IReadOnlyCollection<CanonicalClaimsValidationResult> validation)
    {
        var report = envelope.Report;
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO CanonicalClaimsReport (idempotency_key, client_id, carrier_id, carrier_name, policy_number, division, class,
 import_batch_id, source_document_id, source_document_version_id, mapping_version, validation_status, created_at, updated_at)
VALUES ($key, $client, $carrier_id, $carrier, $policy, $division, $class, $batch, $document, $document_version,
 $mapping, $status, $created, $updated); SELECT last_insert_rowid();");
        AddReportParameters(command, envelope, key, validation);
        return Convert.ToInt64(await command.ExecuteScalarAsync(), CultureInfo.InvariantCulture);
    }

    private static async Task UpdateCanonicalReportAsync(SqliteConnection connection, SqliteTransaction transaction, long reportId, CanonicalClaimsImportEnvelope envelope, IReadOnlyCollection<CanonicalClaimsValidationResult> validation)
    {
        using var command = CreateCommand(connection, transaction, @"
UPDATE CanonicalClaimsReport SET carrier_id=$carrier_id, carrier_name=$carrier, policy_number=$policy, division=$division,
 class=$class, import_batch_id=$batch, source_document_id=$document, source_document_version_id=$document_version,
 mapping_version=$mapping, validation_status=$status, updated_at=$updated
WHERE canonical_claims_report_id=$report_id;");
        AddReportParameters(command, envelope, BuildCanonicalIdempotencyKey(envelope), validation);
        AddParameter(command, "$report_id", reportId);
        await command.ExecuteNonQueryAsync();
    }

    private static void AddReportParameters(SqliteCommand command, CanonicalClaimsImportEnvelope envelope, string key, IReadOnlyCollection<CanonicalClaimsValidationResult> validation)
    {
        var report = envelope.Report;
        AddParameter(command, "$key", key); AddParameter(command, "$client", report.ClientId); AddParameter(command, "$carrier_id", report.CarrierId);
        AddParameter(command, "$carrier", report.Carrier); AddParameter(command, "$policy", report.PolicyNumber); AddParameter(command, "$division", report.Division);
        AddParameter(command, "$class", report.Class); AddParameter(command, "$batch", report.ImportBatchId); AddParameter(command, "$document", report.SourceDocumentId);
        AddParameter(command, "$document_version", report.SourceDocumentVersionId); AddParameter(command, "$mapping", envelope.MappingVersion);
        AddParameter(command, "$status", HasCanonicalWarnings(validation) ? CanonicalClaimsValidationStatus.NeedsReview.ToString() : CanonicalClaimsValidationStatus.Valid.ToString());
        AddParameter(command, "$created", report.CreatedAt.ToString("O", CultureInfo.InvariantCulture)); AddParameter(command, "$updated", DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture));
    }

    private static async Task DeleteCanonicalReportChildrenAsync(SqliteConnection connection, SqliteTransaction transaction, long reportId)
    {
        foreach (var table in new[] { "CanonicalClaimsEvidence", "CanonicalClaimsValidation", "CanonicalClaimsParityResult", "CanonicalClaimsFactRevision", "CanonicalClaimsUtilizationFact", "CanonicalClaimsFinancialFact", "CanonicalClaimsMonthlyFact", "CanonicalClaimsCategoryFact", "CanonicalClaimsDistributionFact", "CanonicalClaimsBenchmarkFact", "CanonicalClaimsPeriod" })
        {
            using var command = CreateCommand(connection, transaction, $"DELETE FROM {table} WHERE canonical_claims_report_id=$report_id;");
            AddParameter(command, "$report_id", reportId);
            await command.ExecuteNonQueryAsync();
        }
    }

    private static async Task InsertCanonicalPeriodAsync(SqliteConnection connection, SqliteTransaction transaction, long reportId, CanonicalClaimsPeriod period)
    {
        using var command = CreateCommand(connection, transaction, "INSERT INTO CanonicalClaimsPeriod (canonical_claims_report_id, period_start, period_end, period_type) VALUES ($report,$start,$end,$type);");
        AddParameter(command, "$report", reportId); AddParameter(command, "$start", DateText(period.Start)); AddParameter(command, "$end", DateText(period.End)); AddParameter(command, "$type", period.PeriodType);
        await command.ExecuteNonQueryAsync();
    }

    private static async Task InsertCanonicalFinancialFactAsync(SqliteConnection c, SqliteTransaction t, long reportId, CanonicalClaimFinancialFact fact)
    {
        using var command = CreateCommand(c, t, @"INSERT INTO CanonicalClaimsFinancialFact
(canonical_claims_report_id,fact_id,benefit,coverage,category,division,class,measurement_type,dimension_type,dimension_value,premium,paid_claims,incurred_claims,loss_ratio,lives,currency,unit_of_measure,availability_state,validation_status,legacy_source_record_id,created_at,updated_at)
VALUES ($report,$id,$benefit,$coverage,$category,$division,$class,$measurement,$dimension_type,$dimension_value,$premium,$paid,$incurred,$loss_ratio,$lives,$currency,$unit,$availability,$validation,$legacy,$created,$updated);");
        AddCommonFactParameters(command, reportId, fact); AddParameter(command, "$premium", fact.Premium); AddParameter(command, "$paid", fact.PaidClaims); AddParameter(command, "$incurred", fact.IncurredClaims); AddParameter(command, "$loss_ratio", fact.LossRatio); AddParameter(command, "$lives", fact.Lives);
        await command.ExecuteNonQueryAsync();
    }

    private static async Task InsertCanonicalMonthlyFactAsync(SqliteConnection c, SqliteTransaction t, long reportId, CanonicalClaimMonthlyFact fact)
    {
        using var command = CreateCommand(c, t, @"INSERT INTO CanonicalClaimsMonthlyFact
(canonical_claims_report_id,fact_id,month,benefit,coverage,category,division,class,measurement_type,dimension_type,dimension_value,premium,paid_claims,incurred_claims,claims_plus_fees,loss_ratio,lives,currency,unit_of_measure,availability_state,validation_status,legacy_source_record_id,created_at,updated_at)
VALUES ($report,$id,$month,$benefit,$coverage,$category,$division,$class,$measurement,$dimension_type,$dimension_value,$premium,$paid,$incurred,$fees,$loss_ratio,$lives,$currency,$unit,$availability,$validation,$legacy,$created,$updated);");
        AddCommonFactParameters(command, reportId, fact); AddParameter(command, "$month", DateText(fact.Month)); AddParameter(command, "$premium", fact.Premium); AddParameter(command, "$paid", fact.PaidClaims); AddParameter(command, "$incurred", fact.IncurredClaims); AddParameter(command, "$fees", fact.ClaimsPlusFees); AddParameter(command, "$loss_ratio", fact.LossRatio); AddParameter(command, "$lives", fact.Lives);
        await command.ExecuteNonQueryAsync();
    }

    private static async Task InsertCanonicalCategoryFactAsync(SqliteConnection c, SqliteTransaction t, long reportId, CanonicalClaimCategoryFact fact)
    {
        using var command=CreateCommand(c,t,"INSERT INTO CanonicalClaimsCategoryFact (canonical_claims_report_id,fact_id,benefit,coverage,category,division,class,measurement_type,dimension_type,dimension_value,monetary_value,count_value,percentage_value,currency,unit_of_measure,availability_state,validation_status,legacy_source_record_id,created_at,updated_at) VALUES ($report,$id,$benefit,$coverage,$category,$division,$class,$measurement,$dimension_type,$dimension_value,$money,$count,$percent,$currency,$unit,$availability,$validation,$legacy,$created,$updated);");
        AddCommonFactParameters(command,reportId,fact); AddParameter(command,"$money",fact.MonetaryValue); AddParameter(command,"$count",fact.CountValue); AddParameter(command,"$percent",fact.PercentageValue); await command.ExecuteNonQueryAsync();
    }

    private static async Task InsertCanonicalDistributionFactAsync(SqliteConnection c, SqliteTransaction t, long reportId, CanonicalClaimDistributionFact fact)
    {
        using var command=CreateCommand(c,t,"INSERT INTO CanonicalClaimsDistributionFact (canonical_claims_report_id,fact_id,benefit,coverage,category,division,class,measurement_type,dimension_type,dimension_value,monetary_value,count_value,percentage_value,average_amount,currency,unit_of_measure,availability_state,validation_status,legacy_source_record_id,created_at,updated_at) VALUES ($report,$id,$benefit,$coverage,$category,$division,$class,$measurement,$dimension_type,$dimension_value,$money,$count,$percent,$average,$currency,$unit,$availability,$validation,$legacy,$created,$updated);");
        AddCommonFactParameters(command,reportId,fact); AddParameter(command,"$money",fact.MonetaryValue); AddParameter(command,"$count",fact.CountValue); AddParameter(command,"$percent",fact.PercentageValue); AddParameter(command,"$average",fact.AverageAmount); await command.ExecuteNonQueryAsync();
    }

    private static async Task InsertCanonicalBenchmarkFactAsync(SqliteConnection c, SqliteTransaction t, long reportId, CanonicalBenchmarkFact fact)
    {
        using var command=CreateCommand(c,t,"INSERT INTO CanonicalClaimsBenchmarkFact (canonical_claims_report_id,fact_id,benefit,coverage,category,division,class,measurement_type,dimension_type,dimension_value,benchmark_definition,monetary_value,count_value,percentage_value,currency,unit_of_measure,availability_state,validation_status,legacy_source_record_id,created_at,updated_at) VALUES ($report,$id,$benefit,$coverage,$category,$division,$class,$measurement,$dimension_type,$dimension_value,$definition,$money,$count,$percent,$currency,$unit,$availability,$validation,$legacy,$created,$updated);");
        AddCommonFactParameters(command,reportId,fact); AddParameter(command,"$definition",fact.BenchmarkDefinition); AddParameter(command,"$money",fact.MonetaryValue); AddParameter(command,"$count",fact.CountValue); AddParameter(command,"$percent",fact.PercentageValue); await command.ExecuteNonQueryAsync();
    }

    private static async Task InsertCanonicalUtilizationFactAsync(SqliteConnection c, SqliteTransaction t, long reportId, CanonicalClaimsUtilizationFact fact)
    {
        using var command = CreateCommand(c, t, @"INSERT INTO CanonicalClaimsUtilizationFact
(canonical_claims_report_id,fact_id,benefit_code,benefit_name,category_code,category_name,service_code,service_name,province_code,relationship_type,division_code,class_code,unique_claimants,occurrences,submitted_amount,eligible_amount,paid_amount,reported_paid_percentage,currency,validation_state,measurement_start_date,measurement_end_date,source_aggregation_level,mapping_version,record_status,created_at_utc)
VALUES ($report,$id,$benefit_code,$benefit_name,$category_code,$category_name,$service_code,$service_name,$province,$relationship,$division,$class,$claimants,$occurrences,$submitted,$eligible,$paid,$reported_percent,$currency,$validation,$start,$end,$level,$mapping,$status,$created);");
        AddParameter(command,"$report",reportId); AddParameter(command,"$id",fact.CanonicalClaimsUtilizationFactId); AddParameter(command,"$benefit_code",fact.BenefitCode); AddParameter(command,"$benefit_name",fact.BenefitName); AddParameter(command,"$category_code",fact.CategoryCode); AddParameter(command,"$category_name",fact.CategoryName); AddParameter(command,"$service_code",fact.ServiceCode); AddParameter(command,"$service_name",fact.ServiceName); AddParameter(command,"$province",fact.ProvinceCode); AddParameter(command,"$relationship",fact.RelationshipType.ToString()); AddParameter(command,"$division",fact.DivisionCode); AddParameter(command,"$class",fact.ClassCode); AddParameter(command,"$claimants",fact.UniqueClaimants); AddParameter(command,"$occurrences",fact.Occurrences); AddParameter(command,"$submitted",fact.SubmittedAmount); AddParameter(command,"$eligible",fact.EligibleAmount); AddParameter(command,"$paid",fact.PaidAmount); AddParameter(command,"$reported_percent",fact.ReportedPaidPercentage); AddParameter(command,"$currency",fact.Currency); AddParameter(command,"$validation",fact.ValidationStatus.ToString()); AddParameter(command,"$start",DateText(fact.MeasurementStartDate)); AddParameter(command,"$end",DateText(fact.MeasurementEndDate)); AddParameter(command,"$level",fact.SourceAggregationLevel.ToString()); AddParameter(command,"$mapping",fact.MappingVersion); AddParameter(command,"$status",fact.RecordStatus); AddParameter(command,"$created",fact.CreatedAtUtc.ToString("O",CultureInfo.InvariantCulture));
        await command.ExecuteNonQueryAsync();
    }

    private static async Task InsertCanonicalFactRevisionAsync(SqliteConnection c, SqliteTransaction t, long reportId, CanonicalClaimsFactRevision revision)
    {
        using var command = CreateCommand(c, t, @"INSERT INTO CanonicalClaimsFactRevision
(canonical_claims_report_id,revision_id,fact_id,natural_key,previous_source_reference,revised_source_reference,previous_submitted_amount,revised_submitted_amount,previous_eligible_amount,revised_eligible_amount,previous_paid_amount,revised_paid_amount,review_status,created_at_utc)
VALUES ($report,$revision,$fact,$key,$previous_source,$revised_source,$previous_submitted,$revised_submitted,$previous_eligible,$revised_eligible,$previous_paid,$revised_paid,$status,$created);");
        AddParameter(command, "$report", reportId);
        AddParameter(command, "$revision", revision.RevisionId);
        AddParameter(command, "$fact", revision.FactId);
        AddParameter(command, "$key", revision.NaturalKey);
        AddParameter(command, "$previous_source", revision.PreviousSourceReference);
        AddParameter(command, "$revised_source", revision.RevisedSourceReference);
        AddParameter(command, "$previous_submitted", revision.PreviousSubmittedAmount);
        AddParameter(command, "$revised_submitted", revision.RevisedSubmittedAmount);
        AddParameter(command, "$previous_eligible", revision.PreviousEligibleAmount);
        AddParameter(command, "$revised_eligible", revision.RevisedEligibleAmount);
        AddParameter(command, "$previous_paid", revision.PreviousPaidAmount);
        AddParameter(command, "$revised_paid", revision.RevisedPaidAmount);
        AddParameter(command, "$status", revision.ReviewStatus);
        AddParameter(command, "$created", revision.CreatedAtUtc.ToString("O", CultureInfo.InvariantCulture));
        await command.ExecuteNonQueryAsync();
    }
    private static async Task InsertCanonicalUtilizationEvidenceAsync(SqliteConnection c, SqliteTransaction t, long reportId, CanonicalClaimsUtilizationFact fact)
    {
        var bridge = new CanonicalClaimCategoryFact { FactId = fact.CanonicalClaimsUtilizationFactId, Evidence = fact.Evidence };
        await InsertCanonicalEvidenceAsync(c, t, reportId, "Utilization", bridge);
    }

    private static void AddCommonFactParameters(SqliteCommand command, long reportId, CanonicalClaimFact fact)
    {
        AddParameter(command,"$report",reportId); AddParameter(command,"$id",fact.FactId); AddParameter(command,"$benefit",fact.Benefit); AddParameter(command,"$coverage",fact.Coverage); AddParameter(command,"$category",fact.Category); AddParameter(command,"$division",fact.Division); AddParameter(command,"$class",fact.Class); AddParameter(command,"$measurement",fact.MeasurementType); AddParameter(command,"$dimension_type",fact.DimensionType); AddParameter(command,"$dimension_value",fact.DimensionValue); AddParameter(command,"$currency",fact.Currency); AddParameter(command,"$unit",fact.UnitOfMeasure); AddParameter(command,"$availability",fact.Availability.ToString()); AddParameter(command,"$validation",fact.ValidationStatus.ToString()); AddParameter(command,"$legacy",fact.LegacySourceRecordId); AddParameter(command,"$created",fact.CreatedAt.ToString("O",CultureInfo.InvariantCulture)); AddParameter(command,"$updated",fact.UpdatedAt.ToString("O",CultureInfo.InvariantCulture));
    }

    private static async Task InsertCanonicalEvidenceAsync(SqliteConnection c, SqliteTransaction t, long reportId, string factType, CanonicalClaimFact fact)
    {
        var e=fact.Evidence; using var command=CreateCommand(c,t,@"INSERT INTO CanonicalClaimsEvidence (canonical_claims_report_id,fact_type,fact_id,source_document_id,source_document_version_id,import_batch_id,source_page_number,worksheet_name,source_row,source_column_or_cell,source_table_or_section,raw_carrier_label,normalized_label,raw_extracted_value,extraction_method,parser_name,parser_version,mapping_version,confidence_score,evidence_precision,mapping_classification,normalization_rule,derivation_formula,was_normalized,warning_code,warning_message) VALUES ($report,$type,$id,$document,$document_version,$batch,$page,$sheet,$row,$cell,$section,$raw_label,$normalized,$raw_value,$method,$parser,$parser_version,$mapping,$confidence,$precision,$classification,$rule,$formula,$was_normalized,$warning_code,$warning_message);");
        AddParameter(command,"$report",reportId); AddParameter(command,"$type",factType); AddParameter(command,"$id",fact.FactId); AddParameter(command,"$document",e.SourceDocumentId); AddParameter(command,"$document_version",e.SourceDocumentVersionId); AddParameter(command,"$batch",e.ImportBatchId); AddParameter(command,"$page",e.SourcePageNumber); AddParameter(command,"$sheet",e.WorksheetName); AddParameter(command,"$row",e.SourceRow); AddParameter(command,"$cell",e.SourceColumnOrCell); AddParameter(command,"$section",e.SourceTableOrSection); AddParameter(command,"$raw_label",e.RawCarrierLabel); AddParameter(command,"$normalized",e.NormalizedLabel); AddParameter(command,"$raw_value",e.RawExtractedValue); AddParameter(command,"$method",e.ExtractionMethod); AddParameter(command,"$parser",e.ParserName); AddParameter(command,"$parser_version",e.ParserVersion); AddParameter(command,"$mapping",e.MappingVersion); AddParameter(command,"$confidence",e.ConfidenceScore); AddParameter(command,"$precision",e.Precision.ToString()); AddParameter(command,"$classification",e.MappingClassification.ToString()); AddParameter(command,"$rule",e.NormalizationRule); AddParameter(command,"$formula",e.DerivationFormula); AddParameter(command,"$was_normalized",e.WasNormalized ? 1 : 0); AddParameter(command,"$warning_code",e.WarningCode); AddParameter(command,"$warning_message",e.WarningMessage); await command.ExecuteNonQueryAsync();
    }

    private static async Task InsertCanonicalValidationAsync(SqliteConnection c, SqliteTransaction t, long reportId, CanonicalClaimsValidationResult item)
    { using var command=CreateCommand(c,t,"INSERT INTO CanonicalClaimsValidation (canonical_claims_report_id,code,severity,field_name,message,fact_id,mapping_classification,source_reference,normalization_rule,created_at) VALUES ($report,$code,$severity,$field,$message,$fact,$classification,$source,$rule,$created);"); AddParameter(command,"$report",reportId); AddParameter(command,"$code",item.Code); AddParameter(command,"$severity",item.Severity); AddParameter(command,"$field",item.FieldName); AddParameter(command,"$message",item.Message); AddParameter(command,"$fact",item.FactId); AddParameter(command,"$classification",item.MappingClassification?.ToString()); AddParameter(command,"$source",item.SourceReference); AddParameter(command,"$rule",item.NormalizationRule); AddParameter(command,"$created",DateTime.UtcNow.ToString("O",CultureInfo.InvariantCulture)); await command.ExecuteNonQueryAsync(); }

    private static async Task InsertCanonicalParityAsync(SqliteConnection c, SqliteTransaction t, long reportId, CanonicalClaimsParityResult item)
    { using var command=CreateCommand(c,t,"INSERT INTO CanonicalClaimsParityResult (canonical_claims_report_id,metric,benefit,period,legacy_value,canonical_value,variance,tolerance,classification,explanation,related_legacy_record_ids,related_canonical_fact_ids,review_status,created_at) VALUES ($report,$metric,$benefit,$period,$legacy,$canonical,$variance,$tolerance,$classification,$explanation,$legacy_ids,$canonical_ids,$review,$created);"); AddParameter(command,"$report",reportId); AddParameter(command,"$metric",item.Metric); AddParameter(command,"$benefit",item.Benefit); AddParameter(command,"$period",item.Period); AddParameter(command,"$legacy",item.LegacyValue); AddParameter(command,"$canonical",item.CanonicalValue); AddParameter(command,"$variance",item.Variance); AddParameter(command,"$tolerance",item.Tolerance); AddParameter(command,"$classification",item.Classification.ToString()); AddParameter(command,"$explanation",item.Explanation); AddParameter(command,"$legacy_ids",item.RelatedLegacyRecordIds); AddParameter(command,"$canonical_ids",item.RelatedCanonicalFactIds); AddParameter(command,"$review",item.ReviewStatus); AddParameter(command,"$created",DateTime.UtcNow.ToString("O",CultureInfo.InvariantCulture)); await command.ExecuteNonQueryAsync(); }

    private static string? DateText(DateOnly? value) => value?.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
}


using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services.Normalization;
using System.Diagnostics;

namespace RenewalReviewBuilder.App.Services.CanonicalClaims;

/// <summary>The single application-level write boundary for canonical claims.</summary>
public sealed class CanonicalClaimsImportService
{
    private readonly DatabaseService _database;
    private readonly CarrierClaimsCanonicalMapperRegistry _mappers;
    private readonly NormalizedDataValidator _validator;
    private readonly CanonicalClaimsParityService _parity;

    public CanonicalClaimsImportService(DatabaseService database)
    {
        _database = database;
        _mappers = new CarrierClaimsCanonicalMapperRegistry();
        _validator = new NormalizedDataValidator();
        _parity = new CanonicalClaimsParityService();
    }

    public async Task<CanonicalClaimsImportResult?> ImportExistingClaimsAsync(
        ClaimsExperienceDataModelResult legacy,
        long clientId,
        long? sourceDocumentId,
        long? importBatchId)
    {
        var carrier = string.IsNullOrWhiteSpace(legacy.Carrier) ? "Manulife" : legacy.Carrier;
        var mapper = _mappers.Resolve<ClaimsExperienceDataModelResult>(carrier);
        if (mapper is null) return null;

        if (mapper is ICarrierClaimsCanonicalMultiPeriodMapper<ClaimsExperienceDataModelResult> multiPeriodMapper)
        {
            return await MapManyValidateAndPersistAsync(
                carrier,
                multiPeriodMapper.SourceParserName,
                multiPeriodMapper.GetType().Name,
                legacy.Summaries.Count + legacy.MonthlyClaims.Count + legacy.ServiceAdjudicationFacts.Count,
                () => multiPeriodMapper.MapMany(legacy, clientId, sourceDocumentId, importBatchId),
                envelope => _parity.ReconcileUtilization(envelope));
        }

        return await MapValidateAndPersistAsync(
            carrier,
            mapper.SourceParserName,
            mapper.GetType().Name,
            legacy.Summaries.Count + legacy.MonthlyClaims.Count,
            () => mapper.Map(legacy, clientId, sourceDocumentId, importBatchId),
            envelope => _parity.Compare(legacy, envelope));
    }

    /// <summary>Canonical import entry point for the Manulife Claims Service Detail by Relationship workbook.</summary>
    public async Task<CanonicalClaimsImportResult> ImportServiceDetailRelationshipWorkbookAsync(string workbookPath, long clientId, long? sourceDocumentId, long? importBatchId)
    {
        var mapper = _mappers.Resolve("Manulife") as ManulifeClaimsCanonicalMapper ?? throw new InvalidOperationException("Manulife canonical mapper is unavailable.");
        return await MapValidateAndPersistAsync("Manulife", mapper.SourceParserName, mapper.GetType().Name, 0,
            () => mapper.MapServiceDetailWorkbook(workbookPath, clientId, sourceDocumentId, importBatchId),
            envelope => _parity.ReconcileUtilization(envelope));
    }

    public async Task<CanonicalClaimsImportResult?> ImportExistingClaimsAsync(
        CarrierPdfParseResult parsed,
        long clientId,
        long? sourceDocumentId,
        long? importBatchId)
    {
        ArgumentNullException.ThrowIfNull(parsed);
        var carrier = string.IsNullOrWhiteSpace(parsed.Carrier) ? parsed.Pdf.Carrier : parsed.Carrier;
        var mapper = _mappers.Resolve<CarrierPdfParseResult>(carrier);
        if (mapper is null) return null;

        return await MapValidateAndPersistAsync(
            carrier,
            mapper.SourceParserName,
            mapper.GetType().Name,
            parsed.Pdf.ExperienceLines.Count + (parsed.WorkbookFromPdf?.ClaimTrends.Count ?? 0) + (parsed.WorkbookFromPdf?.ClaimsDistributionRows.Count ?? 0),
            () => mapper.Map(parsed, clientId, sourceDocumentId, importBatchId),
            envelope => _parity.Compare(parsed, envelope));
    }

    private async Task<CanonicalClaimsImportResult> MapManyValidateAndPersistAsync(
        string carrier,
        string parserName,
        string mapperName,
        int parsedRows,
        Func<IReadOnlyList<CanonicalClaimsImportEnvelope>> map,
        Func<CanonicalClaimsImportEnvelope, List<CanonicalClaimsParityResult>> compare)
    {
        var timer = Stopwatch.StartNew();
        var allValidation = new List<CanonicalClaimsValidationResult>();
        var allParity = new List<CanonicalClaimsParityResult>();
        var persistedFactCount = 0;
        long? firstReportId = null;
        var wasIdempotentRetry = false;
        var completionStatus = ImportStatuses.Completed;
        try
        {
            var envelopes = map();
            if (envelopes.Count == 0)
            {
                timer.Stop();
                return new CanonicalClaimsImportResult { Succeeded = false, CompletionStatus = ImportStatuses.Failed, Duration = timer.Elapsed, Message = "Canonical claims mapping produced no period reports." };
            }

            foreach (var envelope in envelopes)
            {
                var validation = _validator.Validate(envelope);
                validation.AddRange(envelope.MappingWarnings);
                var parity = compare(envelope);
                AddParityValidation(validation, parity);
                var factCount = envelope.FinancialFacts.Count + envelope.MonthlyFacts.Count + envelope.CategoryFacts.Count + envelope.DistributionFacts.Count + envelope.BenchmarkFacts.Count + envelope.UtilizationFacts.Count;
                var reconciliationStatus = validation.Any(item => item.Code.Contains("reconcil", StringComparison.OrdinalIgnoreCase) || item.Code.Contains("total_missing", StringComparison.OrdinalIgnoreCase)) ? "NeedsReview" : "Passed";
                var parityStatus = parity.All(item => item.Classification is CanonicalClaimsParityClassification.ExactMatch or CanonicalClaimsParityClassification.WithinRoundingTolerance) ? "Passed" : "NeedsReview";
                AppLogger.Info($"CanonicalClaims stage=mapped carrier={carrier} parser_selected={parserName} mapper_selected={mapperName} source_document_id={envelope.Report.SourceDocumentId} client_id={envelope.Report.ClientId} period_start={envelope.Report.Period.Start:yyyy-MM-dd} period_end={envelope.Report.Period.End:yyyy-MM-dd} parsed_rows={parsedRows} canonical_rows={factCount} unmapped_terms={validation.Count(item => item.Code.Contains("unmapped", StringComparison.OrdinalIgnoreCase))} validation_warnings={validation.Count(item => item.Severity.Equals("Warning", StringComparison.OrdinalIgnoreCase))} reconciliation_status={reconciliationStatus} parity_status={parityStatus}");

                if (validation.Any(result => result.Severity.Equals("Error", StringComparison.OrdinalIgnoreCase)))
                {
                    timer.Stop();
                    var failed = new CanonicalClaimsImportResult { Succeeded = false, CompletionStatus = ImportStatuses.Failed, Duration = timer.Elapsed, Message = "Canonical claims validation failed; no canonical facts were written for one or more period reports." };
                    failed.ValidationResults.AddRange(validation);
                    failed.ParityResults.AddRange(parity);
                    AppLogger.Info($"CanonicalClaims stage=persistence carrier={carrier} status=Failed duration_ms={timer.ElapsedMilliseconds} legacy_workflow_status=Preserved");
                    return failed;
                }

                var result = await _database.PersistCanonicalClaimsAsync(envelope, validation, parity);
                allValidation.AddRange(result.ValidationResults);
                allParity.AddRange(result.ParityResults);
                persistedFactCount += result.PersistedFactCount;
                firstReportId ??= result.CanonicalReportId;
                wasIdempotentRetry = wasIdempotentRetry || result.WasIdempotentRetry;
                if (!result.Succeeded)
                {
                    timer.Stop();
                    return new CanonicalClaimsImportResult { Succeeded = false, CompletionStatus = result.CompletionStatus, Duration = timer.Elapsed, Message = result.Message };
                }

                if (result.CompletionStatus == ImportStatuses.NeedsReview)
                {
                    completionStatus = ImportStatuses.NeedsReview;
                }
            }

            timer.Stop();
            var aggregate = new CanonicalClaimsImportResult
            {
                Succeeded = true,
                CompletionStatus = completionStatus,
                Duration = timer.Elapsed,
                PersistedFactCount = persistedFactCount,
                CanonicalReportId = firstReportId,
                WasIdempotentRetry = wasIdempotentRetry,
                Message = $"Canonical claims import stored {envelopes.Count} period report(s)."
            };
            aggregate.ValidationResults.AddRange(allValidation);
            aggregate.ParityResults.AddRange(allParity);
            AppLogger.Info($"CanonicalClaims stage=persistence carrier={carrier} status={aggregate.CompletionStatus} persisted=True periods={envelopes.Count} duration_ms={timer.ElapsedMilliseconds} legacy_workflow_status=Preserved");
            return aggregate;
        }
        catch (Exception exception)
        {
            timer.Stop();
            AppLogger.Error(exception);
            AppLogger.Info($"CanonicalClaims stage=mapping carrier={carrier} mapper={mapperName} status=Failed duration_ms={timer.ElapsedMilliseconds} legacy_workflow_status=Preserved");
            return new CanonicalClaimsImportResult { Succeeded = false, CompletionStatus = ImportStatuses.Failed, Duration = timer.Elapsed, Message = "Canonical claims mapping failed; the legacy result was preserved and no canonical facts were written." };
        }
    }
    private async Task<CanonicalClaimsImportResult> MapValidateAndPersistAsync(
        string carrier,
        string parserName,
        string mapperName,
        int parsedRows,
        Func<CanonicalClaimsImportEnvelope> map,
        Func<CanonicalClaimsImportEnvelope, List<CanonicalClaimsParityResult>> compare)
    {
        var timer = Stopwatch.StartNew();
        try
        {
            var envelope = map();
            var validation = _validator.Validate(envelope);
            validation.AddRange(envelope.MappingWarnings);
            var parity = compare(envelope);
            AddParityValidation(validation, parity);
            var factCount = envelope.FinancialFacts.Count + envelope.MonthlyFacts.Count + envelope.CategoryFacts.Count + envelope.DistributionFacts.Count + envelope.BenchmarkFacts.Count + envelope.UtilizationFacts.Count;
            var reconciliationStatus = validation.Any(item => item.Code.Contains("reconcil", StringComparison.OrdinalIgnoreCase) || item.Code.Contains("total_missing", StringComparison.OrdinalIgnoreCase)) ? "NeedsReview" : "Passed";
            var parityStatus = parity.All(item => item.Classification is CanonicalClaimsParityClassification.ExactMatch or CanonicalClaimsParityClassification.WithinRoundingTolerance) ? "Passed" : "NeedsReview";
            AppLogger.Info($"CanonicalClaims stage=mapped carrier={carrier} parser_selected={parserName} mapper_selected={mapperName} source_document_id={envelope.Report.SourceDocumentId} client_id={envelope.Report.ClientId} period_start={envelope.Report.Period.Start:yyyy-MM-dd} period_end={envelope.Report.Period.End:yyyy-MM-dd} parsed_rows={parsedRows} canonical_rows={factCount} unmapped_terms={validation.Count(item => item.Code.Contains("unmapped", StringComparison.OrdinalIgnoreCase))} validation_warnings={validation.Count(item => item.Severity.Equals("Warning", StringComparison.OrdinalIgnoreCase))} reconciliation_status={reconciliationStatus} parity_status={parityStatus}");

            if (validation.Any(result => result.Severity.Equals("Error", StringComparison.OrdinalIgnoreCase)))
            {
                timer.Stop();
                var failed = new CanonicalClaimsImportResult { Succeeded = false, CompletionStatus = ImportStatuses.Failed, Duration = timer.Elapsed, Message = "Canonical claims validation failed; no canonical facts were written." };
                failed.ValidationResults.AddRange(validation);
                failed.ParityResults.AddRange(parity);
                AppLogger.Info($"CanonicalClaims stage=persistence carrier={carrier} status=Failed duration_ms={timer.ElapsedMilliseconds} legacy_workflow_status=Preserved");
                return failed;
            }

            var result = await _database.PersistCanonicalClaimsAsync(envelope, validation, parity);
            timer.Stop();
            result.Duration = timer.Elapsed;
            AppLogger.Info($"CanonicalClaims stage=persistence carrier={carrier} status={result.CompletionStatus} persisted={result.Succeeded} idempotent_retry={result.WasIdempotentRetry} duration_ms={timer.ElapsedMilliseconds} legacy_workflow_status=Preserved");
            return result;
        }
        catch (Exception exception)
        {
            timer.Stop();
            AppLogger.Error(exception);
            AppLogger.Info($"CanonicalClaims stage=mapping carrier={carrier} mapper={mapperName} status=Failed duration_ms={timer.ElapsedMilliseconds} legacy_workflow_status=Preserved");
            return new CanonicalClaimsImportResult { Succeeded = false, CompletionStatus = ImportStatuses.Failed, Duration = timer.Elapsed, Message = "Canonical claims mapping failed; the legacy result was preserved and no canonical facts were written." };
        }
    }

    private static void AddParityValidation(ICollection<CanonicalClaimsValidationResult> validation, IEnumerable<CanonicalClaimsParityResult> parity)
    {
        foreach (var result in parity.Where(item => item.Classification is CanonicalClaimsParityClassification.MissingCanonicalField or CanonicalClaimsParityClassification.UnexpectedMismatch or CanonicalClaimsParityClassification.ValidationFailure))
        {
            validation.Add(new CanonicalClaimsValidationResult
            {
                Code = "canonical_parity_review",
                Severity = "Warning",
                FieldName = result.Metric,
                Message = result.Explanation,
                FactId = string.IsNullOrWhiteSpace(result.RelatedCanonicalFactIds) ? null : result.RelatedCanonicalFactIds
            });
        }
    }
}



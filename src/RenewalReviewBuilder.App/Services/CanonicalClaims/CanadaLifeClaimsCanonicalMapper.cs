using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services.CanonicalClaims;

/// <summary>Maps existing Canada Life parser output; it performs no PDF or workbook extraction.</summary>
public sealed class CanadaLifeClaimsCanonicalMapper : ICarrierClaimsCanonicalMapper<CarrierPdfParseResult>
{
    public const string CurrentMappingVersion = "CANADA_LIFE_CLAIMS_CANONICAL_V1";
    private static readonly CultureInfo CanadaCulture = CultureInfo.GetCultureInfo("en-CA");
    private static readonly Regex PageRegex = new(@"\bpage\s+(?<page>\d+)\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public string CarrierName => "Canada Life";
    public string SourceParserName => nameof(CanadaLifePdfParser);

    public CanonicalClaimsImportEnvelope Map(CarrierPdfParseResult parsedResult, long clientId, long? sourceDocumentId, long? importBatchId)
    {
        ArgumentNullException.ThrowIfNull(parsedResult);
        var pdf = parsedResult.Pdf ?? new PdfExtractionResult();
        var workbook = parsedResult.WorkbookFromPdf ?? new WorkbookExtractionResult();
        var period = ResolveReportPeriod(pdf.ExperienceLines, workbook.ClaimsDistributionRows);
        var envelope = new CanonicalClaimsImportEnvelope
        {
            MappingVersion = CurrentMappingVersion,
            Report = new CanonicalClaimsReport
            {
                ClientId = clientId,
                Carrier = NormalizeCarrier(parsedResult.Carrier, pdf.Carrier),
                PolicyNumber = pdf.PolicyNumber.Trim(),
                Division = SingleDimension(workbook.ClaimsDistributionRows.Select(row => CanadaLifeClaimsTerminology.NormalizeDimensionIdentifier(row.Division, "Division").CanonicalValue)),
                Class = SingleDimension(workbook.ClaimsDistributionRows.Select(row => CanadaLifeClaimsTerminology.NormalizeDimensionIdentifier(row.Class, "Class").CanonicalValue)),
                ImportBatchId = importBatchId,
                SourceDocumentId = sourceDocumentId,
                Period = period
            }
        };

        AddIdentityWarnings(envelope, parsedResult, clientId, sourceDocumentId, importBatchId);
        MapExperience(envelope, pdf.ExperienceLines, sourceDocumentId, importBatchId);
        MapCategories(envelope, workbook.ClaimTrends, sourceDocumentId, importBatchId);
        MapDistributions(envelope, workbook.ClaimsDistributionRows, sourceDocumentId, importBatchId);
        MapDemographics(envelope, pdf.LifeDemographics.Concat(pdf.LtdDemographics), sourceDocumentId, importBatchId);
        MapPooling(envelope, pdf, sourceDocumentId, importBatchId);
        AddReconciliationWarnings(envelope, workbook.ClaimsDistributionRows, workbook.ClaimsDistributionValidation);
        AddUnavailableFieldResults(envelope);

        if (envelope.FinancialFacts.Count == 0 && envelope.CategoryFacts.Count == 0 && envelope.DistributionFacts.Count == 0)
            Warning(envelope, "canada_life_no_meaningful_claims", "Error", "claims_rows", "The Canada Life parsed result contains no meaningful claims rows.", CanonicalClaimsMappingClassification.Unavailable);
        return envelope;
    }

    private static void MapExperience(CanonicalClaimsImportEnvelope envelope, IEnumerable<ExperienceLine> rows, long? documentId, long? batchId)
    {
        var keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var row in rows)
        {
            var benefit = CanadaLifeClaimsTerminology.NormalizeBenefit(row.Benefit);
            var factId = StableId(documentId, "financial", row.Benefit, row.Period);
            if (!keys.Add(factId)) { Warning(envelope, "canada_life_duplicate_dimension", "Warning", "experience", "A duplicate benefit/reporting-period experience row was retained only once.", CanonicalClaimsMappingClassification.Ambiguous, row.SourceReference, factId); continue; }
            WarnIfUnmapped(envelope, "benefit", benefit, row.SourceReference, factId);
            if (!TryParsePeriod(row.Period, out _, out _)) Warning(envelope, "canada_life_period_unmapped", "Error", "reporting_period", $"The explicit Canada Life period '{row.Period}' could not be normalized.", CanonicalClaimsMappingClassification.Ambiguous, row.SourceReference, factId);
            WarnNegative(envelope, "premium", row.BilledPremium, row.SourceReference, factId);
            WarnNegative(envelope, "paid_claims", row.PaidClaims, row.SourceReference, factId);
            WarnNegative(envelope, "incurred_claims", row.IncurredClaims, row.SourceReference, factId);
            WarnMissing(envelope, "premium", row.BilledPremium, row.SourceReference, factId);
            WarnMissing(envelope, "paid_claims", row.PaidClaims, row.SourceReference, factId);
            WarnMissing(envelope, "incurred_claims", row.IncurredClaims, row.SourceReference, factId);
            if (!row.AverageLives.HasValue) Warning(envelope, "canada_life_membership_unavailable", "Warning", "lives", "Average lives are unavailable for this experience row and were not defaulted to zero.", CanonicalClaimsMappingClassification.Unavailable, row.SourceReference, factId);

            envelope.FinancialFacts.Add(new CanonicalClaimFinancialFact
            {
                FactId = factId,
                Benefit = benefit.CanonicalValue,
                MeasurementType = "ClaimsExperience",
                DimensionType = "ReportingPeriod",
                DimensionValue = row.Period,
                Premium = row.BilledPremium,
                PaidClaims = row.PaidClaims,
                IncurredClaims = row.IncurredClaims,
                LossRatio = row.IncurredLossRatio,
                Lives = row.AverageLives,
                Currency = "",
                UnitOfMeasure = "Source currency unavailable",
                Availability = Availability(row.BilledPremium, row.PaidClaims, row.IncurredClaims),
                Evidence = Evidence(row.SourceReference, "Experience Analysis", row.Benefit, benefit, Raw(row.BilledPremium, row.PaidClaims, row.IncurredClaims, row.IncurredLossRatio), documentId, batchId)
            });

            AddPercentageBenchmark(envelope, row, benefit, "PaidLossRatio", UnderwritingEngine.CalculateLossRatioPercent(row.PaidClaims, row.BilledPremium), CanonicalClaimsMappingClassification.Derived, "UnderwritingEngine.CalculateLossRatioPercent(PaidClaims, BilledPremium)", documentId, batchId, Raw(row.PaidClaims, row.BilledPremium));
            AddPercentageBenchmark(envelope, row, benefit, "Credibility", row.Credibility, CanonicalClaimsMappingClassification.Direct, "", documentId, batchId);
            AddPercentageBenchmark(envelope, row, benefit, "TargetLossRatio", row.TargetLossRatio, CanonicalClaimsMappingClassification.Direct, "", documentId, batchId);
            AddPercentageBenchmark(envelope, row, benefit, "AnnualTrendFactor", row.TrendFactor, CanonicalClaimsMappingClassification.Direct, "", documentId, batchId);
            AddPercentageBenchmark(envelope, row, benefit, "FeeGuideAdjustment", row.ProposedAdjustment, CanonicalClaimsMappingClassification.Direct, "", documentId, batchId);
        }
    }

    private static void MapCategories(CanonicalClaimsImportEnvelope envelope, IEnumerable<ClaimTrend> rows, long? documentId, long? batchId)
    {
        var keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var row in rows)
        {
            var benefit = CanadaLifeClaimsTerminology.NormalizeBenefit(row.Benefit);
            var category = CanadaLifeClaimsTerminology.NormalizeCategory(row.Category);
            var factId = StableId(documentId, "category", row.Benefit, row.Category, row.Period);
            if (!keys.Add(factId)) { Warning(envelope, "canada_life_duplicate_dimension", "Warning", "category", "A duplicate benefit/category/reporting-period row was retained only once.", CanonicalClaimsMappingClassification.Ambiguous, row.SourceReference, factId); continue; }
            WarnIfUnmapped(envelope, "benefit", benefit, row.SourceReference, factId);
            WarnIfUnmapped(envelope, "category", category, row.SourceReference, factId);
            WarnNegative(envelope, "category_amount", row.Amount, row.SourceReference, factId, IsAdjustment(category));
            envelope.CategoryFacts.Add(new CanonicalClaimCategoryFact
            {
                FactId = factId, Benefit = benefit.CanonicalValue, Category = category.CanonicalValue,
                MeasurementType = "ClaimsCategory", DimensionType = "ReportingPeriod", DimensionValue = row.Period,
                MonetaryValue = row.Amount, PercentageValue = row.PercentOfTotal, Currency = "", UnitOfMeasure = "Source currency unavailable",
                Availability = Availability(row.Amount),
                Evidence = Evidence(row.SourceReference, "Claims Distribution", $"{row.Benefit}|{row.Category}", Combined(benefit, category), Raw(row.Amount, row.PercentOfTotal), documentId, batchId)
            });
            if (row.BenchmarkPercent.HasValue)
                envelope.BenchmarkFacts.Add(new CanonicalBenchmarkFact
                {
                    FactId = StableId(documentId, "benchmark", "ClaimsDistribution", row.Benefit, row.Category, row.Period), Benefit = benefit.CanonicalValue, Category = category.CanonicalValue,
                    MeasurementType = "CarrierBenchmarkPercent", DimensionType = "ReportingPeriod", DimensionValue = row.Period,
                    PercentageValue = row.BenchmarkPercent, UnitOfMeasure = "Percent", Availability = Availability(row.BenchmarkPercent), BenchmarkDefinition = "Canada Life claims distribution norm percent",
                    Evidence = Evidence(row.SourceReference, "Claims Distribution", $"{row.Benefit}|{row.Category}", Combined(benefit, category), Raw(row.BenchmarkPercent), documentId, batchId)
                });
        }
    }

    private static void MapDistributions(CanonicalClaimsImportEnvelope envelope, IEnumerable<ClaimsDistributionRow> rows, long? documentId, long? batchId)
    {
        var keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var row in rows)
        {
            var benefit = CanadaLifeClaimsTerminology.NormalizeBenefit(row.Benefit);
            var category = CanadaLifeClaimsTerminology.NormalizeCategory(row.ServiceCategory);
            var division = CanadaLifeClaimsTerminology.NormalizeDimensionIdentifier(row.Division, "Division");
            var @class = CanadaLifeClaimsTerminology.NormalizeDimensionIdentifier(row.Class, "Class");
            var factId = StableId(documentId, "distribution", row.Benefit, row.ServiceCategory, row.ReportPeriod, row.Division, row.Class);
            if (!keys.Add(factId)) { Warning(envelope, "canada_life_duplicate_dimension", "Warning", "distribution", "A duplicate claims-distribution dimension was retained only once.", CanonicalClaimsMappingClassification.Ambiguous, row.SourceReference, factId); continue; }
            WarnIfUnmapped(envelope, "benefit", benefit, row.SourceReference, factId);
            WarnIfUnmapped(envelope, "category", category, row.SourceReference, factId);
            WarnIfUnavailableOrAmbiguous(envelope, "division", division, row.SourceReference, factId);
            WarnIfUnavailableOrAmbiguous(envelope, "class", @class, row.SourceReference, factId);
            WarnNegative(envelope, "distribution_amount", row.AmountPaid, row.SourceReference, factId, IsAdjustment(category));
            if (row.Occurrences < 0) Warning(envelope, "canada_life_negative_claim_count", "Warning", "occurrences", "A negative claim count requires source review.", CanonicalClaimsMappingClassification.Ambiguous, row.SourceReference, factId);
            envelope.DistributionFacts.Add(new CanonicalClaimDistributionFact
            {
                FactId = factId, Benefit = benefit.CanonicalValue, Category = category.CanonicalValue, Division = division.CanonicalValue, Class = @class.CanonicalValue,
                MeasurementType = "ClaimsDistribution", DimensionType = "ReportingPeriod", DimensionValue = row.ReportPeriod,
                MonetaryValue = row.AmountPaid, CountValue = row.Occurrences, PercentageValue = row.PercentPaid, AverageAmount = row.AverageClaim,
                Currency = "", UnitOfMeasure = "Source currency unavailable", Availability = Availability(row.AmountPaid, row.PercentPaid),
                Evidence = Evidence(row.SourceReference, "Claims Distribution", $"{row.Benefit}|{row.ServiceCategory}", Combined(benefit, category, division, @class), Raw(row.AmountPaid, row.PercentPaid, row.Occurrences, row.AverageClaim), documentId, batchId)
            });
            if (row.NormPercentPaid.HasValue)
                envelope.BenchmarkFacts.Add(new CanonicalBenchmarkFact
                {
                    FactId = StableId(documentId, "benchmark", "DistributionNorm", row.Benefit, row.ServiceCategory, row.ReportPeriod, row.Division, row.Class), Benefit = benefit.CanonicalValue, Category = category.CanonicalValue,
                    Division = division.CanonicalValue, Class = @class.CanonicalValue, MeasurementType = "CarrierBenchmarkPercent", DimensionType = "ReportingPeriod", DimensionValue = row.ReportPeriod,
                    PercentageValue = row.NormPercentPaid, UnitOfMeasure = "Percent", Availability = Availability(row.NormPercentPaid), BenchmarkDefinition = "Canada Life claims distribution norm percent",
                    Evidence = Evidence(row.SourceReference, "Claims Distribution", $"{row.Benefit}|{row.ServiceCategory}", Combined(benefit, category, division, @class), Raw(row.NormPercentPaid), documentId, batchId)
                });
        }
    }

    private static void MapDemographics(CanonicalClaimsImportEnvelope envelope, IEnumerable<DemographicLine> rows, long? documentId, long? batchId)
    {
        foreach (var row in rows)
        {
            var benefit = CanadaLifeClaimsTerminology.NormalizeBenefit(row.Benefit);
            WarnIfUnmapped(envelope, "benefit", benefit, row.SourceReference, null);
            AddLives(envelope, row, benefit, "Male", row.CurrentMaleLives, documentId, batchId);
            AddLives(envelope, row, benefit, "Female", row.CurrentFemaleLives, documentId, batchId);
        }
    }

    private static void AddLives(CanonicalClaimsImportEnvelope envelope, DemographicLine row, CanadaLifeTermMapping benefit, string sex, int? lives, long? documentId, long? batchId)
    {
        if (!lives.HasValue) return;
        var factId = StableId(documentId, "membership", row.Benefit, row.AgeBand, sex, "Current");
        if (lives < 0) Warning(envelope, "canada_life_negative_membership", "Error", "lives", "Membership counts cannot be negative.", CanonicalClaimsMappingClassification.Ambiguous, row.SourceReference, factId);
        envelope.CategoryFacts.Add(new CanonicalClaimCategoryFact
        {
            FactId = factId, Benefit = benefit.CanonicalValue, Category = row.AgeBand,
            MeasurementType = "CoveredLives", DimensionType = "Sex", DimensionValue = sex,
            CountValue = lives, UnitOfMeasure = "Lives", Availability = lives == 0 ? CanonicalDataAvailability.ReportedZero : CanonicalDataAvailability.Present,
            Evidence = Evidence(row.SourceReference, "Demographics", $"{row.Benefit}|{row.AgeBand}|{sex}", benefit, lives.Value.ToString(CultureInfo.InvariantCulture), documentId, batchId)
        });
    }

    private static void MapPooling(CanonicalClaimsImportEnvelope envelope, PdfExtractionResult pdf, long? documentId, long? batchId)
    {
        if (pdf.PoolingThreshold.HasValue)
            envelope.BenchmarkFacts.Add(PoolingFact("PoolingThreshold", pdf.PoolingThreshold, "Currency amount", documentId, batchId));
        if (pdf.PoolingCharge.HasValue)
            envelope.BenchmarkFacts.Add(PoolingFact("PoolingCharge", pdf.PoolingCharge, "Percent", documentId, batchId));
        if (pdf.PoolingThreshold.HasValue || pdf.PoolingCharge.HasValue)
            Warning(envelope, "canada_life_pooling_evidence_document_only", "Warning", "pooling", "Pooling values were retained, but the existing parser does not retain their page coordinate.", CanonicalClaimsMappingClassification.Direct, pdf.SourcePath);
    }

    private static CanonicalBenchmarkFact PoolingFact(string measurement, decimal? value, string unit, long? documentId, long? batchId) => new()
    {
        FactId = StableId(documentId, "pooling", measurement), MeasurementType = measurement,
        MonetaryValue = unit == "Currency amount" ? value : null, PercentageValue = unit == "Percent" ? value : null,
        Currency = "", UnitOfMeasure = unit, Availability = Availability(value), BenchmarkDefinition = $"Canada Life {measurement}",
        Evidence = new CanonicalClaimsSourceEvidence { SourceDocumentId = documentId, ImportBatchId = batchId, SourceTableOrSection = "Pooling", RawCarrierLabel = measurement, NormalizedLabel = measurement, RawExtractedValue = value?.ToString(CultureInfo.InvariantCulture) ?? "", ExtractionMethod = "Existing Canada Life parser output", ParserName = nameof(CanadaLifePdfParser), ParserVersion = "Legacy-compatible-v1", MappingVersion = CurrentMappingVersion, Precision = CanonicalClaimsEvidencePrecision.Document, MappingClassification = CanonicalClaimsMappingClassification.Direct, WarningCode = "canada_life_pooling_evidence_document_only", WarningMessage = "Existing parser does not retain the pooling page coordinate." }
    };

    private static void AddPercentageBenchmark(CanonicalClaimsImportEnvelope envelope, ExperienceLine row, CanadaLifeTermMapping benefit, string measurement, decimal? value, CanonicalClaimsMappingClassification classification, string formula, long? documentId, long? batchId, string? rawComponents = null)
    {
        if (!value.HasValue) return;
        var evidence = Evidence(row.SourceReference, "Experience Analysis", $"{row.Benefit}|{measurement}", benefit, rawComponents ?? value.Value.ToString(CultureInfo.InvariantCulture), documentId, batchId, classification, formula);
        envelope.BenchmarkFacts.Add(new CanonicalBenchmarkFact { FactId = StableId(documentId, "benchmark", measurement, row.Benefit, row.Period), Benefit = benefit.CanonicalValue, MeasurementType = measurement, DimensionType = "ReportingPeriod", DimensionValue = row.Period, PercentageValue = value, UnitOfMeasure = "Percent", Availability = value == 0m ? CanonicalDataAvailability.ReportedZero : CanonicalDataAvailability.Present, BenchmarkDefinition = $"Canada Life {measurement}", Evidence = evidence });
    }

    private static void AddIdentityWarnings(CanonicalClaimsImportEnvelope envelope, CarrierPdfParseResult parsed, long clientId, long? documentId, long? batchId)
    {
        if (!NormalizeCarrier(parsed.Carrier, parsed.Pdf.Carrier).Equals("Canada Life", StringComparison.OrdinalIgnoreCase)) Warning(envelope, "canada_life_carrier_unconfirmed", "Error", "carrier", "The parsed result is not confirmed as Canada Life.", CanonicalClaimsMappingClassification.Ambiguous);
        if (clientId <= 0) Warning(envelope, "canada_life_client_unconfirmed", "Error", "client_id", "A confirmed client is required.", CanonicalClaimsMappingClassification.Unavailable);
        if (string.IsNullOrWhiteSpace(parsed.Pdf.PolicyNumber)) Warning(envelope, "canada_life_policy_missing", "Error", "policy_number", "Canada Life policy identity is required.", CanonicalClaimsMappingClassification.Unavailable);
        if (!documentId.HasValue) Warning(envelope, "canada_life_source_document_missing", "Error", "source_document_id", "Source document identity is required.", CanonicalClaimsMappingClassification.Unavailable);
        if (!batchId.HasValue) Warning(envelope, "canada_life_import_batch_missing", "Warning", "import_batch_id", "Import-batch identity is missing.", CanonicalClaimsMappingClassification.Unavailable);
    }

    private static void AddReconciliationWarnings(CanonicalClaimsImportEnvelope envelope, IReadOnlyCollection<ClaimsDistributionRow> rows, ClaimsDistributionValidation validation)
    {
        if (rows.Count == 0)
        {
            Warning(envelope, "canada_life_reconciliation_unavailable", "Warning", "reconciliation", "A Canada Life carrier total was not exposed, so required claims reconciliation is unavailable and does not count as passed.", CanonicalClaimsMappingClassification.Unavailable);
            return;
        }
        if (!validation.TotalRowFound) Warning(envelope, "canada_life_distribution_total_missing", "Warning", "distribution_total", "Claims distribution was retained, but the carrier total row is unavailable.", CanonicalClaimsMappingClassification.Unavailable);
        else if (!validation.AmountsBalanceToTotal) Warning(envelope, "canada_life_distribution_not_reconciled", "Warning", "distribution_total", $"Claims distribution service total {validation.SumOfServiceAmounts} does not reconcile to carrier total {validation.TotalAmountPaid} within the existing $1 tolerance.", CanonicalClaimsMappingClassification.Ambiguous);
        foreach (var warning in validation.Warnings) Warning(envelope, "canada_life_distribution_source_warning", "Warning", "claims_distribution", warning, CanonicalClaimsMappingClassification.Ambiguous);
    }

    private static void AddUnavailableFieldResults(CanonicalClaimsImportEnvelope envelope)
    {
        foreach (var field in new[] { "currency", "funding_arrangement", "province", "employee_count", "certificate_count", "dependant_count", "pooled_claims", "non_pooled_claims", "experience_refund", "administrative_expenses", "taxes", "commissions", "reserves" })
            Warning(envelope, $"canada_life_{field}_unavailable", "Info", field, $"{field.Replace('_', ' ')} is not exposed by the existing Canada Life claims parser output and was not inferred.", CanonicalClaimsMappingClassification.Unavailable);
    }

    private static CanonicalClaimsPeriod ResolveReportPeriod(IEnumerable<ExperienceLine> experience, IEnumerable<ClaimsDistributionRow> distributions)
    {
        var periods = new List<(DateOnly Start, DateOnly End)>();
        foreach (var row in experience) if (TryParsePeriod(row.Period, out var start, out var end)) periods.Add((start, end));
        foreach (var row in distributions)
        {
            if (TryParsePeriod(row.PeriodStart, row.PeriodEnd, out var start, out var end) || TryParsePeriod(row.ReportPeriod, out start, out end)) periods.Add((start, end));
        }
        return periods.Count == 0
            ? new CanonicalClaimsPeriod { PeriodType = "Experience" }
            : new CanonicalClaimsPeriod { Start = periods.Min(item => item.Start), End = periods.Max(item => item.End), PeriodType = "Experience" };
    }

    public static bool TryParsePeriod(string label, out DateOnly start, out DateOnly end)
    {
        start = default; end = default;
        var parts = Regex.Split(label.Trim(), @"\s+(?:-|to)\s+", RegexOptions.IgnoreCase);
        if (parts.Length != 2) return false;
        if (TryParseMonth(parts[0], out var startMonth) && TryParseMonth(parts[1], out var endMonth))
        { start = new DateOnly(startMonth.Year, startMonth.Month, 1); end = new DateOnly(endMonth.Year, endMonth.Month, DateTime.DaysInMonth(endMonth.Year, endMonth.Month)); return true; }
        return DateOnly.TryParse(parts[0], CanadaCulture, DateTimeStyles.AllowWhiteSpaces, out start) && DateOnly.TryParse(parts[1], CanadaCulture, DateTimeStyles.AllowWhiteSpaces, out end);
    }

    private static bool TryParsePeriod(string startText, string endText, out DateOnly start, out DateOnly end)
    {
        start = default;
        end = default;
        return DateOnly.TryParse(startText, CanadaCulture, DateTimeStyles.AllowWhiteSpaces, out start)
            && DateOnly.TryParse(endText, CanadaCulture, DateTimeStyles.AllowWhiteSpaces, out end);
    }

    private static bool TryParseMonth(string value, out DateOnly month)
    {
        foreach (var format in new[] { "MMM yyyy", "MMMM yyyy" }) if (DateOnly.TryParseExact(value, format, CanadaCulture, DateTimeStyles.AllowWhiteSpaces, out month)) return true;
        month = default; return false;
    }

    private static CanonicalClaimsSourceEvidence Evidence(string sourceReference, string section, string rawLabel, CanadaLifeTermMapping mapping, string rawValue, long? documentId, long? batchId, CanonicalClaimsMappingClassification? overrideClassification = null, string formula = "")
    {
        var page = Page(sourceReference);
        return new CanonicalClaimsSourceEvidence { SourceDocumentId = documentId, ImportBatchId = batchId, SourcePageNumber = page, SourceTableOrSection = section, RawCarrierLabel = rawLabel, NormalizedLabel = mapping.CanonicalValue, RawExtractedValue = rawValue, ExtractionMethod = "Existing Canada Life parser output", ParserName = nameof(CanadaLifePdfParser), ParserVersion = "Legacy-compatible-v1", MappingVersion = CurrentMappingVersion, Precision = page.HasValue ? CanonicalClaimsEvidencePrecision.Page : CanonicalClaimsEvidencePrecision.Document, MappingClassification = overrideClassification ?? mapping.Classification, NormalizationRule = mapping.RuleName, DerivationFormula = formula, WasNormalized = mapping.Classification == CanonicalClaimsMappingClassification.Normalized, WarningCode = page.HasValue ? "" : "canada_life_evidence_coordinate_incomplete", WarningMessage = page.HasValue ? "" : "The existing parser did not retain a page coordinate." };
    }

    private static void WarnIfUnmapped(CanonicalClaimsImportEnvelope envelope, string field, CanadaLifeTermMapping mapping, string source, string? factId)
    { if (!mapping.IsMapped) Warning(envelope, $"canada_life_unmapped_{field}", "Warning", field, $"Canada Life term '{mapping.RawValue}' was not mapped and requires review.", mapping.Classification, source, factId, mapping.RuleName); }
    private static void WarnIfUnavailableOrAmbiguous(CanonicalClaimsImportEnvelope envelope, string field, CanadaLifeTermMapping mapping, string source, string factId)
    { if (!mapping.IsMapped) Warning(envelope, $"canada_life_{mapping.Classification.ToString().ToLowerInvariant()}_{field}", "Warning", field, $"Canada Life {field} '{mapping.RawValue}' is {mapping.Classification.ToString().ToLowerInvariant()} and requires review.", mapping.Classification, source, factId, mapping.RuleName); }
    private static void WarnNegative(CanonicalClaimsImportEnvelope envelope, string field, decimal? value, string source, string factId, bool explicitlyAdjustment = false)
    { if (value < 0m) Warning(envelope, $"canada_life_negative_{field}", !explicitlyAdjustment && field == "premium" ? "Error" : "Warning", field, explicitlyAdjustment ? $"Negative {field.Replace('_', ' ')} was preserved as an explicitly labelled adjustment or reversal." : $"Negative {field.Replace('_', ' ')} was preserved but requires adjustment or reversal confirmation.", explicitlyAdjustment ? CanonicalClaimsMappingClassification.Direct : CanonicalClaimsMappingClassification.Ambiguous, source, factId); }
    private static void WarnMissing(CanonicalClaimsImportEnvelope envelope, string field, decimal? value, string source, string factId)
    { if (!value.HasValue) Warning(envelope, $"canada_life_{field}_unavailable", "Warning", field, $"{field.Replace('_', ' ')} is unavailable and was not defaulted to zero.", CanonicalClaimsMappingClassification.Unavailable, source, factId); }

    private static void Warning(CanonicalClaimsImportEnvelope envelope, string code, string severity, string field, string message, CanonicalClaimsMappingClassification classification, string source = "", string? factId = null, string rule = "") => envelope.MappingWarnings.Add(new CanonicalClaimsValidationResult { Code = code, Severity = severity, FieldName = field, Message = message, FactId = factId, MappingClassification = classification, SourceReference = source, NormalizationRule = rule });
    private static string NormalizeCarrier(params string[] values)
    {
        var raw = values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value))?.Trim() ?? "";
        return raw.Equals("Canada Life", StringComparison.OrdinalIgnoreCase)
            || raw.Equals("Great-West Life", StringComparison.OrdinalIgnoreCase)
            || raw.Equals("The Great-West Life Assurance Company", StringComparison.OrdinalIgnoreCase)
                ? "Canada Life"
                : raw;
    }
    private static string SingleDimension(IEnumerable<string> values) { var distinct = values.Where(value => !string.IsNullOrWhiteSpace(value) && !value.Equals("All", StringComparison.OrdinalIgnoreCase)).Distinct(StringComparer.OrdinalIgnoreCase).ToList(); return distinct.Count == 1 ? distinct[0] : ""; }
    private static int? Page(string source) { var match = PageRegex.Match(source); return match.Success && int.TryParse(match.Groups["page"].Value, out var page) ? page : null; }
    private static CanonicalDataAvailability Availability(params decimal?[] values) { if (values.All(value => !value.HasValue)) return CanonicalDataAvailability.Missing; return values.Where(value => value.HasValue).All(value => value == 0m) ? CanonicalDataAvailability.ReportedZero : CanonicalDataAvailability.Present; }
    private static bool IsAdjustment(CanadaLifeTermMapping mapping) => mapping.CanonicalValue is "Adjustment" or "Reversal";
    private static CanadaLifeTermMapping Combined(params CanadaLifeTermMapping[] mappings)
    {
        var classification = mappings.Any(mapping => mapping.Classification == CanonicalClaimsMappingClassification.Ambiguous)
            ? CanonicalClaimsMappingClassification.Ambiguous
            : mappings.Any(mapping => mapping.Classification == CanonicalClaimsMappingClassification.Unavailable)
                ? CanonicalClaimsMappingClassification.Unavailable
                : mappings.Any(mapping => mapping.Classification == CanonicalClaimsMappingClassification.Normalized)
                    ? CanonicalClaimsMappingClassification.Normalized
                    : CanonicalClaimsMappingClassification.Direct;
        return new CanadaLifeTermMapping(
            string.Join("|", mappings.Select(mapping => mapping.RawValue)),
            string.Join("|", mappings.Select(mapping => mapping.CanonicalValue)),
            classification,
            string.Join(";", mappings.Select(mapping => mapping.RuleName).Distinct(StringComparer.Ordinal)));
    }
    private static string Raw(params object?[] values) => string.Join("|", values.Select(value => value is IFormattable formattable ? formattable.ToString(null, CultureInfo.InvariantCulture) : value?.ToString() ?? ""));
    private static string StableId(params object?[] values) => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(string.Join("|", new object?[] { CurrentMappingVersion }.Concat(values).Select(value => value?.ToString()?.Trim().ToUpperInvariant() ?? ""))))).ToLowerInvariant();
}

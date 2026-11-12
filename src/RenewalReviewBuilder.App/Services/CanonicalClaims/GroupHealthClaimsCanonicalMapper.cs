using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services.CanonicalClaims;

/// <summary>Maps GroupHEALTH parser output into the carrier-neutral claims write contract.</summary>
public sealed class GroupHealthClaimsCanonicalMapper : ICarrierClaimsCanonicalMapper<CarrierPdfParseResult>
{
    public const string CurrentMappingVersion = "GROUPHEALTH_CLAIMS_CANONICAL_V1";
    private static readonly Regex PageRegex = new(@"\bpage\s+(?<page>\d+)\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private readonly string _carrierName;
    private readonly string _sourceParserName;
    private readonly string _mappingVersion;

    public GroupHealthClaimsCanonicalMapper()
        : this("GroupHEALTH", nameof(GroupHealthPdfParser), CurrentMappingVersion)
    {
    }

    public GroupHealthClaimsCanonicalMapper(string carrierName, string sourceParserName, string mappingVersion)
    {
        _carrierName = carrierName;
        _sourceParserName = sourceParserName;
        _mappingVersion = mappingVersion;
    }

    public string CarrierName => _carrierName;
    public string SourceParserName => _sourceParserName;

    public CanonicalClaimsImportEnvelope Map(CarrierPdfParseResult parsedResult, long clientId, long? sourceDocumentId, long? importBatchId)
    {
        ArgumentNullException.ThrowIfNull(parsedResult);
        var pdf = parsedResult.Pdf ?? new PdfExtractionResult();
        var workbook = parsedResult.WorkbookFromPdf ?? new WorkbookExtractionResult();
        var envelope = new CanonicalClaimsImportEnvelope
        {
            MappingVersion = _mappingVersion,
            Report = new CanonicalClaimsReport
            {
                ClientId = clientId,
                Carrier = _carrierName,
                PolicyNumber = pdf.PolicyNumber.Trim(),
                SourceDocumentId = sourceDocumentId,
                ImportBatchId = importBatchId,
                Period = ResolvePeriod(pdf.ExperienceLines, workbook.ClaimTrends)
            }
        };

        foreach (var row in pdf.ExperienceLines)
        {
            var benefit = NormalizeBenefit(row.Benefit);
            envelope.FinancialFacts.Add(new CanonicalClaimFinancialFact
            {
                FactId = StableId(sourceDocumentId, "financial", row.Benefit, row.Period),
                Benefit = benefit,
                MeasurementType = "ClaimsExperience",
                DimensionType = "ReportingPeriod",
                DimensionValue = row.Period,
                Premium = row.BilledPremium,
                PaidClaims = row.PaidClaims,
                IncurredClaims = row.IncurredClaims,
                LossRatio = row.IncurredLossRatio,
                Lives = row.AverageLives,
                Currency = "CAD",
                UnitOfMeasure = "Canadian dollars",
                Availability = Availability(row.BilledPremium, row.PaidClaims, row.IncurredClaims),
                Evidence = Evidence(sourceDocumentId, importBatchId, row.SourceReference, "Experience Analysis", row.Benefit, benefit, Raw(row.BilledPremium, row.PaidClaims, row.IncurredClaims, row.IncurredLossRatio))
            });

            AddPercent(envelope, sourceDocumentId, importBatchId, row, benefit, "PaidLossRatio", row.PaidLossRatio, CanonicalClaimsMappingClassification.Derived, "PaidClaims / BilledPremium");
            AddPercent(envelope, sourceDocumentId, importBatchId, row, benefit, "AnnualTrendFactor", row.TrendFactor);
            AddPercent(envelope, sourceDocumentId, importBatchId, row, benefit, "TargetLossRatio", row.TargetLossRatio);
            AddPercent(envelope, sourceDocumentId, importBatchId, row, benefit, "Credibility", row.Credibility);
            AddPercent(envelope, sourceDocumentId, importBatchId, row, benefit, "Weighting", row.Weighting);
            AddPercent(envelope, sourceDocumentId, importBatchId, row, benefit, "RequiredRateChange", row.ProposedAdjustment);
        }

        foreach (var row in workbook.ClaimTrends)
        {
            var benefit = NormalizeBenefit(row.Benefit);
            var category = NormalizeCategory(row.Category);
            envelope.CategoryFacts.Add(new CanonicalClaimCategoryFact
            {
                FactId = StableId(sourceDocumentId, "category", row.Benefit, row.Category, row.Period),
                Benefit = benefit,
                Category = category,
                MeasurementType = "ClaimsCategory",
                DimensionType = "ReportingPeriod",
                DimensionValue = row.Period,
                MonetaryValue = row.Amount,
                PercentageValue = row.PercentOfTotal,
                Currency = "CAD",
                UnitOfMeasure = "Canadian dollars",
                Availability = Availability(row.Amount),
                Evidence = Evidence(sourceDocumentId, importBatchId, row.SourceReference, "Claims Detail", $"{row.Benefit}|{row.Category}", $"{benefit}|{category}", Raw(row.Amount, row.PercentOfTotal))
            });
        }

        foreach (var row in pdf.LifeDemographics.Concat(pdf.LtdDemographics))
        {
            AddLives(envelope, sourceDocumentId, importBatchId, row, "Male", row.CurrentMaleLives);
            AddLives(envelope, sourceDocumentId, importBatchId, row, "Female", row.CurrentFemaleLives);
        }

        AddPooling(envelope, sourceDocumentId, importBatchId, pdf, "PoolingThreshold", pdf.PoolingThreshold, "Currency amount");
        AddPooling(envelope, sourceDocumentId, importBatchId, pdf, "PooledPremium", pdf.PooledPremium, "Currency amount");
        AddPooling(envelope, sourceDocumentId, importBatchId, pdf, "PoolingCharge", pdf.PoolingCharge, "Percent");

        foreach (var drug in workbook.DrugTrends.Concat(workbook.PriorDrugTrends))
        {
            envelope.DrugClaims.Add(new CanonicalDrugClaim
            {
                FactId = StableId(sourceDocumentId, "drug", drug.Din, drug.DrugName, drug.Period),
                Benefit = "Drugs",
                Category = drug.Din,
                DrugName = drug.DrugName,
                MeasurementType = "TopDrugPaid",
                DimensionType = "ReportingPeriod",
                DimensionValue = drug.Period,
                MonetaryValue = drug.NetAmountPaid,
                CountValue = drug.PrescriptionCount,
                Currency = "CAD",
                UnitOfMeasure = "Canadian dollars",
                Availability = Availability(drug.NetAmountPaid),
                Evidence = Evidence(sourceDocumentId, importBatchId, drug.SourceReference, "Top Drugs", drug.Din, drug.DrugName, Raw(drug.NetAmountPaid, drug.PrescriptionCount))
            });
        }

        return envelope;
    }

    private static CanonicalClaimsPeriod ResolvePeriod(IEnumerable<ExperienceLine> experience, IEnumerable<ClaimTrend> categories)
    {
        var periods = experience.Select(row => row.Period).Concat(categories.Select(row => row.Period))
            .Select(label => CanadaLifeClaimsCanonicalMapper.TryParsePeriod(label, out var start, out var end) ? (Valid: true, Start: start, End: end) : default)
            .Where(period => period.Valid)
            .OrderBy(period => period.End)
            .ToList();
        var latest = periods.LastOrDefault();
        return latest.Valid
            ? new CanonicalClaimsPeriod { Start = latest.Start, End = latest.End, PeriodType = "Experience" }
            : new CanonicalClaimsPeriod();
    }

    private void AddPercent(CanonicalClaimsImportEnvelope envelope, long? documentId, long? batchId, ExperienceLine row, string benefit, string measurement, decimal? value, CanonicalClaimsMappingClassification classification = CanonicalClaimsMappingClassification.Direct, string formula = "")
    {
        if (!value.HasValue) return;
        envelope.BenchmarkFacts.Add(new CanonicalBenchmarkFact
        {
            FactId = StableId(documentId, "benchmark", measurement, row.Benefit, row.Period),
            Benefit = benefit,
            MeasurementType = measurement,
            DimensionType = "ReportingPeriod",
            DimensionValue = row.Period,
            PercentageValue = value,
            UnitOfMeasure = "Percent",
            Availability = Availability(value),
            BenchmarkDefinition = $"{_carrierName} {measurement}",
            Evidence = Evidence(documentId, batchId, row.SourceReference, "Renewal Factors", $"{row.Benefit}|{measurement}", $"{benefit}|{measurement}", Raw(value), classification, formula)
        });
    }

    private void AddLives(CanonicalClaimsImportEnvelope envelope, long? documentId, long? batchId, DemographicLine row, string sex, int? lives)
    {
        if (!lives.HasValue) return;
        var benefit = NormalizeBenefit(row.Benefit);
        envelope.CategoryFacts.Add(new CanonicalClaimCategoryFact
        {
            FactId = StableId(documentId, "membership", row.Benefit, row.AgeBand, sex, "Current"),
            Benefit = benefit,
            Category = row.AgeBand,
            MeasurementType = "CoveredLives",
            DimensionType = "Sex",
            DimensionValue = sex,
            CountValue = lives,
            UnitOfMeasure = "Lives",
            Availability = lives == 0 ? CanonicalDataAvailability.ReportedZero : CanonicalDataAvailability.Present,
            Evidence = Evidence(documentId, batchId, row.SourceReference, "Demographics", $"{row.Benefit}|{row.AgeBand}|{sex}", $"{benefit}|{row.AgeBand}|{sex}", lives.Value.ToString(CultureInfo.InvariantCulture))
        });
    }

    private void AddPooling(CanonicalClaimsImportEnvelope envelope, long? documentId, long? batchId, PdfExtractionResult pdf, string measurement, decimal? value, string unit)
    {
        if (!value.HasValue) return;
        envelope.BenchmarkFacts.Add(new CanonicalBenchmarkFact
        {
            FactId = StableId(documentId, "pooling", measurement),
            MeasurementType = measurement,
            MonetaryValue = unit == "Currency amount" ? value : null,
            PercentageValue = unit == "Percent" ? value : null,
            Currency = unit == "Currency amount" ? "CAD" : "",
            UnitOfMeasure = unit,
            Availability = Availability(value),
            BenchmarkDefinition = $"{_carrierName} {measurement}",
            Evidence = Evidence(documentId, batchId, pdf.SourcePath, "Pooling", measurement, measurement, Raw(value))
        });
    }

    private CanonicalClaimsSourceEvidence Evidence(long? documentId, long? batchId, string source, string section, string rawLabel, string normalizedLabel, string rawValue, CanonicalClaimsMappingClassification classification = CanonicalClaimsMappingClassification.Direct, string formula = "")
    {
        var pageMatch = PageRegex.Match(source ?? "");
        var page = pageMatch.Success && int.TryParse(pageMatch.Groups["page"].Value, out var parsedPage) ? parsedPage : (int?)null;
        return new CanonicalClaimsSourceEvidence
        {
            SourceDocumentId = documentId,
            ImportBatchId = batchId,
            SourcePageNumber = page,
            SourceTableOrSection = section,
            RawCarrierLabel = rawLabel,
            NormalizedLabel = normalizedLabel,
            RawExtractedValue = rawValue,
            ExtractionMethod = $"Existing {_carrierName} parser output",
            ParserName = _sourceParserName,
            ParserVersion = "1.0",
            MappingVersion = _mappingVersion,
            Precision = page.HasValue ? CanonicalClaimsEvidencePrecision.Page : CanonicalClaimsEvidencePrecision.Document,
            MappingClassification = classification,
            NormalizationRule = $"{_carrierName} shared claims normalization",
            DerivationFormula = formula,
            WasNormalized = !rawLabel.Equals(normalizedLabel, StringComparison.OrdinalIgnoreCase)
        };
    }

    private static string NormalizeBenefit(string value) => value.Trim() switch
    {
        var item when item.Equals("Extended Health Care", StringComparison.OrdinalIgnoreCase) || item.Equals("Health", StringComparison.OrdinalIgnoreCase) => "EHC",
        var item when item.Equals("Dental Care", StringComparison.OrdinalIgnoreCase) => "Dental",
        var item when item.Equals("Long-Term Disability", StringComparison.OrdinalIgnoreCase) => "LTD",
        var item when item.Equals("Life", StringComparison.OrdinalIgnoreCase) => "Basic Life",
        var item => item
    };

    private static string NormalizeCategory(string value) => value.Trim() switch
    {
        var item when item.Equals("Paramedicals", StringComparison.OrdinalIgnoreCase) => "Paramedical",
        var item when item.Equals("Medical Appliances / Supplies", StringComparison.OrdinalIgnoreCase) || item.Equals("Medical Services and Supplies", StringComparison.OrdinalIgnoreCase) => "Medical Supplies",
        var item when item.Equals("Vision Care", StringComparison.OrdinalIgnoreCase) => "Vision",
        var item => item
    };

    private static CanonicalDataAvailability Availability(params decimal?[] values) => values.All(value => !value.HasValue)
        ? CanonicalDataAvailability.Missing
        : values.Where(value => value.HasValue).All(value => value == 0m) ? CanonicalDataAvailability.ReportedZero : CanonicalDataAvailability.Present;
    private static string Raw(params object?[] values) => string.Join("|", values.Select(value => value is IFormattable formattable ? formattable.ToString(null, CultureInfo.InvariantCulture) : value?.ToString() ?? ""));
    private string StableId(params object?[] values) => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(string.Join("|", new object?[] { _mappingVersion }.Concat(values).Select(value => value?.ToString()?.Trim().ToUpperInvariant() ?? ""))))).ToLowerInvariant();
}

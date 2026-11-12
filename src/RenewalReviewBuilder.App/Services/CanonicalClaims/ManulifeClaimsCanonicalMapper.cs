using System.Security.Cryptography;
using System.Text;
using RenewalReviewBuilder.App.Models;
using ClosedXML.Excel;
using System.Globalization;

namespace RenewalReviewBuilder.App.Services.CanonicalClaims;

/// <summary>
/// Phase 1 proof-of-concept mapper for the existing Manulife-compatible claims result.
/// It maps already-extracted values and does not perform new document extraction.
/// </summary>
public sealed class ManulifeClaimsCanonicalMapper : ICarrierClaimsCanonicalMapper<ClaimsExperienceDataModelResult>
{
    public string CarrierName => "Manulife";
    public string SourceParserName => nameof(ClaimsExperienceDataModelBuilder);

    public CanonicalClaimsImportEnvelope Map(
        ClaimsExperienceDataModelResult legacy,
        long clientId,
        long? sourceDocumentId,
        long? importBatchId)
    {
        ArgumentNullException.ThrowIfNull(legacy);
        var envelope = new CanonicalClaimsImportEnvelope
        {
            MappingVersion = "MANULIFE_CLAIMS_CANONICAL_V1",
            Report = new CanonicalClaimsReport
            {
                ClientId = clientId,
                Carrier = string.IsNullOrWhiteSpace(legacy.Carrier) ? "Manulife" : legacy.Carrier.Trim(),
                PolicyNumber = legacy.ClientNumber.Trim(),
                ImportBatchId = importBatchId,
                SourceDocumentId = sourceDocumentId,
                Period = new CanonicalClaimsPeriod
                {
                    Start = legacy.ExperiencePeriodStart,
                    End = legacy.ExperiencePeriodEnd,
                    PeriodType = "Experience"
                }
            }
        };

        foreach (var summary in legacy.Summaries)
        {
            var evidence = Evidence(summary.BenefitLine, sourceDocumentId, importBatchId, envelope.MappingVersion);
            envelope.FinancialFacts.Add(new CanonicalClaimFinancialFact
            {
                FactId = StableId("financial", clientId, summary.BenefitLine, legacy.ExperiencePeriodStart, legacy.ExperiencePeriodEnd),
                Benefit = NormalizeBenefit(summary.BenefitLine),
                MeasurementType = "ClaimsExperienceSummary",
                Premium = summary.Premium,
                PaidClaims = summary.PaidClaims,
                IncurredClaims = summary.IncurredClaims,
                LossRatio = summary.LossRatio,
                Lives = summary.Lives,
                Availability = Availability(summary.PaidClaims, summary.IncurredClaims, summary.Premium),
                Evidence = evidence
            });
        }

        foreach (var monthly in legacy.MonthlyClaims)
        {
            var evidence = Evidence(monthly.BenefitLine, sourceDocumentId, importBatchId, envelope.MappingVersion);
            envelope.MonthlyFacts.Add(new CanonicalClaimMonthlyFact
            {
                FactId = StableId("monthly", clientId, monthly.BenefitLine, monthly.Month, legacy.ExperiencePeriodEnd),
                Benefit = NormalizeBenefit(monthly.BenefitLine),
                MeasurementType = "MonthlyClaims",
                DimensionType = "Month",
                DimensionValue = monthly.Month?.ToString("yyyy-MM") ?? "",
                Month = monthly.Month,
                Premium = monthly.Premium,
                PaidClaims = monthly.PaidClaims,
                IncurredClaims = monthly.IncurredClaims,
                ClaimsPlusFees = monthly.ClaimsPlusFees,
                LossRatio = monthly.LossRatio,
                Lives = monthly.Lives,
                Availability = Availability(monthly.PaidClaims, monthly.IncurredClaims, monthly.Premium),
                Evidence = evidence
            });
        }

        return envelope;
    }

    /// <summary>Maps Manulife's service-detail-by-relationship workbook without changing the legacy claims path.</summary>
    public CanonicalClaimsImportEnvelope MapServiceDetailWorkbook(string workbookPath, long clientId, long? sourceDocumentId, long? importBatchId)
    {
        using var workbook = new XLWorkbook(workbookPath);
        var period = ReadPeriod(workbook.Worksheet("Additional information"));
        var envelope = new CanonicalClaimsImportEnvelope
        {
            MappingVersion = "MANULIFE_SERVICE_DETAIL_RELATIONSHIP_V1",
            Report = new CanonicalClaimsReport { ClientId=clientId, Carrier="Manulife", PolicyNumber=ReadContract(workbook.Worksheet("Additional information")), SourceDocumentId=sourceDocumentId, ImportBatchId=importBatchId, Period=period }
        };
        MapRelationshipSheet(workbook.Worksheet("Sponsor Summary"), 9, 1, 2, 3, 4, null, null, CanonicalClaimsSourceAggregationLevel.Sponsor, true, envelope, sourceDocumentId, importBatchId, period);
        MapRelationshipSheet(workbook.Worksheet("Contract Summary"), 9, 2, 3, 4, 5, 1, null, CanonicalClaimsSourceAggregationLevel.Contract, false, envelope, sourceDocumentId, importBatchId, period);
        MapRelationshipSheet(workbook.Worksheet("Details"), 9, 4, 5, 6, 7, 1, (2,3), CanonicalClaimsSourceAggregationLevel.Class, false, envelope, sourceDocumentId, importBatchId, period);
        ValidateUtilization(envelope);
        return envelope;
    }

    private static void MapRelationshipSheet(IXLWorksheet sheet, int firstRow, int benefitCol, int categoryCol, int serviceCol, int provinceCol, int? contractCol, (int Division, int Class)? divisionClass, CanonicalClaimsSourceAggregationLevel level, bool hasClaimants, CanonicalClaimsImportEnvelope envelope, long? documentId, long? batchId, CanonicalClaimsPeriod period)
    {
        var last = sheet.LastRowUsed()?.RowNumber() ?? firstRow;
        var groups = hasClaimants
            ? new[] { (CanonicalClaimsRelationshipType.Dependant,5), (CanonicalClaimsRelationshipType.Employee,10), (CanonicalClaimsRelationshipType.Partner,15), (CanonicalClaimsRelationshipType.Total,20) }
            : new[] { (CanonicalClaimsRelationshipType.Dependant,6), (CanonicalClaimsRelationshipType.Employee,10), (CanonicalClaimsRelationshipType.Partner,14), (CanonicalClaimsRelationshipType.Total,18) };
        for (var row=firstRow; row<=last; row++)
        {
            var benefit=Text(sheet.Cell(row,benefitCol)); var category=Text(sheet.Cell(row,categoryCol)); var service=Text(sheet.Cell(row,serviceCol));
            if (string.IsNullOrWhiteSpace(benefit) || string.IsNullOrWhiteSpace(service) || service.Contains("Subtotal",StringComparison.OrdinalIgnoreCase)) continue;
            foreach (var (relationship, start) in groups)
            {
                long? claimants = hasClaimants ? Whole(sheet.Cell(row,start)) : null;
                var occurrence = Whole(sheet.Cell(row,start + (hasClaimants ? 1 : 0)));
                var submitted = Money(sheet.Cell(row,start + (hasClaimants ? 2 : 1)));
                var eligible = Money(sheet.Cell(row,start + (hasClaimants ? 3 : 2)));
                var paid = Money(sheet.Cell(row,start + (hasClaimants ? 4 : 3)));
                if (!claimants.HasValue && !occurrence.HasValue && !submitted.HasValue && !eligible.HasValue && !paid.HasValue) continue;
                var cells = hasClaimants ? $"{Column(start)}{row}:{Column(start+4)}{row}" : $"{Column(start)}{row}:{Column(start+3)}{row}";
                envelope.UtilizationFacts.Add(new CanonicalClaimsUtilizationFact
                {
                    CanonicalClaimsUtilizationFactId=StableUtilizationId(sheet.Name,row,relationship), BenefitCode=Code(benefit), BenefitName=benefit, CategoryCode=Code(category), CategoryName=category, ServiceCode=Code(service), ServiceName=service, ProvinceCode=Code(Text(sheet.Cell(row,provinceCol))), RelationshipType=relationship,
                    DivisionCode=divisionClass is null ? null : Text(sheet.Cell(row,divisionClass.Value.Division)), ClassCode=divisionClass is null ? null : Text(sheet.Cell(row,divisionClass.Value.Class)), UniqueClaimants=claimants, Occurrences=occurrence, SubmittedAmount=submitted, EligibleAmount=eligible, PaidAmount=paid, MeasurementStartDate=period.Start, MeasurementEndDate=period.End, SourceAggregationLevel=level, MappingVersion=envelope.MappingVersion,
                    Evidence=new CanonicalClaimsSourceEvidence { SourceDocumentId=documentId, ImportBatchId=batchId, WorksheetName=sheet.Name, SourceRow=row, SourceColumnOrCell=cells, SourceTableOrSection=sheet.Name, RawCarrierLabel=$"{benefit}|{category}|{service}|{relationship}", NormalizedLabel=$"{benefit}|{category}|{service}", ExtractionMethod="Manulife Claims Service Detail by Relationship workbook", ParserName=nameof(ManulifeClaimsCanonicalMapper), MappingVersion=envelope.MappingVersion, Precision=CanonicalClaimsEvidencePrecision.Cell, ConfidenceScore=1m }
                });
            }
        }
    }

    private static void ValidateUtilization(CanonicalClaimsImportEnvelope envelope)
    {
        foreach (var fact in envelope.UtilizationFacts)
        {
            void Warn(string code, string field, string message) => envelope.MappingWarnings.Add(new CanonicalClaimsValidationResult { Code=code, Severity="Warning", FieldName=field, Message=message, FactId=fact.CanonicalClaimsUtilizationFactId, SourceReference=$"{fact.Evidence.WorksheetName}!{fact.Evidence.SourceColumnOrCell}" });
            if (fact.Occurrences < 0) Warn("utilization_negative_occurrences","occurrences","Occurrences must be nonnegative.");
            if (fact.UniqueClaimants < 0) Warn("utilization_negative_unique_claimants","unique_claimants","Unique claimants must be nonnegative.");
            if (fact.Occurrences > 0 && fact.UniqueClaimants > fact.Occurrences) Warn("utilization_claimants_exceed_occurrences","unique_claimants","Unique claimants normally should not exceed occurrences.");
            if (fact.PaidAmount > fact.EligibleAmount) Warn("utilization_paid_exceeds_eligible","paid_amount","Paid amount exceeds eligible amount and was preserved for review.");
            if (fact.EligibleAmount > fact.SubmittedAmount) Warn("utilization_eligible_exceeds_submitted","eligible_amount","Eligible amount exceeds submitted amount and was preserved for review.");
            if (fact.Occurrences == 0 && fact.SubmittedAmount > 0) Warn("utilization_submitted_without_occurrences","submitted_amount","Positive submitted amount with zero occurrences was preserved.");
            if (fact.EligibleAmount == 0 && fact.PaidAmount > 0) Warn("utilization_paid_without_eligible","paid_amount","Paid amount with zero eligible amount was preserved.");
        }
    }
    private static CanonicalClaimsPeriod ReadPeriod(IXLWorksheet sheet) { var text=Text(sheet.Cell(5,1)); var prefix="Reporting period:"; var values=text.StartsWith(prefix,StringComparison.OrdinalIgnoreCase)?text[prefix.Length..].Trim().Split(" to ",2,StringSplitOptions.TrimEntries):[]; return values.Length==2 && DateOnly.TryParse(values[0],CultureInfo.InvariantCulture,DateTimeStyles.AllowWhiteSpaces,out var start) && DateOnly.TryParse(values[1],CultureInfo.InvariantCulture,DateTimeStyles.AllowWhiteSpaces,out var end) ? new CanonicalClaimsPeriod { Start=start, End=end, PeriodType="Experience" } : new CanonicalClaimsPeriod { PeriodType="Experience" }; }
    private static string ReadContract(IXLWorksheet sheet) => Text(sheet.Cell(9,1));
    private static string Text(IXLCell cell) => cell.GetFormattedString().Trim();
    private static decimal? Money(IXLCell cell) => decimal.TryParse(Text(cell),NumberStyles.Any,CultureInfo.InvariantCulture,out var value) ? value : null;
    private static long? Whole(IXLCell cell) => long.TryParse(Text(cell),NumberStyles.Any,CultureInfo.InvariantCulture,out var value) ? value : null;
    private static string Code(string value) => string.Concat(value.Where(char.IsLetterOrDigit)).ToUpperInvariant();
    private static string Column(int index) { var value=""; while(index>0){ index--; value=(char)('A'+index%26)+value; index/=26; } return value; }
    private static string StableUtilizationId(string sheet, int row, CanonicalClaimsRelationshipType relationship) => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes($"MANULIFE_SERVICE_DETAIL_RELATIONSHIP_V1|{sheet}|{row}|{relationship}"))).ToLowerInvariant();

    private static CanonicalClaimsSourceEvidence Evidence(string rawLabel, long? documentId, long? batchId, string mappingVersion)
    {
        return new CanonicalClaimsSourceEvidence
        {
            SourceDocumentId = documentId,
            ImportBatchId = batchId,
            RawCarrierLabel = rawLabel,
            NormalizedLabel = NormalizeBenefit(rawLabel),
            ExtractionMethod = "Existing claims parser compatibility mapping",
            ParserName = nameof(ClaimsExperienceDataModelBuilder),
            ParserVersion = "Legacy-compatible-v1",
            MappingVersion = mappingVersion,
            Precision = CanonicalClaimsEvidencePrecision.Document
        };
    }

    private static CanonicalDataAvailability Availability(params decimal?[] values)
    {
        if (values.All(value => !value.HasValue)) return CanonicalDataAvailability.Missing;
        if (values.Where(value => value.HasValue).All(value => value == 0m)) return CanonicalDataAvailability.ReportedZero;
        return CanonicalDataAvailability.Present;
    }

    private static string NormalizeBenefit(string value)
    {
        var normalized = value.Trim();
        return normalized.Equals("Extended Health Care", StringComparison.OrdinalIgnoreCase) ? "EHC" :
               normalized.Equals("Health", StringComparison.OrdinalIgnoreCase) ? "EHC" :
               normalized.Equals("Dental Care", StringComparison.OrdinalIgnoreCase) ? "Dental" : normalized;
    }

    private static string StableId(string factType, long clientId, string benefit, DateOnly? start, DateOnly? end)
    {
        var raw = $"{factType}|{clientId}|{NormalizeBenefit(benefit).ToUpperInvariant()}|{start:yyyy-MM-dd}|{end:yyyy-MM-dd}";
        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(raw))).ToLowerInvariant();
    }
}

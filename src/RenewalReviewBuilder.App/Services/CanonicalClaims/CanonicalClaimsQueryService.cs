using System.Globalization;
using Microsoft.Data.Sqlite;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services.CanonicalClaims;

/// <summary>Read-only boundary for downstream consumers of canonical claims.</summary>
public sealed class CanonicalClaimsQueryService
{
    private readonly string _databasePath;

    public CanonicalClaimsQueryService(string? databasePath = null)
    {
        var database = new DatabaseService(databasePath);
        var initialized = database.Initialize();
        if (!initialized.Succeeded) throw new InvalidOperationException(initialized.Message);
        _databasePath = database.DatabasePath;
    }

    public async Task<CanonicalClaimsReportReadModel?> GetLatestForClientAsync(long clientId) =>
        (await GetReportsForClientAsync(clientId, latestOnly: true)).FirstOrDefault();

    /// <summary>
    /// Returns saved canonical report periods for History. Monthly facts remain children of their report
    /// and are never promoted into one History event per month.
    /// </summary>
    public Task<IReadOnlyList<CanonicalClaimsReportReadModel>> GetAllForClientAsync(long clientId) =>
        GetReportsForClientAsync(clientId, latestOnly: false);

    private async Task<IReadOnlyList<CanonicalClaimsReportReadModel>> GetReportsForClientAsync(long clientId, bool latestOnly)
    {
        await using var connection = new SqliteConnection($"Data Source={_databasePath}");
        await connection.OpenAsync();
        await using var reportCommand = connection.CreateCommand();
        reportCommand.CommandText = $@"
SELECT r.canonical_claims_report_id, r.client_id, r.carrier_name, r.policy_number, r.validation_status,
       r.source_document_id, r.source_document_version_id, r.created_at, r.updated_at,
       p.period_start, p.period_end, p.period_type
FROM CanonicalClaimsReport r
JOIN CanonicalClaimsPeriod p ON p.canonical_claims_report_id = r.canonical_claims_report_id
WHERE r.client_id = $client_id
ORDER BY p.period_end DESC, r.updated_at DESC{(latestOnly ? " LIMIT 1" : "")};";
        reportCommand.Parameters.AddWithValue("$client_id", clientId);

        var reports = new List<CanonicalClaimsReportReadModel>();
        await using (var reader = await reportCommand.ExecuteReaderAsync())
        {
            while (await reader.ReadAsync())
            {
                reports.Add(new CanonicalClaimsReportReadModel
                {
                    CanonicalReportId = reader.GetInt64(0),
                    ClientId = reader.GetInt64(1),
                    Carrier = reader.GetString(2),
                    PolicyNumber = reader.GetString(3),
                    ValidationStatus = reader.GetString(4),
                    SourceDocumentId = reader.IsDBNull(5) ? null : reader.GetInt64(5),
                    SourceDocumentVersionId = reader.IsDBNull(6) ? null : reader.GetInt64(6),
                    CreatedAt = DateTime.Parse(reader.GetString(7), CultureInfo.InvariantCulture).ToUniversalTime(),
                    UpdatedAt = DateTime.Parse(reader.GetString(8), CultureInfo.InvariantCulture).ToUniversalTime(),
                    Period = new CanonicalClaimsPeriod
                    {
                        Start = ReadDate(reader, 9),
                        End = ReadDate(reader, 10),
                        PeriodType = reader.GetString(11)
                    }
                });
            }
        }

        foreach (var model in reports)
        {
            await LoadFinancialAsync(connection, model);
            await LoadMonthlyAsync(connection, model);
            await LoadCategoryAsync(connection, model);
            await LoadDistributionAsync(connection, model);
            await LoadBenchmarkAsync(connection, model);
            await LoadUtilizationAsync(connection, model);
            await LoadParityAsync(connection, model);
        }

        return reports;
    }

    private static async Task LoadFinancialAsync(SqliteConnection connection, CanonicalClaimsReportReadModel model)
    {
        await using var command = connection.CreateCommand();
        command.CommandText = "SELECT fact_id,benefit,coverage,category,division,class,measurement_type,dimension_type,dimension_value,premium,paid_claims,incurred_claims,loss_ratio,lives,currency,unit_of_measure,availability_state,validation_status FROM CanonicalClaimsFinancialFact WHERE canonical_claims_report_id=$id ORDER BY benefit,dimension_value;";
        command.Parameters.AddWithValue("$id", model.CanonicalReportId);
        await using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync()) model.FinancialFacts.Add(new CanonicalClaimFinancialFact { FactId=reader.GetString(0), Benefit=reader.GetString(1), Coverage=reader.GetString(2), Category=reader.GetString(3), Division=reader.GetString(4), Class=reader.GetString(5), MeasurementType=reader.GetString(6), DimensionType=reader.GetString(7), DimensionValue=reader.GetString(8), Premium=ReadDecimal(reader,9), PaidClaims=ReadDecimal(reader,10), IncurredClaims=ReadDecimal(reader,11), LossRatio=ReadDecimal(reader,12), Lives=ReadInt(reader,13), Currency=reader.GetString(14), UnitOfMeasure=reader.GetString(15), Availability=Enum.Parse<CanonicalDataAvailability>(reader.GetString(16)), ValidationStatus=ReadValidationStatus(reader,17) });
    }

    private static async Task LoadMonthlyAsync(SqliteConnection connection, CanonicalClaimsReportReadModel model)
    {
        await using var command = connection.CreateCommand();
        command.CommandText = "SELECT fact_id,benefit,coverage,category,division,class,measurement_type,dimension_type,dimension_value,month,premium,paid_claims,incurred_claims,claims_plus_fees,loss_ratio,lives,currency,unit_of_measure,availability_state,validation_status FROM CanonicalClaimsMonthlyFact WHERE canonical_claims_report_id=$id ORDER BY month,benefit;";
        command.Parameters.AddWithValue("$id", model.CanonicalReportId);
        await using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync()) model.MonthlyFacts.Add(new CanonicalClaimMonthlyFact { FactId=reader.GetString(0), Benefit=reader.GetString(1), Coverage=reader.GetString(2), Category=reader.GetString(3), Division=reader.GetString(4), Class=reader.GetString(5), MeasurementType=reader.GetString(6), DimensionType=reader.GetString(7), DimensionValue=reader.GetString(8), Month=ReadDate(reader,9), Premium=ReadDecimal(reader,10), PaidClaims=ReadDecimal(reader,11), IncurredClaims=ReadDecimal(reader,12), ClaimsPlusFees=ReadDecimal(reader,13), LossRatio=ReadDecimal(reader,14), Lives=ReadInt(reader,15), Currency=reader.GetString(16), UnitOfMeasure=reader.GetString(17), Availability=Enum.Parse<CanonicalDataAvailability>(reader.GetString(18)), ValidationStatus=ReadValidationStatus(reader,19) });
    }

    private static async Task LoadParityAsync(SqliteConnection connection, CanonicalClaimsReportReadModel model)
    {
        await using var command = connection.CreateCommand();
        command.CommandText = "SELECT metric,benefit,period,legacy_value,canonical_value,variance,tolerance,classification,explanation,related_legacy_record_ids,related_canonical_fact_ids,review_status FROM CanonicalClaimsParityResult WHERE canonical_claims_report_id=$id ORDER BY metric,benefit,period;";
        command.Parameters.AddWithValue("$id", model.CanonicalReportId);
        await using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync()) model.ParityResults.Add(new CanonicalClaimsParityResult { Metric=reader.GetString(0), Benefit=reader.GetString(1), Period=reader.GetString(2), LegacyValue=ReadDecimal(reader,3), CanonicalValue=ReadDecimal(reader,4), Variance=ReadDecimal(reader,5), Tolerance=ReadDecimal(reader,6) ?? 0m, Classification=Enum.Parse<CanonicalClaimsParityClassification>(reader.GetString(7)), Explanation=reader.GetString(8), RelatedLegacyRecordIds=reader.GetString(9), RelatedCanonicalFactIds=reader.GetString(10), ReviewStatus=reader.GetString(11) });
    }

    private static async Task LoadCategoryAsync(SqliteConnection connection, CanonicalClaimsReportReadModel model)
    {
        await using var command = connection.CreateCommand();
        command.CommandText = "SELECT fact_id,benefit,coverage,category,division,class,measurement_type,dimension_type,dimension_value,monetary_value,count_value,percentage_value,currency,unit_of_measure,availability_state,validation_status FROM CanonicalClaimsCategoryFact WHERE canonical_claims_report_id=$id ORDER BY benefit,category,dimension_value;";
        command.Parameters.AddWithValue("$id", model.CanonicalReportId);
        await using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync()) model.CategoryFacts.Add(new CanonicalClaimCategoryFact { FactId=reader.GetString(0), Benefit=reader.GetString(1), Coverage=reader.GetString(2), Category=reader.GetString(3), Division=reader.GetString(4), Class=reader.GetString(5), MeasurementType=reader.GetString(6), DimensionType=reader.GetString(7), DimensionValue=reader.GetString(8), MonetaryValue=ReadDecimal(reader,9), CountValue=ReadLong(reader,10), PercentageValue=ReadDecimal(reader,11), Currency=reader.GetString(12), UnitOfMeasure=reader.GetString(13), Availability=Enum.Parse<CanonicalDataAvailability>(reader.GetString(14)), ValidationStatus=ReadValidationStatus(reader,15) });
    }

    private static async Task LoadDistributionAsync(SqliteConnection connection, CanonicalClaimsReportReadModel model)
    {
        await using var command = connection.CreateCommand();
        command.CommandText = "SELECT fact_id,benefit,coverage,category,division,class,measurement_type,dimension_type,dimension_value,monetary_value,count_value,percentage_value,average_amount,currency,unit_of_measure,availability_state,validation_status FROM CanonicalClaimsDistributionFact WHERE canonical_claims_report_id=$id ORDER BY benefit,category,dimension_value;";
        command.Parameters.AddWithValue("$id", model.CanonicalReportId);
        await using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync()) model.DistributionFacts.Add(new CanonicalClaimDistributionFact { FactId=reader.GetString(0), Benefit=reader.GetString(1), Coverage=reader.GetString(2), Category=reader.GetString(3), Division=reader.GetString(4), Class=reader.GetString(5), MeasurementType=reader.GetString(6), DimensionType=reader.GetString(7), DimensionValue=reader.GetString(8), MonetaryValue=ReadDecimal(reader,9), CountValue=ReadLong(reader,10), PercentageValue=ReadDecimal(reader,11), AverageAmount=ReadDecimal(reader,12), Currency=reader.GetString(13), UnitOfMeasure=reader.GetString(14), Availability=Enum.Parse<CanonicalDataAvailability>(reader.GetString(15)), ValidationStatus=ReadValidationStatus(reader,16) });
    }

    private static async Task LoadBenchmarkAsync(SqliteConnection connection, CanonicalClaimsReportReadModel model)
    {
        await using var command = connection.CreateCommand();
        command.CommandText = "SELECT fact_id,benefit,coverage,category,division,class,measurement_type,dimension_type,dimension_value,benchmark_definition,monetary_value,count_value,percentage_value,currency,unit_of_measure,availability_state,validation_status FROM CanonicalClaimsBenchmarkFact WHERE canonical_claims_report_id=$id ORDER BY benefit,category,measurement_type;";
        command.Parameters.AddWithValue("$id", model.CanonicalReportId);
        await using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync()) model.BenchmarkFacts.Add(new CanonicalBenchmarkFact { FactId=reader.GetString(0), Benefit=reader.GetString(1), Coverage=reader.GetString(2), Category=reader.GetString(3), Division=reader.GetString(4), Class=reader.GetString(5), MeasurementType=reader.GetString(6), DimensionType=reader.GetString(7), DimensionValue=reader.GetString(8), BenchmarkDefinition=reader.GetString(9), MonetaryValue=ReadDecimal(reader,10), CountValue=ReadLong(reader,11), PercentageValue=ReadDecimal(reader,12), Currency=reader.GetString(13), UnitOfMeasure=reader.GetString(14), Availability=Enum.Parse<CanonicalDataAvailability>(reader.GetString(15)), ValidationStatus=ReadValidationStatus(reader,16) });
    }

    /// <summary>Returns raw utilization observations. It intentionally does not aggregate unique claimants.</summary>
    public async Task<IReadOnlyList<CanonicalClaimsUtilizationFact>> GetUtilizationFactsAsync(long reportId, string? benefit = null, string? category = null, string? service = null, string? province = null, CanonicalClaimsRelationshipType? relationship = null, string? division = null, string? @class = null, CanonicalClaimsSourceAggregationLevel? aggregationLevel = null)
    {
        await using var connection = new SqliteConnection($"Data Source={_databasePath}");
        await connection.OpenAsync();
        await using var command = connection.CreateCommand();
        command.CommandText = "SELECT fact_id,benefit_code,benefit_name,category_code,category_name,service_code,service_name,province_code,relationship_type,division_code,class_code,unique_claimants,occurrences,submitted_amount,eligible_amount,paid_amount,reported_paid_percentage,currency,validation_state,measurement_start_date,measurement_end_date,source_aggregation_level,mapping_version,record_status,created_at_utc FROM CanonicalClaimsUtilizationFact WHERE canonical_claims_report_id=$report";
        command.Parameters.AddWithValue("$report", reportId);
        AddFilter(command, "benefit_code", benefit, "$benefit"); AddFilter(command, "category_code", category, "$category"); AddFilter(command, "service_code", service, "$service"); AddFilter(command, "province_code", province, "$province"); AddFilter(command, "division_code", division, "$division"); AddFilter(command, "class_code", @class, "$class");
        if (relationship.HasValue) { command.CommandText += " AND relationship_type=$relationship"; command.Parameters.AddWithValue("$relationship", relationship.Value.ToString()); }
        if (aggregationLevel.HasValue) { command.CommandText += " AND source_aggregation_level=$level"; command.Parameters.AddWithValue("$level", aggregationLevel.Value.ToString()); }
        command.CommandText += " ORDER BY benefit_name,category_name,service_name,province_code,relationship_type;";
        var results = new List<CanonicalClaimsUtilizationFact>();
        await using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync()) results.Add(new CanonicalClaimsUtilizationFact { CanonicalClaimsUtilizationFactId=reader.GetString(0), BenefitCode=reader.GetString(1), BenefitName=reader.GetString(2), CategoryCode=reader.GetString(3), CategoryName=reader.GetString(4), ServiceCode=reader.GetString(5), ServiceName=reader.GetString(6), ProvinceCode=reader.GetString(7), RelationshipType=Enum.Parse<CanonicalClaimsRelationshipType>(reader.GetString(8)), DivisionCode=reader.IsDBNull(9)?null:reader.GetString(9), ClassCode=reader.IsDBNull(10)?null:reader.GetString(10), UniqueClaimants=ReadLong(reader,11), Occurrences=ReadLong(reader,12), SubmittedAmount=ReadDecimal(reader,13), EligibleAmount=ReadDecimal(reader,14), PaidAmount=ReadDecimal(reader,15), ReportedPaidPercentage=ReadDecimal(reader,16), Currency=reader.IsDBNull(17)?"CAD":reader.GetString(17), ValidationStatus=Enum.Parse<CanonicalClaimsValidationStatus>(reader.IsDBNull(18)?"Valid":reader.GetString(18)), MeasurementStartDate=ReadDate(reader,19), MeasurementEndDate=ReadDate(reader,20), SourceAggregationLevel=Enum.Parse<CanonicalClaimsSourceAggregationLevel>(reader.GetString(21)), MappingVersion=reader.GetString(22), RecordStatus=reader.GetString(23), CreatedAtUtc=DateTime.Parse(reader.GetString(24), CultureInfo.InvariantCulture) });
        return results;
    }

    private static void AddFilter(SqliteCommand command, string column, string? value, string parameter)
    { if (!string.IsNullOrWhiteSpace(value)) { command.CommandText += $" AND {column}={parameter}"; command.Parameters.AddWithValue(parameter, value); } }

    private static async Task LoadUtilizationAsync(SqliteConnection connection, CanonicalClaimsReportReadModel model)
    {
        await using var command = connection.CreateCommand();
        command.CommandText = "SELECT fact_id,benefit_code,benefit_name,category_code,category_name,service_code,service_name,province_code,relationship_type,division_code,class_code,unique_claimants,occurrences,submitted_amount,eligible_amount,paid_amount,reported_paid_percentage,currency,validation_state,measurement_start_date,measurement_end_date,source_aggregation_level,mapping_version,record_status,created_at_utc FROM CanonicalClaimsUtilizationFact WHERE canonical_claims_report_id=$report ORDER BY benefit_name,category_name,service_name,province_code,relationship_type;";
        command.Parameters.AddWithValue("$report", model.CanonicalReportId);
        await using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync()) model.UtilizationFacts.Add(new CanonicalClaimsUtilizationFact { CanonicalClaimsUtilizationFactId=reader.GetString(0), BenefitCode=reader.GetString(1), BenefitName=reader.GetString(2), CategoryCode=reader.GetString(3), CategoryName=reader.GetString(4), ServiceCode=reader.GetString(5), ServiceName=reader.GetString(6), ProvinceCode=reader.GetString(7), RelationshipType=Enum.Parse<CanonicalClaimsRelationshipType>(reader.GetString(8)), DivisionCode=reader.IsDBNull(9)?null:reader.GetString(9), ClassCode=reader.IsDBNull(10)?null:reader.GetString(10), UniqueClaimants=ReadLong(reader,11), Occurrences=ReadLong(reader,12), SubmittedAmount=ReadDecimal(reader,13), EligibleAmount=ReadDecimal(reader,14), PaidAmount=ReadDecimal(reader,15), ReportedPaidPercentage=ReadDecimal(reader,16), Currency=reader.IsDBNull(17)?"CAD":reader.GetString(17), ValidationStatus=Enum.Parse<CanonicalClaimsValidationStatus>(reader.IsDBNull(18)?"Valid":reader.GetString(18)), MeasurementStartDate=ReadDate(reader,19), MeasurementEndDate=ReadDate(reader,20), SourceAggregationLevel=Enum.Parse<CanonicalClaimsSourceAggregationLevel>(reader.GetString(21)), MappingVersion=reader.GetString(22), RecordStatus=reader.GetString(23), CreatedAtUtc=DateTime.Parse(reader.GetString(24), CultureInfo.InvariantCulture) });
    }

    private static decimal? ReadDecimal(SqliteDataReader reader, int ordinal) => reader.IsDBNull(ordinal) ? null : Convert.ToDecimal(reader.GetValue(ordinal), CultureInfo.InvariantCulture);
    private static int? ReadInt(SqliteDataReader reader, int ordinal) => reader.IsDBNull(ordinal) ? null : Convert.ToInt32(reader.GetValue(ordinal), CultureInfo.InvariantCulture);
    private static long? ReadLong(SqliteDataReader reader, int ordinal) => reader.IsDBNull(ordinal) ? null : Convert.ToInt64(reader.GetValue(ordinal), CultureInfo.InvariantCulture);
    private static DateOnly? ReadDate(SqliteDataReader reader, int ordinal) => reader.IsDBNull(ordinal) || string.IsNullOrWhiteSpace(reader.GetString(ordinal)) ? null : DateOnly.Parse(reader.GetString(ordinal), CultureInfo.InvariantCulture);
    private static CanonicalClaimsValidationStatus ReadValidationStatus(SqliteDataReader reader, int ordinal)
    {
        if (reader.IsDBNull(ordinal)) return CanonicalClaimsValidationStatus.NeedsReview;
        var value = reader.GetString(ordinal);
        return Enum.TryParse<CanonicalClaimsValidationStatus>(value, ignoreCase: true, out var status) &&
               Enum.IsDefined(status)
            ? status
            : CanonicalClaimsValidationStatus.NeedsReview;
    }
}


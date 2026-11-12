using System.Globalization;
using System.Text.RegularExpressions;
using Microsoft.Data.Sqlite;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed partial class DatabaseService
{
    private static async Task EnsureManulifeExcelBillingSchemaAsync(SqliteConnection connection)
    {
        await EnsureColumnExistsAsync(connection, "manulife_billing_statement_versions", "carrier_id", "INTEGER");
        await EnsureColumnExistsAsync(connection, "manulife_billing_statement_versions", "statement_date", "TEXT");
        await EnsureColumnExistsAsync(connection, "manulife_billing_statement_versions", "source_row_count", "INTEGER NOT NULL DEFAULT 0");
        await EnsureColumnExistsAsync(connection, "manulife_billing_statement_versions", "employee_record_count", "INTEGER NOT NULL DEFAULT 0");
        await EnsureColumnExistsAsync(connection, "manulife_billing_statement_versions", "benefit_record_count", "INTEGER NOT NULL DEFAULT 0");
        await EnsureColumnExistsAsync(connection, "manulife_billing_statement_versions", "total_current_month_premium", "REAL");
        await EnsureColumnExistsAsync(connection, "manulife_billing_statement_versions", "total_prior_period_adjustment", "REAL");
        await EnsureColumnExistsAsync(connection, "manulife_billing_statement_versions", "total_sales_tax", "REAL");
        await EnsureColumnExistsAsync(connection, "manulife_billing_statement_versions", "total_premium_and_tax", "REAL");
        await EnsureColumnExistsAsync(connection, "manulife_billing_statement_versions", "reconciliation_difference", "REAL");
        await EnsureColumnExistsAsync(connection, "manulife_billing_statement_versions", "parser_version", "TEXT NOT NULL DEFAULT ''");
        await EnsureColumnExistsAsync(connection, "manulife_billing_statement_versions", "error_count", "INTEGER NOT NULL DEFAULT 0");
        await EnsureColumnExistsAsync(connection, "manulife_billing_statement_versions", "warning_count", "INTEGER NOT NULL DEFAULT 0");
        await EnsureColumnExistsAsync(connection, "manulife_billing_statement_versions", "header_mapping_json", "TEXT NOT NULL DEFAULT ''");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS manulife_billing_employee_records (
    billing_employee_record_id INTEGER PRIMARY KEY AUTOINCREMENT,
    statement_version_id INTEGER NOT NULL,
    client_id INTEGER NOT NULL,
    carrier_id INTEGER,
    certificate_number TEXT NOT NULL DEFAULT '',
    last_name TEXT NOT NULL DEFAULT '',
    first_name TEXT NOT NULL DEFAULT '',
    class_name TEXT NOT NULL DEFAULT '',
    class_code TEXT NOT NULL DEFAULT '',
    plan_name TEXT NOT NULL DEFAULT '',
    plan_code TEXT NOT NULL DEFAULT '',
    province_of_work TEXT NOT NULL DEFAULT '',
    province_of_work_raw TEXT NOT NULL DEFAULT '',
    province_of_residence TEXT NOT NULL DEFAULT '',
    province_of_residence_raw TEXT NOT NULL DEFAULT '',
    current_month_premium REAL,
    prior_period_adjustment REAL,
    total_sales_tax REAL,
    total_premium_and_tax REAL,
    calculated_current_month_premium REAL,
    calculated_prior_period_adjustment REAL,
    calculated_sales_tax REAL,
    calculated_premium_and_tax REAL,
    reconciliation_difference REAL,
    divisional_billing_code TEXT NOT NULL DEFAULT '',
    bill_total_6 TEXT NOT NULL DEFAULT '',
    bill_total_7 TEXT NOT NULL DEFAULT '',
    source_row_number INTEGER NOT NULL,
    source_row_hash TEXT NOT NULL DEFAULT '',
    validation_status TEXT NOT NULL DEFAULT '',
    FOREIGN KEY (statement_version_id) REFERENCES manulife_billing_statement_versions(statement_version_id),
    FOREIGN KEY (client_id) REFERENCES Clients(client_id),
    FOREIGN KEY (carrier_id) REFERENCES Carrier(carrier_id)
);");
        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_ManulifeBillingEmployee_Statement
ON manulife_billing_employee_records(statement_version_id, certificate_number); ");
        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_ManulifeBillingEmployee_ClientPeriod
ON manulife_billing_employee_records(client_id, certificate_number); ");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS manulife_billing_benefit_records (
    billing_benefit_record_id INTEGER PRIMARY KEY AUTOINCREMENT,
    billing_employee_record_id INTEGER NOT NULL,
    statement_version_id INTEGER NOT NULL,
    client_id INTEGER NOT NULL,
    certificate_number TEXT NOT NULL DEFAULT '',
    benefit_code TEXT NOT NULL DEFAULT '',
    benefit_name TEXT NOT NULL DEFAULT '',
    benefit_plan_code TEXT NOT NULL DEFAULT '',
    dependent_code_or_volume_raw TEXT NOT NULL DEFAULT '',
    dependent_code TEXT NOT NULL DEFAULT '',
    insured_volume REAL,
    option_code TEXT NOT NULL DEFAULT '',
    current_month_premium REAL,
    prior_period_adjustment REAL,
    total_sales_tax REAL,
    benefit_total REAL,
    calculated_benefit_total REAL,
    reconciliation_difference REAL,
    source_row_number INTEGER NOT NULL,
    validation_status TEXT NOT NULL DEFAULT '',
    FOREIGN KEY (billing_employee_record_id) REFERENCES manulife_billing_employee_records(billing_employee_record_id),
    FOREIGN KEY (statement_version_id) REFERENCES manulife_billing_statement_versions(statement_version_id),
    FOREIGN KEY (client_id) REFERENCES Clients(client_id)
);");
        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_ManulifeBillingBenefit_Statement
ON manulife_billing_benefit_records(statement_version_id, certificate_number, benefit_code); ");
        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_ManulifeBillingBenefit_Client
ON manulife_billing_benefit_records(client_id, benefit_code); ");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS manulife_billing_exceptions (
    billing_exception_id INTEGER PRIMARY KEY AUTOINCREMENT,
    statement_version_id INTEGER NOT NULL,
    source_row_number INTEGER,
    certificate_number TEXT NOT NULL DEFAULT '',
    benefit_code TEXT NOT NULL DEFAULT '',
    field_name TEXT NOT NULL DEFAULT '',
    exception_type TEXT NOT NULL DEFAULT '',
    severity TEXT NOT NULL DEFAULT 'Warning',
    source_value TEXT NOT NULL DEFAULT '',
    expected_value TEXT NOT NULL DEFAULT '',
    message TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL,
    resolution_status TEXT NOT NULL DEFAULT 'Open',
    FOREIGN KEY (statement_version_id) REFERENCES manulife_billing_statement_versions(statement_version_id)
);");
        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_ManulifeBillingException_Statement
ON manulife_billing_exceptions(statement_version_id, severity, resolution_status); ");
    }

    public async Task<bool> IsCarrierBillingFileAlreadyImportedAsync(long clientId, string carrier, string sourceFileHash)
    {
        if (clientId <= 0 || string.IsNullOrWhiteSpace(sourceFileHash))
        {
            return false;
        }

        await InitializeCoreAsync();
        using var connection = await OpenConnectionAsync();
        using var command = CreateCommand(connection, null, @"
SELECT COUNT(*)
FROM manulife_billing_statement_versions
WHERE client_id = $client_id
  AND carrier = $carrier COLLATE NOCASE
  AND source_file_hash = $source_file_hash; ");
        AddParameter(command, "$client_id", clientId);
        AddParameter(command, "$carrier", FirstNonBlank(carrier, "Manulife"));
        AddParameter(command, "$source_file_hash", sourceFileHash);
        return Convert.ToInt32(await command.ExecuteScalarAsync(), CultureInfo.InvariantCulture) > 0;
    }

    public async Task<CarrierBillingConfirmationResult> SaveManulifeExcelBillingImportAsync(
        ManulifeExcelBillingImportPreview preview,
        ManulifeExcelBillingImportRequest request)
    {
        var response = new CarrierBillingConfirmationResult { DatabasePath = DatabasePath };
        if (!preview.CanConfirm)
        {
            response.Message = "Billing preview has blocking errors, a duplicate source file, or no selected client. Nothing was saved.";
            return response;
        }

        try
        {
            await InitializeCoreAsync();
            using var connection = await OpenConnectionAsync();
            using var transaction = connection.BeginTransaction();
            var clientName = FirstNonBlank(preview.DetectedClientName, request.SelectedClientName, preview.ClientName);
            var clientId = request.SelectedClientId ??
                           await FindCompatibleManulifeBillingClientIdAsync(connection, transaction, clientName) ??
                           await UpsertClientAsync(connection, transaction, new ClientRecord
                           {
                               ClientName = clientName,
                               Carrier = "Manulife",
                               Status = DefaultClientStatus
                           });
            var sourceFileHash = FirstNonBlank(preview.SourceFileHash, await ComputeFileHashAsync(preview.SourceFilePath));
            var existingVersion = await FindManulifeBillingStatementByHashAsync(connection, transaction, clientId, "Manulife", sourceFileHash);
            if (existingVersion is not null)
            {
                transaction.Rollback();
                response.ClientId = clientId;
                response.Message = "This Manulife billing file has already been imported for the selected client.";
                return response;
            }

            var document = await UpsertDocumentAsync(connection, transaction, clientId, DocumentTypes.Billing, BenefitsIntelligenceModules.BillingPremiumIntelligence, preview.SourceFilePath, "Manulife", null, null, null);
            if (document.DocumentId is null)
            {
                throw new InvalidOperationException("The source billing document could not be recorded.");
            }

            var source = await StartImportTrackingAsync(connection, transaction, document, SourceTypes.Excel, "Manulife", DocumentTypes.Billing, "Manulife Excel/CSV billing parser v1", 0.95m);
            var carrierId = await FindCarrierIdAsync(connection, transaction, "Manulife");
            var versionNumber = await GetNextManulifeBillingVersionNumberAsync(connection, transaction, clientId, "Manulife", preview.BillingPeriod);
            var status = preview.WarningCount > 0 ? ImportStatuses.NeedsReview : ImportStatuses.Completed;
            var versionId = await InsertManulifeExcelBillingVersionAsync(connection, transaction, preview, request, clientId, carrierId, document.DocumentId.Value, source.ImportBatchId, sourceFileHash, versionNumber, status);

            var employeeIds = new Dictionary<int, long>();
            foreach (var employee in preview.EmployeeRecords)
            {
                var employeeId = await InsertManulifeExcelBillingEmployeeAsync(connection, transaction, employee, versionId, clientId, carrierId);
                employeeIds[employee.SourceRowNumber] = employeeId;
            }
            foreach (var benefit in preview.BenefitRecords)
            {
                if (employeeIds.TryGetValue(benefit.SourceRowNumber, out var employeeId))
                {
                    await InsertManulifeExcelBillingBenefitAsync(connection, transaction, benefit, employeeId, versionId, clientId);
                }
            }
            foreach (var exception in preview.Exceptions)
            {
                await InsertManulifeExcelBillingExceptionAsync(connection, transaction, exception, versionId);
                await InsertValidationIssueAsync(connection, transaction, new ValidationIssueRecord
                {
                    ImportBatchId = source.ImportBatchId,
                    SourceDocumentId = document.DocumentId,
                    IssueCode = exception.ExceptionType,
                    Severity = exception.Severity,
                    TableName = "ManulifeBilling",
                    FieldName = exception.FieldName,
                    Message = exception.Message,
                    RawSourceReference = $"row {exception.SourceRowNumber?.ToString(CultureInfo.InvariantCulture) ?? "file"}; {exception.BenefitCode}"
                });
            }
            await InsertImportLogAsync(connection, transaction, source.ImportBatchId!.Value, "Confirm Manulife billing import", status, $"Saved {preview.EmployeeRecordCount} employee records and {preview.BenefitRecordCount} benefit records.");
            await UpdateImportBatchStatusAsync(connection, transaction, source.ImportBatchId.Value, status, $"{preview.ErrorCount} error(s), {preview.WarningCount} warning(s).");
            await UpdateDocumentProcessingStatusAsync(connection, transaction, document.DocumentId.Value, status);
            transaction.Commit();

            response.Succeeded = true;
            response.Message = $"Saved Manulife billing import version {versionNumber} for {preview.BillingPeriod}.";
            response.CarrierBillingImportId = versionId;
            response.ClientId = clientId;
            response.ImportBatchId = source.ImportBatchId;
            response.SavedEmployeeRecordCount = preview.EmployeeRecordCount;
            response.SavedBenefitRecordCount = preview.BenefitRecordCount;
            return response;
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            response.Message = "Manulife Excel billing import could not be saved. The transaction was rolled back.";
            return response;
        }
    }

    public async Task<List<CarrierBillingImportHistoryRecord>> GetManulifeExcelBillingImportHistoryAsync(long clientId)
    {
        await InitializeCoreAsync();
        using var connection = await OpenConnectionAsync();
        using var command = CreateCommand(connection, null, @"
SELECT statement_version_id, client_id, carrier, billing_period, source_file_name, version_number, import_status, created_at,
       employee_record_count, benefit_record_count, total_premium_and_tax, error_count, warning_count
FROM manulife_billing_statement_versions
WHERE client_id = $client_id
  AND carrier = 'Manulife'
ORDER BY billing_period DESC, version_number DESC; ");
        AddParameter(command, "$client_id", clientId);
        var records = new List<CarrierBillingImportHistoryRecord>();
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            records.Add(new CarrierBillingImportHistoryRecord
            {
                CarrierBillingImportId = reader.GetInt64(reader.GetOrdinal("statement_version_id")),
                ClientId = reader.GetInt64(reader.GetOrdinal("client_id")),
                CarrierName = ReadString(reader, "carrier"),
                BillingPeriod = ReadString(reader, "billing_period"),
                SourceFileName = ReadString(reader, "source_file_name"),
                VersionNumber = reader.GetInt32(reader.GetOrdinal("version_number")),
                ImportStatus = ReadString(reader, "import_status"),
                ImportedAt = DateTime.TryParse(ReadString(reader, "created_at"), CultureInfo.InvariantCulture, DateTimeStyles.None, out var importedAt)
                    ? importedAt
                    : DateTime.MinValue,
                EmployeeRecordCount = reader.GetInt32(reader.GetOrdinal("employee_record_count")),
                BenefitRecordCount = reader.GetInt32(reader.GetOrdinal("benefit_record_count")),
                TotalPremiumAndTax = ReadDecimal(reader, "total_premium_and_tax") ?? 0m,
                ErrorCount = reader.GetInt32(reader.GetOrdinal("error_count")),
                WarningCount = reader.GetInt32(reader.GetOrdinal("warning_count"))
            });
        }
        return records;
    }

    public async Task<List<CarrierBillingEmployeeRecord>> GetManulifeExcelBillingEmployeeRecordsAsync(long carrierBillingImportId)
    {
        await InitializeCoreAsync();
        using var connection = await OpenConnectionAsync();
        using var command = CreateCommand(connection, null, @"
SELECT certificate_number, last_name, first_name, class_name, class_code, plan_name, plan_code,
       current_month_premium, prior_period_adjustment, total_sales_tax, total_premium_and_tax,
       reconciliation_difference, divisional_billing_code, validation_status
FROM manulife_billing_employee_records
WHERE statement_version_id = $statement_version_id
ORDER BY class_code, certificate_number;");
        AddParameter(command, "$statement_version_id", carrierBillingImportId);
        var records = new List<CarrierBillingEmployeeRecord>();
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            records.Add(new CarrierBillingEmployeeRecord
            {
                CertificateNumber = ReadString(reader, "certificate_number"),
                LastName = ReadString(reader, "last_name"),
                FirstName = ReadString(reader, "first_name"),
                ClassName = ReadString(reader, "class_name"),
                ClassCode = ReadString(reader, "class_code"),
                PlanName = ReadString(reader, "plan_name"),
                PlanCode = ReadString(reader, "plan_code"),
                CurrentMonthPremium = ReadDecimal(reader, "current_month_premium"),
                PriorPeriodAdjustment = ReadDecimal(reader, "prior_period_adjustment"),
                TotalSalesTax = ReadDecimal(reader, "total_sales_tax"),
                TotalPremiumAndTax = ReadDecimal(reader, "total_premium_and_tax"),
                ReconciliationDifference = ReadDecimal(reader, "reconciliation_difference"),
                DivisionalBillingCode = ReadString(reader, "divisional_billing_code"),
                ValidationStatus = ReadString(reader, "validation_status")
            });
        }
        return records;
    }

    public async Task<List<CarrierBillingException>> GetManulifeExcelBillingExceptionsAsync(long carrierBillingImportId)
    {
        await InitializeCoreAsync();
        using var connection = await OpenConnectionAsync();
        using var command = CreateCommand(connection, null, @"
SELECT source_row_number, certificate_number, benefit_code, field_name, exception_type, severity,
       source_value, expected_value, message, resolution_status
FROM manulife_billing_exceptions
WHERE statement_version_id = $statement_version_id
ORDER BY severity DESC, source_row_number;");
        AddParameter(command, "$statement_version_id", carrierBillingImportId);
        var records = new List<CarrierBillingException>();
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            var sourceRowOrdinal = reader.GetOrdinal("source_row_number");
            records.Add(new CarrierBillingException
            {
                SourceRowNumber = reader.IsDBNull(sourceRowOrdinal) ? null : reader.GetInt32(sourceRowOrdinal),
                CertificateNumber = ReadString(reader, "certificate_number"),
                BenefitCode = ReadString(reader, "benefit_code"),
                FieldName = ReadString(reader, "field_name"),
                ExceptionType = ReadString(reader, "exception_type"),
                Severity = ReadString(reader, "severity"),
                SourceValue = ReadString(reader, "source_value"),
                ExpectedValue = ReadString(reader, "expected_value"),
                Message = ReadString(reader, "message"),
                ResolutionStatus = ReadString(reader, "resolution_status")
            });
        }
        return records;
    }

    public async Task<List<CarrierBillingRateSummary>> GetManulifeExcelBillingRateSummariesAsync(long carrierBillingImportId)
    {
        await InitializeCoreAsync();
        using var connection = await OpenConnectionAsync();
        using var command = CreateCommand(connection, null, @"
SELECT benefit.benefit_code,
       employee.class_code,
       COALESCE(NULLIF(benefit.benefit_plan_code, ''), employee.plan_code, '') AS plan_code,
       benefit.dependent_code,
       benefit.option_code,
       benefit.insured_volume,
       benefit.current_month_premium
FROM manulife_billing_benefit_records benefit
INNER JOIN manulife_billing_employee_records employee
    ON employee.billing_employee_record_id = benefit.billing_employee_record_id
WHERE benefit.statement_version_id = $statement_version_id
  AND benefit.current_month_premium IS NOT NULL
  AND benefit.current_month_premium <> 0
ORDER BY benefit.benefit_code, employee.class_code, benefit.dependent_code, benefit.option_code;");
        AddParameter(command, "$statement_version_id", carrierBillingImportId);

        var inputs = new List<ManulifeBillingRateInput>();
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            inputs.Add(new ManulifeBillingRateInput(
                ReadString(reader, "benefit_code"),
                ReadString(reader, "class_code"),
                ReadString(reader, "plan_code"),
                ReadString(reader, "dependent_code"),
                ReadString(reader, "option_code"),
                ReadDecimal(reader, "insured_volume"),
                ReadDecimal(reader, "current_month_premium") ?? 0m));
        }

        return inputs
            .Select(input => new
            {
                input.BenefitCode,
                input.ClassCode,
                input.PlanCode,
                Coverage = ResolveCoverage(input.DependentCode, input.OptionCode),
                RateBasis = input.InsuredVolume is > 0m ? "Per $1,000 insured volume" : "Monthly premium",
                Rate = input.InsuredVolume is > 0m
                    ? Math.Round(input.CurrentMonthPremium * 1000m / input.InsuredVolume.Value, 2, MidpointRounding.AwayFromZero)
                    : Math.Round(input.CurrentMonthPremium, 2, MidpointRounding.AwayFromZero)
            })
            .GroupBy(item => new { item.BenefitCode, item.ClassCode, item.PlanCode, item.Coverage, item.RateBasis, item.Rate })
            .Select(group => new CarrierBillingRateSummary
            {
                BenefitCode = group.Key.BenefitCode,
                ClassCode = group.Key.ClassCode,
                PlanCode = group.Key.PlanCode,
                Coverage = group.Key.Coverage,
                RateBasis = group.Key.RateBasis,
                Rate = group.Key.Rate,
                RecordCount = group.Count(),
                Source = "Calculated from distinct Manulife billing records"
            })
            .OrderBy(row => row.BenefitCode, StringComparer.OrdinalIgnoreCase)
            .ThenBy(row => row.ClassCode, StringComparer.OrdinalIgnoreCase)
            .ThenBy(row => row.Coverage, StringComparer.OrdinalIgnoreCase)
            .ThenBy(row => row.Rate)
            .ToList();
    }

    public async Task<DatabaseSaveResult> ReassignManulifeExcelBillingClientAsync(long sourceClientId, long targetClientId)
    {
        if (sourceClientId <= 0 || targetClientId <= 0 || sourceClientId == targetClientId)
        {
            return DatabaseSaveResult.Failure("Select two different saved clients before moving Manulife billing history.", DatabasePath);
        }

        try
        {
            await InitializeCoreAsync();
            using var connection = await OpenConnectionAsync();
            using var transaction = connection.BeginTransaction();

            async Task<int> UpdateAsync(string sql)
            {
                using var command = CreateCommand(connection, transaction, sql);
                AddParameter(command, "$source_client_id", sourceClientId);
                AddParameter(command, "$target_client_id", targetClientId);
                return await command.ExecuteNonQueryAsync();
            }

            var documentsUpdated = await UpdateAsync(@"
UPDATE Documents
SET client_id = $target_client_id
WHERE document_id IN (
    SELECT source_document_id
    FROM manulife_billing_statement_versions
    WHERE client_id = $source_client_id
      AND carrier = 'Manulife'
      AND source_document_id IS NOT NULL
);");
            var employeesUpdated = await UpdateAsync(@"
UPDATE manulife_billing_employee_records
SET client_id = $target_client_id
WHERE statement_version_id IN (
    SELECT statement_version_id
    FROM manulife_billing_statement_versions
    WHERE client_id = $source_client_id
      AND carrier = 'Manulife'
);");
            var benefitsUpdated = await UpdateAsync(@"
UPDATE manulife_billing_benefit_records
SET client_id = $target_client_id
WHERE statement_version_id IN (
    SELECT statement_version_id
    FROM manulife_billing_statement_versions
    WHERE client_id = $source_client_id
      AND carrier = 'Manulife'
);");
            var importsUpdated = await UpdateAsync(@"
UPDATE manulife_billing_statement_versions
SET client_id = $target_client_id
WHERE client_id = $source_client_id
  AND carrier = 'Manulife';");

            transaction.Commit();
            return DatabaseSaveResult.Success(
                $"Moved {importsUpdated} Manulife billing import(s) to the selected client.",
                DatabasePath,
                targetClientId,
                insertedRows: 0,
                updatedRows: documentsUpdated + employeesUpdated + benefitsUpdated + importsUpdated);
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            return DatabaseSaveResult.Failure("Manulife billing history could not be moved. No partial changes were saved.", DatabasePath);
        }
    }

    public Task<List<CarrierBillingAggregate>> GetManulifeBillingTotalsByBenefitAsync(long clientId) => GetManulifeBillingAggregatesAsync(clientId, "benefit_code", "manulife_billing_benefit_records");
    public Task<List<CarrierBillingAggregate>> GetManulifeBillingTotalsByClassAsync(long clientId) => GetManulifeBillingAggregatesAsync(clientId, "class_code", "manulife_billing_employee_records");
    public Task<List<CarrierBillingAggregate>> GetManulifeBillingTotalsByPlanAsync(long clientId) => GetManulifeBillingAggregatesAsync(clientId, "plan_code", "manulife_billing_employee_records");
    public Task<List<CarrierBillingAggregate>> GetManulifeBillingTotalsByProvinceAsync(long clientId) => GetManulifeBillingAggregatesAsync(clientId, "province_of_work", "manulife_billing_employee_records");
    public Task<List<CarrierBillingAggregate>> GetManulifeBillingTotalsByDivisionAsync(long clientId) => GetManulifeBillingAggregatesAsync(clientId, "divisional_billing_code", "manulife_billing_employee_records");

    private async Task<List<CarrierBillingAggregate>> GetManulifeBillingAggregatesAsync(long clientId, string groupColumn, string tableName)
    {
        await InitializeCoreAsync();
        using var connection = await OpenConnectionAsync();
        var valueColumns = tableName == "manulife_billing_benefit_records"
            ? "current_month_premium, prior_period_adjustment, total_sales_tax, benefit_total"
            : "current_month_premium, prior_period_adjustment, total_sales_tax, total_premium_and_tax";
        using var command = CreateCommand(connection, null, $@"
SELECT COALESCE(NULLIF({groupColumn}, ''), 'Not specified') AS group_value,
       COUNT(*) AS record_count,
       COALESCE(SUM(current_month_premium), 0) AS current_month_premium,
       COALESCE(SUM(prior_period_adjustment), 0) AS prior_period_adjustment,
       COALESCE(SUM(total_sales_tax), 0) AS total_sales_tax,
       COALESCE(SUM({valueColumns.Split(',').Last().Trim()}), 0) AS total_premium_and_tax
FROM {tableName}
WHERE client_id = $client_id
GROUP BY COALESCE(NULLIF({groupColumn}, ''), 'Not specified')
ORDER BY group_value; ");
        AddParameter(command, "$client_id", clientId);
        var records = new List<CarrierBillingAggregate>();
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            records.Add(new CarrierBillingAggregate
            {
                Group = ReadString(reader, "group_value"),
                BenefitRecordCount = reader.GetInt32(reader.GetOrdinal("record_count")),
                CurrentMonthPremium = ReadDecimal(reader, "current_month_premium") ?? 0m,
                PriorPeriodAdjustment = ReadDecimal(reader, "prior_period_adjustment") ?? 0m,
                SalesTax = ReadDecimal(reader, "total_sales_tax") ?? 0m,
                TotalPremiumAndTax = ReadDecimal(reader, "total_premium_and_tax") ?? 0m
            });
        }
        return records;
    }

    private static async Task<long?> FindCarrierIdAsync(SqliteConnection connection, SqliteTransaction transaction, string carrierName)
    {
        using var command = CreateCommand(connection, transaction, "SELECT carrier_id FROM Carrier WHERE carrier_name = $carrier_name COLLATE NOCASE LIMIT 1;");
        AddParameter(command, "$carrier_name", carrierName);
        var value = await command.ExecuteScalarAsync();
        return value is null || value == DBNull.Value ? null : Convert.ToInt64(value, CultureInfo.InvariantCulture);
    }

    private static async Task<long?> FindCompatibleManulifeBillingClientIdAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        string clientName)
    {
        var exact = await FindClientIdByNameAsync(connection, transaction, clientName);
        if (exact is not null)
        {
            return exact;
        }

        var targetTokens = ClientNameTokens(clientName);
        if (targetTokens.Count < 2)
        {
            return null;
        }

        using var command = CreateCommand(connection, transaction, "SELECT client_id, client_name FROM Clients;");
        var matches = new List<long>();
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            var candidateTokens = ClientNameTokens(ReadString(reader, "client_name"));
            if (candidateTokens.Count < 2)
            {
                continue;
            }

            if (targetTokens.SetEquals(candidateTokens) ||
                targetTokens.IsSubsetOf(candidateTokens) ||
                candidateTokens.IsSubsetOf(targetTokens))
            {
                matches.Add(reader.GetInt64(reader.GetOrdinal("client_id")));
            }
        }

        return matches.Count == 1 ? matches[0] : null;
    }

    private static HashSet<string> ClientNameTokens(string value)
    {
        var ignored = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "THE", "AND", "CO", "COMPANY", "CORP", "CORPORATION", "INC", "LTD", "LIMITED", "LLC"
        };
        return Regex.Matches(value ?? "", "[A-Za-z0-9]+")
            .Select(match => match.Value.ToUpperInvariant())
            .Where(token => !ignored.Contains(token))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
    }

    private static async Task<long> InsertManulifeExcelBillingVersionAsync(SqliteConnection connection, SqliteTransaction transaction, ManulifeExcelBillingImportPreview preview, ManulifeExcelBillingImportRequest request, long clientId, long? carrierId, long documentId, long? importBatchId, string sourceHash, int versionNumber, string status)
    {
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO manulife_billing_statement_versions (
    client_id, source_document_id, import_batch_id, carrier, carrier_id, billing_period, statement_date,
    source_file_name, source_file_path, source_file_hash, version_number, import_status, is_replaced,
    created_at, created_by, source_row_count, employee_record_count, benefit_record_count,
    total_current_month_premium, total_prior_period_adjustment, total_sales_tax, total_premium_and_tax,
    reconciliation_difference, parser_version, error_count, warning_count, header_mapping_json
)
VALUES (
    $client_id, $source_document_id, $import_batch_id, 'Manulife', $carrier_id, $billing_period, $statement_date,
    $source_file_name, $source_file_path, $source_file_hash, $version_number, $import_status, 0,
    $created_at, $created_by, $source_row_count, $employee_record_count, $benefit_record_count,
    $total_current_month_premium, $total_prior_period_adjustment, $total_sales_tax, $total_premium_and_tax,
    $reconciliation_difference, 'Manulife Excel/CSV billing parser v1', $error_count, $warning_count, $header_mapping_json
);
SELECT last_insert_rowid();");
        AddParameter(command, "$client_id", clientId);
        AddParameter(command, "$source_document_id", documentId);
        AddParameter(command, "$import_batch_id", importBatchId);
        AddParameter(command, "$carrier_id", carrierId);
        AddParameter(command, "$billing_period", preview.BillingPeriod);
        AddParameter(command, "$statement_date", preview.StatementDate?.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));
        AddParameter(command, "$source_file_name", preview.SourceFileName);
        AddParameter(command, "$source_file_path", preview.SourceFilePath);
        AddParameter(command, "$source_file_hash", sourceHash);
        AddParameter(command, "$version_number", versionNumber);
        AddParameter(command, "$import_status", status);
        AddParameter(command, "$created_at", ToDateTimeText(DateTime.Now));
        AddParameter(command, "$created_by", FirstNonBlank(request.ImportedBy, Environment.UserName));
        AddParameter(command, "$source_row_count", preview.SourceRowCount);
        AddParameter(command, "$employee_record_count", preview.EmployeeRecordCount);
        AddParameter(command, "$benefit_record_count", preview.BenefitRecordCount);
        AddParameter(command, "$total_current_month_premium", preview.TotalCurrentMonthPremium);
        AddParameter(command, "$total_prior_period_adjustment", preview.TotalPriorPeriodAdjustment);
        AddParameter(command, "$total_sales_tax", preview.TotalSalesTax);
        AddParameter(command, "$total_premium_and_tax", preview.TotalPremiumAndTax);
        AddParameter(command, "$reconciliation_difference", preview.ReconciliationDifference);
        AddParameter(command, "$error_count", preview.ErrorCount);
        AddParameter(command, "$warning_count", preview.WarningCount);
        AddParameter(command, "$header_mapping_json", System.Text.Json.JsonSerializer.Serialize(preview.HeaderMappings));
        return Convert.ToInt64(await command.ExecuteScalarAsync(), CultureInfo.InvariantCulture);
    }

    private static async Task<long> InsertManulifeExcelBillingEmployeeAsync(SqliteConnection connection, SqliteTransaction transaction, CarrierBillingEmployeeRecord record, long versionId, long clientId, long? carrierId)
    {
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO manulife_billing_employee_records (
    statement_version_id, client_id, carrier_id, certificate_number, last_name, first_name, class_name, class_code,
    plan_name, plan_code, province_of_work, province_of_work_raw, province_of_residence, province_of_residence_raw,
    current_month_premium, prior_period_adjustment, total_sales_tax, total_premium_and_tax,
    calculated_current_month_premium, calculated_prior_period_adjustment, calculated_sales_tax, calculated_premium_and_tax,
    reconciliation_difference, divisional_billing_code, bill_total_6, bill_total_7, source_row_number, source_row_hash, validation_status
)
VALUES (
    $statement_version_id, $client_id, $carrier_id, $certificate_number, $last_name, $first_name, $class_name, $class_code,
    $plan_name, $plan_code, $province_of_work, $province_of_work_raw, $province_of_residence, $province_of_residence_raw,
    $current_month_premium, $prior_period_adjustment, $total_sales_tax, $total_premium_and_tax,
    $calculated_current_month_premium, $calculated_prior_period_adjustment, $calculated_sales_tax, $calculated_premium_and_tax,
    $reconciliation_difference, $divisional_billing_code, $bill_total_6, $bill_total_7, $source_row_number, $source_row_hash, $validation_status
);
SELECT last_insert_rowid();");
        AddEmployeeParameters(command, record, versionId, clientId, carrierId);
        return Convert.ToInt64(await command.ExecuteScalarAsync(), CultureInfo.InvariantCulture);
    }

    private static void AddEmployeeParameters(SqliteCommand command, CarrierBillingEmployeeRecord record, long versionId, long clientId, long? carrierId)
    {
        AddParameter(command, "$statement_version_id", versionId); AddParameter(command, "$client_id", clientId); AddParameter(command, "$carrier_id", carrierId);
        AddParameter(command, "$certificate_number", record.CertificateNumber); AddParameter(command, "$last_name", record.LastName); AddParameter(command, "$first_name", record.FirstName); AddParameter(command, "$class_name", record.ClassName); AddParameter(command, "$class_code", record.ClassCode);
        AddParameter(command, "$plan_name", record.PlanName); AddParameter(command, "$plan_code", record.PlanCode); AddParameter(command, "$province_of_work", record.ProvinceOfWork); AddParameter(command, "$province_of_work_raw", record.ProvinceOfWorkRaw); AddParameter(command, "$province_of_residence", record.ProvinceOfResidence); AddParameter(command, "$province_of_residence_raw", record.ProvinceOfResidenceRaw);
        AddParameter(command, "$current_month_premium", record.CurrentMonthPremium); AddParameter(command, "$prior_period_adjustment", record.PriorPeriodAdjustment); AddParameter(command, "$total_sales_tax", record.TotalSalesTax); AddParameter(command, "$total_premium_and_tax", record.TotalPremiumAndTax);
        AddParameter(command, "$calculated_current_month_premium", record.CalculatedCurrentMonthPremium); AddParameter(command, "$calculated_prior_period_adjustment", record.CalculatedPriorPeriodAdjustment); AddParameter(command, "$calculated_sales_tax", record.CalculatedSalesTax); AddParameter(command, "$calculated_premium_and_tax", record.CalculatedPremiumAndTax); AddParameter(command, "$reconciliation_difference", record.ReconciliationDifference);
        AddParameter(command, "$divisional_billing_code", record.DivisionalBillingCode); AddParameter(command, "$bill_total_6", record.BillTotal6); AddParameter(command, "$bill_total_7", record.BillTotal7); AddParameter(command, "$source_row_number", record.SourceRowNumber); AddParameter(command, "$source_row_hash", record.SourceRowHash); AddParameter(command, "$validation_status", record.ValidationStatus);
    }

    private static async Task InsertManulifeExcelBillingBenefitAsync(SqliteConnection connection, SqliteTransaction transaction, CarrierBillingBenefitRecord record, long employeeId, long versionId, long clientId)
    {
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO manulife_billing_benefit_records (
    billing_employee_record_id, statement_version_id, client_id, certificate_number, benefit_code, benefit_name, benefit_plan_code,
    dependent_code_or_volume_raw, dependent_code, insured_volume, option_code, current_month_premium, prior_period_adjustment,
    total_sales_tax, benefit_total, calculated_benefit_total, reconciliation_difference, source_row_number, validation_status
)
VALUES (
    $billing_employee_record_id, $statement_version_id, $client_id, $certificate_number, $benefit_code, $benefit_name, $benefit_plan_code,
    $dependent_code_or_volume_raw, $dependent_code, $insured_volume, $option_code, $current_month_premium, $prior_period_adjustment,
    $total_sales_tax, $benefit_total, $calculated_benefit_total, $reconciliation_difference, $source_row_number, $validation_status
); ");
        AddParameter(command, "$billing_employee_record_id", employeeId); AddParameter(command, "$statement_version_id", versionId); AddParameter(command, "$client_id", clientId);
        AddParameter(command, "$certificate_number", record.CertificateNumber); AddParameter(command, "$benefit_code", record.BenefitCode); AddParameter(command, "$benefit_name", record.BenefitName); AddParameter(command, "$benefit_plan_code", record.BenefitPlanCode); AddParameter(command, "$dependent_code_or_volume_raw", record.DependentCodeOrVolumeRaw); AddParameter(command, "$dependent_code", record.DependentCode); AddParameter(command, "$insured_volume", record.InsuredVolume); AddParameter(command, "$option_code", record.OptionCode);
        AddParameter(command, "$current_month_premium", record.CurrentMonthPremium); AddParameter(command, "$prior_period_adjustment", record.PriorPeriodAdjustment); AddParameter(command, "$total_sales_tax", record.TotalSalesTax); AddParameter(command, "$benefit_total", record.BenefitTotal); AddParameter(command, "$calculated_benefit_total", record.CalculatedBenefitTotal); AddParameter(command, "$reconciliation_difference", record.ReconciliationDifference); AddParameter(command, "$source_row_number", record.SourceRowNumber); AddParameter(command, "$validation_status", record.ValidationStatus);
        await command.ExecuteNonQueryAsync();
    }

    private static string ResolveCoverage(string dependentCode, string optionCode)
    {
        if (!string.IsNullOrWhiteSpace(dependentCode))
        {
            return dependentCode;
        }

        return string.IsNullOrWhiteSpace(optionCode) ? "Standard" : $"Option {optionCode}";
    }

    private sealed record ManulifeBillingRateInput(
        string BenefitCode,
        string ClassCode,
        string PlanCode,
        string DependentCode,
        string OptionCode,
        decimal? InsuredVolume,
        decimal CurrentMonthPremium);

    private static async Task InsertManulifeExcelBillingExceptionAsync(SqliteConnection connection, SqliteTransaction transaction, CarrierBillingException record, long versionId)
    {
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO manulife_billing_exceptions (
    statement_version_id, source_row_number, certificate_number, benefit_code, field_name, exception_type,
    severity, source_value, expected_value, message, created_at, resolution_status
)
VALUES (
    $statement_version_id, $source_row_number, $certificate_number, $benefit_code, $field_name, $exception_type,
    $severity, $source_value, $expected_value, $message, $created_at, $resolution_status
); ");
        AddParameter(command, "$statement_version_id", versionId); AddParameter(command, "$source_row_number", record.SourceRowNumber); AddParameter(command, "$certificate_number", record.CertificateNumber); AddParameter(command, "$benefit_code", record.BenefitCode); AddParameter(command, "$field_name", record.FieldName); AddParameter(command, "$exception_type", record.ExceptionType); AddParameter(command, "$severity", record.Severity); AddParameter(command, "$source_value", record.SourceValue); AddParameter(command, "$expected_value", record.ExpectedValue); AddParameter(command, "$message", record.Message); AddParameter(command, "$created_at", ToDateTimeText(record.CreatedAt)); AddParameter(command, "$resolution_status", record.ResolutionStatus);
        await command.ExecuteNonQueryAsync();
    }
}

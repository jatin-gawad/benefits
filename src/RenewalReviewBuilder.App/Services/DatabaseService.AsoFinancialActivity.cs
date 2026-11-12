using Microsoft.Data.Sqlite;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed partial class DatabaseService
{
    private static async Task EnsureAsoFinancialActivitySchemaAsync(SqliteConnection connection)
    {
        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS aso_financial_statements (
    statement_id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL,
    source_document_id INTEGER,
    contract_number TEXT,
    carrier TEXT,
    billing_arrangement TEXT,
    divisions TEXT,
    benefits TEXT,
    period_start TEXT,
    period_end TEXT,
    statement_date TEXT,
    source_funding_status TEXT,
    source_page_start INTEGER,
    source_page_end INTEGER,
    source_file_name TEXT NOT NULL,
    source_file_path TEXT NOT NULL,
    source_file_hash TEXT NOT NULL,
    source_sheet TEXT,
    version_number INTEGER NOT NULL DEFAULT 1,
    is_current INTEGER NOT NULL DEFAULT 1,
    import_status TEXT NOT NULL,
    validation_summary TEXT,
    imported_at TEXT NOT NULL,
    imported_by TEXT,
    FOREIGN KEY (client_id) REFERENCES Clients(client_id),
    FOREIGN KEY (source_document_id) REFERENCES Documents(document_id)
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS aso_financial_periods (
    period_id INTEGER PRIMARY KEY AUTOINCREMENT,
    statement_id INTEGER NOT NULL,
    activity_month TEXT,
    period_label TEXT NOT NULL,
    is_total INTEGER NOT NULL DEFAULT 0,
    display_order INTEGER NOT NULL DEFAULT 0,
    opening_balance REAL,
    payment_received_not_allocated REAL,
    dental_claims_paid REAL,
    claims_handling_charges REAL,
    general_admin_charge REAL,
    profit_charge REAL,
    total_service_fees REAL,
    consulting_fees REAL,
    transfers REAL,
    deficit_payment_refund REAL,
    total_deposit REAL,
    orst_deposit REAL,
    qst_deposit REAL,
    hst_deposit REAL,
    tvq_deposit REAL,
    skpst_deposit REAL,
    net_deposit REAL,
    unpooled_benefit_payments REAL,
    pooled_benefit_payments REAL,
    total_benefit_payment REAL,
    loss_ratio REAL,
    pooling_charges REAL,
    admin_fees REAL,
    other_charges REAL,
    total_taxes REAL,
    gst REAL,
    hst REAL,
    tvq REAL,
    insured_premium_tax REAL,
    premium_tax_quebec REAL,
    premium_tax_ontario REAL,
    premium_tax_newfoundland REAL,
    retroactive_tax_qst REAL,
    retroactive_tax_orst REAL,
    retroactive_tax_skpst REAL,
    interest_credit_charge REAL,
    total_charges REAL,
    sub_total REAL,
    required_float REAL,
    current_balance REAL,
    balance_status TEXT,
    FOREIGN KEY (statement_id) REFERENCES aso_financial_statements(statement_id) ON DELETE CASCADE
);");

        await EnsureColumnExistsAsync(connection, "aso_financial_statements", "statement_date", "TEXT");
        await EnsureColumnExistsAsync(connection, "aso_financial_statements", "source_funding_status", "TEXT");
        await EnsureColumnExistsAsync(connection, "aso_financial_statements", "source_page_start", "INTEGER");
        await EnsureColumnExistsAsync(connection, "aso_financial_statements", "source_page_end", "INTEGER");
        await EnsureColumnExistsAsync(connection, "aso_financial_periods", "payment_received_not_allocated", "REAL");
        await EnsureColumnExistsAsync(connection, "aso_financial_periods", "dental_claims_paid", "REAL");
        await EnsureColumnExistsAsync(connection, "aso_financial_periods", "claims_handling_charges", "REAL");
        await EnsureColumnExistsAsync(connection, "aso_financial_periods", "general_admin_charge", "REAL");
        await EnsureColumnExistsAsync(connection, "aso_financial_periods", "profit_charge", "REAL");
        await EnsureColumnExistsAsync(connection, "aso_financial_periods", "total_service_fees", "REAL");
        await EnsureColumnExistsAsync(connection, "aso_financial_periods", "consulting_fees", "REAL");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_AsoFinancialStatements_ClientCurrent
ON aso_financial_statements(client_id, is_current, period_end DESC);");
        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_AsoFinancialStatements_SourceHash
ON aso_financial_statements(client_id, source_file_hash);");
        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_AsoFinancialPeriods_StatementMonth
ON aso_financial_periods(statement_id, is_total, activity_month, display_order);");
        await ExecuteNonQueryAsync(connection, null, @"
CREATE UNIQUE INDEX IF NOT EXISTS UX_AsoFinancialPeriods_StatementRow
ON aso_financial_periods(statement_id, period_label COLLATE NOCASE, is_total);");
    }

    public async Task<AsoFinancialActivitySaveResult> SaveAsoFinancialActivityAsync(
        AsoFinancialActivityStatement statement,
        string importedBy = "Benefits Intelligence OS",
        long? selectedClientId = null,
        long? existingSourceDocumentId = null)
    {
        ArgumentNullException.ThrowIfNull(statement);
        if (string.IsNullOrWhiteSpace(statement.SourceFilePath) || !File.Exists(statement.SourceFilePath))
        {
            return FailedAsoSave("The source ASO financial activity file could not be found.");
        }

        if (statement.Periods.Count == 0)
        {
            return FailedAsoSave("No monthly ASO financial activity rows were available to save.");
        }

        await InitializeCoreAsync();
        using var connection = await OpenConnectionAsync();
        using var transaction = connection.BeginTransaction();
        try
        {
            var clientId = await ResolveAsoFinancialClientAsync(connection, transaction, statement, selectedClientId);
            var sourceHash = await ComputeFileHashAsync(statement.SourceFilePath);
            statement.SourceHash = sourceHash;

            var existingStatementId = await FindAsoFinancialStatementByHashAsync(
                connection,
                transaction,
                clientId,
                sourceHash,
                statement.ContractNumber,
                statement.PeriodStart,
                statement.PeriodEnd);
            if (existingStatementId.HasValue)
            {
                transaction.Commit();
                return new AsoFinancialActivitySaveResult
                {
                    Succeeded = true,
                    Message = "This ASO financial activity source is already saved for the client.",
                    DatabasePath = DatabasePath,
                    ClientId = clientId,
                    StatementId = existingStatementId,
                    InsertedRows = 0,
                    WasDuplicate = true
                };
            }

            var sourceDocumentId = existingSourceDocumentId;
            if (!sourceDocumentId.HasValue)
            {
                var document = await UpsertDocumentAsync(
                    connection,
                    transaction,
                    clientId,
                    DocumentTypes.AsoFinancialActivity,
                    BenefitsIntelligenceModules.FinancialIntelligence,
                    statement.SourceFilePath,
                    statement.Carrier,
                    statement.StatementDate ?? statement.PeriodEnd,
                    statement.PeriodStart,
                    statement.PeriodEnd);
                sourceDocumentId = document.DocumentId;
            }

            var versionNumber = await GetNextAsoFinancialVersionAsync(
                connection,
                transaction,
                clientId,
                statement.ContractNumber,
                statement.PeriodStart,
                statement.PeriodEnd);
            var isCurrent = statement.ImportStatus.Equals(AsoFinancialImportStatuses.Complete, StringComparison.OrdinalIgnoreCase);
            if (isCurrent)
            {
                await MarkPriorAsoFinancialVersionsInactiveAsync(
                    connection,
                    transaction,
                    clientId,
                    statement.ContractNumber,
                    statement.PeriodStart,
                    statement.PeriodEnd);
            }

            var statementId = await InsertAsoFinancialStatementAsync(
                connection,
                transaction,
                clientId,
                sourceDocumentId,
                statement,
                sourceHash,
                versionNumber,
                importedBy,
                isCurrent);
            var insertedRows = 0;
            foreach (var period in statement.Periods)
            {
                await InsertAsoFinancialPeriodAsync(connection, transaction, statementId, period);
                insertedRows++;
            }

            if (statement.Total is not null)
            {
                await InsertAsoFinancialPeriodAsync(connection, transaction, statementId, statement.Total);
                insertedRows++;
            }

            transaction.Commit();
            statement.StatementId = statementId;
            statement.ClientId = clientId;
            statement.SourceDocumentId = sourceDocumentId;
            statement.VersionNumber = versionNumber;
            statement.IsCurrent = isCurrent;
            statement.ImportedBy = importedBy;
            InvalidateExecutiveReportingCache();
            return new AsoFinancialActivitySaveResult
            {
                Succeeded = true,
                Message = $"Saved ASO financial activity version {versionNumber} with {insertedRows} rows.",
                DatabasePath = DatabasePath,
                ClientId = clientId,
                DocumentId = sourceDocumentId,
                StatementId = statementId,
                InsertedRows = insertedRows
            };
        }
        catch (AsoFinancialPolicyLinkRequiredException exception)
        {
            transaction.Rollback();
            return new AsoFinancialActivitySaveResult
            {
                Succeeded = false,
                Message = $"ASO financial activity was not saved: {exception.Message}",
                DatabasePath = DatabasePath,
                RequiresVerifiedAsoPolicyLink = true,
                WorkbookContractNumber = exception.WorkbookContractNumber,
                SelectedClientNumber = exception.SelectedClientNumber
            };
        }
        catch (Exception exception)
        {
            transaction.Rollback();
            return FailedAsoSave($"ASO financial activity was not saved: {exception.Message}");
        }
    }

    public async Task<AsoFinancialActivityStatement?> GetLatestAsoFinancialActivityAsync(long clientId)
    {
        await InitializeCoreAsync();
        using var connection = await OpenConnectionAsync();
        using var command = CreateCommand(connection, null, @"
SELECT s.*, c.client_name
FROM aso_financial_statements s
INNER JOIN Clients c ON c.client_id = s.client_id
WHERE s.client_id = $client_id AND s.is_current = 1
ORDER BY COALESCE(s.period_end, '') DESC, s.version_number DESC, s.statement_id DESC
LIMIT 1;");
        AddParameter(command, "$client_id", clientId);
        using var reader = await command.ExecuteReaderAsync();
        if (!await reader.ReadAsync())
        {
            return null;
        }

        var statement = ReadAsoFinancialStatement(reader);
        await reader.CloseAsync();
        statement.Periods.AddRange(await GetAsoFinancialPeriodsAsync(connection, statement.StatementId, false));
        statement.Total = (await GetAsoFinancialPeriodsAsync(connection, statement.StatementId, true)).FirstOrDefault();
        return statement;
    }

    public async Task<IReadOnlyList<AsoFinancialActivityStatement>> GetAsoFinancialActivityHistoryAsync(long clientId)
    {
        await InitializeCoreAsync();
        using var connection = await OpenConnectionAsync();
        using var command = CreateCommand(connection, null, @"
SELECT s.*, c.client_name
FROM aso_financial_statements s
INNER JOIN Clients c ON c.client_id = s.client_id
WHERE s.client_id = $client_id
ORDER BY COALESCE(s.period_end, '') DESC, s.version_number DESC, s.statement_id DESC;");
        AddParameter(command, "$client_id", clientId);
        var statements = new List<AsoFinancialActivityStatement>();
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            statements.Add(ReadAsoFinancialStatement(reader));
        }

        foreach (var statement in statements)
        {
            statement.Periods.AddRange(await GetAsoFinancialPeriodsAsync(connection, statement.StatementId, false));
            statement.Total = (await GetAsoFinancialPeriodsAsync(connection, statement.StatementId, true)).FirstOrDefault();
        }

        return statements;
    }

    public async Task<AsoExperiencePeriodReference?> GetRenewalExperiencePeriodAsync(long clientId)
    {
        await InitializeCoreAsync();
        using var connection = await OpenConnectionAsync();
        using var command = CreateCommand(connection, null, @"
SELECT d.period_start, d.period_end, d.file_name
FROM Documents d
INNER JOIN Renewals r ON r.source_document_id = d.document_id
WHERE r.client_id = $client_id
  AND d.document_type = 'renewal'
  AND d.period_start IS NOT NULL
  AND d.period_end IS NOT NULL
ORDER BY COALESCE(d.effective_date, '') DESC, d.document_id DESC
LIMIT 1;");
        AddParameter(command, "$client_id", clientId);
        using var reader = await command.ExecuteReaderAsync();
        if (!await reader.ReadAsync()) return null;

        return new AsoExperiencePeriodReference
        {
            PeriodStart = DateOnly.TryParse(ReadString(reader, "period_start"), out var start) ? start : null,
            PeriodEnd = DateOnly.TryParse(ReadString(reader, "period_end"), out var end) ? end : null,
            SourceDescription = $"Renewal document: {ReadString(reader, "file_name")}" 
        };
    }

    /// <summary>
    /// Repairs an older renewal import that saved rate rows but predates
    /// persistence of its extracted experience-history period.
    /// </summary>
    public async Task<bool> BackfillRenewalExperiencePeriodAsync(long clientId)
    {
        await InitializeCoreAsync();
        using var connection = await OpenConnectionAsync();
        using var select = CreateCommand(connection, null, @"
SELECT d.document_id, d.file_path, d.carrier
FROM Documents d
INNER JOIN Renewals r ON r.source_document_id = d.document_id
WHERE r.client_id = $client_id
  AND d.document_type = 'renewal'
  AND (d.period_start IS NULL OR d.period_end IS NULL)
ORDER BY COALESCE(d.effective_date, '') DESC, d.document_id DESC
LIMIT 1;");
        AddParameter(select, "$client_id", clientId);
        using var reader = await select.ExecuteReaderAsync();
        if (!await reader.ReadAsync()) return false;
        var documentId = reader.GetInt64(reader.GetOrdinal("document_id"));
        var sourcePath = ReadString(reader, "file_path");
        var carrier = ReadString(reader, "carrier");
        await reader.CloseAsync();

        if (!carrier.Equals("Manulife", StringComparison.OrdinalIgnoreCase) || !File.Exists(sourcePath)) return false;
        var periods = new ManulifePdfParser().Parse(sourcePath).ExperienceLines
            .Select(line => ParseExperiencePeriod(line.Period))
            .Where(period => period.Start.HasValue && period.End.HasValue)
            .OrderByDescending(period => period.End)
            .ToList();
        if (periods.Count == 0) return false;

        using var update = CreateCommand(connection, null, @"
UPDATE Documents
SET period_start = $period_start, period_end = $period_end
WHERE document_id = $document_id;");
        AddParameter(update, "$period_start", periods[0].Start!.Value.ToString("yyyy-MM-dd"));
        AddParameter(update, "$period_end", periods[0].End!.Value.ToString("yyyy-MM-dd"));
        AddParameter(update, "$document_id", documentId);
        return await update.ExecuteNonQueryAsync() == 1;
    }

    private AsoFinancialActivitySaveResult FailedAsoSave(string message)
    {
        return new AsoFinancialActivitySaveResult
        {
            Succeeded = false,
            Message = message,
            DatabasePath = DatabasePath
        };
    }

    private static async Task<long> ResolveAsoFinancialClientAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        AsoFinancialActivityStatement statement,
        long? selectedClientId)
    {
        if (selectedClientId.HasValue)
        {
            using var selected = CreateCommand(connection, transaction, @"
SELECT client_id, client_name, client_number
FROM Clients
WHERE client_id = $client_id
LIMIT 1;");
            AddParameter(selected, "$client_id", selectedClientId.Value);
            using var reader = await selected.ExecuteReaderAsync();
            if (!await reader.ReadAsync())
            {
                throw new InvalidOperationException("The selected client no longer exists.");
            }

            var selectedName = ReadString(reader, "client_name");
            var selectedNumber = ReadString(reader, "client_number");
            if (string.IsNullOrWhiteSpace(statement.ClientName) ||
                !ClientNamesMatch(selectedName, statement.ClientName))
            {
                throw new InvalidOperationException(
                    "The ASO financial activity client name does not match the selected client.");
            }

            if (!string.IsNullOrWhiteSpace(statement.ContractNumber) &&
                !string.IsNullOrWhiteSpace(selectedNumber) &&
                !statement.ContractNumber.Equals(selectedNumber, StringComparison.OrdinalIgnoreCase) &&
                !await IsVerifiedAsoPolicyAsync(
                    connection,
                    transaction,
                    selectedClientId.Value,
                    statement.ContractNumber))
            {
                throw new AsoFinancialPolicyLinkRequiredException(statement.ContractNumber, selectedNumber);
            }

            return selectedClientId.Value;
        }

        var clientName = string.IsNullOrWhiteSpace(statement.ClientName)
            ? "Unknown Client"
            : statement.ClientName;
        return await UpsertClientAsync(
            connection,
            transaction,
            new ClientRecord
            {
                ClientName = clientName,
                ClientNumber = statement.ContractNumber,
                Carrier = statement.Carrier,
                EffectiveDate = statement.PeriodEnd,
                Status = "Active"
            });
    }

    private sealed class AsoFinancialPolicyLinkRequiredException : InvalidOperationException
    {
        public AsoFinancialPolicyLinkRequiredException(string workbookContractNumber, string selectedClientNumber)
            : base($"The workbook contract {workbookContractNumber} does not match the selected client {selectedClientNumber}.")
        {
            WorkbookContractNumber = workbookContractNumber;
            SelectedClientNumber = selectedClientNumber;
        }

        public string WorkbookContractNumber { get; }
        public string SelectedClientNumber { get; }
    }

    private static async Task<long?> FindAsoFinancialStatementByHashAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long clientId,
        string sourceHash,
        string contractNumber,
        DateOnly? periodStart,
        DateOnly? periodEnd)
    {
        using var command = CreateCommand(connection, transaction, @"
SELECT statement_id
FROM aso_financial_statements
WHERE client_id = $client_id
  AND source_file_hash = $source_file_hash
  AND COALESCE(contract_number, '') = COALESCE($contract_number, '')
  AND COALESCE(period_start, '') = COALESCE($period_start, '')
  AND COALESCE(period_end, '') = COALESCE($period_end, '')
LIMIT 1;");
        AddParameter(command, "$client_id", clientId);
        AddParameter(command, "$source_file_hash", sourceHash);
        AddParameter(command, "$contract_number", contractNumber);
        AddParameter(command, "$period_start", ToDateText(periodStart));
        AddParameter(command, "$period_end", ToDateText(periodEnd));
        return ToNullableLong(await command.ExecuteScalarAsync());
    }

    private static async Task<int> GetNextAsoFinancialVersionAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long clientId,
        string contractNumber,
        DateOnly? periodStart,
        DateOnly? periodEnd)
    {
        using var command = CreateCommand(connection, transaction, @"
SELECT COALESCE(MAX(version_number), 0) + 1
FROM aso_financial_statements
WHERE client_id = $client_id
  AND COALESCE(contract_number, '') = COALESCE($contract_number, '')
  AND COALESCE(period_start, '') = COALESCE($period_start, '')
  AND COALESCE(period_end, '') = COALESCE($period_end, '');");
        AddParameter(command, "$client_id", clientId);
        AddParameter(command, "$contract_number", contractNumber);
        AddParameter(command, "$period_start", ToDateText(periodStart));
        AddParameter(command, "$period_end", ToDateText(periodEnd));
        return Convert.ToInt32(await command.ExecuteScalarAsync());
    }

    private static async Task MarkPriorAsoFinancialVersionsInactiveAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long clientId,
        string contractNumber,
        DateOnly? periodStart,
        DateOnly? periodEnd)
    {
        using var command = CreateCommand(connection, transaction, @"
UPDATE aso_financial_statements
SET is_current = 0
WHERE client_id = $client_id
  AND COALESCE(contract_number, '') = COALESCE($contract_number, '')
  AND COALESCE(period_start, '') = COALESCE($period_start, '')
  AND COALESCE(period_end, '') = COALESCE($period_end, '')
  AND is_current = 1;");
        AddParameter(command, "$client_id", clientId);
        AddParameter(command, "$contract_number", contractNumber);
        AddParameter(command, "$period_start", ToDateText(periodStart));
        AddParameter(command, "$period_end", ToDateText(periodEnd));
        await command.ExecuteNonQueryAsync();
    }

    private static async Task<long> InsertAsoFinancialStatementAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long clientId,
        long? documentId,
        AsoFinancialActivityStatement statement,
        string sourceHash,
        int versionNumber,
        string importedBy,
        bool isCurrent)
    {
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO aso_financial_statements (
    client_id, source_document_id, contract_number, carrier, billing_arrangement,
    divisions, benefits, period_start, period_end, statement_date, source_funding_status,
    source_page_start, source_page_end, source_file_name, source_file_path,
    source_file_hash, source_sheet, version_number, is_current, import_status,
    validation_summary, imported_at, imported_by
)
VALUES (
    $client_id, $source_document_id, NULLIF($contract_number, ''), NULLIF($carrier, ''),
    NULLIF($billing_arrangement, ''), NULLIF($divisions, ''), NULLIF($benefits, ''),
    NULLIF($period_start, ''), NULLIF($period_end, ''), NULLIF($statement_date, ''),
    NULLIF($source_funding_status, ''), $source_page_start, $source_page_end,
    $source_file_name, $source_file_path, $source_file_hash, NULLIF($source_sheet, ''),
    $version_number, $is_current, $import_status,
    NULLIF($validation_summary, ''), $imported_at, NULLIF($imported_by, '')
);");
        AddParameter(command, "$client_id", clientId);
        AddParameter(command, "$source_document_id", documentId);
        AddParameter(command, "$contract_number", statement.ContractNumber);
        AddParameter(command, "$carrier", statement.Carrier);
        AddParameter(command, "$billing_arrangement", statement.BillingArrangement);
        AddParameter(command, "$divisions", statement.Divisions);
        AddParameter(command, "$benefits", statement.Benefits);
        AddParameter(command, "$period_start", ToDateText(statement.PeriodStart));
        AddParameter(command, "$period_end", ToDateText(statement.PeriodEnd));
        AddParameter(command, "$statement_date", ToDateText(statement.StatementDate));
        AddParameter(command, "$source_funding_status", statement.SourceFundingStatus);
        AddParameter(command, "$source_page_start", statement.SourcePageStart);
        AddParameter(command, "$source_page_end", statement.SourcePageEnd);
        AddParameter(command, "$source_file_name", statement.SourceFileName);
        AddParameter(command, "$source_file_path", Path.GetFullPath(statement.SourceFilePath));
        AddParameter(command, "$source_file_hash", sourceHash);
        AddParameter(command, "$source_sheet", statement.SourceSheet);
        AddParameter(command, "$version_number", versionNumber);
        AddParameter(command, "$is_current", isCurrent ? 1 : 0);
        AddParameter(command, "$import_status", statement.ImportStatus);
        AddParameter(command, "$validation_summary", statement.ValidationSummary);
        AddParameter(command, "$imported_at", ToDateTimeText(DateTime.Now));
        AddParameter(command, "$imported_by", importedBy);
        await command.ExecuteNonQueryAsync();
        using var idCommand = CreateCommand(connection, transaction, "SELECT last_insert_rowid();");
        return ToLong(await idCommand.ExecuteScalarAsync());
    }

    private static async Task InsertAsoFinancialPeriodAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long statementId,
        AsoFinancialActivityPeriod period)
    {
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO aso_financial_periods (
    statement_id, activity_month, period_label, is_total, display_order, opening_balance,
    payment_received_not_allocated, dental_claims_paid, claims_handling_charges,
    general_admin_charge, profit_charge, total_service_fees, consulting_fees,
    transfers, deficit_payment_refund, total_deposit, orst_deposit, qst_deposit, hst_deposit,
    tvq_deposit, skpst_deposit, net_deposit, unpooled_benefit_payments,
    pooled_benefit_payments, total_benefit_payment, loss_ratio, pooling_charges, admin_fees,
    other_charges, total_taxes, gst, hst, tvq, insured_premium_tax, premium_tax_quebec,
    premium_tax_ontario, premium_tax_newfoundland, retroactive_tax_qst, retroactive_tax_orst,
    retroactive_tax_skpst, interest_credit_charge, total_charges, sub_total, required_float,
    current_balance, balance_status
)
VALUES (
    $statement_id, NULLIF($activity_month, ''), $period_label, $is_total, $display_order,
    $opening_balance, $payment_received_not_allocated, $dental_claims_paid,
    $claims_handling_charges, $general_admin_charge, $profit_charge, $total_service_fees,
    $consulting_fees, $transfers, $deficit_payment_refund, $total_deposit, $orst_deposit,
    $qst_deposit, $hst_deposit, $tvq_deposit, $skpst_deposit, $net_deposit,
    $unpooled_benefit_payments, $pooled_benefit_payments, $total_benefit_payment, $loss_ratio,
    $pooling_charges, $admin_fees, $other_charges, $total_taxes, $gst, $hst, $tvq,
    $insured_premium_tax, $premium_tax_quebec, $premium_tax_ontario,
    $premium_tax_newfoundland, $retroactive_tax_qst, $retroactive_tax_orst,
    $retroactive_tax_skpst, $interest_credit_charge, $total_charges, $sub_total,
    $required_float, $current_balance, NULLIF($balance_status, '')
);");
        AddParameter(command, "$statement_id", statementId);
        AddParameter(command, "$activity_month", ToDateText(period.ActivityMonth));
        AddParameter(command, "$period_label", period.PeriodLabel);
        AddParameter(command, "$is_total", period.IsTotal ? 1 : 0);
        AddParameter(command, "$display_order", period.DisplayOrder);
        AddParameter(command, "$opening_balance", period.OpeningBalance);
        AddParameter(command, "$payment_received_not_allocated", period.PaymentReceivedNotAllocated);
        AddParameter(command, "$dental_claims_paid", period.DentalClaimsPaid);
        AddParameter(command, "$claims_handling_charges", period.ClaimsHandlingCharges);
        AddParameter(command, "$general_admin_charge", period.GeneralAdminCharge);
        AddParameter(command, "$profit_charge", period.ProfitCharge);
        AddParameter(command, "$total_service_fees", period.TotalServiceFees);
        AddParameter(command, "$consulting_fees", period.ConsultingFees);
        AddParameter(command, "$transfers", period.Transfers);
        AddParameter(command, "$deficit_payment_refund", period.DeficitPaymentRefund);
        AddParameter(command, "$total_deposit", period.TotalDeposit);
        AddParameter(command, "$orst_deposit", period.OrstDeposit);
        AddParameter(command, "$qst_deposit", period.QstDeposit);
        AddParameter(command, "$hst_deposit", period.HstDeposit);
        AddParameter(command, "$tvq_deposit", period.TvqDeposit);
        AddParameter(command, "$skpst_deposit", period.SkpstDeposit);
        AddParameter(command, "$net_deposit", period.NetDeposit);
        AddParameter(command, "$unpooled_benefit_payments", period.UnpooledBenefitPayments);
        AddParameter(command, "$pooled_benefit_payments", period.PooledBenefitPayments);
        AddParameter(command, "$total_benefit_payment", period.TotalBenefitPayment);
        AddParameter(command, "$loss_ratio", period.LossRatio);
        AddParameter(command, "$pooling_charges", period.PoolingCharges);
        AddParameter(command, "$admin_fees", period.AdminFees);
        AddParameter(command, "$other_charges", period.OtherCharges);
        AddParameter(command, "$total_taxes", period.TotalTaxes);
        AddParameter(command, "$gst", period.Gst);
        AddParameter(command, "$hst", period.Hst);
        AddParameter(command, "$tvq", period.Tvq);
        AddParameter(command, "$insured_premium_tax", period.InsuredPremiumTax);
        AddParameter(command, "$premium_tax_quebec", period.PremiumTaxQuebec);
        AddParameter(command, "$premium_tax_ontario", period.PremiumTaxOntario);
        AddParameter(command, "$premium_tax_newfoundland", period.PremiumTaxNewfoundland);
        AddParameter(command, "$retroactive_tax_qst", period.RetroactiveTaxQst);
        AddParameter(command, "$retroactive_tax_orst", period.RetroactiveTaxOrst);
        AddParameter(command, "$retroactive_tax_skpst", period.RetroactiveTaxSkpst);
        AddParameter(command, "$interest_credit_charge", period.InterestCreditCharge);
        AddParameter(command, "$total_charges", period.TotalCharges);
        AddParameter(command, "$sub_total", period.SubTotal);
        AddParameter(command, "$required_float", period.RequiredFloat);
        AddParameter(command, "$current_balance", period.CurrentBalance);
        AddParameter(command, "$balance_status", period.BalanceStatus);
        await command.ExecuteNonQueryAsync();
    }

    private static AsoFinancialActivityStatement ReadAsoFinancialStatement(SqliteDataReader reader)
    {
        return new AsoFinancialActivityStatement
        {
            StatementId = ReadLong(reader, "statement_id"),
            ClientId = ReadLong(reader, "client_id"),
            SourceDocumentId = ReadNullableLong(reader, "source_document_id"),
            ClientName = ReadString(reader, "client_name"),
            ContractNumber = ReadString(reader, "contract_number"),
            Carrier = ReadString(reader, "carrier"),
            BillingArrangement = ReadString(reader, "billing_arrangement"),
            Divisions = ReadString(reader, "divisions"),
            Benefits = ReadString(reader, "benefits"),
            PeriodStart = ReadDateOnly(reader, "period_start"),
            PeriodEnd = ReadDateOnly(reader, "period_end"),
            StatementDate = ReadDateOnly(reader, "statement_date"),
            SourceFundingStatus = ReadString(reader, "source_funding_status"),
            SourcePageStart = ReadInt(reader, "source_page_start"),
            SourcePageEnd = ReadInt(reader, "source_page_end"),
            SourceFileName = ReadString(reader, "source_file_name"),
            SourceFilePath = ReadString(reader, "source_file_path"),
            SourceHash = ReadString(reader, "source_file_hash"),
            SourceSheet = ReadString(reader, "source_sheet"),
            VersionNumber = ReadInt(reader, "version_number") ?? 1,
            IsCurrent = (ReadInt(reader, "is_current") ?? 0) == 1,
            ImportStatus = ReadString(reader, "import_status"),
            ValidationSummary = ReadString(reader, "validation_summary"),
            ImportedAt = ReadDateTime(reader, "imported_at") ?? DateTime.MinValue,
            ImportedBy = ReadString(reader, "imported_by")
        };
    }

    private static async Task<List<AsoFinancialActivityPeriod>> GetAsoFinancialPeriodsAsync(
        SqliteConnection connection,
        long statementId,
        bool totalsOnly)
    {
        using var command = CreateCommand(connection, null, @"
SELECT *
FROM aso_financial_periods
WHERE statement_id = $statement_id AND is_total = $is_total
ORDER BY display_order, activity_month;");
        AddParameter(command, "$statement_id", statementId);
        AddParameter(command, "$is_total", totalsOnly ? 1 : 0);
        var periods = new List<AsoFinancialActivityPeriod>();
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            periods.Add(ReadAsoFinancialPeriod(reader));
        }

        return periods;
    }

    private static AsoFinancialActivityPeriod ReadAsoFinancialPeriod(SqliteDataReader reader)
    {
        return new AsoFinancialActivityPeriod
        {
            PeriodId = ReadLong(reader, "period_id"),
            StatementId = ReadLong(reader, "statement_id"),
            ActivityMonth = ReadDateOnly(reader, "activity_month"),
            PeriodLabel = ReadString(reader, "period_label"),
            IsTotal = (ReadInt(reader, "is_total") ?? 0) == 1,
            DisplayOrder = ReadInt(reader, "display_order") ?? 0,
            OpeningBalance = ReadDecimal(reader, "opening_balance"),
            PaymentReceivedNotAllocated = ReadDecimal(reader, "payment_received_not_allocated"),
            DentalClaimsPaid = ReadDecimal(reader, "dental_claims_paid"),
            ClaimsHandlingCharges = ReadDecimal(reader, "claims_handling_charges"),
            GeneralAdminCharge = ReadDecimal(reader, "general_admin_charge"),
            ProfitCharge = ReadDecimal(reader, "profit_charge"),
            TotalServiceFees = ReadDecimal(reader, "total_service_fees"),
            ConsultingFees = ReadDecimal(reader, "consulting_fees"),
            Transfers = ReadDecimal(reader, "transfers"),
            DeficitPaymentRefund = ReadDecimal(reader, "deficit_payment_refund"),
            TotalDeposit = ReadDecimal(reader, "total_deposit"),
            OrstDeposit = ReadDecimal(reader, "orst_deposit"),
            QstDeposit = ReadDecimal(reader, "qst_deposit"),
            HstDeposit = ReadDecimal(reader, "hst_deposit"),
            TvqDeposit = ReadDecimal(reader, "tvq_deposit"),
            SkpstDeposit = ReadDecimal(reader, "skpst_deposit"),
            NetDeposit = ReadDecimal(reader, "net_deposit"),
            UnpooledBenefitPayments = ReadDecimal(reader, "unpooled_benefit_payments"),
            PooledBenefitPayments = ReadDecimal(reader, "pooled_benefit_payments"),
            TotalBenefitPayment = ReadDecimal(reader, "total_benefit_payment"),
            LossRatio = ReadDecimal(reader, "loss_ratio"),
            PoolingCharges = ReadDecimal(reader, "pooling_charges"),
            AdminFees = ReadDecimal(reader, "admin_fees"),
            OtherCharges = ReadDecimal(reader, "other_charges"),
            TotalTaxes = ReadDecimal(reader, "total_taxes"),
            Gst = ReadDecimal(reader, "gst"),
            Hst = ReadDecimal(reader, "hst"),
            Tvq = ReadDecimal(reader, "tvq"),
            InsuredPremiumTax = ReadDecimal(reader, "insured_premium_tax"),
            PremiumTaxQuebec = ReadDecimal(reader, "premium_tax_quebec"),
            PremiumTaxOntario = ReadDecimal(reader, "premium_tax_ontario"),
            PremiumTaxNewfoundland = ReadDecimal(reader, "premium_tax_newfoundland"),
            RetroactiveTaxQst = ReadDecimal(reader, "retroactive_tax_qst"),
            RetroactiveTaxOrst = ReadDecimal(reader, "retroactive_tax_orst"),
            RetroactiveTaxSkpst = ReadDecimal(reader, "retroactive_tax_skpst"),
            InterestCreditCharge = ReadDecimal(reader, "interest_credit_charge"),
            TotalCharges = ReadDecimal(reader, "total_charges"),
            SubTotal = ReadDecimal(reader, "sub_total"),
            RequiredFloat = ReadDecimal(reader, "required_float"),
            CurrentBalance = ReadDecimal(reader, "current_balance"),
            BalanceStatus = ReadString(reader, "balance_status")
        };
    }
}

using Microsoft.Data.Sqlite;

namespace RenewalReviewBuilder.App.Services;

/// <summary>Additive, idempotent Phase 1 schema for canonical claims.</summary>
public static class CanonicalClaimsDatabaseMigration
{
    public static async Task ApplyAsync(SqliteConnection connection)
    {
        ArgumentNullException.ThrowIfNull(connection);
        await ExecuteAsync(connection, @"
CREATE TABLE IF NOT EXISTS CanonicalClaimsReport (
    canonical_claims_report_id INTEGER PRIMARY KEY AUTOINCREMENT,
    idempotency_key TEXT NOT NULL,
    client_id INTEGER NOT NULL,
    carrier_id INTEGER,
    carrier_name TEXT NOT NULL DEFAULT '',
    policy_number TEXT NOT NULL DEFAULT '',
    division TEXT NOT NULL DEFAULT '',
    class TEXT NOT NULL DEFAULT '',
    import_batch_id INTEGER,
    source_document_id INTEGER,
    source_document_version_id INTEGER,
    mapping_version TEXT NOT NULL,
    validation_status TEXT NOT NULL,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    FOREIGN KEY (client_id) REFERENCES Clients(client_id),
    FOREIGN KEY (import_batch_id) REFERENCES ImportBatch(import_batch_id),
    FOREIGN KEY (source_document_id) REFERENCES Documents(document_id)
);
CREATE UNIQUE INDEX IF NOT EXISTS UX_CanonicalClaimsReport_Idempotency
ON CanonicalClaimsReport(idempotency_key);
CREATE INDEX IF NOT EXISTS IX_CanonicalClaimsReport_ClientPeriod
ON CanonicalClaimsReport(client_id, carrier_name, policy_number);

CREATE TABLE IF NOT EXISTS CanonicalClaimsPeriod (
    canonical_claims_period_id INTEGER PRIMARY KEY AUTOINCREMENT,
    canonical_claims_report_id INTEGER NOT NULL UNIQUE,
    period_start TEXT,
    period_end TEXT,
    period_type TEXT NOT NULL,
    FOREIGN KEY (canonical_claims_report_id) REFERENCES CanonicalClaimsReport(canonical_claims_report_id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS CanonicalClaimsFinancialFact (
    canonical_financial_fact_id INTEGER PRIMARY KEY AUTOINCREMENT,
    canonical_claims_report_id INTEGER NOT NULL,
    fact_id TEXT NOT NULL UNIQUE,
    benefit TEXT NOT NULL DEFAULT '', coverage TEXT NOT NULL DEFAULT '', category TEXT NOT NULL DEFAULT '',
    division TEXT NOT NULL DEFAULT '', class TEXT NOT NULL DEFAULT '', measurement_type TEXT NOT NULL DEFAULT '',
    dimension_type TEXT NOT NULL DEFAULT '', dimension_value TEXT NOT NULL DEFAULT '',
    premium NUMERIC, paid_claims NUMERIC, incurred_claims NUMERIC, loss_ratio NUMERIC, lives INTEGER,
    currency TEXT NOT NULL DEFAULT 'CAD', unit_of_measure TEXT NOT NULL DEFAULT '',
    availability_state TEXT NOT NULL, validation_status TEXT NOT NULL,
    legacy_source_record_id INTEGER, created_at TEXT NOT NULL, updated_at TEXT NOT NULL,
    FOREIGN KEY (canonical_claims_report_id) REFERENCES CanonicalClaimsReport(canonical_claims_report_id) ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS IX_CanonicalFinancialFact_ReportBenefit
ON CanonicalClaimsFinancialFact(canonical_claims_report_id, benefit);

CREATE TABLE IF NOT EXISTS CanonicalClaimsMonthlyFact (
    canonical_monthly_fact_id INTEGER PRIMARY KEY AUTOINCREMENT,
    canonical_claims_report_id INTEGER NOT NULL,
    fact_id TEXT NOT NULL UNIQUE,
    month TEXT, benefit TEXT NOT NULL DEFAULT '', coverage TEXT NOT NULL DEFAULT '', category TEXT NOT NULL DEFAULT '',
    division TEXT NOT NULL DEFAULT '', class TEXT NOT NULL DEFAULT '', measurement_type TEXT NOT NULL DEFAULT '',
    dimension_type TEXT NOT NULL DEFAULT '', dimension_value TEXT NOT NULL DEFAULT '',
    premium NUMERIC, paid_claims NUMERIC, incurred_claims NUMERIC, claims_plus_fees NUMERIC, loss_ratio NUMERIC, lives INTEGER,
    currency TEXT NOT NULL DEFAULT 'CAD', unit_of_measure TEXT NOT NULL DEFAULT '',
    availability_state TEXT NOT NULL, validation_status TEXT NOT NULL,
    legacy_source_record_id INTEGER, created_at TEXT NOT NULL, updated_at TEXT NOT NULL,
    FOREIGN KEY (canonical_claims_report_id) REFERENCES CanonicalClaimsReport(canonical_claims_report_id) ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS IX_CanonicalMonthlyFact_ReportMonthBenefit
ON CanonicalClaimsMonthlyFact(canonical_claims_report_id, month, benefit);

CREATE TABLE IF NOT EXISTS CanonicalClaimsCategoryFact (
    canonical_category_fact_id INTEGER PRIMARY KEY AUTOINCREMENT,
    canonical_claims_report_id INTEGER NOT NULL, fact_id TEXT NOT NULL UNIQUE,
    benefit TEXT NOT NULL DEFAULT '', coverage TEXT NOT NULL DEFAULT '', category TEXT NOT NULL DEFAULT '',
    division TEXT NOT NULL DEFAULT '', class TEXT NOT NULL DEFAULT '', measurement_type TEXT NOT NULL DEFAULT '',
    dimension_type TEXT NOT NULL DEFAULT '', dimension_value TEXT NOT NULL DEFAULT '',
    monetary_value NUMERIC, count_value INTEGER,
    percentage_value NUMERIC, availability_state TEXT NOT NULL, validation_status TEXT NOT NULL,
    currency TEXT NOT NULL DEFAULT '', unit_of_measure TEXT NOT NULL DEFAULT '',
    legacy_source_record_id INTEGER, created_at TEXT NOT NULL, updated_at TEXT NOT NULL,
    FOREIGN KEY (canonical_claims_report_id) REFERENCES CanonicalClaimsReport(canonical_claims_report_id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS CanonicalClaimsDistributionFact (
    canonical_distribution_fact_id INTEGER PRIMARY KEY AUTOINCREMENT,
    canonical_claims_report_id INTEGER NOT NULL, fact_id TEXT NOT NULL UNIQUE,
    benefit TEXT NOT NULL DEFAULT '', coverage TEXT NOT NULL DEFAULT '', category TEXT NOT NULL DEFAULT '',
    division TEXT NOT NULL DEFAULT '', class TEXT NOT NULL DEFAULT '', measurement_type TEXT NOT NULL DEFAULT '',
    dimension_type TEXT NOT NULL DEFAULT '',
    dimension_value TEXT NOT NULL DEFAULT '', monetary_value NUMERIC, count_value INTEGER, percentage_value NUMERIC,
    average_amount NUMERIC, availability_state TEXT NOT NULL, validation_status TEXT NOT NULL,
    currency TEXT NOT NULL DEFAULT '', unit_of_measure TEXT NOT NULL DEFAULT '',
    legacy_source_record_id INTEGER, created_at TEXT NOT NULL, updated_at TEXT NOT NULL,
    FOREIGN KEY (canonical_claims_report_id) REFERENCES CanonicalClaimsReport(canonical_claims_report_id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS CanonicalClaimsBenchmarkFact (
    canonical_benchmark_fact_id INTEGER PRIMARY KEY AUTOINCREMENT,
    canonical_claims_report_id INTEGER NOT NULL, fact_id TEXT NOT NULL UNIQUE,
    benefit TEXT NOT NULL DEFAULT '', coverage TEXT NOT NULL DEFAULT '', category TEXT NOT NULL DEFAULT '',
    division TEXT NOT NULL DEFAULT '', class TEXT NOT NULL DEFAULT '', measurement_type TEXT NOT NULL DEFAULT '',
    dimension_type TEXT NOT NULL DEFAULT '', dimension_value TEXT NOT NULL DEFAULT '', benchmark_definition TEXT NOT NULL DEFAULT '',
    monetary_value NUMERIC, count_value INTEGER, percentage_value NUMERIC,
    availability_state TEXT NOT NULL, validation_status TEXT NOT NULL,
    currency TEXT NOT NULL DEFAULT '', unit_of_measure TEXT NOT NULL DEFAULT '',
    legacy_source_record_id INTEGER, created_at TEXT NOT NULL, updated_at TEXT NOT NULL,
    FOREIGN KEY (canonical_claims_report_id) REFERENCES CanonicalClaimsReport(canonical_claims_report_id) ON DELETE CASCADE
);

-- Phase 2: report-centred service utilization. NUMERIC preserves decimal monetary values.
CREATE TABLE IF NOT EXISTS CanonicalClaimsUtilizationFact (
    canonical_claims_utilization_fact_id INTEGER PRIMARY KEY AUTOINCREMENT,
    canonical_claims_report_id INTEGER NOT NULL,
    fact_id TEXT NOT NULL UNIQUE,
    benefit_code TEXT NOT NULL DEFAULT '', benefit_name TEXT NOT NULL DEFAULT '',
    category_code TEXT NOT NULL DEFAULT '', category_name TEXT NOT NULL DEFAULT '',
    service_code TEXT NOT NULL DEFAULT '', service_name TEXT NOT NULL DEFAULT '',
    province_code TEXT NOT NULL DEFAULT '', relationship_type TEXT NOT NULL,
    division_code TEXT, class_code TEXT,
    unique_claimants INTEGER, occurrences INTEGER,
    submitted_amount NUMERIC, eligible_amount NUMERIC, paid_amount NUMERIC,
    reported_paid_percentage NUMERIC, currency TEXT NOT NULL DEFAULT 'CAD', validation_state TEXT NOT NULL DEFAULT 'Valid',
    measurement_start_date TEXT, measurement_end_date TEXT,
    source_aggregation_level TEXT NOT NULL, mapping_version TEXT NOT NULL,
    record_status TEXT NOT NULL, created_at_utc TEXT NOT NULL,
    FOREIGN KEY (canonical_claims_report_id) REFERENCES CanonicalClaimsReport(canonical_claims_report_id) ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS IX_CanonicalUtilization_Report ON CanonicalClaimsUtilizationFact(canonical_claims_report_id);
CREATE INDEX IF NOT EXISTS IX_CanonicalUtilization_Benefit ON CanonicalClaimsUtilizationFact(benefit_code);
CREATE INDEX IF NOT EXISTS IX_CanonicalUtilization_Category ON CanonicalClaimsUtilizationFact(category_code);
CREATE INDEX IF NOT EXISTS IX_CanonicalUtilization_Service ON CanonicalClaimsUtilizationFact(service_code);
CREATE INDEX IF NOT EXISTS IX_CanonicalUtilization_Province ON CanonicalClaimsUtilizationFact(province_code);
CREATE INDEX IF NOT EXISTS IX_CanonicalUtilization_Relationship ON CanonicalClaimsUtilizationFact(relationship_type);
CREATE INDEX IF NOT EXISTS IX_CanonicalUtilization_DivisionClass ON CanonicalClaimsUtilizationFact(division_code, class_code);
CREATE INDEX IF NOT EXISTS IX_CanonicalUtilization_Period ON CanonicalClaimsUtilizationFact(measurement_start_date, measurement_end_date);
CREATE INDEX IF NOT EXISTS IX_CanonicalUtilization_Query ON CanonicalClaimsUtilizationFact(canonical_claims_report_id, benefit_code, category_code, service_code, relationship_type, province_code);

CREATE TABLE IF NOT EXISTS CanonicalClaimsFactRevision (
    canonical_claims_fact_revision_id INTEGER PRIMARY KEY AUTOINCREMENT,
    canonical_claims_report_id INTEGER NOT NULL,
    revision_id TEXT NOT NULL UNIQUE,
    fact_id TEXT NOT NULL,
    natural_key TEXT NOT NULL,
    previous_source_reference TEXT NOT NULL DEFAULT '',
    revised_source_reference TEXT NOT NULL DEFAULT '',
    previous_submitted_amount NUMERIC, revised_submitted_amount NUMERIC,
    previous_eligible_amount NUMERIC, revised_eligible_amount NUMERIC,
    previous_paid_amount NUMERIC, revised_paid_amount NUMERIC,
    review_status TEXT NOT NULL DEFAULT 'NeedsReview',
    created_at_utc TEXT NOT NULL,
    FOREIGN KEY (canonical_claims_report_id) REFERENCES CanonicalClaimsReport(canonical_claims_report_id) ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS IX_CanonicalClaimsFactRevision_Report
ON CanonicalClaimsFactRevision(canonical_claims_report_id, fact_id);
CREATE TABLE IF NOT EXISTS CanonicalClaimsEvidence (
    canonical_evidence_id INTEGER PRIMARY KEY AUTOINCREMENT,
    canonical_claims_report_id INTEGER NOT NULL, fact_type TEXT NOT NULL, fact_id TEXT NOT NULL,
    source_document_id INTEGER, source_document_version_id INTEGER, import_batch_id INTEGER,
    source_page_number INTEGER, worksheet_name TEXT NOT NULL DEFAULT '', source_row INTEGER,
    source_column_or_cell TEXT NOT NULL DEFAULT '', source_table_or_section TEXT NOT NULL DEFAULT '',
    raw_carrier_label TEXT NOT NULL DEFAULT '', normalized_label TEXT NOT NULL DEFAULT '', raw_extracted_value TEXT NOT NULL DEFAULT '',
    extraction_method TEXT NOT NULL DEFAULT '', parser_name TEXT NOT NULL DEFAULT '', parser_version TEXT NOT NULL DEFAULT '',
    mapping_version TEXT NOT NULL DEFAULT '', confidence_score NUMERIC, evidence_precision TEXT NOT NULL,
    mapping_classification TEXT NOT NULL DEFAULT 'Direct', normalization_rule TEXT NOT NULL DEFAULT '',
    derivation_formula TEXT NOT NULL DEFAULT '', was_normalized INTEGER NOT NULL DEFAULT 0,
    warning_code TEXT NOT NULL DEFAULT '', warning_message TEXT NOT NULL DEFAULT '',
    FOREIGN KEY (canonical_claims_report_id) REFERENCES CanonicalClaimsReport(canonical_claims_report_id) ON DELETE CASCADE,
    FOREIGN KEY (source_document_id) REFERENCES Documents(document_id),
    FOREIGN KEY (import_batch_id) REFERENCES ImportBatch(import_batch_id)
);
CREATE UNIQUE INDEX IF NOT EXISTS UX_CanonicalClaimsEvidence_Fact
ON CanonicalClaimsEvidence(canonical_claims_report_id, fact_type, fact_id);

CREATE TABLE IF NOT EXISTS CanonicalClaimsValidation (
    canonical_validation_id INTEGER PRIMARY KEY AUTOINCREMENT,
    canonical_claims_report_id INTEGER NOT NULL, code TEXT NOT NULL, severity TEXT NOT NULL,
    field_name TEXT NOT NULL DEFAULT '', message TEXT NOT NULL DEFAULT '', fact_id TEXT,
    mapping_classification TEXT, source_reference TEXT NOT NULL DEFAULT '', normalization_rule TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL,
    FOREIGN KEY (canonical_claims_report_id) REFERENCES CanonicalClaimsReport(canonical_claims_report_id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS CanonicalClaimsParityResult (
    canonical_parity_result_id INTEGER PRIMARY KEY AUTOINCREMENT,
    canonical_claims_report_id INTEGER NOT NULL, metric TEXT NOT NULL, benefit TEXT NOT NULL DEFAULT '', period TEXT NOT NULL DEFAULT '',
    legacy_value NUMERIC, canonical_value NUMERIC, variance NUMERIC, tolerance NUMERIC NOT NULL,
    classification TEXT NOT NULL, explanation TEXT NOT NULL DEFAULT '', related_legacy_record_ids TEXT NOT NULL DEFAULT '',
    related_canonical_fact_ids TEXT NOT NULL DEFAULT '', review_status TEXT NOT NULL DEFAULT '', created_at TEXT NOT NULL,
    FOREIGN KEY (canonical_claims_report_id) REFERENCES CanonicalClaimsReport(canonical_claims_report_id) ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS IX_CanonicalClaimsParity_ReportClassification
ON CanonicalClaimsParityResult(canonical_claims_report_id, classification);
");
        await EnsureColumnAsync(connection, "CanonicalClaimsEvidence", "mapping_classification", "TEXT NOT NULL DEFAULT 'Direct'");
        await EnsureColumnAsync(connection, "CanonicalClaimsEvidence", "normalization_rule", "TEXT NOT NULL DEFAULT ''");
        await EnsureColumnAsync(connection, "CanonicalClaimsEvidence", "derivation_formula", "TEXT NOT NULL DEFAULT ''");
        await EnsureColumnAsync(connection, "CanonicalClaimsEvidence", "was_normalized", "INTEGER NOT NULL DEFAULT 0");
        await EnsureColumnAsync(connection, "CanonicalClaimsEvidence", "warning_code", "TEXT NOT NULL DEFAULT ''");
        await EnsureColumnAsync(connection, "CanonicalClaimsEvidence", "warning_message", "TEXT NOT NULL DEFAULT ''");
        await EnsureColumnAsync(connection, "CanonicalClaimsValidation", "mapping_classification", "TEXT");
        await EnsureColumnAsync(connection, "CanonicalClaimsValidation", "source_reference", "TEXT NOT NULL DEFAULT ''");
        await EnsureColumnAsync(connection, "CanonicalClaimsValidation", "normalization_rule", "TEXT NOT NULL DEFAULT ''");
        await EnsureColumnAsync(connection, "CanonicalClaimsUtilizationFact", "reported_paid_percentage", "NUMERIC");
        await EnsureColumnAsync(connection, "CanonicalClaimsUtilizationFact", "currency", "TEXT NOT NULL DEFAULT 'CAD'");
        await EnsureColumnAsync(connection, "CanonicalClaimsUtilizationFact", "validation_state", "TEXT NOT NULL DEFAULT 'Valid'");
        foreach (var (table, column, definition) in CarrierNeutralFactColumns())
            await EnsureColumnAsync(connection, table, column, definition);
    }

    private static async Task ExecuteAsync(SqliteConnection connection, string sql)
    {
        await using var command = connection.CreateCommand();
        command.CommandText = sql;
        await command.ExecuteNonQueryAsync();
    }

    private static async Task EnsureColumnAsync(SqliteConnection connection, string table, string column, string definition)
    {
        await using var check = connection.CreateCommand();
        check.CommandText = $"PRAGMA table_info({table});";
        await using var reader = await check.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            if (reader.GetString(1).Equals(column, StringComparison.OrdinalIgnoreCase)) return;
        }
        await reader.DisposeAsync();
        await ExecuteAsync(connection, $"ALTER TABLE {table} ADD COLUMN {column} {definition};");
    }

    private static IEnumerable<(string Table, string Column, string Definition)> CarrierNeutralFactColumns()
    {
        foreach (var table in new[] { "CanonicalClaimsCategoryFact", "CanonicalClaimsDistributionFact", "CanonicalClaimsBenchmarkFact" })
        {
            yield return (table, "coverage", "TEXT NOT NULL DEFAULT ''");
            yield return (table, "division", "TEXT NOT NULL DEFAULT ''");
            yield return (table, "class", "TEXT NOT NULL DEFAULT ''");
            yield return (table, "measurement_type", "TEXT NOT NULL DEFAULT ''");
            yield return (table, "dimension_type", "TEXT NOT NULL DEFAULT ''");
            yield return (table, "dimension_value", "TEXT NOT NULL DEFAULT ''");
            yield return (table, "currency", "TEXT NOT NULL DEFAULT ''");
            yield return (table, "unit_of_measure", "TEXT NOT NULL DEFAULT ''");
        }
    }
}


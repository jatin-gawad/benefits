using Microsoft.Data.Sqlite;

namespace RenewalReviewBuilder.App.Services;

public static class FinancialArrangementDatabaseMigration
{
    public static async Task ApplyAsync(SqliteConnection connection)
    {
        var commands = new[]
        {
            @"CREATE TABLE IF NOT EXISTS financial_arrangements (financial_arrangement_id INTEGER PRIMARY KEY AUTOINCREMENT, client_id INTEGER NOT NULL, carrier TEXT NOT NULL, effective_date TEXT NOT NULL, expiry_date TEXT, status TEXT NOT NULL, source TEXT NOT NULL, source_document_id INTEGER, target_loss_ratio REAL, tlr_basis TEXT NOT NULL DEFAULT '', target_loss_ratio_source_text TEXT NOT NULL DEFAULT '', target_loss_ratio_source_page INTEGER, target_loss_ratio_confidence REAL, created_at TEXT NOT NULL, superseded_at TEXT);",
            @"CREATE TABLE IF NOT EXISTS financial_arrangement_charges (financial_arrangement_charge_id INTEGER PRIMARY KEY AUTOINCREMENT, financial_arrangement_id INTEGER NOT NULL, charge_type TEXT NOT NULL, benefit TEXT NOT NULL, rate REAL, amount REAL, basis TEXT NOT NULL, source_text TEXT NOT NULL DEFAULT '', source_page INTEGER, extraction_confidence REAL, FOREIGN KEY(financial_arrangement_id) REFERENCES financial_arrangements(financial_arrangement_id));",
            @"CREATE TABLE IF NOT EXISTS pooling_arrangements (pooling_arrangement_id INTEGER PRIMARY KEY AUTOINCREMENT, client_id INTEGER NOT NULL, carrier TEXT NOT NULL, benefit TEXT NOT NULL, rate REAL, threshold REAL, basis TEXT NOT NULL, pooling_type TEXT NOT NULL DEFAULT '', effective_date TEXT NOT NULL, expiry_date TEXT, source_document_id INTEGER, source_text TEXT NOT NULL DEFAULT '', source_page INTEGER, extraction_confidence REAL, created_at TEXT NOT NULL, superseded_at TEXT);",
            @"CREATE TABLE IF NOT EXISTS commission_arrangements (commission_arrangement_id INTEGER PRIMARY KEY AUTOINCREMENT, client_id INTEGER NOT NULL, carrier TEXT NOT NULL, rate REAL, amount REAL, basis TEXT NOT NULL, source TEXT NOT NULL, confirmation_status TEXT NOT NULL, effective_date TEXT NOT NULL, expiry_date TEXT, sold_proposal_id INTEGER, source_document_id INTEGER, created_at TEXT NOT NULL, superseded_at TEXT);",
            @"CREATE TABLE IF NOT EXISTS financial_arrangement_versions (financial_arrangement_version_id INTEGER PRIMARY KEY AUTOINCREMENT, client_id INTEGER NOT NULL, carrier TEXT NOT NULL, scenario TEXT NOT NULL, effective_date TEXT NOT NULL, financial_arrangement_id INTEGER NOT NULL, commission_arrangement_id INTEGER, pooling_arrangement_id INTEGER, created_at TEXT NOT NULL, UNIQUE(client_id, carrier, scenario, effective_date, financial_arrangement_id, commission_arrangement_id, pooling_arrangement_id));",
            @"CREATE TABLE IF NOT EXISTS renewal_financial_arrangement_snapshots (renewal_financial_arrangement_snapshot_id INTEGER PRIMARY KEY AUTOINCREMENT, renewal_id INTEGER NOT NULL, client_id INTEGER NOT NULL, carrier TEXT NOT NULL, scenario TEXT NOT NULL, snapshot_type TEXT NOT NULL, snapshot_status TEXT NOT NULL, as_of_utc TEXT NOT NULL, current_arrangement_version_id INTEGER NOT NULL, proposed_arrangement_version_id INTEGER NOT NULL, source_version TEXT NOT NULL, is_stale INTEGER NOT NULL DEFAULT 0, created_at TEXT NOT NULL, finalized_at TEXT, UNIQUE(renewal_id, snapshot_type, source_version));",
            @"CREATE TABLE IF NOT EXISTS financial_arrangement_evidence (financial_arrangement_evidence_id INTEGER PRIMARY KEY AUTOINCREMENT, entity_type TEXT NOT NULL, entity_id INTEGER NOT NULL, source_document_id INTEGER, source_page INTEGER, source_text TEXT NOT NULL DEFAULT '', confidence REAL, created_at TEXT NOT NULL);",
            @"CREATE TABLE IF NOT EXISTS financial_arrangement_audit_events (financial_arrangement_audit_event_id INTEGER PRIMARY KEY AUTOINCREMENT, entity_type TEXT NOT NULL, entity_id INTEGER NOT NULL, client_id INTEGER NOT NULL, renewal_id INTEGER, event_type TEXT NOT NULL, previous_value TEXT NOT NULL DEFAULT '', new_value TEXT NOT NULL DEFAULT '', effective_date TEXT NOT NULL, user_name TEXT NOT NULL, reason TEXT NOT NULL DEFAULT '', source TEXT NOT NULL DEFAULT '', occurred_at_utc TEXT NOT NULL);",
            "CREATE INDEX IF NOT EXISTS IX_CommissionArrangements_Active ON commission_arrangements(client_id, carrier, effective_date, superseded_at);"
            ,"CREATE INDEX IF NOT EXISTS IX_FinancialArrangementAudit_Client ON financial_arrangement_audit_events(client_id, occurred_at_utc DESC);"
        };
        foreach (var sql in commands)
        {
            await using var command = connection.CreateCommand();
            command.CommandText = sql;
            await command.ExecuteNonQueryAsync();
        }
        // A previous schema can already contain this table without the columns
        // introduced by the effective-dated arrangement model. Add them before
        // creating the active-arrangement index.
        await EnsureColumnAsync(connection, "financial_arrangements", "carrier", "TEXT NOT NULL DEFAULT ''");
        await EnsureColumnAsync(connection, "financial_arrangements", "effective_date", "TEXT NOT NULL DEFAULT ''");
        await EnsureColumnAsync(connection, "financial_arrangements", "superseded_at", "TEXT");
        await EnsureColumnAsync(connection, "financial_arrangements", "target_loss_ratio", "REAL");
        await EnsureColumnAsync(connection, "financial_arrangements", "tlr_basis", "TEXT NOT NULL DEFAULT ''");
        await EnsureColumnAsync(connection, "financial_arrangements", "target_loss_ratio_source_text", "TEXT NOT NULL DEFAULT ''");
        await EnsureColumnAsync(connection, "financial_arrangements", "target_loss_ratio_source_page", "INTEGER");
        await EnsureColumnAsync(connection, "financial_arrangements", "target_loss_ratio_confidence", "REAL");

        // pooling_arrangements is an established canonical table. Extend that
        // table in place rather than introducing a second pooling store.
        await EnsureColumnAsync(connection, "pooling_arrangements", "carrier", "TEXT NOT NULL DEFAULT ''");
        await EnsureColumnAsync(connection, "pooling_arrangements", "rate", "REAL");
        await EnsureColumnAsync(connection, "pooling_arrangements", "threshold", "REAL");
        await EnsureColumnAsync(connection, "pooling_arrangements", "basis", "TEXT NOT NULL DEFAULT ''");
        await EnsureColumnAsync(connection, "pooling_arrangements", "pooling_type", "TEXT NOT NULL DEFAULT ''");
        await EnsureColumnAsync(connection, "pooling_arrangements", "effective_date", "TEXT NOT NULL DEFAULT ''");
        await EnsureColumnAsync(connection, "pooling_arrangements", "expiry_date", "TEXT");
        await EnsureColumnAsync(connection, "pooling_arrangements", "created_at", "TEXT NOT NULL DEFAULT ''");
        await EnsureColumnAsync(connection, "pooling_arrangements", "superseded_at", "TEXT");
        await EnsureColumnAsync(connection, "pooling_arrangements", "source_text", "TEXT NOT NULL DEFAULT ''");
        await EnsureColumnAsync(connection, "pooling_arrangements", "source_page", "INTEGER");
        await EnsureColumnAsync(connection, "pooling_arrangements", "extraction_confidence", "REAL");
        // SavePoolingArrangementAsync preserves the established legacy audit timestamp.
        // Add it for clean databases too, so the same write path is valid everywhere.
        await EnsureColumnAsync(connection, "pooling_arrangements", "date_loaded", "TEXT NOT NULL DEFAULT ''");
        if (await HasColumnAsync(connection, "pooling_arrangements", "pooling_charge") &&
            await HasColumnAsync(connection, "pooling_arrangements", "pooling_level"))
        {
            await using var mapLegacyPooling = connection.CreateCommand();
            mapLegacyPooling.CommandText = "UPDATE pooling_arrangements SET rate = COALESCE(rate, pooling_charge), threshold = COALESCE(threshold, pooling_level);";
            await mapLegacyPooling.ExecuteNonQueryAsync();
        }
        await using (var index = connection.CreateCommand())
        {
            index.CommandText = "CREATE INDEX IF NOT EXISTS IX_FinancialArrangements_Active ON financial_arrangements(client_id, carrier, effective_date, superseded_at);";
            await index.ExecuteNonQueryAsync();
        }
        await using (var index = connection.CreateCommand())
        {
            index.CommandText = "CREATE INDEX IF NOT EXISTS IX_PoolingArrangements_Active ON pooling_arrangements(client_id, carrier, benefit, effective_date, superseded_at);";
            await index.ExecuteNonQueryAsync();
        }
        await using (var index = connection.CreateCommand())
        {
            index.CommandText = "CREATE INDEX IF NOT EXISTS IX_FinancialArrangementVersions_Current ON financial_arrangement_versions(client_id, carrier, scenario, effective_date DESC);";
            await index.ExecuteNonQueryAsync();
        }
        await EnsureColumnAsync(connection, "financial_arrangement_charges", "formula", "TEXT NOT NULL DEFAULT ''");
        await EnsureColumnAsync(connection, "renewal_financial_arrangement_snapshots", "finalized_by", "TEXT");
    }

    private static async Task EnsureColumnAsync(SqliteConnection connection, string table, string column, string definition)
    {
        await using var command = connection.CreateCommand();
        command.CommandText = $"PRAGMA table_info({table});";
        await using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            if (string.Equals(reader.GetString(1), column, StringComparison.OrdinalIgnoreCase)) return;
        }
        await using var alter = connection.CreateCommand();
        alter.CommandText = $"ALTER TABLE {table} ADD COLUMN {column} {definition};";
        await alter.ExecuteNonQueryAsync();
    }

    private static async Task<bool> HasColumnAsync(SqliteConnection connection, string table, string column)
    {
        await using var command = connection.CreateCommand();
        command.CommandText = $"PRAGMA table_info({table});";
        await using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            if (string.Equals(reader.GetString(1), column, StringComparison.OrdinalIgnoreCase)) return true;
        }

        return false;
    }
}

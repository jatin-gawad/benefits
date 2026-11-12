using System.Globalization;
using Microsoft.Data.Sqlite;

namespace RenewalReviewBuilder.App.Services;

/// <summary>
/// Idempotent SQLite migration for the Carrier master reference table.
/// </summary>
public static class CarrierDatabaseMigration
{
    private static readonly CarrierSeed[] DefaultCarriers =
    [
        new("MFC", "Manulife", "MANULIFE_V1", "2025", 1, 1, 1, 0, 1, 1),
        new("CL", "Canada Life", "CANADALIFE_V1", "2025", 1, 1, 1, 1, 1, 1),
        new("SLF", "Sun Life", "SUNLIFE_V1", "2024", 1, 1, 1, 0, 1, 1),
        new("RBC", "RBC Insurance", "RBC_V1", "2024", 1, 1, 1, 0, 1, 1),
        new("EQ", "Equitable", "EQUITABLE_V1", "2024", 1, 1, 1, 0, 1, 1),
        new("ALL", "Allstate", "ALLSTATE_V1", "2026", 0, 0, 0, 0, 0, 0),
        new("IA", "iA Financial", "IA_V1", "2025", 1, 1, 1, 0, 1, 1),
        new("GH", "GroupHEALTH", "GROUPHEALTH_V1", "2024", 1, 1, 1, 1, 1, 1),
        new("PBC", "Pacific Blue Cross", "PBC_V1", "2026", 1, 1, 1, 0, 0, 0),
        new("DES", "Desjardins", "DESJARDINS_V1", "2024", 1, 1, 1, 0, 0, 0),
        new("NHC", "National HealthClaim", "NATIONALHEALTHCLAIM_V1", "2023", 0, 0, 0, 0, 0, 0)
    ];

    /// <summary>
    /// Creates or upgrades the Carrier master table, supporting indexes, and default carrier rows.
    /// </summary>
    public static async Task ApplyAsync(SqliteConnection connection)
    {
        try
        {
            await CreateCarrierTableAsync(connection);
            await EnsureCarrierColumnsAsync(connection);
            await NormalizeLegacyCarrierRowsAsync(connection);
            await EnsureCarrierIndexesAsync(connection);
            await SeedCarriersAsync(connection);
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            throw;
        }
    }

    private static async Task CreateCarrierTableAsync(SqliteConnection connection)
    {
        await ExecuteNonQueryAsync(connection, @"
CREATE TABLE IF NOT EXISTS Carrier (
    carrier_id INTEGER PRIMARY KEY AUTOINCREMENT,
    carrier_code TEXT NOT NULL,
    carrier_name TEXT NOT NULL COLLATE NOCASE,
    commission_statement_format TEXT NOT NULL,
    bonus_program_version TEXT,
    has_rate_guarantees INTEGER NOT NULL DEFAULT 0,
    has_persistency_bonus INTEGER NOT NULL DEFAULT 0,
    has_sales_bonus INTEGER NOT NULL DEFAULT 0,
    has_growth_bonus INTEGER NOT NULL DEFAULT 0,
    supports_billing_import INTEGER NOT NULL DEFAULT 0,
    supports_claims_import INTEGER NOT NULL DEFAULT 0,
    is_active INTEGER NOT NULL DEFAULT 1,
    created_date TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    modified_date TEXT,
    status TEXT NOT NULL DEFAULT 'Active',
    date_loaded TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    UNIQUE (carrier_code),
    UNIQUE (carrier_name)
);");
    }

    private static async Task EnsureCarrierColumnsAsync(SqliteConnection connection)
    {
        await EnsureColumnExistsAsync(connection, "Carrier", "carrier_code", "TEXT NOT NULL DEFAULT ''");
        await EnsureColumnExistsAsync(connection, "Carrier", "commission_statement_format", "TEXT NOT NULL DEFAULT ''");
        await EnsureColumnExistsAsync(connection, "Carrier", "bonus_program_version", "TEXT");
        await EnsureColumnExistsAsync(connection, "Carrier", "has_rate_guarantees", "INTEGER NOT NULL DEFAULT 0");
        await EnsureColumnExistsAsync(connection, "Carrier", "has_persistency_bonus", "INTEGER NOT NULL DEFAULT 0");
        await EnsureColumnExistsAsync(connection, "Carrier", "has_sales_bonus", "INTEGER NOT NULL DEFAULT 0");
        await EnsureColumnExistsAsync(connection, "Carrier", "has_growth_bonus", "INTEGER NOT NULL DEFAULT 0");
        await EnsureColumnExistsAsync(connection, "Carrier", "supports_billing_import", "INTEGER NOT NULL DEFAULT 0");
        await EnsureColumnExistsAsync(connection, "Carrier", "supports_claims_import", "INTEGER NOT NULL DEFAULT 0");
        await EnsureColumnExistsAsync(connection, "Carrier", "is_active", "INTEGER NOT NULL DEFAULT 1");
        await EnsureColumnExistsAsync(connection, "Carrier", "created_date", "TEXT NOT NULL DEFAULT ''");
        await EnsureColumnExistsAsync(connection, "Carrier", "modified_date", "TEXT");
        await EnsureColumnExistsAsync(connection, "Carrier", "status", "TEXT NOT NULL DEFAULT 'Active'");
        await EnsureColumnExistsAsync(connection, "Carrier", "date_loaded", "TEXT NOT NULL DEFAULT ''");
    }

    private static async Task NormalizeLegacyCarrierRowsAsync(SqliteConnection connection)
    {
        await ExecuteNonQueryAsync(connection, @"
UPDATE Carrier
SET carrier_code = 'LEGACY_' || carrier_id
WHERE TRIM(COALESCE(carrier_code, '')) = '';");

        await ExecuteNonQueryAsync(connection, @"
UPDATE Carrier
SET commission_statement_format = 'UNKNOWN'
WHERE TRIM(COALESCE(commission_statement_format, '')) = '';");

        await ExecuteNonQueryAsync(connection, @"
UPDATE Carrier
SET created_date = COALESCE(NULLIF(created_date, ''), COALESCE(NULLIF(date_loaded, ''), CURRENT_TIMESTAMP)),
    status = CASE WHEN is_active = 1 THEN 'Active' ELSE 'Inactive' END
WHERE TRIM(COALESCE(created_date, '')) = ''
   OR TRIM(COALESCE(status, '')) = '';");
    }

    private static async Task EnsureCarrierIndexesAsync(SqliteConnection connection)
    {
        await ExecuteNonQueryAsync(connection, @"
CREATE UNIQUE INDEX IF NOT EXISTS IDX_Carrier_Name
ON Carrier(carrier_name COLLATE NOCASE);");

        await ExecuteNonQueryAsync(connection, @"
CREATE UNIQUE INDEX IF NOT EXISTS IDX_Carrier_Code
ON Carrier(carrier_code COLLATE NOCASE);");

        await ExecuteNonQueryAsync(connection, @"
CREATE INDEX IF NOT EXISTS IDX_Carrier_Active
ON Carrier(is_active);");
    }

    private static async Task SeedCarriersAsync(SqliteConnection connection)
    {
        foreach (var carrier in DefaultCarriers)
        {
            await SeedCarrierAsync(connection, carrier);
        }
    }

    private static async Task SeedCarrierAsync(SqliteConnection connection, CarrierSeed carrier)
    {
        var carrierId = await FindCarrierIdAsync(connection, carrier.Code, carrier.Name);
        if (carrierId is null)
        {
            await InsertCarrierAsync(connection, carrier);
            return;
        }

        await UpdateCarrierSeedAsync(connection, carrierId.Value, carrier);
    }

    private static async Task<long?> FindCarrierIdAsync(SqliteConnection connection, string carrierCode, string carrierName)
    {
        await using var command = CreateCommand(connection, @"
SELECT carrier_id
FROM Carrier
WHERE carrier_code = $carrier_code COLLATE NOCASE
   OR carrier_name = $carrier_name COLLATE NOCASE
ORDER BY CASE WHEN carrier_code = $carrier_code COLLATE NOCASE THEN 0 ELSE 1 END,
         carrier_id
LIMIT 1;");
        AddParameter(command, "$carrier_code", carrierCode);
        AddParameter(command, "$carrier_name", carrierName);
        var value = await command.ExecuteScalarAsync();
        return value is null || value == DBNull.Value
            ? null
            : Convert.ToInt64(value, CultureInfo.InvariantCulture);
    }

    private static async Task InsertCarrierAsync(SqliteConnection connection, CarrierSeed carrier)
    {
        await using var command = CreateCommand(connection, @"
INSERT INTO Carrier (
    carrier_code,
    carrier_name,
    commission_statement_format,
    bonus_program_version,
    has_rate_guarantees,
    has_persistency_bonus,
    has_sales_bonus,
    has_growth_bonus,
    supports_billing_import,
    supports_claims_import,
    is_active,
    status,
    created_date
)
VALUES (
    $carrier_code,
    $carrier_name,
    $commission_statement_format,
    $bonus_program_version,
    $has_rate_guarantees,
    $has_persistency_bonus,
    $has_sales_bonus,
    $has_growth_bonus,
    $supports_billing_import,
    $supports_claims_import,
    1,
    'Active',
    CURRENT_TIMESTAMP
);");
        AddCarrierParameters(command, carrier);
        await command.ExecuteNonQueryAsync();
    }

    private static async Task UpdateCarrierSeedAsync(SqliteConnection connection, long carrierId, CarrierSeed carrier)
    {
        await using var command = CreateCommand(connection, @"
UPDATE Carrier
SET carrier_code = $carrier_code,
    carrier_name = $carrier_name,
    commission_statement_format = $commission_statement_format,
    bonus_program_version = $bonus_program_version,
    has_rate_guarantees = $has_rate_guarantees,
    has_persistency_bonus = $has_persistency_bonus,
    has_sales_bonus = $has_sales_bonus,
    has_growth_bonus = $has_growth_bonus,
    supports_billing_import = $supports_billing_import,
    supports_claims_import = $supports_claims_import,
    is_active = 1,
    status = 'Active',
    modified_date = CURRENT_TIMESTAMP
WHERE carrier_id = $carrier_id
  AND (
      carrier_code <> $carrier_code COLLATE NOCASE
      OR carrier_name <> $carrier_name COLLATE NOCASE
      OR commission_statement_format <> $commission_statement_format
      OR COALESCE(bonus_program_version, '') <> COALESCE($bonus_program_version, '')
      OR has_rate_guarantees <> $has_rate_guarantees
      OR has_persistency_bonus <> $has_persistency_bonus
      OR has_sales_bonus <> $has_sales_bonus
      OR has_growth_bonus <> $has_growth_bonus
      OR supports_billing_import <> $supports_billing_import
      OR supports_claims_import <> $supports_claims_import
      OR is_active <> 1
      OR status <> 'Active'
  );");
        AddCarrierParameters(command, carrier);
        AddParameter(command, "$carrier_id", carrierId);
        await command.ExecuteNonQueryAsync();
    }

    private static async Task EnsureColumnExistsAsync(
        SqliteConnection connection,
        string tableName,
        string columnName,
        string columnDefinition)
    {
        await using var command = CreateCommand(connection, $"PRAGMA table_info({tableName});");
        await using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            var existingColumn = reader.GetString(reader.GetOrdinal("name"));
            if (existingColumn.Equals(columnName, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }
        }

        await ExecuteNonQueryAsync(connection, $"ALTER TABLE {tableName} ADD COLUMN {columnName} {columnDefinition};");
    }

    private static async Task ExecuteNonQueryAsync(SqliteConnection connection, string sql)
    {
        await using var command = CreateCommand(connection, sql);
        await command.ExecuteNonQueryAsync();
    }

    private static SqliteCommand CreateCommand(SqliteConnection connection, string sql)
    {
        var command = connection.CreateCommand();
        command.CommandText = sql;
        return command;
    }

    private static void AddCarrierParameters(SqliteCommand command, CarrierSeed carrier)
    {
        AddParameter(command, "$carrier_code", carrier.Code);
        AddParameter(command, "$carrier_name", carrier.Name);
        AddParameter(command, "$commission_statement_format", carrier.Parser);
        AddParameter(command, "$bonus_program_version", carrier.BonusVersion);
        AddParameter(command, "$has_rate_guarantees", carrier.RateGuarantees);
        AddParameter(command, "$has_persistency_bonus", carrier.Persistency);
        AddParameter(command, "$has_sales_bonus", carrier.Sales);
        AddParameter(command, "$has_growth_bonus", carrier.Growth);
        AddParameter(command, "$supports_billing_import", carrier.Billing);
        AddParameter(command, "$supports_claims_import", carrier.Claims);
    }

    private static void AddParameter(SqliteCommand command, string name, object? value)
    {
        command.Parameters.AddWithValue(name, value ?? DBNull.Value);
    }

    private sealed record CarrierSeed(
        string Code,
        string Name,
        string Parser,
        string? BonusVersion,
        int RateGuarantees,
        int Persistency,
        int Sales,
        int Growth,
        int Billing,
        int Claims);
}

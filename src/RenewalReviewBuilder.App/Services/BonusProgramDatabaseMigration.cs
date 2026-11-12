using System.Globalization;
using Microsoft.Data.Sqlite;

namespace RenewalReviewBuilder.App.Services;

/// <summary>
/// Idempotent SQLite migration for Commission Intelligence bonus program configuration.
/// </summary>
public static class BonusProgramDatabaseMigration
{
    private static readonly BonusProgramSeed[] DefaultBonusPrograms =
    [
        new("Manulife", 2025, "Manulife 2025 Bonus Program", 1, 1, 0, 0, 0, 0, 1),
        new("Sun Life", 2024, "Sun Life 2024 Bonus Program", 1, 1, 0, 0, 0, 0, 1),
        new("Canada Life", 2025, "Canada Life 2025 Bonus Program", 1, 1, 1, 0, 0, 0, 1),
        new("RBC Insurance", 2024, "RBC Insurance 2024 Bonus Program", 1, 1, 0, 0, 0, 0, 1),
        new("Equitable", 2024, "Equitable 2024 Bonus Program", 1, 1, 0, 0, 0, 0, 1),
        new("iA Financial", 2025, "iA Financial 2025 Bonus Program", 1, 1, 0, 0, 0, 0, 1),
        new("GroupHEALTH", 2024, "GroupHEALTH 2024 Bonus Program", 1, 1, 1, 0, 0, 0, 1),
        new("Pacific Blue Cross", 2026, "Pacific Blue Cross 2026 Bonus Program", 1, 1, 0, 0, 0, 0, 1),
        new("Desjardins", 2024, "Desjardins 2024 Bonus Program", 1, 1, 0, 0, 0, 0, 1)
    ];

    /// <summary>
    /// Creates or upgrades the BonusProgram table, supporting indexes, and default program rows.
    /// </summary>
    public static async Task ApplyAsync(SqliteConnection connection)
    {
        try
        {
            await CreateBonusProgramTableAsync(connection);
            await EnsureBonusProgramColumnsAsync(connection);
            await NormalizeExistingBonusProgramRowsAsync(connection);
            await EnsureBonusProgramIndexesAsync(connection);
            await SeedBonusProgramsAsync(connection);
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            throw;
        }
    }

    private static async Task CreateBonusProgramTableAsync(SqliteConnection connection)
    {
        await ExecuteNonQueryAsync(connection, @"
CREATE TABLE IF NOT EXISTS BonusProgram (
    BonusProgramID INTEGER PRIMARY KEY AUTOINCREMENT,
    CarrierID INTEGER NOT NULL,
    ProgramName TEXT NOT NULL,
    ProgramYear INTEGER NOT NULL,
    QualificationStartDate TEXT NOT NULL,
    QualificationEndDate TEXT NOT NULL,
    PaymentFrequency TEXT NULL,
    HasSalesBonus INTEGER NOT NULL DEFAULT 0,
    HasPersistencyBonus INTEGER NOT NULL DEFAULT 0,
    HasGrowthBonus INTEGER NOT NULL DEFAULT 0,
    HasRetentionBonus INTEGER NOT NULL DEFAULT 0,
    HasEliteBonus INTEGER NOT NULL DEFAULT 0,
    HasSalesEnhancement INTEGER NOT NULL DEFAULT 0,
    HasRateGuaranteeImpact INTEGER NOT NULL DEFAULT 0,
    SourceDocument TEXT NULL,
    Notes TEXT NULL,
    IsActive INTEGER NOT NULL DEFAULT 1,
    CreatedDate TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    ModifiedDate TEXT NULL,
    FOREIGN KEY (CarrierID) REFERENCES Carrier(carrier_id)
);");
    }

    private static async Task EnsureBonusProgramColumnsAsync(SqliteConnection connection)
    {
        await EnsureColumnExistsAsync(connection, "BonusProgram", "CarrierID", "INTEGER NOT NULL DEFAULT 0");
        await EnsureColumnExistsAsync(connection, "BonusProgram", "ProgramName", "TEXT NOT NULL DEFAULT ''");
        await EnsureColumnExistsAsync(connection, "BonusProgram", "ProgramYear", "INTEGER NOT NULL DEFAULT 0");
        await EnsureColumnExistsAsync(connection, "BonusProgram", "QualificationStartDate", "TEXT NOT NULL DEFAULT ''");
        await EnsureColumnExistsAsync(connection, "BonusProgram", "QualificationEndDate", "TEXT NOT NULL DEFAULT ''");
        await EnsureColumnExistsAsync(connection, "BonusProgram", "PaymentFrequency", "TEXT NULL");
        await EnsureColumnExistsAsync(connection, "BonusProgram", "HasSalesBonus", "INTEGER NOT NULL DEFAULT 0");
        await EnsureColumnExistsAsync(connection, "BonusProgram", "HasPersistencyBonus", "INTEGER NOT NULL DEFAULT 0");
        await EnsureColumnExistsAsync(connection, "BonusProgram", "HasGrowthBonus", "INTEGER NOT NULL DEFAULT 0");
        await EnsureColumnExistsAsync(connection, "BonusProgram", "HasRetentionBonus", "INTEGER NOT NULL DEFAULT 0");
        await EnsureColumnExistsAsync(connection, "BonusProgram", "HasEliteBonus", "INTEGER NOT NULL DEFAULT 0");
        await EnsureColumnExistsAsync(connection, "BonusProgram", "HasSalesEnhancement", "INTEGER NOT NULL DEFAULT 0");
        await EnsureColumnExistsAsync(connection, "BonusProgram", "HasRateGuaranteeImpact", "INTEGER NOT NULL DEFAULT 0");
        await EnsureColumnExistsAsync(connection, "BonusProgram", "SourceDocument", "TEXT NULL");
        await EnsureColumnExistsAsync(connection, "BonusProgram", "Notes", "TEXT NULL");
        await EnsureColumnExistsAsync(connection, "BonusProgram", "IsActive", "INTEGER NOT NULL DEFAULT 1");
        await EnsureColumnExistsAsync(connection, "BonusProgram", "CreatedDate", "TEXT NOT NULL DEFAULT ''");
        await EnsureColumnExistsAsync(connection, "BonusProgram", "ModifiedDate", "TEXT NULL");
    }

    private static async Task NormalizeExistingBonusProgramRowsAsync(SqliteConnection connection)
    {
        await ExecuteNonQueryAsync(connection, @"
UPDATE BonusProgram
SET ProgramName = COALESCE(NULLIF(TRIM(ProgramName), ''), 'Unknown Bonus Program'),
    QualificationStartDate = COALESCE(NULLIF(TRIM(QualificationStartDate), ''), printf('%04d-01-01', ProgramYear)),
    QualificationEndDate = COALESCE(NULLIF(TRIM(QualificationEndDate), ''), printf('%04d-12-31', ProgramYear)),
    CreatedDate = COALESCE(NULLIF(TRIM(CreatedDate), ''), CURRENT_TIMESTAMP)
WHERE TRIM(COALESCE(ProgramName, '')) = ''
   OR TRIM(COALESCE(QualificationStartDate, '')) = ''
   OR TRIM(COALESCE(QualificationEndDate, '')) = ''
   OR TRIM(COALESCE(CreatedDate, '')) = '';");
    }

    private static async Task EnsureBonusProgramIndexesAsync(SqliteConnection connection)
    {
        await ExecuteNonQueryAsync(connection, @"
CREATE INDEX IF NOT EXISTS IDX_BonusProgram_CarrierID
ON BonusProgram(CarrierID);");

        await ExecuteNonQueryAsync(connection, @"
CREATE INDEX IF NOT EXISTS IDX_BonusProgram_ProgramYear
ON BonusProgram(ProgramYear);");

        await ExecuteNonQueryAsync(connection, @"
CREATE INDEX IF NOT EXISTS IDX_BonusProgram_Active
ON BonusProgram(IsActive);");
    }

    private static async Task SeedBonusProgramsAsync(SqliteConnection connection)
    {
        foreach (var bonusProgram in DefaultBonusPrograms)
        {
            var carrierId = await FindCarrierIdByNameAsync(connection, bonusProgram.CarrierName);
            if (carrierId is null)
            {
                continue;
            }

            await SeedBonusProgramAsync(connection, carrierId.Value, bonusProgram);
        }
    }

    private static async Task SeedBonusProgramAsync(SqliteConnection connection, long carrierId, BonusProgramSeed bonusProgram)
    {
        var existingProgramId = await FindBonusProgramIdAsync(connection, carrierId, bonusProgram.ProgramYear);
        if (existingProgramId is null)
        {
            await InsertBonusProgramSeedAsync(connection, carrierId, bonusProgram);
            return;
        }

        await UpdateBonusProgramSeedAsync(connection, existingProgramId.Value, carrierId, bonusProgram);
    }

    private static async Task<long?> FindCarrierIdByNameAsync(SqliteConnection connection, string carrierName)
    {
        await using var command = CreateCommand(connection, @"
SELECT carrier_id
FROM Carrier
WHERE carrier_name = $carrier_name COLLATE NOCASE
LIMIT 1;");
        AddParameter(command, "$carrier_name", carrierName);
        return ToNullableLong(await command.ExecuteScalarAsync());
    }

    private static async Task<long?> FindBonusProgramIdAsync(SqliteConnection connection, long carrierId, int programYear)
    {
        await using var command = CreateCommand(connection, @"
SELECT BonusProgramID
FROM BonusProgram
WHERE CarrierID = $carrier_id
  AND ProgramYear = $program_year
ORDER BY BonusProgramID
LIMIT 1;");
        AddParameter(command, "$carrier_id", carrierId);
        AddParameter(command, "$program_year", programYear);
        return ToNullableLong(await command.ExecuteScalarAsync());
    }

    private static async Task InsertBonusProgramSeedAsync(
        SqliteConnection connection,
        long carrierId,
        BonusProgramSeed bonusProgram)
    {
        await using var command = CreateCommand(connection, @"
INSERT INTO BonusProgram (
    CarrierID,
    ProgramName,
    ProgramYear,
    QualificationStartDate,
    QualificationEndDate,
    PaymentFrequency,
    HasSalesBonus,
    HasPersistencyBonus,
    HasGrowthBonus,
    HasRetentionBonus,
    HasEliteBonus,
    HasSalesEnhancement,
    HasRateGuaranteeImpact,
    SourceDocument,
    Notes,
    IsActive,
    CreatedDate
)
VALUES (
    $carrier_id,
    $program_name,
    $program_year,
    $qualification_start_date,
    $qualification_end_date,
    NULL,
    $has_sales_bonus,
    $has_persistency_bonus,
    $has_growth_bonus,
    $has_retention_bonus,
    $has_elite_bonus,
    $has_sales_enhancement,
    $has_rate_guarantee_impact,
    NULL,
    $notes,
    1,
    CURRENT_TIMESTAMP
);");
        AddBonusProgramParameters(command, carrierId, bonusProgram);
        await command.ExecuteNonQueryAsync();
    }

    private static async Task UpdateBonusProgramSeedAsync(
        SqliteConnection connection,
        long bonusProgramId,
        long carrierId,
        BonusProgramSeed bonusProgram)
    {
        await using var command = CreateCommand(connection, @"
UPDATE BonusProgram
SET CarrierID = $carrier_id,
    ProgramName = $program_name,
    ProgramYear = $program_year,
    QualificationStartDate = $qualification_start_date,
    QualificationEndDate = $qualification_end_date,
    HasSalesBonus = $has_sales_bonus,
    HasPersistencyBonus = $has_persistency_bonus,
    HasGrowthBonus = $has_growth_bonus,
    HasRetentionBonus = $has_retention_bonus,
    HasEliteBonus = $has_elite_bonus,
    HasSalesEnhancement = $has_sales_enhancement,
    HasRateGuaranteeImpact = $has_rate_guarantee_impact,
    IsActive = 1,
    ModifiedDate = CURRENT_TIMESTAMP
WHERE BonusProgramID = $bonus_program_id
  AND (
      CarrierID <> $carrier_id
      OR ProgramName <> $program_name
      OR ProgramYear <> $program_year
      OR QualificationStartDate <> $qualification_start_date
      OR QualificationEndDate <> $qualification_end_date
      OR HasSalesBonus <> $has_sales_bonus
      OR HasPersistencyBonus <> $has_persistency_bonus
      OR HasGrowthBonus <> $has_growth_bonus
      OR HasRetentionBonus <> $has_retention_bonus
      OR HasEliteBonus <> $has_elite_bonus
      OR HasSalesEnhancement <> $has_sales_enhancement
      OR HasRateGuaranteeImpact <> $has_rate_guarantee_impact
      OR IsActive <> 1
  );");
        AddBonusProgramParameters(command, carrierId, bonusProgram);
        AddParameter(command, "$bonus_program_id", bonusProgramId);
        await command.ExecuteNonQueryAsync();
    }

    private static void AddBonusProgramParameters(SqliteCommand command, long carrierId, BonusProgramSeed bonusProgram)
    {
        AddParameter(command, "$carrier_id", carrierId);
        AddParameter(command, "$program_name", bonusProgram.ProgramName);
        AddParameter(command, "$program_year", bonusProgram.ProgramYear);
        AddParameter(command, "$qualification_start_date", $"{bonusProgram.ProgramYear:0000}-01-01");
        AddParameter(command, "$qualification_end_date", $"{bonusProgram.ProgramYear:0000}-12-31");
        AddParameter(command, "$has_sales_bonus", bonusProgram.HasSalesBonus);
        AddParameter(command, "$has_persistency_bonus", bonusProgram.HasPersistencyBonus);
        AddParameter(command, "$has_growth_bonus", bonusProgram.HasGrowthBonus);
        AddParameter(command, "$has_retention_bonus", bonusProgram.HasRetentionBonus);
        AddParameter(command, "$has_elite_bonus", bonusProgram.HasEliteBonus);
        AddParameter(command, "$has_sales_enhancement", bonusProgram.HasSalesEnhancement);
        AddParameter(command, "$has_rate_guarantee_impact", bonusProgram.HasRateGuaranteeImpact);
        AddParameter(command, "$notes", "Seeded for Commission Intelligence configuration. Detailed rules are pending source-document review.");
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

    private static void AddParameter(SqliteCommand command, string name, object? value)
    {
        command.Parameters.AddWithValue(name, value ?? DBNull.Value);
    }

    private static long? ToNullableLong(object? value)
    {
        return value is null || value == DBNull.Value
            ? null
            : Convert.ToInt64(value, CultureInfo.InvariantCulture);
    }

    private sealed record BonusProgramSeed(
        string CarrierName,
        int ProgramYear,
        string ProgramName,
        int HasSalesBonus,
        int HasPersistencyBonus,
        int HasGrowthBonus,
        int HasRetentionBonus,
        int HasEliteBonus,
        int HasSalesEnhancement,
        int HasRateGuaranteeImpact);
}

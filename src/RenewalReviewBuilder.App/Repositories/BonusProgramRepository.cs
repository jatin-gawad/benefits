using System.Globalization;
using Microsoft.Data.Sqlite;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;

namespace RenewalReviewBuilder.App.Repositories;

/// <summary>
/// SQLite repository for Commission Intelligence bonus program configuration.
/// </summary>
public sealed class BonusProgramRepository
{
    private readonly string _databasePath;

    /// <summary>
    /// Creates a repository that uses the shared Benefits Intelligence database unless a test path is supplied.
    /// </summary>
    public BonusProgramRepository(string? databasePath = null)
    {
        _databasePath = string.IsNullOrWhiteSpace(databasePath)
            ? AppPaths.BenefitsIntelligenceDatabasePath
            : databasePath;
    }

    /// <summary>
    /// Gets every bonus program ordered by carrier and program year.
    /// </summary>
    public async Task<List<BonusProgramRecord>> GetAll()
    {
        return await QueryBonusProgramsAsync(BaseSelectSql + @"
ORDER BY c.carrier_name COLLATE NOCASE, bp.ProgramYear DESC, bp.ProgramName COLLATE NOCASE;");
    }

    /// <summary>
    /// Gets active bonus programs ordered by carrier and program year.
    /// </summary>
    public async Task<List<BonusProgramRecord>> GetActive()
    {
        return await QueryBonusProgramsAsync(BaseSelectSql + @"
WHERE bp.IsActive = 1
ORDER BY c.carrier_name COLLATE NOCASE, bp.ProgramYear DESC, bp.ProgramName COLLATE NOCASE;");
    }

    /// <summary>
    /// Gets bonus programs for one carrier.
    /// </summary>
    public async Task<List<BonusProgramRecord>> GetByCarrierId(long carrierId)
    {
        return await QueryBonusProgramsAsync(
            BaseSelectSql + @"
WHERE bp.CarrierID = $carrier_id
ORDER BY bp.ProgramYear DESC, bp.ProgramName COLLATE NOCASE;",
            command => AddParameter(command, "$carrier_id", carrierId));
    }

    /// <summary>
    /// Gets one bonus program for a carrier and program year.
    /// </summary>
    public async Task<BonusProgramRecord?> GetByCarrierAndYear(long carrierId, int programYear)
    {
        return await QuerySingleBonusProgramAsync(
            BaseSelectSql + @"
WHERE bp.CarrierID = $carrier_id
  AND bp.ProgramYear = $program_year
ORDER BY bp.BonusProgramID
LIMIT 1;",
            command =>
            {
                AddParameter(command, "$carrier_id", carrierId);
                AddParameter(command, "$program_year", programYear);
            });
    }

    /// <summary>
    /// Inserts a new bonus program after validating required fields.
    /// </summary>
    public async Task<long> Insert(BonusProgramRecord bonusProgram)
    {
        try
        {
            Normalize(bonusProgram);
            Validate(bonusProgram);
            if (await Exists(bonusProgram.CarrierId, bonusProgram.ProgramYear, bonusProgram.ProgramName))
            {
                throw new InvalidOperationException("A bonus program already exists for this carrier, year, and program name.");
            }

            await using var connection = await OpenConnectionAsync();
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
    $payment_frequency,
    $has_sales_bonus,
    $has_persistency_bonus,
    $has_growth_bonus,
    $has_retention_bonus,
    $has_elite_bonus,
    $has_sales_enhancement,
    $has_rate_guarantee_impact,
    $source_document,
    $notes,
    $is_active,
    CURRENT_TIMESTAMP
);
SELECT last_insert_rowid();");
            AddBonusProgramParameters(command, bonusProgram);
            var id = Convert.ToInt64(await command.ExecuteScalarAsync(), CultureInfo.InvariantCulture);
            bonusProgram.BonusProgramId = id;
            AppLogger.Info($"Inserted bonus program {bonusProgram.ProgramName} ({bonusProgram.ProgramYear}).");
            return id;
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            throw;
        }
    }

    /// <summary>
    /// Updates an existing bonus program after validating required fields.
    /// </summary>
    public async Task<bool> Update(BonusProgramRecord bonusProgram)
    {
        try
        {
            if (bonusProgram.BonusProgramId <= 0)
            {
                throw new ArgumentException("BonusProgramId is required for updates.", nameof(bonusProgram));
            }

            Normalize(bonusProgram);
            Validate(bonusProgram);
            if (await Exists(
                    bonusProgram.CarrierId,
                    bonusProgram.ProgramYear,
                    bonusProgram.ProgramName,
                    bonusProgram.BonusProgramId))
            {
                throw new InvalidOperationException("A bonus program already exists for this carrier, year, and program name.");
            }

            await using var connection = await OpenConnectionAsync();
            await using var command = CreateCommand(connection, @"
UPDATE BonusProgram
SET CarrierID = $carrier_id,
    ProgramName = $program_name,
    ProgramYear = $program_year,
    QualificationStartDate = $qualification_start_date,
    QualificationEndDate = $qualification_end_date,
    PaymentFrequency = $payment_frequency,
    HasSalesBonus = $has_sales_bonus,
    HasPersistencyBonus = $has_persistency_bonus,
    HasGrowthBonus = $has_growth_bonus,
    HasRetentionBonus = $has_retention_bonus,
    HasEliteBonus = $has_elite_bonus,
    HasSalesEnhancement = $has_sales_enhancement,
    HasRateGuaranteeImpact = $has_rate_guarantee_impact,
    SourceDocument = $source_document,
    Notes = $notes,
    IsActive = $is_active,
    ModifiedDate = CURRENT_TIMESTAMP
WHERE BonusProgramID = $bonus_program_id;");
            AddBonusProgramParameters(command, bonusProgram);
            AddParameter(command, "$bonus_program_id", bonusProgram.BonusProgramId);
            var updated = await command.ExecuteNonQueryAsync();
            if (updated > 0)
            {
                AppLogger.Info($"Updated bonus program {bonusProgram.BonusProgramId}.");
            }

            return updated > 0;
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            throw;
        }
    }

    /// <summary>
    /// Marks a bonus program inactive without deleting the record.
    /// </summary>
    public async Task<bool> Deactivate(long bonusProgramId)
    {
        try
        {
            await using var connection = await OpenConnectionAsync();
            await using var command = CreateCommand(connection, @"
UPDATE BonusProgram
SET IsActive = 0,
    ModifiedDate = CURRENT_TIMESTAMP
WHERE BonusProgramID = $bonus_program_id;");
            AddParameter(command, "$bonus_program_id", bonusProgramId);
            var updated = await command.ExecuteNonQueryAsync();
            if (updated > 0)
            {
                AppLogger.Info($"Deactivated bonus program {bonusProgramId}.");
            }

            return updated > 0;
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            throw;
        }
    }

    /// <summary>
    /// Checks whether a bonus program already exists for a carrier, year, and program name.
    /// </summary>
    public async Task<bool> Exists(
        long carrierId,
        int programYear,
        string programName = "",
        long? excludingBonusProgramId = null)
    {
        try
        {
            await using var connection = await OpenConnectionAsync();
            await using var command = CreateCommand(connection, @"
SELECT COUNT(*)
FROM BonusProgram
WHERE CarrierID = $carrier_id
  AND ProgramYear = $program_year
  AND (
      $program_name = ''
      OR ProgramName = $program_name COLLATE NOCASE
  )
  AND ($excluding_bonus_program_id IS NULL OR BonusProgramID <> $excluding_bonus_program_id);");
            AddParameter(command, "$carrier_id", carrierId);
            AddParameter(command, "$program_year", programYear);
            AddParameter(command, "$program_name", Clean(programName));
            AddParameter(command, "$excluding_bonus_program_id", excludingBonusProgramId);
            return Convert.ToInt32(await command.ExecuteScalarAsync(), CultureInfo.InvariantCulture) > 0;
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            throw;
        }
    }

    private async Task<List<BonusProgramRecord>> QueryBonusProgramsAsync(
        string sql,
        Action<SqliteCommand>? addParameters = null)
    {
        try
        {
            await using var connection = await OpenConnectionAsync();
            await using var command = CreateCommand(connection, sql);
            addParameters?.Invoke(command);
            var programs = new List<BonusProgramRecord>();
            await using var reader = await command.ExecuteReaderAsync();
            while (await reader.ReadAsync())
            {
                programs.Add(ReadBonusProgram(reader));
            }

            return programs;
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            throw;
        }
    }

    private async Task<BonusProgramRecord?> QuerySingleBonusProgramAsync(
        string sql,
        Action<SqliteCommand> addParameters)
    {
        try
        {
            await using var connection = await OpenConnectionAsync();
            await using var command = CreateCommand(connection, sql);
            addParameters(command);
            await using var reader = await command.ExecuteReaderAsync();
            return await reader.ReadAsync() ? ReadBonusProgram(reader) : null;
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            throw;
        }
    }

    private async Task<SqliteConnection> OpenConnectionAsync()
    {
        var initialization = new DatabaseService(_databasePath).Initialize();
        if (!initialization.Succeeded)
        {
            throw new InvalidOperationException(initialization.Message);
        }

        var folder = Path.GetDirectoryName(_databasePath);
        if (!string.IsNullOrWhiteSpace(folder))
        {
            Directory.CreateDirectory(folder);
        }

        var connection = new SqliteConnection($"Data Source={_databasePath}");
        await connection.OpenAsync();
        await using var command = CreateCommand(connection, "PRAGMA foreign_keys = ON;");
        await command.ExecuteNonQueryAsync();
        return connection;
    }

    private static BonusProgramRecord ReadBonusProgram(SqliteDataReader reader)
    {
        return new BonusProgramRecord
        {
            BonusProgramId = ReadLong(reader, "BonusProgramID"),
            CarrierId = ReadLong(reader, "CarrierID"),
            CarrierName = ReadString(reader, "CarrierName"),
            ProgramName = ReadString(reader, "ProgramName"),
            ProgramYear = ReadInt(reader, "ProgramYear") ?? 0,
            QualificationStartDate = ReadDateOnly(reader, "QualificationStartDate") ?? DateOnly.MinValue,
            QualificationEndDate = ReadDateOnly(reader, "QualificationEndDate") ?? DateOnly.MinValue,
            PaymentFrequency = ReadNullableString(reader, "PaymentFrequency"),
            HasSalesBonus = ReadBoolean(reader, "HasSalesBonus"),
            HasPersistencyBonus = ReadBoolean(reader, "HasPersistencyBonus"),
            HasGrowthBonus = ReadBoolean(reader, "HasGrowthBonus"),
            HasRetentionBonus = ReadBoolean(reader, "HasRetentionBonus"),
            HasEliteBonus = ReadBoolean(reader, "HasEliteBonus"),
            HasSalesEnhancement = ReadBoolean(reader, "HasSalesEnhancement"),
            HasRateGuaranteeImpact = ReadBoolean(reader, "HasRateGuaranteeImpact"),
            SourceDocument = ReadNullableString(reader, "SourceDocument"),
            Notes = ReadNullableString(reader, "Notes"),
            IsActive = ReadBoolean(reader, "IsActive"),
            CreatedDate = ReadDateTime(reader, "CreatedDate") ?? DateTime.MinValue,
            ModifiedDate = ReadDateTime(reader, "ModifiedDate")
        };
    }

    private static void AddBonusProgramParameters(SqliteCommand command, BonusProgramRecord bonusProgram)
    {
        AddParameter(command, "$carrier_id", bonusProgram.CarrierId);
        AddParameter(command, "$program_name", bonusProgram.ProgramName);
        AddParameter(command, "$program_year", bonusProgram.ProgramYear);
        AddParameter(command, "$qualification_start_date", ToDateText(bonusProgram.QualificationStartDate));
        AddParameter(command, "$qualification_end_date", ToDateText(bonusProgram.QualificationEndDate));
        AddParameter(command, "$payment_frequency", OptionalText(bonusProgram.PaymentFrequency));
        AddParameter(command, "$has_sales_bonus", ToInteger(bonusProgram.HasSalesBonus));
        AddParameter(command, "$has_persistency_bonus", ToInteger(bonusProgram.HasPersistencyBonus));
        AddParameter(command, "$has_growth_bonus", ToInteger(bonusProgram.HasGrowthBonus));
        AddParameter(command, "$has_retention_bonus", ToInteger(bonusProgram.HasRetentionBonus));
        AddParameter(command, "$has_elite_bonus", ToInteger(bonusProgram.HasEliteBonus));
        AddParameter(command, "$has_sales_enhancement", ToInteger(bonusProgram.HasSalesEnhancement));
        AddParameter(command, "$has_rate_guarantee_impact", ToInteger(bonusProgram.HasRateGuaranteeImpact));
        AddParameter(command, "$source_document", OptionalText(bonusProgram.SourceDocument));
        AddParameter(command, "$notes", OptionalText(bonusProgram.Notes));
        AddParameter(command, "$is_active", ToInteger(bonusProgram.IsActive));
    }

    private static void Validate(BonusProgramRecord bonusProgram)
    {
        if (bonusProgram.CarrierId <= 0)
        {
            throw new ArgumentException("CarrierID is required.", nameof(bonusProgram));
        }

        if (string.IsNullOrWhiteSpace(bonusProgram.ProgramName))
        {
            throw new ArgumentException("ProgramName is required.", nameof(bonusProgram));
        }

        if (bonusProgram.ProgramYear <= 0)
        {
            throw new ArgumentException("ProgramYear is required.", nameof(bonusProgram));
        }

        if (bonusProgram.QualificationStartDate == default)
        {
            throw new ArgumentException("QualificationStartDate is required.", nameof(bonusProgram));
        }

        if (bonusProgram.QualificationEndDate == default)
        {
            throw new ArgumentException("QualificationEndDate is required.", nameof(bonusProgram));
        }

        if (bonusProgram.QualificationEndDate < bonusProgram.QualificationStartDate)
        {
            throw new ArgumentException("QualificationEndDate cannot be before QualificationStartDate.", nameof(bonusProgram));
        }
    }

    private static void Normalize(BonusProgramRecord bonusProgram)
    {
        bonusProgram.ProgramName = Clean(bonusProgram.ProgramName);
        bonusProgram.PaymentFrequency = OptionalText(bonusProgram.PaymentFrequency) as string;
        bonusProgram.SourceDocument = OptionalText(bonusProgram.SourceDocument) as string;
        bonusProgram.Notes = OptionalText(bonusProgram.Notes) as string;
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

    private static long ReadLong(SqliteDataReader reader, string name)
    {
        return reader.GetInt64(reader.GetOrdinal(name));
    }

    private static int? ReadInt(SqliteDataReader reader, string name)
    {
        var ordinal = reader.GetOrdinal(name);
        return reader.IsDBNull(ordinal) ? null : Convert.ToInt32(reader.GetValue(ordinal), CultureInfo.InvariantCulture);
    }

    private static string ReadString(SqliteDataReader reader, string name)
    {
        var ordinal = reader.GetOrdinal(name);
        return reader.IsDBNull(ordinal) ? "" : reader.GetString(ordinal);
    }

    private static string? ReadNullableString(SqliteDataReader reader, string name)
    {
        var value = ReadString(reader, name);
        return string.IsNullOrWhiteSpace(value) ? null : value;
    }

    private static bool ReadBoolean(SqliteDataReader reader, string name)
    {
        var ordinal = reader.GetOrdinal(name);
        return !reader.IsDBNull(ordinal) && Convert.ToInt32(reader.GetValue(ordinal), CultureInfo.InvariantCulture) == 1;
    }

    private static DateOnly? ReadDateOnly(SqliteDataReader reader, string name)
    {
        var value = ReadString(reader, name);
        return DateOnly.TryParseExact(value, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var date)
            ? date
            : null;
    }

    private static DateTime? ReadDateTime(SqliteDataReader reader, string name)
    {
        var value = ReadString(reader, name);
        return DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var date)
            ? date
            : null;
    }

    private static string ToDateText(DateOnly date)
    {
        return date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
    }

    private static int ToInteger(bool value)
    {
        return value ? 1 : 0;
    }

    private static object? OptionalText(string? value)
    {
        return string.IsNullOrWhiteSpace(value) ? null : value.Trim();
    }

    private static string Clean(string? value)
    {
        return value?.Trim() ?? "";
    }

    private const string BaseSelectSql = @"
SELECT bp.*,
       c.carrier_name AS CarrierName
FROM BonusProgram bp
INNER JOIN Carrier c ON c.carrier_id = bp.CarrierID
";
}

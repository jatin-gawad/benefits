using System.Globalization;
using Microsoft.Data.Sqlite;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;

namespace RenewalReviewBuilder.App.Repositories;

/// <summary>
/// SQLite repository for the Carrier master reference table.
/// </summary>
public sealed class CarrierRepository
{
    private readonly string _databasePath;

    /// <summary>
    /// Creates a repository that uses the shared Benefits Intelligence database unless a test path is supplied.
    /// </summary>
    public CarrierRepository(string? databasePath = null)
    {
        _databasePath = string.IsNullOrWhiteSpace(databasePath)
            ? AppPaths.BenefitsIntelligenceDatabasePath
            : databasePath;
    }

    /// <summary>
    /// Gets every carrier ordered by name.
    /// </summary>
    public async Task<List<CarrierRecord>> GetAll()
    {
        return await QueryCarriersAsync(@"
SELECT *
FROM Carrier
ORDER BY carrier_name COLLATE NOCASE;");
    }

    /// <summary>
    /// Gets active carriers ordered by name.
    /// </summary>
    public async Task<List<CarrierRecord>> GetActive()
    {
        return await QueryCarriersAsync(@"
SELECT *
FROM Carrier
WHERE is_active = 1
ORDER BY carrier_name COLLATE NOCASE;");
    }

    /// <summary>
    /// Gets one carrier by primary key.
    /// </summary>
    public async Task<CarrierRecord?> GetById(long carrierId)
    {
        return await QuerySingleCarrierAsync(
            "SELECT * FROM Carrier WHERE carrier_id = $carrier_id LIMIT 1;",
            command => AddParameter(command, "$carrier_id", carrierId));
    }

    /// <summary>
    /// Gets one carrier by unique carrier code.
    /// </summary>
    public async Task<CarrierRecord?> GetByCode(string carrierCode)
    {
        return await QuerySingleCarrierAsync(
            "SELECT * FROM Carrier WHERE carrier_code = $carrier_code COLLATE NOCASE LIMIT 1;",
            command => AddParameter(command, "$carrier_code", Clean(carrierCode)));
    }

    /// <summary>
    /// Gets one carrier by unique carrier name.
    /// </summary>
    public async Task<CarrierRecord?> GetByName(string carrierName)
    {
        return await QuerySingleCarrierAsync(
            "SELECT * FROM Carrier WHERE carrier_name = $carrier_name COLLATE NOCASE LIMIT 1;",
            command => AddParameter(command, "$carrier_name", Clean(carrierName)));
    }

    /// <summary>
    /// Inserts a new carrier after validating required fields and uniqueness.
    /// </summary>
    public async Task<long> Insert(CarrierRecord carrier)
    {
        try
        {
            Normalize(carrier);
            Validate(carrier);
            if (await Exists(carrier.CarrierCode, carrier.CarrierName))
            {
                throw new InvalidOperationException("CarrierCode and CarrierName must be unique.");
            }

            await using var connection = await OpenConnectionAsync();
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
    $is_active,
    $status,
    CURRENT_TIMESTAMP
);
SELECT last_insert_rowid();");
            AddCarrierParameters(command, carrier);
            var id = Convert.ToInt64(await command.ExecuteScalarAsync(), CultureInfo.InvariantCulture);
            carrier.CarrierId = id;
            AppLogger.Info($"Inserted carrier {carrier.CarrierCode} - {carrier.CarrierName}.");
            return id;
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            throw;
        }
    }

    /// <summary>
    /// Updates an existing carrier after validating required fields and uniqueness.
    /// </summary>
    public async Task<bool> Update(CarrierRecord carrier)
    {
        try
        {
            if (carrier.CarrierId <= 0)
            {
                throw new ArgumentException("CarrierId is required for updates.", nameof(carrier));
            }

            Normalize(carrier);
            Validate(carrier);
            if (await Exists(carrier.CarrierCode, carrier.CarrierName, carrier.CarrierId))
            {
                throw new InvalidOperationException("CarrierCode and CarrierName must be unique.");
            }

            await using var connection = await OpenConnectionAsync();
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
    is_active = $is_active,
    status = $status,
    modified_date = CURRENT_TIMESTAMP
WHERE carrier_id = $carrier_id;");
            AddCarrierParameters(command, carrier);
            AddParameter(command, "$carrier_id", carrier.CarrierId);
            var updated = await command.ExecuteNonQueryAsync();
            if (updated > 0)
            {
                AppLogger.Info($"Updated carrier {carrier.CarrierCode} - {carrier.CarrierName}.");
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
    /// Marks a carrier inactive without deleting the master record.
    /// </summary>
    public async Task<bool> Deactivate(long carrierId)
    {
        try
        {
            await using var connection = await OpenConnectionAsync();
            await using var command = CreateCommand(connection, @"
UPDATE Carrier
SET is_active = 0,
    status = 'Inactive',
    modified_date = CURRENT_TIMESTAMP
WHERE carrier_id = $carrier_id;");
            AddParameter(command, "$carrier_id", carrierId);
            var updated = await command.ExecuteNonQueryAsync();
            if (updated > 0)
            {
                AppLogger.Info($"Deactivated carrier {carrierId}.");
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
    /// Checks whether a carrier code or name already exists.
    /// </summary>
    public async Task<bool> Exists(string carrierCode, string carrierName = "", long? excludingCarrierId = null)
    {
        try
        {
            await using var connection = await OpenConnectionAsync();
            await using var command = CreateCommand(connection, @"
SELECT COUNT(*)
FROM Carrier
WHERE (
        carrier_code = $carrier_code COLLATE NOCASE
        OR carrier_name = $carrier_name COLLATE NOCASE
      )
  AND ($excluding_carrier_id IS NULL OR carrier_id <> $excluding_carrier_id);");
            AddParameter(command, "$carrier_code", Clean(carrierCode));
            AddParameter(command, "$carrier_name", Clean(carrierName));
            AddParameter(command, "$excluding_carrier_id", excludingCarrierId);
            return Convert.ToInt32(await command.ExecuteScalarAsync(), CultureInfo.InvariantCulture) > 0;
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            throw;
        }
    }

    private async Task<List<CarrierRecord>> QueryCarriersAsync(string sql)
    {
        try
        {
            await using var connection = await OpenConnectionAsync();
            await using var command = CreateCommand(connection, sql);
            var carriers = new List<CarrierRecord>();
            await using var reader = await command.ExecuteReaderAsync();
            while (await reader.ReadAsync())
            {
                carriers.Add(ReadCarrier(reader));
            }

            return carriers;
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            throw;
        }
    }

    private async Task<CarrierRecord?> QuerySingleCarrierAsync(string sql, Action<SqliteCommand> addParameters)
    {
        try
        {
            await using var connection = await OpenConnectionAsync();
            await using var command = CreateCommand(connection, sql);
            addParameters(command);
            await using var reader = await command.ExecuteReaderAsync();
            return await reader.ReadAsync() ? ReadCarrier(reader) : null;
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

    private static CarrierRecord ReadCarrier(SqliteDataReader reader)
    {
        return new CarrierRecord
        {
            CarrierId = ReadLong(reader, "carrier_id"),
            CarrierCode = ReadString(reader, "carrier_code"),
            CarrierName = ReadString(reader, "carrier_name"),
            CommissionStatementFormat = ReadString(reader, "commission_statement_format"),
            BonusProgramVersion = ReadNullableString(reader, "bonus_program_version"),
            HasRateGuarantees = ReadBoolean(reader, "has_rate_guarantees"),
            HasPersistencyBonus = ReadBoolean(reader, "has_persistency_bonus"),
            HasSalesBonus = ReadBoolean(reader, "has_sales_bonus"),
            HasGrowthBonus = ReadBoolean(reader, "has_growth_bonus"),
            SupportsBillingImport = ReadBoolean(reader, "supports_billing_import"),
            SupportsClaimsImport = ReadBoolean(reader, "supports_claims_import"),
            IsActive = ReadBoolean(reader, "is_active"),
            CreatedDate = ReadDateTime(reader, "created_date") ?? DateTime.MinValue,
            ModifiedDate = ReadDateTime(reader, "modified_date")
        };
    }

    private static void AddCarrierParameters(SqliteCommand command, CarrierRecord carrier)
    {
        AddParameter(command, "$carrier_code", carrier.CarrierCode);
        AddParameter(command, "$carrier_name", carrier.CarrierName);
        AddParameter(command, "$commission_statement_format", carrier.CommissionStatementFormat);
        AddParameter(command, "$bonus_program_version", string.IsNullOrWhiteSpace(carrier.BonusProgramVersion) ? null : carrier.BonusProgramVersion);
        AddParameter(command, "$has_rate_guarantees", ToInteger(carrier.HasRateGuarantees));
        AddParameter(command, "$has_persistency_bonus", ToInteger(carrier.HasPersistencyBonus));
        AddParameter(command, "$has_sales_bonus", ToInteger(carrier.HasSalesBonus));
        AddParameter(command, "$has_growth_bonus", ToInteger(carrier.HasGrowthBonus));
        AddParameter(command, "$supports_billing_import", ToInteger(carrier.SupportsBillingImport));
        AddParameter(command, "$supports_claims_import", ToInteger(carrier.SupportsClaimsImport));
        AddParameter(command, "$is_active", ToInteger(carrier.IsActive));
        AddParameter(command, "$status", carrier.IsActive ? "Active" : "Inactive");
    }

    private static void Validate(CarrierRecord carrier)
    {
        if (string.IsNullOrWhiteSpace(carrier.CarrierCode))
        {
            throw new ArgumentException("CarrierCode is required.", nameof(carrier));
        }

        if (string.IsNullOrWhiteSpace(carrier.CarrierName))
        {
            throw new ArgumentException("CarrierName is required.", nameof(carrier));
        }

        if (string.IsNullOrWhiteSpace(carrier.CommissionStatementFormat))
        {
            throw new ArgumentException("CommissionStatementFormat is required.", nameof(carrier));
        }
    }

    private static void Normalize(CarrierRecord carrier)
    {
        carrier.CarrierCode = Clean(carrier.CarrierCode).ToUpperInvariant();
        carrier.CarrierName = Clean(carrier.CarrierName);
        carrier.CommissionStatementFormat = Clean(carrier.CommissionStatementFormat).ToUpperInvariant();
        carrier.BonusProgramVersion = string.IsNullOrWhiteSpace(carrier.BonusProgramVersion)
            ? null
            : carrier.BonusProgramVersion.Trim();
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

    private static DateTime? ReadDateTime(SqliteDataReader reader, string name)
    {
        var value = ReadString(reader, name);
        return DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var date)
            ? date
            : null;
    }

    private static int ToInteger(bool value)
    {
        return value ? 1 : 0;
    }

    private static string Clean(string? value)
    {
        return value?.Trim() ?? "";
    }
}

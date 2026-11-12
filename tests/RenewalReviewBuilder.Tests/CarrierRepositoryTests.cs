using Microsoft.Data.Sqlite;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Repositories;
using RenewalReviewBuilder.App.Services;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class CarrierRepositoryTests
{
    [Fact]
    public async Task Initialize_Creates_Carrier_Master_Table_Indexes_And_Seed_Data()
    {
        var databasePath = NewDatabasePath();
        try
        {
            var service = new DatabaseService(databasePath);
            var initialize = service.Initialize();
            Assert.True(initialize.Succeeded, initialize.Message);

            await using var connection = new SqliteConnection($"Data Source={databasePath}");
            await connection.OpenAsync();

            Assert.True(await TableExistsAsync(connection, "Carrier"));
            Assert.True(await ColumnExistsAsync(connection, "Carrier", "carrier_code"));
            Assert.True(await ColumnExistsAsync(connection, "Carrier", "commission_statement_format"));
            Assert.True(await ColumnExistsAsync(connection, "Carrier", "supports_claims_import"));
            Assert.True(await IndexExistsAsync(connection, "IDX_Carrier_Name"));
            Assert.True(await IndexExistsAsync(connection, "IDX_Carrier_Code"));
            Assert.True(await IndexExistsAsync(connection, "IDX_Carrier_Active"));

            var repository = new CarrierRepository(databasePath);
            var carriers = await repository.GetAll();
            Assert.Equal(11, carriers.Count);
            Assert.Contains(carriers, carrier =>
                carrier.CarrierCode == "ALL" && carrier.CarrierName == "Allstate");
            Assert.Contains(carriers, carrier =>
                carrier.CarrierCode == "NHC" && carrier.CarrierName == "National HealthClaim");

            var canadaLife = Assert.Single(carriers, carrier => carrier.CarrierCode == "CL");
            Assert.Equal("Canada Life", canadaLife.CarrierName);
            Assert.Equal("CANADALIFE_V1", canadaLife.CommissionStatementFormat);
            Assert.Equal("2025", canadaLife.BonusProgramVersion);
            Assert.True(canadaLife.HasGrowthBonus);
            Assert.True(canadaLife.SupportsBillingImport);
            Assert.True(canadaLife.SupportsClaimsImport);
        }
        finally
        {
            DeleteDatabase(databasePath);
        }
    }

    [Fact]
    public async Task Carrier_Repository_CRUD_Validates_Uniqueness_And_Deactivates()
    {
        var databasePath = NewDatabasePath();
        try
        {
            var repository = new CarrierRepository(databasePath);
            var carrier = new CarrierRecord
            {
                CarrierCode = "TST",
                CarrierName = "Test Carrier",
                CommissionStatementFormat = "TEST_V1",
                BonusProgramVersion = "2026",
                HasRateGuarantees = true,
                HasPersistencyBonus = true,
                HasSalesBonus = false,
                HasGrowthBonus = false,
                SupportsBillingImport = false,
                SupportsClaimsImport = true,
                IsActive = true
            };

            var id = await repository.Insert(carrier);
            Assert.True(id > 0);
            Assert.True(await repository.Exists("TST"));

            var byCode = await repository.GetByCode("tst");
            Assert.NotNull(byCode);
            Assert.Equal("Test Carrier", byCode.CarrierName);

            byCode.CarrierName = "Test Carrier Updated";
            byCode.SupportsBillingImport = true;
            Assert.True(await repository.Update(byCode));

            var byName = await repository.GetByName("test carrier updated");
            Assert.NotNull(byName);
            Assert.True(byName.SupportsBillingImport);

            await Assert.ThrowsAsync<InvalidOperationException>(() => repository.Insert(new CarrierRecord
            {
                CarrierCode = "TST",
                CarrierName = "Another Test Carrier",
                CommissionStatementFormat = "TEST_V2"
            }));

            Assert.True(await repository.Deactivate(id));
            var active = await repository.GetActive();
            Assert.DoesNotContain(active, activeCarrier => activeCarrier.CarrierId == id);
        }
        finally
        {
            DeleteDatabase(databasePath);
        }
    }

    [Fact]
    public async Task Carrier_Migration_Upgrades_Legacy_SourceTracking_Table_Idempotently()
    {
        var databasePath = NewDatabasePath();
        try
        {
            await using (var connection = new SqliteConnection($"Data Source={databasePath}"))
            {
                await connection.OpenAsync();
                await using var create = connection.CreateCommand();
                create.CommandText = @"
CREATE TABLE Carrier (
    carrier_id INTEGER PRIMARY KEY AUTOINCREMENT,
    carrier_name TEXT NOT NULL COLLATE NOCASE,
    status TEXT NOT NULL DEFAULT 'Active',
    date_loaded TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    UNIQUE (carrier_name)
);
INSERT INTO Carrier (carrier_name) VALUES ('Manulife');";
                await create.ExecuteNonQueryAsync();
            }

            var service = new DatabaseService(databasePath);
            Assert.True(service.Initialize().Succeeded);
            Assert.True(service.Initialize().Succeeded);

            var repository = new CarrierRepository(databasePath);
            var manulife = await repository.GetByCode("MFC");
            Assert.NotNull(manulife);
            Assert.Equal("Manulife", manulife.CarrierName);
            Assert.Equal("MANULIFE_V1", manulife.CommissionStatementFormat);

            var carriers = await repository.GetAll();
            Assert.Equal(11, carriers.Count);
            Assert.Contains(carriers, carrier =>
                carrier.CarrierCode == "ALL" && carrier.CarrierName == "Allstate");
            Assert.Contains(carriers, carrier =>
                carrier.CarrierCode == "NHC" && carrier.CarrierName == "National HealthClaim");
        }
        finally
        {
            DeleteDatabase(databasePath);
        }
    }

    private static string NewDatabasePath()
    {
        var testFolder = Path.Combine(Path.GetTempPath(), "BenefitsIntelligenceDbTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(testFolder);
        return Path.Combine(testFolder, "Benefits_Intelligence.db");
    }

    private static async Task<bool> TableExistsAsync(SqliteConnection connection, string tableName)
    {
        await using var command = connection.CreateCommand();
        command.CommandText = "SELECT COUNT(*) FROM sqlite_master WHERE type = 'table' AND name = $table_name;";
        command.Parameters.AddWithValue("$table_name", tableName);
        return Convert.ToInt32(await command.ExecuteScalarAsync()) > 0;
    }

    private static async Task<bool> IndexExistsAsync(SqliteConnection connection, string indexName)
    {
        await using var command = connection.CreateCommand();
        command.CommandText = "SELECT COUNT(*) FROM sqlite_master WHERE type = 'index' AND name = $index_name;";
        command.Parameters.AddWithValue("$index_name", indexName);
        return Convert.ToInt32(await command.ExecuteScalarAsync()) > 0;
    }

    private static async Task<bool> ColumnExistsAsync(SqliteConnection connection, string tableName, string columnName)
    {
        await using var command = connection.CreateCommand();
        command.CommandText = $"PRAGMA table_info({tableName});";
        await using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            if (reader.GetString(reader.GetOrdinal("name")).Equals(columnName, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    private static void DeleteDatabase(string databasePath)
    {
        SqliteConnection.ClearAllPools();
        var folder = Path.GetDirectoryName(databasePath);
        if (!string.IsNullOrWhiteSpace(folder) && Directory.Exists(folder))
        {
            Directory.Delete(folder, recursive: true);
        }
    }
}

using Microsoft.Data.Sqlite;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Repositories;
using RenewalReviewBuilder.App.Services;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class BonusProgramRepositoryTests
{
    [Fact]
    public async Task Initialize_Creates_BonusProgram_Table_Indexes_ForeignKey_And_Seed_Data()
    {
        var databasePath = NewDatabasePath();
        try
        {
            var service = new DatabaseService(databasePath);
            Assert.True(service.Initialize().Succeeded);
            Assert.True(service.Initialize().Succeeded);

            await using var connection = new SqliteConnection($"Data Source={databasePath}");
            await connection.OpenAsync();

            Assert.True(await TableExistsAsync(connection, "BonusProgram"));
            Assert.True(await ColumnExistsAsync(connection, "BonusProgram", "BonusProgramID"));
            Assert.True(await ColumnExistsAsync(connection, "BonusProgram", "CarrierID"));
            Assert.True(await ColumnExistsAsync(connection, "BonusProgram", "HasRateGuaranteeImpact"));
            Assert.True(await IndexExistsAsync(connection, "IDX_BonusProgram_CarrierID"));
            Assert.True(await IndexExistsAsync(connection, "IDX_BonusProgram_ProgramYear"));
            Assert.True(await IndexExistsAsync(connection, "IDX_BonusProgram_Active"));
            Assert.True(await ForeignKeyExistsAsync(connection, "BonusProgram", "CarrierID", "Carrier", "carrier_id"));

            var repository = new BonusProgramRepository(databasePath);
            var programs = await repository.GetAll();
            Assert.Equal(9, programs.Count);
            Assert.Equal(9, programs.Select(program => (program.CarrierId, program.ProgramYear)).Distinct().Count());

            var canadaLife = Assert.Single(programs, program => program.CarrierName == "Canada Life");
            Assert.Equal(2025, canadaLife.ProgramYear);
            Assert.Equal(new DateOnly(2025, 1, 1), canadaLife.QualificationStartDate);
            Assert.Equal(new DateOnly(2025, 12, 31), canadaLife.QualificationEndDate);
            Assert.True(canadaLife.HasSalesBonus);
            Assert.True(canadaLife.HasPersistencyBonus);
            Assert.True(canadaLife.HasGrowthBonus);
            Assert.True(canadaLife.HasRateGuaranteeImpact);
        }
        finally
        {
            DeleteDatabase(databasePath);
        }
    }

    [Fact]
    public async Task BonusProgram_Repository_CRUD_Validates_Uniqueness_And_Deactivates()
    {
        var databasePath = NewDatabasePath();
        try
        {
            var carrierRepository = new CarrierRepository(databasePath);
            var carrier = await carrierRepository.GetByCode("MFC");
            Assert.NotNull(carrier);

            var repository = new BonusProgramRepository(databasePath);
            var bonusProgram = new BonusProgramRecord
            {
                CarrierId = carrier.CarrierId,
                ProgramName = "Manulife Test Bonus Program",
                ProgramYear = 2027,
                QualificationStartDate = new DateOnly(2027, 1, 1),
                QualificationEndDate = new DateOnly(2027, 12, 31),
                PaymentFrequency = "Annual",
                HasSalesBonus = true,
                HasPersistencyBonus = true,
                HasGrowthBonus = false,
                HasRetentionBonus = true,
                HasEliteBonus = false,
                HasSalesEnhancement = true,
                HasRateGuaranteeImpact = true,
                SourceDocument = "test-source.pdf",
                Notes = "Test row.",
                IsActive = true
            };

            var id = await repository.Insert(bonusProgram);
            Assert.True(id > 0);
            Assert.True(await repository.Exists(carrier.CarrierId, 2027, "Manulife Test Bonus Program"));

            var byCarrier = await repository.GetByCarrierId(carrier.CarrierId);
            Assert.Contains(byCarrier, program => program.BonusProgramId == id);

            var byCarrierAndYear = await repository.GetByCarrierAndYear(carrier.CarrierId, 2027);
            Assert.NotNull(byCarrierAndYear);
            Assert.Equal("Annual", byCarrierAndYear.PaymentFrequency);

            byCarrierAndYear.ProgramName = "Manulife Test Bonus Program Updated";
            byCarrierAndYear.HasGrowthBonus = true;
            Assert.True(await repository.Update(byCarrierAndYear));

            var updated = await repository.GetByCarrierAndYear(carrier.CarrierId, 2027);
            Assert.NotNull(updated);
            Assert.True(updated.HasGrowthBonus);

            await Assert.ThrowsAsync<InvalidOperationException>(() => repository.Insert(new BonusProgramRecord
            {
                CarrierId = carrier.CarrierId,
                ProgramName = "Manulife Test Bonus Program Updated",
                ProgramYear = 2027,
                QualificationStartDate = new DateOnly(2027, 1, 1),
                QualificationEndDate = new DateOnly(2027, 12, 31)
            }));

            Assert.True(await repository.Deactivate(id));
            var active = await repository.GetActive();
            Assert.DoesNotContain(active, activeProgram => activeProgram.BonusProgramId == id);
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

    private static async Task<bool> ForeignKeyExistsAsync(
        SqliteConnection connection,
        string tableName,
        string fromColumn,
        string toTable,
        string toColumn)
    {
        await using var command = connection.CreateCommand();
        command.CommandText = $"PRAGMA foreign_key_list({tableName});";
        await using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            if (reader.GetString(reader.GetOrdinal("from")).Equals(fromColumn, StringComparison.OrdinalIgnoreCase) &&
                reader.GetString(reader.GetOrdinal("table")).Equals(toTable, StringComparison.OrdinalIgnoreCase) &&
                reader.GetString(reader.GetOrdinal("to")).Equals(toColumn, StringComparison.OrdinalIgnoreCase))
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

using Microsoft.Data.Sqlite;
using RenewalReviewBuilder.App.Services;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class DatabaseSourceTrackingSchemaTests
{
    [Fact]
    public async Task Initialize_CreatesImportTrackingTablesAndSourceColumns()
    {
        var databasePath = Path.Combine(Path.GetTempPath(), $"benefits-intelligence-{Guid.NewGuid():N}.db");
        try
        {
            var service = new DatabaseService(databasePath);
            var result = service.Initialize();

            Assert.True(result.Succeeded, result.Message);

            await using var connection = new SqliteConnection($"Data Source={databasePath}");
            await connection.OpenAsync();

            Assert.True(await TableExistsAsync(connection, "ImportBatch"));
            Assert.True(await TableExistsAsync(connection, "DocumentSource"));
            Assert.True(await TableExistsAsync(connection, "ImportLog"));
            Assert.True(await TableExistsAsync(connection, "ValidationIssue"));
            Assert.True(await TableExistsAsync(connection, "Carrier"));
            Assert.True(await TableExistsAsync(connection, "Policy"));
            Assert.True(await ColumnExistsAsync(connection, "Renewals", "source_type"));
            Assert.True(await ColumnExistsAsync(connection, "Renewals", "import_batch_id"));
            Assert.True(await ColumnExistsAsync(connection, "Rates", "confidence_score"));
            Assert.True(await ColumnExistsAsync(connection, "PlanDesign", "raw_source_reference"));
            Assert.True(await ColumnExistsAsync(connection, "PlanDesign", "class_name"));
            Assert.True(await ColumnExistsAsync(connection, "PlanDesign", "source_section"));
            Assert.True(await ColumnExistsAsync(connection, "PlanDesign", "notes"));
            Assert.True(await ColumnExistsAsync(connection, "PlanDesign", "booklet_source_file_name"));
            Assert.True(await ColumnExistsAsync(connection, "billing_option_premiums", "extraction_method"));
            Assert.True(await TableExistsAsync(connection, "class_comparison_preferences"));
            Assert.True(await TableExistsAsync(connection, "class_comparison_notes"));
            Assert.True(await TableExistsAsync(connection, "class_comparison_audit"));
            Assert.True(await ColumnExistsAsync(connection, "class_comparison_preferences", "recommended_class_name"));
            Assert.True(await ColumnExistsAsync(connection, "class_comparison_notes", "updated_by"));
        }
        finally
        {
            SqliteConnection.ClearAllPools();
            if (File.Exists(databasePath))
            {
                File.Delete(databasePath);
            }
        }
    }

    private static async Task<bool> TableExistsAsync(SqliteConnection connection, string tableName)
    {
        await using var command = connection.CreateCommand();
        command.CommandText = "SELECT COUNT(*) FROM sqlite_master WHERE type = 'table' AND name = $table_name;";
        command.Parameters.AddWithValue("$table_name", tableName);
        var count = Convert.ToInt32(await command.ExecuteScalarAsync());
        return count > 0;
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
}

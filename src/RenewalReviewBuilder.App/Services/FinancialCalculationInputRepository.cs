using Microsoft.Data.Sqlite;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

/// <summary>Builds calculation inputs from existing desktop SQLite records without changing their ownership.</summary>
public sealed class FinancialCalculationInputRepository
{
    private readonly string _databasePath;
    public FinancialCalculationInputRepository(string databasePath) => _databasePath = databasePath;

    public async Task<FinancialCalculationInputSet> LoadAsync(long clientId)
    {
        var inputs = new FinancialCalculationInputSet();
        await using var connection = new SqliteConnection($"Data Source={_databasePath}");
        await connection.OpenAsync();
        await AddScalarAsync(connection, inputs.FinalizedRenewalProjection, "Billed premium", "Finalized renewal projection", "SELECT SUM(COALESCE(negotiated_premium, proposed_premium, required_premium)) FROM renewal_projections WHERE client_id=$client;", clientId);
        await AddScalarAsync(connection, inputs.ValidatedClaimsProjection, "Total benefit payments", "Validated claims projection", "SELECT SUM(COALESCE(incurred_claims, paid_claims)) FROM ClaimsExperience WHERE client_id=$client;", clientId);
        await AddScalarAsync(connection, inputs.ValidatedAsoActivity, "Total benefit payments", "Validated ASO activity", "SELECT total_benefit_payment FROM aso_financial_periods WHERE statement_id IN (SELECT statement_id FROM aso_financial_statements WHERE client_id=$client AND is_current=1) AND is_total=1 ORDER BY period_id DESC LIMIT 1;", clientId);
        await AddScalarAsync(connection, inputs.ActiveClientFinancialData, "Billed premium", "Active client billing", "SELECT SUM(estimated_annual_premium) FROM billing_province_summary WHERE client_id=$client;", clientId);
        return inputs;
    }

    private static async Task AddScalarAsync(SqliteConnection connection, List<FinancialCalculationInput> destination, string basis, string source, string sql, long clientId)
    {
        await using var command = connection.CreateCommand(); command.CommandText = sql; command.Parameters.AddWithValue("$client", clientId);
        try
        {
            var value = await command.ExecuteScalarAsync();
            if (value is null || value == DBNull.Value) return;
            destination.Add(new FinancialCalculationInput { Basis = basis, Value = Convert.ToDecimal(value), Source = source, Period = DateTime.UtcNow.Year.ToString(), IsValidated = true });
        }
        catch (SqliteException)
        {
            // Older client databases can lack a source table; resolver continues to lower-precedence sources.
        }
    }
}

using System.Globalization;
using Microsoft.Data.Sqlite;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

/// <summary>Builds the UI read model from immutable snapshot references and canonical records.</summary>
public sealed class RenewalFinancialArrangementComparisonDataService
{
    private readonly string _databasePath;
    public RenewalFinancialArrangementComparisonDataService(string databasePath) => _databasePath = databasePath;

    public async Task<decimal?> FindLatestSoldProposalCommissionAsync(long clientId)
    {
        if (!File.Exists(_databasePath))
        {
            return null;
        }

        await using var connection = new SqliteConnection($"Data Source={_databasePath};Mode=ReadOnly");
        await connection.OpenAsync();
        await using var command = connection.CreateCommand();
        command.CommandText = """
            SELECT p.commission
            FROM proposal_pooling_retention p
            JOIN proposal_summary s ON s.proposal_id = p.proposal_id
            WHERE CAST(s.client_id AS TEXT) = CAST($client AS TEXT)
              AND p.commission IS NOT NULL
            ORDER BY s.created_at DESC
            LIMIT 1;
            """;
        command.Parameters.AddWithValue("$client", clientId);
        var value = await command.ExecuteScalarAsync();
        return value is null || value == DBNull.Value
            ? null
            : Convert.ToDecimal(value, CultureInfo.InvariantCulture);
    }

    public async Task<RenewalFinancialArrangementComparison?> LoadAsync(long snapshotId)
    {
        await using var connection = new SqliteConnection($"Data Source={_databasePath}"); await connection.OpenAsync();
        var snapshot = await LoadSnapshotAsync(connection, snapshotId); if (snapshot is null) return null;
        var current = await LoadVersionAsync(connection, snapshot.CurrentArrangementVersionId); var proposed = await LoadVersionAsync(connection, snapshot.ProposedArrangementVersionId);
        if (current is null || proposed is null) return null;
        var comparison = new RenewalFinancialArrangementComparisonService().Build(snapshot.RenewalId, current.Arrangement, proposed.Arrangement, current.Commission, proposed.Commission, snapshot.Scenario, snapshot.IsStale);
        comparison.SnapshotStatus = snapshot.SnapshotStatus;
        comparison.CurrentPoolingRate = current.Pooling?.Rate; comparison.ProposedPoolingRate = proposed.Pooling?.Rate;
        comparison.CurrentPoolingThreshold = current.Pooling?.Threshold; comparison.ProposedPoolingThreshold = proposed.Pooling?.Threshold;
        return comparison;
    }

    public async Task<FinancialImpactComparison?> LoadImpactAsync(
        long snapshotId,
        FinancialCalculationInputSet currentInputs,
        FinancialCalculationInputSet proposedInputs)
    {
        await using var connection = new SqliteConnection($"Data Source={_databasePath}"); await connection.OpenAsync();
        var snapshot = await LoadSnapshotAsync(connection, snapshotId); if (snapshot is null) return null;
        var current = await LoadVersionAsync(connection, snapshot.CurrentArrangementVersionId);
        var proposed = await LoadVersionAsync(connection, snapshot.ProposedArrangementVersionId);
        if (current is null || proposed is null) return null;
        return new FinancialImpactCalculationService().Calculate(
            current.Arrangement, proposed.Arrangement,
            current.Commission, proposed.Commission,
            current.Pooling, proposed.Pooling,
            currentInputs, proposedInputs);
    }

    public async Task<FinancialProjectionRequest?> LoadProjectionRequestAsync(
        long snapshotId,
        IReadOnlyList<FinancialHistoricalPeriodSource> completedHistoricalPeriods,
        FinancialCalculationInputSet projectedInputs,
        int minimumCompleteMonths = 6)
    {
        await using var connection = new SqliteConnection($"Data Source={_databasePath}"); await connection.OpenAsync();
        var snapshot = await LoadSnapshotAsync(connection, snapshotId); if (snapshot is null) return null;
        var current = await LoadVersionAsync(connection, snapshot.CurrentArrangementVersionId);
        var proposed = await LoadVersionAsync(connection, snapshot.ProposedArrangementVersionId);
        if (current is null || proposed is null) return null;
        return new FinancialProjectionRequest
        {
            CompletedHistoricalPeriods = completedHistoricalPeriods,
            MinimumCompleteMonths = minimumCompleteMonths,
            CurrentArrangement = current.Arrangement,
            RenewalArrangement = proposed.Arrangement,
            CurrentCommission = current.Commission,
            RenewalCommission = proposed.Commission,
            CurrentPooling = current.Pooling,
            RenewalPooling = proposed.Pooling,
            ProjectedInputs = projectedInputs
        };
    }

    private static async Task<RenewalFinancialArrangementSnapshot?> LoadSnapshotAsync(SqliteConnection connection, long id)
    {
        await using var command = connection.CreateCommand(); command.CommandText = "SELECT * FROM renewal_financial_arrangement_snapshots WHERE renewal_financial_arrangement_snapshot_id=$id;"; command.Parameters.AddWithValue("$id", id); await using var reader = await command.ExecuteReaderAsync();
        return await reader.ReadAsync() ? new RenewalFinancialArrangementSnapshot { RenewalFinancialArrangementSnapshotId = id, RenewalId = reader.GetInt64(reader.GetOrdinal("renewal_id")), ClientId = reader.GetInt64(reader.GetOrdinal("client_id")), Carrier = reader.GetString(reader.GetOrdinal("carrier")), Scenario = Enum.Parse<FinancialArrangementScenario>(reader.GetString(reader.GetOrdinal("scenario"))), SnapshotStatus = reader.GetString(reader.GetOrdinal("snapshot_status")), CurrentArrangementVersionId = reader.GetInt64(reader.GetOrdinal("current_arrangement_version_id")), ProposedArrangementVersionId = reader.GetInt64(reader.GetOrdinal("proposed_arrangement_version_id")), IsStale = reader.GetInt64(reader.GetOrdinal("is_stale")) != 0 } : null;
    }

    private static async Task<VersionData?> LoadVersionAsync(SqliteConnection connection, long versionId)
    {
        long arrangementId; long? commissionId; long? poolingId;
        await using (var command = connection.CreateCommand())
        {
            command.CommandText = "SELECT * FROM financial_arrangement_versions WHERE financial_arrangement_version_id=$id;"; command.Parameters.AddWithValue("$id", versionId); await using var reader = await command.ExecuteReaderAsync(); if (!await reader.ReadAsync()) return null;
            arrangementId = reader.GetInt64(reader.GetOrdinal("financial_arrangement_id"));
            commissionId = reader.IsDBNull(reader.GetOrdinal("commission_arrangement_id")) ? null : reader.GetInt64(reader.GetOrdinal("commission_arrangement_id"));
            poolingId = reader.IsDBNull(reader.GetOrdinal("pooling_arrangement_id")) ? null : reader.GetInt64(reader.GetOrdinal("pooling_arrangement_id"));
        }
        var arrangement = await LoadArrangementAsync(connection, arrangementId);
        if (arrangement is null) return null;
        return new VersionData(arrangement, commissionId.HasValue ? await LoadCommissionAsync(connection, commissionId.Value) : null, poolingId.HasValue ? await LoadPoolingAsync(connection, poolingId.Value) : null);
    }

    private static async Task<FinancialArrangement?> LoadArrangementAsync(SqliteConnection connection, long id)
    {
        await using var command = connection.CreateCommand(); command.CommandText = "SELECT * FROM financial_arrangements WHERE financial_arrangement_id=$id;"; command.Parameters.AddWithValue("$id", id); await using var reader = await command.ExecuteReaderAsync(); if (!await reader.ReadAsync()) return null;
        var arrangement = new FinancialArrangement { FinancialArrangementId = id, ClientId = reader.GetInt64(reader.GetOrdinal("client_id")), Carrier = reader.GetString(reader.GetOrdinal("carrier")), EffectiveDate = DateOnly.Parse(reader.GetString(reader.GetOrdinal("effective_date"))), Source = reader.GetString(reader.GetOrdinal("source")), SourceDocumentId = reader.IsDBNull(reader.GetOrdinal("source_document_id")) ? null : reader.GetInt64(reader.GetOrdinal("source_document_id")) };
        await using var charges = connection.CreateCommand(); charges.CommandText = "SELECT * FROM financial_arrangement_charges WHERE financial_arrangement_id=$id;"; charges.Parameters.AddWithValue("$id", id); await using var chargeReader = await charges.ExecuteReaderAsync();
        while (await chargeReader.ReadAsync()) arrangement.Charges.Add(new FinancialArrangementCharge { FinancialArrangementChargeId = chargeReader.GetInt64(chargeReader.GetOrdinal("financial_arrangement_charge_id")), ChargeType = Enum.Parse<FinancialChargeType>(chargeReader.GetString(chargeReader.GetOrdinal("charge_type"))), Benefit = chargeReader.GetString(chargeReader.GetOrdinal("benefit")), Rate = chargeReader.IsDBNull(chargeReader.GetOrdinal("rate")) ? null : Convert.ToDecimal(chargeReader.GetValue(chargeReader.GetOrdinal("rate"))), Amount = chargeReader.IsDBNull(chargeReader.GetOrdinal("amount")) ? null : Convert.ToDecimal(chargeReader.GetValue(chargeReader.GetOrdinal("amount"))), Basis = chargeReader.GetString(chargeReader.GetOrdinal("basis")), Formula = chargeReader.GetString(chargeReader.GetOrdinal("formula")), SourceText = chargeReader.GetString(chargeReader.GetOrdinal("source_text")), SourcePage = chargeReader.IsDBNull(chargeReader.GetOrdinal("source_page")) ? null : chargeReader.GetInt32(chargeReader.GetOrdinal("source_page")) });
        return arrangement;
    }

    private static async Task<CommissionArrangement?> LoadCommissionAsync(SqliteConnection connection, long id) => await LoadSingleAsync(connection, "commission_arrangements", "commission_arrangement_id", id, reader => new CommissionArrangement { CommissionArrangementId = id, Rate = reader.IsDBNull(reader.GetOrdinal("rate")) ? null : Convert.ToDecimal(reader.GetValue(reader.GetOrdinal("rate"))), Amount = reader.IsDBNull(reader.GetOrdinal("amount")) ? null : Convert.ToDecimal(reader.GetValue(reader.GetOrdinal("amount"))), Basis = reader.GetString(reader.GetOrdinal("basis")), Source = reader.GetString(reader.GetOrdinal("source")), SourceDocumentId = reader.IsDBNull(reader.GetOrdinal("source_document_id")) ? null : reader.GetInt64(reader.GetOrdinal("source_document_id")) });
    private static async Task<PoolingArrangement?> LoadPoolingAsync(SqliteConnection connection, long id) => await LoadSingleAsync(connection, "pooling_arrangements", "pooling_arrangement_id", id, reader => new PoolingArrangement { PoolingArrangementId = id, Rate = reader.IsDBNull(reader.GetOrdinal("rate")) ? null : Convert.ToDecimal(reader.GetValue(reader.GetOrdinal("rate"))), Threshold = reader.IsDBNull(reader.GetOrdinal("threshold")) ? null : Convert.ToDecimal(reader.GetValue(reader.GetOrdinal("threshold"))), Basis = reader.GetString(reader.GetOrdinal("basis")), Benefit = reader.GetString(reader.GetOrdinal("benefit")), SourceDocumentId = reader.IsDBNull(reader.GetOrdinal("source_document_id")) ? null : reader.GetInt64(reader.GetOrdinal("source_document_id")), SourceText = reader.GetString(reader.GetOrdinal("source_text")) });
    private static async Task<T?> LoadSingleAsync<T>(SqliteConnection connection, string table, string key, long id, Func<SqliteDataReader, T> map) where T : class { await using var command = connection.CreateCommand(); command.CommandText = $"SELECT * FROM {table} WHERE {key}=$id;"; command.Parameters.AddWithValue("$id", id); await using var reader = await command.ExecuteReaderAsync(); return await reader.ReadAsync() ? map(reader) : null; }
    private sealed record VersionData(FinancialArrangement Arrangement, CommissionArrangement? Commission, PoolingArrangement? Pooling);
}

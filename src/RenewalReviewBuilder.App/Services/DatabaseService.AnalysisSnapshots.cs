using System.Collections;
using System.IO.Compression;
using System.Reflection;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.Json.Serialization.Metadata;
using Microsoft.Data.Sqlite;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed partial class DatabaseService
{
    private const int CurrentRenewalAnalysisSnapshotSchemaVersion = 1;
    private const string RenewalAnalysisSnapshotPayloadFormat = "json+gzip";

    private static readonly DefaultJsonTypeInfoResolver RenewalAnalysisSnapshotTypeInfoResolver =
        CreateRenewalAnalysisSnapshotTypeInfoResolver();

    private static readonly JsonSerializerOptions RenewalAnalysisSnapshotJsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        PreferredObjectCreationHandling = JsonObjectCreationHandling.Populate,
        TypeInfoResolver = RenewalAnalysisSnapshotTypeInfoResolver
    };

    private static DefaultJsonTypeInfoResolver CreateRenewalAnalysisSnapshotTypeInfoResolver()
    {
        var resolver = new DefaultJsonTypeInfoResolver();
        resolver.Modifiers.Add(typeInfo =>
        {
            if (typeInfo.Kind != JsonTypeInfoKind.Object)
            {
                return;
            }

            foreach (var jsonProperty in typeInfo.Properties)
            {
                if (jsonProperty.AttributeProvider is not PropertyInfo propertyInfo ||
                    propertyInfo.SetMethod is not null)
                {
                    continue;
                }

                // Initialized getter-only collections hold extracted source data and can be populated.
                // Other getter-only properties are calculated views and are recomputed after restore.
                if (IsPopulatableCollection(propertyInfo.PropertyType))
                {
                    continue;
                }

                jsonProperty.ShouldSerialize = static (_, _) => false;
            }
        });

        return resolver;
    }

    private static bool IsPopulatableCollection(Type propertyType)
    {
        if (propertyType == typeof(string) || propertyType.IsArray)
        {
            return false;
        }

        if (typeof(IList).IsAssignableFrom(propertyType) ||
            typeof(IDictionary).IsAssignableFrom(propertyType))
        {
            return true;
        }

        return propertyType.GetInterfaces().Any(interfaceType =>
            interfaceType.IsGenericType &&
            (interfaceType.GetGenericTypeDefinition() == typeof(ICollection<>) ||
             interfaceType.GetGenericTypeDefinition() == typeof(IDictionary<,>)));
    }

    private static async Task EnsureRenewalAnalysisSnapshotSchemaAsync(SqliteConnection connection)
    {
        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS renewal_analysis_snapshots (
    snapshot_id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL,
    source_document_id INTEGER NOT NULL,
    effective_date TEXT NOT NULL DEFAULT '',
    schema_version INTEGER NOT NULL,
    client_name TEXT NOT NULL DEFAULT '',
    policy_number TEXT NOT NULL DEFAULT '',
    carrier TEXT NOT NULL DEFAULT '',
    payload_format TEXT NOT NULL DEFAULT 'json+gzip',
    payload BLOB NOT NULL,
    payload_sha256 TEXT NOT NULL,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    FOREIGN KEY (client_id) REFERENCES Clients(client_id),
    FOREIGN KEY (source_document_id) REFERENCES Documents(document_id),
    UNIQUE (client_id, source_document_id, effective_date)
);
CREATE INDEX IF NOT EXISTS IX_renewal_analysis_snapshots_lookup
ON renewal_analysis_snapshots(client_id, source_document_id, effective_date);");
    }

    private static async Task UpsertRenewalAnalysisSnapshotAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        AnalysisResult analysis,
        long clientId,
        long sourceDocumentId,
        DateOnly? effectiveDate,
        string carrier,
        string clientName,
        string policyNumber)
    {
        var jsonPayload = JsonSerializer.SerializeToUtf8Bytes(analysis, RenewalAnalysisSnapshotJsonOptions);
        var compressedPayload = CompressSnapshot(jsonPayload);
        var payloadHash = Convert.ToHexString(SHA256.HashData(jsonPayload));
        var now = DateTime.UtcNow.ToString("O");

        await ExecuteNonQueryAsync(
            connection,
            transaction,
            @"
INSERT INTO renewal_analysis_snapshots (
    client_id,
    source_document_id,
    effective_date,
    schema_version,
    client_name,
    policy_number,
    carrier,
    payload_format,
    payload,
    payload_sha256,
    created_at,
    updated_at)
VALUES (
    $client_id,
    $source_document_id,
    $effective_date,
    $schema_version,
    $client_name,
    $policy_number,
    $carrier,
    $payload_format,
    $payload,
    $payload_sha256,
    $created_at,
    $updated_at)
ON CONFLICT(client_id, source_document_id, effective_date) DO UPDATE SET
    schema_version = excluded.schema_version,
    client_name = excluded.client_name,
    policy_number = excluded.policy_number,
    carrier = excluded.carrier,
    payload_format = excluded.payload_format,
    payload = excluded.payload,
    payload_sha256 = excluded.payload_sha256,
    updated_at = excluded.updated_at;",
            command =>
            {
                AddParameter(command, "$client_id", clientId);
                AddParameter(command, "$source_document_id", sourceDocumentId);
                AddParameter(command, "$effective_date", ToDateText(effectiveDate));
                AddParameter(command, "$schema_version", CurrentRenewalAnalysisSnapshotSchemaVersion);
                AddParameter(command, "$client_name", clientName);
                AddParameter(command, "$policy_number", policyNumber);
                AddParameter(command, "$carrier", carrier);
                AddParameter(command, "$payload_format", RenewalAnalysisSnapshotPayloadFormat);
                AddParameter(command, "$payload", compressedPayload);
                AddParameter(command, "$payload_sha256", payloadHash);
                AddParameter(command, "$created_at", now);
                AddParameter(command, "$updated_at", now);
            });
    }

    public async Task<RenewalAnalysisSnapshotLoadResult> GetRenewalAnalysisSnapshotAsync(
        long clientId,
        long sourceDocumentId,
        DateOnly? effectiveDate)
    {
        try
        {
            await InitializeCoreAsync();
            using var connection = await OpenConnectionAsync();
            using var command = CreateCommand(
                connection,
                null,
                @"
SELECT
    schema_version,
    payload_format,
    payload,
    payload_sha256
FROM renewal_analysis_snapshots
WHERE client_id = $client_id
  AND source_document_id = $source_document_id
  AND effective_date = $effective_date
LIMIT 1;");
            AddParameter(command, "$client_id", clientId);
            AddParameter(command, "$source_document_id", sourceDocumentId);
            AddParameter(command, "$effective_date", ToDateText(effectiveDate));

            using var reader = await command.ExecuteReaderAsync();
            if (!await reader.ReadAsync())
            {
                return RenewalAnalysisSnapshotLoadResult.NotFound(
                    "No reusable analysis snapshot exists for this saved renewal.");
            }

            var schemaVersion = reader.GetInt32(0);
            var payloadFormat = reader.GetString(1);
            if (schemaVersion != CurrentRenewalAnalysisSnapshotSchemaVersion ||
                !payloadFormat.Equals(RenewalAnalysisSnapshotPayloadFormat, StringComparison.Ordinal))
            {
                return RenewalAnalysisSnapshotLoadResult.Failure(
                    "The saved analysis snapshot uses an unsupported format. Validate this renewal once to refresh it.");
            }

            var compressedPayload = (byte[])reader[2];
            var expectedHash = reader.GetString(3);
            var jsonPayload = DecompressSnapshot(compressedPayload);
            var actualHash = Convert.ToHexString(SHA256.HashData(jsonPayload));
            if (!actualHash.Equals(expectedHash, StringComparison.OrdinalIgnoreCase))
            {
                return RenewalAnalysisSnapshotLoadResult.Failure(
                    "The saved analysis snapshot failed its integrity check. Validate this renewal once to refresh it.");
            }

            var analysis = JsonSerializer.Deserialize<AnalysisResult>(jsonPayload, RenewalAnalysisSnapshotJsonOptions);
            if (analysis is null)
            {
                return RenewalAnalysisSnapshotLoadResult.Failure(
                    "The saved analysis snapshot could not be read. Validate this renewal once to refresh it.");
            }

            analysis.DatabaseClientId = clientId;
            analysis.RenewalCase = analysis.RenewalCase;
            return RenewalAnalysisSnapshotLoadResult.Success(analysis);
        }
        catch (Exception ex)
        {
            return RenewalAnalysisSnapshotLoadResult.Failure(
                $"The saved analysis snapshot could not be loaded: {ex.Message}");
        }
    }

    private static byte[] CompressSnapshot(byte[] jsonPayload)
    {
        using var output = new MemoryStream();
        using (var gzip = new GZipStream(output, CompressionLevel.SmallestSize, leaveOpen: true))
        {
            gzip.Write(jsonPayload, 0, jsonPayload.Length);
        }

        return output.ToArray();
    }

    private static byte[] DecompressSnapshot(byte[] compressedPayload)
    {
        using var input = new MemoryStream(compressedPayload);
        using var gzip = new GZipStream(input, CompressionMode.Decompress);
        using var output = new MemoryStream();
        gzip.CopyTo(output);
        return output.ToArray();
    }
}

public sealed class RenewalAnalysisSnapshotLoadResult
{
    private RenewalAnalysisSnapshotLoadResult()
    {
    }

    public bool Found { get; init; }
    public bool Succeeded { get; init; }
    public string Message { get; init; } = "";
    public AnalysisResult? Analysis { get; init; }

    public static RenewalAnalysisSnapshotLoadResult NotFound(string message) => new()
    {
        Found = false,
        Succeeded = false,
        Message = message
    };

    public static RenewalAnalysisSnapshotLoadResult Failure(string message) => new()
    {
        Found = true,
        Succeeded = false,
        Message = message
    };

    public static RenewalAnalysisSnapshotLoadResult Success(AnalysisResult analysis) => new()
    {
        Found = true,
        Succeeded = true,
        Message = "Saved renewal analysis restored from SQLite.",
        Analysis = analysis
    };
}

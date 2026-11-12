using Microsoft.Data.Sqlite;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed partial class DatabaseService
{
    private static async Task EnsureClientPolicySchemaAsync(SqliteConnection connection)
    {
        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS client_policies (
    client_policy_id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL,
    policy_number TEXT NOT NULL COLLATE NOCASE,
    policy_role TEXT NOT NULL,
    carrier TEXT,
    verification_status TEXT NOT NULL,
    source_file_name TEXT,
    source_description TEXT,
    confirmed_by TEXT,
    confirmed_at TEXT NOT NULL,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    FOREIGN KEY (client_id) REFERENCES Clients(client_id) ON DELETE CASCADE,
    UNIQUE (client_id, policy_number)
);");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_ClientPolicies_ClientRole
ON client_policies(client_id, policy_role, verification_status); ");
        await ExecuteNonQueryAsync(connection, null, @"
CREATE INDEX IF NOT EXISTS IX_ClientPolicies_PolicyNumber
ON client_policies(policy_number); ");

        // Preserve each legacy client number as that group's primary policy.
        await ExecuteNonQueryAsync(connection, null, @"
INSERT OR IGNORE INTO client_policies (
    client_id, policy_number, policy_role, carrier, verification_status,
    source_file_name, source_description, confirmed_by, confirmed_at, created_at, updated_at
)
SELECT
    client_id, TRIM(client_number), 'Primary', carrier, 'Verified',
    NULL, 'Backfilled from Clients.client_number', 'Benefits Intelligence OS',
    COALESCE(date_updated, date_created, CURRENT_TIMESTAMP),
    COALESCE(date_created, CURRENT_TIMESTAMP),
    COALESCE(date_updated, date_created, CURRENT_TIMESTAMP)
FROM Clients
WHERE client_number IS NOT NULL AND TRIM(client_number) <> ''; ");
    }

    public async Task<IReadOnlyList<ClientPolicyRecord>> GetClientPoliciesAsync(long clientId)
    {
        await InitializeCoreAsync();
        using var connection = await OpenConnectionAsync();
        using var command = CreateCommand(connection, null, @"
SELECT *
FROM client_policies
WHERE client_id = $client_id
ORDER BY CASE policy_role WHEN 'Primary' THEN 0 WHEN 'ASO' THEN 1 ELSE 2 END, policy_number; ");
        AddParameter(command, "$client_id", clientId);

        var policies = new List<ClientPolicyRecord>();
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            policies.Add(new ClientPolicyRecord
            {
                ClientPolicyId = ReadLong(reader, "client_policy_id"),
                ClientId = ReadLong(reader, "client_id"),
                PolicyNumber = ReadString(reader, "policy_number"),
                PolicyRole = ReadString(reader, "policy_role"),
                Carrier = ReadString(reader, "carrier"),
                VerificationStatus = ReadString(reader, "verification_status"),
                SourceFileName = ReadString(reader, "source_file_name"),
                SourceDescription = ReadString(reader, "source_description"),
                ConfirmedBy = ReadString(reader, "confirmed_by"),
                ConfirmedAt = ReadDateTime(reader, "confirmed_at") ?? DateTime.MinValue,
                CreatedAt = ReadDateTime(reader, "created_at") ?? DateTime.MinValue,
                UpdatedAt = ReadDateTime(reader, "updated_at") ?? DateTime.MinValue
            });
        }

        return policies;
    }

    public async Task<DatabaseSaveResult> LinkVerifiedAsoPolicyAsync(
        long clientId,
        string policyNumber,
        string workbookClientName,
        string carrier,
        string sourceFileName,
        string confirmedBy)
    {
        var normalizedPolicyNumber = policyNumber.Trim();
        if (clientId <= 0 || string.IsNullOrWhiteSpace(normalizedPolicyNumber))
        {
            return DatabaseSaveResult.Failure("A client and an ASO policy number are required to create a verified policy link.", DatabasePath);
        }

        await InitializeCoreAsync();
        using var connection = await OpenConnectionAsync();
        using var transaction = connection.BeginTransaction();
        try
        {
            string clientName;
            using (var clientCommand = CreateCommand(connection, transaction, @"
SELECT client_name
FROM Clients
WHERE client_id = $client_id
LIMIT 1;"))
            {
                AddParameter(clientCommand, "$client_id", clientId);
                var scalar = await clientCommand.ExecuteScalarAsync();
                clientName = scalar?.ToString() ?? "";
            }

            if (string.IsNullOrWhiteSpace(clientName))
            {
                return DatabaseSaveResult.Failure("The selected client no longer exists.", DatabasePath);
            }

            if (string.IsNullOrWhiteSpace(workbookClientName) ||
                !ClientNamesMatch(clientName, workbookClientName))
            {
                return DatabaseSaveResult.Failure(
                    "The workbook client name does not match the selected client, so no policy link was created.",
                    DatabasePath);
            }

            using (var conflictingClient = CreateCommand(connection, transaction, @"
SELECT client_id
FROM Clients
WHERE client_number = $policy_number COLLATE NOCASE
  AND client_id <> $client_id
LIMIT 1;"))
            {
                AddParameter(conflictingClient, "$policy_number", normalizedPolicyNumber);
                AddParameter(conflictingClient, "$client_id", clientId);
                if ((await conflictingClient.ExecuteScalarAsync()) is not null)
                {
                    return DatabaseSaveResult.Failure(
                        $"Policy {normalizedPolicyNumber} already belongs to a different client and cannot be linked automatically.",
                        DatabasePath);
                }
            }

            using (var conflictingPolicy = CreateCommand(connection, transaction, @"
SELECT client_id
FROM client_policies
WHERE policy_number = $policy_number COLLATE NOCASE
  AND client_id <> $client_id
LIMIT 1;"))
            {
                AddParameter(conflictingPolicy, "$policy_number", normalizedPolicyNumber);
                AddParameter(conflictingPolicy, "$client_id", clientId);
                if ((await conflictingPolicy.ExecuteScalarAsync()) is not null)
                {
                    return DatabaseSaveResult.Failure(
                        $"Policy {normalizedPolicyNumber} is already linked to a different client and requires review.",
                        DatabasePath);
                }
            }

            var now = DateTime.UtcNow;
            using var upsert = CreateCommand(connection, transaction, @"
INSERT INTO client_policies (
    client_id, policy_number, policy_role, carrier, verification_status,
    source_file_name, source_description, confirmed_by, confirmed_at, created_at, updated_at
)
VALUES (
    $client_id, $policy_number, $policy_role, NULLIF($carrier, ''), $verification_status,
    NULLIF($source_file_name, ''), $source_description, NULLIF($confirmed_by, ''),
    $confirmed_at, $created_at, $updated_at
)
ON CONFLICT(client_id, policy_number) DO UPDATE SET
    policy_role = excluded.policy_role,
    carrier = COALESCE(excluded.carrier, client_policies.carrier),
    verification_status = excluded.verification_status,
    source_file_name = COALESCE(excluded.source_file_name, client_policies.source_file_name),
    source_description = excluded.source_description,
    confirmed_by = COALESCE(excluded.confirmed_by, client_policies.confirmed_by),
    confirmed_at = excluded.confirmed_at,
    updated_at = excluded.updated_at; ");
            AddParameter(upsert, "$client_id", clientId);
            AddParameter(upsert, "$policy_number", normalizedPolicyNumber);
            AddParameter(upsert, "$policy_role", ClientPolicyRoles.Aso);
            AddParameter(upsert, "$carrier", carrier.Trim());
            AddParameter(upsert, "$verification_status", ClientPolicyVerificationStatuses.Verified);
            AddParameter(upsert, "$source_file_name", sourceFileName.Trim());
            AddParameter(upsert, "$source_description", "Confirmed ASO policy for this client from Financial Intelligence import.");
            AddParameter(upsert, "$confirmed_by", confirmedBy.Trim());
            AddParameter(upsert, "$confirmed_at", ToDateTimeText(now));
            AddParameter(upsert, "$created_at", ToDateTimeText(now));
            AddParameter(upsert, "$updated_at", ToDateTimeText(now));
            await upsert.ExecuteNonQueryAsync();

            transaction.Commit();
            return DatabaseSaveResult.Success(
                $"Verified ASO policy {normalizedPolicyNumber} for {clientName}.",
                DatabasePath,
                clientId,
                insertedRows: 1);
        }
        catch (Exception exception)
        {
            transaction.Rollback();
            AppLogger.Error(exception);
            return DatabaseSaveResult.Failure($"The ASO policy link was not saved: {exception.Message}", DatabasePath);
        }
    }

    private static async Task<bool> IsVerifiedAsoPolicyAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long clientId,
        string policyNumber)
    {
        using var command = CreateCommand(connection, transaction, @"
SELECT 1
FROM client_policies
WHERE client_id = $client_id
  AND policy_number = $policy_number COLLATE NOCASE
  AND policy_role = $policy_role
  AND verification_status = $verification_status
LIMIT 1; ");
        AddParameter(command, "$client_id", clientId);
        AddParameter(command, "$policy_number", policyNumber.Trim());
        AddParameter(command, "$policy_role", ClientPolicyRoles.Aso);
        AddParameter(command, "$verification_status", ClientPolicyVerificationStatuses.Verified);
        return (await command.ExecuteScalarAsync()) is not null;
    }

    private static bool ClientNamesMatch(string left, string right)
    {
        static string Normalize(string value) => new(value
            .Where(char.IsLetterOrDigit)
            .Select(char.ToUpperInvariant)
            .ToArray());

        var normalizedLeft = Normalize(left);
        var normalizedRight = Normalize(right);
        return normalizedLeft.Length > 0 && normalizedLeft.Equals(normalizedRight, StringComparison.Ordinal);
    }
}

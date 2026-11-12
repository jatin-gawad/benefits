using System.Globalization;
using Microsoft.Data.Sqlite;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;

namespace RenewalReviewBuilder.App.Repositories;

public sealed class BenefitAdministrationRepository
{
    private readonly string _databasePath;

    public BenefitAdministrationRepository(string? databasePath = null)
    {
        _databasePath = string.IsNullOrWhiteSpace(databasePath)
            ? AppPaths.BenefitsIntelligenceDatabasePath
            : databasePath;
    }

    public Task<List<ClientBenefitAdministrationRecord>> GetByClientIdAsync(long clientId)
    {
        return QueryAsync("a.ClientId = $client_id", command => Add(command, "$client_id", clientId));
    }

    public Task<List<ClientBenefitAdministrationRecord>> GetByPlanDesignVersionIdAsync(long planDesignVersionId)
    {
        return QueryAsync(
            "a.PlanDesignVersionId = $plan_design_version_id",
            command => Add(command, "$plan_design_version_id", planDesignVersionId));
    }

    public async Task<ClientBenefitAdministrationRecord?> GetByIdAsync(long administrationId)
    {
        var records = await QueryAsync(
            "a.ClientBenefitAdministrationId = $administration_id",
            command => Add(command, "$administration_id", administrationId));
        return records.FirstOrDefault();
    }

    public Task<List<ClientBenefitAdministrationRecord>> GetActiveAsOfAsync(long clientId, DateTime asOfDate)
    {
        return QueryAsync(
            """
            a.ClientId = $client_id
            AND a.IsActive = 1
            AND (a.EffectiveDate IS NULL OR a.EffectiveDate <= $as_of_date)
            AND (a.TerminationDate IS NULL OR a.TerminationDate >= $as_of_date)
            """,
            command =>
            {
                Add(command, "$client_id", clientId);
                Add(command, "$as_of_date", DateValue(asOfDate));
            });
    }

    public async Task<List<OrganizationRecord>> GetOrganizationsAsync(bool activeOnly = true)
    {
        await using var connection = await OpenConnectionAsync();
        await using var command = connection.CreateCommand();
        command.CommandText = $"""
            SELECT OrganizationId, CanonicalName, DisplayName, OrganizationTypeCode, NormalizedName, IsActive
            FROM Organization
            {(activeOnly ? "WHERE IsActive = 1" : "")}
            ORDER BY DisplayName COLLATE NOCASE;
            """;
        var records = new List<OrganizationRecord>();
        await using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            records.Add(new OrganizationRecord
            {
                OrganizationId = ReadLong(reader, "OrganizationId"),
                CanonicalName = ReadString(reader, "CanonicalName"),
                DisplayName = ReadString(reader, "DisplayName"),
                OrganizationTypeCode = ReadString(reader, "OrganizationTypeCode"),
                NormalizedName = ReadString(reader, "NormalizedName"),
                IsActive = ReadBool(reader, "IsActive")
            });
        }

        return records;
    }

    public Task<bool> UpdateReviewStatusAsync(long administrationId, string reviewStatusCode)
    {
        return ExecuteUpdateAsync(
            """
            UPDATE ClientBenefitAdministration
            SET ReviewStatusCode = $review_status, UpdatedAtUtc = $now
            WHERE ClientBenefitAdministrationId = $administration_id;
            """,
            command =>
            {
                Add(command, "$review_status", reviewStatusCode);
                Add(command, "$administration_id", administrationId);
                Add(command, "$now", DateTime.UtcNow.ToString("O"));
            });
    }

    public Task<bool> CorrectPolicyNumberAsync(long administrationId, string policyNumber)
    {
        return ExecuteUpdateAsync(
            """
            UPDATE ClientBenefitAdministration
            SET PolicyNumber = $policy_number, ReviewStatusCode = 'NEEDS_REVIEW', UpdatedAtUtc = $now
            WHERE ClientBenefitAdministrationId = $administration_id;
            """,
            command =>
            {
                Add(command, "$policy_number", NullIfBlank(policyNumber));
                Add(command, "$administration_id", administrationId);
                Add(command, "$now", DateTime.UtcNow.ToString("O"));
            });
    }

    public Task<bool> CorrectInsurerAsync(long administrationId, long? insurerOrganizationId, string sourceDisplayName)
    {
        return ExecuteUpdateAsync(
            """
            UPDATE ClientBenefitAdministration
            SET InsurerOrganizationId = $insurer_id,
                InsurerSourceDisplayName = $source_display_name,
                ReviewStatusCode = 'NEEDS_REVIEW',
                UpdatedAtUtc = $now
            WHERE ClientBenefitAdministrationId = $administration_id;
            """,
            command =>
            {
                Add(command, "$insurer_id", insurerOrganizationId);
                Add(command, "$source_display_name", NullIfBlank(sourceDisplayName));
                Add(command, "$administration_id", administrationId);
                Add(command, "$now", DateTime.UtcNow.ToString("O"));
            });
    }

    public async Task<long> AddProviderAssignmentAsync(BenefitServiceProviderAssignmentRecord assignment)
    {
        await using var connection = await OpenConnectionAsync();
        await using var transaction = connection.BeginTransaction();
        var id = await InsertProviderAssignmentAsync(connection, transaction, assignment);
        await transaction.CommitAsync();
        return id;
    }

    public Task<bool> UpdateProviderAssignmentAsync(BenefitServiceProviderAssignmentRecord assignment)
    {
        return ExecuteUpdateAsync(
            """
            UPDATE BenefitServiceProviderAssignment
            SET ProviderOrganizationId = $provider_id,
                ProviderRoleCode = $provider_role,
                SourceDisplayName = $source_display_name,
                IsPrimary = $is_primary,
                EffectiveDate = $effective_date,
                TerminationDate = $termination_date,
                ReviewStatusCode = 'NEEDS_REVIEW',
                UpdatedAtUtc = $now
            WHERE BenefitServiceProviderAssignmentId = $assignment_id;
            """,
            command =>
            {
                Add(command, "$provider_id", assignment.ProviderOrganizationId);
                Add(command, "$provider_role", assignment.ProviderRoleCode);
                Add(command, "$source_display_name", NullIfBlank(assignment.SourceDisplayName));
                Add(command, "$is_primary", assignment.IsPrimary ? 1 : 0);
                Add(command, "$effective_date", DateValue(assignment.EffectiveDate));
                Add(command, "$termination_date", DateValue(assignment.TerminationDate));
                Add(command, "$assignment_id", assignment.BenefitServiceProviderAssignmentId);
                Add(command, "$now", DateTime.UtcNow.ToString("O"));
            });
    }

    public Task<bool> DeactivateProviderAssignmentAsync(long assignmentId)
    {
        return ExecuteUpdateAsync(
            """
            UPDATE BenefitServiceProviderAssignment
            SET IsActive = 0, UpdatedAtUtc = $now
            WHERE BenefitServiceProviderAssignmentId = $assignment_id;
            """,
            command =>
            {
                Add(command, "$assignment_id", assignmentId);
                Add(command, "$now", DateTime.UtcNow.ToString("O"));
            });
    }

    public Task<bool> EndDateProviderAssignmentAsync(long assignmentId, DateTime terminationDate)
    {
        return ExecuteUpdateAsync(
            """
            UPDATE BenefitServiceProviderAssignment
            SET TerminationDate = $termination_date, IsActive = 0, UpdatedAtUtc = $now
            WHERE BenefitServiceProviderAssignmentId = $assignment_id;
            """,
            command =>
            {
                Add(command, "$termination_date", DateValue(terminationDate));
                Add(command, "$assignment_id", assignmentId);
                Add(command, "$now", DateTime.UtcNow.ToString("O"));
            });
    }

    public async Task<List<ClientBenefitAdministrationRecord>> FindLikelyDuplicatesAsync(
        long clientId,
        long benefitTypeId,
        string policyNumber)
    {
        return await QueryAsync(
            """
            a.ClientId = $client_id
            AND a.BenefitTypeId = $benefit_type_id
            AND COALESCE(a.PolicyNumber, '') = $policy_number COLLATE NOCASE
            """,
            command =>
            {
                Add(command, "$client_id", clientId);
                Add(command, "$benefit_type_id", benefitTypeId);
                Add(command, "$policy_number", policyNumber.Trim());
            });
    }

    internal static async Task<long?> ResolveBenefitTypeIdAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        string benefitTypeCode)
    {
        await using var command = CreateCommand(connection, transaction, """
            SELECT benefit_id
            FROM Benefit
            WHERE benefit_line = $benefit_line COLLATE NOCASE
            ORDER BY benefit_id
            LIMIT 1;
            """);
        Add(command, "$benefit_line", benefitTypeCode);
        return ToNullableLong(await command.ExecuteScalarAsync());
    }

    internal static async Task<bool> ClientExistsAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long clientId)
    {
        await using var command = CreateCommand(
            connection,
            transaction,
            "SELECT COUNT(*) FROM Clients WHERE client_id = $client_id;");
        Add(command, "$client_id", clientId);
        return Convert.ToInt32(await command.ExecuteScalarAsync(), CultureInfo.InvariantCulture) > 0;
    }

    internal static async Task<bool> SourceDocumentBelongsToClientAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long sourceDocumentId,
        long clientId)
    {
        await using var command = CreateCommand(connection, transaction, """
            SELECT COUNT(*)
            FROM Documents
            WHERE document_id = $document_id AND client_id = $client_id;
            """);
        Add(command, "$document_id", sourceDocumentId);
        Add(command, "$client_id", clientId);
        return Convert.ToInt32(await command.ExecuteScalarAsync(), CultureInfo.InvariantCulture) > 0;
    }

    internal static async Task<long?> ResolveOrganizationIdAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        string normalizedName)
    {
        await using var command = CreateCommand(connection, transaction, """
            SELECT OrganizationId
            FROM Organization
            WHERE NormalizedName = $normalized_name COLLATE NOCASE
              AND IsActive = 1
            ORDER BY OrganizationId
            LIMIT 1;
            """);
        Add(command, "$normalized_name", normalizedName);
        return ToNullableLong(await command.ExecuteScalarAsync());
    }

    internal static async Task<ClientBenefitAdministrationRecord?> FindActiveIdentityAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long clientId,
        long? planDesignVersionId,
        long benefitTypeId,
        string productName)
    {
        await using var command = CreateCommand(connection, transaction, """
            SELECT ClientBenefitAdministrationId, ClientId, PlanDesignVersionId, BenefitTypeId,
                   COALESCE(ProductName, '') AS ProductName,
                   COALESCE(PolicyNumber, '') AS PolicyNumber,
                   InsurerOrganizationId, EffectiveDate, TerminationDate,
                   ReviewStatusCode, IsActive
            FROM ClientBenefitAdministration
            WHERE ClientId = $client_id
              AND COALESCE(PlanDesignVersionId, 0) = COALESCE($plan_design_version_id, 0)
              AND BenefitTypeId = $benefit_type_id
              AND LOWER(TRIM(COALESCE(ProductName, ''))) = LOWER(TRIM($product_name))
              AND IsActive = 1
            ORDER BY ClientBenefitAdministrationId DESC
            LIMIT 1;
            """);
        Add(command, "$client_id", clientId);
        Add(command, "$plan_design_version_id", planDesignVersionId);
        Add(command, "$benefit_type_id", benefitTypeId);
        Add(command, "$product_name", productName);
        await using var reader = await command.ExecuteReaderAsync();
        if (!await reader.ReadAsync())
        {
            return null;
        }

        return new ClientBenefitAdministrationRecord
        {
            ClientBenefitAdministrationId = ReadLong(reader, "ClientBenefitAdministrationId"),
            ClientId = ReadLong(reader, "ClientId"),
            PlanDesignVersionId = ReadNullableLong(reader, "PlanDesignVersionId"),
            BenefitTypeId = ReadLong(reader, "BenefitTypeId"),
            ProductName = ReadString(reader, "ProductName"),
            PolicyNumber = ReadString(reader, "PolicyNumber"),
            InsurerOrganizationId = ReadNullableLong(reader, "InsurerOrganizationId"),
            EffectiveDate = ReadDate(reader, "EffectiveDate"),
            TerminationDate = ReadDate(reader, "TerminationDate"),
            ReviewStatusCode = ReadString(reader, "ReviewStatusCode"),
            IsActive = ReadBool(reader, "IsActive")
        };
    }

    internal static async Task<long> InsertAdministrationAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        ClientBenefitAdministrationRecord record,
        string insurerSourceDisplayName)
    {
        await using var command = CreateCommand(connection, transaction, """
            INSERT INTO ClientBenefitAdministration (
                ClientId, PlanDesignVersionId, BenefitTypeId, BenefitDisplayName, ProductName,
                PolicyNumber, InsurerOrganizationId, InsurerSourceDisplayName, IsOptional,
                EffectiveDate, TerminationDate, SourceDocumentId, SourcePageNumber, SourceText,
                ExtractionConfidence, ReviewStatusCode, IsActive,
                CopiedFromClientBenefitAdministrationId, CreatedAtUtc, UpdatedAtUtc)
            VALUES (
                $client_id, $plan_design_version_id, $benefit_type_id, $benefit_display_name, $product_name,
                $policy_number, $insurer_id, $insurer_source_display_name, $is_optional,
                $effective_date, $termination_date, $source_document_id, $source_page_number, $source_text,
                $confidence, $review_status, $is_active,
                $copied_from_id, $now, $now);
            SELECT last_insert_rowid();
            """);
        Add(command, "$client_id", record.ClientId);
        Add(command, "$plan_design_version_id", record.PlanDesignVersionId);
        Add(command, "$benefit_type_id", record.BenefitTypeId);
        Add(command, "$benefit_display_name", NullIfBlank(record.BenefitDisplayName));
        Add(command, "$product_name", NullIfBlank(record.ProductName));
        Add(command, "$policy_number", NullIfBlank(record.PolicyNumber));
        Add(command, "$insurer_id", record.InsurerOrganizationId);
        Add(command, "$insurer_source_display_name", NullIfBlank(insurerSourceDisplayName));
        Add(command, "$is_optional", record.IsOptional ? 1 : 0);
        Add(command, "$effective_date", DateValue(record.EffectiveDate));
        Add(command, "$termination_date", DateValue(record.TerminationDate));
        Add(command, "$source_document_id", record.SourceDocumentId);
        Add(command, "$source_page_number", record.SourcePageNumber);
        Add(command, "$source_text", NullIfBlank(record.SourceText));
        Add(command, "$confidence", record.ExtractionConfidence);
        Add(command, "$review_status", record.ReviewStatusCode);
        Add(command, "$is_active", record.IsActive ? 1 : 0);
        Add(command, "$copied_from_id", record.CopiedFromClientBenefitAdministrationId);
        Add(command, "$now", DateTime.UtcNow.ToString("O"));
        return Convert.ToInt64(await command.ExecuteScalarAsync(), CultureInfo.InvariantCulture);
    }

    internal static async Task RefreshExistingSourceAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long administrationId,
        BenefitAdministrationExtractionCandidate source,
        string sourceLineage)
    {
        await using var command = CreateCommand(connection, transaction, """
            UPDATE ClientBenefitAdministration
            SET BenefitDisplayName = $benefit_display_name,
                SourceDocumentId = $source_document_id,
                SourcePageNumber = $source_page_number,
                SourceText = $source_text,
                ExtractionConfidence = $confidence,
                UpdatedAtUtc = $now
            WHERE ClientBenefitAdministrationId = $administration_id;
            """);
        Add(command, "$benefit_display_name", source.BenefitDisplayName);
        Add(command, "$source_document_id", source.SourceDocumentId);
        Add(command, "$source_page_number", source.SourcePageNumber);
        Add(command, "$source_text", sourceLineage);
        Add(command, "$confidence", source.ExtractionConfidence);
        Add(command, "$administration_id", administrationId);
        Add(command, "$now", DateTime.UtcNow.ToString("O"));
        await command.ExecuteNonQueryAsync();
    }

    internal static async Task EndDateAdministrationAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long administrationId,
        DateTime terminationDate)
    {
        await using var command = CreateCommand(connection, transaction, """
            UPDATE ClientBenefitAdministration
            SET TerminationDate = $termination_date, IsActive = 0, UpdatedAtUtc = $now
            WHERE ClientBenefitAdministrationId = $administration_id;

            UPDATE BenefitServiceProviderAssignment
            SET TerminationDate = COALESCE(TerminationDate, $termination_date),
                IsActive = 0,
                UpdatedAtUtc = $now
            WHERE ClientBenefitAdministrationId = $administration_id AND IsActive = 1;
            """);
        Add(command, "$termination_date", DateValue(terminationDate));
        Add(command, "$administration_id", administrationId);
        Add(command, "$now", DateTime.UtcNow.ToString("O"));
        await command.ExecuteNonQueryAsync();
    }

    internal static async Task<HashSet<(long ProviderId, string Role)>> GetActiveProviderKeysAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long administrationId)
    {
        await using var command = CreateCommand(connection, transaction, """
            SELECT ProviderOrganizationId, ProviderRoleCode
            FROM BenefitServiceProviderAssignment
            WHERE ClientBenefitAdministrationId = $administration_id AND IsActive = 1;
            """);
        Add(command, "$administration_id", administrationId);
        var keys = new HashSet<(long, string)>();
        await using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            keys.Add((ReadLong(reader, "ProviderOrganizationId"), ReadString(reader, "ProviderRoleCode")));
        }

        return keys;
    }

    internal static async Task<long> InsertProviderAssignmentAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        BenefitServiceProviderAssignmentRecord assignment)
    {
        await using var command = CreateCommand(connection, transaction, """
            INSERT INTO BenefitServiceProviderAssignment (
                ClientBenefitAdministrationId, ProviderOrganizationId, ProviderRoleCode,
                SourceDisplayName, IsPrimary, EffectiveDate, TerminationDate,
                SourceDocumentId, SourcePageNumber, ReviewStatusCode, IsActive,
                CreatedAtUtc, UpdatedAtUtc)
            VALUES (
                $administration_id, $provider_id, $provider_role,
                $source_display_name, $is_primary, $effective_date, $termination_date,
                $source_document_id, $source_page_number, $review_status, $is_active,
                $now, $now);
            SELECT last_insert_rowid();
            """);
        Add(command, "$administration_id", assignment.ClientBenefitAdministrationId);
        Add(command, "$provider_id", assignment.ProviderOrganizationId);
        Add(command, "$provider_role", assignment.ProviderRoleCode);
        Add(command, "$source_display_name", NullIfBlank(assignment.SourceDisplayName));
        Add(command, "$is_primary", assignment.IsPrimary ? 1 : 0);
        Add(command, "$effective_date", DateValue(assignment.EffectiveDate));
        Add(command, "$termination_date", DateValue(assignment.TerminationDate));
        Add(command, "$source_document_id", assignment.SourceDocumentId);
        Add(command, "$source_page_number", assignment.SourcePageNumber);
        Add(command, "$review_status", assignment.ReviewStatusCode);
        Add(command, "$is_active", assignment.IsActive ? 1 : 0);
        Add(command, "$now", DateTime.UtcNow.ToString("O"));
        return Convert.ToInt64(await command.ExecuteScalarAsync(), CultureInfo.InvariantCulture);
    }

    private async Task<List<ClientBenefitAdministrationRecord>> QueryAsync(
        string whereClause,
        Action<SqliteCommand> addParameters)
    {
        await using var connection = await OpenConnectionAsync();
        await using var command = connection.CreateCommand();
        command.CommandText = $"""
            SELECT a.*, b.benefit_line AS BenefitTypeCode,
                   COALESCE(insurer.DisplayName, '') AS InsurerName,
                   COALESCE(insurer.OrganizationTypeCode, '') AS InsurerOrganizationTypeCode,
                   COALESCE(d.file_name, '') AS SourceDocumentName,
                   COALESCE(d.file_path, '') AS SourceDocumentPath
            FROM ClientBenefitAdministration a
            INNER JOIN Benefit b ON b.benefit_id = a.BenefitTypeId
            LEFT JOIN Organization insurer ON insurer.OrganizationId = a.InsurerOrganizationId
            LEFT JOIN Documents d ON d.document_id = a.SourceDocumentId
            WHERE {whereClause}
            ORDER BY a.IsActive DESC, b.normalized_name COLLATE NOCASE,
                     a.EffectiveDate DESC, a.ClientBenefitAdministrationId DESC;
            """;
        addParameters(command);
        var records = new List<ClientBenefitAdministrationRecord>();
        await using (var reader = await command.ExecuteReaderAsync())
        {
            while (await reader.ReadAsync())
            {
                records.Add(ReadAdministration(reader));
            }
        }

        if (records.Count > 0)
        {
            await AttachProvidersAsync(connection, records);
        }

        return records;
    }

    private static async Task AttachProvidersAsync(
        SqliteConnection connection,
        IReadOnlyList<ClientBenefitAdministrationRecord> records)
    {
        var ids = records.Select(record => record.ClientBenefitAdministrationId).ToHashSet();
        var byId = records.ToDictionary(record => record.ClientBenefitAdministrationId);
        await using var command = connection.CreateCommand();
        command.CommandText = """
            SELECT p.*, o.DisplayName AS ProviderName
            FROM BenefitServiceProviderAssignment p
            INNER JOIN Organization o ON o.OrganizationId = p.ProviderOrganizationId
            ORDER BY p.IsPrimary DESC, o.DisplayName COLLATE NOCASE;
            """;
        await using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            var administrationId = ReadLong(reader, "ClientBenefitAdministrationId");
            if (!ids.Contains(administrationId) || !byId.TryGetValue(administrationId, out var administration))
            {
                continue;
            }

            administration.ServiceProviders.Add(new BenefitServiceProviderAssignmentRecord
            {
                BenefitServiceProviderAssignmentId = ReadLong(reader, "BenefitServiceProviderAssignmentId"),
                ClientBenefitAdministrationId = administrationId,
                ProviderOrganizationId = ReadLong(reader, "ProviderOrganizationId"),
                ProviderName = ReadString(reader, "ProviderName"),
                ProviderRoleCode = ReadString(reader, "ProviderRoleCode"),
                SourceDisplayName = ReadString(reader, "SourceDisplayName"),
                IsPrimary = ReadBool(reader, "IsPrimary"),
                EffectiveDate = ReadDate(reader, "EffectiveDate"),
                TerminationDate = ReadDate(reader, "TerminationDate"),
                SourceDocumentId = ReadNullableLong(reader, "SourceDocumentId"),
                SourcePageNumber = ReadNullableInt(reader, "SourcePageNumber"),
                ReviewStatusCode = ReadString(reader, "ReviewStatusCode"),
                IsActive = ReadBool(reader, "IsActive"),
                CreatedAtUtc = ReadDateTime(reader, "CreatedAtUtc"),
                UpdatedAtUtc = ReadDateTime(reader, "UpdatedAtUtc")
            });
        }
    }

    private static ClientBenefitAdministrationRecord ReadAdministration(SqliteDataReader reader)
    {
        return new ClientBenefitAdministrationRecord
        {
            ClientBenefitAdministrationId = ReadLong(reader, "ClientBenefitAdministrationId"),
            ClientId = ReadLong(reader, "ClientId"),
            PlanDesignVersionId = ReadNullableLong(reader, "PlanDesignVersionId"),
            BenefitTypeId = ReadLong(reader, "BenefitTypeId"),
            BenefitTypeCode = ReadString(reader, "BenefitTypeCode"),
            BenefitDisplayName = ReadString(reader, "BenefitDisplayName"),
            ProductName = ReadString(reader, "ProductName"),
            PolicyNumber = ReadString(reader, "PolicyNumber"),
            InsurerOrganizationId = ReadNullableLong(reader, "InsurerOrganizationId"),
            InsurerName = ReadString(reader, "InsurerName"),
            InsurerOrganizationTypeCode = ReadString(reader, "InsurerOrganizationTypeCode"),
            IsOptional = ReadBool(reader, "IsOptional"),
            EffectiveDate = ReadDate(reader, "EffectiveDate"),
            TerminationDate = ReadDate(reader, "TerminationDate"),
            SourceDocumentId = ReadNullableLong(reader, "SourceDocumentId"),
            SourceDocumentName = ReadString(reader, "SourceDocumentName"),
            SourceDocumentPath = ReadString(reader, "SourceDocumentPath"),
            SourcePageNumber = ReadNullableInt(reader, "SourcePageNumber"),
            SourceText = ReadString(reader, "SourceText"),
            ExtractionConfidence = ReadNullableDecimal(reader, "ExtractionConfidence"),
            ReviewStatusCode = ReadString(reader, "ReviewStatusCode"),
            IsActive = ReadBool(reader, "IsActive"),
            CopiedFromClientBenefitAdministrationId = ReadNullableLong(reader, "CopiedFromClientBenefitAdministrationId"),
            CreatedAtUtc = ReadDateTime(reader, "CreatedAtUtc"),
            UpdatedAtUtc = ReadDateTime(reader, "UpdatedAtUtc")
        };
    }

    private async Task<bool> ExecuteUpdateAsync(string sql, Action<SqliteCommand> addParameters)
    {
        await using var connection = await OpenConnectionAsync();
        await using var command = connection.CreateCommand();
        command.CommandText = sql;
        addParameters(command);
        return await command.ExecuteNonQueryAsync() > 0;
    }

    private async Task<SqliteConnection> OpenConnectionAsync()
    {
        var status = new DatabaseService(_databasePath).Initialize();
        if (!status.Succeeded)
        {
            throw new InvalidOperationException(status.Message);
        }

        var connection = new SqliteConnection($"Data Source={_databasePath}");
        await connection.OpenAsync();
        await using var command = connection.CreateCommand();
        command.CommandText = "PRAGMA foreign_keys = ON;";
        await command.ExecuteNonQueryAsync();
        return connection;
    }

    private static SqliteCommand CreateCommand(SqliteConnection connection, SqliteTransaction transaction, string sql)
    {
        var command = connection.CreateCommand();
        command.Transaction = transaction;
        command.CommandText = sql;
        return command;
    }

    private static void Add(SqliteCommand command, string name, object? value)
    {
        command.Parameters.AddWithValue(name, value ?? DBNull.Value);
    }

    private static object? NullIfBlank(string? value) => string.IsNullOrWhiteSpace(value) ? null : value.Trim();
    private static object? DateValue(DateTime? value) => value?.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
    private static long? ToNullableLong(object? value) => value is null || value == DBNull.Value ? null : Convert.ToInt64(value, CultureInfo.InvariantCulture);
    private static long ReadLong(SqliteDataReader reader, string name) => reader.GetInt64(reader.GetOrdinal(name));
    private static long? ReadNullableLong(SqliteDataReader reader, string name) => reader.IsDBNull(reader.GetOrdinal(name)) ? null : reader.GetInt64(reader.GetOrdinal(name));
    private static int? ReadNullableInt(SqliteDataReader reader, string name) => reader.IsDBNull(reader.GetOrdinal(name)) ? null : reader.GetInt32(reader.GetOrdinal(name));
    private static string ReadString(SqliteDataReader reader, string name) => reader.IsDBNull(reader.GetOrdinal(name)) ? "" : reader.GetString(reader.GetOrdinal(name));
    private static bool ReadBool(SqliteDataReader reader, string name) => !reader.IsDBNull(reader.GetOrdinal(name)) && reader.GetInt64(reader.GetOrdinal(name)) != 0;
    private static decimal? ReadNullableDecimal(SqliteDataReader reader, string name) => reader.IsDBNull(reader.GetOrdinal(name)) ? null : Convert.ToDecimal(reader.GetDouble(reader.GetOrdinal(name)), CultureInfo.InvariantCulture);
    private static DateTime? ReadDate(SqliteDataReader reader, string name) => DateTime.TryParse(ReadString(reader, name), CultureInfo.InvariantCulture, DateTimeStyles.None, out var value) ? value : null;
    private static DateTime ReadDateTime(SqliteDataReader reader, string name) => DateTime.TryParse(ReadString(reader, name), CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var value) ? value : default;
}

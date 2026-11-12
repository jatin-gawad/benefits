using System.Globalization;
using System.Text.Json;
using Microsoft.Data.Sqlite;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed class ProposalCanonicalProfileService
{
    private static readonly string[] SourceTableNames =
    [
        "Clients",
        "carriers",
        "proposal_files",
        "proposal_summary",
        "benefits",
        "proposal_benefit_lines",
        "proposal_rates",
        "proposal_premiums",
        "proposal_plan_design",
        "proposal_booklet_validation",
        "proposal_demographics",
        "proposal_underwriting_assumptions",
        "proposal_pooling_retention",
        "proposal_exceptions",
        "proposal_export_audit"
    ];

    private readonly string _databasePath;

    public ProposalCanonicalProfileService(string? databasePath = null)
    {
        _databasePath = string.IsNullOrWhiteSpace(databasePath)
            ? AppPaths.BenefitsIntelligenceDatabasePath
            : databasePath;
    }

    public string DatabasePath => _databasePath;

    public async Task<ProposalCanonicalProfile> BuildAsync(long clientId)
    {
        if (!File.Exists(_databasePath))
        {
            var profile = CreateProfile(clientId);
            AddIssue(
                profile,
                "database_missing",
                ProposalCanonicalIssueSeverities.Error,
                "Benefits Intelligence database was not found.");
            return profile;
        }

        await using var connection = await OpenReadOnlyConnectionAsync();
        var existingTables = await GetExistingTablesAsync(connection);
        return await BuildProfileAsync(connection, existingTables, clientId);
    }

    public async Task<IReadOnlyList<ProposalCanonicalProfile>> BuildAllAsync()
    {
        if (!File.Exists(_databasePath))
        {
            return [];
        }

        await using var connection = await OpenReadOnlyConnectionAsync();
        var existingTables = await GetExistingTablesAsync(connection);
        var clientIds = await ReadSavedProposalClientIdsAsync(connection, existingTables);
        var profiles = new List<ProposalCanonicalProfile>();
        foreach (var clientId in clientIds)
        {
            var profile = await BuildProfileAsync(connection, existingTables, clientId);
            if (profile.HasProposals)
            {
                profiles.Add(profile);
            }
        }

        return profiles
            .OrderBy(profile => profile.ClientName, StringComparer.OrdinalIgnoreCase)
            .ThenBy(profile => profile.ClientId)
            .ToList();
    }

    private static ProposalCanonicalProfile CreateProfile(long clientId) => new()
    {
        ClientId = clientId,
        BuiltAtUtc = DateTime.UtcNow
    };

    private static async Task<ProposalCanonicalProfile> BuildProfileAsync(
        SqliteConnection connection,
        HashSet<string> existingTables,
        long clientId)
    {
        var profile = CreateProfile(clientId);
        var client = await ReadClientAsync(connection, existingTables, clientId);
        profile.ClientName = client.Found ? client.ClientName : "";
        profile.ClientNumber = client.ClientNumber;

        var clientKeys = BuildClientKeys(clientId, client);
        var proposalIds = await ReadProposalIdsAsync(connection, existingTables, clientKeys, client.ClientName);
        foreach (var proposalId in proposalIds.OrderBy(value => value))
        {
            profile.Proposals.Add(await ReadProposalAsync(connection, existingTables, proposalId, client));
        }

        var proposalsById = profile.Proposals.ToDictionary(proposal => proposal.ProposalId);
        await LoadDetailRowsAsync(connection, existingTables, "benefits", proposalsById, ReadBenefit, (proposal, row) => proposal.Benefits.Add(row));
        await LoadDetailRowsAsync(connection, existingTables, "proposal_benefit_lines", proposalsById, ReadBenefitLine, (proposal, row) => proposal.BenefitLines.Add(row));
        await LoadDetailRowsAsync(connection, existingTables, "proposal_rates", proposalsById, ReadRate, (proposal, row) => proposal.Rates.Add(row));
        await LoadDetailRowsAsync(connection, existingTables, "proposal_premiums", proposalsById, ReadPremium, (proposal, row) => proposal.Premiums.Add(row));
        await LoadDetailRowsAsync(connection, existingTables, "proposal_plan_design", proposalsById, reader => ReadPlanDesign(reader, "proposal_plan_design"), (proposal, row) => proposal.PlanDesignFacts.Add(row));
        await LoadDetailRowsAsync(connection, existingTables, "proposal_booklet_validation", proposalsById, reader => ReadPlanDesign(reader, "proposal_booklet_validation"), (proposal, row) => proposal.ValidationFacts.Add(row));
        await LoadDetailRowsAsync(connection, existingTables, "proposal_demographics", proposalsById, ReadDemographic, (proposal, row) => proposal.Demographics.Add(row));
        await LoadDetailRowsAsync(connection, existingTables, "proposal_underwriting_assumptions", proposalsById, ReadUnderwritingAssumption, (proposal, row) => proposal.UnderwritingAssumptions.Add(row));
        await LoadDetailRowsAsync(connection, existingTables, "proposal_pooling_retention", proposalsById, ReadPoolingRetention, (proposal, row) => proposal.PoolingRetention.Add(row));
        await LoadDetailRowsAsync(connection, existingTables, "proposal_exceptions", proposalsById, ReadException, (proposal, row) => proposal.Exceptions.Add(row));
        await LoadDetailRowsAsync(connection, existingTables, "proposal_export_audit", proposalsById, ReadAudit, (proposal, row) => proposal.AuditTrail.Add(row));

        if (string.IsNullOrWhiteSpace(profile.ClientName))
        {
            profile.ClientName = profile.Proposals
                .Select(proposal => proposal.ClientName)
                .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)) ?? $"Client #{clientId.ToString(CultureInfo.InvariantCulture)}";
        }

        foreach (var tableName in SourceTableNames)
        {
            profile.Sources.Add(await BuildSourceUsageAsync(
                connection,
                existingTables,
                tableName,
                client,
                profile.Proposals,
                proposalIds));
        }

        AddProfileIssues(profile);
        return profile;
    }

    private async Task<SqliteConnection> OpenReadOnlyConnectionAsync()
    {
        var builder = new SqliteConnectionStringBuilder
        {
            DataSource = _databasePath,
            Mode = SqliteOpenMode.ReadOnly,
            Cache = SqliteCacheMode.Shared
        };
        var connection = new SqliteConnection(builder.ToString());
        await connection.OpenAsync();

        using var command = connection.CreateCommand();
        command.CommandText = "PRAGMA query_only = ON;";
        await command.ExecuteNonQueryAsync();
        return connection;
    }

    private static async Task<HashSet<string>> GetExistingTablesAsync(SqliteConnection connection)
    {
        using var command = connection.CreateCommand();
        command.CommandText = "SELECT name FROM sqlite_master WHERE type = 'table';";

        var tables = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            tables.Add(reader.GetString(0));
        }

        return tables;
    }

    private static async Task<HashSet<string>> GetColumnsAsync(SqliteConnection connection, string tableName)
    {
        using var command = connection.CreateCommand();
        command.CommandText = $"PRAGMA table_info([{tableName}]);";

        var columns = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            columns.Add(ReadString(reader, "name"));
        }

        return columns;
    }

    private static async Task<ClientLookup> ReadClientAsync(
        SqliteConnection connection,
        HashSet<string> existingTables,
        long clientId)
    {
        if (!existingTables.Contains("Clients"))
        {
            return new ClientLookup(clientId, "", "", false);
        }

        var columns = await GetColumnsAsync(connection, "Clients");
        var idColumn = columns.Contains("client_id") ? "client_id" : columns.Contains("id") ? "id" : "";
        if (string.IsNullOrWhiteSpace(idColumn))
        {
            return new ClientLookup(clientId, "", "", false);
        }

        using var command = connection.CreateCommand();
        command.CommandText = $"SELECT * FROM Clients WHERE CAST([{idColumn}] AS TEXT) = $client_id LIMIT 1;";
        AddParameter(command, "$client_id", clientId.ToString(CultureInfo.InvariantCulture));
        using var reader = await command.ExecuteReaderAsync();
        if (!await reader.ReadAsync())
        {
            return new ClientLookup(clientId, "", "", false);
        }

        return new ClientLookup(
            ReadLong(reader, idColumn) ?? clientId,
            FirstNonBlank(ReadString(reader, "client_name"), $"Client #{clientId.ToString(CultureInfo.InvariantCulture)}"),
            ReadString(reader, "client_number"),
            true);
    }

    private static async Task<IReadOnlyList<long>> ReadSavedProposalClientIdsAsync(
        SqliteConnection connection,
        HashSet<string> existingTables)
    {
        var clientsByKey = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
        if (existingTables.Contains("Clients"))
        {
            var columns = await GetColumnsAsync(connection, "Clients");
            var idColumn = columns.Contains("client_id") ? "client_id" : columns.Contains("id") ? "id" : "";
            if (!string.IsNullOrWhiteSpace(idColumn))
            {
                using var command = connection.CreateCommand();
                command.CommandText = "SELECT * FROM Clients;";
                using var reader = await command.ExecuteReaderAsync();
                while (await reader.ReadAsync())
                {
                    var clientId = ReadLong(reader, idColumn);
                    if (clientId is not > 0)
                    {
                        continue;
                    }

                    var client = new ClientLookup(
                        clientId.Value,
                        ReadString(reader, "client_name"),
                        ReadString(reader, "client_number"),
                        true);
                    foreach (var key in BuildClientKeys(client.ClientId, client))
                    {
                        clientsByKey.TryAdd(key, client.ClientId);
                    }
                }
            }
        }

        var clientIds = new HashSet<long>();
        foreach (var tableName in new[] { "proposal_files", "proposal_summary" })
        {
            if (!existingTables.Contains(tableName))
            {
                continue;
            }

            var columns = await GetColumnsAsync(connection, tableName);
            if (!columns.Contains("client_id") && !columns.Contains("client_name"))
            {
                continue;
            }

            var clientIdExpression = columns.Contains("client_id") ? "CAST(client_id AS TEXT)" : "''";
            var clientNameExpression = columns.Contains("client_name") ? "CAST(client_name AS TEXT)" : "''";
            using var command = connection.CreateCommand();
            command.CommandText = $"SELECT DISTINCT {clientIdExpression} AS stored_client_id, {clientNameExpression} AS stored_client_name FROM [{tableName}];";
            using var reader = await command.ExecuteReaderAsync();
            while (await reader.ReadAsync())
            {
                var storedClientId = ReadString(reader, "stored_client_id").Trim();
                var storedClientName = ReadString(reader, "stored_client_name").Trim();
                var matchedClientId = new[] { storedClientId, storedClientName }
                    .Where(value => !string.IsNullOrWhiteSpace(value))
                    .Select(value => value.ToLowerInvariant())
                    .Select(key => clientsByKey.TryGetValue(key, out var clientId) ? clientId : 0)
                    .FirstOrDefault(clientId => clientId > 0);

                if (matchedClientId > 0)
                {
                    clientIds.Add(matchedClientId);
                }
                else if (long.TryParse(storedClientId, NumberStyles.Integer, CultureInfo.InvariantCulture, out var numericClientId) &&
                         numericClientId > 0)
                {
                    clientIds.Add(numericClientId);
                }
            }
        }

        return clientIds.OrderBy(value => value).ToList();
    }

    private static IReadOnlyList<string> BuildClientKeys(long clientId, ClientLookup client)
    {
        return new[]
            {
                clientId.ToString(CultureInfo.InvariantCulture),
                client.ClientId.ToString(CultureInfo.InvariantCulture),
                client.ClientName,
                client.ClientNumber
            }
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Select(value => value.Trim().ToLowerInvariant())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static async Task<HashSet<long>> ReadProposalIdsAsync(
        SqliteConnection connection,
        HashSet<string> existingTables,
        IReadOnlyList<string> clientKeys,
        string clientName)
    {
        var proposalIds = new HashSet<long>();
        foreach (var tableName in new[] { "proposal_files", "proposal_summary" })
        {
            if (!existingTables.Contains(tableName))
            {
                continue;
            }

            var columns = await GetColumnsAsync(connection, tableName);
            var idExpression = columns.Contains("proposal_id") && columns.Contains("id")
                ? "COALESCE(proposal_id, id)"
                : columns.Contains("proposal_id")
                    ? "proposal_id"
                    : columns.Contains("id")
                        ? "id"
                        : "";
            if (string.IsNullOrWhiteSpace(idExpression))
            {
                continue;
            }

            using var command = connection.CreateCommand();
            var conditions = new List<string>();
            if (columns.Contains("client_id") && clientKeys.Count > 0)
            {
                conditions.Add($"LOWER(CAST(client_id AS TEXT)) IN ({AddTextParameters(command, clientKeys, tableName)})");
            }

            if (columns.Contains("client_name") && !string.IsNullOrWhiteSpace(clientName))
            {
                conditions.Add("LOWER(TRIM(client_name)) = LOWER(TRIM($client_name))");
                AddParameter(command, "$client_name", clientName);
            }

            if (conditions.Count == 0)
            {
                continue;
            }

            command.CommandText = $"SELECT DISTINCT {idExpression} FROM [{tableName}] WHERE {string.Join(" OR ", conditions)};";
            using var reader = await command.ExecuteReaderAsync();
            while (await reader.ReadAsync())
            {
                var proposalId = ToNullableLong(reader.GetValue(0));
                if (proposalId is > 0)
                {
                    proposalIds.Add(proposalId.Value);
                }
            }
        }

        return proposalIds;
    }

    private static async Task<ProposalCanonicalQuote> ReadProposalAsync(
        SqliteConnection connection,
        HashSet<string> existingTables,
        long proposalId,
        ClientLookup client)
    {
        var fileRow = existingTables.Contains("proposal_files")
            ? await ReadSingleRowAsync(connection, "proposal_files", proposalId)
            : new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        var summaryRow = existingTables.Contains("proposal_summary")
            ? await ReadSingleRowAsync(connection, "proposal_summary", proposalId)
            : new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);

        var carrierId = ReadLong(fileRow, "carrier_id") ?? ReadLong(summaryRow, "carrier_id");
        var carrierRow = carrierId.HasValue && existingTables.Contains("carriers")
            ? await ReadSingleRowByIdAsync(connection, "carriers", carrierId.Value)
            : new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        var summaryPayload = ReadString(summaryRow, "payload_json");

        var quote = new ProposalCanonicalQuote
        {
            ProposalId = proposalId,
            StoredClientId = FirstNonBlank(ReadString(fileRow, "client_id"), ReadString(summaryRow, "client_id")),
            CarrierId = carrierId,
            ClientName = FirstNonBlank(ReadString(summaryRow, "client_name"), client.ClientName),
            CarrierKey = ReadString(carrierRow, "carrier_key"),
            CarrierName = FirstNonBlank(ReadString(summaryRow, "carrier_name"), ReadString(carrierRow, "carrier_name"), ReadString(fileRow, "detected_carrier"), "Unknown Carrier"),
            DetectedCarrier = FirstNonBlank(ReadString(summaryRow, "detected_carrier"), ReadString(fileRow, "detected_carrier")),
            ProposalDate = ReadDateOnly(summaryRow, "proposal_date"),
            ProposalDateSourceText = ReadString(summaryRow, "proposal_date"),
            EffectiveDate = ReadDateOnly(summaryRow, "effective_date"),
            EffectiveDateSourceText = ReadString(summaryRow, "effective_date"),
            SourceFilePath = ReadString(fileRow, "file_path"),
            SourceFileName = FirstNonBlank(ReadString(fileRow, "source_file_name"), ReadString(summaryRow, "source_file_name")),
            SourceFileHash = FirstNonBlank(ReadString(fileRow, "source_file_hash"), ReadString(summaryRow, "source_file_hash")),
            DetectionConfidence = FirstNonBlank(ReadString(summaryRow, "detection_confidence"), ReadString(fileRow, "detection_confidence")),
            ExtractionConfidence = ParseConfidence(FirstNonBlank(ReadString(summaryRow, "extraction_confidence"), ReadString(fileRow, "extraction_confidence"))),
            ExtractionMethod = FirstNonBlank(ReadPayloadString(summaryPayload, "extraction_method"), "Proposal Engine"),
            Status = FirstNonBlank(ReadString(summaryRow, "status"), ReadString(fileRow, "status")),
            ReprocessCount = ReadInt(fileRow, "reprocess_count") ?? 0,
            CreatedAt = ReadDateTime(fileRow, "created_at") ?? ReadDateTime(summaryRow, "created_at"),
            UpdatedAt = ReadDateTime(fileRow, "updated_at") ?? ReadDateTime(summaryRow, "updated_at"),
            Evidence = BuildEvidence(fileRow.Count > 0 ? fileRow : summaryRow, fileRow.Count > 0 ? "proposal_files" : "proposal_summary")
        };

        foreach (var sourceFile in ReadPayloadStrings(summaryPayload, "source_files"))
        {
            if (!quote.SourceFiles.Contains(sourceFile, StringComparer.OrdinalIgnoreCase))
            {
                quote.SourceFiles.Add(sourceFile);
            }
        }

        if (quote.SourceFiles.Count == 0)
        {
            var sourceFile = FirstNonBlank(quote.SourceFilePath, quote.SourceFileName);
            if (!string.IsNullOrWhiteSpace(sourceFile))
            {
                quote.SourceFiles.Add(sourceFile);
            }
        }

        return quote;
    }

    private static async Task<Dictionary<string, object?>> ReadSingleRowAsync(
        SqliteConnection connection,
        string tableName,
        long proposalId)
    {
        var columns = await GetColumnsAsync(connection, tableName);
        string condition;
        if (columns.Contains("proposal_id"))
        {
            condition = "proposal_id = $proposal_id";
        }
        else if (columns.Contains("id"))
        {
            condition = "id = $proposal_id";
        }
        else
        {
            return new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        }

        using var command = connection.CreateCommand();
        var orderColumn = columns.Contains("id") ? "id" : "proposal_id";
        command.CommandText = $"SELECT * FROM [{tableName}] WHERE {condition} ORDER BY [{orderColumn}] LIMIT 1;";
        AddParameter(command, "$proposal_id", proposalId);
        using var reader = await command.ExecuteReaderAsync();
        return await reader.ReadAsync() ? ReadRow(reader) : new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
    }

    private static async Task<Dictionary<string, object?>> ReadSingleRowByIdAsync(
        SqliteConnection connection,
        string tableName,
        long id)
    {
        using var command = connection.CreateCommand();
        command.CommandText = $"SELECT * FROM [{tableName}] WHERE id = $id LIMIT 1;";
        AddParameter(command, "$id", id);
        using var reader = await command.ExecuteReaderAsync();
        return await reader.ReadAsync() ? ReadRow(reader) : new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
    }

    private static Dictionary<string, object?> ReadRow(SqliteDataReader reader)
    {
        var row = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        for (var index = 0; index < reader.FieldCount; index++)
        {
            row[reader.GetName(index)] = reader.IsDBNull(index) ? null : reader.GetValue(index);
        }

        return row;
    }

    private static async Task LoadDetailRowsAsync<T>(
        SqliteConnection connection,
        HashSet<string> existingTables,
        string tableName,
        IReadOnlyDictionary<long, ProposalCanonicalQuote> proposalsById,
        Func<SqliteDataReader, T> read,
        Action<ProposalCanonicalQuote, T> add)
    {
        if (!existingTables.Contains(tableName) || proposalsById.Count == 0)
        {
            return;
        }

        var columns = await GetColumnsAsync(connection, tableName);
        if (!columns.Contains("proposal_id"))
        {
            return;
        }

        using var command = connection.CreateCommand();
        var parameters = AddLongParameters(command, proposalsById.Keys, tableName);
        command.CommandText = $"SELECT * FROM [{tableName}] WHERE proposal_id IN ({parameters}) ORDER BY proposal_id, id;";
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            var proposalId = ReadLong(reader, "proposal_id");
            if (proposalId.HasValue && proposalsById.TryGetValue(proposalId.Value, out var proposal))
            {
                add(proposal, read(reader));
            }
        }
    }

    private static ProposalCanonicalBenefitDefinition ReadBenefit(SqliteDataReader reader)
    {
        return new ProposalCanonicalBenefitDefinition
        {
            BenefitKey = ReadString(reader, "benefit_key"),
            BenefitName = ReadString(reader, "benefit_name"),
            NormalizedBenefitName = ReadString(reader, "normalized_benefit_name"),
            DisplayName = ReadString(reader, "display_name"),
            BenefitCategory = ReadString(reader, "benefit_category"),
            SortOrder = ReadInt(reader, "sort_order"),
            IsStandard = ReadBool(reader, "is_standard"),
            IsKnown = ReadBool(reader, "is_known"),
            Evidence = ReadEvidence(reader, "benefits")
        };
    }

    private static ProposalCanonicalBenefitLine ReadBenefitLine(SqliteDataReader reader)
    {
        return new ProposalCanonicalBenefitLine
        {
            BenefitLine = ReadString(reader, "benefit_line"),
            ClassName = ReadString(reader, "class_name"),
            Tier = ReadString(reader, "tier"),
            Lives = ReadDecimal(reader, "lives"),
            Volume = ReadDecimal(reader, "volume"),
            Rate = ReadDecimal(reader, "rate"),
            MonthlyPremium = ReadDecimal(reader, "monthly_premium"),
            AnnualPremium = ReadDecimal(reader, "annual_premium"),
            Evidence = ReadEvidence(reader, "proposal_benefit_lines")
        };
    }

    private static ProposalCanonicalRate ReadRate(SqliteDataReader reader)
    {
        return new ProposalCanonicalRate
        {
            BenefitLine = ReadString(reader, "benefit_line"),
            ClassName = ReadString(reader, "class_name"),
            Tier = ReadString(reader, "tier"),
            AgeBand = ReadString(reader, "age_band"),
            Gender = ReadString(reader, "gender"),
            Province = ReadString(reader, "province"),
            Lives = ReadDecimal(reader, "lives"),
            Volume = ReadDecimal(reader, "volume"),
            Rate = ReadDecimal(reader, "rate"),
            RateBasis = ReadString(reader, "rate_basis"),
            MonthlyPremium = ReadDecimal(reader, "monthly_premium"),
            AnnualPremium = ReadDecimal(reader, "annual_premium"),
            Evidence = ReadEvidence(reader, "proposal_rates")
        };
    }

    private static ProposalCanonicalPremium ReadPremium(SqliteDataReader reader)
    {
        return new ProposalCanonicalPremium
        {
            BenefitLine = ReadString(reader, "benefit_line"),
            ClassName = ReadString(reader, "class_name"),
            Tier = ReadString(reader, "tier"),
            MonthlyPremium = ReadDecimal(reader, "monthly_premium"),
            AnnualPremium = ReadDecimal(reader, "annual_premium"),
            EmployerCost = ReadDecimal(reader, "employer_cost"),
            EmployeeCost = ReadDecimal(reader, "employee_cost"),
            TotalCost = ReadDecimal(reader, "total_cost"),
            Evidence = ReadEvidence(reader, "proposal_premiums")
        };
    }

    private static ProposalCanonicalPlanDesignFact ReadPlanDesign(SqliteDataReader reader, string tableName)
    {
        var payload = ReadString(reader, "payload_json");
        var fact = new ProposalCanonicalPlanDesignFact
        {
            BenefitLine = ReadString(reader, "benefit_line"),
            ProvisionName = ReadString(reader, "provision_name"),
            ProposalValue = ReadString(reader, "provision_value"),
            BookletValue = ReadString(reader, "booklet_value"),
            Variance = ReadString(reader, "variance"),
            ValidationStatus = ReadString(reader, "validation_status"),
            BookletCarrier = ReadPayloadString(payload, "booklet_carrier"),
            ComparisonMethod = ReadPayloadString(payload, "comparison_method"),
            Evidence = ReadEvidence(reader, tableName)
        };

        fact.BookletClassNames.AddRange(ReadPayloadStrings(payload, "booklet_class_names"));
        var bookletSourceFile = ReadPayloadString(payload, "booklet_source_file");
        var bookletSourcePage = ReadPayloadInt(payload, "booklet_source_page");
        var bookletSourceSection = ReadPayloadString(payload, "booklet_source_section");
        var bookletConfidence = ParseConfidence(ReadPayloadString(payload, "booklet_confidence_score"));
        if (!string.IsNullOrWhiteSpace(bookletSourceFile) || bookletSourcePage.HasValue || !string.IsNullOrWhiteSpace(bookletSourceSection))
        {
            fact.BookletEvidence = new ProposalCanonicalEvidence
            {
                SourceTable = "booklet_plan_design",
                ProposalId = ReadLong(reader, "proposal_id"),
                SourceFileName = bookletSourceFile,
                SourcePage = bookletSourcePage,
                SourceSheet = bookletSourceSection,
                ConfidenceScore = bookletConfidence,
                Status = ReadString(reader, "status"),
                SourceType = ResolveSourceType(bookletSourceFile),
                SourceLocation = bookletSourcePage.HasValue
                    ? $"Page {bookletSourcePage.Value.ToString(CultureInfo.InvariantCulture)}"
                    : bookletSourceSection,
                ExtractionMethod = FirstNonBlank(fact.ComparisonMethod, "Structured Booklet Extraction"),
                RawPayloadJson = payload
            };
        }

        return fact;
    }

    private static ProposalCanonicalDemographic ReadDemographic(SqliteDataReader reader)
    {
        return new ProposalCanonicalDemographic
        {
            BenefitLine = ReadString(reader, "benefit_line"),
            ClassName = ReadString(reader, "class_name"),
            Tier = ReadString(reader, "tier"),
            Lives = ReadDecimal(reader, "lives"),
            Volume = ReadDecimal(reader, "volume"),
            Evidence = ReadEvidence(reader, "proposal_demographics")
        };
    }

    private static ProposalCanonicalUnderwritingAssumption ReadUnderwritingAssumption(SqliteDataReader reader)
    {
        return new ProposalCanonicalUnderwritingAssumption
        {
            Name = ReadString(reader, "assumption_name"),
            Value = ReadString(reader, "assumption_value"),
            Evidence = ReadEvidence(reader, "proposal_underwriting_assumptions")
        };
    }

    private static ProposalCanonicalPoolingRetention ReadPoolingRetention(SqliteDataReader reader)
    {
        return new ProposalCanonicalPoolingRetention
        {
            BenefitLine = ReadString(reader, "benefit_line"),
            PoolingLevel = ReadString(reader, "pooling_level"),
            PoolingCharge = ReadDecimal(reader, "pooling_charge"),
            Commission = ReadDecimal(reader, "commission"),
            Retention = ReadDecimal(reader, "retention"),
            TargetLossRatio = ReadDecimal(reader, "target_loss_ratio"),
            RateGuarantee = ReadString(reader, "rate_guarantee"),
            PooledBenefits = ReadString(reader, "pooled_benefits"),
            ExperienceRatedBenefits = ReadString(reader, "experience_rated_benefits"),
            Ibnr = ReadString(reader, "ibnr"),
            RateCaps = ReadString(reader, "rate_caps"),
            FirstRenewalCap = ReadString(reader, "first_renewal_cap"),
            SecondRenewalCap = ReadString(reader, "second_renewal_cap"),
            Evidence = ReadEvidence(reader, "proposal_pooling_retention")
        };
    }

    private static ProposalCanonicalException ReadException(SqliteDataReader reader)
    {
        var payload = ReadString(reader, "payload_json");
        return new ProposalCanonicalException
        {
            ExceptionType = ReadString(reader, "exception_type"),
            FieldName = ReadString(reader, "field_name"),
            Message = ReadString(reader, "message"),
            BenefitLine = ReadPayloadString(payload, "benefit_line"),
            ProvisionName = ReadPayloadString(payload, "provision_name"),
            Evidence = ReadEvidence(reader, "proposal_exceptions")
        };
    }

    private static ProposalCanonicalAuditRecord ReadAudit(SqliteDataReader reader)
    {
        return new ProposalCanonicalAuditRecord
        {
            Action = ReadString(reader, "action"),
            TableName = ReadString(reader, "table_name"),
            RowsLoaded = ReadInt(reader, "rows_loaded"),
            Message = ReadString(reader, "message"),
            Evidence = ReadEvidence(reader, "proposal_export_audit")
        };
    }

    private static ProposalCanonicalEvidence ReadEvidence(SqliteDataReader reader, string tableName)
    {
        var payload = ReadString(reader, "payload_json");
        return new ProposalCanonicalEvidence
        {
            SourceTable = tableName,
            SourceId = ReadLong(reader, "id"),
            ProposalId = ReadLong(reader, "proposal_id"),
            SourceFileName = ReadString(reader, "source_file_name"),
            SourceFileHash = ReadString(reader, "source_file_hash"),
            SourcePage = ReadInt(reader, "source_page_number"),
            SourceSheet = ReadString(reader, "source_sheet_name"),
            ConfidenceScore = ParseConfidence(ReadString(reader, "extraction_confidence")),
            Status = ReadString(reader, "status"),
            SourceType = ReadPayloadString(payload, "source_type"),
            SourceLocation = FirstNonBlank(ReadPayloadString(payload, "source_location"), ReadString(reader, "source_sheet_name")),
            ExtractionMethod = FirstNonBlank(ReadPayloadString(payload, "extraction_method"), "Proposal Engine"),
            RawPayloadJson = payload,
            CreatedAt = ReadDateTime(reader, "created_at"),
            UpdatedAt = ReadDateTime(reader, "updated_at")
        };
    }

    private static ProposalCanonicalEvidence BuildEvidence(IReadOnlyDictionary<string, object?> row, string tableName)
    {
        var payload = ReadString(row, "payload_json");
        return new ProposalCanonicalEvidence
        {
            SourceTable = tableName,
            SourceId = ReadLong(row, "id"),
            ProposalId = ReadLong(row, "proposal_id"),
            SourceFileName = ReadString(row, "source_file_name"),
            SourceFileHash = ReadString(row, "source_file_hash"),
            SourcePage = ReadInt(row, "source_page_number"),
            SourceSheet = ReadString(row, "source_sheet_name"),
            ConfidenceScore = ParseConfidence(ReadString(row, "extraction_confidence")),
            Status = ReadString(row, "status"),
            SourceType = ReadPayloadString(payload, "source_type"),
            SourceLocation = ReadPayloadString(payload, "source_location"),
            ExtractionMethod = FirstNonBlank(ReadPayloadString(payload, "extraction_method"), "Proposal Engine"),
            RawPayloadJson = payload,
            CreatedAt = ReadDateTime(row, "created_at"),
            UpdatedAt = ReadDateTime(row, "updated_at")
        };
    }

    private static async Task<ProposalCanonicalSourceUsage> BuildSourceUsageAsync(
        SqliteConnection connection,
        HashSet<string> existingTables,
        string tableName,
        ClientLookup client,
        IReadOnlyList<ProposalCanonicalQuote> proposals,
        HashSet<long> proposalIds)
    {
        var source = new ProposalCanonicalSourceUsage
        {
            TableName = tableName,
            Exists = existingTables.Contains(tableName)
        };
        if (!source.Exists)
        {
            source.Notes = "Table not present; treated as an optional source.";
            return source;
        }

        if (tableName.Equals("Clients", StringComparison.OrdinalIgnoreCase))
        {
            source.RowsFound = client.Found ? 1 : 0;
        }
        else if (tableName.Equals("carriers", StringComparison.OrdinalIgnoreCase))
        {
            source.RowsFound = proposals.Select(proposal => proposal.CarrierId).Where(value => value.HasValue).Distinct().Count();
        }
        else
        {
            source.RowsFound = await CountProposalRowsAsync(connection, tableName, proposalIds);
        }

        source.Used = source.RowsFound > 0;
        source.Notes = source.Used ? "Used in the read-only canonical profile." : "No rows found for the selected client.";
        return source;
    }

    private static async Task<int> CountProposalRowsAsync(
        SqliteConnection connection,
        string tableName,
        HashSet<long> proposalIds)
    {
        if (proposalIds.Count == 0)
        {
            return 0;
        }

        var columns = await GetColumnsAsync(connection, tableName);
        var idColumn = columns.Contains("proposal_id") ? "proposal_id" : columns.Contains("id") ? "id" : "";
        if (string.IsNullOrWhiteSpace(idColumn))
        {
            return 0;
        }

        using var command = connection.CreateCommand();
        var parameters = AddLongParameters(command, proposalIds, $"count_{tableName}");
        command.CommandText = $"SELECT COUNT(*) FROM [{tableName}] WHERE [{idColumn}] IN ({parameters});";
        return Convert.ToInt32(await command.ExecuteScalarAsync(), CultureInfo.InvariantCulture);
    }

    private static void AddProfileIssues(ProposalCanonicalProfile profile)
    {
        if (profile.Proposals.Count == 0)
        {
            AddIssue(
                profile,
                "proposals_missing",
                ProposalCanonicalIssueSeverities.Error,
                "No saved proposals were found for the selected client.");
            return;
        }

        foreach (var proposal in profile.Proposals)
        {
            if (!proposal.HasCoreFacts)
            {
                AddIssue(profile, "proposal_facts_missing", ProposalCanonicalIssueSeverities.Error, "The saved proposal has no usable normalized facts.", proposal, "proposal_files");
                continue;
            }

            if (string.IsNullOrWhiteSpace(proposal.CarrierName) || proposal.CarrierName.Equals("Unknown Carrier", StringComparison.OrdinalIgnoreCase))
            {
                AddIssue(profile, "proposal_carrier_missing", ProposalCanonicalIssueSeverities.Warning, "The proposal carrier could not be resolved.", proposal, "carriers");
            }

            if (!proposal.HasFinancialFacts)
            {
                AddIssue(profile, "proposal_financial_facts_missing", ProposalCanonicalIssueSeverities.Warning, "No saved proposal rates or premiums were found.", proposal, "proposal_rates / proposal_premiums");
            }

            if (!proposal.HasPlanDesignFacts)
            {
                AddIssue(profile, "proposal_plan_design_missing", ProposalCanonicalIssueSeverities.Warning, "No saved proposal plan design facts were found.", proposal, "proposal_plan_design");
            }

            if (!proposal.HasBookletValidation)
            {
                AddIssue(profile, "proposal_validation_missing", ProposalCanonicalIssueSeverities.Warning, "No proposal-to-booklet validation facts were found.", proposal, "proposal_booklet_validation");
            }

            if (!proposal.HasSourceEvidence)
            {
                AddIssue(profile, "proposal_source_evidence_missing", ProposalCanonicalIssueSeverities.Warning, "The proposal does not contain usable source-file evidence.", proposal, "proposal_files");
            }

            foreach (var validation in proposal.ValidationFacts.Where(row => row.IsMismatch))
            {
                AddIssue(
                    profile,
                    "proposal_booklet_mismatch",
                    ProposalCanonicalIssueSeverities.Warning,
                    $"Proposal value '{validation.ProposalValue}' does not match booklet value '{validation.BookletValue}'.",
                    proposal,
                    "proposal_booklet_validation",
                    validation.Evidence.SourceId);
            }

            if (proposal.Exceptions.Any(row => row.NeedsReview))
            {
                AddIssue(
                    profile,
                    "proposal_exceptions_need_review",
                    ProposalCanonicalIssueSeverities.Warning,
                    $"{proposal.Exceptions.Count(row => row.NeedsReview).ToString(CultureInfo.InvariantCulture)} saved proposal extraction or validation issue(s) need review.",
                    proposal,
                    "proposal_exceptions");
            }
        }
    }

    private static void AddIssue(
        ProposalCanonicalProfile profile,
        string code,
        string severity,
        string message,
        ProposalCanonicalQuote? proposal = null,
        string sourceTable = "",
        long? sourceId = null)
    {
        profile.Issues.Add(new ProposalCanonicalIssue
        {
            Code = code,
            Severity = severity,
            Message = message,
            ProposalId = proposal?.ProposalId,
            Carrier = proposal?.CarrierName ?? "",
            SourceTable = sourceTable,
            SourceId = sourceId
        });
    }

    private static string AddTextParameters(SqliteCommand command, IReadOnlyList<string> values, string prefix)
    {
        var parameterNames = new List<string>();
        for (var index = 0; index < values.Count; index++)
        {
            var name = $"${SanitizeParameterName(prefix)}_{index.ToString(CultureInfo.InvariantCulture)}";
            parameterNames.Add(name);
            AddParameter(command, name, values[index]);
        }

        return string.Join(", ", parameterNames);
    }

    private static string AddLongParameters(SqliteCommand command, IEnumerable<long> values, string prefix)
    {
        var parameterNames = new List<string>();
        var index = 0;
        foreach (var value in values.Distinct())
        {
            var name = $"${SanitizeParameterName(prefix)}_{index.ToString(CultureInfo.InvariantCulture)}";
            parameterNames.Add(name);
            AddParameter(command, name, value);
            index++;
        }

        return string.Join(", ", parameterNames);
    }

    private static string SanitizeParameterName(string value)
    {
        return new string(value.Select(character => char.IsLetterOrDigit(character) ? character : '_').ToArray());
    }

    private static void AddParameter(SqliteCommand command, string name, object? value)
    {
        command.Parameters.AddWithValue(name, value ?? DBNull.Value);
    }

    private static bool TryGetOrdinal(SqliteDataReader reader, string name, out int ordinal)
    {
        for (var index = 0; index < reader.FieldCount; index++)
        {
            if (reader.GetName(index).Equals(name, StringComparison.OrdinalIgnoreCase))
            {
                ordinal = index;
                return true;
            }
        }

        ordinal = -1;
        return false;
    }

    private static string ReadString(SqliteDataReader reader, string name)
    {
        return TryGetOrdinal(reader, name, out var ordinal) && !reader.IsDBNull(ordinal)
            ? Convert.ToString(reader.GetValue(ordinal), CultureInfo.InvariantCulture)?.Trim() ?? ""
            : "";
    }

    private static long? ReadLong(SqliteDataReader reader, string name)
    {
        return TryGetOrdinal(reader, name, out var ordinal) && !reader.IsDBNull(ordinal)
            ? ToNullableLong(reader.GetValue(ordinal))
            : null;
    }

    private static int? ReadInt(SqliteDataReader reader, string name)
    {
        var value = ReadLong(reader, name);
        return value.HasValue && value.Value is >= int.MinValue and <= int.MaxValue ? (int)value.Value : null;
    }

    private static decimal? ReadDecimal(SqliteDataReader reader, string name)
    {
        if (!TryGetOrdinal(reader, name, out var ordinal) || reader.IsDBNull(ordinal))
        {
            return null;
        }

        return ToNullableDecimal(reader.GetValue(ordinal));
    }

    private static bool ReadBool(SqliteDataReader reader, string name)
    {
        var value = ReadString(reader, name);
        return value.Equals("1", StringComparison.OrdinalIgnoreCase) ||
               value.Equals("true", StringComparison.OrdinalIgnoreCase) ||
               value.Equals("yes", StringComparison.OrdinalIgnoreCase);
    }

    private static DateTime? ReadDateTime(SqliteDataReader reader, string name)
    {
        return TryGetOrdinal(reader, name, out var ordinal) && !reader.IsDBNull(ordinal)
            ? ToNullableDateTime(reader.GetValue(ordinal))
            : null;
    }

    private static string ReadString(IReadOnlyDictionary<string, object?> row, string name)
    {
        return row.TryGetValue(name, out var value) && value is not null && value is not DBNull
            ? Convert.ToString(value, CultureInfo.InvariantCulture)?.Trim() ?? ""
            : "";
    }

    private static long? ReadLong(IReadOnlyDictionary<string, object?> row, string name)
    {
        return row.TryGetValue(name, out var value) ? ToNullableLong(value) : null;
    }

    private static int? ReadInt(IReadOnlyDictionary<string, object?> row, string name)
    {
        var value = ReadLong(row, name);
        return value.HasValue && value.Value is >= int.MinValue and <= int.MaxValue ? (int)value.Value : null;
    }

    private static DateTime? ReadDateTime(IReadOnlyDictionary<string, object?> row, string name)
    {
        return row.TryGetValue(name, out var value) ? ToNullableDateTime(value) : null;
    }

    private static DateOnly? ReadDateOnly(IReadOnlyDictionary<string, object?> row, string name)
    {
        var text = ReadString(row, name);
        if (DateOnly.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out var value) ||
            DateOnly.TryParse(text, CultureInfo.CurrentCulture, DateTimeStyles.AllowWhiteSpaces, out value))
        {
            return value;
        }

        return null;
    }

    private static long? ToNullableLong(object? value)
    {
        if (value is null || value is DBNull)
        {
            return null;
        }

        return long.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), NumberStyles.Any, CultureInfo.InvariantCulture, out var result)
            ? result
            : null;
    }

    private static decimal? ToNullableDecimal(object? value)
    {
        if (value is null || value is DBNull)
        {
            return null;
        }

        return decimal.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), NumberStyles.Any, CultureInfo.InvariantCulture, out var result)
            ? result
            : null;
    }

    private static DateTime? ToNullableDateTime(object? value)
    {
        if (value is null || value is DBNull)
        {
            return null;
        }

        var text = Convert.ToString(value, CultureInfo.InvariantCulture);
        if (DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AssumeUniversal, out var result) ||
            DateTime.TryParse(text, CultureInfo.CurrentCulture, DateTimeStyles.AllowWhiteSpaces, out result))
        {
            return result;
        }

        return null;
    }

    private static decimal? ParseConfidence(string value)
    {
        var text = value.Trim();
        if (decimal.TryParse(text.TrimEnd('%'), NumberStyles.Number, CultureInfo.InvariantCulture, out var numeric))
        {
            return text.EndsWith('%') || numeric > 1m ? numeric / 100m : numeric;
        }

        return text.ToUpperInvariant() switch
        {
            "HIGH" => 0.90m,
            "MEDIUM" => 0.70m,
            "LOW" => 0.40m,
            _ => null
        };
    }

    private static string ReadPayloadString(string payload, string propertyName)
    {
        if (string.IsNullOrWhiteSpace(payload))
        {
            return "";
        }

        try
        {
            using var document = JsonDocument.Parse(payload);
            if (TryGetJsonProperty(document.RootElement, propertyName, out var property) && property.ValueKind != JsonValueKind.Null)
            {
                return property.ValueKind == JsonValueKind.String ? property.GetString()?.Trim() ?? "" : property.ToString().Trim();
            }
        }
        catch (JsonException)
        {
        }

        return "";
    }

    private static IReadOnlyList<string> ReadPayloadStrings(string payload, string propertyName)
    {
        if (string.IsNullOrWhiteSpace(payload))
        {
            return [];
        }

        try
        {
            using var document = JsonDocument.Parse(payload);
            if (TryGetJsonProperty(document.RootElement, propertyName, out var property) && property.ValueKind == JsonValueKind.Array)
            {
                return property.EnumerateArray()
                    .Select(item => item.ValueKind == JsonValueKind.String ? item.GetString()?.Trim() ?? "" : item.ToString().Trim())
                    .Where(value => !string.IsNullOrWhiteSpace(value))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();
            }
        }
        catch (JsonException)
        {
        }

        return [];
    }

    private static int? ReadPayloadInt(string payload, string propertyName)
    {
        var value = ReadPayloadString(payload, propertyName);
        return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var result) ? result : null;
    }

    private static string ResolveSourceType(string fileName)
    {
        return Path.GetExtension(fileName).ToLowerInvariant() switch
        {
            ".pdf" => "PDF",
            ".xls" or ".xlsx" or ".xlsm" => "Excel",
            ".doc" or ".docx" => "Word",
            _ => ""
        };
    }

    private static bool TryGetJsonProperty(JsonElement element, string propertyName, out JsonElement property)
    {
        if (element.ValueKind == JsonValueKind.Object)
        {
            foreach (var candidate in element.EnumerateObject())
            {
                if (candidate.Name.Equals(propertyName, StringComparison.OrdinalIgnoreCase))
                {
                    property = candidate.Value;
                    return true;
                }
            }
        }

        property = default;
        return false;
    }

    private static string FirstNonBlank(params string?[] values)
    {
        return values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value))?.Trim() ?? "";
    }

    private sealed record ClientLookup(long ClientId, string ClientName, string ClientNumber, bool Found);
}

using System.Globalization;
using System.Text.RegularExpressions;
using Microsoft.Data.Sqlite;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed partial class PlanDesignIntelligenceService
{
    private static readonly string[] SourceTableNames =
    [
        "PlanDesign",
        "plan_design_features",
        "plan_design_factors",
        "proposal_plan_design",
        "proposal_booklet_validation",
        "Renewals",
        "Rates",
        "ClaimsExperience",
        "MonthlyClaims",
        "claims_distribution"
    ];

    private static readonly IReadOnlyDictionary<string, string> NormalizationMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        ["Healthcare"] = "EHC",
        ["Health Care"] = "EHC",
        ["Extended Healthcare"] = "EHC",
        ["Visioncare"] = "Vision",
        ["Vision Care"] = "Vision",
        ["Dentalcare"] = "Dental",
        ["Dental Care"] = "Dental",
        ["Masseur"] = "Massage Therapist",
        ["Employee Life Insurance"] = "Basic Life",
        ["Accidental Death and Dismemberment"] = "AD&D"
    };

    private readonly string _databasePath;

    public PlanDesignIntelligenceService(string? databasePath = null)
    {
        _databasePath = string.IsNullOrWhiteSpace(databasePath)
            ? AppPaths.BenefitsIntelligenceDatabasePath
            : databasePath;
    }

    public string DatabasePath => _databasePath;

    public async Task<IReadOnlyList<PlanDesignIntelligenceClientOption>> GetClientOptionsAsync()
    {
        if (!File.Exists(_databasePath))
        {
            return [];
        }

        await using var connection = await OpenReadOnlyConnectionAsync();
        if (!await TableExistsAsync(connection, "Clients"))
        {
            return [];
        }

        using var command = connection.CreateCommand();
        command.CommandText = @"
SELECT client_id, client_name, client_number, carrier
FROM Clients
ORDER BY client_name COLLATE NOCASE;";

        var clients = new List<PlanDesignIntelligenceClientOption>();
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            clients.Add(new PlanDesignIntelligenceClientOption
            {
                ClientId = ReadLong(reader, "client_id") ?? 0,
                ClientName = ReadString(reader, "client_name"),
                ClientNumber = ReadString(reader, "client_number"),
                Carrier = ReadString(reader, "carrier")
            });
        }

        return clients;
    }

    public async Task<PlanDesignIntelligenceResult> AnalyzeClientAsync(long clientId)
    {
        var profile = await BuildCanonicalProfileAsync(clientId);
        var result = new PlanDesignIntelligenceResult
        {
            ClientId = clientId,
            ClientName = profile.ClientName,
            ClientNumber = profile.ClientNumber,
            Carrier = profile.Carrier,
            AnalyzedAt = DateTime.Now
        };

        if (profile.Issues.Any(issue => issue.Code == "database_missing"))
        {
            result.Findings.Add(CreateFinding(
                clientId,
                "",
                "",
                PlanDesignFindingTypes.MissingData,
                "Benefits Intelligence database was not found.",
                PlanDesignRiskLevels.High,
                1.00m,
                "",
                "",
                "Open the app database or load client data before running plan design intelligence."));
            return result;
        }

        foreach (var source in profile.Sources)
        {
            result.SourceTables.Add(new PlanDesignSourceTableUsage
            {
                TableName = source.TableName,
                Exists = source.Exists,
                RowsFound = source.RowsFound,
                Used = source.Used,
                Notes = source.Notes
            });
        }

        var planRows = profile.CurrentPlanFacts.Where(row => row.SourceTable == "PlanDesign").ToList();
        var featureRows = profile.CurrentPlanFacts.Where(row => row.SourceTable == "plan_design_features").ToList();
        var proposalRows = profile.QuoteFacts.ToList();
        var validationRows = profile.ProposalValidationFacts.ToList();

        var allPlanRows = planRows
            .Concat(featureRows)
            .Concat(proposalRows)
            .Concat(validationRows)
            .ToList();

        AddSummaryRows(result, allPlanRows);
        AddGapFindings(result, planRows, featureRows, proposalRows, validationRows, profile.ContextBenefitLines);
        AddComparisonFindings(result, planRows, featureRows, proposalRows, validationRows);
        AddNormalizationSuggestions(result, allPlanRows);
        AddPricingImpactOpportunities(result, allPlanRows, profile.ContextBenefitLines, profile.FactorReferences);
        AddReadinessRows(result, planRows, featureRows, proposalRows, validationRows);

        return result;
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

    private static async Task<bool> TableExistsAsync(SqliteConnection connection, string tableName)
    {
        using var command = connection.CreateCommand();
        command.CommandText = "SELECT 1 FROM sqlite_master WHERE type = 'table' AND LOWER(name) = LOWER($table_name) LIMIT 1;";
        AddParameter(command, "$table_name", tableName);
        return await command.ExecuteScalarAsync() is not null;
    }

    private static async Task<HashSet<string>> GetColumnsAsync(SqliteConnection connection, string tableName)
    {
        using var command = connection.CreateCommand();
        command.CommandText = $"PRAGMA table_info({tableName});";

        var columns = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            columns.Add(ReadString(reader, "name"));
        }

        return columns;
    }

    private static async Task<ClientLookup> ReadClientAsync(SqliteConnection connection, HashSet<string> existingTables, long clientId)
    {
        if (!existingTables.Contains("Clients"))
        {
            return new ClientLookup(clientId, $"Client #{clientId}", "", "");
        }

        using var command = connection.CreateCommand();
        command.CommandText = @"
SELECT client_id, client_name, client_number, carrier
FROM Clients
WHERE client_id = $client_id
LIMIT 1;";
        AddParameter(command, "$client_id", clientId);

        using var reader = await command.ExecuteReaderAsync();
        if (await reader.ReadAsync())
        {
            return new ClientLookup(
                ReadLong(reader, "client_id") ?? clientId,
                FirstNonBlank(ReadString(reader, "client_name"), $"Client #{clientId}"),
                ReadString(reader, "client_number"),
                ReadString(reader, "carrier"));
        }

        return new ClientLookup(clientId, $"Client #{clientId}", "", "");
    }

    private static IReadOnlyList<string> BuildClientKeys(long clientId, ClientLookup client)
    {
        return new[]
            {
                clientId.ToString(CultureInfo.InvariantCulture),
                client.ClientId.ToString(CultureInfo.InvariantCulture),
                client.ClientNumber,
                client.ClientName
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
        foreach (var tableName in new[] { "proposal_summary", "proposal_files" })
        {
            if (!existingTables.Contains(tableName))
            {
                continue;
            }

            var columns = await GetColumnsAsync(connection, tableName);
            if (!columns.Contains("proposal_id"))
            {
                continue;
            }

            using var command = connection.CreateCommand();
            var conditions = new List<string>();
            if (columns.Contains("client_id") && clientKeys.Count > 0)
            {
                conditions.Add($"LOWER(CAST(client_id AS TEXT)) IN ({AddTextListParameters(command, clientKeys, "client_key")})");
            }

            if (columns.Contains("client_name") && !string.IsNullOrWhiteSpace(clientName))
            {
                conditions.Add("LOWER(client_name) = LOWER($client_name)");
                AddParameter(command, "$client_name", clientName.Trim());
            }

            if (conditions.Count == 0)
            {
                continue;
            }

            command.CommandText = $"SELECT DISTINCT proposal_id FROM {tableName} WHERE {string.Join(" OR ", conditions)};";
            using var reader = await command.ExecuteReaderAsync();
            while (await reader.ReadAsync())
            {
                var proposalId = ToNullableLong(reader.GetValue(0));
                if (proposalId.HasValue)
                {
                    proposalIds.Add(proposalId.Value);
                }
            }
        }

        return proposalIds;
    }

    private static async Task<PlanDesignCanonicalSourceUsage> BuildSourceUsageAsync(
        SqliteConnection connection,
        HashSet<string> existingTables,
        string tableName,
        long clientId,
        IReadOnlyList<string> clientKeys,
        HashSet<long> proposalIds)
    {
        var usage = new PlanDesignCanonicalSourceUsage
        {
            TableName = tableName,
            Exists = existingTables.Contains(tableName)
        };

        if (!usage.Exists)
        {
            usage.Notes = "Table not present.";
            return usage;
        }

        usage.RowsFound = await CountRowsForClientAsync(connection, tableName, clientId, clientKeys, proposalIds);
        usage.Used = usage.RowsFound > 0;
        usage.Notes = usage.Used ? "Used in read-only analysis." : "No rows found for selected client.";
        return usage;
    }

    private static async Task<int> CountRowsForClientAsync(
        SqliteConnection connection,
        string tableName,
        long clientId,
        IReadOnlyList<string> clientKeys,
        HashSet<long> proposalIds)
    {
        using var command = connection.CreateCommand();
        if (tableName.Equals("plan_design_factors", StringComparison.OrdinalIgnoreCase))
        {
            command.CommandText = "SELECT COUNT(*) FROM plan_design_factors;";
            return Convert.ToInt32(await command.ExecuteScalarAsync(), CultureInfo.InvariantCulture);
        }

        if (tableName is "proposal_plan_design" or "proposal_booklet_validation")
        {
            var columns = await GetColumnsAsync(connection, tableName);
            var where = BuildProposalWhereClause(command, columns, clientKeys, proposalIds);
            if (string.IsNullOrWhiteSpace(where))
            {
                return 0;
            }

            command.CommandText = $"SELECT COUNT(*) FROM {tableName} p WHERE {where};";
            return Convert.ToInt32(await command.ExecuteScalarAsync(), CultureInfo.InvariantCulture);
        }

        var clientColumn = tableName switch
        {
            "PlanDesign" => "client_id",
            "plan_design_features" => "client_id",
            "Renewals" => "client_id",
            "Rates" => "client_id",
            "ClaimsExperience" => "client_id",
            "MonthlyClaims" => "client_id",
            "claims_distribution" => "client_id",
            _ => ""
        };

        if (string.IsNullOrWhiteSpace(clientColumn))
        {
            return 0;
        }

        command.CommandText = $"SELECT COUNT(*) FROM {tableName} WHERE {clientColumn} = $client_id;";
        AddParameter(command, "$client_id", clientId);
        return Convert.ToInt32(await command.ExecuteScalarAsync(), CultureInfo.InvariantCulture);
    }

    private static async Task<List<PlanDesignCanonicalFact>> ReadPlanDesignRowsAsync(
        SqliteConnection connection,
        HashSet<string> existingTables,
        long clientId)
    {
        if (!existingTables.Contains("PlanDesign"))
        {
            return [];
        }

        using var command = connection.CreateCommand();
        command.CommandText = @"
SELECT
    p.plan_design_id AS source_id,
    p.client_id,
    p.class_name,
    p.benefit_category,
    p.benefit_name,
    p.coverage_description,
    p.reimbursement_percent,
    p.maximum_amount,
    p.frequency,
    p.waiting_period,
    p.termination_age,
    p.source_document_id,
    p.source_page,
    p.date_loaded,
    p.source_type,
    p.source_carrier,
    COALESCE(NULLIF(p.source_file_name, ''), NULLIF(p.booklet_source_file_name, ''), '') AS tracked_source_file_name,
    p.source_api_endpoint,
    p.import_batch_id,
    p.imported_at,
    p.imported_by,
    p.confidence_score,
    p.extraction_method,
    p.raw_source_reference,
    COALESCE(NULLIF(d.file_name, ''), '') AS source_document,
    COALESCE(NULLIF(d.document_type, ''), 'plan_design') AS source_document_type,
    COALESCE(NULLIF(d.carrier, ''), NULLIF(c.carrier, ''), '') AS carrier
FROM PlanDesign p
LEFT JOIN Documents d ON d.document_id = p.source_document_id
LEFT JOIN Clients c ON c.client_id = p.client_id
WHERE p.client_id = $client_id
ORDER BY p.benefit_category, p.benefit_name;";
        AddParameter(command, "$client_id", clientId);

        var rows = new List<PlanDesignCanonicalFact>();
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            var category = ReadString(reader, "benefit_category");
            var benefit = ReadString(reader, "benefit_name");
            var sourceDocumentId = ReadLong(reader, "source_document_id");
            var trackedSourceFileName = ReadString(reader, "tracked_source_file_name");
            var linkedSourceFileName = ReadString(reader, "source_document");
            var lineage = CanonicalSourceLineageValueGuardrails.ResolveSourceFileName(
                sourceDocumentId,
                trackedSourceFileName,
                linkedSourceFileName);
            rows.Add(new PlanDesignCanonicalFact
            {
                Scope = PlanDesignCanonicalScopes.CurrentPlan,
                ClientId = ReadLong(reader, "client_id") ?? clientId,
                SourceTable = "PlanDesign",
                SourceId = ReadLong(reader, "source_id")?.ToString(CultureInfo.InvariantCulture) ?? "",
                SourceDocumentId = sourceDocumentId,
                SourceDocument = linkedSourceFileName,
                SourceDocumentType = ReadString(reader, "source_document_type"),
                SourceType = ReadString(reader, "source_type"),
                SourceFileName = lineage.ResolvedSourceFileName,
                TrackedSourceFileName = trackedSourceFileName,
                SourceLineageStatus = lineage.Status,
                SourceLineageMessage = lineage.Message,
                SourceCarrier = FirstNonBlank(ReadString(reader, "source_carrier"), ReadString(reader, "carrier")),
                SourceApiEndpoint = ReadString(reader, "source_api_endpoint"),
                ImportBatchId = ReadLong(reader, "import_batch_id"),
                ImportedAt = ReadDateTime(reader, "imported_at"),
                ImportedBy = ReadString(reader, "imported_by"),
                ConfidenceScore = ReadDecimal(reader, "confidence_score"),
                ExtractionMethod = FirstNonBlank(ReadString(reader, "extraction_method"), "Saved PlanDesign"),
                RawSourceReference = ReadString(reader, "raw_source_reference"),
                Carrier = ReadString(reader, "carrier"),
                BenefitCategory = FirstNonBlank(category, NormalizeBenefitCategory(benefit)),
                BenefitName = FirstNonBlank(benefit, ReadString(reader, "coverage_description"), category),
                ClassName = ReadString(reader, "class_name"),
                FeatureName = FirstNonBlank(benefit, category),
                FeatureValue = ReadString(reader, "coverage_description"),
                ReimbursementPercent = ReadDecimal(reader, "reimbursement_percent"),
                Maximum = ReadString(reader, "maximum_amount"),
                Frequency = ReadString(reader, "frequency"),
                WaitingPeriod = ReadString(reader, "waiting_period"),
                TerminationAge = ReadString(reader, "termination_age"),
                SourceText = ReadString(reader, "coverage_description"),
                SourcePage = ReadInt(reader, "source_page"),
                DateLoaded = ReadDateTime(reader, "imported_at") ?? ReadDateTime(reader, "date_loaded")
            });
        }

        return rows;
    }

    private static async Task<List<PlanDesignCanonicalFact>> ReadFeatureRowsAsync(
        SqliteConnection connection,
        HashSet<string> existingTables,
        long clientId)
    {
        if (!existingTables.Contains("plan_design_features"))
        {
            return [];
        }

        using var command = connection.CreateCommand();
        command.CommandText = @"
SELECT
    f.plan_design_feature_id AS source_id,
    f.client_id,
    f.source_document_id,
    f.benefit,
    f.category,
    f.feature_name,
    f.feature_value,
    f.coinsurance,
    f.deductible,
    f.maximum,
    f.frequency,
    f.termination_age,
    f.class_name,
    f.source_page,
    f.source_text,
    f.confidence_score,
    f.date_loaded,
    f.source_type,
    f.source_carrier,
    f.source_file_name AS tracked_source_file_name,
    f.source_api_endpoint,
    f.import_batch_id,
    f.imported_at,
    f.imported_by,
    f.extraction_method,
    f.raw_source_reference,
    COALESCE(NULLIF(d.file_name, ''), '') AS source_document,
    COALESCE(NULLIF(d.document_type, ''), 'plan_design_features') AS source_document_type,
    COALESCE(NULLIF(d.carrier, ''), NULLIF(c.carrier, ''), '') AS carrier
FROM plan_design_features f
LEFT JOIN Documents d ON d.document_id = f.source_document_id
LEFT JOIN Clients c ON c.client_id = f.client_id
WHERE f.client_id = $client_id
ORDER BY f.category, f.benefit, f.feature_name;";
        AddParameter(command, "$client_id", clientId);

        var rows = new List<PlanDesignCanonicalFact>();
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            var benefit = ReadString(reader, "benefit");
            var category = ReadString(reader, "category");
            var sourceDocumentId = ReadLong(reader, "source_document_id");
            var trackedSourceFileName = ReadString(reader, "tracked_source_file_name");
            var linkedSourceFileName = ReadString(reader, "source_document");
            var lineage = CanonicalSourceLineageValueGuardrails.ResolveSourceFileName(
                sourceDocumentId,
                trackedSourceFileName,
                linkedSourceFileName);
            rows.Add(new PlanDesignCanonicalFact
            {
                Scope = PlanDesignCanonicalScopes.CurrentPlan,
                ClientId = ReadLong(reader, "client_id") ?? clientId,
                SourceTable = "plan_design_features",
                SourceId = ReadLong(reader, "source_id")?.ToString(CultureInfo.InvariantCulture) ?? "",
                SourceDocumentId = sourceDocumentId,
                SourceDocument = linkedSourceFileName,
                SourceDocumentType = ReadString(reader, "source_document_type"),
                SourceType = ReadString(reader, "source_type"),
                SourceFileName = lineage.ResolvedSourceFileName,
                TrackedSourceFileName = trackedSourceFileName,
                SourceLineageStatus = lineage.Status,
                SourceLineageMessage = lineage.Message,
                SourceCarrier = FirstNonBlank(ReadString(reader, "source_carrier"), ReadString(reader, "carrier")),
                SourceApiEndpoint = ReadString(reader, "source_api_endpoint"),
                ImportBatchId = ReadLong(reader, "import_batch_id"),
                ImportedAt = ReadDateTime(reader, "imported_at"),
                ImportedBy = ReadString(reader, "imported_by"),
                ExtractionMethod = FirstNonBlank(ReadString(reader, "extraction_method"), "Plan Design Feature Extraction"),
                RawSourceReference = ReadString(reader, "raw_source_reference"),
                Carrier = ReadString(reader, "carrier"),
                BenefitCategory = FirstNonBlank(category, NormalizeBenefitCategory(benefit)),
                BenefitName = FirstNonBlank(benefit, ReadString(reader, "feature_name"), category),
                ClassName = ReadString(reader, "class_name"),
                FeatureName = ReadString(reader, "feature_name"),
                FeatureValue = ReadString(reader, "feature_value"),
                Coinsurance = ReadString(reader, "coinsurance"),
                Deductible = ReadString(reader, "deductible"),
                Maximum = ReadString(reader, "maximum"),
                Frequency = ReadString(reader, "frequency"),
                TerminationAge = ReadString(reader, "termination_age"),
                SourceText = ReadString(reader, "source_text"),
                SourcePage = ReadInt(reader, "source_page"),
                ConfidenceScore = ReadDecimal(reader, "confidence_score"),
                DateLoaded = ReadDateTime(reader, "imported_at") ?? ReadDateTime(reader, "date_loaded")
            });
        }

        return rows;
    }

    private static async Task<List<PlanDesignCanonicalFact>> ReadProposalRowsAsync(
        SqliteConnection connection,
        HashSet<string> existingTables,
        string tableName,
        IReadOnlyList<string> clientKeys,
        HashSet<long> proposalIds)
    {
        if (!existingTables.Contains(tableName))
        {
            return [];
        }

        var columns = await GetColumnsAsync(connection, tableName);
        var hasCarriers = existingTables.Contains("carriers") && columns.Contains("carrier_id");
        using var command = connection.CreateCommand();
        var where = BuildProposalWhereClause(command, columns, clientKeys, proposalIds);
        if (string.IsNullOrWhiteSpace(where))
        {
            return [];
        }

        var carrierSelect = hasCarriers
            ? "COALESCE(NULLIF(c.carrier_name, ''), '') AS carrier"
            : "'' AS carrier";
        var carrierJoin = hasCarriers
            ? "LEFT JOIN carriers c ON c.id = p.carrier_id"
            : "";
        command.CommandText = $@"
SELECT
    p.id AS source_id,
    p.client_id,
    p.proposal_id,
    p.source_file_name,
    p.source_file_hash,
    p.source_page_number,
    p.source_sheet_name,
    p.extraction_confidence,
    p.status,
    p.created_at,
    p.updated_at,
    p.benefit_line,
    p.provision_name,
    p.provision_value,
    p.booklet_value,
    p.variance,
    p.validation_status,
    {carrierSelect}
FROM {tableName} p
{carrierJoin}
WHERE {where}
ORDER BY p.source_file_name, p.benefit_line, p.provision_name;";

        var rows = new List<PlanDesignCanonicalFact>();
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            var benefit = ReadString(reader, "benefit_line");
            var provisionName = ReadString(reader, "provision_name");
            rows.Add(new PlanDesignCanonicalFact
            {
                Scope = tableName == "proposal_booklet_validation"
                    ? PlanDesignCanonicalScopes.ProposalValidation
                    : PlanDesignCanonicalScopes.Quote,
                ClientId = ReadLong(reader, "client_id") ?? 0,
                ProposalId = ReadLong(reader, "proposal_id"),
                SourceTable = tableName,
                SourceId = ReadLong(reader, "source_id")?.ToString(CultureInfo.InvariantCulture) ?? "",
                SourceDocument = FirstNonBlank(ReadString(reader, "source_file_name"), tableName),
                SourceDocumentType = tableName == "proposal_booklet_validation" ? "proposal_booklet_validation" : DocumentTypes.Proposal,
                SourceFileName = ReadString(reader, "source_file_name"),
                SourceFileHash = ReadString(reader, "source_file_hash"),
                SourceSheet = ReadString(reader, "source_sheet_name"),
                SourceType = PlanDesignCanonicalMapper.ResolveSourceType(ReadString(reader, "source_file_name")),
                ExtractionMethod = "Proposal Engine",
                ConfidenceScore = PlanDesignCanonicalMapper.ParseConfidence(ReadString(reader, "extraction_confidence")),
                Status = ReadString(reader, "status"),
                Carrier = ReadString(reader, "carrier"),
                SourceCarrier = ReadString(reader, "carrier"),
                BenefitCategory = NormalizeBenefitCategory(benefit),
                BenefitName = FirstNonBlank(benefit, provisionName),
                FeatureName = provisionName,
                FeatureValue = ReadString(reader, "provision_value"),
                BookletValue = ReadString(reader, "booklet_value"),
                Variance = ReadString(reader, "variance"),
                ValidationStatus = ReadString(reader, "validation_status"),
                SourcePage = ReadInt(reader, "source_page_number"),
                DateLoaded = ReadDateTime(reader, "updated_at") ?? ReadDateTime(reader, "created_at")
            });
        }

        return rows;
    }

    private static string BuildProposalWhereClause(
        SqliteCommand command,
        HashSet<string> columns,
        IReadOnlyList<string> clientKeys,
        HashSet<long> proposalIds)
    {
        var conditions = new List<string>();
        if (columns.Contains("client_id") && clientKeys.Count > 0)
        {
            conditions.Add($"LOWER(CAST(p.client_id AS TEXT)) IN ({AddTextListParameters(command, clientKeys, "proposal_client")})");
        }

        if (columns.Contains("proposal_id") && proposalIds.Count > 0)
        {
            conditions.Add($"p.proposal_id IN ({AddLongListParameters(command, proposalIds, "proposal_id")})");
        }

        return string.Join(" OR ", conditions);
    }

    private static async Task<List<PlanDesignFactorRecord>> ReadPlanDesignFactorsAsync(
        SqliteConnection connection,
        HashSet<string> existingTables)
    {
        if (!existingTables.Contains("plan_design_factors"))
        {
            return [];
        }

        using var command = connection.CreateCommand();
        command.CommandText = @"
SELECT id, carrier_name, benefit_category, benefit_line, current_design, proposed_design,
       change_type, factor, factor_direction, effective_date, expiry_date, source_document,
       notes, confidence_level, created_at, updated_at
FROM plan_design_factors
WHERE factor > 0
ORDER BY benefit_category, benefit_line, change_type;";

        var rows = new List<PlanDesignFactorRecord>();
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            rows.Add(new PlanDesignFactorRecord
            {
                Id = ReadLong(reader, "id") ?? 0,
                CarrierName = ReadString(reader, "carrier_name"),
                BenefitCategory = ReadString(reader, "benefit_category"),
                BenefitLine = ReadString(reader, "benefit_line"),
                CurrentDesign = ReadString(reader, "current_design"),
                ProposedDesign = ReadString(reader, "proposed_design"),
                ChangeType = ReadString(reader, "change_type"),
                Factor = ReadDecimal(reader, "factor") ?? 1.0000m,
                FactorDirection = ReadString(reader, "factor_direction"),
                EffectiveDate = ReadDateOnly(reader, "effective_date"),
                ExpiryDate = ReadDateOnly(reader, "expiry_date"),
                SourceDocument = ReadString(reader, "source_document"),
                Notes = ReadString(reader, "notes"),
                ConfidenceLevel = ReadString(reader, "confidence_level"),
                CreatedAt = ReadDateTime(reader, "created_at") ?? DateTime.MinValue,
                UpdatedAt = ReadDateTime(reader, "updated_at") ?? DateTime.MinValue
            });
        }

        return rows;
    }

    private static async Task<List<string>> ReadContextBenefitLinesAsync(
        SqliteConnection connection,
        HashSet<string> existingTables,
        long clientId)
    {
        var benefits = new List<string>();
        benefits.AddRange(await ReadBenefitColumnAsync(connection, existingTables, "Renewals", "benefit_line", clientId));
        benefits.AddRange(await ReadBenefitColumnAsync(connection, existingTables, "Rates", "benefit_line", clientId));
        benefits.AddRange(await ReadBenefitColumnAsync(connection, existingTables, "ClaimsExperience", "benefit_line", clientId));
        benefits.AddRange(await ReadBenefitColumnAsync(connection, existingTables, "MonthlyClaims", "benefit_line", clientId));
        benefits.AddRange(await ReadBenefitColumnAsync(connection, existingTables, "claims_distribution", "benefit", clientId));
        return benefits
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static async Task<List<string>> ReadBenefitColumnAsync(
        SqliteConnection connection,
        HashSet<string> existingTables,
        string tableName,
        string columnName,
        long clientId)
    {
        if (!existingTables.Contains(tableName))
        {
            return [];
        }

        var columns = await GetColumnsAsync(connection, tableName);
        if (!columns.Contains(columnName) || !columns.Contains("client_id"))
        {
            return [];
        }

        using var command = connection.CreateCommand();
        command.CommandText = $"SELECT DISTINCT {columnName} FROM {tableName} WHERE client_id = $client_id AND TRIM(COALESCE({columnName}, '')) <> '';";
        AddParameter(command, "$client_id", clientId);

        var values = new List<string>();
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            values.Add(ToStringValue(reader.GetValue(0)));
        }

        return values;
    }

    private static void AddSummaryRows(PlanDesignIntelligenceResult result, IReadOnlyList<PlanDesignCanonicalFact> rows)
    {
        var summaryRows = rows
            .GroupBy(row => string.Join(
                '\u001f',
                row.SourceTable,
                FirstNonBlank(row.SourceDocument, row.SourceTable),
                FirstNonBlank(row.SourceDocumentType, row.SourceTable),
                row.Carrier,
                FirstNonBlank(row.BenefitCategory, NormalizeBenefitCategory(row.BenefitName)),
                FirstNonBlank(row.BenefitName, row.FeatureName, "Plan design")),
                StringComparer.OrdinalIgnoreCase)
            .Select(group =>
            {
                var first = group.First();
                return new PlanDesignSummaryRow
                {
                    SourceTable = first.SourceTable,
                    SourceDocument = FirstNonBlank(first.SourceDocument, first.SourceTable),
                    SourceDocumentType = FirstNonBlank(first.SourceDocumentType, first.SourceTable),
                    Carrier = first.Carrier,
                    BenefitCategory = FirstNonBlank(first.BenefitCategory, NormalizeBenefitCategory(first.BenefitName)),
                    BenefitName = FirstNonBlank(first.BenefitName, first.FeatureName, "Plan design"),
                    FeatureCount = group.Count(),
                    LastLoadedDate = group
                        .Select(row => row.DateLoaded)
                        .Where(value => value.HasValue)
                        .OrderByDescending(value => value)
                        .FirstOrDefault()
                };
            })
            .OrderBy(row => row.SourceTable)
            .ThenBy(row => row.BenefitCategory)
            .ThenBy(row => row.BenefitName)
            .ToList();

        result.SummaryRows.AddRange(summaryRows);
    }

    private static void AddGapFindings(
        PlanDesignIntelligenceResult result,
        IReadOnlyList<PlanDesignCanonicalFact> planRows,
        IReadOnlyList<PlanDesignCanonicalFact> featureRows,
        IReadOnlyList<PlanDesignCanonicalFact> proposalRows,
        IReadOnlyList<PlanDesignCanonicalFact> validationRows,
        IReadOnlyList<string> contextBenefits)
    {
        var currentRows = planRows.Concat(featureRows).ToList();
        var hasBookletData = currentRows.Count > 0 || validationRows.Any(row => HasText(row.BookletValue));
        var hasProposalPlanDesign = proposalRows.Count > 0;
        var hasGranularFeatures = featureRows.Count > 0;

        if (!hasBookletData)
        {
            result.Findings.Add(CreateFinding(
                result.ClientId,
                "",
                "",
                PlanDesignFindingTypes.MissingData,
                "No booklet data found for the selected client.",
                PlanDesignRiskLevels.High,
                0.95m,
                "PlanDesign / plan_design_features",
                "",
                "Load or confirm current booklet plan design data before comparing current benefits."));
        }

        if (!hasProposalPlanDesign)
        {
            result.Findings.Add(CreateFinding(
                result.ClientId,
                "",
                "",
                PlanDesignFindingTypes.MissingData,
                "No proposal plan design rows were found for the selected client.",
                PlanDesignRiskLevels.High,
                0.95m,
                "proposal_plan_design",
                "",
                "Run or load proposal plan design data before proposal comparisons."));
        }

        if (!hasGranularFeatures)
        {
            result.Findings.Add(CreateFinding(
                result.ClientId,
                "",
                "",
                PlanDesignFindingTypes.MissingData,
                "No granular plan_design_features rows were found for the selected client.",
                PlanDesignRiskLevels.Medium,
                0.90m,
                "plan_design_features",
                "",
                "Use booklet feature extraction when detailed feature-level comparisons are required."));
        }

        AddMissingBenefitLineFindings(result, currentRows, contextBenefits);
        AddMissingFieldFindings(result, currentRows);
    }

    private static void AddMissingBenefitLineFindings(
        PlanDesignIntelligenceResult result,
        IReadOnlyList<PlanDesignCanonicalFact> currentRows,
        IReadOnlyList<string> contextBenefits)
    {
        if (contextBenefits.Count == 0 || currentRows.Count == 0)
        {
            return;
        }

        var currentKeys = currentRows.SelectMany(BenefitAliases).ToHashSet(StringComparer.OrdinalIgnoreCase);
        foreach (var benefit in contextBenefits.OrderBy(value => value, StringComparer.OrdinalIgnoreCase))
        {
            var key = NormalizeBenefitKey(benefit);
            if (string.IsNullOrWhiteSpace(key) || currentKeys.Contains(key))
            {
                continue;
            }

            result.Findings.Add(CreateFinding(
                result.ClientId,
                benefit,
                NormalizeBenefitCategory(benefit),
                PlanDesignFindingTypes.MissingData,
                $"Plan design data is missing for benefit line '{benefit}' even though renewal, rate, or claims data exists.",
                PlanDesignRiskLevels.Medium,
                0.80m,
                "Renewals / Rates / ClaimsExperience",
                "",
                "Confirm whether the current plan design includes this benefit line."));
        }
    }

    private static void AddMissingFieldFindings(PlanDesignIntelligenceResult result, IReadOnlyList<PlanDesignCanonicalFact> currentRows)
    {
        foreach (var group in currentRows.GroupBy(row => FirstNonBlank(row.BenefitName, row.BenefitCategory), StringComparer.OrdinalIgnoreCase))
        {
            var rows = group.ToList();
            if (rows.Count == 0)
            {
                continue;
            }

            var first = rows[0];
            AddMissingFieldFinding(result, rows, first, "maximums", HasMaximum, "Maximums are missing or weak for this benefit.");
            AddMissingFieldFinding(result, rows, first, "coinsurance", HasCoinsurance, "Coinsurance or reimbursement data is missing or weak for this benefit.");
            AddMissingFieldFinding(result, rows, first, "deductibles", HasDeductible, "Deductible data is missing or weak for this benefit.");
            AddMissingFieldFinding(result, rows, first, "termination ages", HasTerminationAge, "Termination age data is missing or weak for this benefit.");
            AddMissingFieldFinding(result, rows, first, "waiting periods", HasWaitingPeriod, "Waiting period data is missing or weak for this benefit.");
        }
    }

    private static void AddMissingFieldFinding(
        PlanDesignIntelligenceResult result,
        IReadOnlyList<PlanDesignCanonicalFact> rows,
        PlanDesignCanonicalFact first,
        string fieldName,
        Func<IReadOnlyList<PlanDesignCanonicalFact>, bool> hasField,
        string message)
    {
        if (hasField(rows))
        {
            return;
        }

        result.Findings.Add(CreateFinding(
            result.ClientId,
            FirstNonBlank(first.BenefitName, first.FeatureName),
            first.BenefitCategory,
            PlanDesignFindingTypes.MissingData,
            message,
            fieldName is "maximums" or "coinsurance" ? PlanDesignRiskLevels.Medium : PlanDesignRiskLevels.Low,
            0.78m,
            first.SourceTable,
            first.SourceId,
            $"Review source booklet data and fill the {fieldName} only through the existing extraction or data-loading workflow."));
    }

    private static void AddComparisonFindings(
        PlanDesignIntelligenceResult result,
        IReadOnlyList<PlanDesignCanonicalFact> planRows,
        IReadOnlyList<PlanDesignCanonicalFact> featureRows,
        IReadOnlyList<PlanDesignCanonicalFact> proposalRows,
        IReadOnlyList<PlanDesignCanonicalFact> validationRows)
    {
        var currentRows = planRows.Concat(featureRows).ToList();
        if (currentRows.Count > 0 && proposalRows.Count > 0)
        {
            var currentKeys = currentRows.SelectMany(BenefitAliases).ToHashSet(StringComparer.OrdinalIgnoreCase);
            var proposalBenefits = proposalRows
                .Where(row => BenefitAliases(row).Any(key => currentKeys.Contains(key)))
                .Select(row => FirstNonBlank(row.BenefitName, row.BenefitCategory))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(value => value)
                .Take(10)
                .ToList();

            foreach (var benefit in proposalBenefits)
            {
                result.Findings.Add(CreateFinding(
                    result.ClientId,
                    benefit,
                    NormalizeBenefitCategory(benefit),
                    PlanDesignFindingTypes.ComparisonOpportunity,
                    $"Current and proposal plan design data are both available for '{benefit}'.",
                    PlanDesignRiskLevels.Low,
                    0.85m,
                    "PlanDesign / proposal_plan_design",
                    "",
                    "Review current vs proposal values for client-facing comparison notes."));
            }
        }

        foreach (var row in validationRows.Where(IsMismatch))
        {
            result.Findings.Add(CreateFinding(
                result.ClientId,
                row.BenefitName,
                row.BenefitCategory,
                PlanDesignFindingTypes.InconsistentData,
                $"Proposal booklet validation needs review for {FirstNonBlank(row.BenefitName, "a plan design item")}: {FirstNonBlank(row.ValidationStatus, row.Variance, "variance found")}.",
                IsFailedValidation(row) ? PlanDesignRiskLevels.High : PlanDesignRiskLevels.Medium,
                0.88m,
                row.SourceTable,
                row.SourceId,
                "Review the proposal value against the booklet value before relying on the comparison."));
        }
    }

    private static void AddNormalizationSuggestions(PlanDesignIntelligenceResult result, IReadOnlyList<PlanDesignCanonicalFact> rows)
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var row in rows)
        {
            var textValues = new[]
            {
                row.BenefitCategory,
                row.BenefitName,
                row.FeatureName,
                row.FeatureValue,
                row.BookletValue
            };

            foreach (var value in textValues.Where(HasText))
            {
                foreach (var suggestion in NormalizationMap)
                {
                    if (!ContainsPhrase(value, suggestion.Key))
                    {
                        continue;
                    }

                    var key = $"{suggestion.Key}|{suggestion.Value}|{row.SourceTable}|{row.SourceId}";
                    if (!seen.Add(key))
                    {
                        continue;
                    }

                    result.Findings.Add(CreateFinding(
                        result.ClientId,
                        FirstNonBlank(row.BenefitName, row.FeatureName),
                        row.BenefitCategory,
                        PlanDesignFindingTypes.NormalizationSuggestion,
                        $"{suggestion.Key} should normalize to {suggestion.Value}.",
                        PlanDesignRiskLevels.Low,
                        0.76m,
                        row.SourceTable,
                        row.SourceId,
                        "Keep the stored value unchanged in Phase 1; use this as a review suggestion only."));
                }
            }
        }
    }

    private static void AddPricingImpactOpportunities(
        PlanDesignIntelligenceResult result,
        IReadOnlyList<PlanDesignCanonicalFact> rows,
        IReadOnlyList<string> contextBenefits,
        IReadOnlyList<PlanDesignFactorRecord> factors)
    {
        if (factors.Count == 0)
        {
            return;
        }

        var today = DateOnly.FromDateTime(DateTime.Today);
        var clientKeys = rows.SelectMany(BenefitAliases)
            .Concat(contextBenefits.Select(NormalizeBenefitKey))
            .Where(HasText)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var factor in factors.Where(factor => factor.IsActiveOn(today)))
        {
            var factorKeys = new[]
            {
                NormalizeBenefitKey(factor.BenefitLine),
                NormalizeBenefitKey(factor.BenefitCategory),
                NormalizeBenefitKey(factor.ChangeType)
            }.Where(HasText).ToList();

            if (clientKeys.Count > 0 && !factorKeys.Any(key => clientKeys.Contains(key)))
            {
                continue;
            }

            var key = $"{factor.BenefitLine}|{factor.ChangeType}|{factor.CurrentDesign}|{factor.ProposedDesign}";
            if (!seen.Add(key))
            {
                continue;
            }

            result.Findings.Add(CreateFinding(
                result.ClientId,
                FirstNonBlank(factor.BenefitLine, factor.BenefitCategory),
                factor.BenefitCategory,
                PlanDesignFindingTypes.PricingImpactOpportunity,
                $"Factor library indicates a possible pricing impact for {FirstNonBlank(factor.BenefitLine, factor.ChangeType)}: {factor.CurrentDesign} to {factor.ProposedDesign} ({factor.FactorText}).",
                IsIncreaseFactor(factor) ? PlanDesignRiskLevels.Medium : PlanDesignRiskLevels.Low,
                factor.IsEstimatedPlaceholder ? 0.55m : 0.82m,
                "plan_design_factors",
                factor.Id.ToString(CultureInfo.InvariantCulture),
                "Review the factor as a pricing-impact flag only; do not calculate final pricing unless the existing factor workflow supports it."));

            if (seen.Count >= 25)
            {
                break;
            }
        }
    }

    private static void AddReadinessRows(
        PlanDesignIntelligenceResult result,
        IReadOnlyList<PlanDesignCanonicalFact> planRows,
        IReadOnlyList<PlanDesignCanonicalFact> featureRows,
        IReadOnlyList<PlanDesignCanonicalFact> proposalRows,
        IReadOnlyList<PlanDesignCanonicalFact> validationRows)
    {
        var hasCurrent = planRows.Count > 0 || featureRows.Count > 0 || validationRows.Any(row => HasText(row.BookletValue));
        var hasGranular = featureRows.Count > 0;
        var hasProposal = proposalRows.Count > 0;
        var hasValidation = validationRows.Count > 0;
        var hasValidationVariance = validationRows.Any(IsMismatch);
        var hasRenewalOrRateData = result.SourceTables.Any(row =>
            row.TableName is "Renewals" or "Rates" &&
            row.RowsFound > 0);

        result.ReadinessRows.Add(BuildReadiness(
            "Current booklet vs proposal",
            hasCurrent && hasProposal && hasGranular && !hasValidationVariance
                ? PlanDesignReadinessStatuses.Ready
                : hasCurrent && hasProposal && hasValidationVariance
                    ? PlanDesignReadinessStatuses.NeedsReview
                    : hasCurrent && hasProposal
                        ? PlanDesignReadinessStatuses.Partial
                        : PlanDesignReadinessStatuses.NotReady,
            hasCurrent && hasProposal
                ? "Current and proposal plan design data are available."
                : "Current booklet data and proposal plan design data are both required."));

        result.ReadinessRows.Add(BuildReadiness(
            "Current plan design vs renewal",
            hasCurrent && hasRenewalOrRateData && hasGranular
                ? PlanDesignReadinessStatuses.Ready
                : hasCurrent && hasRenewalOrRateData
                    ? PlanDesignReadinessStatuses.Partial
                    : PlanDesignReadinessStatuses.NotReady,
            hasCurrent && hasRenewalOrRateData
                ? "Current plan design and renewal/rate data are available."
                : "Current plan design plus renewal or rate data are required."));

        result.ReadinessRows.Add(BuildReadiness(
            "Current plan design vs quote",
            hasCurrent && hasProposal && hasGranular
                ? PlanDesignReadinessStatuses.Ready
                : hasCurrent && hasProposal
                    ? PlanDesignReadinessStatuses.Partial
                    : PlanDesignReadinessStatuses.NotReady,
            hasCurrent && hasProposal
                ? "Current plan design and quote/proposal plan design data are available."
                : "Current plan design and proposal plan design data are required."));

        result.ReadinessRows.Add(BuildReadiness(
            "Proposal vs booklet validation",
            hasValidation && hasValidationVariance
                ? PlanDesignReadinessStatuses.NeedsReview
                : hasValidation
                    ? PlanDesignReadinessStatuses.Ready
                    : hasCurrent && hasProposal
                        ? PlanDesignReadinessStatuses.Partial
                        : PlanDesignReadinessStatuses.NotReady,
            hasValidation
                ? "Proposal booklet validation rows are available."
                : "Validation rows are required for direct proposal-to-booklet checks."));
    }

    private static PlanDesignReadinessRow BuildReadiness(string comparison, string status, string reason)
    {
        return new PlanDesignReadinessRow
        {
            Comparison = comparison,
            Status = status,
            Reason = reason
        };
    }

    private static PlanDesignIntelligenceFinding CreateFinding(
        long clientId,
        string benefit,
        string category,
        string findingType,
        string message,
        string riskLevel,
        decimal confidence,
        string sourceTable,
        string sourceId,
        string recommendedAction)
    {
        return new PlanDesignIntelligenceFinding
        {
            ClientId = clientId,
            Benefit = benefit,
            Category = category,
            FindingType = findingType,
            FindingMessage = message,
            RiskLevel = riskLevel,
            Confidence = confidence,
            SourceTable = sourceTable,
            SourceId = sourceId,
            RecommendedAction = recommendedAction
        };
    }

    private static bool HasMaximum(IReadOnlyList<PlanDesignCanonicalFact> rows)
    {
        return rows.Any(row =>
            HasText(row.Maximum) ||
            ContainsAny(row.FeatureName, "maximum", "max") ||
            ContainsAny(row.FeatureValue, "maximum", "max"));
    }

    private static bool HasCoinsurance(IReadOnlyList<PlanDesignCanonicalFact> rows)
    {
        return rows.Any(row =>
            row.ReimbursementPercent.HasValue ||
            HasText(row.Coinsurance) ||
            ContainsAny(row.FeatureName, "coinsurance", "reimbursement") ||
            ContainsAny(row.FeatureValue, "coinsurance", "reimbursement") ||
            PercentPattern().IsMatch(row.FeatureValue ?? ""));
    }

    private static bool HasDeductible(IReadOnlyList<PlanDesignCanonicalFact> rows)
    {
        return rows.Any(row =>
            HasText(row.Deductible) ||
            ContainsAny(row.FeatureName, "deductible") ||
            ContainsAny(row.FeatureValue, "deductible"));
    }

    private static bool HasTerminationAge(IReadOnlyList<PlanDesignCanonicalFact> rows)
    {
        return rows.Any(row =>
            HasText(row.TerminationAge) ||
            ContainsAny(row.FeatureName, "termination", "terminates") ||
            ContainsAny(row.FeatureValue, "termination", "terminates"));
    }

    private static bool HasWaitingPeriod(IReadOnlyList<PlanDesignCanonicalFact> rows)
    {
        return rows.Any(row =>
            HasText(row.WaitingPeriod) ||
            ContainsAny(row.FeatureName, "waiting period", "elimination period") ||
            ContainsAny(row.FeatureValue, "waiting period", "elimination period"));
    }

    private static bool IsMismatch(PlanDesignCanonicalFact row)
    {
        if (IsFailedValidation(row))
        {
            return true;
        }

        if (HasText(row.Variance) &&
            !row.Variance.Equals("none", StringComparison.OrdinalIgnoreCase) &&
            !row.Variance.Equals("match", StringComparison.OrdinalIgnoreCase) &&
            !row.Variance.Equals("matched", StringComparison.OrdinalIgnoreCase) &&
            !row.Variance.Contains("no variance", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return HasText(row.BookletValue) &&
               HasText(row.FeatureValue) &&
               !NormalizeComparableText(row.BookletValue).Equals(NormalizeComparableText(row.FeatureValue), StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsFailedValidation(PlanDesignCanonicalFact row) =>
        PlanDesignCanonicalMapper.IsFailedValidation(row);

    private static bool IsIncreaseFactor(PlanDesignFactorRecord factor)
    {
        return factor.Factor > 1m ||
               factor.FactorDirection.Contains("increase", StringComparison.OrdinalIgnoreCase);
    }

    private static IEnumerable<string> BenefitAliases(PlanDesignCanonicalFact row)
    {
        foreach (var value in new[] { row.BenefitCategory, row.BenefitName, row.FeatureName })
        {
            var key = NormalizeBenefitKey(value);
            if (HasText(key))
            {
                yield return key;
            }
        }
    }

    private static string NormalizeBenefitCategory(string value) =>
        PlanDesignCanonicalMapper.NormalizeBenefitCategory(value);

    private static string NormalizeBenefitKey(string? value) =>
        PlanDesignCanonicalMapper.NormalizeBenefitKey(value);

    private static string NormalizeComparableText(string? value) =>
        PlanDesignCanonicalMapper.NormalizeComparableText(value);

    private static bool ContainsAny(string? value, params string[] needles)
    {
        return HasText(value) && needles.Any(needle => value!.Contains(needle, StringComparison.OrdinalIgnoreCase));
    }

    private static bool ContainsPhrase(string value, string phrase)
    {
        return value.Contains(phrase, StringComparison.OrdinalIgnoreCase);
    }

    private static bool HasText(string? value)
    {
        return !string.IsNullOrWhiteSpace(value);
    }

    private static string FirstNonBlank(params string?[] values)
    {
        return values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value))?.Trim() ?? "";
    }

    private static string AddTextListParameters(SqliteCommand command, IReadOnlyList<string> values, string prefix)
    {
        var parameterNames = new List<string>();
        for (var index = 0; index < values.Count; index++)
        {
            var name = $"${prefix}_{command.Parameters.Count}_{index}";
            parameterNames.Add(name);
            AddParameter(command, name, values[index].Trim().ToLowerInvariant());
        }

        return string.Join(", ", parameterNames);
    }

    private static string AddLongListParameters(SqliteCommand command, IEnumerable<long> values, string prefix)
    {
        var parameterNames = new List<string>();
        var index = 0;
        foreach (var value in values)
        {
            var name = $"${prefix}_{command.Parameters.Count}_{index}";
            parameterNames.Add(name);
            AddParameter(command, name, value);
            index++;
        }

        return string.Join(", ", parameterNames);
    }

    private static void AddParameter(SqliteCommand command, string name, object? value)
    {
        command.Parameters.AddWithValue(name, value ?? DBNull.Value);
    }

    private static string ReadString(SqliteDataReader reader, string name)
    {
        return ToStringValue(reader.GetValue(reader.GetOrdinal(name)));
    }

    private static long? ReadLong(SqliteDataReader reader, string name)
    {
        return ToNullableLong(reader.GetValue(reader.GetOrdinal(name)));
    }

    private static int? ReadInt(SqliteDataReader reader, string name)
    {
        var value = reader.GetValue(reader.GetOrdinal(name));
        if (value is null or DBNull)
        {
            return null;
        }

        return Convert.ToInt32(value, CultureInfo.InvariantCulture);
    }

    private static decimal? ReadDecimal(SqliteDataReader reader, string name)
    {
        var value = reader.GetValue(reader.GetOrdinal(name));
        if (value is null or DBNull)
        {
            return null;
        }

        if (value is decimal decimalValue)
        {
            return decimalValue;
        }

        if (decimal.TryParse(ToStringValue(value), NumberStyles.Any, CultureInfo.InvariantCulture, out var parsed))
        {
            return parsed;
        }

        return null;
    }

    private static DateOnly? ReadDateOnly(SqliteDataReader reader, string name)
    {
        var text = ReadString(reader, name);
        return DateOnly.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out var value)
            ? value
            : null;
    }

    private static DateTime? ReadDateTime(SqliteDataReader reader, string name)
    {
        var text = ReadString(reader, name);
        return DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out var value)
            ? value
            : null;
    }

    private static long? ToNullableLong(object? value)
    {
        if (value is null or DBNull)
        {
            return null;
        }

        if (value is long longValue)
        {
            return longValue;
        }

        return long.TryParse(ToStringValue(value), NumberStyles.Any, CultureInfo.InvariantCulture, out var parsed)
            ? parsed
            : null;
    }

    private static string ToStringValue(object? value)
    {
        return value is null or DBNull ? "" : Convert.ToString(value, CultureInfo.InvariantCulture)?.Trim() ?? "";
    }

    [GeneratedRegex("\\b\\d{1,3}(\\.\\d+)?%\\b", RegexOptions.Compiled)]
    private static partial Regex PercentPattern();

    private sealed record ClientLookup(long ClientId, string ClientName, string ClientNumber, string Carrier);

}

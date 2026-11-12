using System.Collections.Generic;
using System.Globalization;
using Microsoft.Data.Sqlite;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class PlanDesignCanonicalProfileServiceTests
{
    [Fact]
    public async Task BuildAsync_ProducesCompleteReadOnlyProfileFromExistingTables()
    {
        var databasePath = CreateDatabasePath();
        try
        {
            var clientId = await CreateClientAsync(databasePath, "Canonical Profile Ltd.");
            var documentId = await InsertDocumentAsync(databasePath, clientId, "current-booklet.pdf");
            await InsertCurrentPlanAsync(databasePath, clientId, documentId);
            await InsertCurrentFeatureAsync(databasePath, clientId, documentId);
            await CreateProposalTableAsync(databasePath, "proposal_plan_design");
            await CreateProposalTableAsync(databasePath, "proposal_booklet_validation");
            await InsertProposalFactAsync(databasePath, "proposal_plan_design", clientId, "80%", "80%", "match");
            await InsertProposalFactAsync(databasePath, "proposal_booklet_validation", clientId, "80%", "80%", "match");

            var before = await SnapshotTableCountsAsync(databasePath);
            var profile = await new PlanDesignCanonicalProfileService(databasePath).BuildAsync(clientId);
            var after = await SnapshotTableCountsAsync(databasePath);

            Assert.Equal(PlanDesignCanonicalProfileVersions.Current, profile.ProfileVersion);
            Assert.True(profile.IsReadOnly);
            Assert.Equal("Canonical Profile Ltd.", profile.ClientName);
            Assert.True(profile.HasCurrentPlan);
            Assert.True(profile.HasQuotePlan);
            Assert.True(profile.HasProposalValidation);
            Assert.True(profile.CanCompareCurrentVsQuote);
            Assert.True(profile.CanSupportRenewalIntelligence);
            Assert.True(profile.CanSupportProposalIntelligence);
            Assert.True(profile.CanSupportPlanDesignIntelligence);
            Assert.Equal(PlanDesignCanonicalReadinessStatuses.Ready, profile.ReadinessStatus);

            var currentCoinsurance = Assert.Single(
                profile.CurrentPlanFacts,
                fact => fact.SourceTable == "plan_design_features" && fact.FeatureName == "Coinsurance");
            var quoteCoinsurance = Assert.Single(profile.QuoteFacts);
            Assert.Equal(CanonicalBenefitTypeCodes.Ehc, currentCoinsurance.CanonicalBenefitCode);
            Assert.Equal(currentCoinsurance.ComparisonKey, quoteCoinsurance.ComparisonKey);
            Assert.Equal(documentId, currentCoinsurance.Evidence.SourceDocumentId);
            Assert.Equal("current-booklet.pdf", currentCoinsurance.Evidence.SourceDocument);
            Assert.Equal(12, currentCoinsurance.Evidence.SourcePage);
            Assert.Equal(0.95m, currentCoinsurance.Evidence.ConfidenceScore);
            Assert.Equal("proposal.pdf", quoteCoinsurance.Evidence.SourceFileName);
            Assert.Equal(5, quoteCoinsurance.Evidence.SourcePage);
            Assert.NotEmpty(profile.FactorReferences);
            Assert.Contains(profile.Sources, source => source.TableName == "PlanDesign" && source.Used);
            Assert.Contains(profile.Sources, source => source.TableName == "proposal_plan_design" && source.Used);
            AssertTableCountsEqual(before, after);
        }
        finally
        {
            DeleteDatabase(databasePath);
        }
    }

    [Fact]
    public async Task BuildAsync_CurrentOnlyProfileSupportsRenewalButNotQuoteComparison()
    {
        var databasePath = CreateDatabasePath();
        try
        {
            var clientId = await CreateClientAsync(databasePath, "Current Only Ltd.");
            var documentId = await InsertDocumentAsync(databasePath, clientId, "current-only.pdf");
            await InsertCurrentPlanAsync(databasePath, clientId, documentId);

            var profile = await new PlanDesignCanonicalProfileService(databasePath).BuildAsync(clientId);

            Assert.True(profile.HasCurrentPlan);
            Assert.False(profile.HasQuotePlan);
            Assert.False(profile.CanCompareCurrentVsQuote);
            Assert.True(profile.CanSupportRenewalIntelligence);
            Assert.False(profile.CanSupportProposalIntelligence);
            Assert.True(profile.CanSupportPlanDesignIntelligence);
            Assert.Equal(PlanDesignCanonicalReadinessStatuses.Partial, profile.ReadinessStatus);
            Assert.Contains(profile.Issues, issue => issue.Code == "quote_plan_missing");
            Assert.Contains(profile.Issues, issue => issue.Code == "proposal_validation_missing");
        }
        finally
        {
            DeleteDatabase(databasePath);
        }
    }

    [Fact]
    public async Task BuildAsync_QuoteOnlyProfileIsNotReadyWithoutCurrentPlan()
    {
        var databasePath = CreateDatabasePath();
        try
        {
            var clientId = await CreateClientAsync(databasePath, "Quote Only Ltd.");
            await CreateProposalTableAsync(databasePath, "proposal_plan_design");
            await InsertProposalFactAsync(databasePath, "proposal_plan_design", clientId, "80%", "", "valid");

            var profile = await new PlanDesignCanonicalProfileService(databasePath).BuildAsync(clientId);

            Assert.False(profile.HasCurrentPlan);
            Assert.True(profile.HasQuotePlan);
            Assert.False(profile.CanCompareCurrentVsQuote);
            Assert.False(profile.CanSupportRenewalIntelligence);
            Assert.Equal(PlanDesignCanonicalReadinessStatuses.NotReady, profile.ReadinessStatus);
            Assert.Contains(profile.Issues, issue => issue.Code == "current_plan_missing" && issue.Severity == PlanDesignCanonicalIssueSeverities.Error);
        }
        finally
        {
            DeleteDatabase(databasePath);
        }
    }

    [Fact]
    public async Task BuildAsync_FailedProposalValidationRequiresReview()
    {
        var databasePath = CreateDatabasePath();
        try
        {
            var clientId = await CreateClientAsync(databasePath, "Mismatch Ltd.");
            var documentId = await InsertDocumentAsync(databasePath, clientId, "mismatch-current.pdf");
            await InsertCurrentFeatureAsync(databasePath, clientId, documentId);
            await CreateProposalTableAsync(databasePath, "proposal_plan_design");
            await CreateProposalTableAsync(databasePath, "proposal_booklet_validation");
            await InsertProposalFactAsync(databasePath, "proposal_plan_design", clientId, "90%", "80%", "mismatch");
            await InsertProposalFactAsync(databasePath, "proposal_booklet_validation", clientId, "90%", "80%", "mismatch");

            var profile = await new PlanDesignCanonicalProfileService(databasePath).BuildAsync(clientId);

            Assert.True(profile.CanCompareCurrentVsQuote);
            Assert.Equal(PlanDesignCanonicalReadinessStatuses.NeedsReview, profile.ReadinessStatus);
            Assert.Contains(profile.Issues, issue =>
                issue.Code == "proposal_plan_design_mismatch" &&
                issue.SourceTable == "proposal_booklet_validation");
            Assert.All(profile.ProposalValidationFacts, fact => Assert.Equal(PlanDesignCanonicalAvailability.NeedsReview, fact.Availability));
        }
        finally
        {
            DeleteDatabase(databasePath);
        }
    }

    [Fact]
    public async Task BuildAsync_UsesLinkedDocumentAndRejectsUnsafePercentageWithoutWriting()
    {
        var databasePath = CreateDatabasePath();
        try
        {
            var clientId = await CreateClientAsync(databasePath, "Lineage Guardrail Ltd.");
            var documentId = await InsertDocumentAsync(databasePath, clientId, "canonical-booklet.pdf");
            await InsertCurrentPlanAsync(databasePath, clientId, documentId);

            await using (var connection = await OpenConnectionAsync(databasePath))
            {
                using var command = connection.CreateCommand();
                command.CommandText = @"
UPDATE PlanDesign
SET source_file_name = 'wrong-policy.pdf',
    coverage_description = 'Reimbursement: 8050%',
    reimbursement_percent = 8050
WHERE client_id = $client_id;";
                AddParameter(command, "$client_id", clientId);
                await command.ExecuteNonQueryAsync();
            }

            var before = await SnapshotTableCountsAsync(databasePath);
            var profile = await new PlanDesignCanonicalProfileService(databasePath).BuildAsync(clientId);
            var after = await SnapshotTableCountsAsync(databasePath);

            var fact = Assert.Single(profile.CurrentPlanFacts);
            Assert.Equal("canonical-booklet.pdf", fact.SourceFileName);
            Assert.Equal("wrong-policy.pdf", fact.TrackedSourceFileName);
            Assert.Equal(CanonicalSourceLineageStatuses.NeedsReview, fact.SourceLineageStatus);
            Assert.Equal(CanonicalValueGuardrailStatuses.Rejected, fact.ValueGuardrailStatus);
            Assert.Null(fact.ReimbursementPercent);
            Assert.Equal(PlanDesignCanonicalAvailability.NeedsReview, fact.Availability);
            Assert.Equal(CanonicalSourceLineageValueGuardrailVersions.Current, fact.Evidence.GuardrailVersion);
            Assert.Equal(PlanDesignCanonicalReadinessStatuses.NeedsReview, profile.ReadinessStatus);
            Assert.Contains(profile.Issues, issue =>
                issue.Code == CanonicalSourceLineageValueGuardrailIssueCodes.SourceLineageConflict);
            Assert.Contains(profile.Issues, issue =>
                issue.Code == CanonicalSourceLineageValueGuardrailIssueCodes.PercentageOutOfRange);
            AssertTableCountsEqual(before, after);
        }
        finally
        {
            DeleteDatabase(databasePath);
        }
    }

    private static string CreateDatabasePath()
    {
        var folder = Path.Combine(Path.GetTempPath(), "PlanDesignCanonicalProfileTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        var databasePath = Path.Combine(folder, "Benefits_Intelligence.db");
        var result = new DatabaseService(databasePath).Initialize();
        Assert.True(result.Succeeded, result.Message);
        return databasePath;
    }

    private static async Task<long> CreateClientAsync(string databasePath, string clientName)
    {
        await using var connection = await OpenConnectionAsync(databasePath);
        using var command = connection.CreateCommand();
        command.CommandText = @"
INSERT INTO Clients (client_name, client_number, carrier, status, date_created, date_updated)
VALUES ($client_name, $client_number, 'Manulife', 'Active', $now, $now);
SELECT last_insert_rowid();";
        AddParameter(command, "$client_name", clientName);
        AddParameter(command, "$client_number", Guid.NewGuid().ToString("N")[..8]);
        AddParameter(command, "$now", NowText());
        return Convert.ToInt64(await command.ExecuteScalarAsync(), CultureInfo.InvariantCulture);
    }

    private static async Task<long> InsertDocumentAsync(string databasePath, long clientId, string fileName)
    {
        await using var connection = await OpenConnectionAsync(databasePath);
        using var command = connection.CreateCommand();
        command.CommandText = @"
INSERT INTO Documents (
    client_id, document_type, module_name, file_name, file_type, file_path, file_hash,
    carrier, effective_date, upload_date, processed_date, processing_status
)
VALUES (
    $client_id, $document_type, 'Canonical Plan Design Profile', $file_name, 'PDF', $file_name, $file_hash,
    'Manulife', '2026-01-01', $now, $now, 'Processed'
);
SELECT last_insert_rowid();";
        AddParameter(command, "$client_id", clientId);
        AddParameter(command, "$document_type", DocumentTypes.BenefitBooklet);
        AddParameter(command, "$file_name", fileName);
        AddParameter(command, "$file_hash", Guid.NewGuid().ToString("N"));
        AddParameter(command, "$now", NowText());
        return Convert.ToInt64(await command.ExecuteScalarAsync(), CultureInfo.InvariantCulture);
    }

    private static async Task InsertCurrentPlanAsync(string databasePath, long clientId, long documentId)
    {
        await using var connection = await OpenConnectionAsync(databasePath);
        using var command = connection.CreateCommand();
        command.CommandText = @"
INSERT INTO PlanDesign (
    client_id, class_name, benefit_category, benefit_name, coverage_description,
    reimbursement_percent, maximum_amount, frequency, waiting_period, termination_age,
    source_document_id, source_page, date_loaded
)
VALUES (
    $client_id, 'All Classes', 'Extended Health Care', 'Extended Health Care', 'Prescription drugs',
    80, 'Unlimited', 'Per calendar year', 'None', 'Age 70',
    $document_id, 10, $now
);";
        AddParameter(command, "$client_id", clientId);
        AddParameter(command, "$document_id", documentId);
        AddParameter(command, "$now", NowText());
        await command.ExecuteNonQueryAsync();
    }

    private static async Task InsertCurrentFeatureAsync(string databasePath, long clientId, long documentId)
    {
        await using var connection = await OpenConnectionAsync(databasePath);
        using var command = connection.CreateCommand();
        command.CommandText = @"
INSERT INTO plan_design_features (
    client_id, source_document_id, benefit, category, feature_name, feature_value,
    coinsurance, deductible, maximum, frequency, termination_age, class_name,
    source_page, source_text, confidence_score, date_loaded
)
VALUES (
    $client_id, $document_id, 'Extended Health Care', 'Extended Health Care', 'Coinsurance', '80%',
    '80%', '$25', 'Unlimited', 'Per calendar year', 'Age 70', 'All Classes',
    12, 'Coinsurance is 80 percent.', 0.95, $now
);";
        AddParameter(command, "$client_id", clientId);
        AddParameter(command, "$document_id", documentId);
        AddParameter(command, "$now", NowText());
        await command.ExecuteNonQueryAsync();
    }

    private static async Task CreateProposalTableAsync(string databasePath, string tableName)
    {
        await using var connection = await OpenConnectionAsync(databasePath);
        using var command = connection.CreateCommand();
        command.CommandText = $@"
CREATE TABLE IF NOT EXISTS {tableName} (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id TEXT,
    carrier_id INTEGER,
    proposal_id INTEGER,
    source_file_name TEXT,
    source_file_hash TEXT,
    source_page_number INTEGER,
    source_sheet_name TEXT,
    extraction_confidence TEXT,
    status TEXT,
    created_at TEXT,
    updated_at TEXT,
    benefit_line TEXT,
    provision_name TEXT,
    provision_value TEXT,
    booklet_value TEXT,
    variance TEXT,
    validation_status TEXT,
    payload_json TEXT
);";
        await command.ExecuteNonQueryAsync();
    }

    private static async Task InsertProposalFactAsync(
        string databasePath,
        string tableName,
        long clientId,
        string provisionValue,
        string bookletValue,
        string validationStatus)
    {
        await using var connection = await OpenConnectionAsync(databasePath);
        using var command = connection.CreateCommand();
        command.CommandText = $@"
INSERT INTO {tableName} (
    client_id, proposal_id, source_file_name, source_file_hash, source_page_number,
    source_sheet_name, extraction_confidence, status, created_at, updated_at,
    benefit_line, provision_name, provision_value, booklet_value, variance,
    validation_status, payload_json
)
VALUES (
    $client_id, 101, 'proposal.pdf', $source_hash, 5,
    '', 'High', 'VALID', $now, $now,
    'Extended Health Care', 'Coinsurance', $provision_value, $booklet_value, $variance,
    $validation_status, '{{}}'
);";
        AddParameter(command, "$client_id", clientId.ToString(CultureInfo.InvariantCulture));
        AddParameter(command, "$source_hash", Guid.NewGuid().ToString("N"));
        AddParameter(command, "$now", NowText());
        AddParameter(command, "$provision_value", provisionValue);
        AddParameter(command, "$booklet_value", bookletValue);
        AddParameter(command, "$variance", validationStatus.Equals("mismatch", StringComparison.OrdinalIgnoreCase) ? "changed" : "none");
        AddParameter(command, "$validation_status", validationStatus);
        await command.ExecuteNonQueryAsync();
    }

    private static async Task<Dictionary<string, long>> SnapshotTableCountsAsync(string databasePath)
    {
        await using var connection = await OpenConnectionAsync(databasePath);
        using var command = connection.CreateCommand();
        command.CommandText = "SELECT name FROM sqlite_master WHERE type = 'table' ORDER BY name;";

        var tableNames = new List<string>();
        using (var reader = await command.ExecuteReaderAsync())
        {
            while (await reader.ReadAsync())
            {
                tableNames.Add(reader.GetString(0));
            }
        }

        var counts = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
        foreach (var tableName in tableNames)
        {
            using var countCommand = connection.CreateCommand();
            countCommand.CommandText = $"SELECT COUNT(*) FROM {tableName};";
            counts[tableName] = Convert.ToInt64(await countCommand.ExecuteScalarAsync(), CultureInfo.InvariantCulture);
        }

        return counts;
    }

    private static void AssertTableCountsEqual(Dictionary<string, long> expected, Dictionary<string, long> actual)
    {
        Assert.Equal(expected.Count, actual.Count);
        foreach (var expectedPair in expected)
        {
            Assert.True(actual.TryGetValue(expectedPair.Key, out var actualCount), $"Missing table count for {expectedPair.Key}.");
            Assert.Equal(expectedPair.Value, actualCount);
        }
    }

    private static async Task<SqliteConnection> OpenConnectionAsync(string databasePath)
    {
        var connection = new SqliteConnection($"Data Source={databasePath}");
        await connection.OpenAsync();
        return connection;
    }

    private static void AddParameter(SqliteCommand command, string name, object? value)
    {
        command.Parameters.AddWithValue(name, value ?? DBNull.Value);
    }

    private static string NowText()
    {
        return DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture);
    }

    private static void DeleteDatabase(string databasePath)
    {
        SqliteConnection.ClearAllPools();
        var folder = Path.GetDirectoryName(databasePath);
        if (!string.IsNullOrWhiteSpace(folder) && Directory.Exists(folder))
        {
            Directory.Delete(folder, true);
        }
    }
}

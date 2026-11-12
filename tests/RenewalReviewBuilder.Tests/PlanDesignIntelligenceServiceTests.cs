using System.Collections.Generic;
using System.Globalization;
using Microsoft.Data.Sqlite;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class PlanDesignIntelligenceServiceTests
{
    [Fact]
    public async Task AnalyzeClient_Does_Not_Modify_Database()
    {
        var databasePath = CreateDatabasePath();
        var clientId = await CreateClientAsync(databasePath, "Read Only Test Ltd.");
        var documentId = await InsertDocumentAsync(databasePath, clientId, DocumentTypes.BenefitBooklet, "read-only-booklet.pdf");
        await InsertPlanDesignAsync(databasePath, clientId, documentId, "Extended Health", "Drugs", reimbursement: 80m, maximum: "Unlimited");

        var before = await SnapshotTableCountsAsync(databasePath);
        var service = new PlanDesignIntelligenceService(databasePath);
        var result = await service.AnalyzeClientAsync(clientId);
        var after = await SnapshotTableCountsAsync(databasePath);

        Assert.Contains(result.SummaryRows, row => row.SourceTable == "PlanDesign");
        AssertTableCountsEqual(before, after);
    }

    [Fact]
    public async Task AnalyzeClient_Handles_Client_With_Only_PlanDesign_Rows()
    {
        var databasePath = CreateDatabasePath();
        var clientId = await CreateClientAsync(databasePath, "Plan Only Ltd.");
        var documentId = await InsertDocumentAsync(databasePath, clientId, DocumentTypes.BenefitBooklet, "plan-only-booklet.pdf");
        await InsertPlanDesignAsync(databasePath, clientId, documentId, "Extended Health", "Drugs", reimbursement: 80m, maximum: "Unlimited");

        var result = await new PlanDesignIntelligenceService(databasePath).AnalyzeClientAsync(clientId);

        Assert.Contains(result.SummaryRows, row => row.SourceTable == "PlanDesign" && row.BenefitName == "Drugs");
        Assert.DoesNotContain(result.Findings, finding => finding.FindingMessage.Contains("No booklet data", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(result.Findings, finding => finding.FindingMessage.Contains("No proposal plan design", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public async Task AnalyzeClient_Handles_Client_With_Only_ProposalPlanDesign_Rows()
    {
        var databasePath = CreateDatabasePath();
        var clientId = await CreateClientAsync(databasePath, "Proposal Only Ltd.");
        await CreateProposalPlanDesignTableAsync(databasePath, "proposal_plan_design");
        await InsertProposalPlanDesignAsync(databasePath, "proposal_plan_design", clientId, "Dental", "Basic Services", "80%", "", "match");

        var result = await new PlanDesignIntelligenceService(databasePath).AnalyzeClientAsync(clientId);

        Assert.Contains(result.SummaryRows, row => row.SourceTable == "proposal_plan_design" && row.BenefitName == "Dental");
        Assert.Contains(result.Findings, finding => finding.FindingMessage.Contains("No booklet data", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain(result.Findings, finding => finding.FindingMessage.Contains("No proposal plan design", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public async Task AnalyzeClient_Handles_Client_With_No_PlanDesign_Rows()
    {
        var databasePath = CreateDatabasePath();
        var clientId = await CreateClientAsync(databasePath, "No Plan Data Ltd.");

        var result = await new PlanDesignIntelligenceService(databasePath).AnalyzeClientAsync(clientId);

        Assert.Empty(result.SummaryRows);
        Assert.Contains(result.Findings, finding => finding.FindingMessage.Contains("No booklet data", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(result.Findings, finding => finding.FindingMessage.Contains("No proposal plan design", StringComparison.OrdinalIgnoreCase));
        Assert.All(result.ReadinessRows, row => Assert.Equal(PlanDesignReadinessStatuses.NotReady, row.Status));
    }

    [Fact]
    public async Task AnalyzeClient_Identifies_Missing_Maximums_Coinsurance_And_Deductibles()
    {
        var databasePath = CreateDatabasePath();
        var clientId = await CreateClientAsync(databasePath, "Missing Fields Ltd.");
        var documentId = await InsertDocumentAsync(databasePath, clientId, DocumentTypes.BenefitBooklet, "missing-fields-booklet.pdf");
        await InsertPlanDesignAsync(databasePath, clientId, documentId, "Extended Health", "Drugs", reimbursement: null, maximum: "");

        var result = await new PlanDesignIntelligenceService(databasePath).AnalyzeClientAsync(clientId);
        var messages = result.Findings.Select(finding => finding.FindingMessage).ToList();

        Assert.Contains(messages, message => message.Contains("Maximums", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(messages, message => message.Contains("Coinsurance", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(messages, message => message.Contains("Deductible", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public async Task AnalyzeClient_Identifies_Readiness_Statuses()
    {
        var databasePath = CreateDatabasePath();
        var clientId = await CreateClientAsync(databasePath, "Ready Client Ltd.");
        var documentId = await InsertDocumentAsync(databasePath, clientId, DocumentTypes.BenefitBooklet, "ready-booklet.pdf");
        await InsertPlanDesignAsync(databasePath, clientId, documentId, "Extended Health", "Drugs", reimbursement: 80m, maximum: "Unlimited");
        await InsertPlanDesignFeatureAsync(databasePath, clientId, documentId, "Extended Health", "Drugs", "Coinsurance", "80%", "80%", "$25", "Unlimited");
        await InsertRenewalAsync(databasePath, clientId, documentId, "EHC");
        await InsertRateAsync(databasePath, clientId, documentId, "EHC");
        await CreateProposalPlanDesignTableAsync(databasePath, "proposal_plan_design");
        await CreateProposalPlanDesignTableAsync(databasePath, "proposal_booklet_validation");
        await InsertProposalPlanDesignAsync(databasePath, "proposal_plan_design", clientId, "Extended Health", "Coinsurance", "80%", "80%", "match");
        await InsertProposalPlanDesignAsync(databasePath, "proposal_booklet_validation", clientId, "Extended Health", "Coinsurance", "80%", "80%", "match");

        var result = await new PlanDesignIntelligenceService(databasePath).AnalyzeClientAsync(clientId);

        AssertReadiness(result, "Current booklet vs proposal", PlanDesignReadinessStatuses.Ready);
        AssertReadiness(result, "Current plan design vs renewal", PlanDesignReadinessStatuses.Ready);
        AssertReadiness(result, "Current plan design vs quote", PlanDesignReadinessStatuses.Ready);
        AssertReadiness(result, "Proposal vs booklet validation", PlanDesignReadinessStatuses.Ready);
    }

    private static string CreateDatabasePath()
    {
        var testFolder = Path.Combine(Path.GetTempPath(), "PlanDesignIntelligenceTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(testFolder);
        var databasePath = Path.Combine(testFolder, "Benefits_Intelligence.db");
        var initialize = new DatabaseService(databasePath).Initialize();
        Assert.True(initialize.Succeeded, initialize.Message);
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

    private static async Task<long> InsertDocumentAsync(string databasePath, long clientId, string documentType, string fileName)
    {
        await using var connection = await OpenConnectionAsync(databasePath);
        using var command = connection.CreateCommand();
        command.CommandText = @"
INSERT INTO Documents (
    client_id, document_type, module_name, file_name, file_type, file_path, file_hash,
    carrier, effective_date, upload_date, processed_date, processing_status
)
VALUES (
    $client_id, $document_type, 'Plan Design Intelligence Tests', $file_name, 'PDF', $file_name, $file_hash,
    'Manulife', '2026-01-01', $now, $now, 'Processed'
);
SELECT last_insert_rowid();";
        AddParameter(command, "$client_id", clientId);
        AddParameter(command, "$document_type", documentType);
        AddParameter(command, "$file_name", fileName);
        AddParameter(command, "$file_hash", Guid.NewGuid().ToString("N"));
        AddParameter(command, "$now", NowText());
        return Convert.ToInt64(await command.ExecuteScalarAsync(), CultureInfo.InvariantCulture);
    }

    private static async Task InsertPlanDesignAsync(
        string databasePath,
        long clientId,
        long documentId,
        string category,
        string benefitName,
        decimal? reimbursement,
        string maximum)
    {
        await using var connection = await OpenConnectionAsync(databasePath);
        using var command = connection.CreateCommand();
        command.CommandText = @"
INSERT INTO PlanDesign (
    client_id, benefit_category, benefit_name, coverage_description, reimbursement_percent,
    maximum_amount, frequency, waiting_period, termination_age, source_document_id, source_page, date_loaded
)
VALUES (
    $client_id, $category, $benefit_name, 'Prescription drugs', $reimbursement,
    $maximum, 'Per calendar year', 'None', 'Age 70', $document_id, 12, $now
);";
        AddParameter(command, "$client_id", clientId);
        AddParameter(command, "$category", category);
        AddParameter(command, "$benefit_name", benefitName);
        AddParameter(command, "$reimbursement", reimbursement);
        AddParameter(command, "$maximum", maximum);
        AddParameter(command, "$document_id", documentId);
        AddParameter(command, "$now", NowText());
        await command.ExecuteNonQueryAsync();
    }

    private static async Task InsertPlanDesignFeatureAsync(
        string databasePath,
        long clientId,
        long documentId,
        string category,
        string benefit,
        string featureName,
        string featureValue,
        string coinsurance,
        string deductible,
        string maximum)
    {
        await using var connection = await OpenConnectionAsync(databasePath);
        using var command = connection.CreateCommand();
        command.CommandText = @"
INSERT INTO plan_design_features (
    client_id, source_document_id, benefit, category, feature_name, feature_value,
    coinsurance, deductible, maximum, frequency, termination_age, source_page,
    source_text, confidence_score, date_loaded
)
VALUES (
    $client_id, $document_id, $benefit, $category, $feature_name, $feature_value,
    $coinsurance, $deductible, $maximum, 'Per calendar year', 'Age 70', 12,
    'Feature text', 0.95, $now
);";
        AddParameter(command, "$client_id", clientId);
        AddParameter(command, "$document_id", documentId);
        AddParameter(command, "$benefit", benefit);
        AddParameter(command, "$category", category);
        AddParameter(command, "$feature_name", featureName);
        AddParameter(command, "$feature_value", featureValue);
        AddParameter(command, "$coinsurance", coinsurance);
        AddParameter(command, "$deductible", deductible);
        AddParameter(command, "$maximum", maximum);
        AddParameter(command, "$now", NowText());
        await command.ExecuteNonQueryAsync();
    }

    private static async Task InsertRenewalAsync(string databasePath, long clientId, long documentId, string benefitLine)
    {
        await using var connection = await OpenConnectionAsync(databasePath);
        using var command = connection.CreateCommand();
        command.CommandText = @"
INSERT INTO Renewals (
    client_id, renewal_year, carrier, benefit_line, current_rate, renewal_rate,
    rate_change_percent, premium, claims, loss_ratio, lives, renewal_status,
    source_document_id, date_loaded
)
VALUES (
    $client_id, 2026, 'Manulife', $benefit, 100, 110,
    10, 10000, 7500, 75, 10, 'Loaded',
    $document_id, $now
);";
        AddParameter(command, "$client_id", clientId);
        AddParameter(command, "$benefit", benefitLine);
        AddParameter(command, "$document_id", documentId);
        AddParameter(command, "$now", NowText());
        await command.ExecuteNonQueryAsync();
    }

    private static async Task InsertRateAsync(string databasePath, long clientId, long documentId, string benefitLine)
    {
        await using var connection = await OpenConnectionAsync(databasePath);
        using var command = connection.CreateCommand();
        command.CommandText = @"
INSERT INTO Rates (
    client_id, carrier, benefit_line, class, current_rate, proposed_rate,
    negotiated_rate, rate_basis, effective_date, source_document_id, date_loaded
)
VALUES (
    $client_id, 'Manulife', $benefit, 'All', 100, 110,
    NULL, 'composite', '2026-01-01', $document_id, $now
);";
        AddParameter(command, "$client_id", clientId);
        AddParameter(command, "$benefit", benefitLine);
        AddParameter(command, "$document_id", documentId);
        AddParameter(command, "$now", NowText());
        await command.ExecuteNonQueryAsync();
    }

    private static async Task CreateProposalPlanDesignTableAsync(string databasePath, string tableName)
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

    private static async Task InsertProposalPlanDesignAsync(
        string databasePath,
        string tableName,
        long clientId,
        string benefitLine,
        string provisionName,
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
    $client_id, 1, 'proposal.pdf', $hash, 5,
    '', 'High', 'VALID', $now, $now,
    $benefit, $provision, $provision_value, $booklet_value, 'none',
    $validation_status, '{{}}'
);";
        AddParameter(command, "$client_id", clientId.ToString(CultureInfo.InvariantCulture));
        AddParameter(command, "$hash", Guid.NewGuid().ToString("N"));
        AddParameter(command, "$now", NowText());
        AddParameter(command, "$benefit", benefitLine);
        AddParameter(command, "$provision", provisionName);
        AddParameter(command, "$provision_value", provisionValue);
        AddParameter(command, "$booklet_value", bookletValue);
        AddParameter(command, "$validation_status", validationStatus);
        await command.ExecuteNonQueryAsync();
    }

    private static async Task<Dictionary<string, long>> SnapshotTableCountsAsync(string databasePath)
    {
        await using var connection = await OpenConnectionAsync(databasePath);
        using var tableCommand = connection.CreateCommand();
        tableCommand.CommandText = "SELECT name FROM sqlite_master WHERE type = 'table' ORDER BY name;";

        var counts = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
        using var reader = await tableCommand.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            var tableName = reader.GetString(0);
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

    private static void AssertReadiness(PlanDesignIntelligenceResult result, string comparison, string status)
    {
        var row = Assert.Single(result.ReadinessRows, item => item.Comparison == comparison);
        Assert.Equal(status, row.Status);
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
        return DateTime.Now.ToString("O", CultureInfo.InvariantCulture);
    }
}

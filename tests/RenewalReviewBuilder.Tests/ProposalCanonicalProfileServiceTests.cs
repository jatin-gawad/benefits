using System.Collections.Generic;
using System.Globalization;
using Microsoft.Data.Sqlite;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class ProposalCanonicalProfileServiceTests
{
    [Fact]
    public async Task BuildAsync_ProducesCompleteReadOnlyProfileFromExistingTables()
    {
        var databasePath = CreateDatabasePath();
        try
        {
            await CreateFullSchemaAsync(databasePath);
            await SeedProposalAsync(databasePath, 100, "42", 1, "Manulife", "proposal-hash-100");
            var before = await SnapshotTableCountsAsync(databasePath);

            var profile = await new ProposalCanonicalProfileService(databasePath).BuildAsync(42);
            var after = await SnapshotTableCountsAsync(databasePath);

            Assert.Equal(ProposalCanonicalProfileVersions.Current, profile.ProfileVersion);
            Assert.True(profile.IsReadOnly);
            Assert.Equal("Canonical Client Ltd.", profile.ClientName);
            Assert.True(profile.HasProposals);
            Assert.True(profile.CanSupportProposalIntelligence);
            Assert.False(profile.CanCompareCarriers);
            Assert.Equal(ProposalCanonicalReadinessStatuses.Ready, profile.ReadinessStatus);
            Assert.Empty(profile.Issues);

            var proposal = Assert.Single(profile.Proposals);
            Assert.Equal(100, proposal.ProposalId);
            Assert.Equal("Manulife", proposal.CarrierName);
            Assert.Equal("Focused Carrier Comparison", proposal.ExtractionMethod);
            Assert.Equal(0.90m, proposal.ExtractionConfidence);
            Assert.Equal(2, proposal.SourceFiles.Count);
            Assert.Single(proposal.Benefits);
            Assert.Single(proposal.BenefitLines);
            Assert.Single(proposal.Rates);
            Assert.Single(proposal.Premiums);
            Assert.Single(proposal.PlanDesignFacts);
            Assert.Single(proposal.ValidationFacts);
            Assert.Single(proposal.Demographics);
            Assert.Single(proposal.UnderwritingAssumptions);
            Assert.Single(proposal.PoolingRetention);
            Assert.Single(proposal.AuditTrail);
            Assert.Equal(1_000m, proposal.TotalMonthlyPremium);
            Assert.Equal(12_000m, proposal.TotalAnnualPremium);
            Assert.Equal("proposal.pdf", proposal.PlanDesignFacts[0].Evidence.SourceFileName);
            Assert.Equal(5, proposal.PlanDesignFacts[0].Evidence.SourcePage);
            Assert.Equal("current-booklet.pdf", proposal.ValidationFacts[0].BookletEvidence.SourceFileName);
            Assert.Equal(12, proposal.ValidationFacts[0].BookletEvidence.SourcePage);
            Assert.Equal("Canada Life", proposal.ValidationFacts[0].BookletCarrier);
            Assert.Equal("Structured booklet plan design field", proposal.ValidationFacts[0].ComparisonMethod);
            Assert.Equal(ProposalCanonicalReadinessStatuses.Ready, proposal.ReadinessStatus);
            Assert.Contains(profile.Sources, source => source.TableName == "proposal_premiums" && source.Used && source.RowsFound == 1);
            Assert.Contains(profile.Sources, source => source.TableName == "proposal_exceptions" && source.Exists && !source.Used);
            AssertTableCountsEqual(before, after);
        }
        finally
        {
            DeleteDatabase(databasePath);
        }
    }

    [Fact]
    public async Task BuildAsync_PreservesMultipleProposalVersionsAndNameBasedClientKeys()
    {
        var databasePath = CreateDatabasePath();
        try
        {
            await CreateFullSchemaAsync(databasePath);
            await SeedProposalAsync(databasePath, 100, "42", 1, "Manulife", "proposal-hash-old");
            await SeedProposalAsync(databasePath, 101, "Canonical Client Ltd.", 2, "RBC", "proposal-hash-new");

            var profile = await new ProposalCanonicalProfileService(databasePath).BuildAsync(42);

            Assert.Equal(2, profile.Proposals.Count);
            Assert.Equal([100L, 101L], profile.Proposals.Select(proposal => proposal.ProposalId).ToArray());
            Assert.Equal(2, profile.Proposals.Select(proposal => proposal.SourceFileHash).Distinct().Count());
            Assert.True(profile.CanCompareCarriers);
            Assert.True(profile.CanCompareFinancials);
            Assert.True(profile.CanComparePlanDesign);
            Assert.All(profile.Proposals, proposal => Assert.Equal(ProposalCanonicalReadinessStatuses.Ready, proposal.ReadinessStatus));
            Assert.Equal(ProposalCanonicalReadinessStatuses.Ready, profile.ReadinessStatus);
        }
        finally
        {
            DeleteDatabase(databasePath);
        }
    }

    [Fact]
    public async Task BuildAsync_MismatchAndSavedExceptionRequireReview()
    {
        var databasePath = CreateDatabasePath();
        try
        {
            await CreateFullSchemaAsync(databasePath);
            await SeedProposalAsync(
                databasePath,
                100,
                "42",
                1,
                "Manulife",
                "proposal-hash-mismatch",
                proposalValue: "90%",
                bookletValue: "80%",
                validationStatus: "DIFFERENT");
            await InsertDetailAsync(
                databasePath,
                "proposal_exceptions",
                100,
                "exception_type, field_name, message",
                "'NEEDS REVIEW', 'proposal_booklet_validation', 'Multiple booklet class values require review.'");

            var profile = await new ProposalCanonicalProfileService(databasePath).BuildAsync(42);

            var proposal = Assert.Single(profile.Proposals);
            Assert.Equal(1, proposal.ValidationMismatchCount);
            Assert.Equal(2, proposal.ReviewIssueCount);
            Assert.Equal(ProposalCanonicalReadinessStatuses.NeedsReview, proposal.ReadinessStatus);
            Assert.Equal(ProposalCanonicalReadinessStatuses.NeedsReview, profile.ReadinessStatus);
            Assert.Contains(profile.Issues, issue => issue.Code == "proposal_booklet_mismatch");
            Assert.Contains(profile.Issues, issue => issue.Code == "proposal_exceptions_need_review");
        }
        finally
        {
            DeleteDatabase(databasePath);
        }
    }

    [Fact]
    public async Task BuildAsync_TreatsMissingOptionalTablesAsPartialInsteadOfFailing()
    {
        var databasePath = CreateDatabasePath();
        try
        {
            await CreateMinimalSchemaAsync(databasePath);
            await SeedMinimalProposalAsync(databasePath);

            var profile = await new ProposalCanonicalProfileService(databasePath).BuildAsync(42);

            var proposal = Assert.Single(profile.Proposals);
            Assert.True(proposal.HasCoreFacts);
            Assert.False(proposal.HasFinancialFacts);
            Assert.False(proposal.HasBookletValidation);
            Assert.Equal(ProposalCanonicalReadinessStatuses.Partial, proposal.ReadinessStatus);
            Assert.Equal(ProposalCanonicalReadinessStatuses.Partial, profile.ReadinessStatus);
            Assert.Contains(profile.Issues, issue => issue.Code == "proposal_financial_facts_missing");
            Assert.Contains(profile.Issues, issue => issue.Code == "proposal_validation_missing");
            Assert.Contains(profile.Sources, source => source.TableName == "proposal_rates" && !source.Exists);
            Assert.Contains(profile.Sources, source => source.TableName == "proposal_plan_design" && source.Used);
        }
        finally
        {
            DeleteDatabase(databasePath);
        }
    }

    [Fact]
    public async Task BuildAsync_NoSavedProposalIsNotReady()
    {
        var databasePath = CreateDatabasePath();
        try
        {
            await CreateFullSchemaAsync(databasePath);

            var profile = await new ProposalCanonicalProfileService(databasePath).BuildAsync(42);

            Assert.False(profile.HasProposals);
            Assert.False(profile.CanSupportProposalIntelligence);
            Assert.Equal(ProposalCanonicalReadinessStatuses.NotReady, profile.ReadinessStatus);
            Assert.Contains(profile.Issues, issue => issue.Code == "proposals_missing" && issue.Severity == ProposalCanonicalIssueSeverities.Error);
        }
        finally
        {
            DeleteDatabase(databasePath);
        }
    }

    [Fact]
    public async Task BuildAsync_MissingDatabaseDoesNotCreateAFile()
    {
        var databasePath = CreateDatabasePath(createFile: false);
        try
        {
            var profile = await new ProposalCanonicalProfileService(databasePath).BuildAsync(42);

            Assert.False(File.Exists(databasePath));
            Assert.Equal(ProposalCanonicalReadinessStatuses.NotReady, profile.ReadinessStatus);
            Assert.Contains(profile.Issues, issue => issue.Code == "database_missing");
        }
        finally
        {
            DeleteDatabase(databasePath);
        }
    }

    [Fact]
    public async Task BuildAllAsync_ReturnsOnlyClientsWithSavedProposalsAndPreservesHistory()
    {
        var databasePath = CreateDatabasePath();
        try
        {
            await CreateFullSchemaAsync(databasePath);
            await ExecuteWithNewConnectionAsync(databasePath, "INSERT INTO Clients (client_id, client_name, client_number, status) VALUES (84, 'Second Client Ltd.', 'CAN-84', 'Active');");
            await ExecuteWithNewConnectionAsync(databasePath, "INSERT INTO Clients (client_id, client_name, client_number, status) VALUES (126, 'No Proposal Client Ltd.', 'CAN-126', 'Active');");
            await SeedProposalAsync(databasePath, 100, "42", 1, "Manulife", "proposal-hash-100");
            await SeedProposalAsync(databasePath, 101, "Canonical Client Ltd.", 2, "RBC", "proposal-hash-101");
            await SeedProposalAsync(databasePath, 200, "Second Client Ltd.", 3, "Canada Life", "proposal-hash-200");
            await using (var connection = await OpenConnectionAsync(databasePath))
            {
                await ExecuteAsync(connection, "UPDATE proposal_summary SET client_name = 'Second Client Ltd.' WHERE proposal_id = 200;");
            }

            var before = await SnapshotTableCountsAsync(databasePath);
            var profiles = await new ProposalCanonicalProfileService(databasePath).BuildAllAsync();
            var after = await SnapshotTableCountsAsync(databasePath);

            Assert.Equal(2, profiles.Count);
            Assert.Equal([42L, 84L], profiles.Select(profile => profile.ClientId).OrderBy(value => value).ToArray());
            Assert.Equal(2, profiles.Single(profile => profile.ClientId == 42).Proposals.Count);
            Assert.Single(profiles.Single(profile => profile.ClientId == 84).Proposals);
            Assert.DoesNotContain(profiles, profile => profile.ClientId == 126);
            AssertTableCountsEqual(before, after);
        }
        finally
        {
            DeleteDatabase(databasePath);
        }
    }

    [Fact]
    public async Task ProposalPipelineOverview_UsesCanonicalReadinessAndSavedProposalVersions()
    {
        var databasePath = CreateDatabasePath();
        try
        {
            await CreateFullSchemaAsync(databasePath);
            await SeedProposalAsync(databasePath, 100, "42", 1, "Manulife", "proposal-hash-ready");
            await SeedProposalAsync(
                databasePath,
                101,
                "42",
                2,
                "RBC",
                "proposal-hash-review",
                proposalValue: "90%",
                bookletValue: "80%",
                validationStatus: "DIFFERENT");
            var oldTimestamp = DateTime.UtcNow.AddDays(-20).ToString("O", CultureInfo.InvariantCulture);
            await using (var connection = await OpenConnectionAsync(databasePath))
            {
                await ExecuteAsync(connection, "UPDATE proposal_files SET created_at = $timestamp, updated_at = $timestamp;", ("$timestamp", oldTimestamp));
                await ExecuteAsync(connection, "UPDATE proposal_summary SET created_at = $timestamp, updated_at = $timestamp;", ("$timestamp", oldTimestamp));
            }

            var before = await SnapshotTableCountsAsync(databasePath);
            var overview = await new ProposalPipelineOverviewService(databasePath).BuildAsync();
            var after = await SnapshotTableCountsAsync(databasePath);

            Assert.Equal(ProposalCanonicalProfileVersions.Current, overview.ProfileVersion);
            Assert.Equal(9, overview.TotalStages);
            Assert.Equal(1, overview.ClientCount);
            Assert.Equal(2, overview.ProposalCount);
            Assert.Equal(1, overview.NeedsReviewCount);
            Assert.Equal(2, overview.OpenItems);
            Assert.Equal(24_000m, overview.TotalPremium);
            Assert.Equal(2, overview.OverdueItems);

            var comparison = Assert.Single(overview.Stages, stage => stage.Stage == "Comparison generated");
            Assert.Equal(1, comparison.Count);
            Assert.Equal(12_000m, comparison.Premium);
            Assert.Contains("Canonical Client Ltd. - Manulife", comparison.OldestItem, StringComparison.Ordinal);

            var underwriting = Assert.Single(overview.Stages, stage => stage.Stage == "Underwriting review");
            Assert.Equal(1, underwriting.Count);
            Assert.Equal(12_000m, underwriting.Premium);
            Assert.Contains("Canonical Client Ltd. - RBC", underwriting.OldestItem, StringComparison.Ordinal);
            AssertTableCountsEqual(before, after);
        }
        finally
        {
            DeleteDatabase(databasePath);
        }
    }

    [Fact]
    public async Task ProposalPipelineOverview_DoesNotInferLaterStageOverStoredWorkflowStatus()
    {
        var databasePath = CreateDatabasePath();
        try
        {
            await CreateFullSchemaAsync(databasePath);
            await SeedProposalAsync(databasePath, 100, "42", 1, "Manulife", "proposal-hash-presentation");
            await using (var connection = await OpenConnectionAsync(databasePath))
            {
                await ExecuteAsync(connection, "UPDATE proposal_files SET status = 'PRESENTATION READY' WHERE proposal_id = 100;");
                await ExecuteAsync(connection, "UPDATE proposal_summary SET status = 'PRESENTATION READY' WHERE proposal_id = 100;");
            }

            var overview = await new ProposalPipelineOverviewService(databasePath).BuildAsync();

            Assert.Equal(1, overview.Stages.Single(stage => stage.Stage == "Presentation ready").Count);
            Assert.Equal(0, overview.Stages.Single(stage => stage.Stage == "Consultant review").Count);
            Assert.Equal(0, overview.Stages.Single(stage => stage.Stage == "Sold / Lost").Count);
        }
        finally
        {
            DeleteDatabase(databasePath);
        }
    }

    private static string CreateDatabasePath(bool createFile = true)
    {
        var folder = Path.Combine(Path.GetTempPath(), "ProposalCanonicalProfileTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        var databasePath = Path.Combine(folder, "Benefits_Intelligence.db");
        if (createFile)
        {
            using var stream = File.Create(databasePath);
        }

        return databasePath;
    }

    private static async Task CreateFullSchemaAsync(string databasePath)
    {
        await using var connection = await OpenConnectionAsync(databasePath);
        await ExecuteAsync(connection, "CREATE TABLE Clients (client_id INTEGER PRIMARY KEY, client_name TEXT, client_number TEXT, status TEXT);");
        await ExecuteAsync(connection, "INSERT INTO Clients (client_id, client_name, client_number, status) VALUES (42, 'Canonical Client Ltd.', 'CAN-42', 'Active');");
        await ExecuteAsync(connection, "CREATE TABLE carriers (id INTEGER PRIMARY KEY, carrier_key TEXT, carrier_name TEXT);");
        await ExecuteAsync(connection, """
            CREATE TABLE proposal_files (
                id INTEGER PRIMARY KEY, client_id TEXT, carrier_id INTEGER, proposal_id INTEGER,
                source_file_name TEXT, source_file_hash TEXT, source_page_number INTEGER, source_sheet_name TEXT,
                extraction_confidence TEXT, status TEXT, created_at TEXT, updated_at TEXT,
                file_path TEXT, detected_carrier TEXT, detection_confidence TEXT, reprocess_count INTEGER
            );
            """);
        await ExecuteAsync(connection, """
            CREATE TABLE proposal_summary (
                id INTEGER PRIMARY KEY, client_id TEXT, carrier_id INTEGER, proposal_id INTEGER,
                source_file_name TEXT, source_file_hash TEXT, source_page_number INTEGER, source_sheet_name TEXT,
                extraction_confidence TEXT, status TEXT, created_at TEXT, updated_at TEXT,
                client_name TEXT, carrier_name TEXT, proposal_date TEXT, effective_date TEXT,
                detected_carrier TEXT, detection_confidence TEXT, payload_json TEXT
            );
            """);
        await CreateDetailTableAsync(connection, "benefits", "benefit_key TEXT, benefit_name TEXT, normalized_benefit_name TEXT, display_name TEXT, benefit_category TEXT, sort_order INTEGER, is_standard INTEGER, is_known INTEGER");
        await CreateDetailTableAsync(connection, "proposal_benefit_lines", "benefit_line TEXT, class_name TEXT, tier TEXT, lives REAL, volume REAL, rate REAL, monthly_premium REAL, annual_premium REAL");
        await CreateDetailTableAsync(connection, "proposal_rates", "benefit_line TEXT, class_name TEXT, tier TEXT, age_band TEXT, gender TEXT, province TEXT, lives REAL, volume REAL, rate REAL, rate_basis TEXT, monthly_premium REAL, annual_premium REAL");
        await CreateDetailTableAsync(connection, "proposal_premiums", "benefit_line TEXT, class_name TEXT, tier TEXT, monthly_premium REAL, annual_premium REAL, employer_cost REAL, employee_cost REAL, total_cost REAL");
        await CreateDetailTableAsync(connection, "proposal_plan_design", "benefit_line TEXT, provision_name TEXT, provision_value TEXT, booklet_value TEXT, variance TEXT, validation_status TEXT");
        await CreateDetailTableAsync(connection, "proposal_booklet_validation", "benefit_line TEXT, provision_name TEXT, provision_value TEXT, booklet_value TEXT, variance TEXT, validation_status TEXT");
        await CreateDetailTableAsync(connection, "proposal_demographics", "benefit_line TEXT, class_name TEXT, tier TEXT, lives REAL, volume REAL");
        await CreateDetailTableAsync(connection, "proposal_underwriting_assumptions", "assumption_name TEXT, assumption_value TEXT");
        await CreateDetailTableAsync(connection, "proposal_pooling_retention", "benefit_line TEXT, pooling_level TEXT, pooling_charge REAL, commission REAL, retention REAL, target_loss_ratio REAL, rate_guarantee TEXT, pooled_benefits TEXT, experience_rated_benefits TEXT, ibnr TEXT, rate_caps TEXT, first_renewal_cap TEXT, second_renewal_cap TEXT");
        await CreateDetailTableAsync(connection, "proposal_exceptions", "exception_type TEXT, field_name TEXT, message TEXT");
        await CreateDetailTableAsync(connection, "proposal_export_audit", "action TEXT, table_name TEXT, rows_loaded INTEGER, message TEXT");
    }

    private static async Task CreateMinimalSchemaAsync(string databasePath)
    {
        await using var connection = await OpenConnectionAsync(databasePath);
        await ExecuteAsync(connection, "CREATE TABLE Clients (client_id INTEGER PRIMARY KEY, client_name TEXT, client_number TEXT);");
        await ExecuteAsync(connection, "INSERT INTO Clients (client_id, client_name, client_number) VALUES (42, 'Canonical Client Ltd.', 'CAN-42');");
        await ExecuteAsync(connection, "CREATE TABLE carriers (id INTEGER PRIMARY KEY, carrier_key TEXT, carrier_name TEXT);");
        await ExecuteAsync(connection, """
            CREATE TABLE proposal_files (
                id INTEGER PRIMARY KEY, client_id TEXT, carrier_id INTEGER, proposal_id INTEGER,
                source_file_name TEXT, source_file_hash TEXT, extraction_confidence TEXT, status TEXT,
                created_at TEXT, updated_at TEXT, file_path TEXT, detected_carrier TEXT,
                detection_confidence TEXT, reprocess_count INTEGER
            );
            """);
        await ExecuteAsync(connection, """
            CREATE TABLE proposal_summary (
                id INTEGER PRIMARY KEY, client_id TEXT, carrier_id INTEGER, proposal_id INTEGER,
                source_file_name TEXT, source_file_hash TEXT, extraction_confidence TEXT, status TEXT,
                created_at TEXT, updated_at TEXT, client_name TEXT, carrier_name TEXT,
                proposal_date TEXT, effective_date TEXT, detected_carrier TEXT, detection_confidence TEXT, payload_json TEXT
            );
            """);
        await CreateDetailTableAsync(connection, "proposal_plan_design", "benefit_line TEXT, provision_name TEXT, provision_value TEXT, booklet_value TEXT, variance TEXT, validation_status TEXT");
    }

    private static async Task CreateDetailTableAsync(SqliteConnection connection, string tableName, string detailColumns)
    {
        await ExecuteAsync(
            connection,
            $"""
            CREATE TABLE [{tableName}] (
                id INTEGER PRIMARY KEY AUTOINCREMENT, client_id TEXT, carrier_id INTEGER, proposal_id INTEGER,
                source_file_name TEXT, source_file_hash TEXT, source_page_number INTEGER, source_sheet_name TEXT,
                extraction_confidence TEXT, status TEXT, created_at TEXT, updated_at TEXT, payload_json TEXT,
                {detailColumns}
            );
            """);
    }

    private static async Task SeedProposalAsync(
        string databasePath,
        long proposalId,
        string clientKey,
        long carrierId,
        string carrierName,
        string sourceHash,
        string proposalValue = "80%",
        string bookletValue = "80%",
        string validationStatus = "MATCH")
    {
        await using var connection = await OpenConnectionAsync(databasePath);
        await ExecuteAsync(
            connection,
            "INSERT OR IGNORE INTO carriers (id, carrier_key, carrier_name) VALUES ($carrier_id, $carrier_key, $carrier_name);",
            ("$carrier_id", carrierId),
            ("$carrier_key", carrierName.ToLowerInvariant()),
            ("$carrier_name", carrierName));
        await ExecuteAsync(
            connection,
            """
            INSERT INTO proposal_files (
                id, client_id, carrier_id, proposal_id, source_file_name, source_file_hash,
                extraction_confidence, status, created_at, updated_at, file_path,
                detected_carrier, detection_confidence, reprocess_count
            )
            VALUES (
                $proposal_id, $client_id, $carrier_id, $proposal_id, 'proposal.pdf', $source_hash,
                'HIGH', 'VALID', '2026-07-18T12:00:00Z', '2026-07-18T12:00:00Z', 'C:\quotes\proposal.pdf',
                $carrier_key, 'HIGH', 1
            );
            """,
            ("$proposal_id", proposalId),
            ("$client_id", clientKey),
            ("$carrier_id", carrierId),
            ("$source_hash", sourceHash),
            ("$carrier_key", carrierName.ToLowerInvariant()));
        await ExecuteAsync(
            connection,
            """
            INSERT INTO proposal_summary (
                client_id, carrier_id, proposal_id, source_file_name, source_file_hash,
                extraction_confidence, status, created_at, updated_at, client_name, carrier_name,
                proposal_date, effective_date, detected_carrier, detection_confidence, payload_json
            )
            VALUES (
                $client_id, $carrier_id, $proposal_id, 'proposal.pdf', $source_hash,
                'HIGH', 'VALID', '2026-07-18T12:00:00Z', '2026-07-18T12:00:00Z',
                'Canonical Client Ltd.', $carrier_name, '2026-07-01', '2026-08-01',
                $carrier_key, 'HIGH', $payload
            );
            """,
            ("$client_id", clientKey),
            ("$carrier_id", carrierId),
            ("$proposal_id", proposalId),
            ("$source_hash", sourceHash),
            ("$carrier_name", carrierName),
            ("$carrier_key", carrierName.ToLowerInvariant()),
            ("$payload", "{\"extraction_method\":\"Focused Carrier Comparison\",\"source_files\":[\"proposal.pdf\",\"rates.xlsx\"]}"));

        await InsertDetailAsync(databasePath, "benefits", proposalId, "benefit_key, benefit_name, normalized_benefit_name, display_name, benefit_category, sort_order, is_standard, is_known", "'extended-health-care', 'Extended Health Care', 'Extended Health Care', 'Extended Health Care', 'Health', 1, 1, 1", connection);
        await InsertDetailAsync(databasePath, "proposal_benefit_lines", proposalId, "benefit_line, class_name, tier, lives, volume, rate, monthly_premium, annual_premium", "'Extended Health Care', 'All Classes', 'Single', 10, 100000, 1.25, 1000, 12000", connection);
        await InsertDetailAsync(databasePath, "proposal_rates", proposalId, "benefit_line, class_name, tier, lives, volume, rate, rate_basis, monthly_premium, annual_premium", "'Extended Health Care', 'All Classes', 'Single', 10, 100000, 1.25, 'Per $1,000', 1000, 12000", connection);
        await InsertDetailAsync(databasePath, "proposal_premiums", proposalId, "benefit_line, class_name, tier, monthly_premium, annual_premium, employer_cost, employee_cost, total_cost", "'Extended Health Care', 'All Classes', 'Single', 1000, 12000, 9000, 3000, 12000", connection);
        await InsertDetailAsync(databasePath, "proposal_plan_design", proposalId, "benefit_line, provision_name, provision_value, validation_status", $"'Extended Health Care', 'Coinsurance', '{proposalValue}', 'VALID'", connection);
        await InsertDetailAsync(databasePath, "proposal_booklet_validation", proposalId, "benefit_line, provision_name, provision_value, booklet_value, variance, validation_status", $"'Extended Health Care', 'Coinsurance', '{proposalValue}', '{bookletValue}', {(validationStatus == "MATCH" ? "NULL" : "'Different values'")}, '{validationStatus}'", connection);
        await ExecuteAsync(
            connection,
            """
            UPDATE proposal_booklet_validation
            SET payload_json = $payload
            WHERE proposal_id = $proposal_id;
            """,
            ("$payload", "{\"booklet_source_file\":\"current-booklet.pdf\",\"booklet_source_page\":12,\"booklet_source_section\":\"Extended Health Care\",\"booklet_carrier\":\"Canada Life\",\"booklet_confidence_score\":\"HIGH\",\"booklet_class_names\":[\"All Employees\"],\"comparison_method\":\"Structured booklet plan design field\"}"),
            ("$proposal_id", proposalId));
        await InsertDetailAsync(databasePath, "proposal_demographics", proposalId, "benefit_line, class_name, tier, lives, volume", "'Extended Health Care', 'All Classes', 'Single', 10, 100000", connection);
        await InsertDetailAsync(databasePath, "proposal_underwriting_assumptions", proposalId, "assumption_name, assumption_value", "'Rate Guarantee', '24 months'", connection);
        await InsertDetailAsync(databasePath, "proposal_pooling_retention", proposalId, "benefit_line, pooling_level, rate_guarantee", "'All Benefits', '$10,000', '24 months'", connection);
        await InsertDetailAsync(databasePath, "proposal_export_audit", proposalId, "action, table_name, rows_loaded, message", "'load', 'proposal_plan_design', 1, 'Loaded 1 row.'", connection);
    }

    private static async Task SeedMinimalProposalAsync(string databasePath)
    {
        await using var connection = await OpenConnectionAsync(databasePath);
        await ExecuteAsync(connection, "INSERT INTO carriers (id, carrier_key, carrier_name) VALUES (1, 'manulife', 'Manulife');");
        await ExecuteAsync(connection, "INSERT INTO proposal_files (id, client_id, carrier_id, proposal_id, source_file_name, source_file_hash, extraction_confidence, status, file_path, detected_carrier, detection_confidence, reprocess_count) VALUES (100, '42', 1, 100, 'proposal.pdf', 'minimal-hash', 'HIGH', 'VALID', 'proposal.pdf', 'manulife', 'HIGH', 0);");
        await ExecuteAsync(connection, "INSERT INTO proposal_summary (client_id, carrier_id, proposal_id, source_file_name, source_file_hash, extraction_confidence, status, client_name, carrier_name, detected_carrier, detection_confidence, payload_json) VALUES ('42', 1, 100, 'proposal.pdf', 'minimal-hash', 'HIGH', 'VALID', 'Canonical Client Ltd.', 'Manulife', 'manulife', 'HIGH', '{}');");
        await InsertDetailAsync(databasePath, "proposal_plan_design", 100, "benefit_line, provision_name, provision_value, validation_status", "'Extended Health Care', 'Coinsurance', '80%', 'VALID'", connection);
    }

    private static async Task InsertDetailAsync(
        string databasePath,
        string tableName,
        long proposalId,
        string columns,
        string values,
        SqliteConnection? existingConnection = null)
    {
        if (existingConnection is not null)
        {
            await InsertDetailAsync(existingConnection, tableName, proposalId, columns, values);
            return;
        }

        await using var connection = await OpenConnectionAsync(databasePath);
        await InsertDetailAsync(connection, tableName, proposalId, columns, values);
    }

    private static async Task InsertDetailAsync(
        SqliteConnection connection,
        string tableName,
        long proposalId,
        string columns,
        string values)
    {
        await ExecuteAsync(
            connection,
            $"""
            INSERT INTO [{tableName}] (
                client_id, carrier_id, proposal_id, source_file_name, source_file_hash,
                source_page_number, source_sheet_name, extraction_confidence, status,
                created_at, updated_at, payload_json, {columns}
            )
            VALUES (
                '42', 1, {proposalId.ToString(CultureInfo.InvariantCulture)}, 'proposal.pdf', 'row-hash-{proposalId.ToString(CultureInfo.InvariantCulture)}',
                5, 'Plan Design', 'HIGH', 'VALID',
                '2026-07-18T12:00:00Z', '2026-07-18T12:00:00Z', $payload, {values}
            );
            """,
            ("$payload", "{\"source_type\":\"PDF\",\"source_location\":\"Page 5\"}"));
    }

    private static async Task<Dictionary<string, long>> SnapshotTableCountsAsync(string databasePath)
    {
        await using var connection = await OpenConnectionAsync(databasePath);
        using var tablesCommand = connection.CreateCommand();
        tablesCommand.CommandText = "SELECT name FROM sqlite_master WHERE type = 'table' AND name NOT LIKE 'sqlite_%' ORDER BY name;";
        var tables = new List<string>();
        using (var reader = await tablesCommand.ExecuteReaderAsync())
        {
            while (await reader.ReadAsync())
            {
                tables.Add(reader.GetString(0));
            }
        }

        var counts = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
        foreach (var table in tables)
        {
            using var countCommand = connection.CreateCommand();
            countCommand.CommandText = $"SELECT COUNT(*) FROM [{table}];";
            counts[table] = Convert.ToInt64(await countCommand.ExecuteScalarAsync(), CultureInfo.InvariantCulture);
        }

        return counts;
    }

    private static void AssertTableCountsEqual(
        IReadOnlyDictionary<string, long> expected,
        IReadOnlyDictionary<string, long> actual)
    {
        Assert.Equal(expected.Count, actual.Count);
        foreach (var pair in expected)
        {
            Assert.True(actual.TryGetValue(pair.Key, out var count), $"Table '{pair.Key}' was removed.");
            Assert.Equal(pair.Value, count);
        }
    }

    private static async Task<SqliteConnection> OpenConnectionAsync(string databasePath)
    {
        var connection = new SqliteConnection($"Data Source={databasePath}");
        await connection.OpenAsync();
        return connection;
    }

    private static async Task ExecuteWithNewConnectionAsync(string databasePath, string sql)
    {
        await using var connection = await OpenConnectionAsync(databasePath);
        await ExecuteAsync(connection, sql);
    }

    private static async Task ExecuteAsync(
        SqliteConnection connection,
        string sql,
        params (string Name, object Value)[] parameters)
    {
        using var command = connection.CreateCommand();
        command.CommandText = sql;
        foreach (var parameter in parameters)
        {
            command.Parameters.AddWithValue(parameter.Name, parameter.Value);
        }

        await command.ExecuteNonQueryAsync();
    }

    private static void DeleteDatabase(string databasePath)
    {
        SqliteConnection.ClearAllPools();
        var folder = Path.GetDirectoryName(databasePath);
        if (!string.IsNullOrWhiteSpace(folder) && Directory.Exists(folder))
        {
            Directory.Delete(folder, recursive: true);
        }
    }
}

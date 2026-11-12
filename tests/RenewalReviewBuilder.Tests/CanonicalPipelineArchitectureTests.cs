using System.Collections.Generic;
using System.Text.RegularExpressions;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class CanonicalPipelineArchitectureTests
{
    private static readonly Regex UiBoundaryTypePattern = new(
        @"\b[A-Z][A-Za-z0-9_]*(?:Parser|Extractor|Runner)\b",
        RegexOptions.CultureInvariant);

    private static readonly Regex DirectSqlWritePattern = new(
        @"\b(?:INSERT\s+INTO|UPDATE\s+[A-Za-z_][A-Za-z0-9_]*\s+SET|DELETE\s+FROM)\b",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    private static readonly Regex CanonicalPersistenceCallPattern = new(
        @"\b(?:Persist|Save|Insert|Update|Delete)[A-Za-z0-9_]*Canonical[A-Za-z0-9_]*Async\b",
        RegexOptions.CultureInvariant);

    private static readonly IReadOnlyDictionary<string, int> ApprovedUiCompatibilityReferences =
        new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            ["MainWindow.xaml.cs|BookletExtractorRunner"] = 2,
            ["MainWindow.xaml.cs|ProposalEngineRunner"] = 1,
            ["Views/ClientBillingDashboardView.xaml.cs|GroupHealthBillingStatementParser"] = 1
        };

    private static readonly IReadOnlyDictionary<string, int> ApprovedEngineExportCompatibilityReferences =
        new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            ["ExcelReportWriter.cs|NemMaximumExtractor"] = 4,
            ["PowerPointDeckWriter.cs|NemMaximumExtractor"] = 4
        };

    [Fact]
    public void OfficialStandard_IsPresentAndDefinesEveryRequiredLayer()
    {
        var root = FindRepositoryRoot();
        var standardPath = Path.Combine(root, "docs", "architecture", "canonical-data-pipeline-standard.md");

        Assert.True(File.Exists(standardPath), $"Canonical pipeline standard was not found at {standardPath}.");

        var text = File.ReadAllText(standardPath);
        foreach (var requiredText in new[]
                 {
                     "Status: Official desktop architecture standard",
                     "Document discovery",
                     "Raw extraction",
                     "Carrier or source mapper",
                     "Canonical contract",
                     "Validation and evidence",
                     "SQLite persistence",
                     "Query or profile service",
                     "Business engine",
                     "Workspace UI and exports"
                 })
        {
            Assert.Contains(requiredText, text, StringComparison.OrdinalIgnoreCase);
        }
    }

    [Fact]
    public void DesktopUi_CannotAddDatabaseOrCanonicalPersistenceShortcuts()
    {
        var root = FindRepositoryRoot();
        var appRoot = Path.Combine(root, "src", "RenewalReviewBuilder.App");

        foreach (var file in SourceFiles(appRoot, "*.xaml.cs"))
        {
            var text = File.ReadAllText(file);
            var relativePath = RelativeAppPath(appRoot, file);

            Assert.DoesNotContain("Microsoft.Data.Sqlite", text, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("SqliteConnection", text, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("SqliteCommand", text, StringComparison.OrdinalIgnoreCase);
            Assert.False(DirectSqlWritePattern.IsMatch(text), $"UI file {relativePath} contains a direct SQL write.");
            Assert.False(
                CanonicalPersistenceCallPattern.IsMatch(text),
                $"UI file {relativePath} calls canonical persistence directly.");
        }
    }

    [Fact]
    public void DesktopUi_CannotGrowLegacyParserRunnerOrExtractorDependencies()
    {
        var root = FindRepositoryRoot();
        var appRoot = Path.Combine(root, "src", "RenewalReviewBuilder.App");
        var observedReferences = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        foreach (var file in SourceFiles(appRoot, "*.xaml.cs"))
        {
            var relativePath = RelativeAppPath(appRoot, file);
            var text = File.ReadAllText(file);
            foreach (Match match in UiBoundaryTypePattern.Matches(text))
            {
                var key = $"{relativePath}|{match.Value}";
                observedReferences[key] = observedReferences.TryGetValue(key, out var count) ? count + 1 : 1;
            }
        }

        foreach (var reference in observedReferences)
        {
            Assert.True(
                ApprovedUiCompatibilityReferences.TryGetValue(reference.Key, out var maximumCount),
                $"UI introduced a new parser, extractor, or runner dependency: {reference.Key}.");
            Assert.True(
                reference.Value <= maximumCount,
                $"UI compatibility dependency {reference.Key} grew from its approved maximum of {maximumCount} to {reference.Value} references.");
        }
    }

    [Fact]
    public void EnginesAndExports_CannotOpenSqliteOrGrowSourceBoundaryDependencies()
    {
        var root = FindRepositoryRoot();
        var serviceRoot = Path.Combine(root, "src", "RenewalReviewBuilder.App", "Services");
        var protectedFiles = SourceFiles(serviceRoot, "*.cs")
            .Where(file => IsEngineOrExport(Path.GetFileName(file)))
            .ToArray();
        var observedReferences = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        Assert.NotEmpty(protectedFiles);

        foreach (var file in protectedFiles)
        {
            var text = File.ReadAllText(file);

            Assert.DoesNotContain("Microsoft.Data.Sqlite", text, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("SqliteConnection", text, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("SqliteCommand", text, StringComparison.OrdinalIgnoreCase);

            foreach (Match match in UiBoundaryTypePattern.Matches(text))
            {
                var key = $"{Path.GetFileName(file)}|{match.Value}";
                observedReferences[key] = observedReferences.TryGetValue(key, out var count) ? count + 1 : 1;
            }
        }

        foreach (var reference in observedReferences)
        {
            Assert.True(
                ApprovedEngineExportCompatibilityReferences.TryGetValue(reference.Key, out var maximumCount),
                $"Engine or export introduced a new source-boundary dependency: {reference.Key}.");
            Assert.True(
                reference.Value <= maximumCount,
                $"Engine or export compatibility dependency {reference.Key} grew from its approved maximum of {maximumCount} to {reference.Value} references.");
        }
    }

    [Fact]
    public void PlanDesignCanonicalProfile_RemainsReadOnlyAndConnectedToIntelligence()
    {
        var root = FindRepositoryRoot();
        var appRoot = Path.Combine(root, "src", "RenewalReviewBuilder.App");
        var profileModelPath = Path.Combine(appRoot, "Models", "PlanDesignCanonicalProfileModels.cs");
        var profileServicePath = Path.Combine(appRoot, "Services", "PlanDesignCanonicalProfileService.cs");
        var profileBuilderPath = Path.Combine(appRoot, "Services", "PlanDesignIntelligenceService.CanonicalProfile.cs");
        var intelligencePath = Path.Combine(appRoot, "Services", "PlanDesignIntelligenceService.cs");

        foreach (var path in new[] { profileModelPath, profileServicePath, profileBuilderPath, intelligencePath })
        {
            Assert.True(File.Exists(path), $"Canonical Plan Design profile component was not found at {path}.");
        }

        var profileService = File.ReadAllText(profileServicePath);
        var profileBuilder = File.ReadAllText(profileBuilderPath);
        var intelligence = File.ReadAllText(intelligencePath);
        var protectedText = string.Join(Environment.NewLine, profileService, profileBuilder, intelligence);

        Assert.Contains("Task<PlanDesignCanonicalProfile> BuildAsync", profileService, StringComparison.Ordinal);
        Assert.Contains("BuildCanonicalProfileAsync(clientId)", intelligence, StringComparison.Ordinal);
        Assert.Contains("Mode = SqliteOpenMode.ReadOnly", intelligence, StringComparison.Ordinal);
        Assert.Contains("PRAGMA query_only = ON", intelligence, StringComparison.Ordinal);
        Assert.False(DirectSqlWritePattern.IsMatch(protectedText), "Canonical Plan Design profile path contains a direct SQL write.");
    }

    [Fact]
    public void ControlledBookletReprocessing_UsesStagingApprovalAndExistingPlanDesignPersistence()
    {
        var root = FindRepositoryRoot();
        var appRoot = Path.Combine(root, "src", "RenewalReviewBuilder.App");
        var modelPath = Path.Combine(appRoot, "Models", "ControlledBookletReprocessingModels.cs");
        var servicePath = Path.Combine(appRoot, "Services", "CanonicalControlledBookletReprocessingService.cs");
        var persistencePath = Path.Combine(appRoot, "Services", "DatabaseService.ControlledBookletReprocessing.cs");
        var identityPath = Path.Combine(appRoot, "Services", "BookletClientIdentityGuardrail.cs");
        var viewPath = Path.Combine(appRoot, "Views", "ControlledBookletReprocessingView.xaml.cs");
        var viewXamlPath = Path.Combine(appRoot, "Views", "ControlledBookletReprocessingView.xaml");
        var documentPath = Path.Combine(root, "docs", "architecture", "canonical-controlled-booklet-reprocessing-v1.md");
        var workflowTestPath = Path.Combine(root, "tests", "RenewalReviewBuilder.Tests", "CanonicalControlledBookletReprocessingWorkflowTests.cs");

        foreach (var path in new[]
                 {
                     modelPath, servicePath, persistencePath, identityPath, viewPath, viewXamlPath, documentPath, workflowTestPath
                 })
        {
            Assert.True(File.Exists(path), $"Controlled booklet reprocessing component was not found at {path}.");
        }

        var model = File.ReadAllText(modelPath);
        var service = File.ReadAllText(servicePath);
        var persistence = File.ReadAllText(persistencePath);
        var view = File.ReadAllText(viewPath);
        var viewXaml = File.ReadAllText(viewXamlPath);
        var document = File.ReadAllText(documentPath);

        Assert.Contains("CANONICAL_CONTROLLED_BOOKLET_REPROCESSING_V1", model, StringComparison.Ordinal);
        Assert.Contains("GetControlledBookletDocumentAsync", service, StringComparison.Ordinal);
        Assert.Contains("ComputeSha256Async", service, StringComparison.Ordinal);
        Assert.Contains("BuildChanges", service, StringComparison.Ordinal);
        Assert.Contains("CreateControlledBookletBackupAsync", service, StringComparison.Ordinal);
        Assert.Contains("expectedSourceDocumentId", service, StringComparison.Ordinal);
        Assert.Contains("GetActiveRunForDocumentAsync", service, StringComparison.Ordinal);
        Assert.Contains("ControlledBookletReprocessingRun", persistence, StringComparison.Ordinal);
        Assert.Contains("ControlledBookletReprocessingIssue", persistence, StringComparison.Ordinal);
        Assert.Contains("ControlledBookletReprocessingChange", persistence, StringComparison.Ordinal);
        Assert.Contains("source.BackupDatabase(destination)", persistence, StringComparison.Ordinal);
        Assert.Contains("Stage Reprocess", viewXaml, StringComparison.Ordinal);
        Assert.Contains("Approve and Apply", viewXaml, StringComparison.Ordinal);
        Assert.Contains("CanonicalControlledBookletReprocessingService", view, StringComparison.Ordinal);
        Assert.Contains("LoadSelectedDocumentRunAsync", view, StringComparison.Ordinal);
        Assert.Contains("!hasActiveRun", view, StringComparison.Ordinal);
        Assert.Contains("ApplyChangeFilter", view, StringComparison.Ordinal);
        Assert.Contains("ShowUnchangedRowsCheckBox", viewXaml, StringComparison.Ordinal);
        Assert.DoesNotContain("BookletExtractorRunner", view, StringComparison.Ordinal);
        Assert.DoesNotContain("Microsoft.Data.Sqlite", view, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("SqliteConnection", view, StringComparison.OrdinalIgnoreCase);
        Assert.False(DirectSqlWritePattern.IsMatch(view), "Controlled booklet UI contains a direct SQL write.");
        Assert.Contains("does not create another Plan Design fact store", document, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ProposalComparisonWorkspace_ComposesCanonicalAnalysisAndCurrentVsQuoteWithoutDirectSqliteAccess()
    {
        var root = FindRepositoryRoot();
        var appRoot = Path.Combine(root, "src", "RenewalReviewBuilder.App");
        var comparisonServicePath = Path.Combine(appRoot, "Services", "CurrentVsQuotePlanDesignService.cs");
        var analysisServicePath = Path.Combine(appRoot, "Services", "ProposalIntelligenceAnalysisService.cs");
        var workspaceDataServicePath = Path.Combine(appRoot, "Services", "ProposalComparisonDataService.cs");
        var workspaceViewPath = Path.Combine(appRoot, "Views", "ProposalComparisonWorkspaceView.xaml.cs");
        var workspaceXamlPath = Path.Combine(appRoot, "Views", "ProposalComparisonWorkspaceView.xaml");
        var dataServiceTestPath = Path.Combine(root, "tests", "RenewalReviewBuilder.Tests", "ProposalComparisonDataServiceTests.cs");
        var mockServicePath = Path.Combine(appRoot, "Services", "ProposalComparisonMockDataService.cs");

        foreach (var path in new[] { comparisonServicePath, analysisServicePath, workspaceDataServicePath, workspaceViewPath, workspaceXamlPath, dataServiceTestPath })
        {
            Assert.True(File.Exists(path), $"Proposal Comparison canonical consumer was not found at {path}.");
        }

        var comparisonService = File.ReadAllText(comparisonServicePath);
        var workspaceDataService = File.ReadAllText(workspaceDataServicePath);
        var workspaceView = File.ReadAllText(workspaceViewPath);
        var workspaceXaml = File.ReadAllText(workspaceXamlPath);
        var protectedText = string.Join(Environment.NewLine, comparisonService, workspaceDataService, workspaceView);

        Assert.Contains("PlanDesignCanonicalProfileService", comparisonService, StringComparison.Ordinal);
        Assert.Contains("ProposalIntelligenceAnalysisService", workspaceDataService, StringComparison.Ordinal);
        Assert.Contains("CurrentVsQuotePlanDesignService", workspaceDataService, StringComparison.Ordinal);
        Assert.Contains("_analysisService.BuildAsync", workspaceDataService, StringComparison.Ordinal);
        Assert.Contains("_currentVsQuoteService.BuildAsync", workspaceDataService, StringComparison.Ordinal);
        Assert.Contains("ProposalComparisonDataService", workspaceView, StringComparison.Ordinal);
        Assert.Contains("ProposalSummaryGrid", workspaceXaml, StringComparison.Ordinal);
        Assert.Contains("ProposalPremiumGrid", workspaceXaml, StringComparison.Ordinal);
        Assert.Contains("ProposalRateGrid", workspaceXaml, StringComparison.Ordinal);
        Assert.Contains("ProposalCurrentVsQuoteGrid", workspaceXaml, StringComparison.Ordinal);
        Assert.Contains("ProposalCarrierPlanDesignGrid", workspaceXaml, StringComparison.Ordinal);
        Assert.Contains("ProposalUnderwritingGrid", workspaceXaml, StringComparison.Ordinal);
        Assert.Contains("Build Comparison Workbook", workspaceXaml, StringComparison.Ordinal);
        Assert.Contains("RunExistingComparisonRequested", workspaceView, StringComparison.Ordinal);
        Assert.DoesNotContain("Microsoft.Data.Sqlite", protectedText, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("SqliteConnection", protectedText, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("RecommendedCarrier", protectedText, StringComparison.Ordinal);
        Assert.False(DirectSqlWritePattern.IsMatch(protectedText), "Proposal Comparison canonical consumer contains a direct SQL write.");
        Assert.False(File.Exists(mockServicePath), "The obsolete mock proposal comparison service is still present.");
    }

    [Fact]
    public void FocusedProposalComparison_UsesMapperAndExistingSharedSqliteLoader()
    {
        var root = FindRepositoryRoot();
        var appRoot = Path.Combine(root, "src", "RenewalReviewBuilder.App");
        var runnerPath = Path.Combine(appRoot, "Services", "ProposalEngineRunner.cs");
        var extractorPath = Path.Combine(root, "proposal_engine", "quote_test_extractor.py");
        var mapperPath = Path.Combine(root, "proposal_engine", "focused_comparison_mapper.py");
        var validatorPath = Path.Combine(root, "proposal_engine", "booklet_validator.py");
        var loaderPath = Path.Combine(root, "proposal_engine", "sqlite_loader.py");
        var pythonTestPath = Path.Combine(root, "tests", "python", "test_focused_comparison_persistence.py");
        var validationTestPath = Path.Combine(root, "tests", "python", "test_canonical_proposal_booklet_validation.py");

        foreach (var path in new[] { runnerPath, extractorPath, mapperPath, validatorPath, loaderPath, pythonTestPath, validationTestPath })
        {
            Assert.True(File.Exists(path), $"Focused proposal persistence component was not found at {path}.");
        }

        var runner = File.ReadAllText(runnerPath);
        var extractor = File.ReadAllText(extractorPath);
        var mapper = File.ReadAllText(mapperPath);
        var validator = File.ReadAllText(validatorPath);
        var loader = File.ReadAllText(loaderPath);

        Assert.Contains("proposal_engine.quote_test_extractor", runner, StringComparison.Ordinal);
        Assert.Contains("--database-path", runner, StringComparison.Ordinal);
        Assert.Contains("--client-id", runner, StringComparison.Ordinal);
        Assert.Contains("--client-name", runner, StringComparison.Ordinal);
        Assert.Contains("--booklet-path", runner, StringComparison.Ordinal);
        Assert.Contains("build_focused_comparison_envelope", extractor, StringComparison.Ordinal);
        Assert.Contains("validate_focused_proposals_against_booklet", extractor, StringComparison.Ordinal);
        Assert.Contains("load_to_sqlite", extractor, StringComparison.Ordinal);
        Assert.DoesNotContain("sqlite3", mapper, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("extract_plan_design_rows", validator, StringComparison.Ordinal);
        Assert.DoesNotContain("sqlite3", validator, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("_is_core_clients_table", loader, StringComparison.Ordinal);
        Assert.Contains("source_file_hash = ? AND client_id = ? AND carrier_id = ?", loader, StringComparison.Ordinal);
        Assert.Contains("booklet_validation", loader, StringComparison.Ordinal);
    }

    [Fact]
    public void CanonicalProposalProfile_IsReadOnlyAndUsesExistingProposalTables()
    {
        var root = FindRepositoryRoot();
        var appRoot = Path.Combine(root, "src", "RenewalReviewBuilder.App");
        var modelPath = Path.Combine(appRoot, "Models", "ProposalCanonicalProfileModels.cs");
        var servicePath = Path.Combine(appRoot, "Services", "ProposalCanonicalProfileService.cs");
        var testPath = Path.Combine(root, "tests", "RenewalReviewBuilder.Tests", "ProposalCanonicalProfileServiceTests.cs");

        foreach (var path in new[] { modelPath, servicePath, testPath })
        {
            Assert.True(File.Exists(path), $"Canonical proposal profile component was not found at {path}.");
        }

        var model = File.ReadAllText(modelPath);
        var service = File.ReadAllText(servicePath);

        Assert.Contains("PROPOSAL_CANONICAL_PROFILE_V1", model, StringComparison.Ordinal);
        Assert.Contains("Mode = SqliteOpenMode.ReadOnly", service, StringComparison.Ordinal);
        Assert.Contains("PRAGMA query_only = ON", service, StringComparison.Ordinal);
        Assert.Contains("proposal_files", service, StringComparison.Ordinal);
        Assert.Contains("proposal_rates", service, StringComparison.Ordinal);
        Assert.Contains("proposal_premiums", service, StringComparison.Ordinal);
        Assert.Contains("proposal_plan_design", service, StringComparison.Ordinal);
        Assert.Contains("proposal_booklet_validation", service, StringComparison.Ordinal);
        Assert.Contains("proposal_exceptions", service, StringComparison.Ordinal);
        Assert.Contains("BuildAllAsync", service, StringComparison.Ordinal);
        Assert.DoesNotContain("DatabaseService", service, StringComparison.Ordinal);
        Assert.DoesNotContain("BookletExtractorRunner", service, StringComparison.Ordinal);
        Assert.DoesNotContain("ProposalEngineRunner", service, StringComparison.Ordinal);
        Assert.False(DirectSqlWritePattern.IsMatch(service), "Canonical proposal profile contains a direct SQL write.");
    }

    [Fact]
    public void ProposalPipelineOverview_ConsumesCanonicalProfileWithoutChangingExecutiveReports()
    {
        var root = FindRepositoryRoot();
        var appRoot = Path.Combine(root, "src", "RenewalReviewBuilder.App");
        var modelPath = Path.Combine(appRoot, "Models", "ProposalPipelineOverviewModels.cs");
        var servicePath = Path.Combine(appRoot, "Services", "ProposalPipelineOverviewService.cs");
        var mainWindowPath = Path.Combine(appRoot, "MainWindow.xaml.cs");
        var xamlPath = Path.Combine(appRoot, "MainWindow.xaml");
        var testPath = Path.Combine(root, "tests", "RenewalReviewBuilder.Tests", "ProposalCanonicalProfileServiceTests.cs");

        foreach (var path in new[] { modelPath, servicePath, mainWindowPath, xamlPath, testPath })
        {
            Assert.True(File.Exists(path), $"Canonical proposal pipeline overview component was not found at {path}.");
        }

        var service = File.ReadAllText(servicePath);
        var mainWindow = File.ReadAllText(mainWindowPath);
        var xaml = File.ReadAllText(xamlPath);

        Assert.Contains("ProposalCanonicalProfileService", service, StringComparison.Ordinal);
        Assert.Contains("BuildAllAsync", service, StringComparison.Ordinal);
        Assert.DoesNotContain("Microsoft.Data.Sqlite", service, StringComparison.Ordinal);
        Assert.DoesNotContain("DatabaseService", service, StringComparison.Ordinal);
        Assert.DoesNotContain("ProposalEngineRunner", service, StringComparison.Ordinal);
        Assert.False(DirectSqlWritePattern.IsMatch(service), "Proposal pipeline overview contains a direct SQL write.");

        Assert.Contains("ProposalPipelineOverviewGrid.ItemsSource = _proposalPipelineOverviewRows", mainWindow, StringComparison.Ordinal);
        Assert.DoesNotContain("ProposalPipelineOverviewGrid.ItemsSource = _reportsProposalPipelineRows", mainWindow, StringComparison.Ordinal);
        Assert.Contains("ReportsProposalPipelineGrid.ItemsSource = _reportsProposalPipelineRows", mainWindow, StringComparison.Ordinal);
        Assert.Contains("await _proposalPipelineOverviewService.BuildAsync()", mainWindow, StringComparison.Ordinal);
        Assert.Contains("RefreshProposalPipelineOverview_Click", xaml, StringComparison.Ordinal);
    }

    [Fact]
    public void ProposalIntelligenceAnalysis_IsAReadOnlyCanonicalProfileConsumer()
    {
        var root = FindRepositoryRoot();
        var appRoot = Path.Combine(root, "src", "RenewalReviewBuilder.App");
        var modelPath = Path.Combine(appRoot, "Models", "ProposalIntelligenceAnalysisModels.cs");
        var servicePath = Path.Combine(appRoot, "Services", "ProposalIntelligenceAnalysisService.cs");
        var testPath = Path.Combine(root, "tests", "RenewalReviewBuilder.Tests", "ProposalIntelligenceAnalysisServiceTests.cs");
        var documentPath = Path.Combine(root, "docs", "architecture", "canonical-proposal-intelligence-analysis-v1.md");

        foreach (var path in new[] { modelPath, servicePath, testPath, documentPath })
        {
            Assert.True(File.Exists(path), $"Canonical Proposal Intelligence Analysis V1 component was not found at {path}.");
        }

        var model = File.ReadAllText(modelPath);
        var service = File.ReadAllText(servicePath);

        Assert.Contains("PROPOSAL_INTELLIGENCE_ANALYSIS_V1", model, StringComparison.Ordinal);
        Assert.Contains("ProposalCanonicalProfileService", service, StringComparison.Ordinal);
        Assert.Contains("_profileService.BuildAsync", service, StringComparison.Ordinal);
        Assert.Contains("Build(ProposalCanonicalProfile profile)", service, StringComparison.Ordinal);
        Assert.DoesNotContain("Microsoft.Data.Sqlite", service, StringComparison.Ordinal);
        Assert.DoesNotContain("DatabaseService", service, StringComparison.Ordinal);
        Assert.DoesNotContain("ProposalEngineRunner", service, StringComparison.Ordinal);
        Assert.DoesNotContain("BookletExtractorRunner", service, StringComparison.Ordinal);
        Assert.DoesNotContain("UnderwritingEngine", service, StringComparison.Ordinal);
        Assert.DoesNotContain("QuoteEngine", service, StringComparison.Ordinal);
        Assert.DoesNotContain("RecommendedCarrier", model, StringComparison.Ordinal);
        Assert.DoesNotContain("RecommendedCarrier", service, StringComparison.Ordinal);
        Assert.False(DirectSqlWritePattern.IsMatch(service), "Proposal Intelligence analysis contains a direct SQL write.");
    }

    [Fact]
    public void CanonicalSourceLineageAndValueGuardrails_AreEnforcedAcrossPersistenceAndProfileReads()
    {
        var root = FindRepositoryRoot();
        var appRoot = Path.Combine(root, "src", "RenewalReviewBuilder.App");
        var modelPath = Path.Combine(appRoot, "Models", "PlanDesignCanonicalProfileModels.cs");
        var guardrailPath = Path.Combine(appRoot, "Services", "CanonicalSourceLineageValueGuardrails.cs");
        var databasePath = Path.Combine(appRoot, "Services", "DatabaseService.cs");
        var validatorPath = Path.Combine(appRoot, "Services", "Normalization", "NormalizedDataValidator.cs");
        var mapperPath = Path.Combine(appRoot, "Services", "PlanDesignCanonicalMapper.cs");
        var profileReaderPath = Path.Combine(appRoot, "Services", "PlanDesignIntelligenceService.cs");
        var documentPath = Path.Combine(root, "docs", "architecture", "canonical-source-lineage-value-guardrails-v1.md");
        var testPath = Path.Combine(root, "tests", "RenewalReviewBuilder.Tests", "CanonicalSourceLineageValueGuardrailsTests.cs");

        foreach (var path in new[]
                 {
                     modelPath, guardrailPath, databasePath, validatorPath, mapperPath,
                     profileReaderPath, documentPath, testPath
                 })
        {
            Assert.True(File.Exists(path), $"Canonical source-lineage guardrail component was not found at {path}.");
        }

        var model = File.ReadAllText(modelPath);
        var guardrail = File.ReadAllText(guardrailPath);
        var database = File.ReadAllText(databasePath);
        var validator = File.ReadAllText(validatorPath);
        var mapper = File.ReadAllText(mapperPath);
        var profileReader = File.ReadAllText(profileReaderPath);
        var document = File.ReadAllText(documentPath);

        Assert.Contains("CANONICAL_SOURCE_LINEAGE_VALUE_GUARDRAILS_V1", model, StringComparison.Ordinal);
        Assert.Contains("ResolveSourceFileName", guardrail, StringComparison.Ordinal);
        Assert.Contains("EvaluatePercentage", guardrail, StringComparison.Ordinal);
        Assert.Contains("planDesignImportSources", database, StringComparison.Ordinal);
        Assert.Contains("item.SourceDocumentId == importSource.Key", database, StringComparison.Ordinal);
        Assert.Contains("NormalizePercentage(planRow.ReimbursementPercent)", database, StringComparison.Ordinal);
        Assert.Contains("PlanDesignPercentageOutOfRange", validator, StringComparison.Ordinal);
        Assert.Contains("PlanDesignMultiplePercentagesPreserved", validator, StringComparison.Ordinal);
        Assert.Contains("EvaluatePercentage", mapper, StringComparison.Ordinal);
        Assert.Contains("ResolveSourceFileName", profileReader, StringComparison.Ordinal);
        Assert.Contains("No new SQLite tables or migrations are introduced", document, StringComparison.Ordinal);
    }

    [Fact]
    public void CanonicalLabeledDentalRates_MapAtTheSourceBoundaryAndReuseExistingPersistence()
    {
        var root = FindRepositoryRoot();
        var appRoot = Path.Combine(root, "src", "RenewalReviewBuilder.App");
        var mapperPath = Path.Combine(root, "proposal_engine", "booklet_plan_design", "labeled_dental_rate_mapper.py");
        var extractorPath = Path.Combine(root, "proposal_engine", "booklet_plan_design", "booklet_plan_design_extractor.py");
        var validatorPath = Path.Combine(root, "proposal_engine", "booklet_validator.py");
        var canonicalMapperPath = Path.Combine(appRoot, "Services", "PlanDesignCanonicalMapper.cs");
        var documentPath = Path.Combine(root, "docs", "architecture", "canonical-labeled-dental-rates-v1.md");
        var pythonTestPath = Path.Combine(root, "tests", "python", "test_labeled_dental_rate_mapper.py");

        foreach (var path in new[]
                 {
                     mapperPath, extractorPath, validatorPath, canonicalMapperPath, documentPath, pythonTestPath
                 })
        {
            Assert.True(File.Exists(path), $"Canonical labeled dental component was not found at {path}.");
        }

        var mapper = File.ReadAllText(mapperPath);
        var extractor = File.ReadAllText(extractorPath);
        var validator = File.ReadAllText(validatorPath);
        var canonicalMapper = File.ReadAllText(canonicalMapperPath);
        var document = File.ReadAllText(documentPath);

        Assert.Contains("CANONICAL_LABELED_DENTAL_RATE_MAPPER_V1", mapper, StringComparison.Ordinal);
        Assert.Contains("map_labeled_dental_rates", extractor, StringComparison.Ordinal);
        Assert.Contains("map_labeled_dental_maximums", extractor, StringComparison.Ordinal);
        Assert.Contains("remove_superseded_labeled_dental_rows", extractor, StringComparison.Ordinal);
        Assert.Contains("coinsurance.dental_", mapper, StringComparison.Ordinal);
        Assert.Contains("dental_coinsurance_provision_code", validator, StringComparison.Ordinal);
        Assert.Contains("NormalizeDentalProvisionName", canonicalMapper, StringComparison.Ordinal);
        Assert.Contains("Unlabeled text such as `80%; 50%` is not split or guessed", document, StringComparison.Ordinal);
        Assert.Contains("No new SQLite tables or migrations are introduced", document, StringComparison.Ordinal);
    }

    [Fact]
    public void CanonicalProvisionIdentity_IsTheSharedComparisonAndPersistenceBoundary()
    {
        var root = FindRepositoryRoot();
        var appRoot = Path.Combine(root, "src", "RenewalReviewBuilder.App");
        var modelPath = Path.Combine(appRoot, "Models", "BenefitsIntelligenceModels.cs");
        var normalizedModelPath = Path.Combine(appRoot, "Models", "NormalizedImportModels.cs");
        var mapperPath = Path.Combine(appRoot, "Services", "PlanDesignCanonicalMapper.cs");
        var databasePath = Path.Combine(appRoot, "Services", "DatabaseService.cs");
        var validatorPath = Path.Combine(appRoot, "Services", "Normalization", "NormalizedDataValidator.cs");
        var workflowPath = Path.Combine(appRoot, "Services", "CanonicalControlledBookletReprocessingService.cs");
        var documentPath = Path.Combine(root, "docs", "architecture", "canonical-provision-identity-v1.md");

        foreach (var path in new[]
                 {
                     modelPath, normalizedModelPath, mapperPath, databasePath,
                     validatorPath, workflowPath, documentPath
                 })
        {
            Assert.True(File.Exists(path), $"Canonical provision identity component was not found at {path}.");
        }

        var model = File.ReadAllText(modelPath);
        var normalizedModel = File.ReadAllText(normalizedModelPath);
        var mapper = File.ReadAllText(mapperPath);
        var database = File.ReadAllText(databasePath);
        var validator = File.ReadAllText(validatorPath);
        var workflow = File.ReadAllText(workflowPath);
        var document = File.ReadAllText(documentPath);

        Assert.Contains("ProvisionCode", model, StringComparison.Ordinal);
        Assert.Contains("ProvisionLabel", model, StringComparison.Ordinal);
        Assert.Contains("PlanDesignDuplicateProvisionIdentity", normalizedModel, StringComparison.Ordinal);
        Assert.Contains("CANONICAL_PROVISION_IDENTITY_V1", mapper, StringComparison.Ordinal);
        Assert.Contains("NormalizeProvisionRows", mapper, StringComparison.Ordinal);
        Assert.Contains("BuildProvisionComparisonKey", mapper, StringComparison.Ordinal);
        Assert.Contains("canonicalPlanDesignRows = PlanDesignCanonicalMapper.NormalizeProvisionRows", database, StringComparison.Ordinal);
        Assert.Contains("PlanDesignDuplicateProvisionIdentity", validator, StringComparison.Ordinal);
        Assert.Contains("run.BeforeRows.AddRange(PlanDesignCanonicalMapper.NormalizeProvisionRows", workflow, StringComparison.Ordinal);
        Assert.Contains("BuildProvisionComparisonKey", workflow, StringComparison.Ordinal);
        Assert.Contains("No new SQLite tables or migrations are introduced", document, StringComparison.Ordinal);
    }

    [Fact]
    public void DocumentBatchCompletion_RefreshesRenewalQueueClientListAndHomeDashboardInOrder()
    {
        var root = FindRepositoryRoot();
        var mainWindowPath = Path.Combine(root, "src", "RenewalReviewBuilder.App", "MainWindow.xaml.cs");
        var mainWindow = File.ReadAllText(mainWindowPath);
        var methodStart = mainWindow.IndexOf(
            "private async void ProcessBatch_Click",
            StringComparison.Ordinal);
        var methodEnd = mainWindow.IndexOf(
            "private void OpenBatchLog_Click",
            methodStart,
            StringComparison.Ordinal);

        Assert.True(methodStart >= 0, "ProcessBatch_Click was not found.");
        Assert.True(methodEnd > methodStart, "The end of ProcessBatch_Click was not found.");

        var processBatchMethod = mainWindow[methodStart..methodEnd];
        var renewalQueueRefresh = processBatchMethod.IndexOf(
            "await LoadSavedRenewalQueueAsync();",
            StringComparison.Ordinal);
        var clientListRefresh = processBatchMethod.IndexOf(
            "await LoadSavedClientListAsync();",
            StringComparison.Ordinal);
        var homeDashboardRefresh = processBatchMethod.IndexOf(
            "await RefreshHomeDashboardAsync();",
            StringComparison.Ordinal);

        Assert.True(renewalQueueRefresh >= 0, "Batch completion does not refresh the saved renewal queue.");
        Assert.True(clientListRefresh > renewalQueueRefresh, "Batch completion must refresh the client list after the renewal queue.");
        Assert.True(homeDashboardRefresh > clientListRefresh, "Batch completion must refresh the home dashboard after the client list.");
    }

    private static bool IsEngineOrExport(string fileName) =>
        fileName.Equals("UnderwritingEngine.cs", StringComparison.Ordinal) ||
        fileName.Equals("RenewalEngine.cs", StringComparison.Ordinal) ||
        fileName.EndsWith("Writer.cs", StringComparison.Ordinal) ||
        fileName.EndsWith("Exporter.cs", StringComparison.Ordinal) ||
        fileName.EndsWith("ExportService.cs", StringComparison.Ordinal);

    private static IEnumerable<string> SourceFiles(string root, string searchPattern) =>
        Directory.EnumerateFiles(root, searchPattern, SearchOption.AllDirectories)
            .Where(file => !IsGeneratedPath(file));

    private static bool IsGeneratedPath(string file)
    {
        var normalized = file.Replace('\\', '/');
        return normalized.Contains("/bin/", StringComparison.OrdinalIgnoreCase) ||
               normalized.Contains("/obj/", StringComparison.OrdinalIgnoreCase);
    }

    private static string RelativeAppPath(string appRoot, string file) =>
        Path.GetRelativePath(appRoot, file).Replace('\\', '/');

    private static string FindRepositoryRoot()
    {
        var folder = new DirectoryInfo(AppContext.BaseDirectory);
        while (folder is not null && !File.Exists(Path.Combine(folder.FullName, "RenewalReviewBuilder.sln")))
        {
            folder = folder.Parent;
        }

        return folder?.FullName ?? throw new DirectoryNotFoundException("Repository root was not found.");
    }
}

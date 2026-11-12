using Microsoft.Data.Sqlite;
using System.Collections.Generic;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;
using RenewalReviewBuilder.App.Services.Normalization;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class RenewalCanonicalPlaceholderTests
{
    [Fact]
    public void RenewalVersion_Placeholders_Default_To_Null()
    {
        var version = new RenewalVersion();

        Assert.Null(version.RequestedIncreasePercent);
        Assert.Null(version.NegotiatedIncreasePercent);
        Assert.Null(version.FinalAcceptedIncreasePercent);
        Assert.Null(version.NegotiationStatus);
        Assert.Null(version.NegotiatedSavingsType);
        Assert.Null(version.IndustryBenchmarkPercent);
        Assert.Null(version.RenewalRecommendationType);
        Assert.Null(version.RenewalRiskLevel);
        Assert.Null(version.RenewalCompletedDate);
        Assert.Null(version.RenewalVersionNumber);
        Assert.Null(version.RenewalScenarioId);
        Assert.Null(version.PreviousRenewalScenarioId);
        Assert.Null(version.ScenarioType);
        Assert.Null(version.HistoryEventId);
        Assert.Null(version.RelationshipMilestoneId);
        Assert.Null(version.EvolutionGroupId);
    }

    [Fact]
    public void Normalized_RenewalVersion_Maps_Nullable_Placeholders_Without_Calculation()
    {
        var analysis = new AnalysisResult();
        analysis.RenewalVersions.Add(new RenewalVersion
        {
            VersionName = "Future negotiation",
            RequestedIncreasePercent = 0.075m,
            NegotiationStatus = NegotiationStatus.InNegotiation,
            NegotiatedSavingsType = NegotiatedSavingsType.Recurring,
            IndustryBenchmarkSource = "Future benchmark source",
            RenewalRecommendationType = RenewalRecommendationType.MarketSurvey,
            RenewalRiskLevel = RenewalRiskLevel.Moderate,
            RenewalScenarioId = 42,
            ScenarioType = ScenarioType.Negotiated,
            RelationshipMilestoneId = "relationship:future",
            EvolutionGroupId = "renewal:future"
        });

        var payload = NormalizeRenewalData.FromAnalysis(
            analysis,
            new NormalizedSourceMetadata { SourceType = SourceTypes.Manual });

        var version = Assert.Single(payload.RenewalVersions);
        Assert.Equal(0.075m, version.RequestedIncreasePercent);
        Assert.Equal(NegotiationStatus.InNegotiation, version.NegotiationStatus);
        Assert.Equal(NegotiatedSavingsType.Recurring, version.NegotiatedSavingsType);
        Assert.Equal("Future benchmark source", version.IndustryBenchmarkSource);
        Assert.Equal(RenewalRecommendationType.MarketSurvey, version.RenewalRecommendationType);
        Assert.Equal(RenewalRiskLevel.Moderate, version.RenewalRiskLevel);
        Assert.Equal(42, version.RenewalScenarioId);
        Assert.Equal(ScenarioType.Negotiated, version.ScenarioType);
        Assert.Equal("relationship:future", version.RelationshipMilestoneId);
        Assert.Equal("renewal:future", version.EvolutionGroupId);
    }

    [Fact]
    public async Task Existing_Renewal_Import_And_Executive_Read_Model_Accept_Null_Placeholders()
    {
        var (folder, databasePath, sourcePath) = await CreateDatabaseInputsAsync();

        try
        {
            var service = new DatabaseService(databasePath);
            var save = await service.SaveRenewalAnalysisAsync(BuildAnalysis(sourcePath));

            Assert.True(save.Succeeded, save.Message);
            Assert.True(save.ClientId.HasValue);

            var history = await service.GetClientHistoryAsync(save.ClientId.Value);
            Assert.NotEmpty(history.RenewalVersions);
            Assert.All(history.RenewalVersions, AssertNullPlaceholders);

            var reporting = await service.GetExecutiveReportingSnapshotAsync(forceRefresh: true);
            Assert.Contains(reporting.RenewalVersions, version => version.ClientId == save.ClientId.Value);

            await AssertNullablePlaceholderColumnsAsync(databasePath);
        }
        finally
        {
            SqliteConnection.ClearAllPools();
            if (Directory.Exists(folder))
            {
                Directory.Delete(folder, recursive: true);
            }
        }
    }

    [Fact]
    public async Task RenewalVersion_Placeholders_Round_Trip_Without_Changing_Existing_Renewal_Values()
    {
        var (folder, databasePath, sourcePath) = await CreateDatabaseInputsAsync();

        try
        {
            var analysis = BuildAnalysis(sourcePath);
            analysis.RenewalVersions.Add(new RenewalVersion
            {
                VersionName = "Future populated renewal",
                SourceModule = BenefitsIntelligenceModules.RenewalIntelligence,
                Status = "Pending",
                IsSelected = true,
                RequestedIncreasePercent = 0.080m,
                NegotiatedIncreasePercent = 0.055m,
                FinalAcceptedIncreasePercent = 0.050m,
                NegotiationStatus = NegotiationStatus.Accepted,
                NegotiationStage = "Final review",
                NegotiationCompletedDate = new DateOnly(2026, 7, 31),
                NegotiatedSavingsAnnual = 12000m,
                NegotiatedSavingsMonthly = 1000m,
                NegotiatedSavingsType = NegotiatedSavingsType.Recurring,
                NegotiatedBy = "Future consultant",
                IndustryBenchmarkPercent = 0.065m,
                IndustryBenchmarkSource = "Future source",
                IndustryBenchmarkPopulation = "Future population",
                IndustryBenchmarkYear = 2026,
                IndustryBenchmarkConfidence = 0.90m,
                RenewalRecommendationType = RenewalRecommendationType.Maintain,
                RenewalDecisionReason = "Future decision reason",
                RenewalStrategySummary = "Future strategy summary",
                RenewalRiskLevel = RenewalRiskLevel.Low,
                RenewalComplexity = "Low",
                RenewalResult = "Accepted",
                RenewalCompletedDate = new DateOnly(2026, 8, 1),
                RenewalApprovedBy = "Future approver",
                RenewalAcceptedDate = new DateOnly(2026, 8, 1),
                RenewalVersionNumber = 3,
                RenewalScenarioId = 40,
                PreviousRenewalScenarioId = 39,
                ScenarioType = ScenarioType.Accepted,
                HistoryEventId = 12,
                RelationshipMilestoneId = "relationship:renewal-2026",
                EvolutionGroupId = "renewal:2026"
            });

            var service = new DatabaseService(databasePath);
            var save = await service.SaveRenewalAnalysisAsync(analysis);
            Assert.True(save.Succeeded, save.Message);

            var persisted = (await service.GetRenewalVersionsAsync(save.ClientId!.Value))
                .Single(version => version.VersionName == "Future populated renewal");

            Assert.Equal(0.080m, persisted.RequestedIncreasePercent);
            Assert.Equal(0.055m, persisted.NegotiatedIncreasePercent);
            Assert.Equal(0.050m, persisted.FinalAcceptedIncreasePercent);
            Assert.Equal(NegotiationStatus.Accepted, persisted.NegotiationStatus);
            Assert.Equal("Final review", persisted.NegotiationStage);
            Assert.Equal(new DateOnly(2026, 7, 31), persisted.NegotiationCompletedDate);
            Assert.Equal(12000m, persisted.NegotiatedSavingsAnnual);
            Assert.Equal(1000m, persisted.NegotiatedSavingsMonthly);
            Assert.Equal(NegotiatedSavingsType.Recurring, persisted.NegotiatedSavingsType);
            Assert.Equal("Future consultant", persisted.NegotiatedBy);
            Assert.Equal(0.065m, persisted.IndustryBenchmarkPercent);
            Assert.Equal("Future source", persisted.IndustryBenchmarkSource);
            Assert.Equal("Future population", persisted.IndustryBenchmarkPopulation);
            Assert.Equal(2026, persisted.IndustryBenchmarkYear);
            Assert.Equal(0.90m, persisted.IndustryBenchmarkConfidence);
            Assert.Equal(RenewalRecommendationType.Maintain, persisted.RenewalRecommendationType);
            Assert.Equal("Future decision reason", persisted.RenewalDecisionReason);
            Assert.Equal("Future strategy summary", persisted.RenewalStrategySummary);
            Assert.Equal(RenewalRiskLevel.Low, persisted.RenewalRiskLevel);
            Assert.Equal("Low", persisted.RenewalComplexity);
            Assert.Equal("Accepted", persisted.RenewalResult);
            Assert.Equal(new DateOnly(2026, 8, 1), persisted.RenewalCompletedDate);
            Assert.Equal("Future approver", persisted.RenewalApprovedBy);
            Assert.Equal(new DateOnly(2026, 8, 1), persisted.RenewalAcceptedDate);
            Assert.Equal(3, persisted.RenewalVersionNumber);
            Assert.Equal(40, persisted.RenewalScenarioId);
            Assert.Equal(39, persisted.PreviousRenewalScenarioId);
            Assert.Equal(ScenarioType.Accepted, persisted.ScenarioType);
            Assert.Equal(12, persisted.HistoryEventId);
            Assert.Equal("relationship:renewal-2026", persisted.RelationshipMilestoneId);
            Assert.Equal("renewal:2026", persisted.EvolutionGroupId);

            var renewal = Assert.Single(await service.GetRenewalHistoryAsync(save.ClientId.Value));
            Assert.Equal(100m, renewal.CurrentRate);
            Assert.Equal(105m, renewal.RenewalRate);
        }
        finally
        {
            SqliteConnection.ClearAllPools();
            if (Directory.Exists(folder))
            {
                Directory.Delete(folder, recursive: true);
            }
        }
    }

    private static void AssertNullPlaceholders(RenewalVersionRecord version)
    {
        Assert.Null(version.RequestedIncreasePercent);
        Assert.Null(version.NegotiatedIncreasePercent);
        Assert.Null(version.FinalAcceptedIncreasePercent);
        Assert.Null(version.NegotiationStatus);
        Assert.Null(version.NegotiationStage);
        Assert.Null(version.NegotiationCompletedDate);
        Assert.Null(version.NegotiatedSavingsAnnual);
        Assert.Null(version.NegotiatedSavingsMonthly);
        Assert.Null(version.NegotiatedSavingsType);
        Assert.Null(version.NegotiatedBy);
        Assert.Null(version.IndustryBenchmarkPercent);
        Assert.Null(version.IndustryBenchmarkSource);
        Assert.Null(version.IndustryBenchmarkPopulation);
        Assert.Null(version.IndustryBenchmarkYear);
        Assert.Null(version.IndustryBenchmarkConfidence);
        Assert.Null(version.RenewalRecommendationType);
        Assert.Null(version.RenewalDecisionReason);
        Assert.Null(version.RenewalStrategySummary);
        Assert.Null(version.RenewalRiskLevel);
        Assert.Null(version.RenewalComplexity);
        Assert.Null(version.RenewalResult);
        Assert.Null(version.RenewalCompletedDate);
        Assert.Null(version.RenewalApprovedBy);
        Assert.Null(version.RenewalAcceptedDate);
        Assert.Null(version.RenewalVersionNumber);
        Assert.Null(version.RenewalScenarioId);
        Assert.Null(version.PreviousRenewalScenarioId);
        Assert.Null(version.ScenarioType);
        Assert.Null(version.HistoryEventId);
        Assert.Null(version.RelationshipMilestoneId);
        Assert.Null(version.EvolutionGroupId);
    }

    private static async Task AssertNullablePlaceholderColumnsAsync(string databasePath)
    {
        var expectedColumns = new[]
        {
            "requested_increase_percent", "negotiated_increase_percent", "final_accepted_increase_percent",
            "negotiation_status", "negotiation_stage", "negotiation_completed_date", "negotiated_savings_annual",
            "negotiated_savings_monthly", "negotiated_savings_type", "negotiated_by", "industry_benchmark_percent",
            "industry_benchmark_source", "industry_benchmark_population", "industry_benchmark_year",
            "industry_benchmark_confidence", "renewal_recommendation_type", "renewal_decision_reason",
            "renewal_strategy_summary", "renewal_risk_level", "renewal_complexity", "renewal_result",
            "renewal_completed_date", "renewal_approved_by", "renewal_accepted_date", "renewal_version_number",
            "renewal_scenario_id", "previous_renewal_scenario_id", "scenario_type", "history_event_id",
            "relationship_milestone_id", "evolution_group_id"
        };

        await using var connection = new SqliteConnection($"Data Source={databasePath}");
        await connection.OpenAsync();
        await using var command = connection.CreateCommand();
        command.CommandText = "PRAGMA table_info(renewal_versions);";
        await using var reader = await command.ExecuteReaderAsync();
        var nullableColumns = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        while (await reader.ReadAsync())
        {
            nullableColumns[reader.GetString(1)] = reader.GetInt64(3) == 0;
        }

        foreach (var column in expectedColumns)
        {
            Assert.True(nullableColumns.TryGetValue(column, out var isNullable), $"Missing {column}.");
            Assert.True(isNullable, $"{column} must remain nullable.");
        }
    }

    private static async Task<(string Folder, string DatabasePath, string SourcePath)> CreateDatabaseInputsAsync()
    {
        var folder = Path.Combine(Path.GetTempPath(), "RenewalCanonicalPlaceholderTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        var databasePath = Path.Combine(folder, "Benefits_Intelligence.db");
        var sourcePath = Path.Combine(folder, "renewal.pdf");
        await File.WriteAllTextAsync(sourcePath, "placeholder renewal source");
        return (folder, databasePath, sourcePath);
    }

    private static AnalysisResult BuildAnalysis(string sourcePath)
    {
        var analysis = new AnalysisResult
        {
            RenewalCase = new RenewalCase
            {
                ClientName = "Renewal Placeholder Test Ltd.",
                PolicyNumber = "RPT-2026",
                PdfPath = sourcePath,
                RenewalDate = new DateOnly(2026, 8, 1)
            },
            Pdf = new PdfExtractionResult
            {
                ClientName = "Renewal Placeholder Test Ltd.",
                PolicyNumber = "RPT-2026",
                Carrier = "Test Carrier",
                RenewalDate = new DateOnly(2026, 8, 1)
            }
        };
        analysis.Pdf.RateLines.Add(new RateLine
        {
            Benefit = "Extended Health Care",
            CurrentRate = 100m,
            RenewalRate = 105m,
            CurrentPremium = 1000m,
            RenewalPremium = 1050m,
            Lives = 10
        });
        return analysis;
    }
}

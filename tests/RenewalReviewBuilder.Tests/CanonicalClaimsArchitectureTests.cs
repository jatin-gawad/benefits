using System.Globalization;
using System.Security.Cryptography;
using Microsoft.Data.Sqlite;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;
using RenewalReviewBuilder.App.Services.CanonicalClaims;
using RenewalReviewBuilder.App.Services.Normalization;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class CanonicalClaimsArchitectureTests
{
    [Fact]
    public void Utilization_ValidatorMetricsAndReconciliation_PreserveSemiAdditiveClaimants()
    {
        var envelope = new CanonicalClaimsImportEnvelope { Report = new CanonicalClaimsReport { ClientId=1, Carrier="Manulife", PolicyNumber="P", Period=new CanonicalClaimsPeriod { Start=new DateOnly(2025,1,1), End=new DateOnly(2025,12,31) } } };
        envelope.UtilizationFacts.AddRange([
            Utilization(CanonicalClaimsSourceAggregationLevel.Sponsor, CanonicalClaimsRelationshipType.Employee, 2, 4, 400m, 300m, 200m),
            Utilization(CanonicalClaimsSourceAggregationLevel.Sponsor, CanonicalClaimsRelationshipType.Dependant, 1, 2, 200m, 150m, 100m),
            Utilization(CanonicalClaimsSourceAggregationLevel.Sponsor, CanonicalClaimsRelationshipType.Partner, 1, 1, 100m, 100m, 100m),
            Utilization(CanonicalClaimsSourceAggregationLevel.Sponsor, CanonicalClaimsRelationshipType.Total, 4, 7, 700m, 550m, 400m),
            Utilization(CanonicalClaimsSourceAggregationLevel.Contract, CanonicalClaimsRelationshipType.Total, null, 7, 700m, 550m, 400m)
        ]);
        var validation = new NormalizedDataValidator().Validate(envelope);
        Assert.DoesNotContain(validation, x => x.Code == "canonical_utilization_claimants_exceed_occurrences");
        Assert.DoesNotContain(validation, x => x.Code == "canonical_utilization_duplicate_grain");
        Assert.All(new CanonicalClaimsParityService().ReconcileUtilization(envelope), x => Assert.Equal(CanonicalClaimsParityClassification.ExactMatch, x.Classification));
        var metrics=CanonicalClaimsUtilizationMetricCalculator.Calculate(envelope.UtilizationFacts[0]);
        Assert.Equal(2m, metrics.OccurrencesPerClaimant);
        Assert.Equal(50m, metrics.PaidPerOccurrence);
        var diagnostic=CanonicalClaimsUtilizationMetricCalculator.Summarize(envelope.UtilizationFacts, validation.Count);
        Assert.Equal(5, diagnostic.FactCount); Assert.Equal(21, diagnostic.TotalOccurrences); Assert.Equal(1200m, diagnostic.TotalPaid);
    }

    [Fact]
    public void Utilization_Validator_WarnsForOrderingAndDoesNotRejectAdjustment()
    {
        var envelope = new CanonicalClaimsImportEnvelope { Report = new CanonicalClaimsReport { ClientId=1, Carrier="Manulife", PolicyNumber="P", Period=new CanonicalClaimsPeriod { Start=new DateOnly(2025,1,1), End=new DateOnly(2025,12,31) } } };
        envelope.UtilizationFacts.Add(Utilization(CanonicalClaimsSourceAggregationLevel.Sponsor, CanonicalClaimsRelationshipType.Total, 1, 0, 10m, 0m, 5m));
        var results=new NormalizedDataValidator().Validate(envelope);
        Assert.Contains(results, x=>x.Code=="canonical_utilization_submitted_without_occurrences");
        Assert.Contains(results, x=>x.Code=="canonical_utilization_paid_without_eligible");
        Assert.All(results.Where(x=>x.Code.StartsWith("canonical_utilization_")), x=>Assert.NotEqual("Error",x.Severity));
    }

    private static CanonicalClaimsUtilizationFact Utilization(CanonicalClaimsSourceAggregationLevel level, CanonicalClaimsRelationshipType relationship, long? claimants, long? occurrences, decimal? submitted, decimal? eligible, decimal? paid) => new()
    { BenefitCode="EHC", BenefitName="Health", CategoryCode="DRUGS", CategoryName="Drugs", ServiceCode="DRUG", ServiceName="Drug", ProvinceCode="BC", RelationshipType=relationship, UniqueClaimants=claimants, Occurrences=occurrences, SubmittedAmount=submitted, EligibleAmount=eligible, PaidAmount=paid, MeasurementStartDate=new DateOnly(2025,1,1), MeasurementEndDate=new DateOnly(2025,12,31), SourceAggregationLevel=level, Evidence=new CanonicalClaimsSourceEvidence { SourceDocumentId=1, ImportBatchId=1, WorksheetName="Sponsor Summary", SourceRow=9, SourceColumnOrCell="E9:I9" } };
    [Fact]
    public void Mapper_PreservesTotalsLabelsAvailabilityAndEvidence()
    {
        var legacy = BuildLegacyResult();
        legacy.Summaries.Add(new ClaimsExperienceSummary { BenefitLine = "Vision", Premium = 0m, PaidClaims = 0m });
        legacy.Summaries.Add(new ClaimsExperienceSummary { BenefitLine = "Unknown carrier label" });

        var envelope = new ManulifeClaimsCanonicalMapper().Map(legacy, 42, 11, 12);

        Assert.Equal("MANULIFE_CLAIMS_CANONICAL_V1", envelope.MappingVersion);
        Assert.Equal(new DateOnly(2025, 1, 1), envelope.Report.Period.Start);
        Assert.Equal("EHC", envelope.FinancialFacts[0].Benefit);
        Assert.Equal(6500m, envelope.FinancialFacts[0].PaidClaims);
        Assert.Equal(CanonicalDataAvailability.ReportedZero, envelope.FinancialFacts[^2].Availability);
        Assert.Equal(CanonicalDataAvailability.Missing, envelope.FinancialFacts[^1].Availability);
        Assert.Equal("Unknown carrier label", envelope.FinancialFacts[^1].Benefit);
        Assert.Equal(11, envelope.FinancialFacts[0].Evidence.SourceDocumentId);
        Assert.Equal(12, envelope.FinancialFacts[0].Evidence.ImportBatchId);
        Assert.Equal(CanonicalClaimsEvidencePrecision.Document, envelope.FinancialFacts[0].Evidence.Precision);
    }

    [Fact]
    public void Validator_RejectsMissingPeriodAndDuplicateFacts()
    {
        var envelope = new CanonicalClaimsImportEnvelope
        {
            Report = new CanonicalClaimsReport { ClientId = 1, Carrier = "Manulife", PolicyNumber = "123" }
        };
        envelope.FinancialFacts.Add(new CanonicalClaimFinancialFact { FactId = "same", Benefit = "EHC" });
        envelope.MonthlyFacts.Add(new CanonicalClaimMonthlyFact { FactId = "same", Benefit = "EHC", Month = new DateOnly(2025, 1, 1) });

        var results = new NormalizedDataValidator().Validate(envelope);

        Assert.Contains(results, item => item.Code == "canonical_period_missing" && item.Severity == "Error");
        Assert.Contains(results, item => item.Code == "canonical_duplicate_fact" && item.Severity == "Error");
    }

    [Fact]
    public void Parity_ClassifiesExactRoundingAndMissingValues()
    {
        var legacy = BuildLegacyResult();
        legacy.Summaries[0].IncurredClaims = null;
        var envelope = new ManulifeClaimsCanonicalMapper().Map(legacy, 1, 2, 3);
        envelope.FinancialFacts[0] = CopyFinancial(envelope.FinancialFacts[0], paidClaims: 6500.005m);

        var results = new CanonicalClaimsParityService().Compare(legacy, envelope, 0.01m);

        Assert.Contains(results, item => item.Metric == "Premium" && item.Classification == CanonicalClaimsParityClassification.ExactMatch);
        Assert.Contains(results, item => item.Metric == "PaidClaims" && item.Classification == CanonicalClaimsParityClassification.WithinRoundingTolerance);
        Assert.Contains(results, item => item.Metric == "IncurredClaims" && item.Classification == CanonicalClaimsParityClassification.MissingLegacyField);
    }

    [Fact]
    public async Task VerticalSlice_PersistsEvidenceParityAndIdempotentRetry_WithoutChangingLegacyRows()
    {
        var folder = NewTempFolder();
        var databasePath = Path.Combine(folder, "benefits.db");
        var sourcePath = Path.Combine(folder, "claims.pdf");
        await File.WriteAllTextAsync(sourcePath, "stable fixture placeholder");
        try
        {
            var service = new DatabaseService(databasePath);
            var legacy = BuildLegacyResult();
            var first = await service.SaveClaimsExperienceAsync(legacy, [sourcePath], folder);
            var retry = await service.SaveClaimsExperienceAsync(legacy, [sourcePath], folder);

            Assert.True(first.Succeeded, first.Message);
            Assert.True(retry.Succeeded, retry.Message);
            Assert.Contains("Canonical report", first.Message);
            var readModel = await new CanonicalClaimsQueryService(databasePath).GetLatestForClientAsync(first.ClientId!.Value);
            Assert.NotNull(readModel);
            Assert.Equal(2, readModel.FinancialFacts.Count);
            Assert.Equal(2, readModel.MonthlyFacts.Count);
            Assert.True(readModel.RequiredParityComplete);
            var historyReports = await new CanonicalClaimsQueryService(databasePath).GetAllForClientAsync(first.ClientId.Value);
            var historyReport = Assert.Single(historyReports);
            Assert.Equal(readModel.CanonicalReportId, historyReport.CanonicalReportId);
            Assert.NotNull(historyReport.SourceDocumentId);
            Assert.NotEqual(default, historyReport.UpdatedAt);
            await using var connection = await OpenAsync(databasePath);
            Assert.Equal(2, await CountAsync(connection, "ClaimsExperience"));
            Assert.Equal(2, await CountAsync(connection, "MonthlyClaims"));
            Assert.Equal(1, await CountAsync(connection, "CanonicalClaimsReport"));
            Assert.Equal(2, await CountAsync(connection, "CanonicalClaimsFinancialFact"));
            Assert.Equal(2, await CountAsync(connection, "CanonicalClaimsMonthlyFact"));
            Assert.Equal(4, await CountAsync(connection, "CanonicalClaimsEvidence"));
            Assert.Equal(12, await CountAsync(connection, "CanonicalClaimsParityResult"));
            Assert.Equal(12, await ScalarAsync(connection, "SELECT COUNT(*) FROM CanonicalClaimsParityResult WHERE classification='ExactMatch';"));
            Assert.Equal(2, await ScalarAsync(connection, "SELECT COUNT(*) FROM ImportLog WHERE step_name='Canonical Claims Import';"));
        }
        finally
        {
            SqliteConnection.ClearAllPools();
            Directory.Delete(folder, true);
        }
    }

    [Fact]
    public async Task Persistence_RollsBackAllCanonicalChangesOnDatabaseFailure()
    {
        var folder = NewTempFolder();
        var databasePath = Path.Combine(folder, "benefits.db");
        var sourcePath = Path.Combine(folder, "claims.pdf");
        await File.WriteAllTextAsync(sourcePath, "stable fixture placeholder");
        try
        {
            var service = new DatabaseService(databasePath);
            var saved = await service.SaveClaimsExperienceAsync(BuildLegacyResult(), [sourcePath], folder);
            Assert.True(saved.Succeeded, saved.Message);

            var envelope = new ManulifeClaimsCanonicalMapper().Map(BuildLegacyResult(), saved.ClientId!.Value, saved.DocumentId, null);
            envelope.FinancialFacts.Add(CopyFinancial(envelope.FinancialFacts[0])); // duplicate fact_id forces a UNIQUE failure after deletion begins
            var failed = await service.PersistCanonicalClaimsAsync(envelope, [], []);

            Assert.False(failed.Succeeded);
            await using var connection = await OpenAsync(databasePath);
            Assert.Equal(2, await CountAsync(connection, "CanonicalClaimsFinancialFact"));
            Assert.Equal(2, await CountAsync(connection, "CanonicalClaimsMonthlyFact"));
        }
        finally
        {
            SqliteConnection.ClearAllPools();
            Directory.Delete(folder, true);
        }
    }

    [Fact]
    public async Task OldEmptySqliteDatabase_UpgradesWithoutDroppingLegacyTables()
    {
        var folder = NewTempFolder();
        var databasePath = Path.Combine(folder, "old.db");
        try
        {
            await using (var connection = await OpenAsync(databasePath)) { }
            var initialized = new DatabaseService(databasePath).Initialize();
            Assert.True(initialized.Succeeded, initialized.Message);
            var initializedAgain = new DatabaseService(databasePath).Initialize();
            Assert.True(initializedAgain.Succeeded, initializedAgain.Message);
            await using var verify = await OpenAsync(databasePath);
            Assert.Equal(1, await ScalarAsync(verify, "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='ClaimsExperience';"));
            Assert.Equal(1, await ScalarAsync(verify, "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='CanonicalClaimsReport';"));
            Assert.Equal(1, await ScalarAsync(verify, "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='CanonicalClaimsFactRevision';"));
            Assert.Equal(1, await ScalarAsync(verify, "SELECT COUNT(*) FROM pragma_table_info('CanonicalClaimsEvidence') WHERE name='mapping_classification';"));
            Assert.Equal(1, await ScalarAsync(verify, "SELECT COUNT(*) FROM pragma_table_info('CanonicalClaimsCategoryFact') WHERE name='measurement_type';"));
            Assert.Equal(1, await ScalarAsync(verify, "SELECT COUNT(*) FROM pragma_table_info('CanonicalClaimsDistributionFact') WHERE name='division';"));
            Assert.Equal(1, await ScalarAsync(verify, "SELECT COUNT(*) FROM pragma_table_info('CanonicalClaimsBenchmarkFact') WHERE name='unit_of_measure';"));
        }
        finally
        {
            SqliteConnection.ClearAllPools();
            Directory.Delete(folder, true);
        }
    }

    [Fact]
    public void Architecture_UIParsersAndSharedEnginesCannotBypassCanonicalBoundary()
    {
        var root = FindRepositoryRoot();
        var uiFiles = Directory.GetFiles(Path.Combine(root, "src", "RenewalReviewBuilder.App"), "*.xaml.cs", SearchOption.AllDirectories)
            .Append(Path.Combine(root, "src", "RenewalReviewBuilder.App", "MainWindow.xaml.cs"));
        foreach (var file in uiFiles.Distinct(StringComparer.OrdinalIgnoreCase))
        {
            var text = File.ReadAllText(file);
            Assert.DoesNotContain("CanonicalClaimsImportService", text, StringComparison.Ordinal);
            Assert.DoesNotContain("INSERT INTO CanonicalClaims", text, StringComparison.OrdinalIgnoreCase);
        }

        foreach (var fileName in new[] { "UnderwritingEngine.cs", "RenewalEngine.cs" })
        {
            var file = Directory.GetFiles(Path.Combine(root, "src"), fileName, SearchOption.AllDirectories).Single();
            var text = File.ReadAllText(file);
            Assert.DoesNotContain("CanonicalClaims", text, StringComparison.Ordinal);
            Assert.DoesNotContain("carrier ==", text, StringComparison.OrdinalIgnoreCase);
        }
    }

    private static ClaimsExperienceDataModelResult BuildLegacyResult()
    {
        var result = new ClaimsExperienceDataModelResult
        {
            ClientName = "Canonical Claims Test Client",
            ClientNumber = "MFC-12345",
            Carrier = "Manulife",
            ExperiencePeriodStart = new DateOnly(2025, 1, 1),
            ExperiencePeriodEnd = new DateOnly(2025, 12, 31)
        };
        result.Summaries.Add(new ClaimsExperienceSummary { BenefitLine = "Extended Health Care", Premium = 10000m, PaidClaims = 6500m, IncurredClaims = 6500m, LossRatio = 65m, Lives = 10 });
        result.Summaries.Add(new ClaimsExperienceSummary { BenefitLine = "Dental Care", Premium = 8000m, PaidClaims = 4000m, IncurredClaims = 4000m, LossRatio = 50m, Lives = 10 });
        result.MonthlyClaims.Add(new MonthlyClaimsRecord { BenefitLine = "Extended Health Care", Month = new DateOnly(2025, 1, 1), Premium = 1000m, PaidClaims = 650m, IncurredClaims = 650m, LossRatio = 65m, Lives = 10 });
        result.MonthlyClaims.Add(new MonthlyClaimsRecord { BenefitLine = "Dental Care", Month = new DateOnly(2025, 1, 1), Premium = 800m, PaidClaims = 400m, IncurredClaims = 400m, LossRatio = 50m, Lives = 10 });
        return result;
    }

    private static CanonicalClaimFinancialFact CopyFinancial(CanonicalClaimFinancialFact source, decimal? paidClaims = null) => new()
    {
        FactId = source.FactId, Benefit = source.Benefit, Premium = source.Premium,
        PaidClaims = paidClaims ?? source.PaidClaims, IncurredClaims = source.IncurredClaims,
        LossRatio = source.LossRatio, Lives = source.Lives, Availability = source.Availability, Evidence = source.Evidence
    };

    private static string NewTempFolder()
    {
        var folder = Path.Combine(Path.GetTempPath(), "CanonicalClaimsTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        return folder;
    }

    private static async Task<SqliteConnection> OpenAsync(string path)
    {
        var connection = new SqliteConnection($"Data Source={path}");
        await connection.OpenAsync();
        return connection;
    }

    private static async Task<long> CountAsync(SqliteConnection connection, string table) => await ScalarAsync(connection, $"SELECT COUNT(*) FROM {table};");
    private static async Task<long> ScalarAsync(SqliteConnection connection, string sql)
    {
        await using var command = connection.CreateCommand();
        command.CommandText = sql;
        return Convert.ToInt64(await command.ExecuteScalarAsync(), CultureInfo.InvariantCulture);
    }

    private static string FindRepositoryRoot()
    {
        var folder = new DirectoryInfo(AppContext.BaseDirectory);
        while (folder is not null && !File.Exists(Path.Combine(folder.FullName, "RenewalReviewBuilder.sln"))) folder = folder.Parent;
        return folder?.FullName ?? throw new DirectoryNotFoundException("Repository root was not found.");
    }
}

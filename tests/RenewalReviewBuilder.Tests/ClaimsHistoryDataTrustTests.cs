using System.Collections.Generic;
using System.Globalization;
using Microsoft.Data.Sqlite;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Models.Querying;
using RenewalReviewBuilder.App.Services;
using RenewalReviewBuilder.App.Services.CanonicalClaims;
using RenewalReviewBuilder.App.Services.ClientHistory;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class ClaimsHistoryDataTrustTests
{
    [Fact]
    public async Task Persisted_NeedsReview_Facts_Remain_ReviewRequired_And_Do_Not_Qualify_A_History_Milestone()
    {
        var folder = NewTempFolder();
        var databasePath = Path.Combine(folder, "claims-history.db");
        try
        {
            var database = new DatabaseService(databasePath);
            var initialized = database.Initialize();
            Assert.True(initialized.Succeeded, initialized.Message);
            var clientId = await InsertClientAsync(databasePath);
            var envelope = BuildEnvelope(clientId);

            var saved = await database.PersistCanonicalClaimsAsync(envelope, [], []);
            Assert.True(saved.Succeeded, saved.Message);

            // Unknown legacy values are intentionally fail-closed instead of becoming Valid.
            await ExecuteAsync(databasePath,
                "UPDATE CanonicalClaimsBenchmarkFact SET validation_status='LegacyUnknown' WHERE fact_id='benchmark-review';");

            var report = Assert.Single(await new CanonicalClaimsQueryService(databasePath).GetAllForClientAsync(clientId));
            Assert.Equal(CanonicalClaimsValidationStatus.Valid.ToString(), report.ValidationStatus);
            Assert.Equal(CanonicalClaimsValidationStatus.NeedsReview, Assert.Single(report.FinancialFacts).ValidationStatus);
            Assert.Equal(CanonicalClaimsValidationStatus.NeedsReview, Assert.Single(report.MonthlyFacts).ValidationStatus);
            Assert.Equal(CanonicalClaimsValidationStatus.NeedsReview, Assert.Single(report.CategoryFacts).ValidationStatus);
            Assert.Equal(CanonicalClaimsValidationStatus.NeedsReview, Assert.Single(report.DistributionFacts).ValidationStatus);
            Assert.Equal(CanonicalClaimsValidationStatus.NeedsReview, Assert.Single(report.BenchmarkFacts).ValidationStatus);

            var history = await new ClaimsHistoryContributor(database).ContributeAsync(clientId);

            Assert.Empty(history);
        }
        finally
        {
            SqliteConnection.ClearAllPools();
            if (Directory.Exists(folder)) Directory.Delete(folder, recursive: true);
        }
    }

    [Fact]
    public async Task Contributor_Emits_Only_MultiMonth_NonMonthly_Report_Milestones()
    {
        const long clientId = 41;
        var monthlyType = Report(clientId, 1, new DateOnly(2024, 1, 1), new DateOnly(2024, 12, 31), "Monthly");
        var oneMonth = Report(clientId, 2, new DateOnly(2024, 1, 1), new DateOnly(2024, 2, 1), "Experience");
        var multiMonth = Report(clientId, 3, new DateOnly(2024, 1, 1), new DateOnly(2024, 2, 2), "Experience");
        var contributor = Contributor([monthlyType, oneMonth, multiMonth]);

        var events = await contributor.ContributeAsync(clientId);

        var milestone = Assert.Single(events);
        Assert.Contains("2024-02-02", milestone.SourceEntityId, StringComparison.Ordinal);
        Assert.Contains("CanonicalClaimsReport:3", milestone.RelatedRecords);
        Assert.DoesNotContain("CanonicalClaimsReport:1", milestone.RelatedRecords);
        Assert.DoesNotContain("CanonicalClaimsReport:2", milestone.RelatedRecords);
    }

    [Fact]
    public async Task Replay_Excludes_Backdated_Report_Recorded_After_AsOf_And_Identity_Is_ClientScoped()
    {
        const long clientId = 41;
        var report = Report(
            clientId,
            7,
            new DateOnly(2024, 1, 1),
            new DateOnly(2024, 12, 31),
            "Experience",
            updatedAt: new DateTime(2025, 2, 10, 9, 0, 0, DateTimeKind.Utc));
        var contributor = Contributor([report]);

        var beforeImport = await contributor.ContributeAsync(new AsOfQueryContext
        {
            ClientId = clientId,
            AsOfDate = new DateOnly(2025, 1, 31)
        });
        var afterImport = await contributor.ContributeAsync(new AsOfQueryContext
        {
            ClientId = clientId,
            AsOfDate = new DateOnly(2025, 2, 10)
        });
        var otherClientReport = Report(
            42,
            7,
            new DateOnly(2024, 1, 1),
            new DateOnly(2024, 12, 31),
            "Experience",
            updatedAt: report.UpdatedAt);
        var otherClientEvent = Assert.Single(await Contributor([otherClientReport]).ContributeAsync(42));

        Assert.Empty(beforeImport.Events);
        var replayEvent = Assert.Single(afterImport.Events);
        Assert.Equal(HistoricalQuerySupport.Limited, afterImport.HistoricalSupport.Support);
        Assert.Equal("CanonicalClaimsReportPeriod", replayEvent.SourceEntityType);
        Assert.Equal(
            "41|MANULIFE|P-1|2024-01-01|2024-12-31|EXPERIENCE",
            replayEvent.SourceEntityId);
        Assert.Equal(
            $"Claims:CanonicalClaimsReportPeriod:{replayEvent.SourceEntityId}:experience-period",
            replayEvent.StableSourceKey);
        Assert.NotEqual(replayEvent.StableSourceKey, otherClientEvent.StableSourceKey);
    }

    private static CanonicalClaimsImportEnvelope BuildEnvelope(long clientId)
    {
        var envelope = new CanonicalClaimsImportEnvelope
        {
            Report = new CanonicalClaimsReport
            {
                ClientId = clientId,
                Carrier = "Manulife",
                PolicyNumber = "P-1",
                Period = new CanonicalClaimsPeriod
                {
                    Start = new DateOnly(2024, 1, 1),
                    End = new DateOnly(2024, 12, 31),
                    PeriodType = "Experience"
                },
                CreatedAt = new DateTime(2025, 1, 2, 0, 0, 0, DateTimeKind.Utc),
                UpdatedAt = new DateTime(2025, 1, 2, 0, 0, 0, DateTimeKind.Utc)
            }
        };
        envelope.FinancialFacts.Add(new CanonicalClaimFinancialFact
        {
            FactId = "financial-review",
            Benefit = "EHC",
            PaidClaims = 100m,
            Availability = CanonicalDataAvailability.Present,
            ValidationStatus = CanonicalClaimsValidationStatus.NeedsReview
        });
        envelope.MonthlyFacts.Add(new CanonicalClaimMonthlyFact
        {
            FactId = "monthly-review",
            Benefit = "EHC",
            Month = new DateOnly(2024, 1, 1),
            PaidClaims = 10m,
            Availability = CanonicalDataAvailability.Present,
            ValidationStatus = CanonicalClaimsValidationStatus.NeedsReview
        });
        envelope.CategoryFacts.Add(new CanonicalClaimCategoryFact
        {
            FactId = "category-review",
            Benefit = "EHC",
            Category = "Drugs",
            MonetaryValue = 25m,
            Availability = CanonicalDataAvailability.Present,
            ValidationStatus = CanonicalClaimsValidationStatus.NeedsReview
        });
        envelope.DistributionFacts.Add(new CanonicalClaimDistributionFact
        {
            FactId = "distribution-review",
            Benefit = "EHC",
            Category = "Claim size",
            CountValue = 2,
            Availability = CanonicalDataAvailability.Present,
            ValidationStatus = CanonicalClaimsValidationStatus.NeedsReview
        });
        envelope.BenchmarkFacts.Add(new CanonicalBenchmarkFact
        {
            FactId = "benchmark-review",
            Benefit = "EHC",
            BenchmarkDefinition = "Carrier-provided benchmark",
            PercentageValue = 65m,
            Availability = CanonicalDataAvailability.Present,
            ValidationStatus = CanonicalClaimsValidationStatus.NeedsReview
        });
        return envelope;
    }

    private static ClaimsHistoryContributor Contributor(IReadOnlyList<CanonicalClaimsReportReadModel> reports) =>
        new((_, _) => Task.FromResult(reports));

    private static CanonicalClaimsReportReadModel Report(
        long clientId,
        long reportId,
        DateOnly start,
        DateOnly end,
        string periodType,
        DateTime? updatedAt = null)
    {
        var recordedAt = updatedAt ?? new DateTime(2025, 1, 2, 0, 0, 0, DateTimeKind.Utc);
        var report = new CanonicalClaimsReportReadModel
        {
            CanonicalReportId = reportId,
            ClientId = clientId,
            Carrier = "Manulife",
            PolicyNumber = "P-1",
            SourceDocumentId = 700 + reportId,
            SourceDocumentVersionId = 1700 + reportId,
            Period = new CanonicalClaimsPeriod { Start = start, End = end, PeriodType = periodType },
            ValidationStatus = CanonicalClaimsValidationStatus.Valid.ToString(),
            CreatedAt = recordedAt.AddMinutes(-1),
            UpdatedAt = recordedAt
        };
        report.FinancialFacts.Add(new CanonicalClaimFinancialFact
        {
            FactId = $"financial-{clientId}-{reportId}",
            Benefit = "EHC",
            PaidClaims = 100m,
            Availability = CanonicalDataAvailability.Present,
            ValidationStatus = CanonicalClaimsValidationStatus.Valid
        });
        return report;
    }

    private static string NewTempFolder()
    {
        var folder = Path.Combine(Path.GetTempPath(), "ClaimsHistoryDataTrustTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        return folder;
    }

    private static async Task<long> InsertClientAsync(string databasePath)
    {
        await using var connection = new SqliteConnection($"Data Source={databasePath}");
        await connection.OpenAsync();
        await using var command = connection.CreateCommand();
        command.CommandText = @"
INSERT INTO Clients (client_name, client_number, carrier, status, date_created, date_updated)
VALUES ($name, $number, 'Manulife', 'Active', $now, $now);
SELECT last_insert_rowid();";
        command.Parameters.AddWithValue("$name", $"Claims History {Guid.NewGuid():N}");
        command.Parameters.AddWithValue("$number", Guid.NewGuid().ToString("N"));
        command.Parameters.AddWithValue("$now", DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture));
        return Convert.ToInt64(await command.ExecuteScalarAsync(), CultureInfo.InvariantCulture);
    }

    private static async Task ExecuteAsync(string databasePath, string sql)
    {
        await using var connection = new SqliteConnection($"Data Source={databasePath}");
        await connection.OpenAsync();
        await using var command = connection.CreateCommand();
        command.CommandText = sql;
        await command.ExecuteNonQueryAsync();
    }
}

using System.Globalization;
using Microsoft.Data.Sqlite;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services.CanonicalClaims;

namespace RenewalReviewBuilder.App.Services;

public interface IClientClaimsDashboardDataService
{
    Task<ClientClaimsDashboardDataResult> GetAsync(long clientId);
}

/// <summary>
/// Read-only compatibility boundary for the client Claims Experience dashboard.
/// Canonical facts are authoritative; legacy tables remain available as historical fallback.
/// </summary>
public sealed class ClientClaimsDashboardDataService : IClientClaimsDashboardDataService
{
    private static readonly string[] PeriodSeparators = [" - ", " – ", " — ", " to "];

    private readonly DatabaseService _database;
    private readonly CanonicalClaimsQueryService _canonicalClaims;

    public ClientClaimsDashboardDataService(DatabaseService database)
    {
        _database = database;
        _canonicalClaims = new CanonicalClaimsQueryService(database.DatabasePath);
    }

    public async Task<ClientClaimsDashboardDataResult> GetAsync(long clientId)
    {
        var legacyClaimsTask = _database.GetClaimsHistoryAsync(clientId);
        var legacyMonthlyTask = _database.GetMonthlyClaimsAsync(clientId);
        var canonicalReportsTask = _canonicalClaims.GetAllForClientAsync(clientId);
        await Task.WhenAll(legacyClaimsTask, legacyMonthlyTask, canonicalReportsTask);

        var reports = canonicalReportsTask.Result
            .Where(report => !report.ValidationStatus.Equals("Invalid", StringComparison.OrdinalIgnoreCase))
            .ToList();
        var canonicalClaims = MapCanonicalClaims(reports, clientId);
        var canonicalMonthly = MapCanonicalMonthly(reports, clientId);
        var legacyClaims = legacyClaimsTask.Result;
        var legacyMonthly = legacyMonthlyTask.Result;
        var claims = MergeClaims(canonicalClaims, legacyClaims);
        var monthly = MergeMonthly(canonicalMonthly, legacyMonthly);
        var usedCanonical = canonicalClaims.Count > 0 || canonicalMonthly.Count > 0;
        var usedLegacy = claims.Count > canonicalClaims.Count || monthly.Count > canonicalMonthly.Count;
        var needsReview = reports.Any(report =>
            report.ValidationStatus.Equals("NeedsReview", StringComparison.OrdinalIgnoreCase));

        if (claims.Count > 0 || monthly.Count > 0)
        {
            return Result(
                claims,
                monthly,
                usedCanonical
                    ? usedLegacy ? ClientClaimsDashboardSource.CanonicalWithLegacyHistory : ClientClaimsDashboardSource.Canonical
                    : ClientClaimsDashboardSource.Legacy,
                needsReview,
                usedCanonical
                    ? usedLegacy
                        ? "Source: canonical claims with legacy history fallback"
                        : "Source: canonical claims"
                    : "Source: SQLite ClaimsExperience and MonthlyClaims");
        }

        var distributionClaims = await ReadDistributionTotalsAsync(clientId);
        if (distributionClaims.Count > 0)
        {
            return Result(
                distributionClaims,
                [],
                ClientClaimsDashboardSource.DistributionCompatibility,
                needsReview: true,
                "Source: saved claims-distribution totals (premium and lives are not inferred)");
        }

        var renewalClaims = await ReadRenewalClaimsCompatibilityAsync(clientId);
        if (renewalClaims.Count > 0)
        {
            return Result(
                renewalClaims,
                [],
                ClientClaimsDashboardSource.RenewalCompatibility,
                needsReview: true,
                "Source: dated renewal claims values (premium and lives are not inferred)");
        }

        return Result(
            [],
            [],
            ClientClaimsDashboardSource.None,
            needsReview: false,
            "Source: renewal saved; no supported claims facts were found");
    }

    private static List<ClaimsExperienceRecord> MapCanonicalClaims(
        IReadOnlyCollection<CanonicalClaimsReportReadModel> reports,
        long clientId)
    {
        var rows = new List<ClaimsExperienceRecord>();
        var syntheticId = -1L;
        foreach (var report in reports.OrderBy(item => item.Period.End))
        {
            var financialFacts = report.FinancialFacts
                .Where(IsUsable)
                .ToList();
            foreach (var fact in financialFacts)
            {
                var period = ResolvePeriod(fact.DimensionType, fact.DimensionValue, report.Period);
                rows.Add(new ClaimsExperienceRecord
                {
                    ClaimsId = syntheticId--,
                    ClientId = clientId,
                    ExperiencePeriodStart = period.Start,
                    ExperiencePeriodEnd = period.End,
                    BenefitLine = DisplayBenefit(fact.Benefit),
                    Premium = fact.Premium,
                    PaidClaims = fact.PaidClaims,
                    IncurredClaims = fact.IncurredClaims,
                    LossRatio = fact.LossRatio,
                    Lives = fact.Lives,
                    SourceDocumentId = report.SourceDocumentId,
                    DateLoaded = report.UpdatedAt.ToLocalTime()
                });
            }

            if (financialFacts.Count > 0)
            {
                continue;
            }

            // Some canonical carrier reports contain explicit claims-distribution totals but no
            // financial facts. Preserve those paid totals without inventing premium or membership.
            foreach (var fact in report.DistributionFacts.Where(fact =>
                         fact.Category.Equals("Total", StringComparison.OrdinalIgnoreCase) &&
                         fact.MonetaryValue.HasValue &&
                         IsUsable(fact)))
            {
                var period = ResolvePeriod(fact.DimensionType, fact.DimensionValue, report.Period);
                rows.Add(new ClaimsExperienceRecord
                {
                    ClaimsId = syntheticId--,
                    ClientId = clientId,
                    ExperiencePeriodStart = period.Start,
                    ExperiencePeriodEnd = period.End,
                    BenefitLine = DisplayBenefit(fact.Benefit),
                    PaidClaims = fact.MonetaryValue,
                    SourceDocumentId = report.SourceDocumentId,
                    DateLoaded = report.UpdatedAt.ToLocalTime()
                });
            }
        }

        return rows;
    }

    private static List<MonthlyClaimsRecord> MapCanonicalMonthly(
        IReadOnlyCollection<CanonicalClaimsReportReadModel> reports,
        long clientId)
    {
        var rows = new List<MonthlyClaimsRecord>();
        var syntheticId = -1L;
        foreach (var report in reports.OrderBy(item => item.Period.End))
        {
            foreach (var fact in report.MonthlyFacts.Where(IsUsable))
            {
                rows.Add(new MonthlyClaimsRecord
                {
                    ClaimMonthId = syntheticId--,
                    ClientId = clientId,
                    Carrier = report.Carrier,
                    BenefitLine = DisplayBenefit(fact.Benefit),
                    Month = fact.Month,
                    Premium = fact.Premium,
                    PaidClaims = fact.PaidClaims,
                    IncurredClaims = fact.IncurredClaims,
                    ClaimsPlusFees = fact.ClaimsPlusFees,
                    LossRatio = fact.LossRatio,
                    Lives = fact.Lives,
                    SourceDocumentId = report.SourceDocumentId,
                    DateLoaded = report.UpdatedAt.ToLocalTime()
                });
            }
        }

        return rows;
    }

    private async Task<List<ClaimsExperienceRecord>> ReadDistributionTotalsAsync(long clientId)
    {
        await using var connection = OpenReadOnlyConnection();
        await connection.OpenAsync();
        await using var command = connection.CreateCommand();
        command.CommandText = @"
SELECT claims_distribution_report_id, period_start, period_end, benefit_type,
       total_amount_paid, source_document_id, date_loaded, amounts_balance_to_total
FROM claims_distribution_reports
WHERE client_id = $client_id
  AND total_amount_paid IS NOT NULL
  AND period_start IS NOT NULL
  AND period_end IS NOT NULL
ORDER BY date_loaded DESC, claims_distribution_report_id DESC;";
        command.Parameters.AddWithValue("$client_id", clientId);

        var candidates = new List<DistributionTotalRow>();
        await using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            var start = ReadDate(reader, 1);
            var end = ReadDate(reader, 2);
            if (!start.HasValue || !end.HasValue)
            {
                continue;
            }

            candidates.Add(new DistributionTotalRow(
                reader.GetInt64(0),
                start.Value,
                end.Value,
                reader.IsDBNull(3) ? "Overall" : reader.GetString(3),
                Convert.ToDecimal(reader.GetValue(4), CultureInfo.InvariantCulture),
                reader.IsDBNull(5) ? null : reader.GetInt64(5),
                ReadDateTime(reader, 6),
                !reader.IsDBNull(7) && reader.GetInt64(7) == 1));
        }

        return candidates
            .GroupBy(row => ClaimsKey(row.Benefit, row.Start, row.End))
            .Select(group => group
                .OrderByDescending(row => row.BalancesToTotal)
                .ThenByDescending(row => row.DateLoaded)
                .ThenByDescending(row => row.ReportId)
                .First())
            .Select((row, index) => new ClaimsExperienceRecord
            {
                ClaimsId = -(row.ReportId + index + 1),
                ClientId = clientId,
                ExperiencePeriodStart = row.Start,
                ExperiencePeriodEnd = row.End,
                BenefitLine = DisplayBenefit(row.Benefit),
                PaidClaims = row.PaidClaims,
                SourceDocumentId = row.SourceDocumentId,
                DateLoaded = row.DateLoaded
            })
            .ToList();
    }

    private async Task<List<ClaimsExperienceRecord>> ReadRenewalClaimsCompatibilityAsync(long clientId)
    {
        await using var connection = OpenReadOnlyConnection();
        await connection.OpenAsync();
        await using var command = connection.CreateCommand();
        command.CommandText = @"
SELECT r.renewal_id, r.benefit_line, r.claims, r.source_document_id,
       d.period_start, d.period_end, r.date_loaded
FROM Renewals r
JOIN Documents d ON d.document_id = r.source_document_id
WHERE r.client_id = $client_id
  AND r.claims > 0
  AND d.period_start IS NOT NULL
  AND d.period_end IS NOT NULL
ORDER BY r.date_loaded DESC, r.renewal_id DESC;";
        command.Parameters.AddWithValue("$client_id", clientId);

        var candidates = new List<RenewalClaimRow>();
        await using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            var start = ReadDate(reader, 4);
            var end = ReadDate(reader, 5);
            if (!start.HasValue || !end.HasValue)
            {
                continue;
            }

            candidates.Add(new RenewalClaimRow(
                reader.GetInt64(0),
                reader.GetString(1),
                Convert.ToDecimal(reader.GetValue(2), CultureInfo.InvariantCulture),
                reader.GetInt64(3),
                start.Value,
                end.Value,
                ReadDateTime(reader, 6)));
        }

        var unambiguousDocumentRows = candidates
            .GroupBy(row => new { row.SourceDocumentId, Key = ClaimsKey(row.Benefit, row.Start, row.End) })
            .Select(group => new
            {
                Rows = group.ToList(),
                Amounts = group.Select(row => decimal.Round(row.PaidClaims, 2)).Distinct().ToList()
            })
            .Where(group => group.Amounts.Count == 1)
            .Select(group => group.Rows.OrderByDescending(row => row.DateLoaded).ThenByDescending(row => row.RenewalId).First())
            .ToList();

        return unambiguousDocumentRows
            .GroupBy(row => ClaimsKey(row.Benefit, row.Start, row.End))
            .Select(group => group.OrderByDescending(row => row.DateLoaded).ThenByDescending(row => row.SourceDocumentId).First())
            .Select(row => new ClaimsExperienceRecord
            {
                ClaimsId = -row.RenewalId,
                ClientId = clientId,
                ExperiencePeriodStart = row.Start,
                ExperiencePeriodEnd = row.End,
                BenefitLine = DisplayBenefit(row.Benefit),
                PaidClaims = row.PaidClaims,
                SourceDocumentId = row.SourceDocumentId,
                DateLoaded = row.DateLoaded
            })
            .ToList();
    }

    private SqliteConnection OpenReadOnlyConnection()
    {
        return new SqliteConnection(new SqliteConnectionStringBuilder
        {
            DataSource = _database.DatabasePath,
            Mode = SqliteOpenMode.ReadOnly
        }.ToString());
    }

    private static ClientClaimsDashboardDataResult Result(
        IReadOnlyList<ClaimsExperienceRecord> claims,
        IReadOnlyList<MonthlyClaimsRecord> monthly,
        ClientClaimsDashboardSource source,
        bool needsReview,
        string sourceText)
    {
        var count = claims.Count + monthly.Count;
        return new ClientClaimsDashboardDataResult
        {
            Claims = claims,
            MonthlyClaims = monthly,
            Source = source,
            NeedsReview = needsReview,
            SourceText = sourceText,
            DataStatusText = needsReview
                ? "Data status: source-backed values loaded; review incomplete fields"
                : count > 0 ? "Data accuracy: High" : "Data status: claims not available in saved sources",
            StatusMessage = count > 0
                ? $"Loaded {count:#,0} claims row(s) from {SourceLabel(source)}."
                : "The renewal is saved, but no supported claims facts were found in its saved sources."
        };
    }

    private static List<ClaimsExperienceRecord> MergeClaims(
        IReadOnlyCollection<ClaimsExperienceRecord> canonical,
        IReadOnlyCollection<ClaimsExperienceRecord> legacy)
    {
        var rows = canonical.ToList();
        var keys = rows.Select(row => ClaimsKey(row.BenefitLine, row.ExperiencePeriodStart, row.ExperiencePeriodEnd))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        rows.AddRange(legacy.Where(row => keys.Add(ClaimsKey(row.BenefitLine, row.ExperiencePeriodStart, row.ExperiencePeriodEnd))));
        return rows;
    }

    private static List<MonthlyClaimsRecord> MergeMonthly(
        IReadOnlyCollection<MonthlyClaimsRecord> canonical,
        IReadOnlyCollection<MonthlyClaimsRecord> legacy)
    {
        var rows = canonical.ToList();
        var keys = rows.Select(MonthlyKey).ToHashSet(StringComparer.OrdinalIgnoreCase);
        rows.AddRange(legacy.Where(row => keys.Add(MonthlyKey(row))));
        return rows;
    }

    private static bool IsUsable(CanonicalClaimFact fact)
    {
        return fact.ValidationStatus != CanonicalClaimsValidationStatus.Invalid &&
               fact.Availability is CanonicalDataAvailability.Present or CanonicalDataAvailability.ReportedZero;
    }

    private static (DateOnly? Start, DateOnly? End) ResolvePeriod(
        string dimensionType,
        string dimensionValue,
        CanonicalClaimsPeriod reportPeriod)
    {
        if (dimensionType.Equals("ReportingPeriod", StringComparison.OrdinalIgnoreCase) &&
            TryParsePeriod(dimensionValue, out var start, out var end))
        {
            return (start, end);
        }

        return (reportPeriod.Start, reportPeriod.End);
    }

    internal static bool TryParsePeriod(string value, out DateOnly start, out DateOnly end)
    {
        foreach (var separator in PeriodSeparators)
        {
            var index = value.IndexOf(separator, StringComparison.OrdinalIgnoreCase);
            if (index <= 0)
            {
                continue;
            }

            var startText = value[..index].Trim();
            var endText = value[(index + separator.Length)..].Trim();
            if (TryParsePeriodEndpoint(startText, isEnd: false, out start) &&
                TryParsePeriodEndpoint(endText, isEnd: true, out end) &&
                start <= end)
            {
                return true;
            }
        }

        start = default;
        end = default;
        return false;
    }

    private static bool TryParsePeriodEndpoint(string value, bool isEnd, out DateOnly date)
    {
        foreach (var format in new[] { "MMM yyyy", "MMMM yyyy" })
        {
            if (!DateTime.TryParseExact(value, format, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out var month))
            {
                continue;
            }

            date = new DateOnly(month.Year, month.Month, isEnd ? DateTime.DaysInMonth(month.Year, month.Month) : 1);
            return true;
        }

        if (DateOnly.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out date))
        {
            return true;
        }

        return false;
    }

    private static DateOnly? ReadDate(SqliteDataReader reader, int ordinal)
    {
        return reader.IsDBNull(ordinal) || string.IsNullOrWhiteSpace(reader.GetString(ordinal))
            ? null
            : DateOnly.Parse(reader.GetString(ordinal), CultureInfo.InvariantCulture);
    }

    private static DateTime ReadDateTime(SqliteDataReader reader, int ordinal)
    {
        return reader.IsDBNull(ordinal) || string.IsNullOrWhiteSpace(reader.GetString(ordinal))
            ? DateTime.MinValue
            : DateTime.Parse(reader.GetString(ordinal), CultureInfo.InvariantCulture);
    }

    private static string ClaimsKey(string benefit, DateOnly? start, DateOnly? end) =>
        $"{NormalizeKey(benefit)}|{start:yyyy-MM-dd}|{end:yyyy-MM-dd}";

    private static string MonthlyKey(MonthlyClaimsRecord row) =>
        $"{NormalizeKey(row.BenefitLine)}|{row.Month:yyyy-MM-dd}";

    private static string NormalizeKey(string value) =>
        string.Concat(value.Where(char.IsLetterOrDigit)).ToUpperInvariant();

    private static string DisplayBenefit(string value) =>
        string.IsNullOrWhiteSpace(value) ? "Overall" : value.Trim();

    private static string SourceLabel(ClientClaimsDashboardSource source) => source switch
    {
        ClientClaimsDashboardSource.Canonical => "canonical claims",
        ClientClaimsDashboardSource.CanonicalWithLegacyHistory => "canonical claims and legacy history",
        ClientClaimsDashboardSource.Legacy => "legacy claims history",
        ClientClaimsDashboardSource.DistributionCompatibility => "saved claims-distribution totals",
        ClientClaimsDashboardSource.RenewalCompatibility => "dated renewal claims values",
        _ => "saved sources"
    };

    private sealed record DistributionTotalRow(
        long ReportId,
        DateOnly Start,
        DateOnly End,
        string Benefit,
        decimal PaidClaims,
        long? SourceDocumentId,
        DateTime DateLoaded,
        bool BalancesToTotal);

    private sealed record RenewalClaimRow(
        long RenewalId,
        string Benefit,
        decimal PaidClaims,
        long SourceDocumentId,
        DateOnly Start,
        DateOnly End,
        DateTime DateLoaded);
}

public enum ClientClaimsDashboardSource
{
    None,
    Canonical,
    CanonicalWithLegacyHistory,
    Legacy,
    DistributionCompatibility,
    RenewalCompatibility
}

public sealed class ClientClaimsDashboardDataResult
{
    public IReadOnlyList<ClaimsExperienceRecord> Claims { get; init; } = [];
    public IReadOnlyList<MonthlyClaimsRecord> MonthlyClaims { get; init; } = [];
    public ClientClaimsDashboardSource Source { get; init; }
    public bool NeedsReview { get; init; }
    public string SourceText { get; init; } = "Source: no claims source loaded";
    public string DataStatusText { get; init; } = "Data status: claims not available";
    public string StatusMessage { get; init; } = "No supported claims facts were found.";
}

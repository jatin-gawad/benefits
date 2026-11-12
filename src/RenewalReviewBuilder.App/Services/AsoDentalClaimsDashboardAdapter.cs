using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

/// <summary>
/// Projects verified Dental ASO operating-account claims into the read-only
/// Claims dashboard without duplicating those facts in the legacy claims tables.
/// Premium and lives deliberately remain null because the operating-account
/// statement does not report either value.
/// </summary>
public static class AsoDentalClaimsDashboardAdapter
{
    public const string BenefitLine = "Dental ASO paid claims";

    public static AsoDentalClaimsDashboardData Merge(
        IReadOnlyCollection<ClaimsExperienceRecord> claims,
        IReadOnlyCollection<MonthlyClaimsRecord> monthlyClaims,
        AsoFinancialActivityStatement? statement)
    {
        var mergedClaims = claims.ToList();
        var mergedMonthlyClaims = monthlyClaims.ToList();
        if (!CanProject(statement))
        {
            return new AsoDentalClaimsDashboardData(mergedClaims, mergedMonthlyClaims, false);
        }

        var clientId = ResolveClientId(statement!, claims, monthlyClaims);
        if (clientId <= 0)
        {
            return new AsoDentalClaimsDashboardData(mergedClaims, mergedMonthlyClaims, false);
        }

        var loadedAt = ResolveLoadedAt(statement!);
        var added = false;
        if (!mergedClaims.Any(row =>
                IsDental(row.BenefitLine) &&
                row.ExperiencePeriodStart == statement!.PeriodStart &&
                row.ExperiencePeriodEnd == statement.PeriodEnd))
        {
            mergedClaims.Add(new ClaimsExperienceRecord
            {
                ClaimsId = -1,
                ClientId = clientId,
                ExperiencePeriodStart = statement!.PeriodStart,
                ExperiencePeriodEnd = statement.PeriodEnd,
                BenefitLine = BenefitLine,
                PaidClaims = statement.Total!.DentalClaimsPaid,
                Premium = null,
                Lives = null,
                LossRatio = null,
                SourceDocumentId = statement.SourceDocumentId,
                DateLoaded = loadedAt
            });
            added = true;
        }

        var existingDentalMonths = mergedMonthlyClaims
            .Where(row => IsDental(row.BenefitLine) && row.Month.HasValue)
            .Select(row => row.Month!.Value)
            .ToHashSet();
        var nextSyntheticId = -1L;
        foreach (var period in statement!.Periods
                     .Where(period => period.ActivityMonth.HasValue && period.DentalClaimsPaid.HasValue)
                     .OrderBy(period => period.ActivityMonth))
        {
            if (!existingDentalMonths.Add(period.ActivityMonth!.Value))
            {
                continue;
            }

            mergedMonthlyClaims.Add(new MonthlyClaimsRecord
            {
                ClaimMonthId = nextSyntheticId--,
                ClientId = clientId,
                Carrier = statement.Carrier,
                BenefitLine = BenefitLine,
                Month = period.ActivityMonth,
                PaidClaims = period.DentalClaimsPaid,
                Premium = null,
                Lives = null,
                LossRatio = null,
                SourceDocumentId = statement.SourceDocumentId,
                DateLoaded = loadedAt
            });
            added = true;
        }

        return new AsoDentalClaimsDashboardData(mergedClaims, mergedMonthlyClaims, added);
    }

    private static bool CanProject(AsoFinancialActivityStatement? statement)
    {
        return statement is not null &&
               statement.IsCurrent &&
               statement.ImportStatus.Equals(AsoFinancialImportStatuses.Complete, StringComparison.OrdinalIgnoreCase) &&
               statement.Benefits.Contains("dental", StringComparison.OrdinalIgnoreCase) &&
               statement.PeriodStart.HasValue &&
               statement.PeriodEnd.HasValue &&
               statement.Total?.DentalClaimsPaid is not null &&
               statement.Periods.Count > 0;
    }

    private static long ResolveClientId(
        AsoFinancialActivityStatement statement,
        IReadOnlyCollection<ClaimsExperienceRecord> claims,
        IReadOnlyCollection<MonthlyClaimsRecord> monthlyClaims)
    {
        if (statement.ClientId > 0)
        {
            return statement.ClientId;
        }

        return claims.Select(row => row.ClientId)
            .Concat(monthlyClaims.Select(row => row.ClientId))
            .FirstOrDefault(clientId => clientId > 0);
    }

    private static DateTime ResolveLoadedAt(AsoFinancialActivityStatement statement)
    {
        if (statement.ImportedAt > DateTime.MinValue)
        {
            return statement.ImportedAt;
        }

        return statement.StatementDate?.ToDateTime(TimeOnly.MinValue) ?? DateTime.MinValue;
    }

    private static bool IsDental(string benefitLine) =>
        benefitLine.Contains("dental", StringComparison.OrdinalIgnoreCase);
}

public sealed record AsoDentalClaimsDashboardData(
    IReadOnlyList<ClaimsExperienceRecord> Claims,
    IReadOnlyList<MonthlyClaimsRecord> MonthlyClaims,
    bool AddedAsoClaims);

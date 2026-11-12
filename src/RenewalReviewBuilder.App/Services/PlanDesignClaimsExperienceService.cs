using Microsoft.Data.Sqlite;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services.CanonicalClaims;

namespace RenewalReviewBuilder.App.Services;

public interface IPlanDesignClaimsExperienceService
{
    Task<PlanDesignClaimsExperienceResult> GetClaimsAsync(long clientId, PlanDesignClaimsReadMode mode = PlanDesignClaimsReadMode.CanonicalWithLegacyFallback);
}

/// <summary>Single read-only bridge from Plan Design to the report-centred canonical claims aggregate.</summary>
public sealed class PlanDesignClaimsExperienceService(string? databasePath = null) : IPlanDesignClaimsExperienceService
{
    private readonly string _databasePath = string.IsNullOrWhiteSpace(databasePath) ? AppPaths.BenefitsIntelligenceDatabasePath : databasePath;

    public async Task<PlanDesignClaimsExperienceResult> GetClaimsAsync(long clientId, PlanDesignClaimsReadMode mode = PlanDesignClaimsReadMode.CanonicalWithLegacyFallback)
    {
        if (mode != PlanDesignClaimsReadMode.Legacy)
        {
            var canonical = await ReadCanonicalAsync(clientId);
            if (canonical.Rows.Count > 0)
            {
                if (mode == PlanDesignClaimsReadMode.Compare) canonical.ParityResults.AddRange(await CompareLegacyAsync(clientId, canonical.Rows));
                return canonical;
            }
            if (mode == PlanDesignClaimsReadMode.CanonicalOnly) return new PlanDesignClaimsExperienceResult { ReadMode=mode, UsedCanonical=true, FallbackReason="No validated canonical claims report is available." };
            var legacy = await ReadLegacyAsync(clientId, mode);
            var fallback = new PlanDesignClaimsExperienceResult { ReadMode=mode, UsedCanonical=false, FallbackReason="No usable canonical claims report is available." };
            fallback.Rows.AddRange(legacy.Rows); return fallback;
        }
        return await ReadLegacyAsync(clientId, mode);
    }

    private async Task<PlanDesignClaimsExperienceResult> ReadCanonicalAsync(long clientId)
    {
        var report = await new CanonicalClaimsQueryService(_databasePath).GetLatestForClientAsync(clientId);
        if (report is null || report.ValidationStatus.Equals("Invalid", StringComparison.OrdinalIgnoreCase)) return new PlanDesignClaimsExperienceResult { ReadMode=PlanDesignClaimsReadMode.CanonicalWithLegacyFallback, UsedCanonical=true };
        var rows = new PlanDesignClaimsExperienceResult { ReadMode=PlanDesignClaimsReadMode.CanonicalWithLegacyFallback, UsedCanonical=true };
        foreach (var fact in report.UtilizationFacts)
            rows.Rows.Add(new PlanDesignClaimsExperienceReadModel { ClientId=clientId, CanonicalClaimsReportId=report.CanonicalReportId, BenefitCode=fact.BenefitCode, BenefitName=fact.BenefitName, CategoryCode=fact.CategoryCode, CategoryName=fact.CategoryName, ServiceCode=fact.ServiceCode, ServiceName=fact.ServiceName, ProvinceCode=fact.ProvinceCode, RelationshipType=fact.RelationshipType, DivisionCode=fact.DivisionCode, ClassCode=fact.ClassCode, PeriodStart=fact.MeasurementStartDate, PeriodEnd=fact.MeasurementEndDate, UniqueClaimants=fact.UniqueClaimants, Occurrences=fact.Occurrences, SubmittedAmount=fact.SubmittedAmount, EligibleAmount=fact.EligibleAmount, PaidAmount=fact.PaidAmount, ValidationStatus=report.ValidationStatus, EvidenceStatus=fact.Evidence.SourceDocumentId.HasValue ? "Available" : "Missing", SourceDocumentIds=fact.Evidence.SourceDocumentId.HasValue ? [fact.Evidence.SourceDocumentId.Value] : [], MappingConfidence=fact.Evidence.ConfidenceScore, MatchStatus=string.IsNullOrWhiteSpace(fact.ServiceCode) ? PlanDesignClaimsMatchStatuses.ServiceMappingRequired : PlanDesignClaimsMatchStatuses.ExactMatch, ConfidenceLevel=Confidence(report.ValidationStatus, fact) });
        return rows;
    }

    private async Task<PlanDesignClaimsExperienceResult> ReadLegacyAsync(long clientId, PlanDesignClaimsReadMode mode)
    {
        var result = new PlanDesignClaimsExperienceResult { ReadMode=mode, UsedCanonical=false };
        if (!File.Exists(_databasePath)) return result;
        await using var connection = new SqliteConnection(new SqliteConnectionStringBuilder { DataSource=_databasePath, Mode=SqliteOpenMode.ReadOnly }.ToString()); await connection.OpenAsync();
        await using var command=connection.CreateCommand(); command.CommandText="SELECT benefit_line, SUM(paid_claims) FROM ClaimsExperience WHERE client_id=$client GROUP BY benefit_line;"; command.Parameters.AddWithValue("$client",clientId);
        await using var reader=await command.ExecuteReaderAsync(); while(await reader.ReadAsync()) result.Rows.Add(new PlanDesignClaimsExperienceReadModel { ClientId=clientId, BenefitCode=Code(reader.GetString(0)), BenefitName=reader.GetString(0), PaidAmount=reader.IsDBNull(1)?null:Convert.ToDecimal(reader.GetValue(1)), ValidationStatus="Legacy", EvidenceStatus="Legacy", MatchStatus=PlanDesignClaimsMatchStatuses.BenefitLevelFallback, ConfidenceLevel="Limited", FallbackReason="Legacy claims read mode." });
        return result;
    }

    private async Task<List<CanonicalClaimsParityResult>> CompareLegacyAsync(long clientId, IReadOnlyCollection<PlanDesignClaimsExperienceReadModel> canonical)
    {
        var legacy=await ReadLegacyAsync(clientId, PlanDesignClaimsReadMode.Legacy); var results=new List<CanonicalClaimsParityResult>();
        foreach(var benefit in canonical.GroupBy(x=>x.BenefitCode)) { var c=benefit.Sum(x=>x.PaidAmount??0m); var l=legacy.Rows.Where(x=>x.BenefitCode.Equals(benefit.Key,StringComparison.OrdinalIgnoreCase)).Sum(x=>x.PaidAmount??0m); results.Add(new CanonicalClaimsParityResult { Metric="PlanDesignPaidClaims", Benefit=benefit.Key, LegacyValue=l, CanonicalValue=c, Variance=c-l, Classification=Math.Abs(c-l)<=.01m?CanonicalClaimsParityClassification.ExactMatch:CanonicalClaimsParityClassification.UnexpectedMismatch, Explanation="Plan Design legacy/canonical paid-claims comparison." }); }
        return results;
    }
    private static string Confidence(string validation, CanonicalClaimsUtilizationFact fact) => validation.Equals("Valid",StringComparison.OrdinalIgnoreCase) && fact.Evidence.SourceDocumentId.HasValue && fact.UniqueClaimants.HasValue ? "High" : fact.Evidence.SourceDocumentId.HasValue ? "Moderate" : "Limited";
    private static string Code(string value) => string.Concat(value.Where(char.IsLetterOrDigit)).ToUpperInvariant();
}

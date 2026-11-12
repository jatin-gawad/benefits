using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services.Normalization;

public static class NormalizeClaimsData
{
    public static NormalizedImportPayload FromClaimsExperienceResult(
        ClaimsExperienceDataModelResult result,
        NormalizedSourceMetadata source)
    {
        var payload = new NormalizedImportPayload
        {
            DocumentType = DocumentTypes.ClaimsExperience,
            Source = source.Copy(),
            Client = new NormalizedClient
            {
                ClientName = result.ClientName,
                ClientNumber = result.ClientNumber
            },
            Carrier = new NormalizedCarrier
            {
                CarrierName = result.Carrier
            }
        };

        foreach (var summary in result.Summaries)
        {
            payload.ClaimsExperience.Add(new NormalizedClaimsExperience
            {
                BenefitLine = summary.BenefitLine,
                PeriodStart = result.ExperiencePeriodStart,
                PeriodEnd = result.ExperiencePeriodEnd,
                Premium = summary.Premium,
                PaidClaims = summary.PaidClaims,
                IncurredClaims = summary.IncurredClaims,
                LossRatio = summary.LossRatio,
                Lives = summary.Lives
            });
        }

        foreach (var row in result.MonthlyClaims)
        {
            payload.Claims.Add(new NormalizedClaim
            {
                BenefitLine = row.BenefitLine,
                PeriodLabel = row.Month?.ToString("yyyy-MM") ?? "",
                PaidClaims = row.PaidClaims,
                IncurredClaims = row.IncurredClaims
            });
        }

        payload.ApplySourceToRecords();
        return payload;
    }

    public static NormalizedImportPayload FromWorkbook(WorkbookExtractionResult workbook, NormalizedSourceMetadata source)
    {
        var payload = new NormalizedImportPayload
        {
            DocumentType = DocumentTypes.ClaimsExperience,
            Source = source.Copy()
        };

        foreach (var trend in workbook.ClaimTrends)
        {
            payload.Claims.Add(new NormalizedClaim
            {
                BenefitLine = trend.Benefit,
                Category = trend.Category,
                PeriodLabel = trend.Period,
                PaidClaims = trend.Amount
            });
        }

        payload.ApplySourceToRecords();
        return payload;
    }
}

using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services.Normalization;

public static class NormalizeBillingData
{
    public static NormalizedImportPayload FromBilling(BillingExtractionResult result, NormalizedSourceMetadata source)
    {
        var payload = NormalizePremiumData.FromBilling(result, source);

        foreach (var detail in result.Details)
        {
            payload.BillingStatements.Add(new NormalizedBillingStatement
            {
                BillingPeriod = result.BillingPeriod,
                Province = detail.Province,
                BenefitLine = detail.Benefit,
                ClassCode = detail.ClassCode,
                OptionLabel = detail.OptionLabel,
                Lives = detail.Lives,
                MonthlyPremium = detail.CurrentMonthPremium,
                PremiumAdjustments = detail.PriorPeriodAdjustment,
                PremiumBeforeTax = detail.PremiumBeforeTax
            });
        }

        payload.ApplySourceToRecords();
        return payload;
    }
}

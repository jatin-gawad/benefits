using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services.Normalization;

public static class NormalizePremiumData
{
    public static NormalizedImportPayload FromBilling(BillingExtractionResult result, NormalizedSourceMetadata source)
    {
        var payload = new NormalizedImportPayload
        {
            DocumentType = DocumentTypes.Billing,
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

        foreach (var detail in result.Details)
        {
            payload.Premiums.Add(new NormalizedPremium
            {
                BenefitLine = detail.Benefit,
                ClassName = detail.ClassCode,
                PeriodLabel = result.BillingPeriod,
                // Normalized premium is the amount billed before tax, including prior-period adjustments.
                Amount = detail.PremiumBeforeTax
            });
        }

        payload.ApplySourceToRecords();
        return payload;
    }
}

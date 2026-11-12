using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services.Normalization;

public static class NormalizePlanDesignData
{
    public static NormalizedImportPayload FromRows(IEnumerable<PlanDesignRecord> rows, NormalizedSourceMetadata source)
    {
        var rowList = PlanDesignCanonicalMapper.NormalizeProvisionRows(rows);
        var first = rowList.FirstOrDefault();
        var payload = new NormalizedImportPayload
        {
            DocumentType = DocumentTypes.BenefitBooklet,
            Source = source.Copy(),
            Client = new NormalizedClient
            {
                ClientName = first?.ClientName ?? "",
                ClientNumber = first?.ClientNumber ?? ""
            },
            Carrier = new NormalizedCarrier
            {
                CarrierName = first?.Carrier ?? source.SourceCarrier
            }
        };

        foreach (var row in rowList)
        {
            payload.PlanDesign.Add(new NormalizedPlanDesign
            {
                BenefitCategory = row.BenefitCategory,
                BenefitName = row.BenefitName,
                ProvisionCode = row.ProvisionCode,
                ProvisionLabel = row.ProvisionLabel,
                CoverageDescription = row.CoverageDescription,
                ReimbursementPercent = row.ReimbursementPercent,
                ReimbursementText = row.ReimbursementText,
                MaximumAmount = row.MaximumAmount,
                Frequency = row.Frequency,
                WaitingPeriod = row.WaitingPeriod,
                TerminationAge = row.TerminationAge,
                ClassName = row.ClassName
            });
        }

        payload.ApplySourceToRecords();
        return payload;
    }
}

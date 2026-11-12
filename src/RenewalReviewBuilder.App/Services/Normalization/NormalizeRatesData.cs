using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services.Normalization;

public static class NormalizeRatesData
{
    public static NormalizedImportPayload FromAnalysis(AnalysisResult analysis, NormalizedSourceMetadata source)
    {
        var payload = new NormalizedImportPayload
        {
            DocumentType = DocumentTypes.Renewal,
            Source = source.Copy()
        };

        foreach (var line in analysis.Pdf.RateLines)
        {
            payload.Rates.Add(new NormalizedRate
            {
                BenefitLine = line.Benefit,
                ClassName = string.IsNullOrWhiteSpace(line.TierOrClass) ? line.AccountClass : line.TierOrClass,
                CurrentRate = line.CurrentRate,
                ProposedRate = line.RenewalRate,
                EffectiveDate = analysis.RenewalCase.RenewalDate ?? analysis.Pdf.RenewalDate
            });
        }

        payload.ApplySourceToRecords();
        return payload;
    }
}

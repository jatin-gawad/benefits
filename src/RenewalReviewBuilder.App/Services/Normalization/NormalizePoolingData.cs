using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services.Normalization;

public static class NormalizePoolingData
{
    public static NormalizedImportPayload FromAnalysis(AnalysisResult analysis, NormalizedSourceMetadata source)
    {
        var payload = new NormalizedImportPayload
        {
            DocumentType = DocumentTypes.Renewal,
            Source = source.Copy()
        };

        if (analysis.Pdf.PoolingCharge.HasValue || analysis.Pdf.PoolingThreshold.HasValue)
        {
            payload.Pooling.Add(new NormalizedPooling
            {
                BenefitLine = "All",
                PoolingCharge = analysis.Pdf.PoolingCharge,
                PoolingThreshold = analysis.Pdf.PoolingThreshold
            });
        }

        payload.ApplySourceToRecords();
        return payload;
    }
}

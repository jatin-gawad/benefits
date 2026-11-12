using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services.Normalization;

public static class NormalizeQuoteData
{
    public static NormalizedImportPayload Empty(NormalizedSourceMetadata source)
    {
        var payload = new NormalizedImportPayload
        {
            DocumentType = DocumentTypes.Proposal,
            Source = source.Copy()
        };
        payload.ApplySourceToRecords();
        return payload;
    }
}

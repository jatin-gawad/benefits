using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services.CanonicalClaims;
public interface ICarrierClaimsCanonicalMapper
{
    string CarrierName { get; }
    string SourceParserName { get; }
}

public interface ICarrierClaimsCanonicalMapper<in TParsedResult> : ICarrierClaimsCanonicalMapper
{
    CanonicalClaimsImportEnvelope Map(
        TParsedResult parsedResult,
        long clientId,
        long? sourceDocumentId,
        long? importBatchId);
}
public interface ICarrierClaimsCanonicalMultiPeriodMapper<in TParsedResult> : ICarrierClaimsCanonicalMapper<TParsedResult>
{
    IReadOnlyList<CanonicalClaimsImportEnvelope> MapMany(
        TParsedResult parsedResult,
        long clientId,
        long? sourceDocumentId,
        long? importBatchId);
}


using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services.Normalization;

namespace RenewalReviewBuilder.App.Connectors.Pdf;

public sealed class PdfPlanDesignConnector : IBenefitsConnector
{
    public string SourceType => SourceTypes.Pdf;
    public string Carrier => "Any";
    public string DocumentType => DocumentTypes.BenefitBooklet;
    public string ConnectorName => nameof(PdfPlanDesignConnector);
    public bool IsEnabled => true;
    public string Status => "Connected";

    public Task<NormalizedImportPayload> ExtractAsync(BenefitsConnectorRequest request)
    {
        var firstPath = request.AllInputPaths().FirstOrDefault() ?? request.InputPath;
        var source = NormalizedSourceMetadata.ForFile(SourceType, Carrier, firstPath, PdfExtractionMethods.ExternalBookletExtractor, null);
        return Task.FromResult(NormalizePlanDesignData.FromRows([], source));
    }
}

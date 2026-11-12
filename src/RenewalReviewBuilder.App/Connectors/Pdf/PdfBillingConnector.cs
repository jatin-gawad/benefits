using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;
using RenewalReviewBuilder.App.Services.Normalization;

namespace RenewalReviewBuilder.App.Connectors.Pdf;

public sealed class PdfBillingConnector : IBenefitsConnector
{
    private readonly ManulifeBillingStatementParser _parser;

    public PdfBillingConnector(ManulifeBillingStatementParser? parser = null)
    {
        _parser = parser ?? new ManulifeBillingStatementParser();
    }

    public string SourceType => SourceTypes.Pdf;
    public string Carrier => "Manulife";
    public string DocumentType => DocumentTypes.Billing;
    public string ConnectorName => nameof(PdfBillingConnector);
    public bool IsEnabled => true;
    public string Status => "Connected";

    public Task<NormalizedImportPayload> ExtractAsync(BenefitsConnectorRequest request)
    {
        var result = _parser.Extract(request.InputPath);
        var source = NormalizedSourceMetadata.ForFile(SourceType, result.Carrier, request.InputPath, PdfExtractionMethods.PdfPig, result.ImportStatus == BillingImportStatuses.Completed ? 0.92m : 0.75m);
        return Task.FromResult(NormalizeBillingData.FromBilling(result, source));
    }
}

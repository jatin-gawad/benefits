using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;
using RenewalReviewBuilder.App.Services.Normalization;

namespace RenewalReviewBuilder.App.Connectors.Pdf;

public class PdfRenewalConnector : IBenefitsConnector
{
    private readonly CarrierPdfParser _parser;

    public PdfRenewalConnector(CarrierPdfParser? parser = null)
    {
        _parser = parser ?? new CarrierPdfParser();
    }

    public virtual string SourceType => SourceTypes.Pdf;
    public virtual string Carrier => "Any";
    public virtual string DocumentType => DocumentTypes.Renewal;
    public virtual string ConnectorName => nameof(PdfRenewalConnector);
    public virtual bool IsEnabled => true;
    public virtual string Status => "Connected";

    public virtual Task<NormalizedImportPayload> ExtractAsync(BenefitsConnectorRequest request)
    {
        var parsed = _parser.Parse(request.InputPath);
        var carrier = FirstNonBlank(parsed.Carrier, parsed.Pdf.Carrier, Carrier);
        var source = NormalizedSourceMetadata.ForFile(SourceType, carrier, request.InputPath, PdfExtractionMethods.PdfPig, parsed.Pdf.RateLines.Count > 0 ? 0.90m : 0.50m);
        return Task.FromResult(NormalizeRenewalData.FromPdfExtraction(parsed.Pdf, source));
    }

    private static string FirstNonBlank(params string[] values)
    {
        return values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)) ?? "";
    }
}

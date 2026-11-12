using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services.Normalization;

namespace RenewalReviewBuilder.App.Connectors.Excel;

public sealed class ExcelRenewalConnector : IBenefitsConnector
{
    public string SourceType => SourceTypes.Excel;
    public string Carrier => "Any";
    public string DocumentType => DocumentTypes.Renewal;
    public string ConnectorName => nameof(ExcelRenewalConnector);
    public bool IsEnabled => true;
    public string Status => "Connected";

    public Task<NormalizedImportPayload> ExtractAsync(BenefitsConnectorRequest request)
    {
        var source = NormalizedSourceMetadata.ForFile(SourceType, Carrier, request.InputPath, "Excel workbook parser", null);
        return Task.FromResult(NormalizeQuoteData.Empty(source));
    }
}

using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services.Normalization;

namespace RenewalReviewBuilder.App.Connectors.Excel;

public sealed class ExcelBillingConnector : IBenefitsConnector
{
    public string SourceType => SourceTypes.Excel;
    public string Carrier => "Any";
    public string DocumentType => DocumentTypes.Billing;
    public string ConnectorName => nameof(ExcelBillingConnector);
    public bool IsEnabled => true;
    public string Status => "Connected";

    public Task<NormalizedImportPayload> ExtractAsync(BenefitsConnectorRequest request)
    {
        var source = NormalizedSourceMetadata.ForFile(SourceType, Carrier, request.InputPath, "Excel billing parser placeholder", null);
        return Task.FromResult(NormalizeBillingData.FromBilling(new BillingExtractionResult { SourceFile = request.InputPath }, source));
    }
}

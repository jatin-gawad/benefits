using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Connectors.Manual;

public sealed class ManualUploadConnector : IBenefitsConnector
{
    public string SourceType => SourceTypes.Manual;
    public string Carrier => "Any";
    public string DocumentType => DocumentTypes.Other;
    public string ConnectorName => nameof(ManualUploadConnector);
    public bool IsEnabled => true;
    public string Status => "Connected";

    public Task<NormalizedImportPayload> ExtractAsync(BenefitsConnectorRequest request)
    {
        var payload = request.ManualPayload ?? new NormalizedImportPayload
        {
            Source = new NormalizedSourceMetadata
            {
                SourceType = SourceType,
                SourceCarrier = Carrier,
                ImportedAt = DateTime.Now,
                ImportedBy = Environment.UserName,
                ExtractionMethod = "Manual entry"
            }
        };
        payload.ApplySourceToRecords();
        return Task.FromResult(payload);
    }
}

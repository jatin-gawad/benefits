using RenewalReviewBuilder.App.Connectors;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;
using RenewalReviewBuilder.App.Services.PolicySchedules;
using System.Threading;

namespace RenewalReviewBuilder.App.Connectors.Pdf;

/// <summary>
/// Phase 2 registration only. It deliberately does not interpret Manulife schedule fields.
/// </summary>
public interface IPolicyScheduleConnector : IBenefitsConnector
{
    Task<PolicyScheduleExtractionResult> ExtractPolicyScheduleAsync(BenefitsConnectorRequest request, CancellationToken cancellationToken = default);
}

public sealed class ManulifePolicyScheduleConnector : IPolicyScheduleConnector
{
    private readonly PdfTextExtractor _pdfTextExtractor;
    private readonly ManulifePolicyScheduleExtractor _extractor;

    public ManulifePolicyScheduleConnector(PdfTextExtractor? pdfTextExtractor = null, ManulifePolicyScheduleExtractor? extractor = null)
    {
        _pdfTextExtractor = pdfTextExtractor ?? new PdfTextExtractor();
        _extractor = extractor ?? new ManulifePolicyScheduleExtractor();
    }
    public string SourceType => SourceTypes.Pdf;
    public string Carrier => "Manulife";
    public string DocumentType => DocumentTypes.GroupPolicySchedule;
    public string ConnectorName => nameof(ManulifePolicyScheduleConnector);
    public bool IsEnabled => true;
    public string Status => "Connected";

    public Task<NormalizedImportPayload> ExtractAsync(BenefitsConnectorRequest request)
    {
        return ExtractPayloadAsync(request);
    }

    public Task<PolicyScheduleExtractionResult> ExtractPolicyScheduleAsync(BenefitsConnectorRequest request, CancellationToken cancellationToken = default)
    {
        var path = request.AllInputPaths().FirstOrDefault() ?? request.InputPath;
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path)) throw new FileNotFoundException("The Manulife policy schedule PDF was not found.", path);
        cancellationToken.ThrowIfCancellationRequested();
        var pages = _pdfTextExtractor.Extract(path);
        return Task.FromResult(_extractor.Extract(path, pages));
    }

    private async Task<NormalizedImportPayload> ExtractPayloadAsync(BenefitsConnectorRequest request)
    {
        var extraction = await ExtractPolicyScheduleAsync(request);
        var source = NormalizedSourceMetadata.ForFile(SourceType, Carrier, extraction.SourcePath, "Manulife policy schedule extractor v1", null);
        return new NormalizedImportPayload { DocumentType = DocumentType, Source = source, Carrier = new NormalizedCarrier { CarrierName = Carrier } };
    }
}

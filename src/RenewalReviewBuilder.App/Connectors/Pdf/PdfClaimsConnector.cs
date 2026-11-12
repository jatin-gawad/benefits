using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;
using RenewalReviewBuilder.App.Services.Normalization;

namespace RenewalReviewBuilder.App.Connectors.Pdf;

public sealed class PdfClaimsConnector : IBenefitsConnector
{
    private readonly ClaimsExperienceDataModelBuilder _builder;

    public PdfClaimsConnector(ClaimsExperienceDataModelBuilder? builder = null)
    {
        _builder = builder ?? new ClaimsExperienceDataModelBuilder();
    }

    public string SourceType => SourceTypes.Pdf;
    public string Carrier => "Any";
    public string DocumentType => DocumentTypes.ClaimsExperience;
    public string ConnectorName => nameof(PdfClaimsConnector);
    public bool IsEnabled => true;
    public string Status => "Connected";

    public async Task<NormalizedImportPayload> ExtractAsync(BenefitsConnectorRequest request)
    {
        var paths = request.AllInputPaths().ToList();
        var result = await _builder.ExportAsync(new ClaimsExperienceDataModelRequest
        {
            SourceFolder = request.SourceFolder,
            SourceFiles = paths,
            OutputFolder = request.OutputFolder
        });
        var firstPath = paths.FirstOrDefault() ?? "";
        var source = NormalizedSourceMetadata.ForFile(SourceType, result.Carrier, firstPath, PdfExtractionMethods.PdfPig, result.Summaries.Count > 0 ? 0.85m : 0.45m);
        return NormalizeClaimsData.FromClaimsExperienceResult(result, source);
    }
}

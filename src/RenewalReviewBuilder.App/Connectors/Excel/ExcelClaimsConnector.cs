using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;
using RenewalReviewBuilder.App.Services.Normalization;

namespace RenewalReviewBuilder.App.Connectors.Excel;

public sealed class ExcelClaimsConnector : IBenefitsConnector
{
    private readonly ManulifeWorkbookParser _parser;

    public ExcelClaimsConnector(ManulifeWorkbookParser? parser = null)
    {
        _parser = parser ?? new ManulifeWorkbookParser();
    }

    public string SourceType => SourceTypes.Excel;
    public string Carrier => "Any";
    public string DocumentType => DocumentTypes.ClaimsExperience;
    public string ConnectorName => nameof(ExcelClaimsConnector);
    public bool IsEnabled => true;
    public string Status => "Connected";

    public Task<NormalizedImportPayload> ExtractAsync(BenefitsConnectorRequest request)
    {
        var workbook = _parser.Parse(request.InputPath);
        var source = NormalizedSourceMetadata.ForFile(SourceType, Carrier, request.InputPath, "Claims workbook parser", workbook.ClaimTrends.Count > 0 ? 0.85m : 0.45m);
        return Task.FromResult(NormalizeClaimsData.FromWorkbook(workbook, source));
    }
}

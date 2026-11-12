using System.Threading;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services.FlexibleBenefits;

namespace RenewalReviewBuilder.App.Connectors.Pdf;

public interface IFlexibleBenefitsConnector : IBenefitsConnector
{
    Task<FlexibleBenefitExtractionResult> ExtractFlexibleBenefitsAsync(
        BenefitsConnectorRequest request,
        CancellationToken cancellationToken = default);

    Task<FlexibleBenefitExtractionResult> ParseFlexibleBenefitsAsync(
        FlexibleBenefitOcrDocument document,
        CancellationToken cancellationToken = default);
}

public sealed class NationalHealthClaimFlexibleBenefitsConnector : IFlexibleBenefitsConnector
{
    private readonly IFlexibleBenefitDocumentTextExtractor _textExtractor;
    private readonly IFlexibleBenefitDocumentParser _parser;
    private readonly FlexibleBenefitValidationService _validator;

    public NationalHealthClaimFlexibleBenefitsConnector(
        IFlexibleBenefitDocumentTextExtractor? textExtractor = null,
        IFlexibleBenefitDocumentParser? parser = null,
        FlexibleBenefitValidationService? validator = null)
    {
        _textExtractor = textExtractor ?? new WindowsOcrPdfExtractor();
        _parser = parser ?? new NationalHealthClaimFlexibleBenefitParser(_textExtractor);
        _validator = validator ?? new FlexibleBenefitValidationService();
    }

    public string SourceType => SourceTypes.Pdf;
    public string Carrier => "National HealthClaim";
    public string DocumentType => DocumentTypes.FlexibleBenefitsPlan;
    public string ConnectorName => nameof(NationalHealthClaimFlexibleBenefitsConnector);
    public bool IsEnabled => true;
    public string Status => "Connected";

    public async Task<NormalizedImportPayload> ExtractAsync(BenefitsConnectorRequest request)
    {
        var extraction = await ExtractFlexibleBenefitsAsync(request);
        var source = NormalizedSourceMetadata.ForFile(
            SourceType,
            extraction.Classification.Administrator,
            extraction.SourcePath,
            PdfExtractionMethods.WindowsLocalOcr,
            extraction.Confidence);
        var payload = new NormalizedImportPayload
        {
            DocumentType = DocumentType,
            Source = source,
            Client = new NormalizedClient
            {
                ClientName = FirstNonBlank(extraction.Client.DbaName, extraction.Client.LegalName),
                ClientNumber = request.ClientId,
                Province = extraction.Client.Province
            },
            Carrier = new NormalizedCarrier { CarrierName = extraction.Classification.Administrator },
            Policy = new NormalizedPolicy
            {
                PolicyNumber = extraction.Plan.PolicyNumber,
                ClientName = FirstNonBlank(extraction.Client.DbaName, extraction.Client.LegalName),
                CarrierName = extraction.Classification.Administrator,
                EffectiveDate = extraction.Plan.EffectiveDate
            }
        };
        payload.FlexibleBenefitPlans.Add(new NormalizedFlexibleBenefitPlan
        {
            AdministratorName = extraction.Classification.Administrator,
            PlanName = extraction.Plan.PlanName,
            PlanType = extraction.Plan.PlanType.ToString(),
            EffectiveDate = extraction.Plan.EffectiveDate,
            ClassCount = extraction.Classes.Count,
            EnabledCategoryCount = extraction.Categories.Count(category => category.Enabled == true),
            ReviewStatus = extraction.Status,
            Extraction = extraction
        });

        long? clientId = long.TryParse(request.ClientId, out var parsedClientId) ? parsedClientId : null;
        var validation = _validator.Validate(extraction, clientId);
        foreach (var issue in validation.Issues)
        {
            payload.ValidationIssues.Add(new ValidationIssueRecord
            {
                IssueCode = issue.Code,
                Severity = issue.Severity,
                TableName = "FlexibleBenefitPlan",
                FieldName = issue.FieldName,
                Message = issue.Message,
                RawSourceReference = extraction.SourcePath,
                CreatedAt = DateTime.Now
            });
        }

        payload.ApplySourceToRecords();
        return payload;
    }

    public async Task<FlexibleBenefitExtractionResult> ExtractFlexibleBenefitsAsync(
        BenefitsConnectorRequest request,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(request);
        var path = request.AllInputPaths().FirstOrDefault() ?? request.InputPath;
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            throw new FileNotFoundException("The flexible-benefits source PDF could not be found.", path);
        }

        var document = await _textExtractor.ExtractAsync(path, cancellationToken);
        return await ParseFlexibleBenefitsAsync(document, cancellationToken);
    }

    public async Task<FlexibleBenefitExtractionResult> ParseFlexibleBenefitsAsync(
        FlexibleBenefitOcrDocument document,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(document);
        if (!_parser.CanParse(document))
        {
            throw new InvalidDataException("The PDF is not a supported National HealthClaim flexible-benefits document.");
        }

        return await _parser.ParseAsync(document, cancellationToken);
    }

    private static string FirstNonBlank(params string?[] values)
    {
        return values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value))?.Trim() ?? "";
    }
}

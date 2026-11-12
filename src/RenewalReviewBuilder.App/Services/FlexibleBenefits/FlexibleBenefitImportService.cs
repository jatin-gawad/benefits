using System.IO;
using System.Threading;
using RenewalReviewBuilder.App.Connectors;
using RenewalReviewBuilder.App.Connectors.Pdf;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Repositories;

namespace RenewalReviewBuilder.App.Services.FlexibleBenefits;

public sealed class FlexibleBenefitImportService
{
    private readonly ConnectorRegistry _connectorRegistry;
    private readonly IFlexibleBenefitDocumentTextExtractor _textExtractor;
    private readonly FlexibleBenefitDocumentClassifier _classifier;
    private readonly FlexibleBenefitValidationService _validator;
    private readonly IFlexibleBenefitRepository _repository;

    public FlexibleBenefitImportService(
        string? databasePath = null,
        ConnectorRegistry? connectorRegistry = null,
        IFlexibleBenefitDocumentTextExtractor? textExtractor = null,
        FlexibleBenefitDocumentClassifier? classifier = null,
        FlexibleBenefitValidationService? validator = null,
        IFlexibleBenefitRepository? repository = null)
    {
        _connectorRegistry = connectorRegistry ?? new ConnectorRegistry();
        _textExtractor = textExtractor ?? new WindowsOcrPdfExtractor();
        _classifier = classifier ?? new FlexibleBenefitDocumentClassifier();
        _validator = validator ?? new FlexibleBenefitValidationService();
        _repository = repository ?? new FlexibleBenefitRepository(databasePath);
    }

    public async Task<FlexibleBenefitImportResult> ImportAsync(
        FlexibleBenefitImportRequest request,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(request);
        if (!request.ClientId.HasValue || request.ClientId <= 0)
        {
            return Failure("Confirm the client before importing a flexible-benefits document.");
        }

        if (string.IsNullOrWhiteSpace(request.FilePath) || !File.Exists(request.FilePath))
        {
            return Failure("The selected flexible-benefits PDF could not be found.");
        }

        try
        {
            var classificationDocument = await _textExtractor.ExtractAsync(request.FilePath, cancellationToken);
            var classification = _classifier.Classify(classificationDocument);
            if (string.IsNullOrWhiteSpace(classification.Administrator))
            {
                return Failure("The flexible-benefits administrator could not be confirmed from this PDF.");
            }

            var connector = _connectorRegistry.Resolve(
                SourceTypes.Pdf,
                classification.Administrator,
                DocumentTypes.FlexibleBenefitsPlan);
            if (connector is not IFlexibleBenefitsConnector flexibleConnector)
            {
                return Failure("The detected administrator does not have an enabled flexible-benefits parser yet.");
            }

            var extraction = await flexibleConnector.ParseFlexibleBenefitsAsync(
                classificationDocument,
                cancellationToken);
            var validation = _validator.Validate(extraction, request.ClientId);
            return await _repository.SaveAsync(request, extraction, validation);
        }
        catch (OperationCanceledException)
        {
            return new FlexibleBenefitImportResult
            {
                Outcome = FlexibleBenefitImportOutcome.Cancelled,
                Message = "The flexible-benefits import was cancelled.",
                ClientId = request.ClientId,
                Status = FlexibleBenefitExtractionStatuses.Failed
            };
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            return Failure("The flexible-benefits document could not be extracted. No normalized records were saved.", request.ClientId);
        }
    }

    private static FlexibleBenefitImportResult Failure(string message, long? clientId = null)
    {
        return new FlexibleBenefitImportResult
        {
            Outcome = FlexibleBenefitImportOutcome.Failed,
            Message = message,
            ClientId = clientId,
            Status = FlexibleBenefitExtractionStatuses.Failed
        };
    }
}

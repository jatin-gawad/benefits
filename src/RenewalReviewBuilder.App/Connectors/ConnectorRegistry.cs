using RenewalReviewBuilder.App.Connectors.Api;
using RenewalReviewBuilder.App.Connectors.Excel;
using RenewalReviewBuilder.App.Connectors.Manual;
using RenewalReviewBuilder.App.Connectors.Pdf;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Connectors;

public sealed class ConnectorRegistry
{
    private readonly List<IBenefitsConnector> _connectors;

    public ConnectorRegistry(IEnumerable<IBenefitsConnector>? connectors = null)
    {
        _connectors = connectors?.ToList() ?? CreateDefaultConnectors();
    }

    public IBenefitsConnector Resolve(string sourceType, string carrier, string documentType)
    {
        if (TryResolve(sourceType, carrier, documentType, out var connector))
        {
            return connector;
        }

        throw new InvalidOperationException($"No connector is registered for source '{sourceType}', carrier '{carrier}', document type '{documentType}'.");
    }

    public bool TryResolve(string sourceType, string carrier, string documentType, out IBenefitsConnector connector)
    {
        var normalizedSource = Normalize(sourceType);
        var normalizedCarrier = Normalize(carrier);
        var normalizedDocument = Normalize(documentType);

        var exact = _connectors.FirstOrDefault(connector =>
            Normalize(connector.SourceType) == normalizedSource &&
            Normalize(connector.Carrier) == normalizedCarrier &&
            Normalize(connector.DocumentType) == normalizedDocument);
        if (exact is not null)
        {
            connector = exact;
            return true;
        }

        var generic = _connectors.FirstOrDefault(connector =>
            Normalize(connector.SourceType) == normalizedSource &&
            IsGenericCarrier(connector.Carrier) &&
            Normalize(connector.DocumentType) == normalizedDocument);
        if (generic is not null)
        {
            connector = generic;
            return true;
        }

        connector = null!;
        return false;
    }

    public IReadOnlyList<ConnectorDescriptor> ListConnectors()
    {
        return _connectors
            .OrderBy(connector => connector.SourceType)
            .ThenBy(connector => connector.Carrier)
            .ThenBy(connector => connector.DocumentType)
            .Select(connector => new ConnectorDescriptor
            {
                SourceType = connector.SourceType,
                Carrier = connector.Carrier,
                DocumentType = connector.DocumentType,
                ConnectorName = connector.ConnectorName,
                Status = connector.Status,
                IsEnabled = connector.IsEnabled
            })
            .ToList();
    }

    private static List<IBenefitsConnector> CreateDefaultConnectors()
    {
        return
        [
            new PdfRenewalConnector(),
            new ManulifePdfRenewalConnector(),
            new PdfClaimsConnector(),
            new PdfPlanDesignConnector(),
            new ManulifePolicyScheduleConnector(),
            new NationalHealthClaimFlexibleBenefitsConnector(),
            new PdfBillingConnector(),
            new ExcelRenewalConnector(),
            new ExcelClaimsConnector(),
            new ExcelBillingConnector(),
            new ManualUploadConnector(),
            new CanadaLifeApiConnector(),
            new SunLifeApiConnector(),
            new ManulifeApiConnector(),
            new EquitableApiConnector()
        ];
    }

    private static bool IsGenericCarrier(string carrier)
    {
        return string.IsNullOrWhiteSpace(carrier) ||
               carrier.Equals("Any", StringComparison.OrdinalIgnoreCase) ||
               carrier.Equals("*", StringComparison.OrdinalIgnoreCase);
    }

    private static string Normalize(string value)
    {
        return (value ?? "")
            .Replace(" ", "", StringComparison.Ordinal)
            .Replace("_", "", StringComparison.Ordinal)
            .Replace("-", "", StringComparison.Ordinal)
            .Trim()
            .ToLowerInvariant();
    }
}

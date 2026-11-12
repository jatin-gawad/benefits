using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Connectors.Api;

public abstract class FutureCarrierApiConnector : IBenefitsConnector
{
    public string SourceType => SourceTypes.Api;
    public abstract string Carrier { get; }
    public string DocumentType => DocumentTypes.Renewal;
    public abstract string ConnectorName { get; }
    public bool IsEnabled => false;
    public string Status => "Not Connected";

    public Task<NormalizedImportPayload> ExtractAsync(BenefitsConnectorRequest request)
    {
        return FetchRenewalAsync(request.ClientId, request.PolicyId, request.Period);
    }

    public virtual Task<NormalizedImportPayload> FetchRenewalAsync(string clientId, string policyId, BenefitsConnectorPeriod period)
    {
        throw new NotImplementedException("Future API connector placeholder");
    }

    public virtual Task<NormalizedImportPayload> FetchClaimsAsync(string clientId, string policyId, BenefitsConnectorPeriod period)
    {
        throw new NotImplementedException("Future API connector placeholder");
    }

    public virtual Task<NormalizedImportPayload> FetchRatesAsync(string clientId, string policyId, BenefitsConnectorPeriod period)
    {
        throw new NotImplementedException("Future API connector placeholder");
    }

    public virtual Task<NormalizedImportPayload> FetchPlanDesignAsync(string clientId, string policyId)
    {
        throw new NotImplementedException("Future API connector placeholder");
    }
}

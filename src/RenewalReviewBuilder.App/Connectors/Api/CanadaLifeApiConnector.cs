namespace RenewalReviewBuilder.App.Connectors.Api;

public sealed class CanadaLifeApiConnector : FutureCarrierApiConnector
{
    public override string Carrier => "Canada Life";
    public override string ConnectorName => nameof(CanadaLifeApiConnector);
}

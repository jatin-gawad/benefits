namespace RenewalReviewBuilder.App.Connectors.Api;

public sealed class SunLifeApiConnector : FutureCarrierApiConnector
{
    public override string Carrier => "Sun Life";
    public override string ConnectorName => nameof(SunLifeApiConnector);
}

namespace RenewalReviewBuilder.App.Connectors.Api;

public sealed class EquitableApiConnector : FutureCarrierApiConnector
{
    public override string Carrier => "Equitable";
    public override string ConnectorName => nameof(EquitableApiConnector);
}

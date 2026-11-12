namespace RenewalReviewBuilder.App.Connectors.Api;

public sealed class ManulifeApiConnector : FutureCarrierApiConnector
{
    public override string Carrier => "Manulife";
    public override string ConnectorName => nameof(ManulifeApiConnector);
}

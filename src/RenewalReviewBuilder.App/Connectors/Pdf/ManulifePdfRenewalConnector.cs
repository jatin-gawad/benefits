namespace RenewalReviewBuilder.App.Connectors.Pdf;

public sealed class ManulifePdfRenewalConnector : PdfRenewalConnector
{
    public override string Carrier => "Manulife";
    public override string ConnectorName => nameof(ManulifePdfRenewalConnector);
}

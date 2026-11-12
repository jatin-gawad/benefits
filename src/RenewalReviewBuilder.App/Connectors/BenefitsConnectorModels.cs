using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Connectors;

public sealed class BenefitsConnectorRequest
{
    public string ClientId { get; set; } = "";
    public string PolicyId { get; set; } = "";
    public BenefitsConnectorPeriod Period { get; set; } = new();
    public string InputPath { get; set; } = "";
    public List<string> InputPaths { get; } = [];
    public string SourceFolder { get; set; } = "";
    public string OutputFolder { get; set; } = "";
    public NormalizedImportPayload? ManualPayload { get; set; }

    public IReadOnlyList<string> AllInputPaths()
    {
        return InputPaths.Count > 0
            ? InputPaths
            : string.IsNullOrWhiteSpace(InputPath) ? [] : [InputPath];
    }
}

public sealed class BenefitsConnectorPeriod
{
    public DateOnly? Start { get; set; }
    public DateOnly? End { get; set; }
}

public sealed class ConnectorDescriptor
{
    public string SourceType { get; set; } = "";
    public string Carrier { get; set; } = "";
    public string DocumentType { get; set; } = "";
    public string ConnectorName { get; set; } = "";
    public string Status { get; set; } = "Connected";
    public bool IsEnabled { get; set; } = true;
}

public interface IBenefitsConnector
{
    string SourceType { get; }
    string Carrier { get; }
    string DocumentType { get; }
    string ConnectorName { get; }
    bool IsEnabled { get; }
    string Status { get; }
    Task<NormalizedImportPayload> ExtractAsync(BenefitsConnectorRequest request);
}

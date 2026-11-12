using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed class RenewalBatchProcessor
{
    private readonly GenericRenewalProcessingEngine _engine;
    private readonly DocumentDiscoveryService _documentDiscoveryService;

    public RenewalBatchProcessor(
        GenericRenewalProcessingEngine? engine = null,
        DocumentDiscoveryService? documentDiscoveryService = null)
    {
        _engine = engine ?? new GenericRenewalProcessingEngine();
        _documentDiscoveryService = documentDiscoveryService ?? new DocumentDiscoveryService();
    }

    public async Task<List<RenewalProcessingResult>> ProcessFolderAsync(
        string inputFolder,
        string outputFolder,
        IProgress<string>? progress = null)
    {
        if (!Directory.Exists(inputFolder))
        {
            return
            [
                new RenewalProcessingResult
                {
                    OutputFolder = outputFolder,
                    Errors = { $"Input folder was not found: {inputFolder}" }
                }
            ];
        }

        var clientFolders = ResolveClientFolders(inputFolder);
        var results = new List<RenewalProcessingResult>();
        foreach (var clientFolder in clientFolders)
        {
            progress?.Report($"Discovering documents in {Path.GetFileName(clientFolder)}...");
            var discovery = _documentDiscoveryService.DiscoverClientFolder(clientFolder);
            if (discovery.RenewalDocument is null)
            {
                results.Add(CreateNeedsReviewResult(discovery, outputFolder));
                continue;
            }

            progress?.Report($"Processing renewal for {Path.GetFileName(clientFolder)}...");
            results.Add(await _engine.process_renewal_file(discovery.RenewalDocument.FilePath, outputFolder));
        }

        return results;
    }

    private static RenewalProcessingResult CreateNeedsReviewResult(
        RenewalDocumentDiscoveryResult discovery,
        string outputFolder)
    {
        var result = new RenewalProcessingResult
        {
            OutputFolder = outputFolder,
            Status = RenewalProcessingStatuses.NeedsReview,
            PlanDesignStatus = discovery.PlanDesign.Status,
            PlanDesignSource = discovery.PlanDesign.Source,
            DocumentDiscovery = discovery
        };
        result.Warnings.Add($"No renewal PDF was found in {discovery.ClientFolder}. Supporting documents were classified, but this client needs review.");
        result.ValidationSummary.Add(new RenewalValidationItem
        {
            Rule = "Renewal PDF",
            Status = "Needs Review",
            Message = "No renewal PDF was found. Batch processing can continue with the next client."
        });
        return result;
    }

    private static List<string> ResolveClientFolders(string inputFolder)
    {
        var hasRootPdfs = Directory.EnumerateFiles(inputFolder, "*.pdf", SearchOption.TopDirectoryOnly).Any();
        if (hasRootPdfs)
        {
            return [inputFolder];
        }

        var childFolders = Directory
            .EnumerateDirectories(inputFolder, "*", SearchOption.TopDirectoryOnly)
            .Where(folder => Directory.EnumerateFiles(folder, "*.pdf", SearchOption.AllDirectories).Any())
            .OrderBy(folder => folder, StringComparer.OrdinalIgnoreCase)
            .ToList();

        return childFolders.Count > 0 ? childFolders : [inputFolder];
    }
}

using System.Globalization;

namespace RenewalReviewBuilder.App.Models;

public sealed class ProposalPipelineOverview
{
    public string ProfileVersion { get; init; } = ProposalCanonicalProfileVersions.Current;
    public DateTime BuiltAtUtc { get; init; } = DateTime.UtcNow;
    public int ClientCount { get; init; }
    public int ProposalCount { get; init; }
    public int NeedsReviewCount { get; init; }
    public List<ProposalPipelineStageRow> Stages { get; } = [];

    public int TotalStages => Stages.Count;
    public int OpenItems => Stages
        .Where(stage => !stage.Stage.Equals("Sold / Lost", StringComparison.OrdinalIgnoreCase))
        .Sum(stage => stage.Count);
    public decimal TotalPremium => Stages
        .Where(stage => stage.Premium.HasValue)
        .Sum(stage => stage.Premium!.Value);
    public int OverdueItems => Stages.Sum(stage => stage.OverdueItems);
}

public sealed class ProposalPipelineStageRow
{
    public string Stage { get; init; } = "";
    public int Count { get; init; }
    public decimal? Premium { get; init; }
    public string PremiumText => Premium.HasValue
        ? Premium.Value.ToString("$#,0", CultureInfo.InvariantCulture)
        : "N/A";
    public string OldestItem { get; init; } = "N/A";
    public string AverageDaysInStage { get; init; } = "N/A";
    public int OverdueItems { get; init; }
}

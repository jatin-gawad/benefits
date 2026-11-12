using System.Globalization;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed class ProposalPipelineOverviewService
{
    private static readonly string[] PipelineStages =
    [
        "Intake received",
        "Carrier quotes pending",
        "Quotes received",
        "Comparison generated",
        "Underwriting review",
        "Consultant review",
        "Presentation ready",
        "Finalist selected",
        "Sold / Lost"
    ];

    private readonly ProposalCanonicalProfileService _profileService;

    public ProposalPipelineOverviewService(string? databasePath = null)
        : this(new ProposalCanonicalProfileService(databasePath))
    {
    }

    public ProposalPipelineOverviewService(ProposalCanonicalProfileService profileService)
    {
        _profileService = profileService ?? throw new ArgumentNullException(nameof(profileService));
    }

    public async Task<ProposalPipelineOverview> BuildAsync()
    {
        var profiles = await _profileService.BuildAllAsync();
        var items = profiles
            .SelectMany(profile => profile.Proposals.Select(proposal => BuildItem(profile, proposal)))
            .ToList();
        var overview = new ProposalPipelineOverview
        {
            ProfileVersion = ProposalCanonicalProfileVersions.Current,
            BuiltAtUtc = DateTime.UtcNow,
            ClientCount = profiles.Count,
            ProposalCount = items.Count,
            NeedsReviewCount = items.Count(item => item.NeedsReview)
        };

        foreach (var stage in PipelineStages)
        {
            var stageItems = items
                .Where(item => item.Stage.Equals(stage, StringComparison.OrdinalIgnoreCase))
                .ToList();
            var oldest = stageItems
                .Where(item => item.ActivityDate.HasValue)
                .OrderBy(item => item.ActivityDate)
                .FirstOrDefault();
            var daysInStage = stageItems
                .Where(item => item.ActivityDate.HasValue)
                .Select(item => Math.Max(0, (DateTime.Today - item.ActivityDate!.Value.Date).Days))
                .ToList();
            var premiums = stageItems
                .Where(item => item.Premium.HasValue)
                .Select(item => item.Premium!.Value)
                .ToList();

            overview.Stages.Add(new ProposalPipelineStageRow
            {
                Stage = stage,
                Count = stageItems.Count,
                Premium = premiums.Count == 0 ? null : premiums.Sum(),
                OldestItem = oldest is null
                    ? "N/A"
                    : $"{oldest.Label} ({oldest.ActivityDate:MMM d})",
                AverageDaysInStage = daysInStage.Count == 0
                    ? "N/A"
                    : ((int)Math.Round(daysInStage.Average(), MidpointRounding.AwayFromZero)).ToString(CultureInfo.InvariantCulture),
                OverdueItems = daysInStage.Count(days => days > 14)
            });
        }

        return overview;
    }

    private static PipelineItem BuildItem(
        ProposalCanonicalProfile profile,
        ProposalCanonicalQuote proposal)
    {
        var clientName = FirstNonBlank(
            profile.ClientName,
            proposal.ClientName,
            profile.ClientId > 0 ? $"Client #{profile.ClientId.ToString(CultureInfo.InvariantCulture)}" : "Client");
        var carrierName = FirstNonBlank(proposal.CarrierName, proposal.DetectedCarrier, "Proposal");
        return new PipelineItem(
            ResolveStage(profile, proposal),
            FirstNonBlank($"{clientName} - {carrierName}", clientName, carrierName),
            proposal.TotalAnnualPremium ?? proposal.TotalCost,
            proposal.UpdatedAt ??
            proposal.CreatedAt ??
            proposal.ProposalDate?.ToDateTime(TimeOnly.MinValue),
            proposal.ReadinessStatus == ProposalCanonicalReadinessStatuses.NeedsReview);
    }

    private static string ResolveStage(
        ProposalCanonicalProfile profile,
        ProposalCanonicalQuote proposal)
    {
        var status = proposal.Status;
        if (status.Contains("finalist", StringComparison.OrdinalIgnoreCase))
        {
            return "Finalist selected";
        }

        if (IsClosedStatus(status))
        {
            return "Sold / Lost";
        }

        if (status.Contains("presentation", StringComparison.OrdinalIgnoreCase))
        {
            return "Presentation ready";
        }

        if (status.Contains("consultant", StringComparison.OrdinalIgnoreCase))
        {
            return "Consultant review";
        }

        if (status.Contains("underwriting", StringComparison.OrdinalIgnoreCase) ||
            status.Contains("needs review", StringComparison.OrdinalIgnoreCase))
        {
            return "Underwriting review";
        }

        if (status.Contains("comparison", StringComparison.OrdinalIgnoreCase))
        {
            return "Comparison generated";
        }

        if (status.Contains("pending", StringComparison.OrdinalIgnoreCase) ||
            status.Contains("not available", StringComparison.OrdinalIgnoreCase))
        {
            return "Carrier quotes pending";
        }

        if (status.Contains("intake", StringComparison.OrdinalIgnoreCase) || !proposal.HasCoreFacts)
        {
            return "Intake received";
        }

        if (proposal.ReadinessStatus == ProposalCanonicalReadinessStatuses.NeedsReview)
        {
            return "Underwriting review";
        }

        if (!proposal.HasFinancialFacts)
        {
            return "Carrier quotes pending";
        }

        return profile.CanCompareCarriers ? "Comparison generated" : "Quotes received";
    }

    private static bool IsClosedStatus(string status) =>
        status.Contains("sold", StringComparison.OrdinalIgnoreCase) ||
        status.Contains("lost", StringComparison.OrdinalIgnoreCase) ||
        status.Contains("complete", StringComparison.OrdinalIgnoreCase) ||
        status.Contains("cancel", StringComparison.OrdinalIgnoreCase) ||
        status.Contains("closed", StringComparison.OrdinalIgnoreCase) ||
        status.Contains("declined", StringComparison.OrdinalIgnoreCase);

    private static string FirstNonBlank(params string?[] values) =>
        values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value))?.Trim() ?? "";

    private sealed record PipelineItem(
        string Stage,
        string Label,
        decimal? Premium,
        DateTime? ActivityDate,
        bool NeedsReview);
}

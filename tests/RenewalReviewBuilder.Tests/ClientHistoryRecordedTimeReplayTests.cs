using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Models.History;
using RenewalReviewBuilder.App.Models.Querying;
using RenewalReviewBuilder.App.Services.ClientHistory;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class ClientHistoryRecordedTimeReplayTests
{
    private const long ClientId = 42;
    private static readonly DateOnly ReplayDate = new(2024, 12, 31);

    [Fact]
    public async Task Carrier_Replay_Excludes_Backdated_Relationships_Not_Known_By_AsOf_Date()
    {
        var known = Administration(1, 1, "Manulife", 100, new DateTime(2020, 1, 1),
            new DateTime(2023, 1, 2, 0, 0, 0, DateTimeKind.Utc));
        var correctedLater = Administration(2, 2, "Sun Life", 200, new DateTime(2021, 1, 1),
            new DateTime(2022, 1, 2, 0, 0, 0, DateTimeKind.Utc));
        correctedLater.UpdatedAtUtc = new DateTime(2025, 1, 2, 0, 0, 0, DateTimeKind.Utc);
        var missingKnownAt = Administration(3, 3, "Canada Life", 300, new DateTime(2019, 1, 1), default);
        missingKnownAt.CreatedAtUtc = default;
        missingKnownAt.UpdatedAtUtc = default;
        var contributor = Carrier([known, correctedLater, missingKnownAt]);

        var current = await contributor.ContributeAsync(ClientId);
        var replay = await contributor.ContributeAsync(AsOf());

        Assert.Equal(3, current.Count);
        var knownEvent = Assert.Single(replay.Events);
        Assert.Contains("Carrier:100", knownEvent.RelatedRecords);
        Assert.DoesNotContain(replay.Events, item => item.RelatedRecords.Contains("Carrier:200"));
        Assert.DoesNotContain(replay.Events, item => item.RelatedRecords.Contains("Carrier:300"));
        Assert.Equal(HistoricalQuerySupport.Limited, replay.HistoricalSupport.Support);
        Assert.Contains("saved or updated after", replay.HistoricalSupport.Detail, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("without a canonical recorded time", replay.HistoricalSupport.Detail, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Collapsed_Carrier_Identity_Is_Client_Scoped_Coherent_And_Preserves_Evidence_Ownership()
    {
        var health = Administration(10, 1, "Manulife", 100, new DateTime(2024, 1, 1),
            new DateTime(2024, 1, 2, 0, 0, 0, DateTimeKind.Utc));
        health.SourceDocumentId = 900;
        var dental = Administration(20, 2, "Manulife", 100, new DateTime(2024, 1, 1),
            new DateTime(2024, 1, 3, 0, 0, 0, DateTimeKind.Utc));
        dental.SourceDocumentId = 901;

        var historyEvent = Assert.Single((await Carrier([health, dental]).ContributeAsync(AsOf())).Events);

        const string entityId = "client-42|20240101|none|id-100";
        const string stableKey = "CarrierVendor:CarrierRelationship:client-42|20240101|none|id-100:assigned";
        Assert.Equal("CarrierRelationship", historyEvent.SourceEntityType);
        Assert.Equal(entityId, historyEvent.SourceEntityId);
        Assert.Equal(stableKey, historyEvent.EventId);
        Assert.Equal(stableKey, historyEvent.StableSourceKey);
        Assert.Equal(new long?[] { 900, 901 }, historyEvent.Evidence.Select(item => item.DocumentId).OrderBy(item => item));
        Assert.All(historyEvent.Evidence, evidence => Assert.Equal(stableKey, evidence.EventId));
        Assert.Contains("BenefitAdministration:10", historyEvent.RelatedRecords);
        Assert.Contains("BenefitAdministration:20", historyEvent.RelatedRecords);
        Assert.Equal(new DateTime(2024, 1, 3, 0, 0, 0, DateTimeKind.Utc), historyEvent.RecordedDate);
    }

    [Fact]
    public async Task Plan_Design_Replay_Excludes_Backdated_Rules_Not_Known_By_AsOf_Date()
    {
        var known = Rule(1, 10, new DateOnly(2020, 1, 1), 80m,
            new DateTime(2023, 1, 2, 0, 0, 0, DateTimeKind.Utc));
        var correctedLater = Rule(2, 10, new DateOnly(2021, 1, 1), 90m,
            new DateTime(2022, 1, 2, 0, 0, 0, DateTimeKind.Utc));
        correctedLater.UpdatedAtUtc = new DateTime(2025, 1, 2, 0, 0, 0, DateTimeKind.Utc);
        var missingKnownAt = Rule(3, 20, new DateOnly(2019, 1, 1), 70m, default);
        missingKnownAt.CreatedAtUtc = default;
        missingKnownAt.UpdatedAtUtc = default;
        var contributor = Plan([known, correctedLater, missingKnownAt]);

        var current = await contributor.ContributeAsync(ClientId);
        var replay = await contributor.ContributeAsync(AsOf());

        Assert.Equal(3, current.Count);
        var knownEvent = Assert.Single(replay.Events);
        Assert.Equal("1", knownEvent.SourceEntityId);
        Assert.Equal("Plan design rule effective", knownEvent.EventType);
        Assert.DoesNotContain(replay.Events, item => item.EventType == "Plan design rule changed");
        Assert.Equal(HistoricalQuerySupport.Limited, replay.HistoricalSupport.Support);
        Assert.Contains("saved or updated after", replay.HistoricalSupport.Detail, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("without a canonical recorded time", replay.HistoricalSupport.Detail, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Plan_Rules_Without_Feature_Stay_Standalone_And_Report_Limited_Support()
    {
        var first = Rule(11, null, new DateOnly(2020, 1, 1), 80m,
            new DateTime(2020, 1, 2, 0, 0, 0, DateTimeKind.Utc));
        var second = Rule(12, null, new DateOnly(2021, 1, 1), 90m,
            new DateTime(2021, 1, 2, 0, 0, 0, DateTimeKind.Utc));

        var result = await Plan([first, second]).ContributeAsync(AsOf());

        Assert.Equal(2, result.Events.Count);
        Assert.All(result.Events, item => Assert.Equal("Plan design rule effective", item.EventType));
        Assert.DoesNotContain(result.Events, item => item.EventType == "Plan design rule changed");
        Assert.Equal(HistoricalQuerySupport.Limited, result.HistoricalSupport.Support);
        Assert.Contains("not linked to a canonical plan-design feature", result.HistoricalSupport.Detail, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Proposal_Replay_Excludes_Backdated_Profiles_Not_Known_By_AsOf_Date()
    {
        var profile = new ProposalCanonicalProfile { ClientId = ClientId };
        profile.Proposals.Add(Proposal(100, "2020-01-01", new DateTime(2023, 1, 2, 0, 0, 0, DateTimeKind.Utc), null));
        profile.Proposals.Add(Proposal(101, "2021-01-01", new DateTime(2022, 1, 2, 0, 0, 0, DateTimeKind.Utc),
            new DateTime(2025, 1, 2, 0, 0, 0, DateTimeKind.Utc)));
        profile.Proposals.Add(Proposal(102, "2019-01-01", null, null));
        var contributor = ProposalContributor(profile);

        var current = await contributor.ContributeAsync(ClientId);
        var replay = await contributor.ContributeAsync(AsOf());

        Assert.Equal(3, current.Count);
        var knownEvent = Assert.Single(replay.Events);
        Assert.Equal("100", knownEvent.SourceEntityId);
        Assert.DoesNotContain("Saved", knownEvent.Summary, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("status", knownEvent.Summary, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(HistoricalQuerySupport.Limited, replay.HistoricalSupport.Support);
        Assert.Contains("saved or updated after", replay.HistoricalSupport.Detail, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("without a canonical recorded time", replay.HistoricalSupport.Detail, StringComparison.OrdinalIgnoreCase);
    }

    private static AsOfQueryContext AsOf() => new()
    {
        ClientId = ClientId,
        AsOfDate = ReplayDate
    };

    private static CarrierVendorHistoryContributor Carrier(IReadOnlyList<ClientBenefitAdministrationRecord> records) =>
        new((_, _) => Task.FromResult(records));

    private static PlanDesignHistoryContributor Plan(IReadOnlyList<PlanBenefitRule> rules) =>
        new((_, _) => Task.FromResult(rules));

    private static ProposalMarketHistoryContributor ProposalContributor(ProposalCanonicalProfile profile) =>
        new((_, _) => Task.FromResult(profile));

    private static ClientBenefitAdministrationRecord Administration(
        long id,
        long benefitTypeId,
        string carrier,
        long carrierId,
        DateTime effectiveDate,
        DateTime recordedAt) => new()
    {
        ClientBenefitAdministrationId = id,
        ClientId = ClientId,
        BenefitTypeId = benefitTypeId,
        BenefitTypeCode = $"BENEFIT-{benefitTypeId}",
        BenefitDisplayName = $"Benefit {benefitTypeId}",
        InsurerOrganizationId = carrierId,
        InsurerName = carrier,
        InsurerOrganizationTypeCode = BenefitAdministrationOrganizationTypes.Insurer,
        EffectiveDate = effectiveDate,
        SourceDocumentId = 800 + id,
        SourceText = $"Carrier: {carrier}",
        ExtractionConfidence = .95m,
        ReviewStatusCode = BenefitAdministrationReviewStatuses.Confirmed,
        CreatedAtUtc = recordedAt,
        UpdatedAtUtc = recordedAt
    };

    private static PlanBenefitRule Rule(
        long id,
        long? featureId,
        DateOnly effectiveDate,
        decimal tier1,
        DateTime recordedAt) => new()
    {
        PlanBenefitRuleId = id,
        PlanDesignFeatureId = featureId,
        ClientId = ClientId,
        SourceDocumentId = 900 + id,
        RuleType = BenefitRuleType.TieredReimbursement,
        EffectiveDate = effectiveDate,
        Tier1ReimbursementPercent = tier1,
        Tier1ThresholdAmount = 500m,
        Tier2ReimbursementPercent = 100m,
        ThresholdBasis = BenefitRuleThresholdBasis.Person,
        ResetPeriod = BenefitRuleResetPeriod.CalendarYear,
        CurrencyCode = "CAD",
        SourceText = $"{tier1}% to 500; 100% thereafter",
        NormalizationStatus = BenefitRuleNormalizationStatus.Normalized,
        ConfidenceScore = .95m,
        DisplayOrder = 10,
        CreatedAtUtc = recordedAt,
        UpdatedAtUtc = recordedAt
    };

    private static ProposalCanonicalQuote Proposal(
        long id,
        string proposalDate,
        DateTime? createdAt,
        DateTime? updatedAt)
    {
        var parsed = DateOnly.ParseExact(proposalDate, "yyyy-MM-dd", CultureInfo.InvariantCulture);
        var proposal = new ProposalCanonicalQuote
        {
            ProposalId = id,
            CarrierId = 77,
            CarrierName = "Manulife",
            ProposalDate = parsed,
            ProposalDateSourceText = proposalDate,
            Status = "Saved",
            CreatedAt = createdAt,
            UpdatedAt = updatedAt,
            ExtractionConfidence = .90m,
            Evidence = new ProposalCanonicalEvidence
            {
                SourceTable = "proposal_summary",
                SourceId = id,
                SourceFileName = $"proposal-{id}.pdf",
                SourcePage = 2,
                ConfidenceScore = .90m
            }
        };
        proposal.BenefitLines.Add(new ProposalCanonicalBenefitLine
        {
            BenefitLine = "EHC",
            MonthlyPremium = 1000m
        });
        return proposal;
    }
}

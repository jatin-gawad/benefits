using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Models.History;
using RenewalReviewBuilder.App.Models.Querying;
using RenewalReviewBuilder.App.Services.ClientHistory;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class ClientHistoryCanonicalContributorTests
{
    private const long ClientId = 42;

    [Fact]
    public async Task Carrier_And_Provider_Map_Confirmed_Effective_Dated_Relationships_With_Evidence()
    {
        var record = Administration(10, "Manulife", 100, new DateTime(2021, 1, 1));
        record.ServiceProviders.Add(new BenefitServiceProviderAssignmentRecord
        {
            BenefitServiceProviderAssignmentId = 20,
            ClientBenefitAdministrationId = 10,
            ProviderOrganizationId = 200,
            ProviderName = "Homewood Health",
            ProviderRoleCode = BenefitAdministrationProviderRoles.TravelAssistance,
            EffectiveDate = new DateTime(2021, 2, 1),
            ReviewStatusCode = BenefitAdministrationReviewStatuses.Confirmed,
            SourceDocumentId = 901,
            SourcePageNumber = 8,
            CreatedAtUtc = new DateTime(2021, 2, 2, 0, 0, 0, DateTimeKind.Utc)
        });
        var contributor = Carrier([record]);

        var first = await contributor.ContributeAsync(ClientId);
        var second = await contributor.ContributeAsync(ClientId);

        Assert.Equal(2, first.Count);
        Assert.Equal(first.Select(item => item.StableSourceKey), second.Select(item => item.StableSourceKey));
        var carrier = Assert.Single(first, item => item.EventType == "Carrier assignment");
        Assert.Equal(new DateOnly(2021, 1, 1), carrier.EffectiveDate);
        Assert.Contains("Carrier:100", carrier.RelatedRecords);
        Assert.Equal(900, Assert.Single(carrier.Evidence).DocumentId);
        Assert.Equal(HistoryValidationStatus.Valid, carrier.ValidationStatus);
        var carrierReference = Assert.Single(carrier.ChangeSet.Changes, item => item.FieldKey == "Carrier");
        Assert.Equal(HistoryValueType.CarrierReference, carrierReference.ValueType);
        Assert.Equal("Manulife", carrierReference.NewValue);
        Assert.Equal("Manulife", new ClientHistoryTimelineService()
            .CalculateKpis(first, new DateOnly(2022, 1, 1)).CurrentCarrier);
        var provider = Assert.Single(first, item => item.EventType == "Service provider assignment");
        Assert.Contains("Organization:200", provider.RelatedRecords);
        Assert.DoesNotContain(provider.RelatedRecords, item => item.StartsWith("Carrier:", StringComparison.Ordinal));
        Assert.Equal(HistoryValidationStatus.Valid, provider.ValidationStatus);
    }

    [Fact]
    public async Task Carrier_Explicit_Terminations_Are_Emitted_Once_And_Replacement_Is_A_New_Assignment()
    {
        var oldCarrier = Administration(10, "GroupHealth", 100, new DateTime(2019, 1, 1));
        oldCarrier.TerminationDate = new DateTime(2024, 9, 30);
        var newCarrier = Administration(11, "Manulife", 101, new DateTime(2024, 10, 1));
        newCarrier.TerminationDate = new DateTime(2026, 9, 30);
        var contributor = Carrier([oldCarrier, newCarrier]);

        var events = await contributor.ContributeAsync(ClientId);

        Assert.Equal(4, events.Count);
        Assert.Equal(2, events.Count(item => item.EventType == "Carrier assignment"));
        Assert.Equal(2, events.Count(item => item.EventType == "Carrier termination"));
        Assert.DoesNotContain(events, item => item.EventType == "Carrier change");
        Assert.Contains(events, item => item.EventType == "Carrier termination" && item.RelatedRecords.Contains("BenefitAdministration:10"));
        Assert.Contains(events, item => item.EventType == "Carrier termination" && item.RelatedRecords.Contains("BenefitAdministration:11"));
        Assert.Equal("Not available", new ClientHistoryTimelineService()
            .CalculateKpis(events, new DateOnly(2027, 1, 1)).CurrentCarrier);
        Assert.Equal("Manulife", new ClientHistoryTimelineService()
            .CalculateKpis(events.Where(item => item.EffectiveDate <= new DateOnly(2025, 1, 1)), new DateOnly(2025, 1, 1)).CurrentCarrier);
    }

    [Fact]
    public async Task Carrier_Gap_And_Provider_Terminations_Preserve_Every_Confirmed_Snapshot()
    {
        var oldCarrier = Administration(10, "GroupHealth", 100, new DateTime(2018, 1, 1));
        oldCarrier.TerminationDate = new DateTime(2020, 12, 31);
        oldCarrier.ServiceProviders.Add(new BenefitServiceProviderAssignmentRecord
        {
            BenefitServiceProviderAssignmentId = 20,
            ClientBenefitAdministrationId = 10,
            ProviderOrganizationId = 200,
            ProviderName = "Homewood Health",
            ProviderRoleCode = BenefitAdministrationProviderRoles.Other,
            EffectiveDate = new DateTime(2018, 1, 1),
            TerminationDate = new DateTime(2020, 12, 31),
            ReviewStatusCode = BenefitAdministrationReviewStatuses.Confirmed,
            SourceDocumentId = 901,
            SourcePageNumber = 8
        });
        var newCarrier = Administration(11, "Manulife", 101, new DateTime(2022, 1, 1));
        var contributor = Carrier([oldCarrier, newCarrier]);

        var first = await contributor.ContributeAsync(ClientId);
        var second = await contributor.ContributeAsync(ClientId);

        Assert.Equal(first.Select(item => item.StableSourceKey), second.Select(item => item.StableSourceKey));
        Assert.DoesNotContain(first, item => item.EventType == "Carrier change");
        var carrierTermination = Assert.Single(first, item => item.EventType == "Carrier termination");
        Assert.Equal("CarrierRelationship", carrierTermination.SourceEntityType);
        Assert.Contains("BenefitAdministration:10", carrierTermination.RelatedRecords);
        Assert.Equal(900, Assert.Single(carrierTermination.Evidence).DocumentId);
        var providerTermination = Assert.Single(first, item => item.EventType == "Service provider termination");
        Assert.Equal("20", providerTermination.SourceEntityId);
        Assert.Equal(901, Assert.Single(providerTermination.Evidence).DocumentId);
        Assert.Equal("Not available", new ClientHistoryTimelineService()
            .CalculateKpis(first.Where(item => item.EffectiveDate <= new DateOnly(2021, 6, 1)), new DateOnly(2021, 6, 1)).CurrentCarrier);
    }

    [Fact]
    public async Task Carrier_Exact_Termination_Remains_A_Fact_When_Assignment_Start_Is_Unknown()
    {
        var record = Administration(10, "GroupHealth", 100, null);
        record.TerminationDate = new DateTime(2020, 12, 31);

        var result = await Carrier([record]).ContributeAsync(new AsOfQueryContext
        {
            ClientId = ClientId,
            AsOfDate = new DateOnly(2021, 1, 3)
        });

        var termination = Assert.Single(result.Events);
        Assert.Equal("Carrier termination", termination.EventType);
        Assert.Equal(new DateOnly(2020, 12, 31), termination.EffectiveDate);
        Assert.Equal(HistoricalQuerySupport.Limited, result.HistoricalSupport.Support);
    }

    [Fact]
    public async Task Carrier_Duplicate_Cross_Benefit_Transitions_Collapse_And_Merge_Evidence()
    {
        var oldHealth = Administration(10, "GroupHealth", 100, new DateTime(2019, 1, 1));
        var oldDental = Administration(20, "GroupHealth", 100, new DateTime(2019, 1, 1));
        oldDental.BenefitTypeId = 2;
        oldDental.BenefitTypeCode = "DENTAL";
        oldDental.SourceDocumentId = 901;
        var newHealth = Administration(11, "Manulife", 101, new DateTime(2024, 10, 1));
        newHealth.SourceDocumentId = 902;
        var newDental = Administration(21, "Manulife", 101, new DateTime(2024, 10, 1));
        newDental.BenefitTypeId = 2;
        newDental.BenefitTypeCode = "DENTAL";
        newDental.SourceDocumentId = 903;
        var contributor = Carrier([oldHealth, oldDental, newHealth, newDental]);

        var first = await contributor.ContributeAsync(ClientId);
        var second = await contributor.ContributeAsync(ClientId);

        Assert.Equal(2, first.Count);
        Assert.Equal(first.Select(item => item.StableSourceKey), second.Select(item => item.StableSourceKey));
        var change = Assert.Single(first, item => item.EventType == "Carrier change");
        Assert.Equal(new long?[] { 902, 903 }, change.Evidence.Select(item => item.DocumentId).OrderBy(value => value));
        Assert.Contains("BenefitType:1", change.RelatedRecords);
        Assert.Contains("BenefitType:2", change.RelatedRecords);
        Assert.Equal("Manulife", new ClientHistoryTimelineService()
            .CalculateKpis(first, new DateOnly(2025, 1, 1)).CurrentCarrier);
    }

    [Fact]
    public async Task Carrier_Kpi_Does_Not_Arbitrarily_Choose_One_Of_Multiple_Active_Benefit_Carriers()
    {
        var health = Administration(10, "Manulife", 100, new DateTime(2024, 1, 1));
        var dental = Administration(20, "Sun Life", 200, new DateTime(2024, 1, 1));
        dental.BenefitTypeId = 2;
        dental.BenefitTypeCode = "DENTAL";

        var events = await Carrier([health, dental]).ContributeAsync(ClientId);

        Assert.Equal("Multiple carriers", new ClientHistoryTimelineService()
            .CalculateKpis(events, new DateOnly(2025, 1, 1)).CurrentCarrier);
    }

    [Fact]
    public async Task Provider_Needs_Independent_Confirmation_Before_It_Becomes_History()
    {
        var record = Administration(10, "Manulife", 100, new DateTime(2021, 1, 1));
        record.ServiceProviders.Add(new BenefitServiceProviderAssignmentRecord
        {
            BenefitServiceProviderAssignmentId = 20,
            ClientBenefitAdministrationId = 10,
            ProviderOrganizationId = 200,
            ProviderName = "Homewood Health",
            ProviderRoleCode = BenefitAdministrationProviderRoles.Other,
            EffectiveDate = new DateTime(2021, 1, 1),
            ReviewStatusCode = BenefitAdministrationReviewStatuses.NeedsReview
        });

        var result = await Carrier([record]).ContributeAsync(new AsOfQueryContext
        {
            ClientId = ClientId,
            AsOfDate = new DateOnly(2022, 1, 1)
        });

        Assert.Single(result.Events, item => item.EventType == "Carrier assignment");
        Assert.DoesNotContain(result.Events, item => item.EventType.StartsWith("Service provider", StringComparison.Ordinal));
        Assert.Equal(HistoricalQuerySupport.Limited, result.HistoricalSupport.Support);
    }

    [Fact]
    public async Task Carrier_Missing_Identity_Date_And_Review_State_Surface_Limited_Without_Fabrication()
    {
        var undated = Administration(10, "Manulife", 100, null);
        undated.ReviewStatusCode = BenefitAdministrationReviewStatuses.NeedsReview;
        var contributor = Carrier([undated]);

        var result = await contributor.ContributeAsync(new AsOfQueryContext { ClientId = ClientId, AsOfDate = new DateOnly(2025, 1, 1) });

        Assert.Empty(result.Events);
        Assert.Equal(HistoricalQuerySupport.Limited, result.HistoricalSupport.Support);
    }

    [Fact]
    public async Task Plan_Design_Uses_Normalized_Effective_Dated_Rules_And_Typed_Comparable_Changes()
    {
        var oldRule = Rule(1, new DateOnly(2021, 1, 1), 80m, 500m, 100m);
        oldRule.IsActive = false;
        oldRule.TerminationDate = new DateOnly(2023, 12, 31);
        var currentRule = Rule(2, new DateOnly(2024, 1, 1), 90m, 750m, 100m);
        currentRule.TerminationDate = new DateOnly(2026, 12, 31);
        var contributor = Plan([oldRule, currentRule]);

        var first = await contributor.ContributeAsync(ClientId);
        var second = await contributor.ContributeAsync(ClientId);

        Assert.Equal(3, first.Count);
        Assert.Equal(first.Select(item => item.StableSourceKey), second.Select(item => item.StableSourceKey));
        var effective = Assert.Single(first, item => item.EventType == "Plan design rule effective");
        Assert.Collection(
            effective.ChangeSet.Changes.OrderBy(item => item.FieldKey),
            item => AssertInitial(item, "CurrencyCode", HistoryValueType.Text, "CAD"),
            item => AssertInitial(item, "ResetPeriod", HistoryValueType.Text, "CalendarYear"),
            item => AssertInitial(item, "RuleType", HistoryValueType.Text, "TieredReimbursement"),
            item => AssertInitial(item, "ThresholdBasis", HistoryValueType.Text, "Person"),
            item => AssertInitial(item, "Tier1ReimbursementPercent", HistoryValueType.Percentage, "80", "%"),
            item => AssertInitial(item, "Tier1ThresholdAmount", HistoryValueType.Money, "500", "CAD"),
            item => AssertInitial(item, "Tier2ReimbursementPercent", HistoryValueType.Percentage, "100", "%"));
        var changed = Assert.Single(first, item => item.EventType == "Plan design rule changed");
        Assert.Contains(changed.ChangeSet.Changes, item => item.FieldKey == "Tier1ReimbursementPercent" && item.PreviousValue == "80" && item.NewValue == "90");
        Assert.Equal(501, Assert.Single(changed.Evidence).DocumentId);
        Assert.All(first, item => Assert.Equal(HistoryDatePrecision.ExactDate, item.EffectiveDateReference.Precision));
        var status = await contributor.ContributeAsync(new AsOfQueryContext { ClientId = ClientId, AsOfDate = new DateOnly(2025, 1, 1) });
        Assert.Equal(HistoricalQuerySupport.Limited, status.HistoricalSupport.Support);
    }

    [Fact]
    public async Task Plan_Design_Stable_Event_Identity_Uses_Persisted_Rule_Id_Not_Mutable_Provenance()
    {
        var original = Rule(19, new DateOnly(2024, 1, 1), 80m, 500m, 100m);
        var updatedProvenance = Rule(19, new DateOnly(2025, 2, 1), 80m, 500m, 100m);
        updatedProvenance.SourceDocumentId = 999;
        updatedProvenance.PlanDesignFeatureId = 88;
        updatedProvenance.DisplayOrder = 30;

        var originalEvent = Assert.Single(await Plan([original]).ContributeAsync(ClientId));
        var updatedEvent = Assert.Single(await Plan([updatedProvenance]).ContributeAsync(ClientId));

        Assert.Equal(originalEvent.StableSourceKey, updatedEvent.StableSourceKey);
        Assert.Equal("19", originalEvent.SourceEntityId);
        Assert.Equal("19", updatedEvent.SourceEntityId);
    }

    [Fact]
    public async Task Plan_Design_Omits_Undated_And_Unnormalized_Rules_Instead_Of_Using_Load_Time()
    {
        var undated = Rule(1, null, 80m, 500m, 100m);
        var needsReview = Rule(2, new DateOnly(2024, 1, 1), 80m, 500m, 100m);
        needsReview.NormalizationStatus = BenefitRuleNormalizationStatus.NeedsReview;
        var contributor = Plan([undated, needsReview]);

        var result = await contributor.ContributeAsync(new AsOfQueryContext { ClientId = ClientId, AsOfDate = new DateOnly(2025, 1, 1) });

        Assert.Empty(result.Events);
        Assert.Equal(HistoricalQuerySupport.Limited, result.HistoricalSupport.Support);
        Assert.Contains("undated", result.HistoricalSupport.Detail, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Plan_Design_Identical_Contiguous_Successor_Advances_Termination_And_Evidence()
    {
        var original = Rule(1, new DateOnly(2021, 1, 1), 80m, 500m, 100m);
        original.TerminationDate = new DateOnly(2023, 12, 31);
        var successor = Rule(2, new DateOnly(2024, 1, 1), 80m, 500m, 100m);
        successor.SourceDocumentId = 502;
        successor.SourceText = "Renewed unchanged; termination December 31, 2026";
        successor.TerminationDate = new DateOnly(2026, 12, 31);
        var contributor = Plan([original, successor]);

        var first = await contributor.ContributeAsync(ClientId);
        var second = await contributor.ContributeAsync(ClientId);

        Assert.Equal(2, first.Count);
        Assert.DoesNotContain(first, item => item.EventType == "Plan design rule changed");
        var termination = Assert.Single(first, item => item.EventType == "Plan design rule termination");
        Assert.Equal(new DateOnly(2026, 12, 31), termination.EffectiveDate);
        Assert.Equal("2", termination.SourceEntityId);
        Assert.Equal(502, Assert.Single(termination.Evidence).DocumentId);
        Assert.Equal(first.Select(item => item.StableSourceKey), second.Select(item => item.StableSourceKey));
    }

    [Fact]
    public async Task Plan_Design_Exact_Intermediate_Termination_And_Gap_Restart_Do_Not_Invent_Change()
    {
        var original = Rule(1, new DateOnly(2021, 1, 1), 80m, 500m, 100m);
        original.TerminationDate = new DateOnly(2022, 12, 31);
        var afterGap = Rule(2, new DateOnly(2024, 1, 1), 90m, 750m, 100m);
        afterGap.SourceDocumentId = 502;
        var contributor = Plan([original, afterGap]);

        var events = await contributor.ContributeAsync(ClientId);

        Assert.Equal(3, events.Count);
        var termination = Assert.Single(events, item => item.EventType == "Plan design rule termination");
        Assert.Equal(new DateOnly(2022, 12, 31), termination.EffectiveDate);
        Assert.Equal("1", termination.SourceEntityId);
        Assert.Equal(501, Assert.Single(termination.Evidence).DocumentId);
        var effectiveEvents = events.Where(item => item.EventType == "Plan design rule effective").ToList();
        Assert.Equal(2, effectiveEvents.Count);
        var restarted = Assert.Single(effectiveEvents, item => item.SourceEntityId == "2");
        Assert.Equal(new DateOnly(2024, 1, 1), restarted.EffectiveDate);
        Assert.Equal(502, Assert.Single(restarted.Evidence).DocumentId);
        Assert.Equal(7, restarted.ChangeSet.Changes.Count);
        Assert.All(restarted.ChangeSet.Changes, item => Assert.Equal("", item.PreviousValue));
        Assert.DoesNotContain(events, item => item.EventType == "Plan design rule changed");
        Assert.Equal(events.Select(item => item.StableSourceKey),
            (await contributor.ContributeAsync(ClientId)).Select(item => item.StableSourceKey));
    }

    [Fact]
    public async Task Financial_Maps_Confirmed_Version_And_Current_Complete_Aso_Milestones_Without_Kpi_Value_Fabrication()
    {
        var version = new FinancialArrangementHistoryVersion
        {
            FinancialArrangementVersionId = 1,
            FinancialArrangementId = 11,
            ClientId = ClientId,
            Carrier = "Manulife",
            Scenario = FinancialArrangementScenario.Current,
            EffectiveDate = new DateOnly(2024, 1, 1),
            ArrangementStatus = "Confirmed",
            SourceDocumentId = 600,
            TargetLossRatio = .68m,
            TargetLossRatioBasis = "Benefits / deposits",
            TargetLossRatioSourceText = "Target loss ratio 68%",
            TargetLossRatioSourcePage = 4,
            TargetLossRatioConfidence = .96m,
            CreatedAtUtc = new DateTime(2024, 1, 2, 0, 0, 0, DateTimeKind.Utc)
        };
        var audit = new FinancialArrangementAuditEvent
        {
            FinancialArrangementAuditEventId = 2,
            EntityType = "FinancialArrangement",
            EntityId = 11,
            ClientId = ClientId,
            EventType = "FinancialArrangementConfirmed",
            EffectiveDate = new DateOnly(2024, 1, 1),
            User = "consultant",
            Reason = "Confirmed against contract",
            OccurredAtUtc = new DateTime(2024, 1, 3, 0, 0, 0, DateTimeKind.Utc)
        };
        var statement = Statement(7, new DateOnly(2024, 12, 31));
        statement.Total = new AsoFinancialActivityPeriod
        {
            PeriodId = 70,
            StatementId = 7,
            IsTotal = true,
            CurrentBalance = -1000m,
            BalanceStatus = "Surplus",
            DeficitPaymentRefund = 250m
        };
        var contributor = Financial([version], [audit], [statement]);

        var first = await contributor.ContributeAsync(ClientId);
        var second = await contributor.ContributeAsync(ClientId);

        Assert.Equal(4, first.Count);
        Assert.Equal(first.Select(item => item.StableSourceKey), second.Select(item => item.StableSourceKey));
        Assert.All(first, item => Assert.Empty(item.FinancialImpact));
        var funding = Assert.Single(first, item => item.EventType == "ASO funding position");
        Assert.Contains(funding.ChangeSet.Changes, item => item.FieldKey == "BalanceStatus" && item.NewValue == "Surplus");
        Assert.Contains(funding.ChangeSet.Changes, item => item.FieldKey == "CurrentBalance" && item.NewValue == "-1000");
        var arrangement = Assert.Single(first, item => item.SourceEntityType == "FinancialArrangementVersion");
        Assert.Equal("consultant", arrangement.ApprovedBy);
        Assert.Equal(600, Assert.Single(arrangement.Evidence).DocumentId);
        var support = await contributor.ContributeAsync(new AsOfQueryContext { ClientId = ClientId, AsOfDate = new DateOnly(2025, 1, 1) });
        Assert.Equal(HistoricalQuerySupport.Limited, support.HistoricalSupport.Support);
    }

    [Fact]
    public async Task Financial_Maps_Only_Approved_Linked_Commission_Pooling_And_Charge_Terms()
    {
        var version = new FinancialArrangementHistoryVersion
        {
            FinancialArrangementVersionId = 41,
            FinancialArrangementId = 410,
            ClientId = ClientId,
            Carrier = "Manulife",
            Scenario = FinancialArrangementScenario.Current,
            EffectiveDate = new DateOnly(2025, 1, 1),
            ArrangementStatus = "Confirmed",
            SourceDocumentId = 600,
            CommissionArrangementId = 420,
            CommissionConfirmationStatus = "Confirmed",
            CommissionRate = .05m,
            CommissionAmount = 12000m,
            CommissionBasis = "Total benefit payments",
            CommissionSource = "AOR",
            CommissionSourceDocumentId = 601,
            PoolingArrangementId = 430,
            PoolingBenefit = "EHC",
            PoolingRate = .077m,
            PoolingThreshold = 15000m,
            PoolingBasis = "Non-pooled EHC benefit payments",
            PoolingType = "Large amount pooling",
            PoolingSourceDocumentId = 602,
            PoolingSourceText = "Pooling 7.7% above $15,000",
            PoolingSourcePage = 8,
            PoolingConfidence = .97m
        };
        version.Charges.Add(new FinancialArrangementCharge
        {
            FinancialArrangementChargeId = 440,
            ChargeType = FinancialChargeType.GeneralAdministration,
            Benefit = "EHC",
            Rate = .065m,
            Amount = 9000m,
            Basis = "Total benefit payments",
            SourceText = "Administration charge 6.5%",
            SourcePage = 9,
            ExtractionConfidence = .96m
        });
        var poolingAudit = new FinancialArrangementAuditEvent
        {
            EntityType = "PoolingArrangement",
            EntityId = 430,
            ClientId = ClientId,
            EventType = "PoolingConfirmed",
            EffectiveDate = version.EffectiveDate,
            OccurredAtUtc = new DateTime(2025, 1, 2, 0, 0, 0, DateTimeKind.Utc)
        };

        var historyEvent = Assert.Single(await Financial([version], [poolingAudit], []).ContributeAsync(ClientId));

        Assert.Empty(historyEvent.FinancialImpact);
        Assert.Contains(historyEvent.ChangeSet.Changes, item => item.FieldKey == "CommissionRate" && item.NewValue == "5");
        Assert.Contains(historyEvent.ChangeSet.Changes, item => item.FieldKey == "CommissionAmount" && item.NewValue == "12000");
        Assert.Contains(historyEvent.ChangeSet.Changes, item => item.FieldKey == "PoolingRate" && item.NewValue == "7.7");
        Assert.Contains(historyEvent.ChangeSet.Changes, item => item.FieldKey == "PoolingThreshold" && item.NewValue == "15000");
        Assert.Contains(historyEvent.ChangeSet.Changes, item => item.FieldKey.StartsWith("Charge:GeneralAdministration:EHC:", StringComparison.Ordinal) && item.NewValue == "6.5");
        Assert.Contains("CommissionArrangement:420", historyEvent.RelatedRecords);
        Assert.Contains("PoolingArrangement:430", historyEvent.RelatedRecords);
        Assert.Contains(historyEvent.Evidence, item => item.DocumentId == 601);
        Assert.Contains(historyEvent.Evidence, item => item.DocumentId == 602 && item.PageNumber == 8);

        version.CommissionConfirmationStatus = "RequiresReview";
        var unapproved = Assert.Single(await Financial([version], [], []).ContributeAsync(ClientId));
        Assert.DoesNotContain(unapproved.ChangeSet.Changes, item => item.FieldKey.StartsWith("Commission", StringComparison.Ordinal));
        Assert.DoesNotContain(unapproved.ChangeSet.Changes, item => item.FieldKey.StartsWith("Pooling", StringComparison.Ordinal));
        Assert.DoesNotContain(unapproved.RelatedRecords, item => item.StartsWith("CommissionArrangement:", StringComparison.Ordinal));
        Assert.DoesNotContain(unapproved.RelatedRecords, item => item.StartsWith("PoolingArrangement:", StringComparison.Ordinal));
        Assert.Contains(unapproved.ChangeSet.Changes, item => item.FieldKey.StartsWith("Charge:GeneralAdministration:EHC:", StringComparison.Ordinal));
    }

    [Fact]
    public async Task Financial_Replay_Exposes_Pooling_Only_After_Its_Canonical_Confirmation_Time()
    {
        var version = new FinancialArrangementHistoryVersion
        {
            FinancialArrangementVersionId = 42,
            FinancialArrangementId = 411,
            ClientId = ClientId,
            Carrier = "Manulife",
            Scenario = FinancialArrangementScenario.Current,
            EffectiveDate = new DateOnly(2025, 1, 1),
            ArrangementStatus = "Confirmed",
            TargetLossRatio = .68m,
            PoolingArrangementId = 431,
            PoolingBenefit = "EHC",
            PoolingRate = .077m,
            PoolingThreshold = 15000m,
            PoolingBasis = "Non-pooled EHC benefit payments",
            CreatedAtUtc = new DateTime(2025, 1, 2, 0, 0, 0, DateTimeKind.Utc)
        };
        var poolingAudit = new FinancialArrangementAuditEvent
        {
            EntityType = "PoolingArrangement",
            EntityId = 431,
            ClientId = ClientId,
            EventType = "PoolingConfirmed",
            EffectiveDate = version.EffectiveDate,
            OccurredAtUtc = new DateTime(2025, 2, 15, 0, 0, 0, DateTimeKind.Utc)
        };
        var contributor = Financial([version], [poolingAudit], []);

        var before = Assert.Single((await contributor.ContributeAsync(new AsOfQueryContext
        {
            ClientId = ClientId,
            AsOfDate = new DateOnly(2025, 1, 31)
        })).Events);
        var after = Assert.Single((await contributor.ContributeAsync(new AsOfQueryContext
        {
            ClientId = ClientId,
            AsOfDate = new DateOnly(2025, 2, 15)
        })).Events);

        Assert.DoesNotContain(before.ChangeSet.Changes, item => item.FieldKey.StartsWith("Pooling", StringComparison.Ordinal));
        Assert.Contains(after.ChangeSet.Changes, item => item.FieldKey == "PoolingRate" && item.NewValue == "7.7");
        Assert.Equal(before.StableSourceKey, after.StableSourceKey);
    }

    [Fact]
    public async Task Financial_First_Version_Emits_NewValues_Then_Skips_Repeats_And_Emits_Canonical_Removals()
    {
        static FinancialArrangementHistoryVersion Version(long versionId, DateOnly effectiveDate, bool includeTerms)
        {
            var version = new FinancialArrangementHistoryVersion
            {
                FinancialArrangementVersionId = versionId,
                FinancialArrangementId = 500 + versionId,
                ClientId = ClientId,
                Carrier = "Manulife",
                Scenario = FinancialArrangementScenario.Current,
                EffectiveDate = effectiveDate,
                ArrangementStatus = "Confirmed",
                TargetLossRatio = .68m,
                TargetLossRatioBasis = "Benefits / deposits",
                CommissionArrangementId = includeTerms ? 600 + versionId : null,
                CommissionConfirmationStatus = includeTerms ? "Confirmed" : "",
                CommissionRate = includeTerms ? .05m : null,
                CommissionBasis = includeTerms ? "Total benefit payments" : ""
            };
            if (includeTerms)
            {
                version.Charges.Add(new FinancialArrangementCharge
                {
                    FinancialArrangementChargeId = 700 + versionId,
                    ChargeType = FinancialChargeType.GeneralAdministration,
                    Benefit = "EHC",
                    Rate = .065m,
                    Basis = "Total benefit payments"
                });
            }
            return version;
        }

        var firstVersion = Version(1, new DateOnly(2024, 1, 1), includeTerms: true);
        var repeatedVersion = Version(2, new DateOnly(2025, 1, 1), includeTerms: true);
        var removalVersion = Version(3, new DateOnly(2026, 1, 1), includeTerms: false);

        var events = await Financial([firstVersion, repeatedVersion, removalVersion], [], []).ContributeAsync(ClientId);

        var first = Assert.Single(events, item => item.SourceEntityId == "1");
        Assert.All(first.ChangeSet.Changes, item => Assert.Equal("", item.PreviousValue));
        Assert.Contains(first.ChangeSet.Changes, item => item.FieldKey == "TargetLossRatio" && item.NewValue == "68");
        Assert.Contains(first.ChangeSet.Changes, item => item.FieldKey == "CommissionRate" && item.NewValue == "5");
        var repeated = Assert.Single(events, item => item.SourceEntityId == "2");
        Assert.Empty(repeated.ChangeSet.Changes);
        var removal = Assert.Single(events, item => item.SourceEntityId == "3");
        Assert.Contains(removal.ChangeSet.Changes, item => item.FieldKey == "CommissionRate" && item.PreviousValue == "5" && item.NewValue == "");
        Assert.Contains(removal.ChangeSet.Changes, item => item.FieldKey.StartsWith("Charge:GeneralAdministration:EHC:", StringComparison.Ordinal) && item.PreviousValue == "6.5" && item.NewValue == "");
        Assert.DoesNotContain(removal.ChangeSet.Changes, item => item.FieldKey == "TargetLossRatio");
    }

    [Fact]
    public async Task Financial_Replay_Does_Not_Expose_A_Retroactive_Version_Before_It_Was_Recorded_And_Confirmed()
    {
        var version = new FinancialArrangementHistoryVersion
        {
            FinancialArrangementVersionId = 51,
            FinancialArrangementId = 510,
            ClientId = ClientId,
            Carrier = "Manulife",
            Scenario = FinancialArrangementScenario.Current,
            EffectiveDate = new DateOnly(2025, 1, 1),
            ArrangementStatus = "Confirmed",
            TargetLossRatio = .68m,
            CreatedAtUtc = new DateTime(2025, 2, 1, 0, 0, 0, DateTimeKind.Utc)
        };
        var audit = new FinancialArrangementAuditEvent
        {
            EntityType = "FinancialArrangement",
            EntityId = version.FinancialArrangementId,
            ClientId = ClientId,
            EventType = "FinancialArrangementConfirmed",
            EffectiveDate = version.EffectiveDate,
            OccurredAtUtc = new DateTime(2025, 2, 2, 0, 0, 0, DateTimeKind.Utc)
        };
        var contributor = Financial([version], [audit], []);

        var beforeConfirmation = await contributor.ContributeAsync(new AsOfQueryContext
        {
            ClientId = ClientId,
            AsOfDate = new DateOnly(2025, 1, 31)
        });
        var afterConfirmation = await contributor.ContributeAsync(new AsOfQueryContext
        {
            ClientId = ClientId,
            AsOfDate = new DateOnly(2025, 2, 2)
        });

        Assert.Empty(beforeConfirmation.Events);
        Assert.Single(afterConfirmation.Events);
        Assert.Equal(audit.OccurredAtUtc, afterConfirmation.Events[0].RecordedDate!.Value);
    }

    [Fact]
    public async Task Financial_Replay_Omits_Confirmed_Version_When_No_Canonical_Record_Time_Is_Available()
    {
        var version = new FinancialArrangementHistoryVersion
        {
            FinancialArrangementVersionId = 52,
            FinancialArrangementId = 520,
            ClientId = ClientId,
            Carrier = "Manulife",
            Scenario = FinancialArrangementScenario.Current,
            EffectiveDate = new DateOnly(2025, 1, 1),
            ArrangementStatus = "Confirmed",
            TargetLossRatio = .68m
        };
        var contributor = Financial([version], [], []);

        var replay = await contributor.ContributeAsync(new AsOfQueryContext
        {
            ClientId = ClientId,
            AsOfDate = new DateOnly(2025, 2, 1)
        });

        Assert.Empty(replay.Events);
    }

    [Fact]
    public async Task Financial_Replay_Uses_The_Latest_Complete_Aso_Version_Imported_By_AsOfDate()
    {
        var original = Statement(30, new DateOnly(2024, 12, 31));
        original.IsCurrent = false;
        original.ImportedAt = new DateTime(2025, 1, 10, 0, 0, 0, DateTimeKind.Utc);
        original.Total = new AsoFinancialActivityPeriod
        {
            PeriodId = 300,
            StatementId = original.StatementId,
            IsTotal = true,
            CurrentBalance = -100m,
            BalanceStatus = "Surplus"
        };
        var corrected = Statement(31, new DateOnly(2024, 12, 31));
        corrected.VersionNumber = 2;
        corrected.IsCurrent = true;
        corrected.Carrier = "Corrected carrier name";
        corrected.StatementDate = new DateOnly(2025, 1, 5);
        corrected.ImportedAt = new DateTime(2025, 2, 10, 0, 0, 0, DateTimeKind.Utc);
        corrected.Total = new AsoFinancialActivityPeriod
        {
            PeriodId = 310,
            StatementId = corrected.StatementId,
            IsTotal = true,
            CurrentBalance = 500m,
            BalanceStatus = "Deficit"
        };
        var contributor = Financial([], [], [original, corrected]);

        var current = await contributor.ContributeAsync(ClientId);
        var replay = await contributor.ContributeAsync(new AsOfQueryContext
        {
            ClientId = ClientId,
            AsOfDate = new DateOnly(2025, 1, 31)
        });
        var replayAfterCorrection = await contributor.ContributeAsync(new AsOfQueryContext
        {
            ClientId = ClientId,
            AsOfDate = new DateOnly(2025, 2, 28)
        });

        var currentFunding = Assert.Single(current, item => item.EventType == "ASO funding position");
        Assert.Contains(currentFunding.ChangeSet.Changes, item => item.FieldKey == "CurrentBalance" && item.NewValue == "500");
        var replayFunding = Assert.Single(replay.Events, item => item.EventType == "ASO funding position");
        Assert.Contains(replayFunding.ChangeSet.Changes, item => item.FieldKey == "CurrentBalance" && item.NewValue == "-100");
        Assert.DoesNotContain(replay.Events, item => item.SourceEntityId == "31" || item.RelatedRecords.Contains("AsoFinancialStatement:31"));
        Assert.Equal(HistoryDatePrecision.ExactDate, replayFunding.EffectiveDateReference.Precision);
        var correctedFunding = Assert.Single(replayAfterCorrection.Events, item => item.EventType == "ASO funding position");
        Assert.Contains(correctedFunding.ChangeSet.Changes, item => item.FieldKey == "CurrentBalance" && item.NewValue == "500");
        Assert.DoesNotContain(replayAfterCorrection.Events, item => item.SourceEntityId == "30" || item.RelatedRecords.Contains("AsoFinancialStatement:30"));
        Assert.Equal(replayFunding.StableSourceKey, correctedFunding.StableSourceKey);
    }

    [Fact]
    public async Task Financial_Rejects_Unconfirmed_And_Inactive_Data_And_Preserves_Month_Precision()
    {
        var unconfirmed = new FinancialArrangementHistoryVersion
        {
            FinancialArrangementVersionId = 1,
            FinancialArrangementId = 11,
            ClientId = ClientId,
            Carrier = "Manulife",
            Scenario = FinancialArrangementScenario.Current,
            EffectiveDate = new DateOnly(2024, 1, 1),
            ArrangementStatus = "Extracted"
        };
        var old = Statement(8, null);
        old.IsCurrent = false;
        var monthOnly = Statement(9, null);
        monthOnly.Periods.Add(new AsoFinancialActivityPeriod { PeriodId = 90, StatementId = 9, ActivityMonth = new DateOnly(2025, 7, 1) });
        var contributor = Financial([unconfirmed], [], [old, monthOnly]);

        var events = await contributor.ContributeAsync(ClientId);

        var statement = Assert.Single(events);
        Assert.Equal(HistoryDatePrecision.MonthAndYear, statement.EffectiveDateReference.Precision);
        Assert.Null(statement.EffectiveDate);
        Assert.Equal(7, statement.EffectiveDateReference.Month);
        Assert.DoesNotContain(events, item => item.SourceEntityId == "1" || item.SourceEntityId == "8");
    }

    [Fact]
    public async Task Financial_Distinguishes_Month_Only_Statements_And_Uses_Natural_Stable_Identity_Across_Reimports()
    {
        var july = Statement(20, null);
        july.Periods.Add(new AsoFinancialActivityPeriod
        {
            PeriodId = 200,
            StatementId = july.StatementId,
            ActivityMonth = new DateOnly(2025, 7, 1)
        });
        var correctedJuly = Statement(21, null);
        correctedJuly.VersionNumber = 2;
        correctedJuly.ImportedAt = july.ImportedAt.AddDays(1);
        correctedJuly.Periods.Add(new AsoFinancialActivityPeriod
        {
            PeriodId = 210,
            StatementId = correctedJuly.StatementId,
            ActivityMonth = new DateOnly(2025, 7, 1)
        });
        var august = Statement(22, null);
        august.Periods.Add(new AsoFinancialActivityPeriod
        {
            PeriodId = 220,
            StatementId = august.StatementId,
            ActivityMonth = new DateOnly(2025, 8, 1)
        });

        var original = Assert.Single(await Financial([], [], [july]).ContributeAsync(ClientId));
        var correction = Assert.Single(await Financial([], [], [correctedJuly]).ContributeAsync(ClientId));
        var combined = await Financial([], [], [july, correctedJuly, august]).ContributeAsync(ClientId);

        Assert.Equal(original.StableSourceKey, correction.StableSourceKey);
        Assert.Equal(2, combined.Count);
        Assert.Contains(combined, item => item.SourceEntityId == correctedJuly.StatementId.ToString());
        Assert.Contains(combined, item => item.EffectiveDateReference.Month == 7);
        Assert.Contains(combined, item => item.EffectiveDateReference.Month == 8);
        Assert.Equal(2, combined.Select(item => item.StableSourceKey).Distinct(StringComparer.Ordinal).Count());
    }

    [Fact]
    public async Task Proposal_Maps_Only_Factual_Received_Date_With_Source_Provenance()
    {
        var profile = new ProposalCanonicalProfile { ClientId = ClientId };
        profile.Proposals.Add(Proposal(100, "2026-07-01", new DateOnly(2026, 7, 1), "Manulife"));
        var contributor = Proposal(profile);

        var first = await contributor.ContributeAsync(ClientId);
        var second = await contributor.ContributeAsync(ClientId);

        var historyEvent = Assert.Single(first);
        Assert.Equal(first[0].StableSourceKey, second[0].StableSourceKey);
        Assert.Equal("Proposal received", historyEvent.EventType);
        Assert.DoesNotContain("selected", historyEvent.Title, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("recommended", historyEvent.Title, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("Saved", historyEvent.Summary, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("source status", historyEvent.Summary, StringComparison.OrdinalIgnoreCase);
        var evidence = Assert.Single(historyEvent.Evidence);
        Assert.Equal(3, evidence.PageNumber);
        Assert.Equal("proposal_summary:100", evidence.SourceFactId);
        Assert.Contains("Carrier:77", historyEvent.RelatedRecords);
    }

    [Fact]
    public async Task Proposal_Year_And_Month_Dates_Keep_Partial_Precision_And_Review_Rows_Are_Omitted()
    {
        var profile = new ProposalCanonicalProfile { ClientId = ClientId };
        profile.Proposals.Add(Proposal(100, "2025", null, "Manulife"));
        profile.Proposals.Add(Proposal(101, "2026-07", new DateOnly(2026, 7, 1), "Sun Life"));
        var review = Proposal(102, "2026-08-01", new DateOnly(2026, 8, 1), "Canada Life");
        review.Exceptions.Add(new ProposalCanonicalException { ExceptionType = "Needs review" });
        profile.Proposals.Add(review);
        var contributor = Proposal(profile);

        var events = await contributor.ContributeAsync(ClientId);

        Assert.Equal(2, events.Count);
        Assert.Equal(HistoryDatePrecision.YearOnly, events.Single(item => item.SourceEntityId == "100").EffectiveDateReference.Precision);
        var month = events.Single(item => item.SourceEntityId == "101");
        Assert.Equal(HistoryDatePrecision.MonthAndYear, month.EffectiveDateReference.Precision);
        Assert.Null(month.EffectiveDate);
        Assert.DoesNotContain(events, item => item.SourceEntityId == "102");
    }

    [Fact]
    public async Task Claims_Emits_One_Validated_Period_Milestone_And_Merges_Report_Version_Evidence()
    {
        var firstReport = ClaimsReport(1, 700, new DateTime(2025, 1, 2, 0, 0, 0, DateTimeKind.Utc));
        var secondReport = ClaimsReport(2, 701, new DateTime(2025, 1, 3, 0, 0, 0, DateTimeKind.Utc));
        firstReport.MonthlyFacts.Add(new CanonicalClaimMonthlyFact { FactId = "month-1", Month = new DateOnly(2024, 12, 1), Availability = CanonicalDataAvailability.Present });
        secondReport.MonthlyFacts.Add(new CanonicalClaimMonthlyFact { FactId = "month-2", Month = new DateOnly(2024, 12, 1), Availability = CanonicalDataAvailability.Present });
        var contributor = Claims([firstReport, secondReport]);

        var first = await contributor.ContributeAsync(ClientId);
        var second = await contributor.ContributeAsync(ClientId);

        var milestone = Assert.Single(first);
        Assert.Equal(first[0].StableSourceKey, second[0].StableSourceKey);
        Assert.Equal("Claims experience period", milestone.EventType);
        Assert.Equal(2, milestone.Evidence.Count);
        Assert.Contains("CanonicalClaimsReport:1", milestone.RelatedRecords);
        Assert.Contains("CanonicalClaimsReport:2", milestone.RelatedRecords);
        Assert.DoesNotContain("trend", milestone.Summary, StringComparison.OrdinalIgnoreCase);
        var status = await contributor.ContributeAsync(new AsOfQueryContext { ClientId = ClientId, AsOfDate = new DateOnly(2025, 1, 1) });
        Assert.Equal(HistoricalQuerySupport.Limited, status.HistoricalSupport.Support);
    }

    [Fact]
    public async Task Claims_Omits_Review_Undated_And_Empty_Reports_Without_Monthly_Noise()
    {
        var review = ClaimsReport(1, 700, DateTime.UtcNow);
        review = CopyClaimsReport(review, validationStatus: CanonicalClaimsValidationStatus.NeedsReview.ToString());
        var undated = CopyClaimsReport(ClaimsReport(2, 701, DateTime.UtcNow), clearEnd: true);
        var empty = CopyClaimsReport(ClaimsReport(3, 702, DateTime.UtcNow), includeFact: false);
        var contributor = Claims([review, undated, empty]);

        var result = await contributor.ContributeAsync(new AsOfQueryContext { ClientId = ClientId, AsOfDate = new DateOnly(2025, 1, 1) });

        Assert.Empty(result.Events);
        Assert.Equal(HistoricalQuerySupport.Limited, result.HistoricalSupport.Support);
    }

    [Fact]
    public async Task Claims_Omits_Valid_Report_When_Its_Only_Present_Fact_Is_Not_Validated()
    {
        var report = ClaimsReport(4, 703, DateTime.UtcNow);
        report.FinancialFacts.Clear();
        report.FinancialFacts.Add(new CanonicalClaimFinancialFact
        {
            FactId = "financial-4-review",
            Benefit = "EHC",
            PaidClaims = 100m,
            Availability = CanonicalDataAvailability.Present,
            ValidationStatus = CanonicalClaimsValidationStatus.NeedsReview
        });
        var contributor = Claims([report]);

        var events = await contributor.ContributeAsync(ClientId);

        Assert.Empty(events);
    }

    [Fact]
    public async Task Canonical_AsOf_Projection_Truncates_Future_Contributor_Facts()
    {
        var carrier = Administration(10, "Manulife", 100, new DateTime(2026, 1, 1));
        var plan = Rule(1, new DateOnly(2026, 1, 1), 80m, 500m, 100m);
        var proposalProfile = new ProposalCanonicalProfile { ClientId = ClientId };
        proposalProfile.Proposals.Add(Proposal(100, "2026-01-01", new DateOnly(2026, 1, 1), "Manulife"));
        var claim = CopyClaimsReport(ClaimsReport(1, 700, DateTime.UtcNow), end: new DateOnly(2026, 1, 1));
        var financialVersion = new FinancialArrangementHistoryVersion
        {
            FinancialArrangementVersionId = 1,
            FinancialArrangementId = 11,
            ClientId = ClientId,
            Carrier = "Manulife",
            Scenario = FinancialArrangementScenario.Current,
            EffectiveDate = new DateOnly(2024, 1, 1),
            ArrangementStatus = "Confirmed",
            CreatedAtUtc = new DateTime(2024, 1, 2, 0, 0, 0, DateTimeKind.Utc)
        };
        var builder = new ClientHistoryProjectionBuilder([
            Carrier([carrier]),
            Plan([plan]),
            Financial([financialVersion], [], []),
            Proposal(proposalProfile),
            Claims([claim])
        ]);

        var result = await builder.BuildAsOfAsync(new AsOfQueryContext
        {
            ClientId = ClientId,
            AsOfDate = new DateOnly(2025, 12, 31),
            IncludeManualEvents = true
        });

        var only = Assert.Single(result.Events);
        Assert.Equal("Financial", only.SourceModule);
        Assert.Equal(5, result.ContributorSupport.Count);
        Assert.All(result.ContributorSupport.Where(item => item.SourceModule is "PlanDesign" or "Financial" or "ProposalMarket" or "Claims"),
            item => Assert.Equal(HistoricalQuerySupport.Limited, item.Support));
    }

    private static CarrierVendorHistoryContributor Carrier(IReadOnlyList<ClientBenefitAdministrationRecord> records) =>
        new((_, _) => Task.FromResult(records));

    private static PlanDesignHistoryContributor Plan(IReadOnlyList<PlanBenefitRule> rules) =>
        new((_, _) => Task.FromResult(rules));

    private static FinancialHistoryContributor Financial(
        IReadOnlyList<FinancialArrangementHistoryVersion> versions,
        IReadOnlyList<FinancialArrangementAuditEvent> audits,
        IReadOnlyList<AsoFinancialActivityStatement> statements) =>
        new((_, _) => Task.FromResult(new FinancialHistorySourceData(versions, audits, statements)));

    private static ProposalMarketHistoryContributor Proposal(ProposalCanonicalProfile profile) =>
        new((_, _) => Task.FromResult(profile));

    private static ClaimsHistoryContributor Claims(IReadOnlyList<CanonicalClaimsReportReadModel> reports) =>
        new((_, _) => Task.FromResult(reports));

    private static ClientBenefitAdministrationRecord Administration(long id, string carrier, long organizationId, DateTime? effective) => new()
    {
        ClientBenefitAdministrationId = id,
        ClientId = ClientId,
        BenefitTypeId = 1,
        BenefitTypeCode = "EHC",
        BenefitDisplayName = "Extended Health Care",
        InsurerOrganizationId = organizationId,
        InsurerName = carrier,
        InsurerOrganizationTypeCode = BenefitAdministrationOrganizationTypes.Insurer,
        EffectiveDate = effective,
        SourceDocumentId = 900,
        SourcePageNumber = 7,
        SourceText = $"Insurer: {carrier}",
        ExtractionConfidence = .95m,
        ReviewStatusCode = BenefitAdministrationReviewStatuses.Confirmed,
        CreatedAtUtc = new DateTime(2021, 1, 2, 0, 0, 0, DateTimeKind.Utc),
        UpdatedAtUtc = new DateTime(2021, 1, 3, 0, 0, 0, DateTimeKind.Utc)
    };

    private static PlanBenefitRule Rule(long id, DateOnly? effective, decimal tier1, decimal threshold, decimal tier2) => new()
    {
        PlanBenefitRuleId = id,
        PlanDesignFeatureId = 20,
        ClientId = ClientId,
        SourceDocumentId = 501,
        EffectiveDate = effective,
        Tier1ReimbursementPercent = tier1,
        Tier1ThresholdAmount = threshold,
        Tier2ReimbursementPercent = tier2,
        ThresholdBasis = BenefitRuleThresholdBasis.Person,
        ResetPeriod = BenefitRuleResetPeriod.CalendarYear,
        CurrencyCode = "CAD",
        SourceText = $"{tier1}% to {threshold}; {tier2}% thereafter",
        SourceLocation = "Page 6",
        NormalizationStatus = BenefitRuleNormalizationStatus.Normalized,
        ConfidenceScore = .94m,
        DisplayOrder = 10,
        IsActive = true,
        CreatedAtUtc = new DateTime(2021, 1, 2, 0, 0, 0, DateTimeKind.Utc),
        UpdatedAtUtc = new DateTime(2021, 1, 3, 0, 0, 0, DateTimeKind.Utc)
    };

    private static void AssertInitial(
        HistoryChangeItem item,
        string key,
        HistoryValueType valueType,
        string newValue,
        string unit = "")
    {
        Assert.Equal(key, item.FieldKey);
        Assert.Equal(valueType, item.ValueType);
        Assert.Equal("", item.PreviousValue);
        Assert.Equal(newValue, item.NewValue);
        Assert.Equal(unit, item.Unit);
    }

    private static AsoFinancialActivityStatement Statement(long id, DateOnly? periodEnd) => new()
    {
        StatementId = id,
        ClientId = ClientId,
        SourceDocumentId = 800 + id,
        Carrier = "Manulife",
        ContractNumber = "12345",
        PeriodStart = periodEnd?.AddMonths(-11),
        PeriodEnd = periodEnd,
        SourceFileName = $"statement-{id}.xlsx",
        SourceSheet = "Activity",
        SourceHash = $"hash-{id}",
        VersionNumber = 1,
        IsCurrent = true,
        ImportStatus = AsoFinancialImportStatuses.Complete,
        ImportedAt = new DateTime(2025, 1, 2, 0, 0, 0, DateTimeKind.Utc)
    };

    private static ProposalCanonicalQuote Proposal(long id, string rawDate, DateOnly? date, string carrier)
    {
        var quote = new ProposalCanonicalQuote
        {
            ProposalId = id,
            CarrierId = 77,
            CarrierName = carrier,
            ProposalDate = date,
            ProposalDateSourceText = rawDate,
            EffectiveDate = date?.AddMonths(1),
            Status = "Saved",
            ExtractionConfidence = .90m,
            CreatedAt = new DateTime(2026, 8, 1),
            Evidence = new ProposalCanonicalEvidence
            {
                SourceTable = "proposal_summary",
                SourceId = id,
                SourceFileName = "proposal.pdf",
                SourcePage = 3,
                SourceSheet = "Summary",
                ConfidenceScore = .90m,
                RawPayloadJson = "{\"proposal_date\":\"" + rawDate + "\"}"
            }
        };
        quote.BenefitLines.Add(new ProposalCanonicalBenefitLine { BenefitLine = "EHC", MonthlyPremium = 1000m });
        return quote;
    }

    private static CanonicalClaimsReportReadModel ClaimsReport(long id, long documentId, DateTime updatedAt)
    {
        var report = new CanonicalClaimsReportReadModel
        {
            CanonicalReportId = id,
            ClientId = ClientId,
            Carrier = "Manulife",
            PolicyNumber = "12345",
            SourceDocumentId = documentId,
            SourceDocumentVersionId = documentId + 100,
            Period = new CanonicalClaimsPeriod
            {
                Start = new DateOnly(2024, 1, 1),
                End = new DateOnly(2024, 12, 31),
                PeriodType = "Experience"
            },
            ValidationStatus = CanonicalClaimsValidationStatus.Valid.ToString(),
            CreatedAt = updatedAt.AddDays(-1),
            UpdatedAt = updatedAt
        };
        report.FinancialFacts.Add(new CanonicalClaimFinancialFact { FactId = $"financial-{id}", Benefit = "EHC", PaidClaims = 100m, Availability = CanonicalDataAvailability.Present });
        return report;
    }

    private static CanonicalClaimsReportReadModel CopyClaimsReport(
        CanonicalClaimsReportReadModel source,
        string? validationStatus = null,
        DateOnly? end = default,
        bool clearEnd = false,
        bool includeFact = true)
    {
        var resolvedEnd = clearEnd ? null : end ?? source.Period.End;
        var copy = new CanonicalClaimsReportReadModel
        {
            CanonicalReportId = source.CanonicalReportId,
            ClientId = source.ClientId,
            Carrier = source.Carrier,
            PolicyNumber = source.PolicyNumber,
            SourceDocumentId = source.SourceDocumentId,
            SourceDocumentVersionId = source.SourceDocumentVersionId,
            Period = new CanonicalClaimsPeriod { Start = source.Period.Start, End = resolvedEnd, PeriodType = source.Period.PeriodType },
            ValidationStatus = validationStatus ?? source.ValidationStatus,
            CreatedAt = source.CreatedAt,
            UpdatedAt = source.UpdatedAt
        };
        if (includeFact) copy.FinancialFacts.Add(new CanonicalClaimFinancialFact { FactId = $"financial-{copy.CanonicalReportId}", Availability = CanonicalDataAvailability.Present, PaidClaims = 100m });
        return copy;
    }
}

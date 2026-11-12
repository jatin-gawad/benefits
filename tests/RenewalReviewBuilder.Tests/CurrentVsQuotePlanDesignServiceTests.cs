using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class CurrentVsQuotePlanDesignServiceTests
{
    [Fact]
    public void Build_MatchesEquivalentCurrentAndQuoteValues()
    {
        var profile = CreateProfile();
        profile.Facts.Add(CreateFact(
            PlanDesignCanonicalScopes.CurrentPlan,
            "Extended Health Care",
            "Coinsurance",
            "80 percent",
            sourceFileName: "current-booklet.pdf",
            sourcePage: 12));
        profile.Facts.Add(CreateFact(
            PlanDesignCanonicalScopes.Quote,
            "Extended Health Care",
            "Coinsurance",
            "80%",
            proposalId: 101,
            carrier: "Manulife",
            sourceFileName: "manulife-proposal.pdf",
            sourcePage: 5));
        profile.Facts.Add(CreateFact(
            PlanDesignCanonicalScopes.ProposalValidation,
            "Extended Health Care",
            "Coinsurance",
            "80%",
            proposalId: 101,
            carrier: "Manulife",
            validationStatus: "match"));

        var result = new CurrentVsQuotePlanDesignService().Build(profile);

        var quote = Assert.Single(result.Quotes);
        var row = Assert.Single(result.Rows);
        var value = Assert.Single(row.QuoteValues);
        Assert.Equal("Manulife", quote.DisplayName);
        Assert.Equal(CurrentVsQuotePlanDesignStatuses.Same, value.Status);
        Assert.Equal(CurrentVsQuotePlanDesignStatuses.Same, row.Status);
        Assert.Equal("current-booklet.pdf", row.CurrentEvidence?.SourceFileName);
        Assert.Equal("manulife-proposal.pdf", value.Evidence?.SourceFileName);
    }

    [Fact]
    public void Build_ReportsDifferentWithoutGuessingBetterOrWorse()
    {
        var profile = CreateProfile();
        profile.Facts.Add(CreateFact(
            PlanDesignCanonicalScopes.CurrentPlan,
            "Dental",
            "Coinsurance",
            "80%"));
        profile.Facts.Add(CreateFact(
            PlanDesignCanonicalScopes.Quote,
            "Dental",
            "Coinsurance",
            "100%",
            proposalId: 201,
            carrier: "Canada Life"));

        var result = new CurrentVsQuotePlanDesignService().Build(profile);

        var row = Assert.Single(result.Rows);
        var value = Assert.Single(row.QuoteValues);
        Assert.Equal(CurrentVsQuotePlanDesignStatuses.Different, value.Status);
        Assert.Equal(CurrentVsQuotePlanDesignStatuses.Different, row.Status);
        Assert.DoesNotContain("Better", value.Status, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("Worse", value.Status, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Build_MatchesCanonicalDentalProvisionAliasesWithoutCombiningRates()
    {
        var profile = CreateProfile();
        profile.Facts.Add(CreateFact(
            PlanDesignCanonicalScopes.CurrentPlan,
            "Dental Basic Services",
            "Dental Basic Services",
            "80%"));
        profile.Facts.Add(CreateFact(
            PlanDesignCanonicalScopes.Quote,
            "Dental Care",
            "Basic",
            "80%",
            proposalId: 211,
            carrier: "Manulife"));

        var result = new CurrentVsQuotePlanDesignService().Build(profile);

        var row = Assert.Single(result.Rows);
        Assert.Equal("Basic Services", row.ProvisionName);
        Assert.Equal(CurrentVsQuotePlanDesignStatuses.Same, Assert.Single(row.QuoteValues).Status);
        Assert.Equal(
            profile.CurrentPlanFacts[0].ComparisonKey,
            profile.QuoteFacts[0].ComparisonKey);
    }

    [Fact]
    public void Build_PreservesMissingSidesAcrossMultipleSavedQuotes()
    {
        var profile = CreateProfile();
        profile.Facts.Add(CreateFact(
            PlanDesignCanonicalScopes.CurrentPlan,
            "Extended Health Care",
            "Coinsurance",
            "80%"));
        profile.Facts.Add(CreateFact(
            PlanDesignCanonicalScopes.Quote,
            "Extended Health Care",
            "Coinsurance",
            "80%",
            proposalId: 301,
            carrier: "Manulife"));
        profile.Facts.Add(CreateFact(
            PlanDesignCanonicalScopes.Quote,
            "Extended Health Care",
            "Maximum",
            "$1,000",
            proposalId: 302,
            carrier: "RBC"));

        var result = new CurrentVsQuotePlanDesignService().Build(profile);

        Assert.Equal(2, result.Quotes.Count);
        Assert.Equal(2, result.Rows.Count);

        var coinsurance = Assert.Single(result.Rows, row => row.ProvisionName == "Coinsurance");
        Assert.Equal(CurrentVsQuotePlanDesignStatuses.CurrentOnly, coinsurance.Status);
        Assert.Contains(coinsurance.QuoteValues, value => value.Carrier == "Manulife" && value.Status == CurrentVsQuotePlanDesignStatuses.Same);
        Assert.Contains(coinsurance.QuoteValues, value => value.Carrier == "RBC" && value.Status == CurrentVsQuotePlanDesignStatuses.CurrentOnly);

        var maximum = Assert.Single(result.Rows, row => row.ProvisionName == "Maximum");
        Assert.Equal("Not found", maximum.CurrentValue);
        Assert.Equal(CurrentVsQuotePlanDesignStatuses.QuoteOnly, maximum.Status);
        Assert.Contains(maximum.QuoteValues, value => value.Carrier == "RBC" && value.Status == CurrentVsQuotePlanDesignStatuses.QuoteOnly);
    }

    [Fact]
    public void Build_RequiresReviewWhenProposalValidationFails()
    {
        var profile = CreateProfile();
        profile.Facts.Add(CreateFact(
            PlanDesignCanonicalScopes.CurrentPlan,
            "Long Term Disability",
            "Elimination Period",
            "119 days"));
        profile.Facts.Add(CreateFact(
            PlanDesignCanonicalScopes.Quote,
            "Long Term Disability",
            "Elimination Period",
            "112 days",
            proposalId: 401,
            carrier: "Sun Life"));
        profile.Facts.Add(CreateFact(
            PlanDesignCanonicalScopes.ProposalValidation,
            "Long Term Disability",
            "Elimination Period",
            "112 days",
            proposalId: 401,
            carrier: "Sun Life",
            validationStatus: "mismatch"));

        var result = new CurrentVsQuotePlanDesignService().Build(profile);

        var row = Assert.Single(result.Rows);
        Assert.Equal(CurrentVsQuotePlanDesignStatuses.NeedsReview, row.Status);
        Assert.Equal(CurrentVsQuotePlanDesignStatuses.NeedsReview, Assert.Single(row.QuoteValues).Status);
    }

    [Fact]
    public void Build_CurrentOnlyProfileRemainsUsableButNotComparable()
    {
        var profile = CreateProfile();
        profile.Facts.Add(CreateFact(
            PlanDesignCanonicalScopes.CurrentPlan,
            "Basic Life",
            "Benefit Amount",
            "2x annual salary"));
        var factCount = profile.Facts.Count;

        var result = new CurrentVsQuotePlanDesignService().Build(profile);

        Assert.False(result.CanCompare);
        Assert.Empty(result.Quotes);
        Assert.Equal(CurrentVsQuotePlanDesignStatuses.CurrentOnly, Assert.Single(result.Rows).Status);
        Assert.Equal(factCount, profile.Facts.Count);
    }

    private static PlanDesignCanonicalProfile CreateProfile()
    {
        return new PlanDesignCanonicalProfile
        {
            ClientId = 42,
            ClientName = "Canonical Comparison Ltd.",
            ClientNumber = "CC-42",
            Carrier = "Current Carrier",
            BuiltAtUtc = new DateTime(2026, 7, 18, 12, 0, 0, DateTimeKind.Utc)
        };
    }

    private static PlanDesignCanonicalFact CreateFact(
        string scope,
        string benefit,
        string provision,
        string value,
        long? proposalId = null,
        string carrier = "",
        string sourceFileName = "source.pdf",
        int? sourcePage = null,
        string validationStatus = "")
    {
        var fact = new PlanDesignCanonicalFact
        {
            Scope = scope,
            ClientId = 42,
            ProposalId = proposalId,
            Carrier = carrier,
            BenefitCategory = benefit,
            BenefitName = benefit,
            ClassName = "All Classes",
            FeatureName = provision,
            FeatureValue = value,
            ValidationStatus = validationStatus,
            SourceTable = scope switch
            {
                PlanDesignCanonicalScopes.CurrentPlan => "plan_design_features",
                PlanDesignCanonicalScopes.ProposalValidation => "proposal_booklet_validation",
                _ => "proposal_plan_design"
            },
            SourceFileName = sourceFileName,
            SourceDocument = sourceFileName,
            SourcePage = sourcePage,
            ConfidenceScore = 0.95m
        };
        PlanDesignCanonicalMapper.CompleteFact(fact);
        return fact;
    }
}

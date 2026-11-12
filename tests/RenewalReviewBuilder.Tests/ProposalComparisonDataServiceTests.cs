using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class ProposalComparisonDataServiceTests
{
    [Fact]
    public void Build_MapsCanonicalProposalAnalysisAndPreservesCurrentVsQuoteRows()
    {
        var analysis = CreateAnalysis();
        var currentVsQuote = CreateCurrentVsQuote();

        var model = new ProposalComparisonDataService().Build(analysis, currentVsQuote);

        Assert.False(model.IsEmpty);
        Assert.Equal("Canonical Client Ltd.", model.ClientName);
        Assert.Equal("Canada Life", model.CurrentCarrier);
        Assert.Equal("2 saved proposals", model.QuoteCountText);
        Assert.Equal("$12,000", model.LowestAnnualPremiumText);
        Assert.Equal("$2,400 (20%)", model.AnnualPremiumSpreadText);
        Assert.Equal(ProposalCanonicalReadinessStatuses.Ready, model.ReadinessStatus);
        Assert.Equal(ProposalIntelligenceAnalysisVersions.Current, model.AnalysisVersion);

        Assert.Equal(2, model.ProposalSummaries.Count);
        var manulife = model.ProposalSummaries.Single(row => row.CarrierName == "Manulife");
        Assert.Equal("$12,000", manulife.AnnualPremiumText);
        Assert.Equal("Lowest saved premium", manulife.LowestPremiumLabel);
        Assert.Equal("Validated", manulife.ValidationStatusText);

        var premium = Assert.Single(model.PremiumComparisons);
        Assert.Equal("Extended Health Care", premium.Subject);
        Assert.Equal(ProposalIntelligenceComparisonStatuses.Different, premium.ComparisonStatus);
        Assert.Equal(2, premium.Values.Count);
        Assert.Contains(premium.Values, value => value.CarrierName.Contains("Manulife", StringComparison.Ordinal) && value.Value.Contains("$12,000", StringComparison.Ordinal));

        var rate = Assert.Single(model.RateComparisons);
        Assert.Equal("1.25", rate.Values.Single(value => value.ProposalId == 100).Value);
        Assert.Contains("Per $1,000", rate.Detail, StringComparison.Ordinal);

        var carrierPlanDesign = Assert.Single(model.CarrierPlanDesignComparisons);
        Assert.Equal("Coinsurance", carrierPlanDesign.Detail);
        Assert.Contains(carrierPlanDesign.Values, value => value.Value == "80%" && value.Detail.Contains("Booklet: 80%", StringComparison.Ordinal));

        var underwriting = Assert.Single(model.UnderwritingComparisons);
        Assert.Equal("Rate Guarantee", underwriting.Subject);
        Assert.Equal(2, underwriting.Values.Count);

        var currentPlan = Assert.Single(model.BenefitComparisons);
        Assert.Equal("80%", currentPlan.CurrentValue);
        Assert.Equal("90%", Assert.Single(currentPlan.QuoteValues).Value);
        Assert.Equal(CurrentVsQuotePlanDesignStatuses.Different, currentPlan.ComparisonStatus);
        Assert.Contains(model.Issues, issue => issue.Code == "annual_premium_spread");
        Assert.Contains(model.Issues, issue => issue.Code == "current_plan_review");
        Assert.Equal("2 findings", model.ReviewItemCountText);
        Assert.Contains("canonical proposal analysis", model.StatusText, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Build_CurrentPlanOnlyRemainsVisibleWhileProposalAnalysisIsNotReady()
    {
        var analysis = new ProposalIntelligenceAnalysis
        {
            ClientId = 42,
            ClientName = "Canonical Client Ltd.",
            ReadinessStatus = ProposalCanonicalReadinessStatuses.NotReady,
            ReadinessMessage = "No usable saved proposal facts were found for this client."
        };
        var currentVsQuote = CreateCurrentVsQuote();

        var model = new ProposalComparisonDataService().Build(analysis, currentVsQuote);

        Assert.False(model.IsEmpty);
        Assert.Empty(model.ProposalSummaries);
        Assert.Single(model.BenefitComparisons);
        Assert.Equal(ProposalCanonicalReadinessStatuses.NotReady, model.ReadinessStatus);
        Assert.Contains("no saved carrier proposals", model.StatusText, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Build_NoProposalOrCurrentPlanDataUsesEmptyState()
    {
        var analysis = new ProposalIntelligenceAnalysis
        {
            ClientId = 42,
            ClientName = "Canonical Client Ltd.",
            ReadinessStatus = ProposalCanonicalReadinessStatuses.NotReady
        };
        var currentVsQuote = new CurrentVsQuotePlanDesignComparison
        {
            ClientId = 42,
            ClientName = "Canonical Client Ltd."
        };

        var model = new ProposalComparisonDataService().Build(analysis, currentVsQuote);

        Assert.True(model.IsEmpty);
        Assert.Empty(model.ProposalSummaries);
        Assert.Empty(model.BenefitComparisons);
        Assert.Contains("No saved proposal", model.StatusText, StringComparison.Ordinal);
    }

    private static ProposalIntelligenceAnalysis CreateAnalysis()
    {
        var analysis = new ProposalIntelligenceAnalysis
        {
            ClientId = 42,
            ClientName = "Canonical Client Ltd.",
            ClientNumber = "CAN-42",
            ReadinessStatus = ProposalCanonicalReadinessStatuses.Ready,
            ReadinessMessage = "Saved carrier proposals are ready for factual comparison.",
            CanCompareCarriers = true,
            CanCompareFinancials = true,
            CanComparePlanDesign = true,
            BuiltAtUtc = new DateTime(2026, 7, 18, 12, 0, 0, DateTimeKind.Utc)
        };
        analysis.Quotes.Add(CreateQuote(100, "Manulife", 12_000m, 0m, 0m, true));
        analysis.Quotes.Add(CreateQuote(200, "RBC", 14_400m, 2_400m, 20m, false));

        var premium = new ProposalIntelligencePremiumComparisonRow
        {
            ComparisonKey = "premium:ehc",
            BenefitLine = "Extended Health Care",
            ClassName = "All Classes",
            Tier = "Single",
            ComparisonStatus = ProposalIntelligenceComparisonStatuses.Different,
            ExpectedProposalCount = 2,
            MissingProposalCount = 0,
            HasPremiumDifference = true,
            HasCostShareDifference = true
        };
        premium.Values.Add(CreatePremiumValue(100, "Manulife | #100", 12_000m));
        premium.Values.Add(CreatePremiumValue(200, "RBC | #200", 14_400m));
        analysis.PremiumComparisons.Add(premium);

        var rate = new ProposalIntelligenceRateComparisonRow
        {
            ComparisonKey = "rate:ehc",
            BenefitLine = "Extended Health Care",
            ClassName = "All Classes",
            Tier = "Single",
            RateBasis = "Per $1,000",
            ComparisonStatus = ProposalIntelligenceComparisonStatuses.Different,
            ExpectedProposalCount = 2,
            MissingProposalCount = 0,
            HasRateDifference = true
        };
        rate.Values.Add(CreateRateValue(100, "Manulife | #100", 1.25m, 12_000m));
        rate.Values.Add(CreateRateValue(200, "RBC | #200", 1.50m, 14_400m));
        analysis.RateComparisons.Add(rate);

        var planDesign = new ProposalIntelligencePlanDesignComparisonRow
        {
            ComparisonKey = "plan:ehc:coinsurance",
            BenefitLine = "Extended Health Care",
            ProvisionName = "Coinsurance",
            ComparisonStatus = ProposalIntelligenceComparisonStatuses.Different,
            ExpectedProposalCount = 2,
            HasProposalDifference = true
        };
        planDesign.Values.Add(CreatePlanDesignValue(100, "Manulife | #100", "80%"));
        planDesign.Values.Add(CreatePlanDesignValue(200, "RBC | #200", "90%"));
        analysis.PlanDesignComparisons.Add(planDesign);

        var term = new ProposalIntelligenceTermComparisonRow
        {
            ComparisonKey = "term:rate-guarantee",
            Category = "Underwriting Assumption",
            BenefitLine = "All Benefits",
            TermName = "Rate Guarantee",
            ComparisonStatus = ProposalIntelligenceComparisonStatuses.Different,
            ExpectedProposalCount = 2,
            HasDifference = true
        };
        term.Values.Add(CreateTermValue(100, "Manulife | #100", "24 months"));
        term.Values.Add(CreateTermValue(200, "RBC | #200", "12 months"));
        analysis.UnderwritingComparisons.Add(term);

        analysis.Findings.Add(new ProposalIntelligenceFinding
        {
            Code = "annual_premium_spread",
            Severity = ProposalCanonicalIssueSeverities.Information,
            Category = ProposalIntelligenceFindingCategories.Financial,
            Message = "The saved annual proposal premium spread is $2,400."
        });
        return analysis;
    }

    private static ProposalIntelligenceQuoteAnalysis CreateQuote(
        long proposalId,
        string carrier,
        decimal annualPremium,
        decimal difference,
        decimal percent,
        bool isLowest)
    {
        return new ProposalIntelligenceQuoteAnalysis
        {
            QuoteKey = $"proposal:{proposalId}",
            ProposalId = proposalId,
            CarrierName = carrier,
            DisplayName = $"{carrier} | #{proposalId}",
            AnnualPremium = annualPremium,
            MonthlyPremium = annualPremium / 12m,
            AnnualPremiumDifferenceFromLowest = difference,
            AnnualPremiumDifferencePercentFromLowest = percent,
            IsLowestAnnualPremium = isLowest,
            ReadinessStatus = ProposalCanonicalReadinessStatuses.Ready,
            HasCoreFacts = true,
            HasFinancialFacts = true,
            HasPlanDesignFacts = true,
            HasBookletValidation = true,
            HasSourceEvidence = true,
            SourceFileName = $"{carrier.ToLowerInvariant()}-proposal.pdf",
            Evidence = ProposalEvidence(proposalId, "proposal_files")
        };
    }

    private static ProposalIntelligencePremiumValue CreatePremiumValue(long proposalId, string carrier, decimal annualPremium)
    {
        return new ProposalIntelligencePremiumValue
        {
            ProposalId = proposalId,
            CarrierName = carrier.Split('|')[0].Trim(),
            QuoteDisplayName = carrier,
            MonthlyPremium = annualPremium / 12m,
            AnnualPremium = annualPremium,
            EmployerCost = annualPremium * 0.75m,
            EmployeeCost = annualPremium * 0.25m,
            TotalCost = annualPremium,
            Evidence = ProposalEvidence(proposalId, "proposal_premiums")
        };
    }

    private static ProposalIntelligenceRateValue CreateRateValue(long proposalId, string carrier, decimal rate, decimal annualPremium)
    {
        return new ProposalIntelligenceRateValue
        {
            ProposalId = proposalId,
            CarrierName = carrier.Split('|')[0].Trim(),
            QuoteDisplayName = carrier,
            Rate = rate,
            Lives = 10,
            Volume = 100_000m,
            AnnualPremium = annualPremium,
            Evidence = ProposalEvidence(proposalId, "proposal_rates")
        };
    }

    private static ProposalIntelligencePlanDesignValue CreatePlanDesignValue(long proposalId, string carrier, string value)
    {
        return new ProposalIntelligencePlanDesignValue
        {
            ProposalId = proposalId,
            CarrierName = carrier.Split('|')[0].Trim(),
            QuoteDisplayName = carrier,
            ProposalValue = value,
            BookletValue = value,
            ValidationStatus = "MATCH",
            Evidence = ProposalEvidence(proposalId, "proposal_plan_design"),
            BookletEvidence = ProposalEvidence(proposalId, "PlanDesign")
        };
    }

    private static ProposalIntelligenceTermValue CreateTermValue(long proposalId, string carrier, string value)
    {
        return new ProposalIntelligenceTermValue
        {
            ProposalId = proposalId,
            CarrierName = carrier.Split('|')[0].Trim(),
            QuoteDisplayName = carrier,
            Value = value,
            Evidence = ProposalEvidence(proposalId, "proposal_underwriting_assumptions")
        };
    }

    private static CurrentVsQuotePlanDesignComparison CreateCurrentVsQuote()
    {
        var comparison = new CurrentVsQuotePlanDesignComparison
        {
            ClientId = 42,
            ClientName = "Canonical Client Ltd.",
            CurrentCarrier = "Canada Life",
            ReadinessStatus = PlanDesignCanonicalReadinessStatuses.Ready,
            HasCurrentPlan = true,
            HasQuotePlan = true,
            CanCompare = true,
            BuiltAtUtc = new DateTime(2026, 7, 18, 12, 0, 0, DateTimeKind.Utc)
        };
        comparison.Quotes.Add(new CurrentVsQuotePlanDesignQuote
        {
            QuoteKey = "proposal:200",
            ProposalId = 200,
            Carrier = "RBC",
            DisplayName = "RBC | #200",
            Reference = "rbc-proposal.pdf"
        });
        var row = new CurrentVsQuotePlanDesignRow
        {
            ComparisonKey = "ehc:coinsurance",
            BenefitCategory = "Health",
            BenefitName = "Extended Health Care",
            ClassName = "All Classes",
            ProvisionName = "Coinsurance",
            CurrentValue = "80%",
            CurrentAvailability = PlanDesignCanonicalAvailability.Available,
            Status = CurrentVsQuotePlanDesignStatuses.Different,
            CurrentEvidence = CurrentEvidence("current-gir.pdf")
        };
        row.QuoteValues.Add(new CurrentVsQuotePlanDesignValue
        {
            QuoteKey = "proposal:200",
            ProposalId = 200,
            Carrier = "RBC",
            QuoteDisplayName = "RBC | #200",
            Value = "90%",
            Status = CurrentVsQuotePlanDesignStatuses.Different,
            Availability = PlanDesignCanonicalAvailability.Available,
            Evidence = CurrentEvidence("rbc-proposal.pdf")
        });
        comparison.Rows.Add(row);
        comparison.Issues.Add(new PlanDesignCanonicalIssue
        {
            Code = "current_plan_review",
            Severity = PlanDesignCanonicalIssueSeverities.Warning,
            Message = "Current plan source requires review.",
            Benefit = "Extended Health Care",
            Provision = "Coinsurance"
        });
        return comparison;
    }

    private static ProposalCanonicalEvidence ProposalEvidence(long proposalId, string table)
    {
        return new ProposalCanonicalEvidence
        {
            ProposalId = proposalId,
            SourceTable = table,
            SourceFileName = $"proposal-{proposalId}.pdf",
            SourcePage = 5,
            ConfidenceScore = 0.95m
        };
    }

    private static PlanDesignCanonicalEvidence CurrentEvidence(string fileName)
    {
        return new PlanDesignCanonicalEvidence
        {
            SourceTable = "PlanDesign",
            SourceFileName = fileName,
            SourcePage = 10,
            ConfidenceScore = 0.90m
        };
    }
}

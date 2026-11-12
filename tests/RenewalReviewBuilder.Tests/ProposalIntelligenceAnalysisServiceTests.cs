using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class ProposalIntelligenceAnalysisServiceTests
{
    [Fact]
    public void Build_ProducesCompleteFactualCarrierComparison()
    {
        var profile = CreateProfile(
            CreateCompleteProposal(100, "Manulife", 12_000m, 1.25m, "80%", "24 months"),
            CreateCompleteProposal(200, "RBC", 14_400m, 1.50m, "90%", "12 months"));

        var analysis = new ProposalIntelligenceAnalysisService().Build(profile);

        Assert.Equal(ProposalIntelligenceAnalysisVersions.Current, analysis.AnalysisVersion);
        Assert.Equal(ProposalCanonicalProfileVersions.Current, analysis.ProfileVersion);
        Assert.True(analysis.IsReadOnly);
        Assert.Equal(ProposalCanonicalReadinessStatuses.Ready, analysis.ReadinessStatus);
        Assert.True(analysis.CanCompareCarriers);
        Assert.True(analysis.CanCompareFinancials);
        Assert.True(analysis.CanComparePlanDesign);
        Assert.True(analysis.HasComparativeAnalysis);
        Assert.Equal(2, analysis.Quotes.Count);
        Assert.Equal(12_000m, analysis.LowestAnnualPremium);
        Assert.Equal(14_400m, analysis.HighestAnnualPremium);
        Assert.Equal(2_400m, analysis.AnnualPremiumSpread);
        Assert.Equal(20m, analysis.AnnualPremiumSpreadPercent);

        var manulife = analysis.Quotes.Single(quote => quote.CarrierName == "Manulife");
        var rbc = analysis.Quotes.Single(quote => quote.CarrierName == "RBC");
        Assert.True(manulife.IsLowestAnnualPremium);
        Assert.Equal(0m, manulife.AnnualPremiumDifferenceFromLowest);
        Assert.False(rbc.IsLowestAnnualPremium);
        Assert.Equal(2_400m, rbc.AnnualPremiumDifferenceFromLowest);
        Assert.Equal(20m, rbc.AnnualPremiumDifferencePercentFromLowest);
        Assert.Equal("manulife-proposal.pdf", manulife.SourceFileName);

        var rate = Assert.Single(analysis.RateComparisons);
        Assert.True(rate.IsComparable);
        Assert.True(rate.HasRateDifference);
        Assert.Equal(ProposalIntelligenceComparisonStatuses.Different, rate.ComparisonStatus);
        Assert.Equal([1.25m, 1.50m], rate.Values.OrderBy(value => value.Rate).Select(value => value.Rate!.Value).ToArray());

        var premium = Assert.Single(analysis.PremiumComparisons);
        Assert.True(premium.IsComparable);
        Assert.True(premium.HasPremiumDifference);
        Assert.True(premium.HasCostShareDifference);
        Assert.Equal(ProposalIntelligenceComparisonStatuses.Different, premium.ComparisonStatus);

        var planDesign = Assert.Single(analysis.PlanDesignComparisons);
        Assert.True(planDesign.IsComparable);
        Assert.True(planDesign.HasProposalDifference);
        Assert.False(planDesign.HasBookletMismatch);
        Assert.Equal(ProposalIntelligenceComparisonStatuses.Different, planDesign.ComparisonStatus);
        Assert.All(planDesign.Values, value => Assert.Equal("proposal_plan_design", value.Evidence.SourceTable));

        var rateGuarantee = Assert.Single(
            analysis.UnderwritingComparisons,
            row => row.Category == "Underwriting Assumption" && row.TermName == "Rate Guarantee");
        Assert.True(rateGuarantee.HasDifference);
        Assert.Equal(ProposalIntelligenceComparisonStatuses.Different, rateGuarantee.ComparisonStatus);
        Assert.Contains(analysis.Findings, finding => finding.Code == "annual_premium_spread" && finding.Severity == ProposalCanonicalIssueSeverities.Information);
        Assert.Contains(analysis.Findings, finding => finding.Code == "rate_differences_found");
        Assert.Contains(analysis.Findings, finding => finding.Code == "premium_differences_found");
        Assert.Contains(analysis.Findings, finding => finding.Code == "plan_design_differences_found");
        Assert.Contains(analysis.Findings, finding => finding.Code == "underwriting_term_differences_found");
        Assert.Equal(0, analysis.ReviewFindingCount);
    }

    [Fact]
    public void Build_BookletMismatchRequiresReviewWithoutChoosingACarrier()
    {
        var manulife = CreateCompleteProposal(100, "Manulife", 12_000m, 1.25m, "80%", "24 months");
        var rbc = CreateCompleteProposal(200, "RBC", 14_400m, 1.50m, "90%", "12 months");
        rbc.ValidationFacts.Clear();
        rbc.ValidationFacts.Add(CreateValidationFact(200, "RBC", "90%", "80%", "DIFFERENT"));
        var profile = CreateProfile(manulife, rbc);

        var analysis = new ProposalIntelligenceAnalysisService().Build(profile);

        Assert.Equal(ProposalCanonicalReadinessStatuses.NeedsReview, analysis.ReadinessStatus);
        var planDesign = Assert.Single(analysis.PlanDesignComparisons);
        Assert.True(planDesign.HasBookletMismatch);
        Assert.Equal(ProposalIntelligenceComparisonStatuses.NeedsReview, planDesign.ComparisonStatus);
        var finding = Assert.Single(analysis.Findings, item => item.Code == "proposal_booklet_mismatch");
        Assert.Equal(200, finding.ProposalId);
        Assert.Equal("RBC", finding.Carrier);
        Assert.Equal(ProposalIntelligenceFindingCategories.Validation, finding.Category);
        Assert.DoesNotContain(analysis.Findings, item => item.Code.Contains("recommend", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Build_SinglePartialProposalIsUsableButNotComparable()
    {
        var proposal = new ProposalCanonicalQuote
        {
            ProposalId = 100,
            CarrierName = "Manulife",
            SourceFileName = "partial-proposal.pdf",
            SourceFileHash = "partial-hash",
            Evidence = Evidence(100, "proposal_files", "partial-proposal.pdf")
        };
        proposal.PlanDesignFacts.Add(new ProposalCanonicalPlanDesignFact
        {
            BenefitLine = "Extended Health Care",
            ProvisionName = "Coinsurance",
            ProposalValue = "80%",
            Evidence = Evidence(100, "proposal_plan_design", "partial-proposal.pdf")
        });
        var profile = CreateProfile(proposal);

        var analysis = new ProposalIntelligenceAnalysisService().Build(profile);

        Assert.Equal(ProposalCanonicalReadinessStatuses.Partial, analysis.ReadinessStatus);
        Assert.False(analysis.CanCompareCarriers);
        Assert.False(analysis.CanCompareFinancials);
        Assert.False(analysis.HasComparativeAnalysis);
        Assert.Empty(analysis.RateComparisons);
        Assert.Empty(analysis.PremiumComparisons);
        Assert.Equal(ProposalIntelligenceComparisonStatuses.NotComparable, Assert.Single(analysis.PlanDesignComparisons).ComparisonStatus);
        Assert.Contains(analysis.Findings, finding => finding.Code == "proposal_financial_facts_missing");
        Assert.Contains(analysis.Findings, finding => finding.Code == "proposal_validation_missing");
        Assert.Contains(analysis.Findings, finding => finding.Code == "carrier_comparison_unavailable");
    }

    [Fact]
    public void Build_MissingCarrierRateStaysIncompleteInsteadOfBeingInferred()
    {
        var manulife = CreateCompleteProposal(100, "Manulife", 12_000m, 1.25m, "80%", "24 months");
        var rbc = CreateCompleteProposal(200, "RBC", 14_400m, 1.50m, "90%", "12 months");
        var canadaLife = CreateCompleteProposal(300, "Canada Life", 13_200m, 1.40m, "80%", "18 months");
        canadaLife.Rates.Clear();
        var profile = CreateProfile(manulife, rbc, canadaLife);

        var analysis = new ProposalIntelligenceAnalysisService().Build(profile);

        var rate = Assert.Single(analysis.RateComparisons);
        Assert.Equal(3, rate.ExpectedProposalCount);
        Assert.Equal(1, rate.MissingProposalCount);
        Assert.True(rate.HasRateDifference);
        Assert.Equal(ProposalIntelligenceComparisonStatuses.Incomplete, rate.ComparisonStatus);
        Assert.DoesNotContain(rate.Values, value => value.ProposalId == 300);
        Assert.True(analysis.CanCompareFinancials);
        Assert.Equal(ProposalCanonicalReadinessStatuses.Ready, analysis.ReadinessStatus);
    }

    [Fact]
    public void Build_PreservesMultipleVersionsFromTheSameCarrier()
    {
        var oldVersion = CreateCompleteProposal(100, "Manulife", 12_000m, 1.25m, "80%", "24 months");
        oldVersion.ProposalDate = new DateOnly(2026, 6, 1);
        var newVersion = CreateCompleteProposal(101, "Manulife", 11_500m, 1.20m, "80%", "24 months");
        newVersion.ProposalDate = new DateOnly(2026, 7, 1);
        var profile = CreateProfile(oldVersion, newVersion);

        var analysis = new ProposalIntelligenceAnalysisService().Build(profile);

        Assert.Equal(2, analysis.Quotes.Count);
        Assert.Equal(2, analysis.Quotes.Select(quote => quote.QuoteKey).Distinct().Count());
        Assert.Contains(analysis.Quotes, quote => quote.DisplayName.Contains("#100", StringComparison.Ordinal));
        Assert.Contains(analysis.Quotes, quote => quote.DisplayName.Contains("#101", StringComparison.Ordinal));
        Assert.Equal([100L, 101L], analysis.Quotes.Select(quote => quote.ProposalId).OrderBy(value => value).ToArray());
    }

    [Fact]
    public void Build_NoProposalsIsNotReady()
    {
        var profile = CreateProfile();

        var analysis = new ProposalIntelligenceAnalysisService().Build(profile);

        Assert.Equal(ProposalCanonicalReadinessStatuses.NotReady, analysis.ReadinessStatus);
        Assert.False(analysis.HasProposals);
        Assert.False(analysis.HasComparativeAnalysis);
        Assert.Empty(analysis.Quotes);
        Assert.Contains(analysis.Findings, finding => finding.Code == "proposals_missing" && finding.Severity == ProposalCanonicalIssueSeverities.Error);
    }

    [Fact]
    public async Task BuildAsync_MissingDatabaseRemainsReadOnly()
    {
        var folder = Path.Combine(Path.GetTempPath(), "ProposalIntelligenceAnalysisTests", Guid.NewGuid().ToString("N"));
        var databasePath = Path.Combine(folder, "Benefits_Intelligence.db");
        try
        {
            var analysis = await new ProposalIntelligenceAnalysisService(databasePath).BuildAsync(42);

            Assert.False(File.Exists(databasePath));
            Assert.Equal(ProposalCanonicalReadinessStatuses.NotReady, analysis.ReadinessStatus);
            Assert.Contains(analysis.Findings, finding => finding.Code == "database_missing");
        }
        finally
        {
            if (Directory.Exists(folder))
            {
                Directory.Delete(folder, recursive: true);
            }
        }
    }

    private static ProposalCanonicalProfile CreateProfile(params ProposalCanonicalQuote[] proposals)
    {
        var profile = new ProposalCanonicalProfile
        {
            ClientId = 42,
            ClientName = "Canonical Client Ltd.",
            ClientNumber = "CAN-42"
        };
        profile.Proposals.AddRange(proposals);
        return profile;
    }

    private static ProposalCanonicalQuote CreateCompleteProposal(
        long proposalId,
        string carrier,
        decimal annualPremium,
        decimal rate,
        string planDesignValue,
        string rateGuarantee)
    {
        var monthlyPremium = annualPremium / 12m;
        var fileName = $"{carrier.ToLowerInvariant().Replace(" ", "-")}-proposal.pdf";
        var proposal = new ProposalCanonicalQuote
        {
            ProposalId = proposalId,
            CarrierId = proposalId,
            CarrierKey = carrier.ToLowerInvariant().Replace(" ", "-"),
            CarrierName = carrier,
            ProposalDate = new DateOnly(2026, 7, 1),
            EffectiveDate = new DateOnly(2026, 8, 1),
            SourceFileName = fileName,
            SourceFileHash = $"hash-{proposalId}",
            ExtractionConfidence = 0.95m,
            ExtractionMethod = "Focused Carrier Comparison",
            Status = "VALID",
            Evidence = Evidence(proposalId, "proposal_files", fileName)
        };
        proposal.Rates.Add(new ProposalCanonicalRate
        {
            BenefitLine = "Extended Health Care",
            ClassName = "All Classes",
            Tier = "Single",
            Lives = 10,
            Volume = 100_000m,
            Rate = rate,
            RateBasis = "Per $1,000",
            MonthlyPremium = monthlyPremium,
            AnnualPremium = annualPremium,
            Evidence = Evidence(proposalId, "proposal_rates", fileName)
        });
        proposal.Premiums.Add(new ProposalCanonicalPremium
        {
            BenefitLine = "Extended Health Care",
            ClassName = "All Classes",
            Tier = "Single",
            MonthlyPremium = monthlyPremium,
            AnnualPremium = annualPremium,
            EmployerCost = annualPremium * 0.75m,
            EmployeeCost = annualPremium * 0.25m,
            TotalCost = annualPremium,
            Evidence = Evidence(proposalId, "proposal_premiums", fileName)
        });
        proposal.PlanDesignFacts.Add(new ProposalCanonicalPlanDesignFact
        {
            BenefitLine = "Extended Health Care",
            ProvisionName = "Coinsurance",
            ProposalValue = planDesignValue,
            ValidationStatus = "VALID",
            Evidence = Evidence(proposalId, "proposal_plan_design", fileName)
        });
        proposal.ValidationFacts.Add(CreateValidationFact(proposalId, carrier, planDesignValue, planDesignValue, "MATCH"));
        proposal.UnderwritingAssumptions.Add(new ProposalCanonicalUnderwritingAssumption
        {
            Name = "Rate Guarantee",
            Value = rateGuarantee,
            Evidence = Evidence(proposalId, "proposal_underwriting_assumptions", fileName)
        });
        proposal.PoolingRetention.Add(new ProposalCanonicalPoolingRetention
        {
            BenefitLine = "All Benefits",
            PoolingLevel = "$10,000",
            RateGuarantee = rateGuarantee,
            Evidence = Evidence(proposalId, "proposal_pooling_retention", fileName)
        });
        return proposal;
    }

    private static ProposalCanonicalPlanDesignFact CreateValidationFact(
        long proposalId,
        string carrier,
        string proposalValue,
        string bookletValue,
        string status)
    {
        var fileName = $"{carrier.ToLowerInvariant().Replace(" ", "-")}-proposal.pdf";
        return new ProposalCanonicalPlanDesignFact
        {
            BenefitLine = "Extended Health Care",
            ProvisionName = "Coinsurance",
            ProposalValue = proposalValue,
            BookletValue = bookletValue,
            ValidationStatus = status,
            Variance = status == "MATCH" ? "" : "Different values",
            Evidence = Evidence(proposalId, "proposal_booklet_validation", fileName),
            BookletEvidence = Evidence(proposalId, "PlanDesign", "current-booklet.pdf")
        };
    }

    private static ProposalCanonicalEvidence Evidence(long proposalId, string sourceTable, string fileName)
    {
        return new ProposalCanonicalEvidence
        {
            SourceTable = sourceTable,
            ProposalId = proposalId,
            SourceFileName = fileName,
            SourceFileHash = $"evidence-{sourceTable}-{proposalId}",
            SourcePage = 5,
            ConfidenceScore = 0.95m,
            Status = "VALID",
            ExtractionMethod = "Test"
        };
    }
}

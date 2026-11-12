using System.Collections.Generic;
using Microsoft.Data.Sqlite;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class TieredReimbursementRuleTests
{
    private const string SourceWording =
        "Eligible expenses are reimbursed at 80% for the first $1,000 paid per certificate in a calendar year and reimbursed at 100% for eligible expenses over $1,000 paid in that calendar year.";

    [Fact]
    public void Domain_ValidTieredRule_IsComplete()
    {
        var rule = BuildValidRule();

        Assert.True(rule.IsComplete);
        Assert.True(rule.IsDisplayable);
        Assert.Empty(rule.Validate());
    }

    [Fact]
    public void Domain_MissingTier2_IsIncompleteAndRequiresReview()
    {
        var rule = BuildValidRule();
        rule.Tier2ReimbursementPercent = null;
        rule.NormalizationStatus = BenefitRuleNormalizationStatus.NeedsReview;

        Assert.False(rule.IsComplete);
        Assert.False(rule.IsDisplayable);
        Assert.Contains(rule.Validate(), error => error.Contains("Tier 2", StringComparison.OrdinalIgnoreCase));
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(101)]
    public void Domain_InvalidPercentages_AreRejected(decimal percentage)
    {
        var rule = BuildValidRule();
        rule.Tier1ReimbursementPercent = percentage;

        Assert.False(rule.IsComplete);
        Assert.Contains(rule.Validate(), error => error.Contains("between 0 and 100", StringComparison.OrdinalIgnoreCase));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1000)]
    public void Domain_NonPositiveThreshold_IsRejected(decimal threshold)
    {
        var rule = BuildValidRule();
        rule.Tier1ThresholdAmount = threshold;

        Assert.False(rule.IsComplete);
        Assert.Contains(rule.Validate(), error => error.Contains("greater than zero", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Formatter_ProducesRequiredFullAndCompactWording()
    {
        var rule = BuildValidRule();

        Assert.Equal(
            "80% of the first $1,000 per certificate per calendar year, then 100%",
            BenefitRuleDisplayFormatter.FormatFull(rule));
        Assert.Equal(
            "80% of first $1,000/certificate/year, then 100%",
            BenefitRuleDisplayFormatter.FormatCompact(rule));
    }

    [Fact]
    public void Formatter_UsesStructuredValuesRatherThanTheExample()
    {
        var rule = BuildValidRule();
        rule.Tier1ReimbursementPercent = 70m;
        rule.Tier1ThresholdAmount = 500m;
        rule.Tier2ReimbursementPercent = 90m;
        rule.ThresholdBasis = BenefitRuleThresholdBasis.Person;
        rule.ResetPeriod = BenefitRuleResetPeriod.PolicyYear;

        Assert.Equal(
            "70% of the first $500 per person per policy year, then 90%",
            BenefitRuleDisplayFormatter.FormatFull(rule));
        Assert.Equal(
            "70% of first $500/person/policy year, then 90%",
            BenefitRuleDisplayFormatter.FormatCompact(rule));
    }

    [Fact]
    public void Formatter_NullOrIncompleteRules_DoNotCrashOrDisplay()
    {
        var incomplete = BuildValidRule();
        incomplete.Tier2ReimbursementPercent = null;

        Assert.Equal("", BenefitRuleDisplayFormatter.FormatFull(null));
        Assert.Equal("", BenefitRuleDisplayFormatter.FormatCompact(incomplete));
    }

    [Fact]
    public void Normalizer_MapsCarrierWordingToCanonicalFields()
    {
        var rule = Assert.IsType<PlanBenefitRule>(TieredReimbursementRuleNormalizer.Parse(
            SourceWording,
            "Benefit Booklet.pdf",
            12,
            0.96m));

        Assert.Equal(BenefitRuleType.TieredReimbursement, rule.RuleType);
        Assert.Equal(80m, rule.Tier1ReimbursementPercent);
        Assert.Equal(1000m, rule.Tier1ThresholdAmount);
        Assert.Equal(100m, rule.Tier2ReimbursementPercent);
        Assert.Equal(BenefitRuleThresholdBasis.Certificate, rule.ThresholdBasis);
        Assert.Equal(BenefitRuleResetPeriod.CalendarYear, rule.ResetPeriod);
        Assert.Equal("CAD", rule.CurrencyCode);
        Assert.Equal(BenefitRuleNormalizationStatus.Normalized, rule.NormalizationStatus);
        Assert.Equal(SourceWording, rule.SourceText);
        Assert.Contains("Benefit Booklet.pdf", rule.SourceLocation);
        Assert.Contains("Page 12", rule.SourceLocation);
    }

    [Theory]
    [InlineData("80% for the first $1,000 per certificate per calendar year, then 100%")]
    [InlineData("80% up to $1,000 per certificate per calendar year; 100% thereafter")]
    [InlineData("80% of eligible expenses until the annual $1,000 threshold per certificate is reached in a calendar year, then 100%")]
    [InlineData("Eligible expenses reimbursed at 80% for the first $1,000 paid per certificate and 100% over $1,000 in a calendar year")]
    public void Normalizer_MapsEquivalentCarrierWording(string sourceText)
    {
        var rule = Assert.IsType<PlanBenefitRule>(TieredReimbursementRuleNormalizer.Parse(sourceText));

        Assert.Equal(80m, rule.Tier1ReimbursementPercent);
        Assert.Equal(1000m, rule.Tier1ThresholdAmount);
        Assert.Equal(100m, rule.Tier2ReimbursementPercent);
        Assert.Equal(BenefitRuleThresholdBasis.Certificate, rule.ThresholdBasis);
        Assert.Equal(BenefitRuleResetPeriod.CalendarYear, rule.ResetPeriod);
        Assert.Equal(BenefitRuleNormalizationStatus.Normalized, rule.NormalizationStatus);
    }

    [Fact]
    public void Normalizer_AssociatesSharedFootnoteOnlyWithMarkedApplicableBenefits()
    {
        var features = BuildMappedFeatures();

        var mappings = TieredReimbursementRuleNormalizer.MapPlanDesignFeatures(features);

        Assert.Equal([0, 1, 2, 3], mappings.Select(mapping => mapping.SourceIndex).Order().ToArray());
        Assert.All(mappings, mapping => Assert.True(mapping.Rule.IsDisplayable));
        Assert.DoesNotContain(mappings, mapping => mapping.SourceIndex is 4 or 5 or 6 or 7 or 8);
    }

    [Fact]
    public void Normalizer_IncompleteTieredEvidence_RemainsNeedsReview()
    {
        const string incompleteWording =
            "Eligible expenses are reimbursed at 80% for the first amount paid per certificate in a calendar year, then 100%.";
        var rule = Assert.IsType<PlanBenefitRule>(TieredReimbursementRuleNormalizer.Parse(incompleteWording));

        Assert.False(rule.IsComplete);
        Assert.Equal(BenefitRuleNormalizationStatus.NeedsReview, rule.NormalizationStatus);
        Assert.Equal(incompleteWording, rule.SourceText);
    }

    [Fact]
    public void Normalizer_MissingSecondTier_RemainsNeedsReviewWithoutCopyingTierOne()
    {
        const string incompleteWording =
            "Eligible expenses are reimbursed at 80% for the first $1,000 per certificate in a calendar year, then the remaining percentage is unreadable.";
        var rule = Assert.IsType<PlanBenefitRule>(TieredReimbursementRuleNormalizer.Parse(incompleteWording));

        Assert.Equal(80m, rule.Tier1ReimbursementPercent);
        Assert.Equal(1000m, rule.Tier1ThresholdAmount);
        Assert.Null(rule.Tier2ReimbursementPercent);
        Assert.Equal(BenefitRuleNormalizationStatus.NeedsReview, rule.NormalizationStatus);
    }

    [Fact]
    public void Normalizer_UnconfirmedAsteriskAssociation_PreservesEvidenceForReview()
    {
        const string sourceFile = "unconfirmed-tiered-benefit-booklet.pdf";
        var features = new List<PlanDesignFeatureExtraction>
        {
            Feature("Prescription Drugs", "80%", "80%", sourceFile, 10),
            Feature("Health Care Practitioners", "80%", "80%", sourceFile, 10),
            Feature("Vision Care", "80%", "80%", sourceFile, 10),
            Feature("Other Medical Expenses", "80%", "80%", sourceFile, 10),
            Feature("Extended Health Care Footnote", SourceWording, "", sourceFile, 12)
        };

        var mapping = Assert.Single(TieredReimbursementRuleNormalizer.MapPlanDesignFeatures(features));

        Assert.Equal(4, mapping.SourceIndex);
        Assert.Equal(BenefitRuleNormalizationStatus.NeedsReview, mapping.Rule.NormalizationStatus);
        Assert.Equal(SourceWording, mapping.Rule.SourceText);
        Assert.False(mapping.Rule.IsDisplayable);
    }

    [Fact]
    public void Normalizer_DoesNotCombineUnrelatedBookletFactsIntoTieredRule()
    {
        const string sourceFile = "multi-rate-benefit-booklet.pdf";
        var features = new List<PlanDesignFeatureExtraction>
        {
            Feature(
                "Long Term Disability",
                "66.7% of monthly earnings to a maximum benefit",
                "66.7%",
                sourceFile,
                10),
            Feature(
                "Prescription Drugs",
                "80% for preferred drugs; 90% at a preferred pharmacy",
                "80%; 90%",
                sourceFile,
                11),
            Feature(
                "Vision Care",
                "The first exam is covered, then every 24 months.",
                "",
                sourceFile,
                12)
        };

        var mappings = TieredReimbursementRuleNormalizer.MapPlanDesignFeatures(features);

        Assert.Empty(mappings);
    }

    [Fact]
    public void BenefitSummary_PrefersTieredWordingAndKeepsUnaffectedFlatPercentages()
    {
        var summary = Assert.Single(BenefitSummaryDisplayFormatter.Build(new BenefitSummaryBuildRequest
        {
            ClientName = "Tiered Display Client",
            Rows =
            [
                SummaryRow("Prescription Drugs*", "80%", BuildValidRule()),
                SummaryRow("Hospital", "100%"),
                SummaryRow("Travel Insurance and Assistance", "100%")
            ]
        }));

        var drugs = Assert.Single(summary.Sections, section => section.SectionTitle == "Prescription Drugs");
        Assert.Contains(drugs.Lines, line =>
            line.Label == "Coinsurance" &&
            line.DisplayValue == "80% of the first $1,000 per certificate per calendar year, then 100%");
        Assert.DoesNotContain(drugs.Lines, line => line.DisplayValue == "80%");

        var hospital = Assert.Single(summary.Sections, section => section.SectionTitle == "Hospital");
        Assert.Contains(hospital.Lines, line => line.DisplayValue == "100%");

        var travel = Assert.Single(summary.Sections, section => section.SectionTitle == "Emergency Travel");
        Assert.Contains(travel.Lines, line => line.DisplayValue == "100%");
    }

    [Fact]
    public void BenefitSummary_IncompleteTieredRuleDoesNotFallBackToFlatPercentage()
    {
        var incomplete = BuildValidRule();
        incomplete.Tier2ReimbursementPercent = null;
        incomplete.NormalizationStatus = BenefitRuleNormalizationStatus.NeedsReview;

        var summary = Assert.Single(BenefitSummaryDisplayFormatter.Build(new BenefitSummaryBuildRequest
        {
            ClientName = "Tiered Review Client",
            Rows = [SummaryRow("Prescription Drugs*", "80%", incomplete)]
        }));

        var drugs = Assert.Single(summary.Sections, section => section.SectionTitle == "Prescription Drugs");
        Assert.Contains(drugs.Lines, line =>
            line.Label == "Coinsurance" &&
            line.DisplayValue == BenefitRuleDisplayFormatter.ReviewRequiredDisplayText);
        Assert.DoesNotContain(drugs.Lines, line => line.DisplayValue == "80%");
    }

    [Fact]
    public async Task RepositoryAndClientWorkspace_RoundTripStructuredRulesAndEvidence()
    {
        var testFolder = Path.Combine(Path.GetTempPath(), "TieredReimbursementTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(testFolder);
        var databasePath = Path.Combine(testFolder, "Benefits_Intelligence.db");
        var sourceFile = Path.Combine(testFolder, "tiered-benefit-booklet.pdf");
        await File.WriteAllTextAsync(sourceFile, "booklet test source");

        try
        {
            var database = new DatabaseService(databasePath);
            var rows = BuildMappedFeatures()
                .Select(feature => new PlanDesignRecord
                {
                    ClassName = "Class 1 - All Eligible Employees",
                    BenefitCategory = feature.Category,
                    BenefitName = feature.Benefit,
                    CoverageDescription = feature.FeatureValue,
                    ReimbursementPercent = feature.Coinsurance == "80%" ? 80m : feature.Coinsurance == "100%" ? 100m : null,
                    SourceFileName = Path.GetFileName(sourceFile),
                    SourcePage = feature.SourcePage,
                    Notes = feature.SourceText == feature.FeatureValue ? "" : feature.SourceText
                })
                .ToList();

            var save = await database.SavePlanDesignAsync(
                new ClientRecord
                {
                    ClientName = "Canonical Rule Test Client",
                    ClientNumber = "RULE-001",
                    Carrier = "Test Carrier",
                    EffectiveDate = new DateOnly(2026, 1, 1)
                },
                [sourceFile],
                rows);

            Assert.True(save.Succeeded, save.Message);
            Assert.True(save.ClientId.HasValue, save.Message);

            var reloadedRules = await database.GetPlanBenefitRulesAsync(save.ClientId.Value);
            Assert.Equal(4, reloadedRules.Count);
            Assert.All(reloadedRules, rule =>
            {
                Assert.Equal(80m, rule.Tier1ReimbursementPercent);
                Assert.Equal(1000m, rule.Tier1ThresholdAmount);
                Assert.Equal(100m, rule.Tier2ReimbursementPercent);
                Assert.Equal(BenefitRuleThresholdBasis.Certificate, rule.ThresholdBasis);
                Assert.Equal(BenefitRuleResetPeriod.CalendarYear, rule.ResetPeriod);
                Assert.Equal("CAD", rule.CurrencyCode);
                Assert.Equal(SourceWording, rule.SourceText);
                Assert.True(rule.SourceDocumentId.HasValue);
                Assert.True(rule.PlanDesignFeatureId.HasValue);
            });

            var workspace = await new PlanDesignWorkspaceDataService(databasePath)
                .LoadAsync(save.ClientId, "Canonical Rule Test Client", "Test Carrier");
            var summary = Assert.Single(workspace.BenefitSummaries);
            var expected = "80% of the first $1,000 per certificate per calendar year, then 100%";

            AssertTieredSection(summary, "Prescription Drugs", expected);
            AssertTieredSection(summary, "Paramedicals", expected);
            AssertTieredSection(summary, "Vision Care", expected);
            AssertTieredSection(summary, "Medical Services & Supplies", expected);

            var hospital = Assert.Single(summary.Sections, section => section.SectionTitle == "Hospital");
            Assert.Contains(hospital.Lines, line => line.DisplayValue == "100%");
            var travel = Assert.Single(summary.Sections, section => section.SectionTitle == "Emergency Travel");
            Assert.Contains(travel.Lines, line => line.DisplayValue == "100%");

            Assert.Contains(workspace.ClassComparison.ComparisonRows, row =>
                row.Benefit.Contains("Prescription", StringComparison.OrdinalIgnoreCase) &&
                row.Cells.Any(value => value.DisplayValue.Contains(
                    "80% of first $1,000/certificate/year, then 100%",
                    StringComparison.OrdinalIgnoreCase)));
        }
        finally
        {
            SqliteConnection.ClearAllPools();
            if (Directory.Exists(testFolder))
            {
                Directory.Delete(testFolder, recursive: true);
            }
        }
    }

    private static PlanBenefitRule BuildValidRule()
    {
        return new PlanBenefitRule
        {
            RuleType = BenefitRuleType.TieredReimbursement,
            Tier1ReimbursementPercent = 80m,
            Tier1ThresholdAmount = 1000m,
            Tier2ReimbursementPercent = 100m,
            ThresholdBasis = BenefitRuleThresholdBasis.Certificate,
            ResetPeriod = BenefitRuleResetPeriod.CalendarYear,
            CurrencyCode = "CAD",
            SourceText = SourceWording,
            NormalizationStatus = BenefitRuleNormalizationStatus.Normalized,
            IsActive = true
        };
    }

    private static BenefitSummarySourceRow SummaryRow(
        string benefit,
        string coinsurance,
        PlanBenefitRule? rule = null)
    {
        return new BenefitSummarySourceRow
        {
            ClassName = "All Classes",
            Benefit = benefit,
            Category = "Extended Health Care",
            ProvisionName = benefit,
            RawValue = coinsurance,
            Coinsurance = coinsurance,
            Rules = rule is null ? [] : [rule]
        };
    }

    private static List<PlanDesignFeatureExtraction> BuildMappedFeatures()
    {
        const string sourceFile = "tiered-benefit-booklet.pdf";
        return
        [
            Feature("Prescription Drugs*", "80%", "80%", sourceFile, 10),
            Feature("Health Care Practitioners*", "80%", "80%", sourceFile, 10),
            Feature("Vision Care*", "80%", "80%", sourceFile, 10),
            Feature("Other Medical Expenses*", "80%", "80%", sourceFile, 10),
            Feature("Hospital", "100%", "100%", sourceFile, 10),
            Feature("Travel Insurance and Assistance", "100%", "100%", sourceFile, 10),
            Feature("Travel Cancellation Insurance", "100%", "100%", sourceFile, 10),
            Feature("Medical Second Opinion", "100%", "100%", sourceFile, 10),
            Feature("Extended Health Care Footnote", SourceWording, "", sourceFile, 12)
        ];
    }

    private static PlanDesignFeatureExtraction Feature(
        string benefit,
        string value,
        string coinsurance,
        string sourceFile,
        int sourcePage)
    {
        return new PlanDesignFeatureExtraction
        {
            Benefit = benefit,
            Category = "Extended Health Care",
            FeatureName = benefit,
            FeatureValue = value,
            Coinsurance = coinsurance,
            ClassName = "Class 1 - All Eligible Employees",
            SourceFileName = sourceFile,
            SourcePage = sourcePage,
            SourceText = value,
            ConfidenceScore = 0.96m
        };
    }

    private static void AssertTieredSection(
        BenefitSummaryViewModel summary,
        string sectionTitle,
        string expected)
    {
        var section = Assert.Single(summary.Sections, candidate => candidate.SectionTitle == sectionTitle);
        Assert.Contains(section.Lines, line => line.Label == "Coinsurance" && line.DisplayValue == expected);
        Assert.DoesNotContain(section.Lines, line => line.DisplayValue == "80%");
    }
}

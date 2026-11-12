using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services.PolicySchedules;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class ManulifePolicyScheduleExtractorTests
{
    [Fact]
    public void Extract_Detects_Supported_Sections_Classes_And_Evidence()
    {
        var result = new ManulifePolicyScheduleExtractor().Extract("schedule.pdf",
        [
            new PdfPageText { PageNumber = 1, Lines = { "Manulife", "Employee Life Insurance", "Classification Code A,B", "Benefit Amount 2 x annual Earnings rounded to the next higher $1,000 up to a maximum benefit of $500,000", "Non-Evidence Limit $250,000" } },
            new PdfPageText { PageNumber = 2, Lines = { "Dependent Life Insurance", "Classification Code A", "Spouse Benefit $10,000", "Child Benefit $5,000", "Accidental Death and Dismemberment", "Benefit Amount 2 x annual Earnings", "Long Term Disability", "Classification Code B", "Benefit Amount 66.7% of first $2,500 of monthly Earnings, plus 50% of the next $2,500, up to a maximum benefit of $10,000", "CPP/QPP Primary CPP/QPP benefits" } }
        ]);

        Assert.Equal(4, result.Sections.Count);
        var life = Assert.Single(result.Sections, section => section.Benefit == "Employee Life");
        Assert.Equal(["A", "B"], life.ClassificationCodes);
        Assert.Contains(life.Fields, field => field.Label == "Benefit Amount" && field.PageNumber == 1);
        Assert.Contains(result.Sections, section => section.Benefit == "LTD");
    }

    [Fact]
    public void Map_Preserves_PolicySchedule_Evidence_And_Maps_Classes()
    {
        var extraction = new ManulifePolicyScheduleExtractor().Extract("schedule.pdf",
        [new PdfPageText { PageNumber = 3, Lines = { "Employee Life Insurance", "Classification Code A,B", "Non-Evidence Limit $250,000", "Waiting Period 3 months" } }]);
        var mapped = new ManulifePolicyScheduleMapper().Map(extraction, new ClientRecord { ClientName = "Test Client", Carrier = "Manulife" });

        Assert.Contains(mapped.Rows, row => row.ProvisionCode == "non_evidence_limit" && row.ClassName == "Class A" && row.SourcePage == 3);
        Assert.Contains(mapped.Rows, row => row.ProvisionCode == "waiting_period" && row.ClassName == "Class B");
        Assert.All(mapped.Rows, row => Assert.Contains(PlanDesignSources.PolicySchedule, row.SourceSection));
    }
}

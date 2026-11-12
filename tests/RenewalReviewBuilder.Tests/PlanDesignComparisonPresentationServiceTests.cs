using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class PlanDesignComparisonPresentationServiceTests
{
    [Fact]
    public void DomainService_ClassifiesSupportedDirectionsWithoutStringRanking()
    {
        var service = new PlanDesignComparisonDomainService();

        Assert.Equal(PlanDesignComparisonStatuses.Improved, service.Classify("Coinsurance", "80%", "100%", "Different", "Available", "Available").Status);
        Assert.Equal(PlanDesignComparisonStatuses.Reduced, service.Classify("Waiting Period", "90 days", "120 days", "Different", "Available", "Available").Status);
        Assert.Equal(PlanDesignComparisonStatuses.Changed, service.Classify("Special condition", "Option A", "Option B", "Different", "Available", "Available").Status);
        Assert.Equal(PlanDesignComparisonStatuses.Same, service.Classify("Maximum", "$500", "$500", "Same", "Available", "Available").Status);
        Assert.Equal(PlanDesignComparisonStatuses.ReviewRequired, service.Classify("Maximum", "$500", "$750", "Needs Review", "Available", "Available").Status);
    }

    [Fact]
    public void DomainService_LabelsCostPreviewAndUnavailableWithoutInventingANumber()
    {
        Assert.Equal("Preview", PlanDesignComparisonDomainService.FormatCost(new PlanDesignComparisonCostInput { Kind = PlanDesignComparisonCostKinds.Preview }));
        Assert.Equal("Not available", PlanDesignComparisonDomainService.FormatCost(null));
        Assert.Equal("+$0.44 PEPM", PlanDesignComparisonDomainService.FormatCost(new PlanDesignComparisonCostInput { Kind = PlanDesignComparisonCostKinds.Pepm, Amount = 0.44m }));
    }

    [Fact]
    public void Presentation_BuildsCanonicalGroupsAndDefaultsToChangedOnly()
    {
        var comparison = CreateComparison();
        var model = new PlanDesignComparisonPresentationService().Build(comparison);

        Assert.Equal(PlanDesignComparisonViewModes.ChangedOnly, model.ViewMode);
        Assert.Contains(model.Groups, group => group.Category == "Dental");
        Assert.Contains(model.DisplayRows, row => row.Provision?.ProvisionName == "Coinsurance");
        Assert.DoesNotContain(model.DisplayRows, row => row.Provision?.ProvisionName == "Maximum");
        Assert.Equal("Not available", Assert.Single(model.SummaryMetrics, metric => metric.Label == "Estimated Cost Impact").Value);
        Assert.True(Assert.Single(model.DisplayRows, row => row.Provision?.ProvisionName == "Coinsurance").Provision!.HasEvidence);
    }

    [Fact]
    public void Presentation_ChangesRowsWhenAnotherScenarioIsSelected()
    {
        var comparison = CreateComparison();
        comparison.Quotes.Add(new CurrentVsQuotePlanDesignQuote { QuoteKey = "quote:two", DisplayName = "Scenario Two" });
        comparison.Rows[0].QuoteValues.Add(new CurrentVsQuotePlanDesignValue { QuoteKey = "quote:two", Value = "80%", Status = CurrentVsQuotePlanDesignStatuses.Same, Availability = PlanDesignCanonicalAvailability.Available });
        comparison.Rows[1].QuoteValues.Add(new CurrentVsQuotePlanDesignValue { QuoteKey = "quote:two", Value = "$1,000", Status = CurrentVsQuotePlanDesignStatuses.Different, Availability = PlanDesignCanonicalAvailability.Available });

        var model = new PlanDesignComparisonPresentationService().Build(comparison);
        model.SelectedScenarioKey = "quote:two";

        Assert.Contains(model.DisplayRows, row => row.Provision?.ProvisionName == "Maximum");
        Assert.DoesNotContain(model.DisplayRows, row => row.Provision?.ProvisionName == "Coinsurance");
    }

    private static CurrentVsQuotePlanDesignComparison CreateComparison()
    {
        var comparison = new CurrentVsQuotePlanDesignComparison { ClientId = 1, CurrentCarrier = "Current Carrier", CanCompare = true };
        comparison.Quotes.Add(new CurrentVsQuotePlanDesignQuote { QuoteKey = "quote:one", DisplayName = "Scenario One" });
        comparison.Rows.Add(CreateRow("Dental", "Coinsurance", "80%", "100%", "Different", "current.pdf", "quote.pdf"));
        comparison.Rows.Add(CreateRow("Dental", "Maximum", "$750", "$750", "Same", "current.pdf", "quote.pdf"));
        return comparison;
    }

    private static CurrentVsQuotePlanDesignRow CreateRow(string category, string provision, string current, string selected, string status, string currentFile, string selectedFile)
    {
        var row = new CurrentVsQuotePlanDesignRow
        {
            BenefitCategory = category,
            BenefitName = category,
            ProvisionName = provision,
            CurrentValue = current,
            CurrentAvailability = PlanDesignCanonicalAvailability.Available,
            CurrentEvidence = new PlanDesignCanonicalEvidence { SourceFileName = currentFile }
        };
        row.QuoteValues.Add(new CurrentVsQuotePlanDesignValue
        {
            QuoteKey = "quote:one",
            Value = selected,
            Status = status,
            Availability = PlanDesignCanonicalAvailability.Available,
            Evidence = new PlanDesignCanonicalEvidence { SourceFileName = selectedFile }
        });
        return row;
    }
}

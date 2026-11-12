using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class PlanDesignVerificationServiceTests
{
    [Theory]
    [InlineData("2x Earnings", "2 x Annual Earnings")]
    [InlineData("$500,000", "500000")]
    [InlineData("119 days", "119 Day Waiting Period")]
    [InlineData("Age 85 or retirement", "Age 85 or retirement, whichever occurs first")]
    public void Compare_Recognizes_Equivalent_Wording(string left, string right)
    {
        var result = new PlanDesignVerificationService().Compare(Evidence(left), Evidence(right));
        Assert.True(result.Status is PlanDesignVerificationStatuses.Confirmed or PlanDesignVerificationStatuses.EquivalentWording);
    }

    [Fact]
    public void Compare_DifferentPlanState_IsProposedChange_NotConflict()
    {
        var result = new PlanDesignVerificationService().Compare(Evidence("Age 85"), Evidence("Age 70", PlanDesignStates.FutureRenewal));
        Assert.Equal(PlanDesignVerificationStatuses.ProposedPlanChange, result.Status);
    }

    [Fact]
    public void Compare_ApprovedManualValue_IsProtected()
    {
        var result = new PlanDesignVerificationService().Compare(Evidence("Age 85", approved: true), Evidence("Age 70"));
        Assert.Equal(PlanDesignVerificationStatuses.ManualReviewRequired, result.Status);
    }

    private static PlanDesignEvidenceFact Evidence(string value, string state = PlanDesignStates.CurrentContract, bool approved = false) => new()
    {
        PlanState = state, IsManualApproved = approved, SourceAuthority = 2,
        Fact = new PlanDesignCanonicalFact { CanonicalBenefitCode = "LIFE", ClassName = "Class A", FeatureName = "Termination Age", FeatureValue = value }
    };
}

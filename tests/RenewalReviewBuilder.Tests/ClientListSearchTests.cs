using RenewalReviewBuilder.App;
using Xunit;

namespace RenewalReviewBuilder.Tests;

public sealed class ClientListSearchTests
{
    private static readonly ClientListItem FamilyServices = new(
        "AFSOGV",
        "Family Services Of Greater Vancouver",
        "Active",
        "Not set",
        206,
        "TBD",
        "GroupHEALTH",
        "TBD",
        100,
        "Jul 1, 2026",
        45);

    [Theory]
    [InlineData("Family Services")]
    [InlineData("family vancouver")]
    [InlineData("AFSOGV")]
    [InlineData("GroupHEALTH")]
    [InlineData("active grouphealth")]
    public void MatchesSearchFindsClientAcrossVisibleFields(string search)
    {
        Assert.True(FamilyServices.MatchesSearch(search));
    }

    [Fact]
    public void MatchesSearchRejectsUnrelatedClient()
    {
        Assert.False(FamilyServices.MatchesSearch("Manulife CL-0041"));
    }
}

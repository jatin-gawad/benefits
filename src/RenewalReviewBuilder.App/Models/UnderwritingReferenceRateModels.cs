namespace RenewalReviewBuilder.App.Models;

public static class UnderwritingRateTypes
{
    public const string Mortality = "Mortality";
    public const string Morbidity = "Morbidity";
    public const string SoaBasicLifeMortality = "SOA Basic Life Mortality Benchmark";
    public const string SoaGroupLifeWaiver = "SOA Group Life Waiver Benchmark";
}

public static class UnderwritingRateGenders
{
    public const string Female = "Female";
    public const string Male = "Male";
}

public sealed class UnderwritingReferenceRate
{
    public string RateType { get; init; } = "";
    public string VersionName { get; init; } = "";
    public string AgeBand { get; init; } = "";
    public string Gender { get; init; } = "";
    public decimal Rate { get; init; }
    public string UnitBasis { get; init; } = "";
    public string SourceName { get; init; } = "";
    public string SourceSheet { get; init; } = "";
    public int DisplayOrder { get; init; }
}

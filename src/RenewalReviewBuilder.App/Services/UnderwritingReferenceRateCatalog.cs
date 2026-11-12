using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public static class UnderwritingReferenceRateCatalog
{
    public const string VersionName = "WB Renewal master v1";
    public const string SourceName = "Copy of WB Renewal - 2  3 Option.xlsx";
    public const string MortalityUnitBasis = "Monthly rate per $1,000 of Life volume";
    public const string MorbidityUnitBasis = "Monthly rate per $100 of LTD volume";
    public const string SoaVersionName = "SOA Group Life 2010-13 Basic Life";
    public const string SoaSourceName = "Group Life 2010-13 Basic Life Death and Dis AE Ratios with Additional Variables - Final.xlsx";
    public const string SoaMortalityUnitBasis = "Expected death rate by amount per $1,000";
    public const string SoaWaiverUnitBasis = "Expected Group Life disability-waiver rate by amount; normalized per $1,000 for mix comparison";

    public static IReadOnlyList<UnderwritingReferenceRate> Defaults { get; } = BuildDefaults();

    private static IReadOnlyList<UnderwritingReferenceRate> BuildDefaults()
    {
        var rates = new List<UnderwritingReferenceRate>();
        AddRates(
            rates,
            UnderwritingRateTypes.Mortality,
            "Mortality Master",
            MortalityUnitBasis,
            [
                ("Under 25", 0.02666666666666667m, 0.059166666666666666m),
                ("25-29", 0.03216666666666666m, 0.07483333333333334m),
                ("30-34", 0.041166666666666664m, 0.08616666666666667m),
                ("35-39", 0.05050000000000001m, 0.09566666666666668m),
                ("40-44", 0.07816666666666668m, 0.12949999999999998m),
                ("45-49", 0.12483333333333334m, 0.19216666666666668m),
                ("50-54", 0.1908333333333333m, 0.29283333333333333m),
                ("55-59", 0.2936666666666667m, 0.4543333333333333m),
                ("60-64", 0.4626666666666667m, 0.7163333333333334m),
                ("65-69", 0.7475m, 1.1468333333333334m)
            ]);
        AddRates(
            rates,
            UnderwritingRateTypes.Morbidity,
            "Morbidity Master",
            MorbidityUnitBasis,
            [
                ("Under 25", 0.3381068m, 0.23158m),
                ("25-29", 0.4722762m, 0.28974m),
                ("30-34", 0.809088m, 0.4704m),
                ("35-39", 1.0668528m, 0.65052m),
                ("40-44", 1.643863m, 0.9727m),
                ("45-49", 2.2753228m, 1.60234m),
                ("50-54", 2.9946318m, 2.62687m),
                ("55-59", 3.7544648m, 4.08094m),
                ("60-64", 3.3087352m, 3.71768m)
            ]);
        AddRates(
            rates,
            UnderwritingRateTypes.SoaBasicLifeMortality,
            "Death & Disability",
            SoaMortalityUnitBasis,
            [
                ("Under 25", 0.126591232536174m, 0.5815611568449806m),
                ("25-29", 0.12197606099061097m, 0.350981312647081m),
                ("30-34", 0.16931230388626914m, 0.3677433450701918m),
                ("35-39", 0.2527888535493852m, 0.43568997670088555m),
                ("40-44", 0.4046851127379416m, 0.58576797087722m),
                ("45-49", 0.6483739654372774m, 0.9699766764093085m),
                ("50-54", 0.8986209338862339m, 1.4664978745030506m),
                ("55-59", 1.3894657117962066m, 2.3638084828251222m),
                ("60-64", 2.106901160727606m, 3.661189178655246m),
                ("65-69", 3.5204015925398386m, 5.142268568321122m)
            ],
            SoaVersionName,
            SoaSourceName);
        AddRates(
            rates,
            UnderwritingRateTypes.SoaGroupLifeWaiver,
            "Death & Disability",
            SoaWaiverUnitBasis,
            [
                ("Under 25", 0.08907702513122426m, 0.12009521728140837m),
                ("25-29", 0.14801896332704295m, 0.1128277603457981m),
                ("30-34", 0.28227465682547903m, 0.164359786987985m),
                ("35-39", 0.4047105463080239m, 0.24838128530022735m),
                ("40-44", 0.6416839571590526m, 0.40410063114144573m),
                ("45-49", 0.862570848493921m, 0.5441742194125911m),
                ("50-54", 1.2880052905159698m, 0.9914337215866743m),
                ("55-59", 1.7787822538056905m, 1.5726006295170794m),
                ("60-64", 0.4163437751187688m, 0.39729635631026766m)
            ],
            SoaVersionName,
            SoaSourceName);
        return rates;
    }

    private static void AddRates(
        ICollection<UnderwritingReferenceRate> target,
        string rateType,
        string sourceSheet,
        string unitBasis,
        IReadOnlyList<(string AgeBand, decimal FemaleRate, decimal MaleRate)> rows,
        string versionName = VersionName,
        string sourceName = SourceName)
    {
        for (var index = 0; index < rows.Count; index++)
        {
            var row = rows[index];
            target.Add(Create(rateType, versionName, sourceName, sourceSheet, unitBasis, row.AgeBand, UnderwritingRateGenders.Female, row.FemaleRate, index));
            target.Add(Create(rateType, versionName, sourceName, sourceSheet, unitBasis, row.AgeBand, UnderwritingRateGenders.Male, row.MaleRate, index));
        }
    }

    private static UnderwritingReferenceRate Create(
        string rateType,
        string versionName,
        string sourceName,
        string sourceSheet,
        string unitBasis,
        string ageBand,
        string gender,
        decimal rate,
        int displayOrder)
    {
        return new UnderwritingReferenceRate
        {
            RateType = rateType,
            VersionName = versionName,
            AgeBand = ageBand,
            Gender = gender,
            Rate = rate,
            UnitBasis = unitBasis,
            SourceName = sourceName,
            SourceSheet = sourceSheet,
            DisplayOrder = displayOrder
        };
    }
}

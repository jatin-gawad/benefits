namespace RenewalReviewBuilder.App.Services.CanonicalClaims;

/// <summary>Keeps carrier selection before the canonical contract and away from consumers.</summary>
public sealed class CarrierClaimsCanonicalMapperRegistry
{
    private readonly IReadOnlyDictionary<string, ICarrierClaimsCanonicalMapper> _mappers;

    public CarrierClaimsCanonicalMapperRegistry()
    {
        var manulife = new ManulifeClaimsCanonicalMapper();
        var canadaLife = new CanadaLifeClaimsCanonicalMapper();
        var sunLife = new SunLifeClaimsCanonicalMapper();
        var groupHealth = new GroupHealthClaimsCanonicalMapper();
        var groupSource = new GroupHealthClaimsCanonicalMapper(
            "GroupSource",
            nameof(GroupSourcePdfParser),
            "GROUPSOURCE_CLAIMS_CANONICAL_V1");
        _mappers = new Dictionary<string, ICarrierClaimsCanonicalMapper>(StringComparer.OrdinalIgnoreCase)
        {
            ["Manulife"] = manulife,
            ["MFC"] = manulife,
            ["Canada Life"] = canadaLife,
            ["Great-West Life"] = canadaLife,
            ["Sun Life"] = sunLife,
            ["SunLife"] = sunLife,
            ["Sun Life Financial"] = sunLife,
            ["GroupHEALTH"] = groupHealth,
            ["Group Health"] = groupHealth,
            ["GroupPLUS"] = groupHealth,
            ["GroupSource"] = groupSource,
            ["Group Source"] = groupSource
        };
    }

    public ICarrierClaimsCanonicalMapper? Resolve(string carrier)
    {
        return _mappers.TryGetValue(carrier.Trim(), out var mapper) ? mapper : null;
    }

    public ICarrierClaimsCanonicalMapper<TParsedResult>? Resolve<TParsedResult>(string carrier)
    {
        return Resolve(carrier) as ICarrierClaimsCanonicalMapper<TParsedResult>;
    }
}


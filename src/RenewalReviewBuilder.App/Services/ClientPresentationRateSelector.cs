using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public static class ClientPresentationRateSelector
{
    public static decimal? CurrentRate(RenewalVersionRateLine line)
    {
        return Select(
            line.ClientPresentationRateBasis,
            line.CurrentNetRate,
            line.CurrentGrossRate,
            line.CurrentRate);
    }

    public static decimal? ProposedRate(RenewalVersionRateLine line)
    {
        return Select(
            line.ClientPresentationRateBasis,
            line.ProposedNetRate,
            line.ProposedGrossRate,
            line.ProposedRate);
    }

    public static decimal? VersionRate(RenewalVersionRateLine line)
    {
        var fallback = line.VersionRate ?? line.ProposedRate;
        return Select(line.ClientPresentationRateBasis, line.VersionNetRate, line.VersionGrossRate, fallback);
    }

    private static decimal? Select(string basis, decimal? net, decimal? gross, decimal? carrierStated)
    {
        if (string.Equals(basis, RateValueBasisNames.Gross, StringComparison.OrdinalIgnoreCase))
        {
            return gross ?? throw new InvalidOperationException("A client presentation rate is designated as Gross, but its explicit Gross value is missing.");
        }

        if (string.Equals(basis, RateValueBasisNames.Net, StringComparison.OrdinalIgnoreCase))
        {
            return net ?? throw new InvalidOperationException("A client presentation rate is designated as Net, but its explicit Net value is missing.");
        }

        return carrierStated;
    }

    public static IReadOnlyList<string> ValidateSelectedVersion(AnalysisResult analysis)
    {
        var version = analysis.SelectedRenewalVersion;
        if (version is null)
        {
            return [];
        }

        var errors = new List<string>();
        foreach (var line in version.RateLines)
        {
            if (!string.Equals(line.ClientPresentationRateBasis, RateValueBasisNames.Gross, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var label = string.IsNullOrWhiteSpace(line.TierOrClass)
                ? line.Benefit
                : $"{line.Benefit} / {line.TierOrClass}";
            if (!line.CurrentGrossRate.HasValue)
            {
                errors.Add($"{label}: Current Gross rate is missing.");
            }

            if (!line.ProposedGrossRate.HasValue)
            {
                errors.Add($"{label}: Renewal Gross rate is missing.");
            }

            if (!line.VersionGrossRate.HasValue)
            {
                errors.Add($"{label}: Negotiated/Final Gross rate is missing.");
            }
        }

        return errors;
    }
}

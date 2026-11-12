using System.Text.RegularExpressions;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public static class DemographicAgeBandNormalizer
{
    public static string[] GetLifeAgeBands(IReadOnlyCollection<DemographicLine> demographics)
    {
        return UsesCombinedLifeSeniorBucket(demographics)
            ? ["Under 25", "25-29", "30-34", "35-39", "40-44", "45-49", "50-54", "55-59", "60-64", "65+"]
            : ["Under 25", "25-29", "30-34", "35-39", "40-44", "45-49", "50-54", "55-59", "60-64", "65-69", "70 & over"];
    }

    public static string[] GetLtdAgeBands(IReadOnlyCollection<DemographicLine> demographics)
    {
        return UsesCombinedLtdSeniorBucket(demographics)
            ? ["Under 25", "25-29", "30-34", "35-39", "40-44", "45-49", "50-54", "55-59", "60 & over"]
            : ["Under 25", "25-29", "30-34", "35-39", "40-44", "45-49", "50-54", "55-59", "60-64", "65+"];
    }

    public static bool UsesCombinedLifeSeniorBucket(IReadOnlyCollection<DemographicLine> demographics)
    {
        return demographics.Any(row => NormalizeLifeAgeBand(row.AgeBand).Equals("65+", StringComparison.OrdinalIgnoreCase));
    }

    public static bool UsesCombinedLtdSeniorBucket(IReadOnlyCollection<DemographicLine> demographics)
    {
        return demographics.Any(row => NormalizeLtdAgeBand(row.AgeBand).Equals("60 & over", StringComparison.OrdinalIgnoreCase));
    }

    public static string NormalizeLifeAgeBand(string? ageBand)
    {
        if (string.IsNullOrWhiteSpace(ageBand))
        {
            return "";
        }

        var text = ageBand.Trim()
            .Replace('\u2013', '-')
            .Replace('\u2014', '-')
            .Replace('\u2212', '-');
        text = Regex.Replace(text, @"\s+", " ");
        var compact = Regex.Replace(text.ToLowerInvariant(), @"[\s_]", "")
            .Replace("&", "and")
            .Replace("-and", "and");

        if (compact is "<20" or "0-19" or "20-24" or "<25" or "under25" or "underage25")
        {
            return "Under 25";
        }

        var rangeMatch = Regex.Match(compact, @"^(?<start>\d{2})-(?<end>\d{2})\+?$");
        if (rangeMatch.Success)
        {
            if (int.TryParse(rangeMatch.Groups["start"].Value, out var start) && start >= 70)
            {
                return "70 & over";
            }

            return $"{rangeMatch.Groups["start"].Value}-{rangeMatch.Groups["end"].Value}";
        }

        if (compact is "65+" or "65plus" or "65andover" or "65andolder" or "65orolder" or "65older" or "65over" or ">=65" or ">65")
        {
            return "65+";
        }

        if (compact is "70+" or "70plus" or "70andover" or "70andolder" or "70orolder" or "70older" or "70over" or ">=70" or ">70")
        {
            return "70 & over";
        }

        return text;
    }

    public static string NormalizeLtdAgeBand(string? ageBand)
    {
        if (string.IsNullOrWhiteSpace(ageBand))
        {
            return "";
        }

        var normalized = NormalizeLifeAgeBand(ageBand);
        var compact = Regex.Replace(normalized.ToLowerInvariant(), @"[\s_]", "")
            .Replace("&", "and")
            .Replace("-and", "and");

        if (compact is "60+" or "60plus" or "60andover" or "60andolder" or "60orolder" or "60older" or "60over" or ">=60" or ">60")
        {
            return "60 & over";
        }

        if (compact is "65+" or "65plus" or "65andover" or "65andolder" or "65orolder" or "65older" or "65over" or ">=65" or ">65")
        {
            return "65+";
        }

        return normalized;
    }

    public static string SelectLifeChartBucket(string? ageBand, bool usesCombinedSeniorBucket)
    {
        var normalized = NormalizeLifeAgeBand(ageBand);
        if (usesCombinedSeniorBucket && normalized is "65-69" or "70 & over" or "65+")
        {
            return "65+";
        }

        return normalized;
    }
}

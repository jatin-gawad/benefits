using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public static class ExperienceLineSelector
{
    public static ExperienceLine? FindBestMatch(IEnumerable<ExperienceLine> experienceLines, string benefit)
    {
        return experienceLines
            .Select(line => new
            {
                Line = line,
                Score = BenefitMatcher.MatchScore(line.Benefit, benefit),
                PeriodStart = PeriodStart(line)
            })
            .Where(item => item.Score > 0)
            .OrderByDescending(item => item.Score)
            .ThenByDescending(item => item.PeriodStart ?? DateOnly.MinValue)
            .FirstOrDefault()
            ?.Line;
    }

    public static DateOnly? PeriodStart(ExperienceLine line)
    {
        var periodStart = line.Period.Split(" - ", StringSplitOptions.None).FirstOrDefault();
        return ParsingHelpers.ParseDateOnly(periodStart);
    }
}

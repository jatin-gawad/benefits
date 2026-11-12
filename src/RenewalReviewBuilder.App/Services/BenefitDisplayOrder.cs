namespace RenewalReviewBuilder.App.Services;

internal static class BenefitDisplayOrder
{
    private const int MiscellaneousBenefitRank = 8;

    public static List<T> Sort<T>(
        IEnumerable<T> rows,
        Func<T, string?> benefitSelector,
        Func<T, string?>? tierSelector = null,
        Func<T, string?>? accountClassSelector = null)
    {
        return rows
            .Select((row, index) => new
            {
                Row = row,
                Index = index,
                Benefit = benefitSelector(row) ?? "",
                Tier = tierSelector?.Invoke(row) ?? "",
                AccountClass = accountClassSelector?.Invoke(row) ?? ""
            })
            .OrderBy(item => BenefitRank(item.Benefit))
            .ThenBy(item => Normalize(item.Benefit), StringComparer.Ordinal)
            .ThenBy(item => Normalize(item.AccountClass), StringComparer.Ordinal)
            .ThenBy(item => TierRank(item.Tier))
            .ThenBy(item => Normalize(item.Tier), StringComparer.Ordinal)
            .ThenBy(item => item.Index)
            .Select(item => item.Row)
            .ToList();
    }

    public static int BenefitRank(string? benefit)
    {
        var normalized = Normalize(benefit);
        if (string.IsNullOrWhiteSpace(normalized) || IsOptionalCoverage(normalized))
        {
            return MiscellaneousBenefitRank;
        }

        if (IsDependentLife(normalized))
        {
            return 2;
        }

        if (IsAccidentalDeathAndDismemberment(normalized))
        {
            return 1;
        }

        if (IsCriticalIllness(normalized))
        {
            return 3;
        }

        if (IsShortTermDisability(normalized))
        {
            return 4;
        }

        if (IsLongTermDisability(normalized))
        {
            return 5;
        }

        if (IsExtendedHealthCare(normalized))
        {
            return 6;
        }

        if (IsDental(normalized))
        {
            return 7;
        }

        return IsEmployeeLife(normalized) ? 0 : MiscellaneousBenefitRank;
    }

    public static int TierRank(string? tier)
    {
        var normalized = Normalize(tier);
        if (normalized.Contains("single parent", StringComparison.Ordinal))
        {
            return 1;
        }

        if (normalized.Contains("single", StringComparison.Ordinal) ||
            normalized.Equals("employee only", StringComparison.Ordinal) ||
            normalized.Equals("member only", StringComparison.Ordinal))
        {
            return 0;
        }

        if (normalized.Contains("couple", StringComparison.Ordinal) ||
            normalized.Contains("employee spouse", StringComparison.Ordinal) ||
            normalized.Contains("two person", StringComparison.Ordinal))
        {
            return 2;
        }

        if (normalized.Contains("family", StringComparison.Ordinal))
        {
            return 3;
        }

        if (string.IsNullOrWhiteSpace(normalized) ||
            normalized.Equals("all", StringComparison.Ordinal) ||
            normalized.Equals("all classes", StringComparison.Ordinal))
        {
            return 90;
        }

        return 50;
    }

    private static bool IsOptionalCoverage(string benefit)
    {
        return benefit.Contains("optional", StringComparison.Ordinal) ||
               benefit.Contains("voluntary", StringComparison.Ordinal);
    }

    private static bool IsDependentLife(string benefit)
    {
        return benefit.Contains("dependent life", StringComparison.Ordinal) ||
               benefit.Contains("dependant life", StringComparison.Ordinal) ||
               benefit.Contains("dependent term life", StringComparison.Ordinal) ||
               benefit.Contains("dependant term life", StringComparison.Ordinal) ||
               benefit.Contains("spouse life", StringComparison.Ordinal) ||
               benefit.Contains("spousal life", StringComparison.Ordinal) ||
               benefit.Contains("child life", StringComparison.Ordinal) ||
               benefit.Equals("dep life", StringComparison.Ordinal);
    }

    private static bool IsAccidentalDeathAndDismemberment(string benefit)
    {
        return benefit.Equals("ad d", StringComparison.Ordinal) ||
               benefit.Equals("add", StringComparison.Ordinal) ||
               benefit.Contains("accidental death", StringComparison.Ordinal) ||
               benefit.Contains("dismemberment", StringComparison.Ordinal);
    }

    private static bool IsCriticalIllness(string benefit)
    {
        return benefit.Equals("ci", StringComparison.Ordinal) ||
               benefit.Contains("critical illness", StringComparison.Ordinal);
    }

    private static bool IsShortTermDisability(string benefit)
    {
        return benefit.Equals("std", StringComparison.Ordinal) ||
               benefit.Contains("short term disability", StringComparison.Ordinal) ||
               benefit.Contains("weekly indemnity", StringComparison.Ordinal);
    }

    private static bool IsLongTermDisability(string benefit)
    {
        return benefit.Equals("ltd", StringComparison.Ordinal) ||
               benefit.Contains("long term disability", StringComparison.Ordinal);
    }

    private static bool IsExtendedHealthCare(string benefit)
    {
        if (benefit.Contains("virtual", StringComparison.Ordinal) ||
            benefit.Contains("health spending", StringComparison.Ordinal))
        {
            return false;
        }

        return benefit.Equals("ehc", StringComparison.Ordinal) ||
               benefit.Equals("health", StringComparison.Ordinal) ||
               benefit.Equals("health care", StringComparison.Ordinal) ||
               benefit.Equals("healthcare", StringComparison.Ordinal) ||
               benefit.Contains("extended health care", StringComparison.Ordinal);
    }

    private static bool IsDental(string benefit)
    {
        return benefit.Equals("dental", StringComparison.Ordinal) ||
               benefit.Contains("dental care", StringComparison.Ordinal);
    }

    private static bool IsEmployeeLife(string benefit)
    {
        return benefit.Equals("life", StringComparison.Ordinal) ||
               benefit.Contains("basic life", StringComparison.Ordinal) ||
               benefit.Contains("employee life", StringComparison.Ordinal) ||
               benefit.Contains("employee term life", StringComparison.Ordinal) ||
               benefit.Contains("group life", StringComparison.Ordinal) ||
               benefit.Equals("life insurance", StringComparison.Ordinal);
    }

    private static string Normalize(string? value)
    {
        return BenefitMatcher.NormalizeWords(value ?? "");
    }
}

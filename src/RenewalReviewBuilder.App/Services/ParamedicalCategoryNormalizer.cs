namespace RenewalReviewBuilder.App.Services;

public static class ParamedicalCategoryNormalizer
{
    public static string Normalize(string value)
    {
        var cleaned = ParsingHelpers.Clean(value);
        var key = NormalizeKey(cleaned);
        if (string.IsNullOrWhiteSpace(key))
        {
            return "Other Paramedical";
        }

        if (key is "total")
        {
            return "Total";
        }

        if (key is "paramedical" or "paramedicalservices" or "paramedicaltotal")
        {
            return "Paramedical";
        }

        if (key.Contains("physio", StringComparison.Ordinal))
        {
            return "Physiotherapy";
        }

        if (key.Contains("psycholog", StringComparison.Ordinal) ||
            key.Contains("psychologist", StringComparison.Ordinal) ||
            key.Contains("psychotherapist", StringComparison.Ordinal))
        {
            return "Psychologist";
        }

        if (key.Contains("massage", StringComparison.Ordinal) ||
            key.Contains("masseur", StringComparison.Ordinal))
        {
            return "Massage";
        }

        if (key.Contains("chiro", StringComparison.Ordinal))
        {
            return "Chiropractic";
        }

        if (key.Contains("naturopath", StringComparison.Ordinal))
        {
            return "Naturopath";
        }

        if (key.Contains("acupuncture", StringComparison.Ordinal))
        {
            return "Acupuncture";
        }

        if (key.Contains("clinicalcounsel", StringComparison.Ordinal))
        {
            return "Clinical Counsellor";
        }

        if (key.Contains("socialworker", StringComparison.Ordinal))
        {
            return "Social Worker";
        }

        if (key.Contains("osteopath", StringComparison.Ordinal))
        {
            return "Osteopath";
        }

        if (key.Contains("podiat", StringComparison.Ordinal) ||
            key.Contains("chiropod", StringComparison.Ordinal))
        {
            return "Podiatrist";
        }

        if (key.Contains("dietitian", StringComparison.Ordinal) ||
            key.Contains("dietician", StringComparison.Ordinal))
        {
            return "Dietitian";
        }

        if (key.Contains("speech", StringComparison.Ordinal))
        {
            return "Speech Therapy";
        }

        return cleaned;
    }

    public static bool IsKnownParamedicalCategory(string value)
    {
        var normalized = Normalize(value);
        return !normalized.Equals(ParsingHelpers.Clean(value), StringComparison.OrdinalIgnoreCase) ||
               normalized is "Physiotherapy" or "Psychologist" or "Massage" or "Chiropractic" or "Naturopath" or
                   "Acupuncture" or "Clinical Counsellor" or "Social Worker" or "Osteopath" or "Podiatrist" or
                   "Dietitian" or "Speech Therapy";
    }

    public static bool IsRollup(string value)
    {
        var key = NormalizeKey(value);
        return key is "paramedical" or "paramedicalservices" or "paramedicaltotal" or "total";
    }

    private static string NormalizeKey(string value)
    {
        return new string(ParsingHelpers.Clean(value)
            .Where(char.IsLetterOrDigit)
            .Select(char.ToLowerInvariant)
            .ToArray());
    }
}

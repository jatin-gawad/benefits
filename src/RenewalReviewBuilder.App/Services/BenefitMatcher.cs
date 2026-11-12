using System.Text.RegularExpressions;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public static class BenefitMatcher
{
    public static string ResolveCanonicalBenefitTypeCode(string sourceBenefit)
    {
        var value = NormalizeWords(sourceBenefit);
        if (string.IsNullOrWhiteSpace(value))
        {
            return "";
        }

        if (value.Contains("optional accident serious illness", StringComparison.Ordinal) ||
            value.Contains("optional accident and serious illness", StringComparison.Ordinal))
        {
            return CanonicalBenefitTypeCodes.OptionalAccidentSeriousIllness;
        }

        if (value.Contains("optional critical illness", StringComparison.Ordinal))
        {
            return CanonicalBenefitTypeCodes.OptionalCriticalIllness;
        }

        if (value.Contains("optional life insurance", StringComparison.Ordinal))
        {
            return CanonicalBenefitTypeCodes.OptionalLife;
        }

        if (value.Contains("accidental death and dismemberment", StringComparison.Ordinal) || value == "ad d")
        {
            return CanonicalBenefitTypeCodes.AdAndD;
        }

        if (value.Contains("long term disability", StringComparison.Ordinal) || value == "ltd")
        {
            return CanonicalBenefitTypeCodes.Ltd;
        }

        if (value.Contains("extended health care", StringComparison.Ordinal) || value == "ehc")
        {
            return CanonicalBenefitTypeCodes.Ehc;
        }

        if (value.Contains("travel insurance", StringComparison.Ordinal) || value.Contains("travel cancellation", StringComparison.Ordinal))
        {
            return CanonicalBenefitTypeCodes.Travel;
        }

        if (value.Contains("medical second opinion", StringComparison.Ordinal))
        {
            return CanonicalBenefitTypeCodes.MedicalSecondOpinion;
        }

        if (value.Contains("virtual healthcare", StringComparison.Ordinal) || value.Contains("virtual health care", StringComparison.Ordinal))
        {
            return CanonicalBenefitTypeCodes.VirtualHealthcare;
        }

        if (value.Contains("dental care", StringComparison.Ordinal) || value == "dental")
        {
            return CanonicalBenefitTypeCodes.Dental;
        }

        return value.Contains("group life insurance", StringComparison.Ordinal) || value == "life"
            ? CanonicalBenefitTypeCodes.Life
            : "";
    }

    public static bool IsMatch(string candidateBenefit, string targetBenefit)
    {
        var candidate = NormalizeWords(candidateBenefit);
        var target = NormalizeWords(targetBenefit);
        if (string.IsNullOrWhiteSpace(candidate) || string.IsNullOrWhiteSpace(target))
        {
            return false;
        }

        return candidate.Equals(target, StringComparison.OrdinalIgnoreCase) ||
               ContainsPhrase(candidate, target) ||
               ContainsPhrase(target, candidate);
    }

    public static bool ContainsPhrase(string value, string phrase)
    {
        var normalizedValue = NormalizeWords(value);
        var normalizedPhrase = NormalizeWords(phrase);
        if (string.IsNullOrWhiteSpace(normalizedValue) || string.IsNullOrWhiteSpace(normalizedPhrase))
        {
            return false;
        }

        return $" {normalizedValue} ".Contains($" {normalizedPhrase} ", StringComparison.OrdinalIgnoreCase);
    }

    public static int MatchScore(string candidateBenefit, string targetBenefit)
    {
        var candidate = NormalizeWords(candidateBenefit);
        var target = NormalizeWords(targetBenefit);
        if (candidate.Equals(target, StringComparison.OrdinalIgnoreCase))
        {
            return 100;
        }

        if (ContainsPhrase(candidate, target))
        {
            return 80;
        }

        if (ContainsPhrase(target, candidate))
        {
            return 70;
        }

        return 0;
    }

    public static string NormalizeWords(string value)
    {
        var normalized = Regex.Replace(value ?? "", "[^A-Za-z0-9]+", " ");
        return Regex.Replace(normalized, "\\s+", " ").Trim().ToLowerInvariant();
    }
}

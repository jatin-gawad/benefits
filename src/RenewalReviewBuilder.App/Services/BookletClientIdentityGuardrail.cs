using System.Text.RegularExpressions;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public static class BookletClientIdentityGuardrail
{
    private static readonly HashSet<string> IgnoredClientTokens = new(StringComparer.OrdinalIgnoreCase)
    {
        "THE", "AND", "LTD", "LIMITED", "INC", "INCORPORATED", "CORP", "CORPORATION",
        "COMPANY", "CO", "GROUP", "BENEFITS"
    };

    public static bool IsConfirmed(BookletClientIdentity? identity)
    {
        var acceptedIdentitySource = identity?.IdentitySource.Equals("Document", StringComparison.OrdinalIgnoreCase) == true ||
                                     identity?.IdentitySource.Equals("Structured file name", StringComparison.OrdinalIgnoreCase) == true ||
                                     identity?.IdentitySource.Equals("Document and structured file name", StringComparison.OrdinalIgnoreCase) == true;
        if (identity is null ||
            !acceptedIdentitySource ||
            identity.ConfidenceScore < 0.80m ||
            ClientNameResolver.IsInvalidClientName(identity.ClientName))
        {
            return false;
        }

        var normalized = NormalizeGroupKey(identity.ClientName);
        return normalized.Length >= 3 &&
               !normalized.Contains("BOOKLET", StringComparison.Ordinal) &&
               !normalized.Contains("SCHEDULEOFBENEFITS", StringComparison.Ordinal) &&
               !normalized.Contains("EMPLOYEEBENEFITS", StringComparison.Ordinal) &&
               !normalized.Contains("GROUPBENEFITS", StringComparison.Ordinal);
    }

    public static bool GroupKeysLikelyMatch(string left, string right)
    {
        if (left.Equals(right, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        var sharedMinimumLength = Math.Min(left.Length, right.Length);
        return sharedMinimumLength >= 12 &&
               (left.Contains(right, StringComparison.OrdinalIgnoreCase) ||
                right.Contains(left, StringComparison.OrdinalIgnoreCase));
    }

    public static bool ClientNamesLikelyMatch(string expected, string detected)
    {
        var expectedNormalized = NormalizeClientName(expected);
        var detectedNormalized = NormalizeClientName(detected);
        if (string.IsNullOrWhiteSpace(expectedNormalized) || string.IsNullOrWhiteSpace(detectedNormalized))
        {
            return false;
        }

        if (expectedNormalized.Equals(detectedNormalized, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        var sharedMinimumLength = Math.Min(expectedNormalized.Length, detectedNormalized.Length);
        return sharedMinimumLength >= 8 &&
               (expectedNormalized.Contains(detectedNormalized, StringComparison.OrdinalIgnoreCase) ||
                detectedNormalized.Contains(expectedNormalized, StringComparison.OrdinalIgnoreCase));
    }

    public static string NormalizeGroupKey(string value)
    {
        return new string((value ?? "")
            .ToUpperInvariant()
            .Where(char.IsLetterOrDigit)
            .ToArray());
    }

    private static string NormalizeClientName(string value)
    {
        return string.Concat(
            Regex.Matches(value ?? "", @"[A-Za-z0-9]+")
                .Cast<Match>()
                .Select(match => match.Value.ToUpperInvariant())
                .Where(token => !IgnoredClientTokens.Contains(token))
                .Where(token => !token.All(char.IsDigit)));
    }
}

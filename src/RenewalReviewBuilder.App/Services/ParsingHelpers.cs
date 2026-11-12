using System.Globalization;
using System.Text.RegularExpressions;

namespace RenewalReviewBuilder.App.Services;

public static class ParsingHelpers
{
    private static readonly Regex NumberRegex = new(
        @"\(?-?\$?\d{1,3}(?:,\d{3})*(?:\.\d+)?\)?|-?\d+(?:\.\d+)?",
        RegexOptions.Compiled);

    public static string Clean(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return "";
        }

        return Regex.Replace(value.Replace('\u00a0', ' '), @"\s+", " ").Trim();
    }

    public static decimal? ParseDecimal(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        var cleaned = value.Trim()
            .Replace("$", "", StringComparison.Ordinal)
            .Replace("%", "", StringComparison.Ordinal)
            .Replace(",", "", StringComparison.Ordinal);

        var isNegativeParentheses = cleaned.StartsWith('(') && cleaned.EndsWith(')');
        cleaned = cleaned.Trim('(', ')');

        if (decimal.TryParse(
                cleaned,
                NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint | NumberStyles.AllowThousands,
                CultureInfo.InvariantCulture,
                out var result))
        {
            return isNegativeParentheses ? -result : result;
        }

        return null;
    }

    public static int? ParseInt(string? value)
    {
        var number = ParseDecimal(value);
        return number.HasValue ? (int)Math.Round(number.Value, MidpointRounding.AwayFromZero) : null;
    }

    public static List<decimal> ExtractNumbers(string text)
    {
        return NumberRegex.Matches(text)
            .Select(match => ParseDecimal(match.Value))
            .Where(value => value.HasValue)
            .Select(value => value!.Value)
            .ToList();
    }

    public static DateOnly? ParseDateOnly(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        var formats = new[]
        {
            "dd-MMM-yyyy",
            "d-MMM-yyyy",
            "dd-MMMM-yyyy",
            "d-MMMM-yyyy",
            "MMMM dd, yyyy",
            "MMMM d, yyyy",
            "MMM dd, yyyy",
            "MMM d, yyyy",
            "yyyy-MM-dd",
            "M/d/yyyy",
            "MM/dd/yyyy"
        };
        if (DateOnly.TryParseExact(value.Trim(), formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out var date))
        {
            return date;
        }

        return null;
    }

    public static string Money(decimal? value)
    {
        return value.HasValue ? value.Value.ToString("C2", CultureInfo.CurrentCulture) : "";
    }

    public static string Percent(decimal? value)
    {
        return value.HasValue ? $"{value.Value:0.0}%" : "";
    }
}

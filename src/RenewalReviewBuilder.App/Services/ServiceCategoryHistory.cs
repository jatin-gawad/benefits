using System.Globalization;
using System.Text.RegularExpressions;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

/// <summary>
/// Owns carrier-neutral dated history for claims service categories while keeping
/// the legacy prior/current properties synchronized for existing consumers.
/// </summary>
public static class ServiceCategoryHistory
{
    public static IReadOnlyList<ServiceCategoryPeriodValue> GetPeriods(ServiceCategoryTrend row, int maximum = 3)
    {
        var values = new List<ServiceCategoryPeriodValue>();
        AddOrReplace(values, row.PriorPeriodLabel, row.PriorAmount, row.SourceReference);
        AddOrReplace(values, row.CurrentPeriodLabel, row.CurrentAmount, row.SourceReference);
        foreach (var period in row.Periods)
        {
            AddOrReplace(values, period.PeriodLabel, period.Amount, period.SourceReference);
        }

        var ordered = values
            .OrderBy(value => ParsePeriodStart(value.PeriodLabel) ?? DateOnly.MinValue)
            .ThenBy(value => value.PeriodLabel, StringComparer.OrdinalIgnoreCase)
            .ToList();
        return SelectWindow(ordered, Math.Max(1, maximum));
    }

    public static bool Merge(ServiceCategoryTrend target, ServiceCategoryTrend source)
    {
        var before = GetPeriods(target, int.MaxValue).Select(PeriodKey).ToHashSet(StringComparer.OrdinalIgnoreCase);
        foreach (var period in GetPeriods(source, int.MaxValue))
        {
            Upsert(target, period.PeriodLabel, period.Amount, period.SourceReference);
        }

        target.Benefit = FirstNonBlank(source.Benefit, target.Benefit);
        target.Category = FirstNonBlank(source.Category, target.Category);
        target.SourceReference = MergeSource(target.SourceReference, source.SourceReference);
        SynchronizeLegacyWindow(target);
        return GetPeriods(target, int.MaxValue).Any(period => !before.Contains(PeriodKey(period)));
    }

    public static void AlignToSharedWindow(IEnumerable<ServiceCategoryTrend> rows, int maximum = 3)
    {
        var materialized = rows.ToList();
        if (materialized.Count == 0) return;

        var sharedPeriods = materialized
            .SelectMany(row => GetPeriods(row, int.MaxValue))
            .Where(period => !string.IsNullOrWhiteSpace(period.PeriodLabel))
            .GroupBy(period => PeriodIdentity(period.PeriodLabel), StringComparer.OrdinalIgnoreCase)
            .Select(group => group.First())
            .OrderBy(period => ParsePeriodStart(period.PeriodLabel) ?? DateOnly.MinValue)
            .ThenBy(period => period.PeriodLabel, StringComparer.OrdinalIgnoreCase)
            .TakeLast(Math.Max(1, maximum))
            .ToList();
        if (sharedPeriods.Count == 0) return;

        foreach (var row in materialized)
        {
            foreach (var sharedPeriod in sharedPeriods)
            {
                if (!GetPeriods(row, int.MaxValue).Any(period => SamePeriod(period.PeriodLabel, sharedPeriod.PeriodLabel)))
                {
                    Upsert(row, sharedPeriod.PeriodLabel, 0m);
                }
            }

            // Keep the newest shared period as current, including an explicit zero
            // when a category is absent. Legacy prior remains the latest actual
            // non-zero history so carrier-change periods do not displace it.
            SynchronizeLegacyWindow(row);
        }
    }

    public static void Upsert(ServiceCategoryTrend row, string label, decimal amount, string sourceReference = "")
    {
        if (string.IsNullOrWhiteSpace(label)) return;
        var key = Normalize(label);
        var existing = row.Periods.FirstOrDefault(period => Normalize(period.PeriodLabel).Equals(key, StringComparison.OrdinalIgnoreCase));
        if (existing is null)
        {
            row.Periods.Add(new ServiceCategoryPeriodValue { PeriodLabel = ParsingHelpers.Clean(label), Amount = amount, SourceReference = sourceReference });
        }
        else
        {
            existing.Amount = amount;
            existing.SourceReference = MergeSource(existing.SourceReference, sourceReference);
        }

        SynchronizeLegacyWindow(row);
    }

    public static void SynchronizeLegacyWindow(ServiceCategoryTrend row)
    {
        var values = new List<ServiceCategoryPeriodValue>();
        AddOrReplace(values, row.PriorPeriodLabel, row.PriorAmount, row.SourceReference);
        AddOrReplace(values, row.CurrentPeriodLabel, row.CurrentAmount, row.SourceReference);
        foreach (var period in row.Periods) AddOrReplace(values, period.PeriodLabel, period.Amount, period.SourceReference);
        values = values.OrderBy(value => ParsePeriodStart(value.PeriodLabel) ?? DateOnly.MinValue).ThenBy(value => value.PeriodLabel, StringComparer.OrdinalIgnoreCase).ToList();
        row.Periods.Clear();
        row.Periods.AddRange(values);
        if (values.Count == 0) return;
        var current = values[^1];
        var previousValues = values.Take(values.Count - 1).ToList();
        var prior = previousValues.LastOrDefault(value => value.Amount != 0m) ?? previousValues.LastOrDefault() ?? new ServiceCategoryPeriodValue();
        row.PriorPeriodLabel = prior.PeriodLabel;
        row.PriorAmount = prior.Amount;
        row.CurrentPeriodLabel = current.PeriodLabel;
        row.CurrentAmount = current.Amount;
    }

    public static DateOnly? ParsePeriodStart(string? label)
    {
        if (string.IsNullOrWhiteSpace(label)) return null;
        var dayMonthYear = Regex.Match(label, @"(?<day>\d{1,2})[-\s]+(?<month>Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Sept(?:ember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)[-\s,]+(?<year>\d{4})", RegexOptions.IgnoreCase);
        if (dayMonthYear.Success && DateOnly.TryParse(
                $"{dayMonthYear.Groups["month"].Value} {dayMonthYear.Groups["day"].Value}, {dayMonthYear.Groups["year"].Value}",
                CultureInfo.InvariantCulture,
                DateTimeStyles.AllowWhiteSpaces,
                out var dayFirstDate))
        {
            return dayFirstDate;
        }

        var match = Regex.Match(label, @"(?<month>Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Sept(?:ember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\.?\s+(?<day>\d{1,2}),?\s+(?<year>\d{4})", RegexOptions.IgnoreCase);
        if (match.Success && DateOnly.TryParse($"{match.Groups["month"].Value} {match.Groups["day"].Value}, {match.Groups["year"].Value}", CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out var dated)) return dated;
        var monthYear = Regex.Match(label, @"(?<month>Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Sept(?:ember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\.?\s+(?<year>\d{4})", RegexOptions.IgnoreCase);
        return monthYear.Success && DateOnly.TryParse($"{monthYear.Groups["month"].Value} 1, {monthYear.Groups["year"].Value}", CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out var monthDate) ? monthDate : null;
    }

    private static IReadOnlyList<ServiceCategoryPeriodValue> SelectWindow(List<ServiceCategoryPeriodValue> ordered, int maximum)
    {
        if (ordered.Count <= maximum) return ordered;

        var selected = new List<ServiceCategoryPeriodValue> { ordered[^1] };
        foreach (var value in ordered.Take(ordered.Count - 1).Reverse().Where(value => value.Amount != 0m))
        {
            if (selected.Count >= maximum) break;
            selected.Add(value);
        }

        foreach (var value in ordered.Take(ordered.Count - 1).Reverse())
        {
            if (selected.Count >= maximum) break;
            if (!selected.Contains(value)) selected.Add(value);
        }

        return selected
            .OrderBy(value => ParsePeriodStart(value.PeriodLabel) ?? DateOnly.MinValue)
            .ThenBy(value => value.PeriodLabel, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }
    private static void AddOrReplace(List<ServiceCategoryPeriodValue> values, string label, decimal amount, string sourceReference)
    {
        if (string.IsNullOrWhiteSpace(label)) return;
        var key = Normalize(label);
        var existing = values.FirstOrDefault(value => Normalize(value.PeriodLabel).Equals(key, StringComparison.OrdinalIgnoreCase));
        if (existing is null) values.Add(new ServiceCategoryPeriodValue { PeriodLabel = ParsingHelpers.Clean(label), Amount = amount, SourceReference = sourceReference });
        else { existing.Amount = amount; existing.SourceReference = MergeSource(existing.SourceReference, sourceReference); }
    }

    private static string PeriodKey(ServiceCategoryPeriodValue period) => $"{Normalize(period.PeriodLabel)}|{period.Amount.ToString(CultureInfo.InvariantCulture)}";
    private static bool SamePeriod(string first, string second) => PeriodIdentity(first).Equals(PeriodIdentity(second), StringComparison.OrdinalIgnoreCase);
    private static string PeriodIdentity(string label) => Normalize(label);
    private static string Normalize(string value) => new(value.Where(char.IsLetterOrDigit).Select(char.ToUpperInvariant).ToArray());
    private static string FirstNonBlank(params string[] values) => values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)) ?? "";
    private static string MergeSource(string first, string second) => string.Join("; ", new[] { first, second }.Where(value => !string.IsNullOrWhiteSpace(value)).Distinct(StringComparer.OrdinalIgnoreCase));
}

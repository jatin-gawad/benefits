using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services.FlexibleBenefits;

public sealed class FlexibleBenefitNormalizationService
{
    public FlexibleBenefitPlanType NormalizePlanType(bool? hcsaEnabled, bool? lsaEnabled, string? sourceText = null)
    {
        var normalized = NormalizeSearchText(sourceText);
        if (hcsaEnabled == true && lsaEnabled == true)
        {
            return FlexibleBenefitPlanType.CombinedHcsaAndLsa;
        }

        if (hcsaEnabled == true)
        {
            return FlexibleBenefitPlanType.Hcsa;
        }

        if (lsaEnabled == true)
        {
            return normalized.Contains("wellnessspendingaccount", StringComparison.Ordinal)
                ? FlexibleBenefitPlanType.WellnessSpendingAccount
                : FlexibleBenefitPlanType.Lsa;
        }

        if (normalized.Contains("personalspendingaccount", StringComparison.Ordinal))
        {
            return FlexibleBenefitPlanType.PersonalSpendingAccount;
        }

        return FlexibleBenefitPlanType.Other;
    }

    public FlexibleBenefitYearType NormalizeBenefitYear(string? value)
    {
        var normalized = NormalizeSearchText(value);
        if (normalized.Contains("januarytodecember", StringComparison.Ordinal) ||
            normalized.Contains("calendaryear", StringComparison.Ordinal))
        {
            return FlexibleBenefitYearType.CalendarYear;
        }

        if (normalized.Contains("fiscal", StringComparison.Ordinal))
        {
            return FlexibleBenefitYearType.FiscalYear;
        }

        if (normalized.Contains("policyyear", StringComparison.Ordinal))
        {
            return FlexibleBenefitYearType.PolicyYear;
        }

        return string.IsNullOrWhiteSpace(value)
            ? FlexibleBenefitYearType.Unknown
            : FlexibleBenefitYearType.Custom;
    }

    public FlexibleBenefitCarryForwardType NormalizeCarryForward(string? value)
    {
        var normalized = NormalizeSearchText(value);
        var hasCredit = normalized.Contains("creditcarry", StringComparison.Ordinal);
        var hasExpense = normalized.Contains("expensecarry", StringComparison.Ordinal);
        if (hasCredit && hasExpense)
        {
            return FlexibleBenefitCarryForwardType.Both;
        }

        if (hasCredit)
        {
            return FlexibleBenefitCarryForwardType.Credit;
        }

        if (hasExpense)
        {
            return FlexibleBenefitCarryForwardType.Expense;
        }

        if (normalized.Contains("nocarry", StringComparison.Ordinal) || normalized == "none")
        {
            return FlexibleBenefitCarryForwardType.None;
        }

        return FlexibleBenefitCarryForwardType.Unknown;
    }

    public (FlexibleBenefitAllocationType Type, decimal? HcsaPercentage, decimal? LsaPercentage) NormalizeRatio(
        string? value)
    {
        var normalized = NormalizeSearchText(value);
        if (normalized.Contains("100hcsa", StringComparison.Ordinal) ||
            normalized.Contains("100hsa", StringComparison.Ordinal))
        {
            return (FlexibleBenefitAllocationType.AllHcsa, 100m, 0m);
        }

        if (normalized.Contains("100lsa", StringComparison.Ordinal))
        {
            return (FlexibleBenefitAllocationType.AllLsa, 0m, 100m);
        }

        if (normalized.Contains("5050", StringComparison.Ordinal) ||
            normalized.Contains("50hcsa50lsa", StringComparison.Ordinal) ||
            normalized.Contains("50hsalsa", StringComparison.Ordinal))
        {
            return (FlexibleBenefitAllocationType.Split, 50m, 50m);
        }

        if (normalized.Contains("employeeelection", StringComparison.Ordinal) ||
            normalized.Contains("employeeallocation", StringComparison.Ordinal))
        {
            return (FlexibleBenefitAllocationType.EmployeeElection, null, null);
        }

        return (FlexibleBenefitAllocationType.Unknown, null, null);
    }

    public DateOnly? ParseDate(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        var cleaned = value.Trim();
        var formats = new[]
        {
            "yyyy/MM/dd", "yyyy-MM-dd", "yyyy MM dd", "MM/dd/yyyy", "MMMM d yyyy", "MMMM d, yyyy", "d MMMM yyyy"
        };
        foreach (var format in formats)
        {
            if (DateOnly.TryParseExact(cleaned, format, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out var date))
            {
                return date;
            }
        }

        var match = Regex.Match(cleaned, @"(?<year>20\d{2})\s*[/\-]\s*(?<month>\d{1,2})\s*[/\-]\s*(?<day>\d{1,2})");
        if (match.Success &&
            int.TryParse(match.Groups["year"].Value, out var year) &&
            int.TryParse(match.Groups["month"].Value, out var month) &&
            int.TryParse(match.Groups["day"].Value, out var day))
        {
            try
            {
                return new DateOnly(year, month, day);
            }
            catch (ArgumentOutOfRangeException)
            {
                return null;
            }
        }

        return DateOnly.TryParse(cleaned, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out var parsed)
            ? parsed
            : null;
    }

    public decimal? ParseMoney(string? value)
    {
        if (string.IsNullOrWhiteSpace(value) || value.Contains('%', StringComparison.Ordinal))
        {
            return null;
        }

        var match = Regex.Match(value, @"(?<!\d)(?:CAD\s*)?\$?\s*(?<amount>-?\d[\d,]*(?:\.\d{1,2})?)(?!\s*%)", RegexOptions.IgnoreCase);
        if (!match.Success)
        {
            return null;
        }

        return decimal.TryParse(
            match.Groups["amount"].Value.Replace(",", "", StringComparison.Ordinal),
            NumberStyles.Number | NumberStyles.AllowLeadingSign,
            CultureInfo.InvariantCulture,
            out var amount)
            ? amount
            : null;
    }

    public decimal? ParsePercentage(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        var match = Regex.Match(value, @"(?<percent>\d{1,3}(?:\.\d+)?)\s*%");
        return match.Success &&
               decimal.TryParse(match.Groups["percent"].Value, NumberStyles.Number, CultureInfo.InvariantCulture, out var percent) &&
               percent is >= 0m and <= 100m
            ? percent
            : null;
    }

    public int? ParseDays(string? value)
    {
        if (!string.IsNullOrWhiteSpace(value))
        {
            var dayMatch = Regex.Match(value, @"(?<!\d)(?<value>\d{1,4})\s*(?:-|\s)?days?\b", RegexOptions.IgnoreCase);
            if (dayMatch.Success && int.TryParse(dayMatch.Groups["value"].Value, out var days) && days >= 0)
            {
                return days;
            }
        }

        return ParseNonNegativeInteger(value);
    }

    public int? ParseDependentAge(string? value)
    {
        var age = ParseNonNegativeInteger(value);
        return age is >= 0 and <= 100 ? age : null;
    }

    public bool IsValidEmail(string? value)
    {
        return !string.IsNullOrWhiteSpace(value) &&
               Regex.IsMatch(value.Trim(), @"^[^\s@]+@[^\s@]+\.[^\s@]+$", RegexOptions.CultureInvariant);
    }

    public string NormalizeLegalName(string? value)
    {
        return Regex.Replace((value ?? "").Trim(), @"\s+", " ");
    }

    public static string NormalizeSearchText(string? value)
    {
        return new string((value ?? "")
            .Normalize()
            .Where(char.IsLetterOrDigit)
            .Select(char.ToLowerInvariant)
            .ToArray());
    }

    private static int? ParseNonNegativeInteger(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        var match = Regex.Match(value, @"(?<!\d)(?<value>\d{1,4})(?!\d)");
        return match.Success && int.TryParse(match.Groups["value"].Value, out var result) && result >= 0
            ? result
            : null;
    }
}

public sealed class FlexibleBenefitLayoutReader
{
    public FlexibleBenefitLineMatch? FindLine(FlexibleBenefitOcrDocument document, params string[] aliases)
    {
        return document.Pages
            .SelectMany(page => FindLines(page, aliases))
            .OrderByDescending(match => match.MatchScore)
            .FirstOrDefault();
    }

    public IReadOnlyList<FlexibleBenefitLineMatch> FindLines(FlexibleBenefitOcrPage page, params string[] aliases)
    {
        var normalizedAliases = aliases
            .Where(alias => !string.IsNullOrWhiteSpace(alias))
            .Select(alias => (Raw: alias, Normalized: FlexibleBenefitNormalizationService.NormalizeSearchText(alias)))
            .Where(alias => alias.Normalized.Length >= 3)
            .ToList();
        var matches = new List<FlexibleBenefitLineMatch>();
        foreach (var line in page.Lines)
        {
            foreach (var alias in normalizedAliases)
            {
                var score = MatchScore(line.Text, alias.Raw);
                if (score < 0.60m)
                {
                    continue;
                }

                matches.Add(new FlexibleBenefitLineMatch(page, line, alias.Raw, score));
            }
        }

        return matches;
    }

    public FlexibleBenefitSelectionMatch? FindSelection(
        FlexibleBenefitOcrPage page,
        params string[] aliases)
    {
        var choices = new List<FlexibleBenefitSelectionMatch>();
        foreach (var lineMatch in FindLines(page, aliases))
        {
            var line = lineMatch.Line;
            var verticalTolerance = Math.Max(line.Bounds.Height * 2.5, page.PixelHeight * 0.035);
            var horizontalTolerance = page.PixelWidth * 0.12;
            foreach (var mark in page.Marks)
            {
                var minimumMarkSize = page.PixelWidth * 0.016;
                var maximumMarkSize = page.PixelWidth * 0.035;
                if (mark.Shape != "Selection mark" ||
                    mark.Bounds.Width < minimumMarkSize || mark.Bounds.Height < minimumMarkSize ||
                    mark.Bounds.Width > maximumMarkSize || mark.Bounds.Height > maximumMarkSize)
                {
                    continue;
                }

                var markCenterY = mark.Bounds.Y + mark.Bounds.Height / 2d;
                var lineCenterY = line.Bounds.Y + line.Bounds.Height / 2d;
                var verticalDistance = Math.Abs(markCenterY - lineCenterY);
                if (verticalDistance > verticalTolerance)
                {
                    continue;
                }

                var horizontalGap = line.Bounds.X - mark.Bounds.Right;
                if (horizontalGap > horizontalTolerance || horizontalGap < -mark.Bounds.Width * 1.5)
                {
                    continue;
                }

                var sizePenalty = mark.Bounds.Width < page.PixelWidth * 0.012 ? page.PixelWidth * 0.04 : 0d;
                var distance = verticalDistance + Math.Abs(Math.Min(0d, horizontalGap)) * 2d + Math.Max(0d, horizontalGap) + sizePenalty;
                choices.Add(new FlexibleBenefitSelectionMatch(lineMatch, mark, distance));
            }
        }

        return choices.OrderBy(choice => choice.Distance).FirstOrDefault();
    }

    public FlexibleBenefitOrdinalSelection? FindOrdinalSelection(
        FlexibleBenefitOcrPage page,
        string anchorText,
        IReadOnlyList<string> optionValues)
    {
        var anchor = FindLines(page, anchorText).OrderByDescending(match => match.MatchScore).FirstOrDefault();
        if (anchor is null || optionValues.Count == 0)
        {
            return null;
        }

        var minSize = page.PixelWidth * 0.014;
        var maxSize = page.PixelWidth * 0.030;
        var verticalTolerance = page.PixelHeight * 0.04;
        var anchorCenterY = anchor.Line.Bounds.Y + anchor.Line.Bounds.Height / 2d;
        var candidates = page.Marks
            .Where(mark => mark.Bounds.Width >= minSize && mark.Bounds.Height >= minSize)
            .Where(mark => mark.Bounds.Width <= maxSize && mark.Bounds.Height <= maxSize)
            .Where(mark => mark.Bounds.X > anchor.Line.Bounds.Right)
            .Where(mark => Math.Abs(mark.Bounds.Y + mark.Bounds.Height / 2d - anchorCenterY) <= verticalTolerance)
            .OrderBy(mark => mark.Bounds.X)
            .ToList();

        if (candidates.Count < optionValues.Count)
        {
            return null;
        }

        var bestWindow = Enumerable.Range(0, candidates.Count - optionValues.Count + 1)
            .Select(index => candidates.Skip(index).Take(optionValues.Count).ToList())
            .OrderBy(window => WindowSpread(window))
            .First();
        var selectedIndexes = bestWindow
            .Select((mark, index) => (mark, index))
            .Where(item => item.mark.IsSelected == true)
            .Select(item => item.index)
            .ToList();
        if (selectedIndexes.Count != 1)
        {
            return new FlexibleBenefitOrdinalSelection(anchor, null, null, bestWindow, 0.55m);
        }

        var selectedIndex = selectedIndexes[0];
        return new FlexibleBenefitOrdinalSelection(
            anchor,
            optionValues[selectedIndex],
            bestWindow[selectedIndex],
            bestWindow,
            bestWindow.All(mark => mark.IsSelected.HasValue) ? 0.94m : 0.72m);
    }

    public FlexibleBenefitLineMatch? FindNearestValueLine(
        FlexibleBenefitOcrPage page,
        FlexibleBenefitLineMatch label,
        Func<string, bool> valuePredicate,
        bool toRight = true)
    {
        var labelCenterY = label.Line.Bounds.Y + label.Line.Bounds.Height / 2d;
        return page.Lines
            .Where(line => valuePredicate(line.Text))
            .Where(line => !toRight || line.Bounds.X >= label.Line.Bounds.X)
            .Select(line => new
            {
                Line = line,
                Distance = Math.Abs(line.Bounds.Y + line.Bounds.Height / 2d - labelCenterY) +
                           (toRight ? Math.Max(0d, label.Line.Bounds.Right - line.Bounds.X) : 0d)
            })
            .OrderBy(item => item.Distance)
            .Select(item => new FlexibleBenefitLineMatch(page, item.Line, label.Alias, 0.80m))
            .FirstOrDefault();
    }

    public FlexibleBenefitPlanSourceField BuildSourceField(
        string entityType,
        string entityKey,
        string fieldName,
        string? rawValue,
        string? normalizedValue,
        FlexibleBenefitLineMatch? line,
        FlexibleBenefitDetectedMark? mark,
        string extractionMethod,
        decimal confidence,
        string reviewStatus = FlexibleBenefitReviewStatuses.Pending)
    {
        var bounds = new List<FlexibleBenefitBounds>();
        if (line is not null)
        {
            bounds.Add(line.Line.Bounds);
        }

        if (mark is not null)
        {
            bounds.Add(mark.Bounds);
        }

        return new FlexibleBenefitPlanSourceField
        {
            EntityType = entityType,
            EntityKey = entityKey,
            FieldName = fieldName,
            RawValue = rawValue?.Trim() ?? "",
            NormalizedValue = normalizedValue?.Trim() ?? "",
            SourcePageNumber = line?.Page.PageNumber,
            BoundingBoxJson = bounds.Count == 0 ? "" : JsonSerializer.Serialize(Union(bounds)),
            ExtractionMethod = extractionMethod,
            ExtractionConfidence = confidence,
            ReviewStatus = reviewStatus
        };
    }

    private static decimal MatchScore(string line, string alias)
    {
        var normalizedLine = FlexibleBenefitNormalizationService.NormalizeSearchText(line);
        var normalizedAlias = FlexibleBenefitNormalizationService.NormalizeSearchText(alias);
        if (normalizedLine.Length == 0 || normalizedAlias.Length == 0)
        {
            return 0m;
        }

        if (normalizedLine.Contains(normalizedAlias, StringComparison.Ordinal))
        {
            return 1m;
        }

        if (normalizedAlias.Contains(normalizedLine, StringComparison.Ordinal) && normalizedLine.Length >= 8)
        {
            return 0.90m;
        }

        var lineTokens = Tokenize(line);
        var aliasTokens = Tokenize(alias);
        if (aliasTokens.Count == 0)
        {
            return 0m;
        }

        var common = aliasTokens.Count(token => lineTokens.Any(lineToken =>
            lineToken.Equals(token, StringComparison.Ordinal) ||
            Math.Min(lineToken.Length, token.Length) >= 4 &&
            (lineToken.Contains(token, StringComparison.Ordinal) || token.Contains(lineToken, StringComparison.Ordinal))));
        return (decimal)common / aliasTokens.Count;
    }

    private static List<string> Tokenize(string value)
    {
        return Regex.Matches(value.ToLowerInvariant(), @"[a-z0-9]+")
            .Select(match => match.Value)
            .Where(token => token.Length >= 3 || char.IsDigit(token[0]))
            .ToList();
    }

    private static double WindowSpread(IReadOnlyList<FlexibleBenefitDetectedMark> marks)
    {
        if (marks.Count < 2)
        {
            return 0d;
        }

        var gaps = marks.Zip(marks.Skip(1), (left, right) => right.Bounds.X - left.Bounds.X).ToList();
        return gaps.Max() - gaps.Min();
    }

    private static FlexibleBenefitBounds Union(IReadOnlyList<FlexibleBenefitBounds> values)
    {
        var left = values.Min(value => value.X);
        var top = values.Min(value => value.Y);
        var right = values.Max(value => value.Right);
        var bottom = values.Max(value => value.Bottom);
        return new FlexibleBenefitBounds(left, top, right - left, bottom - top);
    }
}

public sealed record FlexibleBenefitLineMatch(
    FlexibleBenefitOcrPage Page,
    FlexibleBenefitOcrLine Line,
    string Alias,
    decimal MatchScore);

public sealed record FlexibleBenefitSelectionMatch(
    FlexibleBenefitLineMatch LineMatch,
    FlexibleBenefitDetectedMark Mark,
    double Distance);

public sealed record FlexibleBenefitOrdinalSelection(
    FlexibleBenefitLineMatch Anchor,
    string? SelectedValue,
    FlexibleBenefitDetectedMark? SelectedMark,
    IReadOnlyList<FlexibleBenefitDetectedMark> OptionMarks,
    decimal Confidence);

public sealed class FlexibleBenefitSummaryService
{
    public string BuildHcsaSummary(FlexibleBenefitPlan plan, IReadOnlyList<FlexibleBenefitClass> classes)
    {
        if (plan.HcsaEnabled == false)
        {
            return "Not included";
        }

        if (plan.HcsaEnabled != true)
        {
            return "Not confirmed";
        }

        var amounts = classes.Select(item => item.HcsaAmount).Where(value => value.HasValue).Select(value => value!.Value).ToList();
        if (amounts.Count > 0)
        {
            return FormatRange("HCSA", amounts);
        }

        if (plan.DefaultAllocationType == FlexibleBenefitAllocationType.EmployeeElection)
        {
            return "Employee allocation";
        }

        if (plan.DefaultHcsaPercentage.HasValue)
        {
            var annual = FormatAllocatedClassRange(classes, plan.DefaultHcsaPercentage.Value);
            return annual == "Not confirmed" ? annual : $"{annual} default allocation";
        }

        return plan.LsaEnabled == true ? "Not confirmed" : FormatAnnualClassRange(classes);
    }

    public string BuildLsaSummary(FlexibleBenefitPlan plan, IReadOnlyList<FlexibleBenefitClass> classes)
    {
        if (plan.LsaEnabled == false)
        {
            return "Not included";
        }

        if (plan.LsaEnabled != true)
        {
            return "Not confirmed";
        }

        var amounts = classes.Select(item => item.LsaAmount).Where(value => value.HasValue).Select(value => value!.Value).ToList();
        if (amounts.Count > 0)
        {
            return FormatRange("LSA", amounts);
        }

        if (plan.DefaultAllocationType == FlexibleBenefitAllocationType.EmployeeElection)
        {
            return "Employee allocation";
        }

        if (plan.DefaultLsaPercentage.HasValue)
        {
            var annual = FormatAllocatedClassRange(classes, plan.DefaultLsaPercentage.Value);
            return annual == "Not confirmed" ? annual : $"{annual} default allocation";
        }

        return plan.HcsaEnabled == true ? "Not confirmed" : FormatAnnualClassRange(classes);
    }

    public string BuildCombinedSummary(FlexibleBenefitPlan plan, IReadOnlyList<FlexibleBenefitClass> classes)
    {
        if (plan.HcsaEnabled == true && plan.LsaEnabled == true &&
            plan.DefaultAllocationType == FlexibleBenefitAllocationType.EmployeeElection)
        {
            return "HCSA/LSA: Employee allocation";
        }

        if (plan.HcsaEnabled == true && plan.LsaEnabled == true)
        {
            return $"HCSA: {BuildHcsaSummary(plan, classes)}; LSA: {BuildLsaSummary(plan, classes)}";
        }

        if (plan.HcsaEnabled == true)
        {
            return $"HCSA: {BuildHcsaSummary(plan, classes)}";
        }

        if (plan.LsaEnabled == true)
        {
            return $"LSA: {BuildLsaSummary(plan, classes)}";
        }

        return plan.HcsaEnabled == false && plan.LsaEnabled == false ? "Not included" : "Not confirmed";
    }

    private static string FormatAnnualClassRange(IReadOnlyList<FlexibleBenefitClass> classes)
    {
        var amounts = classes
            .SelectMany(item => new[] { item.AnnualSingleAmount, item.AnnualCoupleAmount, item.AnnualFamilyAmount })
            .Where(value => value.HasValue)
            .Select(value => value!.Value)
            .Distinct()
            .OrderBy(value => value)
            .ToList();
        if (amounts.Count == 0)
        {
            return "Not confirmed";
        }

        return amounts.Count == 1
            ? $"{amounts[0]:C0} per year"
            : $"{amounts[0]:C0}-{amounts[^1]:C0} by class";
    }

    private static string FormatRange(string account, IReadOnlyList<decimal> amounts)
    {
        var ordered = amounts.Distinct().OrderBy(value => value).ToList();
        return ordered.Count == 1
            ? $"{ordered[0]:C0} per year"
            : $"{ordered[0]:C0}-{ordered[^1]:C0} by class";
    }

    private static string FormatAllocatedClassRange(IReadOnlyList<FlexibleBenefitClass> classes, decimal percentage)
    {
        var amounts = classes
            .SelectMany(item => new[] { item.AnnualSingleAmount, item.AnnualCoupleAmount, item.AnnualFamilyAmount })
            .Where(value => value.HasValue)
            .Select(value => decimal.Round(value!.Value * percentage / 100m, 2, MidpointRounding.AwayFromZero))
            .Distinct()
            .OrderBy(value => value)
            .ToList();
        if (amounts.Count == 0)
        {
            return "Not confirmed";
        }

        return amounts.Count == 1
            ? $"{amounts[0]:C0} per year"
            : $"{amounts[0]:C0}-{amounts[^1]:C0} by class";
    }
}

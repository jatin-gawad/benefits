using System.Globalization;
using System.Text.RegularExpressions;
using System.Windows.Media;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public static partial class ClassComparisonEngine
{
    // The workspace is horizontally scrollable; do not impose the former five-class UI limit.
    public const int MaximumSelectedClasses = 12;

    private static readonly Brush[] ClassPalette =
    [
        new SolidColorBrush(Color.FromRgb(71, 85, 105)),
        new SolidColorBrush(Color.FromRgb(22, 163, 74)),
        new SolidColorBrush(Color.FromRgb(37, 99, 235)),
        new SolidColorBrush(Color.FromRgb(15, 139, 141)),
        new SolidColorBrush(Color.FromRgb(245, 158, 11))
    ];

    private static readonly IReadOnlyDictionary<string, decimal> ScoreWeights = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase)
    {
        ["Plan richness"] = 0.28m,
        ["Employee cost"] = 0.16m,
        ["Employer cost"] = 0.16m,
        ["Eligibility"] = 0.10m,
        ["Coverage breadth"] = 0.14m,
        ["Competitiveness"] = 0.10m,
        ["Administration complexity"] = 0.03m,
        ["Employee impact"] = 0.03m
    };

    private static readonly string[] SectionOrder =
    [
        "Eligibility",
        "Life Insurance",
        "Disability",
        "Extended Health Care",
        "Dental Care",
        "Cost Sharing",
        "Other"
    ];

    public static ClassComparisonWorkspaceViewModel Build(ClassComparisonBuildRequest request)
    {
        ArgumentNullException.ThrowIfNull(request);

        var classNames = ResolveClassNames(request.BenefitRows);
        var sourceFeatures = BuildFeatureValues(request.BenefitRows, classNames);
        var sourceRates = BuildRateRows(request.RateRows, classNames);
        var sourceCosts = BuildCostRows(request.CostRows, sourceRates, classNames);
        var preferredSelected = request.Preferences.SelectedClassNames
            .Where(selected => classNames.Any(name => name.Equals(selected, StringComparison.OrdinalIgnoreCase)))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Take(MaximumSelectedClasses)
            .ToList();
        if (preferredSelected.Count == 0)
        {
            preferredSelected = classNames.Take(MaximumSelectedClasses).ToList();
        }

        var baseline = classNames.FirstOrDefault(name =>
                           name.Equals(request.Preferences.BaselineClassName, StringComparison.OrdinalIgnoreCase) &&
                           preferredSelected.Contains(name, StringComparer.OrdinalIgnoreCase))
                       ?? preferredSelected.FirstOrDefault()
                       ?? classNames.FirstOrDefault()
                       ?? "";

        var model = new ClassComparisonWorkspaceViewModel
        {
            ClientId = request.ClientId,
            ClientName = request.ClientName,
            EffectiveDateText = FirstNonBlank(request.EffectiveDateText, "Not available"),
            TotalEligibleEmployees = request.TotalEligibleEmployees,
            Consultant = FirstNonBlank(request.Consultant, "Not available"),
            PlanStatus = FirstNonBlank(request.PlanStatus, "Not available"),
            SourceFeatures = sourceFeatures,
            SourceCosts = sourceCosts,
            SourceRates = sourceRates,
            BaselineClassName = baseline,
            ActiveTab = FirstNonBlank(request.Preferences.ActiveTab, "Class Comparison"),
            KeyBenefitsOnly = request.Preferences.KeyBenefitsOnly,
            HighlightDifferences = request.Preferences.HighlightDifferences,
            RecommendedClassName = request.Preferences.RecommendedClassName
        };

        foreach (var section in request.Preferences.CollapsedSections.Where(section => !string.IsNullOrWhiteSpace(section)))
        {
            model.CollapsedSections.Add(section.Trim());
        }

        for (var index = 0; index < classNames.Count; index++)
        {
            var className = classNames[index];
            var cost = sourceCosts.FirstOrDefault(item => item.ClassName.Equals(className, StringComparison.OrdinalIgnoreCase));
            var (displayName, description, classCode) = ClassDisplayParts(className);
            model.Classes.Add(new ClassComparisonClassCard
            {
                ClassName = className,
                DisplayName = displayName,
                Description = description,
                ClassCode = classCode,
                EmployeeCount = cost?.EmployeeCount,
                AnnualEmployerCost = cost?.AnnualEmployerCost,
                AnnualEmployeeCost = cost?.AnnualEmployeeCost,
                AnnualTotalCost = cost?.AnnualTotalCost,
                TotalPlanCostPercent = cost?.TotalClientCostPercent,
                ProjectedChangePercent = cost?.ProjectedChangePercent,
                AccentBrush = ClassPalette[index % ClassPalette.Length],
                IsSelected = preferredSelected.Contains(className, StringComparer.OrdinalIgnoreCase),
                IsBaseline = className.Equals(baseline, StringComparison.OrdinalIgnoreCase)
            });
        }

        model.ComparisonClassName = preferredSelected
            .FirstOrDefault(name => !name.Equals(baseline, StringComparison.OrdinalIgnoreCase))
            ?? classNames.FirstOrDefault(name => !name.Equals(baseline, StringComparison.OrdinalIgnoreCase))
            ?? "";

        foreach (var note in request.Notes)
        {
            model.Notes.Add(note);
        }

        foreach (var historyRow in request.History)
        {
            model.History.Add(historyRow);
        }

        Refresh(model);
        return model;
    }

    public static void Refresh(ClassComparisonWorkspaceViewModel model)
    {
        ArgumentNullException.ThrowIfNull(model);

        var selected = model.Classes.Where(item => item.IsSelected).Take(MaximumSelectedClasses).ToList();
        if (selected.Count == 0 && model.Classes.Count > 0)
        {
            model.Classes[0].IsSelected = true;
            selected.Add(model.Classes[0]);
        }

        if (!selected.Any(item => item.ClassName.Equals(model.BaselineClassName, StringComparison.OrdinalIgnoreCase)))
        {
            model.BaselineClassName = selected.FirstOrDefault()?.ClassName ?? "";
        }

        foreach (var classCard in model.Classes)
        {
            classCard.IsBaseline = classCard.ClassName.Equals(model.BaselineClassName, StringComparison.OrdinalIgnoreCase);
        }

        Replace(model.SelectedClasses, selected);
        BuildComparisonRows(model, selected);
        BuildVisibleCosts(model, selected);
        BuildVisibleRates(model, selected);
        BuildScores(model, selected);
        BuildDifferenceSummaries(model, selected);
        RefreshPairComparison(model);
        model.StatusText = $"Comparing {model.SelectedClassCountText} against {model.BaselineClass?.DisplayName ?? "the selected baseline"}.";
        model.NotifyCollectionsChanged();
    }

    public static bool TrySetClassSelection(
        ClassComparisonWorkspaceViewModel model,
        string className,
        bool isSelected,
        out string validationMessage)
    {
        validationMessage = "";
        var target = model.Classes.FirstOrDefault(item => item.ClassName.Equals(className, StringComparison.OrdinalIgnoreCase));
        if (target is null)
        {
            validationMessage = "That class is no longer available for this client.";
            return false;
        }

        var currentCount = model.Classes.Count(item => item.IsSelected);
        if (isSelected && !target.IsSelected && currentCount >= MaximumSelectedClasses)
        {
            validationMessage = $"Select up to {MaximumSelectedClasses} classes at a time.";
            return false;
        }

        if (!isSelected && target.IsSelected && currentCount <= 1)
        {
            validationMessage = "At least one class must remain selected.";
            return false;
        }

        target.IsSelected = isSelected;
        if (!isSelected && target.IsBaseline)
        {
            model.BaselineClassName = model.Classes.First(item => item.IsSelected).ClassName;
        }

        Refresh(model);
        return true;
    }

    public static bool TrySetBaseline(ClassComparisonWorkspaceViewModel model, string className)
    {
        if (!model.Classes.Any(item => item.IsSelected && item.ClassName.Equals(className, StringComparison.OrdinalIgnoreCase)))
        {
            return false;
        }

        model.BaselineClassName = className;
        Refresh(model);
        return true;
    }

    public static bool TrySetRecommendedClass(ClassComparisonWorkspaceViewModel model, string className)
    {
        if (!model.Classes.Any(item => item.IsSelected && item.ClassName.Equals(className, StringComparison.OrdinalIgnoreCase)))
        {
            return false;
        }

        model.RecommendedClassName = className;
        Refresh(model);
        return true;
    }

    public static bool TrySetPairClasses(
        ClassComparisonWorkspaceViewModel model,
        string baselineClassName,
        string comparisonClassName,
        out string validationMessage)
    {
        validationMessage = "";
        var baseline = model.Classes.FirstOrDefault(item => item.ClassName.Equals(baselineClassName, StringComparison.OrdinalIgnoreCase));
        var comparison = model.Classes.FirstOrDefault(item => item.ClassName.Equals(comparisonClassName, StringComparison.OrdinalIgnoreCase));
        if (baseline is null || comparison is null)
        {
            validationMessage = "Both selected classes must belong to this client.";
            return false;
        }

        if (baseline.ClassName.Equals(comparison.ClassName, StringComparison.OrdinalIgnoreCase))
        {
            validationMessage = "Choose two different plan classes to compare.";
            return false;
        }

        foreach (var classCard in model.Classes)
        {
            classCard.IsSelected = classCard == baseline || classCard == comparison;
        }
        model.BaselineClassName = baseline.ClassName;
        model.ComparisonClassName = comparison.ClassName;
        Refresh(model);
        return true;
    }

    public static bool SwapPairClasses(ClassComparisonWorkspaceViewModel model)
    {
        if (!model.HasEnoughClasses || string.IsNullOrWhiteSpace(model.ComparisonClassName))
        {
            return false;
        }

        var baseline = model.BaselineClassName;
        return TrySetPairClasses(model, model.ComparisonClassName, baseline, out _);
    }

    public static void SetShowDifferencesOnly(ClassComparisonWorkspaceViewModel model, bool showDifferencesOnly)
    {
        model.ShowDifferencesOnly = showDifferencesOnly;
        model.ApplyPairFilters();
    }

    public static void SetSearchText(ClassComparisonWorkspaceViewModel model, string searchText)
    {
        model.SearchText = searchText ?? "";
    }

    public static void SetKeyBenefitsOnly(ClassComparisonWorkspaceViewModel model, bool keyBenefitsOnly)
    {
        model.KeyBenefitsOnly = keyBenefitsOnly;
        Refresh(model);
    }

    public static void SetHighlightDifferences(ClassComparisonWorkspaceViewModel model, bool highlightDifferences)
    {
        model.HighlightDifferences = highlightDifferences;
        Refresh(model);
    }

    public static void ToggleSection(ClassComparisonWorkspaceViewModel model, string section)
    {
        if (!model.CollapsedSections.Add(section))
        {
            model.CollapsedSections.Remove(section);
        }

        Refresh(model);
    }

    public static void SetAllSectionsExpanded(ClassComparisonWorkspaceViewModel model, bool expanded)
    {
        ArgumentNullException.ThrowIfNull(model);
        if (expanded)
        {
            model.CollapsedSections.Clear();
        }
        else
        {
            foreach (var section in model.SourceFeatures
                         .Select(value => value.Section)
                         .Where(section => !string.IsNullOrWhiteSpace(section))
                         .Distinct(StringComparer.OrdinalIgnoreCase))
            {
                model.CollapsedSections.Add(section);
            }
        }

        Refresh(model);
    }

    public static void ResetFilters(ClassComparisonWorkspaceViewModel model)
    {
        model.KeyBenefitsOnly = false;
        model.HighlightDifferences = true;
        model.CollapsedSections.Clear();
        foreach (var item in model.Classes.Select((card, index) => new { card, index }))
        {
            item.card.IsSelected = item.index < MaximumSelectedClasses;
        }

        model.BaselineClassName = model.Classes.FirstOrDefault(item => item.IsSelected)?.ClassName ?? "";
        model.RecommendedClassName = "";
        Refresh(model);
    }

    public static ClassComparisonRule ResolveRule(string section, string benefit, string displayValue)
    {
        var text = $"{section} {benefit}".ToLowerInvariant();
        var normalizedDisplay = displayValue.ToLowerInvariant();
        var isKey = ContainsAny(text,
            "employee basic life", "life benefit", "ad&d", "dependent life", "ltd", "long term disability",
            "prescription", "drug", "paramedical", "mental health", "vision", "hospital", "travel",
            "deductible", "annual maximum", "major restorative", "orthodont", "waiting period", "eligibility");

        if (ContainsAny(text, "taxability", "taxable", "cost sharing", "plan type", "provider", "employee definition", "actively", "fee guide", "mandatory generic", "substitution"))
        {
            return new ClassComparisonRule(ClassComparisonDirection.Neutral, ClassComparisonValueKind.Text, isKey, 0m, "Informational provision; no automatic better-or-worse result is assigned.");
        }

        if (ContainsAny(text, "waiting period", "elimination period", "recall", "complete exam", "supply limit", "frequency"))
        {
            return new ClassComparisonRule(ClassComparisonDirection.LowerIsBetter, ClassComparisonValueKind.Duration, isKey, 1m, "A shorter interval generally provides earlier or more frequent access.");
        }

        if (ContainsAny(text, "deductible", "employee contribution", "employee cost"))
        {
            return new ClassComparisonRule(ClassComparisonDirection.LowerIsBetter, ClassComparisonValueKind.Money, isKey, 1m, "A lower employee-paid amount is treated as better value.");
        }

        if (ContainsAny(text, "termination age", "own occupation", "benefit period", "reduction age"))
        {
            return new ClassComparisonRule(ClassComparisonDirection.HigherIsBetter, ClassComparisonValueKind.Duration, isKey, 0.8m, "A longer covered period is treated as richer coverage.");
        }

        if (ContainsAny(text, "coinsurance", "reimbursement", "coverage level", "basic services", "major services", "diagnostic", "orthodontics"))
        {
            return new ClassComparisonRule(ClassComparisonDirection.HigherIsBetter, ClassComparisonValueKind.Percent, isKey, 1m, "A higher reimbursement percentage is treated as richer coverage.");
        }

        if (ContainsAny(text, "maximum", "benefit amount", "benefit formula", "annual earnings", "paramedical", "mental health", "vision", "health spending", "flex credit"))
        {
            return new ClassComparisonRule(ClassComparisonDirection.HigherIsBetter, ClassComparisonValueKind.Money, isKey, 1m, "A higher covered maximum is treated as richer coverage when values are comparable.");
        }

        if (ContainsAny(normalizedDisplay, "included", "covered", "not included", "no coverage", "not covered"))
        {
            return new ClassComparisonRule(ClassComparisonDirection.IncludedIsBetter, ClassComparisonValueKind.Boolean, isKey, 0.8m, "Included coverage is treated as better than excluded coverage.");
        }

        return new ClassComparisonRule(ClassComparisonDirection.Neutral, ClassComparisonValueKind.Auto, isKey, 0m, "Different wording is shown for review without assigning a value direction.");
    }

    private static List<string> ResolveClassNames(IReadOnlyList<ClassBenefitSourceRow> rows)
    {
        var names = rows
            .Select(row => FirstNonBlank(row.ClassName, "All Classes"))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(ClassSortKey, StringComparer.OrdinalIgnoreCase)
            .ToList();
        var concrete = names.Where(name => !IsGenericClass(name)).ToList();
        return concrete.Count > 0 ? concrete : names;
    }

    private static IReadOnlyList<ClassComparisonFeatureValue> BuildFeatureValues(
        IReadOnlyList<ClassBenefitSourceRow> rows,
        IReadOnlyList<string> classNames)
    {
        var values = new List<ClassComparisonFeatureValue>();
        foreach (var row in rows)
        {
            var targetClasses = IsGenericClass(row.ClassName) && classNames.Any(name => !IsGenericClass(name))
                ? classNames
                : classNames.Where(name => ClassNamesMatch(name, row.ClassName)).ToList();
            if (targetClasses.Count == 0)
            {
                continue;
            }

            var section = ResolveSection(row);
            var (benefit, primaryValue) = ResolveBenefitLabelAndPrimaryValue(row);
            if (!TryBuildClientFacingValue(row, section, benefit, primaryValue, out benefit, out var displayValue))
            {
                continue;
            }
            var rule = ResolveRule(section, benefit, displayValue);
            var normalized = ClassComparisonValueNormalizer.Normalize(displayValue, rule.ValueKind);
            var key = $"{NormalizeKey(section)}::{NormalizeKey(benefit)}";

            foreach (var className in targetClasses)
            {
                values.Add(new ClassComparisonFeatureValue(
                    section,
                    benefit,
                    key,
                    className,
                    NormalizeDisplayValue(displayValue),
                    rule,
                    normalized,
                    row.SourceDocument,
                    row.SourcePage,
                    FirstNonBlank(row.SourceText, row.Notes),
                    row.ConfidenceScore));
            }
        }

        return values
            .GroupBy(value => new { value.RowKey, Class = value.ClassName.ToUpperInvariant() })
            .Select(group => MergeFeatureValues(group.ToList()))
            .OrderBy(value => SectionIndex(value.Section))
            .ThenBy(value => BenefitOrder(value.Benefit))
            .ThenBy(value => value.Benefit, StringComparer.OrdinalIgnoreCase)
            .ThenBy(value => value.ClassName, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static ClassComparisonFeatureValue MergeFeatureValues(IReadOnlyList<ClassComparisonFeatureValue> values)
    {
        var first = values[0];
        var displays = values.Select(value => value.DisplayValue)
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        var display = displays.Count == 0 ? "—" : string.Join(Environment.NewLine, displays);
        return first with
        {
            DisplayValue = display,
            NormalizedValue = ClassComparisonValueNormalizer.Normalize(display, first.Rule.ValueKind),
            SourceDocument = FirstNonBlank(values.Select(value => value.SourceDocument)),
            SourcePage = string.Join(", ", values.Select(value => value.SourcePage).Where(value => !string.IsNullOrWhiteSpace(value)).Distinct(StringComparer.OrdinalIgnoreCase)),
            SourceText = string.Join(" | ", values.Select(value => value.SourceText).Where(value => !string.IsNullOrWhiteSpace(value)).Distinct(StringComparer.OrdinalIgnoreCase))
            ,ConfidenceScore = values.Where(value => value.ConfidenceScore.HasValue).Select(value => value.ConfidenceScore).DefaultIfEmpty().Min()
        };
    }

    public static ClassComparisonResult BuildPairResult(
        ClassComparisonWorkspaceViewModel model,
        string baselineClassName,
        string comparisonClassName)
    {
        ArgumentNullException.ThrowIfNull(model);
        if (string.IsNullOrWhiteSpace(baselineClassName) ||
            string.IsNullOrWhiteSpace(comparisonClassName) ||
            baselineClassName.Equals(comparisonClassName, StringComparison.OrdinalIgnoreCase))
        {
            return new ClassComparisonResult
            {
                BaselineClassId = baselineClassName,
                ComparisonClassId = comparisonClassName,
                ValidationStatus = "Two different classes are required."
            };
        }

        var selectedNames = new HashSet<string>([baselineClassName, comparisonClassName], StringComparer.OrdinalIgnoreCase);
        var rows = new List<ClassComparisonRow>();
        foreach (var group in model.SourceFeatures
                     .Where(value => selectedNames.Contains(value.ClassName))
                     .GroupBy(value => value.RowKey, StringComparer.OrdinalIgnoreCase)
                     .OrderBy(group => SectionIndex(group.First().Section))
                     .ThenBy(group => BenefitOrder(group.First().Benefit))
                     .ThenBy(group => group.First().Benefit, StringComparer.OrdinalIgnoreCase))
        {
            var first = group.First();
            var baseline = group.FirstOrDefault(value => value.ClassName.Equals(baselineClassName, StringComparison.OrdinalIgnoreCase));
            var comparison = group.FirstOrDefault(value => value.ClassName.Equals(comparisonClassName, StringComparison.OrdinalIgnoreCase));
            rows.Add(BuildPairRow(first, baseline, comparison, model.HighlightDifferences));
        }

        return new ClassComparisonResult
        {
            BaselineClassId = baselineClassName,
            ComparisonClassId = comparisonClassName,
            Rows = rows,
            ValidationStatus = rows.Count == 0 ? "No normalized comparable features are available." : "Ready"
        };
    }

    private static void RefreshPairComparison(ClassComparisonWorkspaceViewModel model)
    {
        var baseline = model.Classes.FirstOrDefault(item => item.ClassName.Equals(model.BaselineClassName, StringComparison.OrdinalIgnoreCase));
        if (baseline is null && model.Classes.Count > 0)
        {
            baseline = model.Classes[0];
            model.BaselineClassName = baseline.ClassName;
        }

        var comparison = model.Classes.FirstOrDefault(item => item.ClassName.Equals(model.ComparisonClassName, StringComparison.OrdinalIgnoreCase));
        if (comparison is null || comparison == baseline)
        {
            comparison = model.Classes.FirstOrDefault(item => item != baseline);
            model.ComparisonClassName = comparison?.ClassName ?? "";
        }

        Replace(model.ComparisonClassOptions, model.Classes.Where(item => item != baseline));
        model.ComparisonResult = BuildPairResult(model, baseline?.ClassName ?? "", comparison?.ClassName ?? "");
    }

    private static ClassComparisonRow BuildPairRow(
        ClassComparisonFeatureValue first,
        ClassComparisonFeatureValue? baseline,
        ClassComparisonFeatureValue? comparison,
        bool highlightDifferences)
    {
        var reliable = IsReliable(baseline) && IsReliable(comparison);
        var kind = ResolveDifferenceKind(baseline, comparison, first.Rule, reliable);
        var eligibility = IsEligibilityFeature(first);
        if (eligibility && kind is ClassComparisonDifferenceKind.Different or ClassComparisonDifferenceKind.Higher or ClassComparisonDifferenceKind.Lower)
        {
            kind = ClassComparisonDifferenceKind.DifferentEligibility;
        }
        var material = kind is not ClassComparisonDifferenceKind.Same and not ClassComparisonDifferenceKind.NeedsReview;
        var cost = material && IsCostFeature(first);
        var evidence = new[] { EvidenceReference(baseline), EvidenceReference(comparison) }
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        return new ClassComparisonRow
        {
            Category = first.Section,
            FeatureKey = first.RowKey,
            DisplayName = first.Benefit,
            BaselineRawValue = FirstNonBlank(baseline?.SourceText, baseline?.DisplayValue),
            ComparisonRawValue = FirstNonBlank(comparison?.SourceText, comparison?.DisplayValue),
            BaselineDisplayValue = baseline?.DisplayValue ?? "Not available",
            ComparisonDisplayValue = comparison?.DisplayValue ?? "Not available",
            DifferenceKind = kind,
            DifferenceLabel = DifferenceLabel(kind, first),
            IsMaterialDifference = material,
            IsCostDifference = cost,
            IsEligibilityDifference = material && eligibility,
            HasReliableEvidence = reliable,
            EvidenceReferences = evidence,
            HighlightDifferences = highlightDifferences
        };
    }

    private static ClassComparisonDifferenceKind ResolveDifferenceKind(
        ClassComparisonFeatureValue? baseline,
        ClassComparisonFeatureValue? comparison,
        ClassComparisonRule rule,
        bool reliable)
    {
        if (!reliable)
        {
            return ClassComparisonDifferenceKind.NeedsReview;
        }
        if (baseline is null && comparison is null)
        {
            return ClassComparisonDifferenceKind.MissingValue;
        }
        if (baseline is not null && comparison is null)
        {
            return ClassComparisonDifferenceKind.BaselineOnly;
        }
        if (baseline is null)
        {
            return ClassComparisonDifferenceKind.ComparisonOnly;
        }
        if (baseline.NormalizedValue.IsMissing || comparison!.NormalizedValue.IsMissing)
        {
            return ClassComparisonDifferenceKind.MissingValue;
        }
        if (ClassComparisonValueNormalizer.AreEquivalent(baseline.NormalizedValue, comparison.NormalizedValue))
        {
            return ClassComparisonDifferenceKind.Same;
        }
        if (baseline.NormalizedValue.NumericValue.HasValue && comparison.NormalizedValue.NumericValue.HasValue)
        {
            return comparison.NormalizedValue.NumericValue.Value > baseline.NormalizedValue.NumericValue.Value
                ? ClassComparisonDifferenceKind.Higher
                : ClassComparisonDifferenceKind.Lower;
        }
        return ClassComparisonDifferenceKind.Different;
    }

    private static bool IsReliable(ClassComparisonFeatureValue? value) =>
        value is null || !value.ConfidenceScore.HasValue || value.ConfidenceScore.Value >= 0.75m;

    private static bool IsEligibilityFeature(ClassComparisonFeatureValue feature) =>
        feature.Section.Equals("Eligibility", StringComparison.OrdinalIgnoreCase) ||
        ContainsAny($"{feature.Section} {feature.Benefit}", "eligibility", "waiting period", "hours worked", "participation", "termination age", "employee definition");

    private static bool IsCostFeature(ClassComparisonFeatureValue feature) =>
        ContainsAny($"{feature.Section} {feature.Benefit}", "premium", "rate", "employee contribution", "employer contribution", "cost", "taxable", "taxability");

    private static string DifferenceLabel(ClassComparisonDifferenceKind kind, ClassComparisonFeatureValue feature)
    {
        var name = feature.Benefit.ToLowerInvariant();
        return kind switch
        {
            ClassComparisonDifferenceKind.Same => "Same",
            ClassComparisonDifferenceKind.Higher when name.Contains("maximum") => "Higher maximum",
            ClassComparisonDifferenceKind.Lower when name.Contains("maximum") => "Lower maximum",
            ClassComparisonDifferenceKind.Higher when ContainsAny(name, "coinsurance", "coverage", "reimbursement") => "Higher coverage",
            ClassComparisonDifferenceKind.Lower when ContainsAny(name, "coinsurance", "coverage", "reimbursement") => "Lower coverage",
            ClassComparisonDifferenceKind.Higher => "Higher value",
            ClassComparisonDifferenceKind.Lower => "Lower value",
            ClassComparisonDifferenceKind.DifferentEligibility => "Different eligibility",
            ClassComparisonDifferenceKind.BaselineOnly => "Baseline only",
            ClassComparisonDifferenceKind.ComparisonOnly => "Comparison only",
            ClassComparisonDifferenceKind.MissingValue => "Missing value",
            ClassComparisonDifferenceKind.NeedsReview => "Needs review",
            _ => "Different"
        };
    }

    private static string EvidenceReference(ClassComparisonFeatureValue? value)
    {
        if (value is null) return "";
        return string.IsNullOrWhiteSpace(value.SourcePage)
            ? value.SourceDocument
            : $"{value.SourceDocument}, page {value.SourcePage}";
    }

    private static string NormalizeDisplayValue(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return "—";
        }

        return value.Trim() switch
        {
            var text when text.Equals("n/a", StringComparison.OrdinalIgnoreCase) || text.Equals("not applicable", StringComparison.OrdinalIgnoreCase) => "N/A",
            var text when text.Equals("not covered", StringComparison.OrdinalIgnoreCase) || text.Equals("not included", StringComparison.OrdinalIgnoreCase) => "Not covered",
            var text when text.Equals("not available", StringComparison.OrdinalIgnoreCase) || text.Equals("not provided", StringComparison.OrdinalIgnoreCase) => "—",
            var text => text
        };
    }

    private static IReadOnlyList<ClassRateComparisonRow> BuildRateRows(
        IReadOnlyList<ClassRateSourceRow> rates,
        IReadOnlyList<string> classNames)
    {
        var output = new List<ClassRateComparisonRow>();
        foreach (var rate in rates)
        {
            var matchingClass = classNames.FirstOrDefault(className => ClassNamesMatch(className, rate.ClassName));
            if (matchingClass is null)
            {
                continue;
            }

            output.Add(new ClassRateComparisonRow
            {
                Benefit = rate.BenefitLine,
                ClassName = matchingClass,
                ClassDisplayName = ClassDisplayParts(matchingClass).DisplayName,
                CurrentRate = rate.CurrentRate,
                ProposedRate = rate.ProposedRate,
                NegotiatedRate = rate.NegotiatedRate,
                RateBasis = rate.RateBasis,
                EffectiveDateText = rate.EffectiveDateText,
                SourceDocument = rate.SourceDocument
            });
        }

        return output
            .GroupBy(row => new
            {
                Class = row.ClassName.ToUpperInvariant(),
                Benefit = row.Benefit.ToUpperInvariant(),
                row.CurrentRate,
                row.ProposedRate,
                row.NegotiatedRate,
                Basis = row.RateBasis.ToUpperInvariant(),
                row.EffectiveDateText
            })
            .Select(group => group.First())
            .OrderBy(row => row.Benefit, StringComparer.OrdinalIgnoreCase)
            .ThenBy(row => row.ClassName, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static IReadOnlyList<ClassCostSummaryRow> BuildCostRows(
        IReadOnlyList<ClassCostSourceRow> costs,
        IReadOnlyList<ClassRateComparisonRow> rates,
        IReadOnlyList<string> classNames)
    {
        var output = new List<ClassCostSummaryRow>();
        for (var index = 0; index < classNames.Count; index++)
        {
            var className = classNames[index];
            var matches = costs.Where(cost => ClassNamesMatch(className, cost.ClassName)).ToList();
            var monthlyEmployer = SumNullable(matches.Select(item => item.MonthlyEmployerCost));
            var monthlyEmployee = SumNullable(matches.Select(item => item.MonthlyEmployeeCost));
            var monthlyTotal = SumNullable(matches.Select(item => item.MonthlyTotalCost))
                               ?? AddNullable(monthlyEmployer, monthlyEmployee);
            var employeeCount = matches.Select(item => item.EmployeeCount).Where(value => value.HasValue).Sum(value => value!.Value);
            var classRates = rates.Where(rate => rate.ClassName.Equals(className, StringComparison.OrdinalIgnoreCase) && rate.VariancePercent.HasValue).ToList();
            var projectedChange = classRates.Count == 0
                ? (decimal?)null
                : classRates.Average(rate => rate.VariancePercent!.Value);
            var annualTotal = monthlyTotal * 12m;

            output.Add(new ClassCostSummaryRow
            {
                ClassName = className,
                DisplayName = ClassDisplayParts(className).DisplayName,
                EmployeeCount = employeeCount > 0 ? employeeCount : null,
                MonthlyEmployerCost = monthlyEmployer,
                MonthlyEmployeeCost = monthlyEmployee,
                MonthlyTotalCost = monthlyTotal,
                AnnualEmployerCost = monthlyEmployer * 12m,
                AnnualEmployeeCost = monthlyEmployee * 12m,
                AnnualTotalCost = annualTotal,
                CostPerEmployee = annualTotal.HasValue && employeeCount > 0 ? annualTotal.Value / employeeCount : null,
                ProjectedChangePercent = projectedChange,
                SourceDocument = FirstNonBlank(matches.Select(item => item.SourceDocument)),
                BarBrush = ClassPalette[index % ClassPalette.Length]
            });
        }

        var totalClientCost = output.Where(item => item.AnnualTotalCost.HasValue).Sum(item => item.AnnualTotalCost!.Value);
        foreach (var item in output.Where(item => item.AnnualTotalCost.HasValue && totalClientCost > 0m))
        {
            item.TotalClientCostPercent = item.AnnualTotalCost!.Value / totalClientCost;
        }

        return output;
    }

    private static void BuildComparisonRows(ClassComparisonWorkspaceViewModel model, IReadOnlyList<ClassComparisonClassCard> selected)
    {
        model.ComparisonRows.Clear();
        var selectedNames = selected.Select(item => item.ClassName).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var baselineName = model.BaselineClassName;
        var sections = model.SourceFeatures
            .Where(value => selectedNames.Contains(value.ClassName))
            .GroupBy(value => value.Section, StringComparer.OrdinalIgnoreCase)
            .OrderBy(group => SectionIndex(group.Key))
            .ThenBy(group => group.Key, StringComparer.OrdinalIgnoreCase);

        foreach (var section in sections)
        {
            var isExpanded = !model.CollapsedSections.Contains(section.Key);
            model.ComparisonRows.Add(new ClassComparisonMatrixRow
            {
                Section = section.Key,
                Benefit = section.Key,
                RowKey = $"section::{NormalizeKey(section.Key)}",
                IsSectionHeader = true,
                IsExpanded = isExpanded
            });

            if (!isExpanded)
            {
                continue;
            }

            foreach (var rowGroup in section
                         .GroupBy(value => value.RowKey, StringComparer.OrdinalIgnoreCase)
                         .OrderBy(group => BenefitOrder(group.First().Benefit))
                         .ThenBy(group => group.First().Benefit, StringComparer.OrdinalIgnoreCase))
            {
                var first = rowGroup.First();
                if (model.KeyBenefitsOnly && !first.Rule.IsKeyBenefit)
                {
                    continue;
                }

                var row = new ClassComparisonMatrixRow
                {
                    Section = first.Section,
                    Benefit = first.Benefit,
                    RowKey = first.RowKey,
                    IsKeyBenefit = first.Rule.IsKeyBenefit,
                    Rule = first.Rule
                };
                var baseline = rowGroup.FirstOrDefault(value => value.ClassName.Equals(baselineName, StringComparison.OrdinalIgnoreCase));
                foreach (var classCard in selected)
                {
                    var value = rowGroup.FirstOrDefault(item => item.ClassName.Equals(classCard.ClassName, StringComparison.OrdinalIgnoreCase));
                    row.Cells.Add(BuildCell(classCard.ClassName, value, baseline, first.Rule, model.HighlightDifferences));
                }

                (row.BestValueText, row.BestValueDetail) = BestValue(rowGroup.ToList(), selected, first.Rule);
                model.ComparisonRows.Add(row);
            }
        }
    }

    private static ClassComparisonCell BuildCell(
        string className,
        ClassComparisonFeatureValue? value,
        ClassComparisonFeatureValue? baseline,
        ClassComparisonRule rule,
        bool highlightDifferences)
    {
        var status = Compare(value?.NormalizedValue, baseline?.NormalizedValue, rule.Direction);
        return new ClassComparisonCell
        {
            ClassName = className,
            DisplayValue = value?.DisplayValue ?? "Not available",
            Status = status,
            HighlightDifferences = highlightDifferences,
            SourceDocument = value?.SourceDocument ?? "",
            SourcePage = value?.SourcePage ?? ""
        };
    }

    private static ClassComparisonStatus Compare(
        NormalizedClassComparisonValue? value,
        NormalizedClassComparisonValue? baseline,
        ClassComparisonDirection direction)
    {
        if (value is null || value.IsMissing || baseline is null || baseline.IsMissing)
        {
            return value is not null && baseline is not null && value.IsMissing && baseline.IsMissing
                ? ClassComparisonStatus.Same
                : ClassComparisonStatus.Missing;
        }

        if (ClassComparisonValueNormalizer.AreEquivalent(value, baseline))
        {
            return ClassComparisonStatus.Same;
        }

        if (direction == ClassComparisonDirection.Neutral)
        {
            return ClassComparisonStatus.Different;
        }

        if (direction == ClassComparisonDirection.IncludedIsBetter && value.BooleanValue.HasValue && baseline.BooleanValue.HasValue)
        {
            return value.BooleanValue.Value ? ClassComparisonStatus.Better : ClassComparisonStatus.Worse;
        }

        if (value.NumericValue.HasValue && baseline.NumericValue.HasValue)
        {
            var isBetter = direction == ClassComparisonDirection.HigherIsBetter
                ? value.NumericValue.Value > baseline.NumericValue.Value
                : value.NumericValue.Value < baseline.NumericValue.Value;
            return isBetter ? ClassComparisonStatus.Better : ClassComparisonStatus.Worse;
        }

        return ClassComparisonStatus.Different;
    }

    private static (string Text, string Detail) BestValue(
        IReadOnlyList<ClassComparisonFeatureValue> rowValues,
        IReadOnlyList<ClassComparisonClassCard> selected,
        ClassComparisonRule rule)
    {
        if (rule.Direction == ClassComparisonDirection.Neutral)
        {
            return ("Informational", rule.Explanation);
        }

        var selectedNames = selected.Select(item => item.ClassName).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var candidates = rowValues
            .Where(value => selectedNames.Contains(value.ClassName) && !value.NormalizedValue.IsMissing)
            .ToList();
        if (candidates.Count == 0)
        {
            return ("Not available", "No selected class has a comparable value.");
        }

        List<ClassComparisonFeatureValue> best;
        if (rule.Direction == ClassComparisonDirection.IncludedIsBetter && candidates.Any(value => value.NormalizedValue.BooleanValue.HasValue))
        {
            var bestBoolean = candidates.Max(value => value.NormalizedValue.BooleanValue == true ? 1 : 0);
            best = candidates.Where(value => (value.NormalizedValue.BooleanValue == true ? 1 : 0) == bestBoolean).ToList();
        }
        else if (candidates.All(value => value.NormalizedValue.NumericValue.HasValue))
        {
            var bestNumber = rule.Direction == ClassComparisonDirection.HigherIsBetter
                ? candidates.Max(value => value.NormalizedValue.NumericValue!.Value)
                : candidates.Min(value => value.NormalizedValue.NumericValue!.Value);
            best = candidates.Where(value => Math.Abs(value.NormalizedValue.NumericValue!.Value - bestNumber) <= 0.001m).ToList();
        }
        else
        {
            return ("Review", "Values differ but cannot be ranked reliably with the selected rule.");
        }

        return best.Count == 1
            ? (ClassDisplayParts(best[0].ClassName).DisplayName, rule.Explanation)
            : ("Tie", $"{best.Count} selected classes share the best comparable value.");
    }

    private static void BuildVisibleCosts(ClassComparisonWorkspaceViewModel model, IReadOnlyList<ClassComparisonClassCard> selected)
    {
        model.CostSummaries.Clear();
        var selectedNames = selected.Select(item => item.ClassName).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var costs = model.SourceCosts.Where(item => selectedNames.Contains(item.ClassName)).ToList();
        var max = costs.Where(item => item.AnnualTotalCost.HasValue).Select(item => item.AnnualTotalCost!.Value).DefaultIfEmpty(0m).Max();
        foreach (var cost in costs)
        {
            cost.BarHeight = cost.AnnualTotalCost.HasValue && max > 0m
                ? Math.Max(24d, (double)(cost.AnnualTotalCost.Value / max) * 148d)
                : 0d;
            model.CostSummaries.Add(cost);
        }
    }

    private static void BuildVisibleRates(ClassComparisonWorkspaceViewModel model, IReadOnlyList<ClassComparisonClassCard> selected)
    {
        var selectedNames = selected.Select(item => item.ClassName).ToHashSet(StringComparer.OrdinalIgnoreCase);
        Replace(model.RateRows, model.SourceRates.Where(item => selectedNames.Contains(item.ClassName)));
    }

    private static void BuildScores(ClassComparisonWorkspaceViewModel model, IReadOnlyList<ClassComparisonClassCard> selected)
    {
        model.ValueScores.Clear();
        var selectedNames = selected.Select(item => item.ClassName).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var comparableRows = model.SourceFeatures
            .Where(value => selectedNames.Contains(value.ClassName))
            .GroupBy(value => value.RowKey, StringComparer.OrdinalIgnoreCase)
            .ToList();
        var rawScores = new List<ClassValueScoreRow>();

        foreach (var classCard in selected)
        {
            var richness = AverageScore(comparableRows.Where(group => IsRichnessSection(group.First().Section)), classCard.ClassName);
            var eligibility = AverageScore(comparableRows.Where(group => group.First().Section.Equals("Eligibility", StringComparison.OrdinalIgnoreCase)), classCard.ClassName);
            var breadth = CoverageBreadth(comparableRows, classCard.ClassName);
            var competitiveness = AverageNullable(richness, breadth);
            var cost = model.SourceCosts.FirstOrDefault(item => item.ClassName.Equals(classCard.ClassName, StringComparison.OrdinalIgnoreCase));
            var employeeCost = CostScore(model.SourceCosts, classCard.ClassName, item => item.AnnualEmployeeCost);
            var employerCost = CostScore(model.SourceCosts, classCard.ClassName, item => item.AnnualEmployerCost);
            var employeeImpact = AverageNullable(richness, breadth, employeeCost);
            var components = new List<ClassValueScoreComponent>
            {
                Component("Plan richness", richness, "Normalized benefit values using explicit higher, lower, included, and neutral rules."),
                Component("Employee cost", employeeCost, "Lower annual employee cost scores better when class-level cost data exists."),
                Component("Employer cost", employerCost, "Lower annual employer cost scores better when class-level cost data exists."),
                Component("Eligibility", eligibility, "Eligibility provisions are scored only where a directional rule is defined."),
                Component("Coverage breadth", breadth, "Share of comparison rows with a saved, included value."),
                Component("Competitiveness", competitiveness, "Average of plan richness and coverage breadth; no market benchmark is implied."),
                Component("Administration complexity", null, "Not scored until structured administration data is saved."),
                Component("Employee impact", employeeImpact, "Average of richness, coverage breadth, and employee cost where available.")
            };
            var weighted = components.Where(component => component.Score.HasValue && component.Weight > 0m).ToList();
            var weightTotal = weighted.Sum(component => component.Weight);
            var overall = weightTotal <= 0m
                ? 0m
                : weighted.Sum(component => component.Score!.Value * component.Weight) / weightTotal;
            rawScores.Add(new ClassValueScoreRow
            {
                ClassName = classCard.ClassName,
                DisplayName = classCard.DisplayName,
                Score = (int)Math.Round(Math.Clamp(overall, 0m, 100m), MidpointRounding.AwayFromZero),
                BarBrush = classCard.AccentBrush,
                Components = components
            });
        }

        var ranked = rawScores.OrderByDescending(item => item.Score).ThenBy(item => item.DisplayName, StringComparer.OrdinalIgnoreCase).ToList();
        for (var index = 0; index < ranked.Count; index++)
        {
            ranked[index].Rank = index + 1;
        }

        var savedRecommendation = selected.FirstOrDefault(item => item.ClassName.Equals(model.RecommendedClassName, StringComparison.OrdinalIgnoreCase));
        if (savedRecommendation is null)
        {
            var hasUniqueLeader = ranked.Count == 1 || ranked[0].Score > ranked[1].Score;
            model.RecommendedClassName = hasUniqueLeader ? ranked[0].ClassName : "";
        }

        foreach (var score in ranked)
        {
            score.IsRecommended = score.ClassName.Equals(model.RecommendedClassName, StringComparison.OrdinalIgnoreCase);
            model.ValueScores.Add(score);
        }

        foreach (var classCard in model.Classes)
        {
            var score = ranked.FirstOrDefault(item => item.ClassName.Equals(classCard.ClassName, StringComparison.OrdinalIgnoreCase));
            classCard.ValueScore = score?.Score;
            classCard.IsRecommended = classCard.ClassName.Equals(model.RecommendedClassName, StringComparison.OrdinalIgnoreCase);
        }
    }

    private static ClassValueScoreComponent Component(string category, decimal? score, string explanation)
    {
        return new ClassValueScoreComponent
        {
            Category = category,
            Weight = ScoreWeights[category],
            Score = score,
            Explanation = explanation
        };
    }

    private static decimal? AverageScore(
        IEnumerable<IGrouping<string, ClassComparisonFeatureValue>> groups,
        string className)
    {
        var scores = groups
            .Select(group => RelativeFeatureScore(group.ToList(), className))
            .Where(value => value.HasValue)
            .Select(value => value!.Value)
            .ToList();
        return scores.Count == 0 ? null : scores.Average();
    }

    private static decimal? RelativeFeatureScore(IReadOnlyList<ClassComparisonFeatureValue> values, string className)
    {
        var target = values.FirstOrDefault(value => value.ClassName.Equals(className, StringComparison.OrdinalIgnoreCase));
        if (target is null || target.NormalizedValue.IsMissing || target.Rule.Direction == ClassComparisonDirection.Neutral)
        {
            return null;
        }

        if (target.Rule.Direction == ClassComparisonDirection.IncludedIsBetter && target.NormalizedValue.BooleanValue.HasValue)
        {
            return target.NormalizedValue.BooleanValue.Value ? 100m : 40m;
        }

        var numeric = values.Where(value => !value.NormalizedValue.IsMissing && value.NormalizedValue.NumericValue.HasValue).ToList();
        if (!target.NormalizedValue.NumericValue.HasValue || numeric.Count == 0)
        {
            return null;
        }

        var min = numeric.Min(value => value.NormalizedValue.NumericValue!.Value);
        var max = numeric.Max(value => value.NormalizedValue.NumericValue!.Value);
        if (Math.Abs(max - min) <= 0.001m)
        {
            return 80m;
        }

        var position = (target.NormalizedValue.NumericValue.Value - min) / (max - min);
        if (target.Rule.Direction == ClassComparisonDirection.LowerIsBetter)
        {
            position = 1m - position;
        }

        return 50m + (50m * position);
    }

    private static decimal CoverageBreadth(
        IReadOnlyList<IGrouping<string, ClassComparisonFeatureValue>> groups,
        string className)
    {
        if (groups.Count == 0)
        {
            return 0m;
        }

        decimal points = 0m;
        foreach (var group in groups)
        {
            var value = group.FirstOrDefault(item => item.ClassName.Equals(className, StringComparison.OrdinalIgnoreCase));
            if (value is null || value.NormalizedValue.IsMissing)
            {
                continue;
            }

            points += value.NormalizedValue.BooleanValue == false ? 0.4m : 1m;
        }

        return points / groups.Count * 100m;
    }

    private static decimal? CostScore(
        IReadOnlyList<ClassCostSummaryRow> costs,
        string className,
        Func<ClassCostSummaryRow, decimal?> selector)
    {
        var target = costs.FirstOrDefault(item => item.ClassName.Equals(className, StringComparison.OrdinalIgnoreCase));
        var available = costs.Select(selector).Where(value => value.HasValue).Select(value => value!.Value).ToList();
        var targetValue = target is null ? null : selector(target);
        if (!targetValue.HasValue || available.Count == 0)
        {
            return null;
        }

        var min = available.Min();
        var max = available.Max();
        if (Math.Abs(max - min) <= 0.01m)
        {
            return 80m;
        }

        return 50m + (50m * (1m - ((targetValue.Value - min) / (max - min))));
    }

    private static decimal? AverageNullable(params decimal?[] values)
    {
        var available = values.Where(value => value.HasValue).Select(value => value!.Value).ToList();
        return available.Count == 0 ? null : available.Average();
    }

    private static void BuildDifferenceSummaries(ClassComparisonWorkspaceViewModel model, IReadOnlyList<ClassComparisonClassCard> selected)
    {
        model.DifferenceSummaries.Clear();
        var baseline = model.BaselineClass;
        if (baseline is null)
        {
            return;
        }

        foreach (var row in model.ComparisonRows.Where(row => !row.IsSectionHeader))
        {
            for (var index = 0; index < row.Cells.Count && index < selected.Count; index++)
            {
                var cell = row.Cells[index];
                if (cell.ClassName.Equals(baseline.ClassName, StringComparison.OrdinalIgnoreCase) ||
                    cell.Status is ClassComparisonStatus.Same or ClassComparisonStatus.Missing)
                {
                    continue;
                }

                var classDisplay = selected[index].DisplayName;
                var relationship = cell.Status switch
                {
                    ClassComparisonStatus.Better => "provides a better comparable value for",
                    ClassComparisonStatus.Worse => "provides a lower comparable value for",
                    _ => "differs from the baseline for"
                };
                model.DifferenceSummaries.Add(new ClassDifferenceSummaryRow
                {
                    Summary = $"{classDisplay} {relationship} {row.Benefit}.",
                    ClassName = cell.ClassName,
                    Benefit = row.Benefit,
                    Status = cell.Status
                });
                if (model.DifferenceSummaries.Count >= 12)
                {
                    return;
                }
            }
        }

        if (model.DifferenceSummaries.Count == 0)
        {
            model.DifferenceSummaries.Add(new ClassDifferenceSummaryRow
            {
                Summary = "No material normalized differences were detected across the selected classes.",
                Status = ClassComparisonStatus.Same
            });
        }
    }

    private static string ResolveSection(ClassBenefitSourceRow row)
    {
        var text = $"{row.Benefit} {row.Category} {row.FeatureName} {row.SourceSection}".ToLowerInvariant();
        if (ContainsAny(text, "eligibility", "employee definition", "minimum hours", "actively at work", "dependent eligibility")) return "Eligibility";
        if (ContainsAny(text, "life", "ad&d", "accidental death", "dependent life", "optional life")) return "Life Insurance";
        if (ContainsAny(text, "disability", "ltd", "std")) return "Disability";
        if (ContainsAny(text, "dental", "orthodont", "restorative", "recall", "fluoride", "scaling", "x-ray", "sealant")) return "Dental Care";
        if (ContainsAny(text, "cost sharing", "employee contribution", "employer contribution", "single rate", "family rate")) return "Cost Sharing";
        if (ContainsAny(text, "health", "drug", "paramedical", "practitioner", "mental", "vision", "hospital", "travel", "nursing", "medical supplies", "eap", "efap", "virtual care", "medical second opinion")) return "Extended Health Care";
        return "Other";
    }

    private static (string Benefit, string PrimaryValue) ResolveBenefitLabelAndPrimaryValue(ClassBenefitSourceRow row)
    {
        var feature = FirstNonBlank(row.FeatureName, row.Category, row.Benefit, "Plan Design");
        var primary = FirstNonBlank(row.FeatureValue, row.Maximum, row.Coinsurance, row.Deductible);
        var prefixMatch = LabelPrefixRegex().Match(primary);
        if (prefixMatch.Success && IsBroadFeatureName(feature, row))
        {
            feature = prefixMatch.Groups["label"].Value.Trim();
            primary = prefixMatch.Groups["value"].Value.Trim();
        }

        return (CanonicalBenefitLabel(feature, row), primary);
    }

    private static bool IsBroadFeatureName(string feature, ClassBenefitSourceRow row)
    {
        return feature.Equals(row.Benefit, StringComparison.OrdinalIgnoreCase) ||
               feature.Equals(row.Category, StringComparison.OrdinalIgnoreCase) ||
               feature.Equals(row.SourceSection, StringComparison.OrdinalIgnoreCase) ||
               ContainsAny(feature.ToLowerInvariant(), "life / ad&d", "ltd", "prescription drugs", "medical services", "paramedical", "hospital", "vision care", "travel", "dental care", "virtual care", "eap");
    }

    private static string CanonicalBenefitLabel(string value, ClassBenefitSourceRow row)
    {
        var text = Regex.Replace(value.Replace('_', ' '), @"\s+", " ").Trim();
        var normalized = text.ToLowerInvariant();
        if (normalized.Contains("life benefit amount", StringComparison.Ordinal)) return "Employee Basic Life";
        if (normalized.Contains("ad&d benefit amount", StringComparison.Ordinal)) return "AD&D";
        if (normalized.Contains("no-evidence", StringComparison.Ordinal) || normalized.Contains("non-evidence", StringComparison.Ordinal)) return "Non-Evidence Maximum";
        if (normalized.Contains("benefit reduction", StringComparison.Ordinal)) return "Reduction Schedule";
        if (normalized.Equals("benefit schedule", StringComparison.Ordinal) && ContainsAny($"{row.Benefit} {row.SourceSection}".ToLowerInvariant(), "ltd", "disability")) return "LTD Benefit Formula";
        if (normalized.Equals("coinsurance", StringComparison.Ordinal) && ContainsAny($"{row.Benefit} {row.SourceSection}".ToLowerInvariant(), "drug")) return "Prescription Drugs Reimbursement";
        if (normalized.Equals("coverage", StringComparison.Ordinal)) return FirstNonBlank(row.Benefit, row.SourceSection, text);
        return CultureInfo.InvariantCulture.TextInfo.ToTitleCase(text.ToLowerInvariant())
            .Replace("Ad&D", "AD&D", StringComparison.Ordinal)
            .Replace("Ltd", "LTD", StringComparison.Ordinal)
            .Replace("Eap", "EAP", StringComparison.Ordinal)
            .Replace("Efap", "EFAP", StringComparison.Ordinal);
    }

    private static string BuildDisplayValue(ClassBenefitSourceRow row, string benefit, string primaryValue)
    {
        var value = FirstNonBlank(primaryValue, row.Maximum, row.Coinsurance, row.Deductible,
            row.Frequency, row.WaitingPeriod, row.TerminationAge);
        return string.IsNullOrWhiteSpace(value)
            ? "Not available"
            : BenefitSummaryDisplayFormatter.FormatDisplayValue(value, benefit);
    }

    private static bool TryBuildClientFacingValue(
        ClassBenefitSourceRow row,
        string section,
        string benefit,
        string primaryValue,
        out string normalizedBenefit,
        out string displayValue)
    {
        if (BenefitSummaryDisplayFormatter.TryFormatLongTermDisabilityBenefitFormula(
                section,
                primaryValue,
                out var longTermDisabilityFormula))
        {
            normalizedBenefit = benefit;
            displayValue = longTermDisabilityFormula;
            return true;
        }

        var candidates = new (string Label, string Value)[]
        {
            (benefit, primaryValue),
            ("Coinsurance", row.Coinsurance),
            ("Deductible", row.Deductible),
            ("Maximum", row.Maximum),
            ("Frequency", row.Frequency),
            ("Waiting Period", row.WaitingPeriod),
            ("Termination", row.TerminationAge)
        };

        foreach (var candidate in candidates.Where(candidate => !string.IsNullOrWhiteSpace(candidate.Value)))
        {
            if (BenefitSummaryDisplayFormatter.TryFormatClientFacingValue(
                    section,
                    candidate.Label,
                    candidate.Value,
                    out _,
                    out var summaryValue))
            {
                // Keep the established comparison key/row label stable. Benefit Summary controls whether
                // the value is client-facing and how it is formatted; it does not rename comparison rows.
                normalizedBenefit = benefit;
                displayValue = summaryValue;
                return true;
            }
        }

        normalizedBenefit = benefit;
        displayValue = "";
        return false;
    }

    private static bool IsRichnessSection(string section)
    {
        return section is "Life Insurance" or "Disability" or "Extended Health Care" or "Dental Care" or "Other";
    }

    private static int SectionIndex(string section)
    {
        var index = Array.FindIndex(SectionOrder, value => value.Equals(section, StringComparison.OrdinalIgnoreCase));
        return index < 0 ? SectionOrder.Length : index;
    }

    private static int BenefitOrder(string benefit)
    {
        var text = benefit.ToLowerInvariant();
        if (ContainsAny(text, "employee basic life", "employee life")) return 1;
        if (ContainsAny(text, "ad&d", "accidental death")) return 2;
        if (ContainsAny(text, "dependent life")) return 3;
        if (ContainsAny(text, "waiting", "elimination")) return 10;
        if (ContainsAny(text, "deductible")) return 20;
        if (ContainsAny(text, "prescription", "drug")) return 21;
        if (ContainsAny(text, "paramedical")) return 22;
        if (ContainsAny(text, "mental health")) return 23;
        if (ContainsAny(text, "vision")) return 24;
        if (ContainsAny(text, "hospital")) return 25;
        if (ContainsAny(text, "travel")) return 26;
        if (ContainsAny(text, "major restorative")) return 40;
        if (ContainsAny(text, "orthodont")) return 41;
        return 50;
    }

    private static (string DisplayName, string Description, string ClassCode) ClassDisplayParts(string className)
    {
        var text = FirstNonBlank(className, "All Classes");
        var splitIndex = text.IndexOf(" - ", StringComparison.Ordinal);
        if (splitIndex <= 0)
        {
            return (text.Length <= 38 ? text : $"{text[..35]}...", "", ExtractClassCode(text));
        }

        var prefix = text[..splitIndex].Trim();
        var suffix = text[(splitIndex + 3)..].Trim();
        var display = prefix.StartsWith("Class ", StringComparison.OrdinalIgnoreCase)
            ? prefix
            : text.Length <= 38
                ? text
                : suffix.Length <= 32
                    ? suffix
                    : prefix;
        return (display, suffix, ExtractClassCode(text));
    }

    private static string ExtractClassCode(string className)
    {
        var match = ClassCodeRegex().Match(className.Trim());
        if (match.Success)
        {
            return match.Groups["code"].Value.ToUpperInvariant();
        }

        var prefix = className.Split('-', 2)[0].Trim();
        return prefix.Length is > 0 and <= 4 ? prefix.ToUpperInvariant() : "";
    }

    private static bool ClassNamesMatch(string planClassName, string sourceClassName)
    {
        if (string.IsNullOrWhiteSpace(sourceClassName))
        {
            return IsGenericClass(planClassName);
        }

        var left = NormalizeKey(planClassName);
        var right = NormalizeKey(sourceClassName);
        if (left.Equals(right, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        var leftCode = ExtractClassCode(planClassName);
        var rightCode = ExtractClassCode(sourceClassName);
        return leftCode.Length > 0 && rightCode.Length > 0 &&
               leftCode.Equals(rightCode, StringComparison.OrdinalIgnoreCase) &&
               !ContainsAny(leftCode.ToLowerInvariant(), "single", "family", "couple");
    }

    private static bool IsGenericClass(string value)
    {
        var text = NormalizeKey(FirstNonBlank(value, "All Classes"));
        return text is "all" or "allclass" or "allclasses" or "allemployees" or "alleligibleemployees";
    }

    private static string ClassSortKey(string value)
    {
        var code = ExtractClassCode(value);
        return string.IsNullOrWhiteSpace(code) ? value : $"{code}|{value}";
    }

    private static string NormalizeKey(string value)
    {
        return new string((value ?? "").Where(char.IsLetterOrDigit).Select(char.ToLowerInvariant).ToArray());
    }

    private static decimal? SumNullable(IEnumerable<decimal?> values)
    {
        var available = values.Where(value => value.HasValue).Select(value => value!.Value).ToList();
        return available.Count == 0 ? null : available.Sum();
    }

    private static decimal? AddNullable(decimal? left, decimal? right)
    {
        return left.HasValue || right.HasValue ? (left ?? 0m) + (right ?? 0m) : null;
    }

    private static bool ContainsAny(string text, params string[] terms)
    {
        return terms.Any(term =>
            !string.IsNullOrWhiteSpace(term) &&
            Regex.IsMatch(
                text,
                $@"(?<![A-Za-z0-9]){Regex.Escape(term)}",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant));
    }

    private static string FirstNonBlank(params string?[] values)
    {
        return FirstNonBlank((IEnumerable<string?>)values);
    }

    private static string FirstNonBlank(IEnumerable<string?> values)
    {
        foreach (var value in values)
        {
            if (!string.IsNullOrWhiteSpace(value))
            {
                return value.Trim();
            }
        }

        return "";
    }

    private static void Replace<T>(ICollection<T> target, IEnumerable<T> values)
    {
        target.Clear();
        foreach (var value in values)
        {
            target.Add(value);
        }
    }

    [GeneratedRegex(@"^(?<label>[A-Za-z][A-Za-z0-9 &/()'\-]{1,60}):\s*(?<value>.+)$", RegexOptions.CultureInvariant)]
    private static partial Regex LabelPrefixRegex();

    [GeneratedRegex(@"^(?:class\s+)?(?<code>[A-Za-z0-9]{1,4})(?:\s*-|\s|$)", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex ClassCodeRegex();
}

public static partial class ClassComparisonValueNormalizer
{
    public static NormalizedClassComparisonValue Normalize(string? value, ClassComparisonValueKind preferredKind = ClassComparisonValueKind.Auto)
    {
        var text = Regex.Replace(value ?? "", @"\s+", " ").Trim();
        if (IsMissing(text))
        {
            return new NormalizedClassComparisonValue(preferredKind, null, null, "", true);
        }

        var lower = text.ToLowerInvariant();
        if (preferredKind is ClassComparisonValueKind.Money or ClassComparisonValueKind.Number &&
            ContainsAny(lower, "nil", "none", "no deductible", "$0", "zero"))
        {
            return new NormalizedClassComparisonValue(preferredKind, 0m, null, "0", false);
        }

        var falseValue = ContainsAny(lower, "not included", "no coverage", "not covered", "excluded", "false");
        var trueValue = !falseValue && ContainsAny(lower, "included", "covered", "true", "yes", "available");
        if (preferredKind == ClassComparisonValueKind.Boolean || falseValue || trueValue)
        {
            return new NormalizedClassComparisonValue(ClassComparisonValueKind.Boolean, null, trueValue, trueValue ? "true" : "false", false);
        }

        if (preferredKind == ClassComparisonValueKind.Percent || PercentRegex().IsMatch(text))
        {
            var number = FirstNumber(text);
            if (number.HasValue)
            {
                var normalized = text.Contains('%') || number.Value > 1m ? number.Value / 100m : number.Value;
                return new NormalizedClassComparisonValue(ClassComparisonValueKind.Percent, normalized, null, normalized.ToString(CultureInfo.InvariantCulture), false);
            }
        }

        if (preferredKind == ClassComparisonValueKind.Duration || DurationRegex().IsMatch(lower))
        {
            var duration = DurationInDays(lower);
            if (duration.HasValue)
            {
                return new NormalizedClassComparisonValue(ClassComparisonValueKind.Duration, duration, null, duration.Value.ToString(CultureInfo.InvariantCulture), false);
            }
        }

        if (preferredKind == ClassComparisonValueKind.Money || MoneyRegex().IsMatch(text))
        {
            var number = FirstNumber(text);
            if (number.HasValue)
            {
                return new NormalizedClassComparisonValue(ClassComparisonValueKind.Money, number, null, number.Value.ToString(CultureInfo.InvariantCulture), false);
            }
        }

        if (preferredKind == ClassComparisonValueKind.Number)
        {
            var number = FirstNumber(text);
            if (number.HasValue)
            {
                return new NormalizedClassComparisonValue(ClassComparisonValueKind.Number, number, null, number.Value.ToString(CultureInfo.InvariantCulture), false);
            }
        }

        return new NormalizedClassComparisonValue(
            ClassComparisonValueKind.Text,
            null,
            null,
            NormalizeText(text),
            false);
    }

    public static bool AreEquivalent(NormalizedClassComparisonValue left, NormalizedClassComparisonValue right)
    {
        if (left.IsMissing || right.IsMissing)
        {
            return left.IsMissing && right.IsMissing;
        }

        if (left.BooleanValue.HasValue || right.BooleanValue.HasValue)
        {
            return left.BooleanValue.HasValue && right.BooleanValue.HasValue && left.BooleanValue.Value == right.BooleanValue.Value;
        }

        if (left.NumericValue.HasValue && right.NumericValue.HasValue)
        {
            var tolerance = left.Kind == ClassComparisonValueKind.Duration || right.Kind == ClassComparisonValueKind.Duration ? 1m : 0.001m;
            return Math.Abs(left.NumericValue.Value - right.NumericValue.Value) <= tolerance;
        }

        return left.TextValue.Equals(right.TextValue, StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsMissing(string value)
    {
        return string.IsNullOrWhiteSpace(value) ||
               value.Equals("n/a", StringComparison.OrdinalIgnoreCase) ||
               value.Equals("not available", StringComparison.OrdinalIgnoreCase) ||
               value.Equals("-", StringComparison.OrdinalIgnoreCase) ||
               value.Equals("—", StringComparison.OrdinalIgnoreCase);
    }

    private static decimal? FirstNumber(string value)
    {
        var match = NumberRegex().Match(value);
        if (!match.Success)
        {
            return null;
        }

        var cleaned = match.Groups["number"].Value.Replace(",", "", StringComparison.Ordinal);
        return decimal.TryParse(cleaned, NumberStyles.Number, CultureInfo.InvariantCulture, out var number) ? number : null;
    }

    private static decimal? DurationInDays(string value)
    {
        var match = DurationRegex().Match(value);
        if (!match.Success || !decimal.TryParse(match.Groups["number"].Value, NumberStyles.Number, CultureInfo.InvariantCulture, out var number))
        {
            return null;
        }

        var unit = match.Groups["unit"].Value.ToLowerInvariant();
        return unit switch
        {
            "day" or "days" => number,
            "week" or "weeks" => number * 7m,
            "month" or "months" => number * 30.4375m,
            "year" or "years" => number * 365.25m,
            "age" => number * 365.25m,
            _ => null
        };
    }

    private static string NormalizeText(string value)
    {
        return new string(value.ToLowerInvariant().Where(char.IsLetterOrDigit).ToArray());
    }

    private static bool ContainsAny(string text, params string[] terms)
    {
        return terms.Any(term =>
            !string.IsNullOrWhiteSpace(term) &&
            Regex.IsMatch(
                text,
                $@"(?<![A-Za-z0-9]){Regex.Escape(term)}",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant));
    }

    [GeneratedRegex(@"(?<number>-?\d[\d,]*(?:\.\d+)?)", RegexOptions.CultureInvariant)]
    private static partial Regex NumberRegex();

    [GeneratedRegex(@"-?\d[\d,]*(?:\.\d+)?\s*%", RegexOptions.CultureInvariant)]
    private static partial Regex PercentRegex();

    [GeneratedRegex(@"\$\s*-?\d", RegexOptions.CultureInvariant)]
    private static partial Regex MoneyRegex();

    [GeneratedRegex(@"(?:(?:age)\s*)?(?<number>\d+(?:\.\d+)?)\s*(?<unit>days?|weeks?|months?|years?)|(?<unit>age)\s*(?<number>\d+(?:\.\d+)?)", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex DurationRegex();
}

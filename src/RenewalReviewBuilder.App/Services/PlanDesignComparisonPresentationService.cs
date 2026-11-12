using System.Windows.Media;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed class PlanDesignComparisonPresentationService
{
    private readonly PlanDesignComparisonDomainService _domainService;

    public PlanDesignComparisonPresentationService(PlanDesignComparisonDomainService? domainService = null)
    {
        _domainService = domainService ?? new PlanDesignComparisonDomainService();
    }

    public PlanDesignComparisonWorkspaceViewModel Build(CurrentVsQuotePlanDesignComparison comparison)
    {
        ArgumentNullException.ThrowIfNull(comparison);

        var model = new PlanDesignComparisonWorkspaceViewModel();
        model.BaselineOptions.Add(new PlanDesignComparisonOption
        {
            Key = "current",
            DisplayName = "Current Plan",
            Detail = comparison.CurrentCarrier,
            IsBaseline = true
        });

        foreach (var quote in comparison.Quotes)
        {
            model.ScenarioOptions.Add(new PlanDesignComparisonOption
            {
                Key = quote.QuoteKey,
                DisplayName = FirstNonBlank(quote.DisplayName, "Saved scenario"),
                Detail = FirstNonBlank(quote.Reference, quote.SourceFileName)
            });
        }

        if (model.ScenarioOptions.Count == 0)
        {
            model.RefreshPresentation();
            return model;
        }

        model.SelectedScenarioKey = model.ScenarioOptions[0].Key;
        foreach (var sourceRow in comparison.Rows)
        {
            foreach (var selected in sourceRow.QuoteValues)
            {

            var classification = _domainService.Classify(
                sourceRow.ProvisionName,
                sourceRow.CurrentValue,
                selected.Value,
                selected.Status,
                sourceRow.CurrentAvailability,
                selected.Availability);
            var groupName = MapCanonicalCategory(sourceRow.BenefitCategory, sourceRow.CanonicalBenefitCode);
            var group = model.Groups.FirstOrDefault(item => item.Category.Equals(groupName, StringComparison.OrdinalIgnoreCase));
            if (group is null)
            {
                group = new PlanDesignComparisonGroup { Category = groupName, Order = CategoryOrder(groupName) };
                model.Groups.Add(group);
            }

            group.Rows.Add(new PlanDesignComparisonProvisionRow
            {
                ScenarioKey = selected.QuoteKey,
                Category = groupName,
                BenefitName = FirstNonBlank(sourceRow.BenefitName, sourceRow.BenefitCategory),
                ProvisionName = sourceRow.ProvisionName,
                CurrentValue = DisplayValue(sourceRow.CurrentValue),
                SelectedValue = DisplayValue(selected.Value),
                DifferenceStatus = classification.Status,
                DifferenceRationale = classification.Rationale,
                CostImpactText = PlanDesignComparisonDomainService.FormatCost(null),
                ValueRecommendation = classification.Recommendation,
                ValidationWarning = selected.ValidationStatus,
                IsKeyBenefit = IsKeyBenefit(sourceRow.CanonicalBenefitCode, sourceRow.BenefitCategory),
                CurrentEvidence = ToEvidence(sourceRow.CurrentEvidence),
                SelectedEvidence = ToEvidence(selected.Evidence)
            });
            }
        }

        model.ViewMode = model.HasDifferences
            ? PlanDesignComparisonViewModes.ChangedOnly
            : PlanDesignComparisonViewModes.AllBenefits;
        AddSummaryMetrics(model);
        model.RefreshPresentation();
        return model;
    }

    private static void AddSummaryMetrics(PlanDesignComparisonWorkspaceViewModel model)
    {
        var rows = model.Groups.SelectMany(group => group.Rows).ToList();
        AddMetric("Benefits Changed", rows.Count(row => !row.DifferenceStatus.Equals(PlanDesignComparisonStatuses.Same, StringComparison.OrdinalIgnoreCase)).ToString(), "Comparable provisions", Brushes.SlateGray);
        AddMetric("Coverage Improved", rows.Count(row => row.DifferenceStatus == PlanDesignComparisonStatuses.Improved).ToString(), "Domain-supported", Brushes.SeaGreen);
        AddMetric("Coverage Reduced", rows.Count(row => row.DifferenceStatus == PlanDesignComparisonStatuses.Reduced).ToString(), "Domain-supported", Brushes.Firebrick);
        AddMetric("Administrative Changes", rows.Count(row => row.Category == "Eligibility & Administration" && !row.DifferenceStatus.Equals(PlanDesignComparisonStatuses.Same, StringComparison.OrdinalIgnoreCase)).ToString(), "Review as applicable", Brushes.DarkOrange);
        AddMetric("Estimated Cost Impact", "Not available", "No mapped scenario cost", Brushes.SlateGray);
        AddMetric("Value Score Impact", "Not available", "No traceable score", Brushes.SlateGray);
        AddMetric("Compliance Score", rows.Any(row => row.DifferenceStatus == PlanDesignComparisonStatuses.ReviewRequired) ? "Review required" : "Not available", "No mapped compliance score", Brushes.SlateGray);

        void AddMetric(string label, string value, string detail, Brush foreground) => model.SummaryMetrics.Add(new PlanDesignComparisonMetric
        {
            Label = label,
            Value = value,
            Detail = detail,
            ForegroundBrush = foreground
        });
    }

    private static PlanDesignComparisonEvidence ToEvidence(PlanDesignCanonicalEvidence? evidence)
    {
        if (evidence is null)
        {
            return new PlanDesignComparisonEvidence();
        }

        var document = FirstNonBlank(evidence.SourceFileName, evidence.SourceDocument, evidence.TrackedSourceFileName);
        var reference = string.IsNullOrWhiteSpace(document)
            ? "Not available"
            : evidence.SourcePage.HasValue
                ? $"{document}, p. {evidence.SourcePage.Value}"
                : document;
        return new PlanDesignComparisonEvidence
        {
            ReferenceText = reference,
            SourceText = evidence.SourceText,
            SourceDocument = document,
            SourcePage = evidence.SourcePage,
            IsAvailable = !string.IsNullOrWhiteSpace(document) || !string.IsNullOrWhiteSpace(evidence.SourceText)
        };
    }

    private static string MapCanonicalCategory(string category, string canonicalBenefitCode)
    {
        var normalized = FirstNonBlank(canonicalBenefitCode, category).Trim().ToLowerInvariant();
        if (normalized.Contains("life") || normalized.Contains("ad&d") || normalized.Contains("add")) return "Life & AD&D";
        if (normalized.Contains("disability") || normalized.Contains("ltd") || normalized.Contains("std")) return "Disability";
        if (normalized.Contains("drug")) return "Drugs";
        if (normalized.Contains("paramedical") || normalized.Contains("practitioner")) return "Paramedicals";
        if (normalized.Contains("vision")) return "Vision";
        if (normalized.Contains("dental")) return "Dental";
        if (normalized.Contains("spending") || normalized.Contains("hcsa") || normalized.Contains("lsa")) return "Spending Accounts";
        if (normalized.Contains("eligib") || normalized.Contains("administration") || normalized.Contains("termination")) return "Eligibility & Administration";
        if (normalized.Contains("health") || normalized.Contains("ehc")) return "Health";
        return "Other";
    }

    private static int CategoryOrder(string category) => category switch
    {
        "Life & AD&D" => 1,
        "Disability" => 2,
        "Health" => 3,
        "Drugs" => 4,
        "Paramedicals" => 5,
        "Vision" => 6,
        "Dental" => 7,
        "Spending Accounts" => 8,
        "Eligibility & Administration" => 9,
        _ => 10
    };

    private static bool IsKeyBenefit(string canonicalBenefitCode, string category)
    {
        var normalized = FirstNonBlank(canonicalBenefitCode, category).ToLowerInvariant();
        return normalized.Contains("life") || normalized.Contains("disability") || normalized.Contains("health") || normalized.Contains("drug") || normalized.Contains("dental");
    }

    private static string DisplayValue(string value) => string.IsNullOrWhiteSpace(value) || value.Equals("Not found", StringComparison.OrdinalIgnoreCase)
        ? "Not available"
        : value.Trim();

    private static string FirstNonBlank(params string?[] values) => values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value))?.Trim() ?? "";
}

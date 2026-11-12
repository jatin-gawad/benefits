using System.Text.RegularExpressions;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services.PolicySchedules;

/// <summary>Maps Manulife source facts to existing carrier-neutral Plan Design rows.</summary>
public sealed class ManulifePolicyScheduleMapper
{
    public PolicyScheduleMappingResult Map(PolicyScheduleExtractionResult extraction, ClientRecord client)
    {
        var result = new PolicyScheduleMappingResult();
        foreach (var section in extraction.Sections)
        {
            var classes = section.ClassificationCodes.Count == 0 ? ["All Classes"] : section.ClassificationCodes.Select(code => $"Class {code}").ToList();
            if (section.RawClassifications.Count > 0 && section.ClassificationCodes.Count == 0)
            {
                result.ValidationIssues.Add(Issue(ValidationIssueCodes.PolicyScheduleClassMappingRequired, "classification_code", $"Class mapping is required for '{string.Join(", ", section.RawClassifications)}'.", extraction.SourcePath));
            }

            foreach (var field in section.Fields)
            {
                var provision = NormalizeProvision(field.Label, field.Value, section.Benefit);
                if (provision is null)
                {
                    result.ValidationIssues.Add(Issue(ValidationIssueCodes.PolicyScheduleUnsupportedTerm, "source_field", $"Unsupported Manulife term '{field.Label}' was retained for review.", field.SourceExcerpt));
                    continue;
                }

                foreach (var className in classes)
                {
                    var row = new PlanDesignRecord
                    {
                        ClientName = client.ClientName,
                        ClientNumber = client.ClientNumber,
                        Carrier = "Manulife",
                        ClassName = className,
                        BenefitCategory = section.Benefit,
                        BenefitName = section.Benefit,
                        ProvisionCode = provision.Value.Code,
                        ProvisionLabel = provision.Value.Label,
                        CoverageDescription = $"{provision.Value.Label}: {field.Value}",
                        MaximumAmount = provision.Value.IsMaximum ? ExtractMoney(field.Value) : "",
                        WaitingPeriod = provision.Value.IsWaiting ? field.Value : "",
                        TerminationAge = provision.Value.IsTermination ? field.Value : "",
                        SourceFileName = Path.GetFileName(extraction.SourcePath),
                        SourceSection = $"{PlanDesignSources.PolicySchedule} | {field.SectionHeading}",
                        SourcePage = field.PageNumber,
                        Notes = $"Extractor={extraction.ExtractorVersion}; confidence={field.Confidence:0.00}; excerpt={field.SourceExcerpt}"
                    };
                    PlanDesignCanonicalMapper.CompleteFact(new PlanDesignCanonicalFact
                    {
                        BenefitCategory = row.BenefitCategory,
                        BenefitName = row.BenefitName,
                        FeatureName = row.ProvisionLabel,
                        FeatureValue = field.Value,
                        ClassName = row.ClassName,
                        SourceFileName = row.SourceFileName,
                        SourceText = field.SourceExcerpt,
                        Carrier = row.Carrier
                    });
                    result.Rows.Add(row);
                }
            }
        }

        foreach (var warning in extraction.Warnings)
            result.ValidationIssues.Add(Issue(warning.Code, "policy_schedule", warning.Message, warning.SourceExcerpt));
        return result;
    }

    private static (string Code, string Label, bool IsMaximum, bool IsWaiting, bool IsTermination)? NormalizeProvision(string label, string value, string benefit)
    {
        var key = $"{label} {value}";
        if (label.Equals("Benefit Formula", StringComparison.OrdinalIgnoreCase)) return ("benefit_formula", "Benefit Formula", false, false, false);
        if (key.Contains("non-evidence", StringComparison.OrdinalIgnoreCase)) return ("non_evidence_limit", "Non-Evidence Limit", true, false, false);
        if (key.Contains("maximum", StringComparison.OrdinalIgnoreCase)) return ("maximum_benefit", "Maximum Benefit", true, false, false);
        if (key.Contains("waiting", StringComparison.OrdinalIgnoreCase)) return ("waiting_period", "Waiting Period", false, true, false);
        if (key.Contains("qualifying", StringComparison.OrdinalIgnoreCase)) return ("qualifying_period", "Qualifying Period", false, false, false);
        if (key.Contains("elimination", StringComparison.OrdinalIgnoreCase)) return ("elimination_period", "Elimination Period", false, false, false);
        if (key.Contains("termination", StringComparison.OrdinalIgnoreCase)) return ("termination_condition", "Termination Condition", false, false, true);
        if (key.Contains("waiver", StringComparison.OrdinalIgnoreCase)) return ("waiver_of_premium", "Waiver of Premium", false, false, false);
        if (key.Contains("participation", StringComparison.OrdinalIgnoreCase)) return ("participation_basis", "Participation Basis", false, false, false);
        if (key.Contains("reduction", StringComparison.OrdinalIgnoreCase)) return ("benefit_reduction", "Benefit Reduction", false, false, false);
        if (key.Contains("cpp/qpp", StringComparison.OrdinalIgnoreCase)) return ("cpp_qpp_integration", "CPP/QPP Integration", false, false, false);
        if (key.Contains("survivor", StringComparison.OrdinalIgnoreCase)) return ("survivor_benefit", "Survivor Benefit", false, false, false);
        if (key.Contains("earnings", StringComparison.OrdinalIgnoreCase)) return ("earnings_definition", "Earnings Definition", false, false, false);
        if (key.Contains("rounding", StringComparison.OrdinalIgnoreCase)) return ("rounding_rule", "Rounding Rule", false, false, false);
        if (key.Contains("spouse", StringComparison.OrdinalIgnoreCase)) return ("spouse_benefit", "Spouse Benefit", true, false, false);
        if (key.Contains("child", StringComparison.OrdinalIgnoreCase)) return ("child_benefit", "Child Benefit", true, false, false);
        if (label.Equals("Supplementary Benefit", StringComparison.OrdinalIgnoreCase)) return ("supplementary_rider", "Supplementary Benefit", false, false, false);
        return null;
    }

    private static string ExtractMoney(string value)
    {
        var match = Regex.Match(value, @"\$[\d,]+(?:\.\d{1,2})?");
        return match.Success ? match.Value : value;
    }
    private static ValidationIssueRecord Issue(string code, string field, string message, string source) => new()
    { IssueCode = code, Severity = "Warning", TableName = "PlanDesign", FieldName = field, Message = message, RawSourceReference = source, CreatedAt = DateTime.Now };
}

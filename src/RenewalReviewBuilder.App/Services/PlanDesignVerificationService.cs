using System.Text.RegularExpressions;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

/// <summary>State-aware extension of the existing canonical Plan Design comparison pattern.</summary>
public sealed class PlanDesignVerificationService
{
    public IReadOnlyList<PlanDesignVerificationResult> CompareProfile(PlanDesignCanonicalProfile profile)
    {
        var facts = profile.CurrentPlanFacts
            .Where(fact => !string.IsNullOrWhiteSpace(fact.ComparisonKey))
            .Select(fact => new PlanDesignEvidenceFact
            {
                Fact = fact,
                PlanState = PlanDesignStates.CurrentContract,
                EffectiveDate = TryReadEffectiveDate(fact),
                SourceAuthority = SourceAuthority(fact)
            })
            .GroupBy(item => item.Fact.ComparisonKey, StringComparer.OrdinalIgnoreCase);
        var results = new List<PlanDesignVerificationResult>();
        foreach (var group in facts)
        {
            var ordered = group.OrderBy(item => item.SourceAuthority).ThenByDescending(item => item.Fact.ConfidenceScore ?? 0m).ToList();
            if (ordered.Count == 1)
            {
                results.Add(Result(ordered[0], null, PlanDesignVerificationStatuses.SourceOnlyDetail, "Only one source currently provides this provision."));
                continue;
            }
            var approved = ordered[0];
            results.AddRange(ordered.Skip(1).Select(evidence => Compare(approved, evidence)));
        }
        return results;
    }
    public PlanDesignVerificationResult Compare(PlanDesignEvidenceFact? approved, PlanDesignEvidenceFact? evidence)
    {
        if (approved is null || evidence is null)
            return Result(approved, evidence, PlanDesignVerificationStatuses.MissingFromSource, "A comparable source value is missing.");
        if (!SameIdentity(approved.Fact, evidence.Fact))
            return Result(approved, evidence, PlanDesignVerificationStatuses.UnsupportedComparison, "Benefit, class, or provision identity differs.");
        if (approved.IsManualApproved && !Equivalent(Value(approved.Fact), Value(evidence.Fact)))
            return Result(approved, evidence, PlanDesignVerificationStatuses.ManualReviewRequired, "An approved manual value is protected and requires reviewer action.");
        if (!SameState(approved.PlanState, evidence.PlanState))
            return Result(approved, evidence, PlanDesignVerificationStatuses.ProposedPlanChange, "Evidence belongs to a different plan state.");
        if (approved.EffectiveDate.HasValue && evidence.EffectiveDate.HasValue && approved.EffectiveDate != evidence.EffectiveDate)
            return Result(approved, evidence, PlanDesignVerificationStatuses.EffectiveDateDifference, "Evidence effective dates differ.");
        var left = Value(approved.Fact); var right = Value(evidence.Fact);
        if (Equivalent(left, right))
            return Result(approved, evidence, left.Equals(right, StringComparison.OrdinalIgnoreCase) ? PlanDesignVerificationStatuses.Confirmed : PlanDesignVerificationStatuses.EquivalentWording, "Values agree semantically.");
        return Result(approved, evidence, PlanDesignVerificationStatuses.MaterialConflict, "Same-state evidence differs materially.");
    }

    public static int SourceAuthority(PlanDesignCanonicalFact fact) => fact.SourceDocumentType.ToLowerInvariant() switch
    {
        "policy_amendment" => 1,
        "group_policy_schedule" => 2,
        "benefit_booklet" => 3,
        "renewal" => 5,
        "proposal" => 6,
        _ => 7
    };

    private static bool SameIdentity(PlanDesignCanonicalFact left, PlanDesignCanonicalFact right) =>
        left.CanonicalBenefitCode.Equals(right.CanonicalBenefitCode, StringComparison.OrdinalIgnoreCase) &&
        left.ClassName.Equals(right.ClassName, StringComparison.OrdinalIgnoreCase) &&
        left.FeatureName.Equals(right.FeatureName, StringComparison.OrdinalIgnoreCase);
    private static bool SameState(string left, string right) => left.Equals(right, StringComparison.OrdinalIgnoreCase) || left == PlanDesignStates.Approved && right == PlanDesignStates.CurrentContract;
    private static string Value(PlanDesignCanonicalFact fact) => string.Join(" ", fact.FeatureValue, fact.Maximum, fact.WaitingPeriod, fact.TerminationAge).Trim();
    private static bool Equivalent(string left, string right) => Normalize(left) == Normalize(right);
    private static DateOnly? TryReadEffectiveDate(PlanDesignCanonicalFact fact) =>
        DateOnly.TryParse(fact.RawSourceReference, out var effectiveDate) ? effectiveDate : null;
    private static string Normalize(string value)
    {
        var text = value.ToLowerInvariant().Replace("annual", "").Replace("participation", "").Replace("waiting period", "").Replace("whichever occurs first", "").Replace("whichever is earlier", "");
        text = Regex.Replace(text, @"\bday(s)?\b", "days");
        text = Regex.Replace(text, @"[$,\s\-.,]", "");
        text = text.Replace("earnings", "earnings").Replace("x", "x");
        return text;
    }
    private static PlanDesignVerificationResult Result(PlanDesignEvidenceFact? approved, PlanDesignEvidenceFact? evidence, string status, string message) => new()
    { Approved = approved, Evidence = evidence, Status = status, Confidence = Math.Clamp(((approved?.SourceAuthority ?? 7) + (evidence?.SourceAuthority ?? 7)) == 0 ? 1m : 0.80m, 0m, 1m), Message = message };
}

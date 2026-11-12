using System.Globalization;
using System.Text.RegularExpressions;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed record BenefitRuleMappingResult(int SourceIndex, PlanBenefitRule Rule);

public static partial class TieredReimbursementRuleNormalizer
{
    public static IReadOnlyList<BenefitRuleMappingResult> MapPlanDesignFeatures(
        IReadOnlyList<PlanDesignFeatureExtraction> features)
    {
        var candidates = features.Select((feature, index) => new SourceCandidate(
            index,
            feature.Benefit,
            feature.Category,
            feature.FeatureName,
            feature.FeatureValue,
            feature.Coinsurance,
            feature.SourceText,
            feature.SourceFileName,
            feature.SourcePage,
            FirstNonBlank(feature.ClassNumber, feature.ClassName),
            feature.ConfidenceScore)).ToList();

        return MapCandidates(candidates);
    }

    public static IReadOnlyList<BenefitRuleMappingResult> MapBenefitSummaryRows(
        IReadOnlyList<BenefitSummarySourceRow> rows)
    {
        var candidates = rows.Select((row, index) => new SourceCandidate(
            index,
            row.Benefit,
            row.Category,
            row.ProvisionName,
            row.RawValue,
            row.Coinsurance,
            row.SourceText,
            row.SourceDocument,
            row.SourcePage,
            row.ClassName,
            null)).ToList();

        return MapCandidates(candidates);
    }

    public static PlanBenefitRule? Parse(
        string sourceText,
        string sourceDocument = "",
        int? sourcePage = null,
        decimal? confidenceScore = null)
    {
        var evidence = ExtractRuleEvidence(sourceText);
        if (!LooksLikeTieredReimbursement(evidence))
        {
            return null;
        }

        var percentageMatches = PercentageRegex().Matches(evidence).ToList();
        var percentages = percentageMatches
            .Select(match => ParseDecimal(match.Groups["value"].Value))
            .Where(value => value.HasValue)
            .Select(value => value!.Value)
            .ToList();
        var (tier1Percent, tier2Percent) = ResolveTierPercentages(evidence, percentageMatches, percentages);
        var thresholdMatch = ThresholdRegex().Match(evidence);
        var threshold = thresholdMatch.Success
            ? ParseDecimal(thresholdMatch.Groups["value"].Value.Replace(",", "", StringComparison.Ordinal))
            : null;

        var rule = new PlanBenefitRule
        {
            RuleType = BenefitRuleType.TieredReimbursement,
            Tier1ReimbursementPercent = tier1Percent,
            Tier1ThresholdAmount = threshold,
            Tier2ReimbursementPercent = tier2Percent,
            ThresholdBasis = ParseBasis(evidence),
            ResetPeriod = ParsePeriod(evidence),
            CurrencyCode = evidence.Contains('$') ? "CAD" : "",
            SourceText = evidence,
            SourceLocation = BuildSourceLocation(sourceDocument, sourcePage),
            ConfidenceScore = confidenceScore,
            DisplayOrder = 10,
            IsActive = true,
            CreatedAtUtc = DateTime.UtcNow,
            UpdatedAtUtc = DateTime.UtcNow
        };
        rule.NormalizationStatus = rule.IsComplete
            ? BenefitRuleNormalizationStatus.Normalized
            : BenefitRuleNormalizationStatus.NeedsReview;
        return rule;
    }

    private static IReadOnlyList<BenefitRuleMappingResult> MapCandidates(IReadOnlyList<SourceCandidate> candidates)
    {
        var results = new List<BenefitRuleMappingResult>();
        foreach (var group in candidates.GroupBy(candidate => NormalizeGroupKey(candidate.SourceDocument)))
        {
            var resultCountBeforeGroup = results.Count;
            var parsedEvidence = group
                .Select(candidate => new
                {
                    Candidate = candidate,
                    Rule = Parse(candidate.Evidence, candidate.SourceDocument, candidate.SourcePage, candidate.ConfidenceScore)
                })
                .Where(item => item.Rule is not null)
                .OrderByDescending(item => item.Rule!.IsComplete)
                .ThenByDescending(item => item.Candidate.ConfidenceScore ?? 0m)
                .FirstOrDefault();
            var sharedRule = parsedEvidence?.Rule;

            // Never infer one tier rule by joining unrelated provisions from an entire document.
            // If no individual source row contains the rule, preserve the facts without inventing one.

            foreach (var candidate in group.Where(IsApplicableBenefit))
            {
                var directRule = Parse(candidate.Evidence, candidate.SourceDocument, candidate.SourcePage, candidate.ConfidenceScore);
                var sourceRule = HasDirectBenefitRule(candidate)
                    ? directRule
                    : HasAsterisk(candidate)
                        ? directRule ?? sharedRule
                        : null;
                if (sourceRule is null)
                {
                    continue;
                }

                results.Add(new BenefitRuleMappingResult(candidate.SourceIndex, Clone(sourceRule)));
            }

            if (sharedRule is not null && results.Count == resultCountBeforeGroup)
            {
                var reviewRule = Clone(sharedRule);
                reviewRule.NormalizationStatus = BenefitRuleNormalizationStatus.NeedsReview;
                var evidenceCandidate = parsedEvidence?.Candidate ??
                                        group.FirstOrDefault(candidate => LooksLikeTieredReimbursement(candidate.Evidence)) ??
                                        group.First();
                results.Add(new BenefitRuleMappingResult(evidenceCandidate.SourceIndex, reviewRule));
            }
        }

        return results;
    }

    private static bool IsApplicableBenefit(SourceCandidate candidate)
    {
        var identity = candidate.Identity;
        if (Regex.IsMatch(identity, @"\bhospital\b|\btravel\b|medical second opinion", RegexOptions.IgnoreCase))
        {
            return false;
        }

        return Regex.IsMatch(
            identity,
            @"prescription drugs?|health care practitioners?|paramedical|vision care|other medical expenses?|medical services\s*&\s*supplies",
            RegexOptions.IgnoreCase);
    }

    private static bool HasAsterisk(SourceCandidate candidate)
    {
        return candidate.Identity.Contains('*') ||
               candidate.FeatureValue.Contains('*') ||
               candidate.Coinsurance.Contains('*');
    }

    private static bool HasDirectBenefitRule(SourceCandidate candidate)
    {
        return LooksLikeTieredReimbursement(candidate.FeatureValue) ||
               LooksLikeTieredReimbursement(candidate.Coinsurance);
    }

    private static bool LooksLikeTieredReimbursement(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }

        var percentageCount = PercentageRegex().Matches(value).Count;
        var hasTierCue = Regex.IsMatch(
            value,
            @"\b(first|up to|threshold|thereafter|then|over)\b",
            RegexOptions.IgnoreCase);
        if (percentageCount >= 2 && hasTierCue)
        {
            return true;
        }

        return percentageCount >= 1 &&
               Regex.IsMatch(value, @"\b(first|threshold)\b", RegexOptions.IgnoreCase) &&
               Regex.IsMatch(value, @"\b(then|thereafter|over)\b", RegexOptions.IgnoreCase);
    }

    private static (decimal? Tier1, decimal? Tier2) ResolveTierPercentages(
        string evidence,
        IReadOnlyList<Match> matches,
        IReadOnlyList<decimal> percentages)
    {
        if (percentages.Count >= 2)
        {
            return (percentages[0], percentages[1]);
        }

        if (percentages.Count == 0 || matches.Count == 0)
        {
            return (null, null);
        }

        var prefix = evidence[..matches[0].Index];
        return Regex.IsMatch(prefix, @"\b(then|thereafter)\b", RegexOptions.IgnoreCase)
            ? (null, percentages[0])
            : (percentages[0], null);
    }

    private static string ExtractRuleEvidence(string value)
    {
        var text = CleanWhitespace(value);
        if (string.IsNullOrWhiteSpace(text))
        {
            return "";
        }

        var sentences = Regex.Split(text, @"(?<=[.!?])\s+");
        var matching = sentences.FirstOrDefault(LooksLikeTieredReimbursement);
        return matching ?? text;
    }

    private static BenefitRuleThresholdBasis? ParseBasis(string value)
    {
        if (Regex.IsMatch(value, @"\bper\s+certificate\b", RegexOptions.IgnoreCase))
        {
            return BenefitRuleThresholdBasis.Certificate;
        }

        if (Regex.IsMatch(value, @"\bper\s+person\b", RegexOptions.IgnoreCase))
        {
            return BenefitRuleThresholdBasis.Person;
        }

        return null;
    }

    private static BenefitRuleResetPeriod? ParsePeriod(string value)
    {
        if (Regex.IsMatch(value, @"\bcalendar\s+year\b", RegexOptions.IgnoreCase))
        {
            return BenefitRuleResetPeriod.CalendarYear;
        }

        if (Regex.IsMatch(value, @"\bpolicy\s+year\b", RegexOptions.IgnoreCase))
        {
            return BenefitRuleResetPeriod.PolicyYear;
        }

        return null;
    }

    private static PlanBenefitRule Clone(PlanBenefitRule source)
    {
        return new PlanBenefitRule
        {
            RuleType = source.RuleType,
            Tier1ReimbursementPercent = source.Tier1ReimbursementPercent,
            Tier1ThresholdAmount = source.Tier1ThresholdAmount,
            Tier2ReimbursementPercent = source.Tier2ReimbursementPercent,
            ThresholdBasis = source.ThresholdBasis,
            ResetPeriod = source.ResetPeriod,
            CurrencyCode = source.CurrencyCode,
            SourceText = source.SourceText,
            SourceLocation = source.SourceLocation,
            NormalizationStatus = source.NormalizationStatus,
            ConfidenceScore = source.ConfidenceScore,
            DisplayOrder = source.DisplayOrder,
            IsActive = source.IsActive,
            EffectiveDate = source.EffectiveDate,
            TerminationDate = source.TerminationDate,
            CreatedAtUtc = DateTime.UtcNow,
            UpdatedAtUtc = DateTime.UtcNow
        };
    }

    private static string BuildSourceLocation(string sourceDocument, int? sourcePage)
    {
        return string.Join(", ", new[]
        {
            string.IsNullOrWhiteSpace(sourceDocument) ? "" : sourceDocument.Trim(),
            sourcePage.HasValue ? $"Page {sourcePage.Value}" : ""
        }.Where(value => !string.IsNullOrWhiteSpace(value)));
    }

    private static string NormalizeGroupKey(string value)
    {
        return string.IsNullOrWhiteSpace(value) ? "__current_import__" : value.Trim().ToUpperInvariant();
    }

    private static decimal? ParseDecimal(string value)
    {
        return decimal.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out var parsed)
            ? parsed
            : null;
    }

    private static string CleanWhitespace(string value)
    {
        return Regex.Replace(value ?? "", @"\s+", " ").Trim();
    }

    private static string FirstNonBlank(params string[] values)
    {
        return values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)) ?? "";
    }

    [GeneratedRegex(@"(?<!\d)(?<value>100|[1-9]?\d(?:\.\d+)?)\s*%(?!\d)")]
    private static partial Regex PercentageRegex();

    [GeneratedRegex(@"\$\s*(?<value>\d[\d,]*(?:\.\d{1,2})?)")]
    private static partial Regex ThresholdRegex();

    private sealed record SourceCandidate(
        int SourceIndex,
        string Benefit,
        string Category,
        string FeatureName,
        string FeatureValue,
        string Coinsurance,
        string SourceText,
        string SourceDocument,
        int? SourcePage,
        string ClassKey,
        decimal? ConfidenceScore)
    {
        public string Identity => $"{Benefit} {Category} {FeatureName}";
        public string Evidence => string.Join(Environment.NewLine, new[]
        {
            FeatureValue,
            Coinsurance,
            SourceText
        }.Where(value => !string.IsNullOrWhiteSpace(value)));
    }

}

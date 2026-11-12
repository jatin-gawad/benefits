using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public static class RevisedRateNegotiationRecovery
{
    public static bool NeedsProposalRecovery(IReadOnlyCollection<NegotiationRateRow> rows)
    {
        if (rows.Count == 0)
        {
            return false;
        }

        return rows.All(row =>
        {
            var benefit = Normalize(row.BenefitLine);
            var isAggregate = benefit is "overall" or "total" or "totals" or "allbenefits";
            var lacksRateDetail = !row.CurrentRate.HasValue && !row.ProposedRate.HasValue &&
                                  !row.Volume.HasValue && !row.Lives.HasValue;
            return isAggregate && lacksRateDetail;
        });
    }

    public static List<NegotiationRateRow> BuildProposalRows(
        RevisedRateSchedule proposal,
        decimal? queueCurrentTotal,
        decimal? queueProposalTotal)
    {
        var sourceRows = proposal.Rows
            .Where(row => row.CurrentRate.HasValue || row.CurrentPremium.HasValue ||
                          row.RevisedRate.HasValue || row.RevisedPremium.HasValue)
            .ToList();
        if (sourceRows.Count == 0)
        {
            return [];
        }

        var currentScale = InferPremiumScale(sourceRows.Sum(row => row.CurrentPremium ?? 0m), queueCurrentTotal) ?? 1m;
        var proposalScale = InferPremiumScale(sourceRows.Sum(row => row.RevisedPremium ?? 0m), queueProposalTotal) ?? currentScale;

        return sourceRows.Select(source =>
        {
            var currentPremium = Scale(source.CurrentPremium, currentScale);
            var proposedPremium = Scale(source.RevisedPremium, proposalScale);
            var rateBasis = InferRateBasis(source);
            var proposedRate = source.RevisedRate;

            return new NegotiationRateRow
            {
                BenefitLine = FirstNonBlank(source.CanonicalBenefit, source.SourceBenefit, "Overall"),
                ClassName = FirstNonBlank(source.TierOrClass, source.SourceClass, "All"),
                RateBasis = rateBasis,
                Volume = source.Volume,
                Lives = source.Lives,
                CurrentRate = source.CurrentRate,
                ProposedRate = proposedRate,
                ModelledRate = proposedRate,
                NegotiatedRate = proposedRate,
                CurrentPremium = currentPremium,
                ProposedPremium = proposedPremium,
                ModelledPremium = proposedPremium,
                NegotiatedPremium = proposedPremium,
                RateValueBasis = RateValueBasisNames.CarrierStated,
                ClientPresentationRateBasis = RateValueBasisNames.CarrierStated,
                SourceReference = source.SourceReference,
                UseSavedModelValues = true,
                Status = RenewalQueueStatuses.Negotiated,
                Notes = "Detailed carrier proposal rows recovered from the original renewal source."
            };
        }).ToList();
    }

    public static decimal? ScaleImportedPremium(RevisedRateScheduleRow source, NegotiationRateRow target)
    {
        if (!source.RevisedPremium.HasValue)
        {
            return null;
        }

        var scale = InferPremiumScale(source.CurrentPremium, target.CurrentPremium)
                    ?? InferPremiumScale(source.RevisedPremium, target.ProposedPremium)
                    ?? 1m;
        return Scale(source.RevisedPremium, scale);
    }

    private static decimal? InferPremiumScale(decimal? sourcePremium, decimal? targetPremium)
    {
        if (sourcePremium is not > 0m || targetPremium is not > 0m)
        {
            return null;
        }

        var ratio = targetPremium.Value / sourcePremium.Value;
        if (ratio is > 10m and < 14m)
        {
            return 12m;
        }

        return ratio is > 0.75m and < 1.25m ? 1m : null;
    }

    private static decimal? Scale(decimal? value, decimal scale)
    {
        return value.HasValue ? decimal.Round(value.Value * scale, 2) : null;
    }

    private static string InferRateBasis(RevisedRateScheduleRow row)
    {
        var normalized = Normalize($"{row.CanonicalBenefit} {row.SourceBenefit} {row.TierOrClass} {row.SourceClass}");
        if (normalized.Contains("single", StringComparison.Ordinal) ||
            normalized.Contains("employeeonly", StringComparison.Ordinal))
        {
            return "single";
        }

        if (normalized.Contains("family", StringComparison.Ordinal) ||
            normalized.Contains("couple", StringComparison.Ordinal) ||
            normalized.Contains("spouse", StringComparison.Ordinal) ||
            normalized.Contains("dependent", StringComparison.Ordinal))
        {
            return "family";
        }

        if (row.Volume.HasValue || normalized.Contains("life", StringComparison.Ordinal) ||
            normalized.Contains("adandd", StringComparison.Ordinal) ||
            normalized.Contains("std", StringComparison.Ordinal) ||
            normalized.Contains("ltd", StringComparison.Ordinal))
        {
            return "volume";
        }

        return "composite";
    }

    private static string Normalize(string value)
    {
        return new string((value ?? "").Where(char.IsLetterOrDigit).Select(char.ToLowerInvariant).ToArray());
    }

    private static string FirstNonBlank(params string[] values)
    {
        return values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value))?.Trim() ?? "";
    }
}

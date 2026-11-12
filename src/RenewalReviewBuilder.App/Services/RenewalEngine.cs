using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed class RenewalEngine
{
    private readonly UnderwritingEngine _underwritingEngine;
    private readonly NegotiationWorkbenchService _negotiationWorkbench;
    private readonly QuoteEngine _quoteEngine;

    public RenewalEngine(UnderwritingEngine? underwritingEngine = null)
    {
        _underwritingEngine = underwritingEngine ?? new UnderwritingEngine();
        _negotiationWorkbench = new NegotiationWorkbenchService(_underwritingEngine);
        _quoteEngine = new QuoteEngine(_underwritingEngine);
    }

    public void BuildRenewalVersions(AnalysisResult analysis)
    {
        analysis.RenewalVersions.Clear();

        var carrierProposed = BuildVersion(
            analysis,
            RenewalVersionNames.CarrierProposed,
            BenefitsIntelligenceModules.RenewalIntelligence,
            "Original carrier proposal",
            line => line.RenewalRate,
            line => line.RenewalPremium,
            useRequiredPremium: false);
        carrierProposed.IsSelected = true;

        var consultantModel = BuildVersion(
            analysis,
            RenewalVersionNames.ConsultantModel,
            BenefitsIntelligenceModules.UnderwritingIntelligence,
            "Shared underwriting model",
            line => line.RenewalRate,
            line => line.CalculatedRenewalPremium ?? line.RenewalPremium,
            useRequiredPremium: true);

        var negotiated = _negotiationWorkbench.BuildNegotiatedVersion(analysis, consultantModel);
        var finalApproved = CloneVersion(negotiated, RenewalVersionNames.FinalApproved, "Awaiting final approval");
        var alternateOption = _quoteEngine.BuildAlternateOption(analysis, consultantModel);

        analysis.RenewalVersions.Add(carrierProposed);
        analysis.RenewalVersions.Add(consultantModel);
        analysis.RenewalVersions.Add(negotiated);
        analysis.RenewalVersions.Add(finalApproved);
        analysis.RenewalVersions.Add(alternateOption);
        analysis.SelectedRenewalVersionName = RenewalVersionNames.CarrierProposed;
    }

    private RenewalVersion BuildVersion(
        AnalysisResult analysis,
        string versionName,
        string sourceModule,
        string status,
        Func<RateLine, decimal?> rateSelector,
        Func<RateLine, decimal?> premiumSelector,
        bool useRequiredPremium)
    {
        var version = new RenewalVersion
        {
            VersionName = versionName,
            SourceModule = sourceModule,
            Status = status
        };

        var sourceLines = analysis.Pdf.RateLines.Count > 0
            ? analysis.Pdf.RateLines
            : BuildOverallFallbackRateLines(analysis);

        foreach (var line in sourceLines)
        {
            var versionRate = rateSelector(line);
            var versionPremium = premiumSelector(line);
            var calculation = CalculateLine(analysis, line, versionName, sourceModule, versionPremium, versionRate, null);
            if (useRequiredPremium && calculation.RequiredPremium is > 0m)
            {
                versionPremium = calculation.RequiredPremium;
                versionRate = calculation.RequiredRate ?? versionRate;
            }

            version.Calculations.Add(calculation);
            var versionLine = CreateVersionRateLine(line, versionRate, versionPremium ?? calculation.ProposedPremium);
            versionLine.CurrentPremium ??= calculation.CurrentPremium;
            version.RateLines.Add(versionLine);
        }

        return version;
    }

    internal UnderwritingCalculationResult CalculateLine(
        AnalysisResult analysis,
        RateLine line,
        string versionName,
        string sourceModule,
        decimal? proposedPremium,
        decimal? proposedRate,
        decimal? negotiatedPremium)
    {
        var experience = FindMatchingExperience(analysis.Pdf.ExperienceLines, line.Benefit);
        var currentPremium = line.CurrentPremium ?? line.CalculatedCurrentPremium;
        var currentClaims = ResolveClaims(experience, currentPremium);
        var assumptions = DefaultAssumptions(line.Benefit);

        return _underwritingEngine.Calculate(new UnderwritingCalculationInput
        {
            SourceModule = sourceModule,
            VersionName = versionName,
            BenefitLine = line.Benefit,
            CurrentClaims = currentClaims,
            CurrentPremium = currentPremium,
            ProposedPremium = proposedPremium,
            NegotiatedPremium = negotiatedPremium,
            CurrentRate = line.CurrentRate,
            ProposedRate = proposedRate,
            NegotiatedRate = null,
            Volume = line.Volume,
            Lives = line.Lives,
            RateBasis = line.RateBasis,
            Trend = assumptions.Trend,
            TargetLossRatio = assumptions.TargetLossRatio,
            PoolingAdjustment = assumptions.PoolingAdjustment,
            IbnrAdjustment = assumptions.IbnrAdjustment,
            RetentionAdminFee = assumptions.RetentionAdminFee,
            ManualAdjustment = assumptions.ManualAdjustment
        });
    }

    internal RenewalVersionRateLine CreateVersionRateLine(RateLine line, decimal? versionRate, decimal? versionPremium)
    {
        return new RenewalVersionRateLine
        {
            Benefit = string.IsNullOrWhiteSpace(line.Benefit) ? "Overall" : line.Benefit,
            TierOrClass = line.TierOrClass,
            AccountClass = line.AccountClass,
            Volume = line.Volume,
            Lives = line.Lives,
            CurrentRate = line.CurrentRate,
            CurrentPremium = line.CurrentPremium ?? line.CalculatedCurrentPremium,
            ProposedRate = line.RenewalRate,
            ProposedPremium = line.RenewalPremium,
            VersionRate = versionRate,
            VersionPremium = versionPremium,
            ChangePercent = UnderwritingEngine.CalculatePercentChange(line.CurrentPremium ?? line.CalculatedCurrentPremium, versionPremium),
            RateBasis = string.IsNullOrWhiteSpace(line.RateBasis) ? InferRateBasis(line) : line.RateBasis,
            FundingArrangement = line.FundingArrangement,
            SourceReference = line.SourceReference,
            IncludeInPremiumTotal = line.IncludeInPremiumTotal,
            SourceRowRole = line.SourceRowRole,
            ValidationStatus = line.ValidationStatus,
            ValidationMessage = line.ValidationMessage,
            SourceCarrier = line.SourceCarrier,
            SourceFileName = line.SourceFileName,
            SourcePage = line.SourcePage,
            SourceText = line.SourceText,
            CurrentRatePrecision = line.CurrentRatePrecision,
            RenewalRatePrecision = line.RenewalRatePrecision,
            CurrentPeriodLabel = line.CurrentPeriodLabel,
            RenewalPeriodLabel = line.RenewalPeriodLabel,
            RenewalEffectiveDate = line.RenewalEffectiveDate
        };
    }

    private RenewalVersion CloneVersion(RenewalVersion source, string versionName, string status)
    {
        var clone = new RenewalVersion
        {
            VersionName = versionName,
            SourceModule = source.SourceModule,
            Status = status
        };

        foreach (var line in source.RateLines)
        {
            clone.RateLines.Add(new RenewalVersionRateLine
            {
                Benefit = line.Benefit,
                TierOrClass = line.TierOrClass,
                AccountClass = line.AccountClass,
                Volume = line.Volume,
                Lives = line.Lives,
                CurrentRate = line.CurrentRate,
                CurrentPremium = line.CurrentPremium,
                ProposedRate = line.ProposedRate,
                ProposedPremium = line.ProposedPremium,
                VersionRate = line.VersionRate,
                VersionPremium = line.VersionPremium,
                ChangePercent = line.ChangePercent,
                RateBasis = line.RateBasis,
                RateValueBasis = line.RateValueBasis,
                ClientPresentationRateBasis = line.ClientPresentationRateBasis,
                CurrentNetRate = line.CurrentNetRate,
                CurrentGrossRate = line.CurrentGrossRate,
                ProposedNetRate = line.ProposedNetRate,
                ProposedGrossRate = line.ProposedGrossRate,
                VersionNetRate = line.VersionNetRate,
                VersionGrossRate = line.VersionGrossRate,
                RateScheduleSourceFile = line.RateScheduleSourceFile,
                RateScheduleSourceSheet = line.RateScheduleSourceSheet,
                RateScheduleSourceRow = line.RateScheduleSourceRow,
                FundingArrangement = line.FundingArrangement,
                SourceReference = line.SourceReference,
                IncludeInPremiumTotal = line.IncludeInPremiumTotal,
                SourceRowRole = line.SourceRowRole,
                ValidationStatus = line.ValidationStatus,
                ValidationMessage = line.ValidationMessage,
                SourceCarrier = line.SourceCarrier,
                SourceFileName = line.SourceFileName,
                SourcePage = line.SourcePage,
                SourceText = line.SourceText,
                CurrentRatePrecision = line.CurrentRatePrecision,
                RenewalRatePrecision = line.RenewalRatePrecision,
                CurrentPeriodLabel = line.CurrentPeriodLabel,
                RenewalPeriodLabel = line.RenewalPeriodLabel,
                RenewalEffectiveDate = line.RenewalEffectiveDate
            });
        }

        foreach (var calculation in source.Calculations)
        {
            clone.Calculations.Add(CloneCalculation(calculation, versionName));
        }

        return clone;
    }

    private static UnderwritingCalculationResult CloneCalculation(UnderwritingCalculationResult source, string versionName)
    {
        return new UnderwritingCalculationResult
        {
            SourceModule = source.SourceModule,
            VersionName = versionName,
            BenefitLine = source.BenefitLine,
            CurrentClaims = source.CurrentClaims,
            CurrentPremium = source.CurrentPremium,
            TrendedClaims = source.TrendedClaims,
            PoolingAdjustment = source.PoolingAdjustment,
            IbnrAdjustment = source.IbnrAdjustment,
            RetentionAdminFeeAdjustment = source.RetentionAdminFeeAdjustment,
            ManualAdjustment = source.ManualAdjustment,
            ProjectedClaims = source.ProjectedClaims,
            RequiredPremium = source.RequiredPremium,
            ProposedPremium = source.ProposedPremium,
            NegotiatedPremium = source.NegotiatedPremium,
            RequiredRate = source.RequiredRate,
            TargetLossRatio = source.TargetLossRatio,
            ProjectedLossRatio = source.ProjectedLossRatio,
            RenewalChangePercent = source.RenewalChangePercent,
            NegotiatedChangePercent = source.NegotiatedChangePercent,
            Savings = source.Savings,
            Trend = source.Trend
        };
    }

    private static List<RateLine> BuildOverallFallbackRateLines(AnalysisResult analysis)
    {
        return
        [
            new RateLine
            {
                Benefit = "Overall",
                CurrentPremium = analysis.Pdf.CurrentMonthlyPremium,
                RenewalPremium = analysis.Pdf.RenewalMonthlyPremium,
                PercentChange = analysis.Pdf.MonthlyPremiumChangePercent,
                SourceReference = "PDF renewal rate illustration"
            }
        ];
    }

    private static ExperienceLine? FindMatchingExperience(IEnumerable<ExperienceLine> experienceLines, string benefit)
    {
        return ExperienceLineSelector.FindBestMatch(experienceLines, benefit);
    }

    private static decimal? ResolveClaims(ExperienceLine? experience, decimal? currentPremium)
    {
        if (experience is null)
        {
            return null;
        }

        var lossRatio = UnderwritingEngine.NormalizeRate(experience.PaidLossRatio ?? experience.IncurredLossRatio);
        if (currentPremium.HasValue && lossRatio.HasValue)
        {
            return decimal.Round(currentPremium.Value * lossRatio.Value, 2);
        }

        if (experience.PaidClaims.HasValue || experience.IncurredClaims.HasValue)
        {
            return experience.PaidClaims ?? experience.IncurredClaims;
        }

        return null;
    }

    private static UnderwritingCalculationInput DefaultAssumptions(string benefit)
    {
        if (BenefitMatcher.ContainsPhrase(benefit, "dental"))
        {
            return new UnderwritingCalculationInput
            {
                Trend = 0.06m,
                TargetLossRatio = 0.763m
            };
        }

        if (BenefitMatcher.ContainsPhrase(benefit, "health") ||
            BenefitMatcher.ContainsPhrase(benefit, "ehc"))
        {
            return new UnderwritingCalculationInput
            {
                Trend = 0.08m,
                TargetLossRatio = 0.763m
            };
        }

        return new UnderwritingCalculationInput
        {
            TargetLossRatio = 0.85m
        };
    }

    private static string InferRateBasis(RateLine rateLine)
    {
        var normalized = NormalizeForMatch($"{rateLine.Benefit} {rateLine.TierOrClass} {rateLine.AccountClass}");
        if (normalized.Contains("single", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("employeeonly", StringComparison.OrdinalIgnoreCase))
        {
            return "single";
        }

        if (normalized.Contains("family", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("couple", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("spouse", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("dependent", StringComparison.OrdinalIgnoreCase))
        {
            return "family";
        }

        if (rateLine.Volume.HasValue ||
            normalized.Contains("volume", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("life", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("adandd", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("std", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("ltd", StringComparison.OrdinalIgnoreCase))
        {
            return "volume";
        }

        return "composite";
    }

    private static string NormalizeForMatch(string value)
    {
        return new string((value ?? "")
            .Where(char.IsLetterOrDigit)
            .Select(char.ToLowerInvariant)
            .ToArray());
    }
}

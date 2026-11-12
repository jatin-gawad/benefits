using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed class NegotiationWorkbenchService
{
    private readonly UnderwritingEngine _underwritingEngine;

    public NegotiationWorkbenchService(UnderwritingEngine? underwritingEngine = null)
    {
        _underwritingEngine = underwritingEngine ?? new UnderwritingEngine();
    }

    public RenewalVersion BuildNegotiatedVersion(AnalysisResult analysis, RenewalVersion baseline)
    {
        var version = new RenewalVersion
        {
            VersionName = RenewalVersionNames.Negotiated,
            SourceModule = BenefitsIntelligenceModules.RenewalIntelligence,
            Status = "Negotiation workbench scenario"
        };

        foreach (var line in baseline.RateLines)
        {
            var input = new UnderwritingCalculationInput
            {
                SourceModule = BenefitsIntelligenceModules.RenewalIntelligence,
                VersionName = RenewalVersionNames.Negotiated,
                BenefitLine = line.Benefit,
                CurrentPremium = line.CurrentPremium,
                ProposedPremium = line.ProposedPremium,
                NegotiatedPremium = line.VersionPremium,
                CurrentRate = line.CurrentRate,
                ProposedRate = line.ProposedRate,
                NegotiatedRate = line.VersionRate,
                Volume = line.Volume,
                Lives = line.Lives,
                RateBasis = line.RateBasis
            };

            var calculation = _underwritingEngine.Calculate(input);
            version.Calculations.Add(calculation);
            version.RateLines.Add(new RenewalVersionRateLine
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
                ChangePercent = calculation.NegotiatedChangePercent ?? calculation.RenewalChangePercent,
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
                SourceReference = line.SourceReference
            });
        }

        return version;
    }

    public RenewalVersion ApplySavedVersion(
        AnalysisResult analysis,
        NegotiationWorkbenchSaveRequest request,
        string versionName)
    {
        ArgumentNullException.ThrowIfNull(analysis);
        ArgumentNullException.ThrowIfNull(request);

        if (versionName is not RenewalVersionNames.Negotiated and not RenewalVersionNames.FinalApproved)
        {
            throw new ArgumentOutOfRangeException(nameof(versionName), versionName, "Only negotiated or final approved versions can be applied from a saved negotiation.");
        }

        var version = analysis.RenewalVersions.FirstOrDefault(item =>
            item.VersionName.Equals(versionName, StringComparison.OrdinalIgnoreCase));
        if (version is null)
        {
            version = new RenewalVersion { VersionName = versionName };
            analysis.RenewalVersions.Add(version);
        }

        version.SourceModule = BenefitsIntelligenceModules.RenewalIntelligence;
        version.Status = request.Status;
        version.CreatedAt = DateTime.Now;
        version.RequestedIncreasePercent = request.RequestedIncreasePercent;
        version.NegotiatedIncreasePercent = request.NegotiatedIncreasePercent;
        version.FinalAcceptedIncreasePercent = request.FinalAcceptedIncreasePercent;
        version.NegotiationStatus = request.NegotiationStatus;
        version.NegotiationStage = request.NegotiationStage;
        version.NegotiationCompletedDate = request.NegotiationCompletedDate;
        version.NegotiatedSavingsAnnual = request.NegotiatedSavingsAnnual;
        version.NegotiatedSavingsMonthly = request.NegotiatedSavingsMonthly;
        version.NegotiatedSavingsType = request.NegotiatedSavingsType;
        version.NegotiatedBy = request.NegotiatedBy;
        version.IndustryBenchmarkPercent = request.IndustryBenchmarkPercent;
        version.IndustryBenchmarkSource = request.IndustryBenchmarkSource;
        version.IndustryBenchmarkPopulation = request.IndustryBenchmarkPopulation;
        version.IndustryBenchmarkYear = request.IndustryBenchmarkYear;
        version.IndustryBenchmarkConfidence = request.IndustryBenchmarkConfidence;
        version.RenewalRecommendationType = request.RenewalRecommendationType;
        version.RenewalDecisionReason = request.RenewalDecisionReason;
        version.RenewalStrategySummary = request.RenewalStrategySummary;
        version.RenewalRiskLevel = request.RenewalRiskLevel;
        version.RenewalComplexity = request.RenewalComplexity;
        version.RenewalResult = request.RenewalResult;
        version.RenewalCompletedDate = request.RenewalCompletedDate;
        version.RenewalApprovedBy = request.RenewalApprovedBy;
        version.RenewalAcceptedDate = request.RenewalAcceptedDate;
        version.RenewalVersionNumber = request.RenewalVersionNumber;
        version.RenewalScenarioId = request.RenewalScenarioId;
        version.PreviousRenewalScenarioId = request.PreviousRenewalScenarioId;
        version.ScenarioType = request.ScenarioType;
        version.HistoryEventId = request.HistoryEventId;
        version.RelationshipMilestoneId = request.RelationshipMilestoneId;
        version.EvolutionGroupId = request.EvolutionGroupId;
        version.RateLines.Clear();
        version.Calculations.Clear();

        foreach (var rate in request.Rates)
        {
            var versionPremium = versionName == RenewalVersionNames.FinalApproved
                ? rate.NegotiatedPremium ?? rate.ModelledPremium ?? rate.ProposedPremium
                : rate.NegotiatedPremium ?? rate.ProposedPremium;
            var hasNegotiatedRate = rate.NegotiatedRate.HasValue;
            var versionRate = !hasNegotiatedRate
                ? null
                : versionName == RenewalVersionNames.FinalApproved
                    ? rate.NegotiatedRate ?? rate.ModelledRate ?? rate.ProposedRate
                    : rate.NegotiatedRate ?? rate.ProposedRate;

            version.RateLines.Add(new RenewalVersionRateLine
            {
                Benefit = rate.BenefitLine,
                TierOrClass = rate.ClassName,
                AccountClass = rate.AccountClass,
                Volume = rate.Volume,
                Lives = rate.Lives,
                CurrentRate = rate.CurrentRate,
                CurrentPremium = rate.CurrentPremium,
                ProposedRate = rate.ProposedRate,
                ProposedPremium = rate.ProposedPremium,
                VersionRate = versionRate,
                VersionPremium = versionPremium,
                ChangePercent = UnderwritingEngine.CalculatePercentChange(rate.CurrentPremium, versionPremium),
                RateBasis = rate.RateBasis,
                RateValueBasis = rate.RateValueBasis,
                ClientPresentationRateBasis = rate.ClientPresentationRateBasis,
                CurrentNetRate = rate.CurrentNetRate,
                CurrentGrossRate = rate.CurrentGrossRate,
                ProposedNetRate = rate.ProposedNetRate,
                ProposedGrossRate = rate.ProposedGrossRate,
                VersionNetRate = rate.VersionNetRate,
                VersionGrossRate = rate.ClientPresentationRateBasis.Equals(RateValueBasisNames.Gross, StringComparison.OrdinalIgnoreCase)
                    ? versionRate
                    : rate.VersionGrossRate,
                RateScheduleSourceFile = rate.RateScheduleSourceFile,
                RateScheduleSourceSheet = rate.RateScheduleSourceSheet,
                RateScheduleSourceRow = rate.RateScheduleSourceRow,
                SourceReference = string.IsNullOrWhiteSpace(rate.Source)
                    ? "Renewal Negotiation Workbench"
                    : rate.Source
            });
        }

        foreach (var calculation in request.Calculations)
        {
            version.Calculations.Add(new UnderwritingCalculationResult
            {
                SourceModule = calculation.SourceModule,
                VersionName = versionName,
                BenefitLine = calculation.BenefitLine,
                CurrentClaims = calculation.CurrentClaims,
                CurrentPremium = calculation.CurrentPremium,
                TrendedClaims = calculation.TrendedClaims,
                PoolingAdjustment = calculation.PoolingAdjustment,
                IbnrAdjustment = calculation.IbnrAdjustment,
                RetentionAdminFeeAdjustment = calculation.RetentionAdminFeeAdjustment,
                ManualAdjustment = calculation.ManualAdjustment,
                ProjectedClaims = calculation.ProjectedClaims,
                RequiredPremium = calculation.RequiredPremium,
                ProposedPremium = calculation.ProposedPremium,
                NegotiatedPremium = calculation.NegotiatedPremium,
                RequiredRate = calculation.RequiredRate,
                TargetLossRatio = calculation.TargetLossRatio,
                ProjectedLossRatio = calculation.ProjectedLossRatio,
                RenewalChangePercent = calculation.RenewalChangePercent,
                NegotiatedChangePercent = calculation.NegotiatedChangePercent,
                Savings = calculation.Savings,
                Trend = calculation.Trend
            });
        }

        analysis.SelectedRenewalVersionName = versionName;
        return version;
    }
}

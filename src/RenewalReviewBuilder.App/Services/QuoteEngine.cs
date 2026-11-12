using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed class QuoteEngine
{
    private readonly UnderwritingEngine _underwritingEngine;

    public QuoteEngine(UnderwritingEngine? underwritingEngine = null)
    {
        _underwritingEngine = underwritingEngine ?? new UnderwritingEngine();
    }

    public RenewalVersion BuildAlternateOption(AnalysisResult analysis, RenewalVersion baseline)
    {
        var version = new RenewalVersion
        {
            VersionName = RenewalVersionNames.AlternateOption,
            SourceModule = BenefitsIntelligenceModules.ProposalComparisonIntelligence,
            Status = "Alternate carrier or plan design option"
        };

        foreach (var line in baseline.RateLines)
        {
            var input = new UnderwritingCalculationInput
            {
                SourceModule = BenefitsIntelligenceModules.ProposalComparisonIntelligence,
                VersionName = RenewalVersionNames.AlternateOption,
                BenefitLine = line.Benefit,
                CurrentPremium = line.CurrentPremium,
                ProposedPremium = line.VersionPremium ?? line.ProposedPremium,
                CurrentRate = line.CurrentRate,
                ProposedRate = line.VersionRate ?? line.ProposedRate,
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
                ChangePercent = calculation.RenewalChangePercent,
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
}

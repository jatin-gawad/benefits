using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services.Normalization;

public static class NormalizeRenewalData
{
    public static NormalizedImportPayload FromAnalysis(AnalysisResult analysis, NormalizedSourceMetadata source)
    {
        var payload = FromPdfExtraction(analysis.Pdf, source);
        payload.DocumentType = DocumentTypes.Renewal;
        payload.Policy ??= new NormalizedPolicy();
        payload.Policy.ClientName = FirstNonBlank(payload.Policy.ClientName, analysis.RenewalCase.ClientName, analysis.Pdf.ClientName);
        payload.Policy.PolicyNumber = FirstNonBlank(payload.Policy.PolicyNumber, analysis.RenewalCase.PolicyNumber, analysis.Pdf.PolicyNumber);
        payload.Policy.RenewalDate = analysis.RenewalCase.RenewalDate ?? analysis.Pdf.RenewalDate;

        foreach (var version in analysis.RenewalVersions)
        {
            payload.RenewalVersions.Add(new NormalizedRenewalVersion
            {
                VersionName = version.VersionName,
                Status = version.Status,
                EffectiveDate = analysis.RenewalCase.RenewalDate ?? analysis.Pdf.RenewalDate,
                RequestedIncreasePercent = version.RequestedIncreasePercent,
                NegotiatedIncreasePercent = version.NegotiatedIncreasePercent,
                FinalAcceptedIncreasePercent = version.FinalAcceptedIncreasePercent,
                NegotiationStatus = version.NegotiationStatus,
                NegotiationStage = version.NegotiationStage,
                NegotiationCompletedDate = version.NegotiationCompletedDate,
                NegotiatedSavingsAnnual = version.NegotiatedSavingsAnnual,
                NegotiatedSavingsMonthly = version.NegotiatedSavingsMonthly,
                NegotiatedSavingsType = version.NegotiatedSavingsType,
                NegotiatedBy = version.NegotiatedBy,
                IndustryBenchmarkPercent = version.IndustryBenchmarkPercent,
                IndustryBenchmarkSource = version.IndustryBenchmarkSource,
                IndustryBenchmarkPopulation = version.IndustryBenchmarkPopulation,
                IndustryBenchmarkYear = version.IndustryBenchmarkYear,
                IndustryBenchmarkConfidence = version.IndustryBenchmarkConfidence,
                RenewalRecommendationType = version.RenewalRecommendationType,
                RenewalDecisionReason = version.RenewalDecisionReason,
                RenewalStrategySummary = version.RenewalStrategySummary,
                RenewalRiskLevel = version.RenewalRiskLevel,
                RenewalComplexity = version.RenewalComplexity,
                RenewalResult = version.RenewalResult,
                RenewalCompletedDate = version.RenewalCompletedDate,
                RenewalApprovedBy = version.RenewalApprovedBy,
                RenewalAcceptedDate = version.RenewalAcceptedDate,
                RenewalVersionNumber = version.RenewalVersionNumber,
                RenewalScenarioId = version.RenewalScenarioId,
                PreviousRenewalScenarioId = version.PreviousRenewalScenarioId,
                ScenarioType = version.ScenarioType,
                HistoryEventId = version.HistoryEventId,
                RelationshipMilestoneId = version.RelationshipMilestoneId,
                EvolutionGroupId = version.EvolutionGroupId
            });
        }

        payload.ApplySourceToRecords();
        return payload;
    }

    public static NormalizedImportPayload FromPdfExtraction(PdfExtractionResult pdf, NormalizedSourceMetadata source)
    {
        var carrier = FirstNonBlank(pdf.Carrier, source.SourceCarrier);
        var payload = new NormalizedImportPayload
        {
            DocumentType = DocumentTypes.Renewal,
            Source = source.Copy(),
            Client = new NormalizedClient
            {
                ClientName = pdf.ClientName,
                ClientNumber = pdf.PolicyNumber
            },
            Carrier = new NormalizedCarrier
            {
                CarrierName = carrier
            },
            Policy = new NormalizedPolicy
            {
                PolicyNumber = pdf.PolicyNumber,
                ClientName = pdf.ClientName,
                CarrierName = carrier,
                RenewalDate = pdf.RenewalDate,
                EffectiveDate = pdf.RenewalDate
            }
        };

        foreach (var line in pdf.RateLines)
        {
            payload.Rates.Add(new NormalizedRate
            {
                BenefitLine = FirstNonBlank(line.Benefit, "Overall"),
                ClassName = FirstNonBlank(line.TierOrClass, line.AccountClass),
                CurrentRate = line.CurrentRate,
                ProposedRate = line.RenewalRate,
                EffectiveDate = pdf.RenewalDate
            });

            payload.Premiums.Add(new NormalizedPremium
            {
                BenefitLine = FirstNonBlank(line.Benefit, "Overall"),
                ClassName = FirstNonBlank(line.TierOrClass, line.AccountClass),
                PeriodLabel = "Current",
                Amount = line.CurrentPremium
            });

            payload.Premiums.Add(new NormalizedPremium
            {
                BenefitLine = FirstNonBlank(line.Benefit, "Overall"),
                ClassName = FirstNonBlank(line.TierOrClass, line.AccountClass),
                PeriodLabel = "Renewal",
                Amount = line.RenewalPremium
            });
        }

        foreach (var line in pdf.ExperienceLines)
        {
            payload.ClaimsExperience.Add(new NormalizedClaimsExperience
            {
                BenefitLine = FirstNonBlank(line.Benefit, "Overall"),
                Premium = line.BilledPremium,
                PaidClaims = line.PaidClaims,
                IncurredClaims = line.IncurredClaims,
                LossRatio = line.PaidLossRatio ?? line.IncurredLossRatio
            });
        }

        payload.Renewals.Add(new NormalizedRenewal
        {
            CarrierName = carrier,
            BenefitLine = "Overall",
            EffectiveDate = pdf.RenewalDate,
            RenewalYear = pdf.RenewalDate?.Year,
            CurrentPremium = pdf.CurrentMonthlyPremium,
            RenewalPremium = pdf.RenewalMonthlyPremium,
            Claims = pdf.ExperienceLines.Sum(line => line.PaidClaims ?? line.IncurredClaims ?? 0m),
            Lives = pdf.RateLines.Select(line => line.Lives).FirstOrDefault(value => value.HasValue)
        });

        if (pdf.PoolingCharge.HasValue || pdf.PoolingThreshold.HasValue)
        {
            payload.Pooling.Add(new NormalizedPooling
            {
                BenefitLine = "All",
                PoolingCharge = pdf.PoolingCharge,
                PoolingThreshold = pdf.PoolingThreshold
            });
        }

        payload.ApplySourceToRecords();
        return payload;
    }

    private static string FirstNonBlank(params string?[] values)
    {
        return values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)) ?? "";
    }
}

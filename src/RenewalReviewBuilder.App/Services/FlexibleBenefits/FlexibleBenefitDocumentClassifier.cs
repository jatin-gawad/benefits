using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services.FlexibleBenefits;

public sealed class FlexibleBenefitDocumentClassifier
{
    private static readonly (string Carrier, string[] Signals)[] CarrierSignals =
    [
        ("National HealthClaim", ["nationalhealthclaim", "lifehealth+"]),
        ("Manulife", ["manulife"]),
        ("Canada Life", ["canadalife"]),
        ("Sun Life", ["sunlife"]),
        ("GreenShield", ["greenshield", "green shield"])
    ];

    public FlexibleBenefitDocumentClassification Classify(FlexibleBenefitOcrDocument document)
    {
        var normalized = FlexibleBenefitNormalizationService.NormalizeSearchText(document.CombinedText);
        var carrier = CarrierSignals
            .Select(signal => new
            {
                signal.Carrier,
                Count = signal.Signals.Count(value => normalized.Contains(
                    FlexibleBenefitNormalizationService.NormalizeSearchText(value),
                    StringComparison.Ordinal))
            })
            .OrderByDescending(result => result.Count)
            .FirstOrDefault(result => result.Count > 0);

        var hasEnrollment = normalized.Contains("enrollmentform", StringComparison.Ordinal);
        var hasAgreement = normalized.Contains("administrativeservicesindemnityagreement", StringComparison.Ordinal) ||
                           normalized.Contains("indemnitycontract", StringComparison.Ordinal);
        var subtype = hasEnrollment && hasAgreement
            ? FlexibleBenefitDocumentSubtype.CombinedEnrollmentAndAgreement
            : hasEnrollment
                ? FlexibleBenefitDocumentSubtype.EnrollmentForm
                : hasAgreement
                    ? FlexibleBenefitDocumentSubtype.IndemnityAgreement
                    : normalized.Contains("feeschedule", StringComparison.Ordinal) || normalized.Contains("feesandterms", StringComparison.Ordinal)
                        ? FlexibleBenefitDocumentSubtype.FeeSchedule
                        : normalized.Contains("amendment", StringComparison.Ordinal)
                            ? FlexibleBenefitDocumentSubtype.PlanAmendment
                            : FlexibleBenefitDocumentSubtype.Unknown;
        var hasElectronicSignature = normalized.Contains("esignedby", StringComparison.Ordinal);

        return new FlexibleBenefitDocumentClassification
        {
            Carrier = carrier?.Carrier ?? "",
            Administrator = carrier?.Carrier ?? "",
            DocumentType = FlexibleBenefitDocumentTypes.FlexibleBenefitsPlanDocument,
            DocumentSubtype = subtype,
            IsSigned = hasElectronicSignature ? true : null,
            IsAmendment = subtype == FlexibleBenefitDocumentSubtype.PlanAmendment,
            Confidence = carrier is null || subtype == FlexibleBenefitDocumentSubtype.Unknown ? 0.45m : 0.94m
        };
    }
}

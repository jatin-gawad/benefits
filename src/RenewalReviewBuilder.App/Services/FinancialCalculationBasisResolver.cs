using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed class FinancialCalculationBasisResolver
{
    public ResolvedFinancialCalculationBasis Resolve(string basis, FinancialCalculationInputSet inputs)
    {
        var requestedCanonicalBasis = FinancialCalculationBasisCatalog.Canonicalize(basis);
        var requestedBasis = FinancialCalculationBasisCatalog.Normalize(basis);
        var sources = new[]
        {
            inputs.FinalizedRenewalProjection,
            inputs.ConfirmedRenewalFinancialAnalysis,
            inputs.ValidatedClaimsProjection,
            inputs.ValidatedAsoActivity,
            inputs.ActiveClientFinancialData
        };
        foreach (var source in sources)
        {
            var match = source.FirstOrDefault(input =>
                input.IsValidated &&
                input.Value.HasValue &&
                BasesMatch(requestedCanonicalBasis, requestedBasis, input.Basis));
            if (match is not null) return new ResolvedFinancialCalculationBasis { Basis = basis, Value = match.Value, Source = match.Source, Period = match.Period, ResolvedAtUtc = DateTime.UtcNow };
        }
        return new ResolvedFinancialCalculationBasis { Basis = basis, Source = "Unavailable", ResolvedAtUtc = DateTime.UtcNow, Warning = $"No usable value was found for {basis}." };
    }

    private static bool BasesMatch(FinancialCalculationBasisId requestedCanonicalBasis, string requestedBasis, string candidateBasis)
    {
        var candidateCanonicalBasis = FinancialCalculationBasisCatalog.Canonicalize(candidateBasis);
        return requestedCanonicalBasis != FinancialCalculationBasisId.Unknown
            ? candidateCanonicalBasis == requestedCanonicalBasis
            : FinancialCalculationBasisCatalog.Normalize(candidateBasis).Equals(requestedBasis, StringComparison.OrdinalIgnoreCase);
    }
}

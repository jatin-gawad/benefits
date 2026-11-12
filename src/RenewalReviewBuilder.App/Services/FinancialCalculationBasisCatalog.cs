using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

/// <summary>Canonical identifiers for financial calculation bases. Display labels remain source-facing only.</summary>
public static class FinancialCalculationBasisCatalog
{
    public static FinancialCalculationBasisId Canonicalize(string? basis)
    {
        var value = Normalize(basis);
        if (value.Equals("total benefit payments", StringComparison.OrdinalIgnoreCase) ||
            (value.StartsWith("annualized total benefit payments (", StringComparison.OrdinalIgnoreCase) && value.EndsWith("-month rolling window)", StringComparison.OrdinalIgnoreCase)))
            return FinancialCalculationBasisId.TotalBenefitPayments;
        return value.ToLowerInvariant() switch
        {
            "non-pooled ehc benefit payments" => FinancialCalculationBasisId.NonPooledEhcBenefitPayments,
            "dental benefit payments" => FinancialCalculationBasisId.DentalBenefitPayments,
            "billed premium" => FinancialCalculationBasisId.BilledPremium,
            "lives" or "covered lives" => FinancialCalculationBasisId.Lives,
            "flat amount" => FinancialCalculationBasisId.FlatAmount,
            _ => FinancialCalculationBasisId.Unknown
        };
    }

    public static string Display(FinancialCalculationBasisId basis) => basis switch
    {
        FinancialCalculationBasisId.TotalBenefitPayments => "Total benefit payments",
        FinancialCalculationBasisId.NonPooledEhcBenefitPayments => "Non-pooled EHC benefit payments",
        FinancialCalculationBasisId.DentalBenefitPayments => "Dental benefit payments",
        FinancialCalculationBasisId.BilledPremium => "Billed premium",
        FinancialCalculationBasisId.Lives => "Lives",
        FinancialCalculationBasisId.FlatAmount => "Flat amount",
        _ => "Unknown basis"
    };

    public static string Normalize(string? value) => string.Join(' ', (value ?? "").Trim().Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
}

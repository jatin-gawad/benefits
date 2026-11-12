using System.Globalization;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services;

namespace RenewalReviewBuilder.App;

public sealed partial class NegotiationRateRow
{
    public decimal? EditableCurrentRate
    {
        get => CurrentRate;
        set => ApplyBaselineRateEdit(value, isCurrent: true);
    }

    public decimal? EditableProposedRate
    {
        get => ProposedRate;
        set => ApplyBaselineRateEdit(value, isCurrent: false);
    }

    private void ApplyBaselineRateEdit(decimal? rate, bool isCurrent)
    {
        if (!rate.HasValue || rate < 0m)
        {
            throw new ArgumentException("Enter a rate of zero or greater.", nameof(rate));
        }

        var previousRate = isCurrent ? CurrentRate : ProposedRate;
        if (rate == previousRate)
        {
            return;
        }

        // Establish monthly/annual units before replacing the old rate/premium pair.
        var scale = MainWindow.PremiumScaleForRateCalculation(this);
        var monthlyPremium = UnderwritingEngine.CalculatePremium(BenefitLine, Volume, Lives, rate, RateBasis);
        var premium = monthlyPremium.HasValue ? decimal.Round(monthlyPremium.Value * scale, 2) : (decimal?)null;
        if (isCurrent)
        {
            CurrentRate = rate;
            CurrentPremium = premium ?? CurrentPremium;
            if (RateValueBasis.Equals(RateValueBasisNames.Gross, StringComparison.OrdinalIgnoreCase))
                CurrentGrossRate = rate;
            else if (RateValueBasis.Equals(RateValueBasisNames.Net, StringComparison.OrdinalIgnoreCase))
                CurrentNetRate = rate;
        }
        else
        {
            ProposedRate = rate;
            ProposedPremium = premium ?? ProposedPremium;
            if (RateValueBasis.Equals(RateValueBasisNames.Gross, StringComparison.OrdinalIgnoreCase))
                ProposedGrossRate = rate;
            else if (RateValueBasis.Equals(RateValueBasisNames.Net, StringComparison.OrdinalIgnoreCase))
                ProposedNetRate = rate;
        }

        var label = isCurrent ? "Current" : "Proposed";
        var oldText = previousRate?.ToString("0.####", CultureInfo.InvariantCulture) ?? "missing";
        var newText = rate.Value.ToString("0.####", CultureInfo.InvariantCulture);
        var note = $"{label} rate corrected from {oldText} to {newText}.";
        if (!premium.HasValue)
            note += " Premium retained because the rate basis, volume, or lives do not support recalculation.";
        Notes = string.IsNullOrWhiteSpace(Notes) ? note : $"{Notes} {note}";
    }
}

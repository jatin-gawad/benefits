using System.Globalization;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public static class BenefitRuleDisplayFormatter
{
    public const string ReviewRequiredDisplayText = "Tiered reimbursement details require review";

    private static readonly CultureInfo CanadianCulture = CultureInfo.GetCultureInfo("en-CA");

    public static string FormatFull(PlanBenefitRule? rule)
    {
        if (rule is null || !rule.IsDisplayable)
        {
            return "";
        }

        return $"{FormatPercent(rule.Tier1ReimbursementPercent!.Value)} of the first " +
               $"{FormatCurrency(rule.Tier1ThresholdAmount!.Value)} per {FormatBasis(rule.ThresholdBasis!.Value)} " +
               $"per {FormatPeriod(rule.ResetPeriod!.Value)}, then {FormatPercent(rule.Tier2ReimbursementPercent!.Value)}";
    }

    public static string FormatCompact(PlanBenefitRule? rule)
    {
        if (rule is null || !rule.IsDisplayable)
        {
            return "";
        }

        return $"{FormatPercent(rule.Tier1ReimbursementPercent!.Value)} of first " +
               $"{FormatCurrency(rule.Tier1ThresholdAmount!.Value)}/{FormatBasis(rule.ThresholdBasis!.Value)}/" +
               $"{FormatCompactPeriod(rule.ResetPeriod!.Value)}, then {FormatPercent(rule.Tier2ReimbursementPercent!.Value)}";
    }

    private static string FormatPercent(decimal value)
    {
        return $"{value.ToString(value == decimal.Truncate(value) ? "0" : "0.##", CanadianCulture)}%";
    }

    private static string FormatCurrency(decimal value)
    {
        return value == decimal.Truncate(value)
            ? value.ToString("C0", CanadianCulture)
            : value.ToString("C2", CanadianCulture);
    }

    private static string FormatBasis(BenefitRuleThresholdBasis basis)
    {
        return basis switch
        {
            BenefitRuleThresholdBasis.Certificate => "certificate",
            BenefitRuleThresholdBasis.Person => "person",
            _ => "benefit"
        };
    }

    private static string FormatPeriod(BenefitRuleResetPeriod period)
    {
        return period switch
        {
            BenefitRuleResetPeriod.CalendarYear => "calendar year",
            BenefitRuleResetPeriod.PolicyYear => "policy year",
            _ => "period"
        };
    }

    private static string FormatCompactPeriod(BenefitRuleResetPeriod period)
    {
        return period switch
        {
            BenefitRuleResetPeriod.CalendarYear => "year",
            BenefitRuleResetPeriod.PolicyYear => "policy year",
            _ => "period"
        };
    }
}

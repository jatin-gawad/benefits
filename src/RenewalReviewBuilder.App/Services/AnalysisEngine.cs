using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed class AnalysisEngine
{
    private readonly RenewalEngine _renewalEngine = new();

    public void Analyze(AnalysisResult result)
    {
        AddSourceFindings(result);
        AnalyzeRates(result);
        _renewalEngine.BuildRenewalVersions(result);
        AnalyzePremiumChange(result);
        AnalyzeExperience(result);
        AddPaidLossRatioSourceChecks(result);
        AnalyzeClaims(result);
        AddFutureInputChecks(result);
    }

    private static void AddSourceFindings(AnalysisResult result)
    {
        foreach (var warning in result.Pdf.Warnings)
        {
            AddFinding(result, FindingSeverity.Warning, "PDF extraction needs review", warning, "Review the PDF manually and add a parser rule if the carrier changed the layout.", warning, "PDF");
        }

        foreach (var warning in result.Workbook.Warnings)
        {
            AddFinding(result, FindingSeverity.Warning, "Workbook extraction needs review", warning, "Review the workbook tab manually and add a parser rule if the carrier changed the layout.", warning, "XLSM");
        }

        foreach (var missingSheet in result.Workbook.MissingSheets)
        {
            AddFinding(
                result,
                FindingSeverity.Critical,
                "Missing carrier workbook sheet",
                $"The sheet '{missingSheet}' was not found.",
                "Confirm this is the correct carrier or claims-history workbook.",
                "The detailed claim analysis could be incomplete because one expected carrier worksheet was missing.",
                "XLSM");
        }
    }

    private static void AnalyzeRates(AnalysisResult result)
    {
        foreach (var line in result.Pdf.RateLines)
        {
            var canCalculate = !line.ValidationStatus.Equals(
                RateLineValidationStatuses.NeedsReview,
                StringComparison.OrdinalIgnoreCase);
            line.CalculatedCurrentPremium = canCalculate
                ? UnderwritingEngine.CalculatePremium(line.Benefit, line.Volume, line.Lives, line.CurrentRate, line.RateBasis)
                : null;
            line.CalculatedRenewalPremium = canCalculate
                ? UnderwritingEngine.CalculatePremium(line.Benefit, line.Volume, line.Lives, line.RenewalRate, line.RateBasis)
                : null;
            line.PercentChange ??= UnderwritingEngine.CalculatePercentChange(line.CurrentPremium, line.RenewalPremium);

            if (line.CurrentPremium.HasValue && line.CalculatedCurrentPremium.HasValue)
            {
                line.CurrentVariance = decimal.Round(line.CurrentPremium.Value - line.CalculatedCurrentPremium.Value, 2);
            }

            if (line.RenewalPremium.HasValue && line.CalculatedRenewalPremium.HasValue)
            {
                line.RenewalVariance = decimal.Round(line.RenewalPremium.Value - line.CalculatedRenewalPremium.Value, 2);
            }

            var largestVariance = new[] { line.CurrentVariance, line.RenewalVariance }
                .Where(value => value.HasValue)
                .Select(value => Math.Abs(value!.Value))
                .DefaultIfEmpty(0m)
                .Max();

            if (largestVariance > 2m)
            {
                AddFinding(
                    result,
                    FindingSeverity.Warning,
                    "Rate illustration premium does not recalculate cleanly",
                    $"{line.Benefit} {line.TierOrClass}: largest variance is {ParsingHelpers.Money(largestVariance)}.",
                    "Check the rate basis, volume, lives, and carrier rounding before presenting this line.",
                    "This line should be validated before it is used as support for the renewal rate change.",
                    line.SourceReference);
            }
        }

        AddRateSourceChecks(result);
        AddPoolingSourceCheck(result);
        AddPooledPremiumCrossCheck(result);
    }

    private static void AddRateSourceChecks(AnalysisResult result)
    {
        AddPremiumCheck(
            result,
            "Rate total vs PDF",
            result.Pdf.RateLines.Where(line => line.IncludeInPremiumTotal).Sum(line => line.CurrentPremium ?? 0m),
            result.Pdf.RateLines.Where(line => line.IncludeInPremiumTotal).Sum(line => line.RenewalPremium ?? 0m),
            result.Pdf.CurrentMonthlyPremium,
            result.Pdf.RenewalMonthlyPremium,
            "extracted benefit rows",
            "PDF Total Monthly Premium");

        var ehcLines = result.Pdf.RateLines
            .Where(line => line.Benefit.StartsWith("Extended Health Care", StringComparison.OrdinalIgnoreCase) &&
                           !line.Benefit.Equals("Extended Health Care - Total", StringComparison.OrdinalIgnoreCase))
            .ToList();

        AddPremiumCheck(
            result,
            "Extended Health Care detail vs PDF",
            ehcLines.Sum(line => line.CurrentPremium ?? 0m),
            ehcLines.Sum(line => line.RenewalPremium ?? 0m),
            result.Pdf.CurrentEhcPremium,
            result.Pdf.RenewalEhcPremium,
            "extracted EHC detail rows",
            "PDF Total Extended Health Care");

        var dentalLines = result.Pdf.RateLines
            .Where(line => line.Benefit.Equals("Dental", StringComparison.OrdinalIgnoreCase))
            .ToList();

        AddPremiumCheck(
            result,
            "Dental detail vs PDF",
            dentalLines.Sum(line => line.CurrentPremium ?? 0m),
            dentalLines.Sum(line => line.RenewalPremium ?? 0m),
            result.Pdf.CurrentDentalPremium,
            result.Pdf.RenewalDentalPremium,
            "extracted Dental detail rows",
            "PDF Total Dental");
    }

    private static void AddPoolingSourceCheck(AnalysisResult result)
    {
        if (!result.Pdf.PoolingThreshold.HasValue &&
            !result.Pdf.PoolingCharge.HasValue &&
            !result.Pdf.PooledPremium.HasValue)
        {
            return;
        }

        var details = new List<string>();
        if (result.Pdf.PoolingThreshold.HasValue)
        {
            details.Add($"threshold {ParsingHelpers.Money(result.Pdf.PoolingThreshold)}");
        }

        if (result.Pdf.PoolingCharge.HasValue)
        {
            details.Add($"charge {ParsingHelpers.Percent(result.Pdf.PoolingCharge)}");
        }

        if (result.Pdf.PooledPremium.HasValue)
        {
            details.Add($"pooled premium {ParsingHelpers.Money(result.Pdf.PooledPremium)}");
        }

        result.SourceChecks.Add(new SourceCheck
        {
            Area = "Pooling arrangement",
            Status = result.Pdf.PoolingThreshold.HasValue && result.Pdf.PoolingCharge.HasValue ? "Read" : "Needs review",
            Detail = string.Join("; ", details)
        });
    }

    private static void AddPooledPremiumCrossCheck(AnalysisResult result)
    {
        var carrierAnnual = result.Pdf.PooledPremium;
        var estimatedAnnual = CalculateEstimatedAnnualPooledPremium(result);

        if (!carrierAnnual.HasValue && !estimatedAnnual.HasValue)
        {
            return;
        }

        if (!carrierAnnual.HasValue || !estimatedAnnual.HasValue)
        {
            result.SourceChecks.Add(new SourceCheck
            {
                Area = "Pooled premium cross-check",
                Status = "Needs review",
                Detail = !carrierAnnual.HasValue
                    ? "Carrier pooled premium was not extracted from the renewal analysis block."
                    : "Estimated annual pooled premium could not be calculated from Extended Health Care - Pooled rate rows."
            });
            return;
        }

        var variance = decimal.Round(estimatedAnnual.Value - carrierAnnual.Value, 2);
        var variancePercent = carrierAnnual.Value == 0m ? (decimal?)null : decimal.Round(variance / carrierAnnual.Value * 100m, 1);
        var status = Math.Abs(variance) <= 25m || Math.Abs(variancePercent ?? 0m) <= 1m ? "Matched" : "Mismatch";

        result.SourceChecks.Add(new SourceCheck
        {
            Area = "Pooled premium cross-check",
            Status = status,
            Detail = $"Estimated annual pooled premium {ParsingHelpers.Money(estimatedAnnual)} vs carrier pooled premium {ParsingHelpers.Money(carrierAnnual)}; variance {ParsingHelpers.Money(variance)} ({ParsingHelpers.Percent(variancePercent)})."
        });

        if (status == "Mismatch")
        {
            AddFinding(
                result,
                FindingSeverity.Warning,
                "Pooled premium estimate does not tie to carrier renewal",
                $"Estimated annual pooled premium differs from the Manulife pooled premium by {ParsingHelpers.Money(variance)} ({ParsingHelpers.Percent(variancePercent)}).",
                "Review the pooled EHC rate rows and renewal analysis pooled premium before presenting the renewal exhibit.",
                "The renewal pooling charge needs a manual check before it is used as support for pricing.",
                "Manulife pooled premium section");
        }
    }

    private static decimal? CalculateEstimatedAnnualPooledPremium(AnalysisResult result)
    {
        var pooledLines = result.Pdf.RateLines
            .Where(line => line.Benefit.Equals("Extended Health Care - Pooled", StringComparison.OrdinalIgnoreCase))
            .ToList();

        if (pooledLines.Count == 0)
        {
            return null;
        }

        var monthly = pooledLines.Sum(line =>
            line.CalculatedRenewalPremium ??
            UnderwritingEngine.CalculatePremium(line.Benefit, line.Volume, line.Lives, line.RenewalRate, line.RateBasis) ??
            line.RenewalPremium ??
            0m);

        return monthly > 0m ? decimal.Round(monthly * 12m, 2) : null;
    }

    private static void AddPremiumCheck(
        AnalysisResult result,
        string area,
        decimal extractedCurrent,
        decimal extractedRenewal,
        decimal? pdfCurrent,
        decimal? pdfRenewal,
        string extractedLabel,
        string pdfLabel)
    {
        if (!pdfCurrent.HasValue || !pdfRenewal.HasValue)
        {
            result.SourceChecks.Add(new SourceCheck
            {
                Area = area,
                Status = "Needs review",
                Detail = $"{pdfLabel} was not extracted from the PDF."
            });
            return;
        }

        var currentVariance = decimal.Round(extractedCurrent - pdfCurrent.Value, 2);
        var renewalVariance = decimal.Round(extractedRenewal - pdfRenewal.Value, 2);
        var status = Math.Abs(currentVariance) <= 0.05m && Math.Abs(renewalVariance) <= 0.05m ? "Matched" : "Mismatch";

        result.SourceChecks.Add(new SourceCheck
        {
            Area = area,
            Status = status,
            Detail = $"{extractedLabel}: current {ParsingHelpers.Money(extractedCurrent)} vs {pdfLabel} {ParsingHelpers.Money(pdfCurrent)}; renewal {ParsingHelpers.Money(extractedRenewal)} vs {pdfLabel} {ParsingHelpers.Money(pdfRenewal)}."
        });

        if (status == "Mismatch")
        {
            AddFinding(
                result,
                FindingSeverity.Warning,
                $"{area} mismatch",
                $"Current variance is {ParsingHelpers.Money(currentVariance)} and renewal variance is {ParsingHelpers.Money(renewalVariance)}.",
                "Review the rate illustration rows against the PDF before presenting the pack.",
                "One or more carrier rate rows did not tie to the PDF total, so the exhibit needs manual review.",
                pdfLabel);
        }
    }

    private static void AnalyzePremiumChange(AnalysisResult result)
    {
        if (result.MonthlyPremiumChangePercent is null)
        {
            return;
        }

        var severity = result.MonthlyPremiumChangePercent >= 15m ? FindingSeverity.Warning : FindingSeverity.Info;
        AddFinding(
            result,
            severity,
            "Total monthly premium change",
            $"The carrier renewal moves from {ParsingHelpers.Money(result.CurrentMonthlyPremium)} to {ParsingHelpers.Money(result.RenewalMonthlyPremium)}, an increase of {ParsingHelpers.Money(result.MonthlyPremiumChange)} or {ParsingHelpers.Percent(result.MonthlyPremiumChangePercent)}.",
            "Use the loss-ratio and trend findings to explain whether the increase is justified.",
            "The client will feel the total monthly change first; lead with why the funding level needs to move.",
            "PDF renewal rate illustration");

        if (result.Pdf.CurrentEhcPremium.HasValue && result.Pdf.RenewalEhcPremium.HasValue)
        {
            var increase = UnderwritingEngine.CalculatePercentChange(result.Pdf.CurrentEhcPremium.Value, result.Pdf.RenewalEhcPremium.Value);
            if (increase >= 15m)
            {
                AddFinding(
                    result,
                    FindingSeverity.Warning,
                    "Extended Health Care is a major increase driver",
                    $"EHC premium changes from {ParsingHelpers.Money(result.Pdf.CurrentEhcPremium)} to {ParsingHelpers.Money(result.Pdf.RenewalEhcPremium)}, or {increase:0.0}%.",
                    "Tie the EHC increase to category trends, benchmarks, and claim concentration.",
                    "EHC is one of the main explanations for the renewal increase.",
                    "PDF renewal rate illustration");
            }
        }

        if (result.Pdf.CurrentDentalPremium.HasValue && result.Pdf.RenewalDentalPremium.HasValue)
        {
            var increase = UnderwritingEngine.CalculatePercentChange(result.Pdf.CurrentDentalPremium.Value, result.Pdf.RenewalDentalPremium.Value);
            if (increase >= 5m)
            {
                AddFinding(
                    result,
                    FindingSeverity.Info,
                    "Dental premium pressure present",
                    $"Dental premium changes from {ParsingHelpers.Money(result.Pdf.CurrentDentalPremium)} to {ParsingHelpers.Money(result.Pdf.RenewalDentalPremium)}, or {increase:0.0}%.",
                    "Use paid/incurred loss ratios and basic-service spend to explain the increase.",
                    "Dental experience supports part of the renewal adjustment.",
                    "PDF renewal rate illustration");
            }
        }
    }

    private static void AnalyzeExperience(AnalysisResult result)
    {
        foreach (var group in result.Pdf.ExperienceLines.GroupBy(line => line.Benefit))
        {
            var latest = group
                .OrderByDescending(line => line.Period)
                .FirstOrDefault(line => line.PaidLossRatio.HasValue || line.IncurredLossRatio.HasValue);

            if (latest is null)
            {
                continue;
            }

            var lossRatio = latest.PaidLossRatio ?? latest.IncurredLossRatio;
            if (lossRatio is null)
            {
                continue;
            }

            var severity = lossRatio >= 100m ? FindingSeverity.Critical :
                lossRatio >= 85m ? FindingSeverity.Warning :
                FindingSeverity.Info;

            if (severity == FindingSeverity.Info && !group.Key.Contains("Health", StringComparison.OrdinalIgnoreCase) && !group.Key.Contains("Dental", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            AddFinding(
                result,
                severity,
                $"{group.Key} loss ratio signal",
                $"Latest extracted period {latest.Period} shows a {(latest.PaidLossRatio.HasValue ? "paid" : "reported")} loss ratio of {ParsingHelpers.Percent(lossRatio)}.",
                lossRatio >= 85m ? "Use this as renewal justification and confirm whether any large claims are ongoing." : "Keep as supporting context; this line is not the main pressure point.",
                lossRatio >= 85m ? "Claims are running high relative to premium, so the rate change is easier to justify." : "This line appears more stable than the main increase drivers.",
                latest.SourceReference);
        }
    }

    private static void AddPaidLossRatioSourceChecks(AnalysisResult result)
    {
        AddPaidLossRatioSourceCheck(
            result,
            "EHC paid loss ratio vs source",
            row => BenefitMatcher.ContainsPhrase(row.Benefit, "health") || BenefitMatcher.ContainsPhrase(row.Benefit, "ehc"));

        AddPaidLossRatioSourceCheck(
            result,
            "Dental paid loss ratio vs source",
            row => BenefitMatcher.ContainsPhrase(row.Benefit, "dental"));
    }

    private static void AddPaidLossRatioSourceCheck(
        AnalysisResult result,
        string area,
        Func<ExperienceLine, bool> benefitFilter)
    {
        var benefitRows = result.Pdf.ExperienceLines
            .Where(benefitFilter)
            .ToList();

        if (benefitRows.Count == 0)
        {
            return;
        }

        var latest = benefitRows
            .Where(row => row.BilledPremium is > 0m && row.PaidClaims.HasValue)
            .OrderByDescending(ExperienceLineSelector.PeriodStart)
            .FirstOrDefault();

        if (latest is null)
        {
            result.SourceChecks.Add(new SourceCheck
            {
                Area = area,
                Status = "Needs review",
                Detail = "Paid claims and billed premium were not available together for this benefit."
            });
            return;
        }

        var calculated = decimal.Round(latest.PaidClaims!.Value / latest.BilledPremium!.Value * 100m, 2);
        var stored = latest.PaidLossRatio;
        var status = stored.HasValue && Math.Abs(stored.Value - calculated) <= 0.2m
            ? "Matched"
            : "Needs review";

        result.SourceChecks.Add(new SourceCheck
        {
            Area = area,
            Status = status,
            Detail = stored.HasValue
                ? $"Paid claims {ParsingHelpers.Money(latest.PaidClaims)} / billed premium {ParsingHelpers.Money(latest.BilledPremium)} = {ParsingHelpers.Percent(calculated)}; app value {ParsingHelpers.Percent(stored)}."
                : $"Paid claims {ParsingHelpers.Money(latest.PaidClaims)} / billed premium {ParsingHelpers.Money(latest.BilledPremium)} = {ParsingHelpers.Percent(calculated)}; app value was not populated."
        });
    }

    private static void AnalyzeClaims(AnalysisResult result)
    {
        var ehcParamedical = result.Workbook.ClaimTrends
            .Where(trend => trend.Benefit.Contains("Health", StringComparison.OrdinalIgnoreCase) &&
                            trend.Category.Contains("Paramedical", StringComparison.OrdinalIgnoreCase))
            .OrderByDescending(trend => trend.PercentOfTotal ?? 0m)
            .FirstOrDefault();

        if (ehcParamedical?.PercentOfTotal >= 40m)
        {
            AddFinding(
                result,
                FindingSeverity.Warning,
                "High paramedical utilization",
                $"Paramedical represents {ParsingHelpers.Percent(ehcParamedical.PercentOfTotal)} of extracted EHC claim cost.",
                "Review paramedical maximums and practitioner usage. Massage-specific detail should be added when the source includes that subcategory.",
                "EHC utilization is concentrated in paramedical services, which is a practical explanation for renewal pressure.",
                ehcParamedical.SourceReference);
        }

        var dentalBasic = result.Workbook.ClaimTrends
            .Where(trend => trend.Benefit.Equals("Dental", StringComparison.OrdinalIgnoreCase) &&
                            trend.Category.Contains("Basic", StringComparison.OrdinalIgnoreCase))
            .OrderByDescending(trend => trend.PercentOfTotal ?? 0m)
            .FirstOrDefault();

        if (dentalBasic?.PercentOfTotal >= 65m)
        {
            AddFinding(
                result,
                FindingSeverity.Info,
                "Dental claims are mainly basic services",
                $"Basic dental accounts for {ParsingHelpers.Percent(dentalBasic.PercentOfTotal)} of extracted dental claim cost.",
                "Use this to explain that dental pressure is driven by recurring utilization, not only one-off major claims.",
                "Dental cost pressure appears broad and recurring.",
                dentalBasic.SourceReference);
        }

        var indicationDrivers = result.Workbook.IndicationTrends
            .Where(indication => !IsSummaryRow(indication.Indication))
            .ToList();
        var totalIndicationSpend = indicationDrivers.Sum(indication => indication.AllowableTotal ?? indication.SubmittedTotal ?? 0m);
        var diabetes = indicationDrivers
            .Where(indication => indication.Indication.Contains("Diabetes", StringComparison.OrdinalIgnoreCase))
            .OrderByDescending(indication => indication.AllowableTotal ?? indication.SubmittedTotal ?? 0m)
            .FirstOrDefault();

        if (diabetes is not null && totalIndicationSpend > 0)
        {
            var share = decimal.Round((diabetes.AllowableTotal ?? diabetes.SubmittedTotal ?? 0m) / totalIndicationSpend * 100m, 1);
            if (share >= 20m)
            {
                AddFinding(
                    result,
                    FindingSeverity.Warning,
                    "Drug spend concentration in diabetes",
                    $"Diabetes-related drugs represent about {share:0.0}% of top indication spend.",
                    "Review whether high-cost diabetes drugs are recurring and whether plan design or prior authorization applies.",
                    "A meaningful share of drug cost is concentrated in diabetes-related medications.",
                    diabetes.SourceReference);
            }
        }

        var topDrug = result.Workbook.DrugTrends
            .Where(drug => !IsSummaryRow(drug.DrugName))
            .OrderByDescending(drug => drug.NetAmountPaid ?? drug.SubmittedTotal ?? 0m)
            .FirstOrDefault();
        var topDrugAmount = topDrug?.NetAmountPaid ?? topDrug?.SubmittedTotal;

        if (topDrug is not null && topDrugAmount is >= 5000m)
        {
            AddFinding(
                result,
                FindingSeverity.Info,
                "Top drug cost should be reviewed",
                $"{topDrug.DrugName} is the largest extracted DIN by net amount paid at {ParsingHelpers.Money(topDrugAmount)}.",
                "Check whether this is a recurring medication and whether the carrier renewal already reflects it.",
                "A small number of drugs may be influencing EHC experience.",
                topDrug.SourceReference);
        }

        var totalCertificateClaims = result.Workbook.CertificateClaims.Sum(certificate => certificate.Total ?? 0m);
        var topCertificate = result.Workbook.CertificateClaims
            .OrderByDescending(certificate => certificate.Total ?? 0m)
            .FirstOrDefault();

        if (topCertificate is not null && totalCertificateClaims > 0)
        {
            var share = decimal.Round((topCertificate.Total ?? 0m) / totalCertificateClaims * 100m, 1);
            if (share >= 15m)
            {
                AddFinding(
                    result,
                    FindingSeverity.Warning,
                    "Claims concentrated in top certificate",
                    $"Certificate {topCertificate.Certificate} accounts for about {share:0.0}% of extracted EHC + dental claims.",
                    "Review whether this claimant is ongoing, pooled, or likely to normalize.",
                    "A concentrated claimant can explain volatility and should be separated from ordinary utilization.",
                    topCertificate.SourceReference);
            }
        }
    }

    private static void AddFutureInputChecks(AnalysisResult result)
    {
        AddFinding(
            result,
            FindingSeverity.Info,
            "Billing statement not provided in Phase 1",
            "The current billing statement is not part of the Phase 1 input set.",
            "Add billing statement ingestion in a future phase to reconcile live billed rates against the renewal illustration.",
            "This review verifies the renewal package and supporting claims workbook, but not the current bill.",
            "Phase 1 scope");

        if (result.PreviousRenewals.Count > 0)
        {
            return;
        }

        var suppliedPriorRenewalInputs = result.RenewalCase.AdditionalPdfPaths
            .Where(path => !string.IsNullOrWhiteSpace(path))
            .Where(path =>
            {
                var fileName = Path.GetFileName(path);
                return fileName.Contains("renewal", StringComparison.OrdinalIgnoreCase) ||
                       fileName.Contains("amendment", StringComparison.OrdinalIgnoreCase);
            })
            .Concat(string.IsNullOrWhiteSpace(result.RenewalCase.PriorYearFolder)
                ? []
                : [result.RenewalCase.PriorYearFolder])
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (suppliedPriorRenewalInputs.Count > 0)
        {
            AddFinding(
                result,
                FindingSeverity.Warning,
                "Prior renewal supplied but not recognized",
                $"A prior renewal input was supplied, but it could not be matched to this renewal: {string.Join(", ", suppliedPriorRenewalInputs.Select(Path.GetFileName))}.",
                "Confirm that the prior document belongs to the same client or policy and contains readable renewal details.",
                "The prior document is present, but historical comparisons are unavailable until it can be matched.",
                string.Join("; ", suppliedPriorRenewalInputs));
            return;
        }

        AddFinding(
            result,
            FindingSeverity.Info,
            "Prior renewal not provided",
            "No prior renewal or amendment document was identified in Additional PDFs or the Prior Year Folder.",
            "Add the prior renewal or amendment when historical rate-change validation is required.",
            "Historical renewal and amendment comparisons are unavailable for this review.",
            "Renewal inputs");
    }

    private static bool IsSummaryRow(string label)
    {
        return string.IsNullOrWhiteSpace(label) ||
               label.Equals("Grand Total", StringComparison.OrdinalIgnoreCase) ||
               label.Contains("Top 50", StringComparison.OrdinalIgnoreCase) ||
               label.Contains("Total", StringComparison.OrdinalIgnoreCase);
    }

    private static void AddFinding(
        AnalysisResult result,
        FindingSeverity severity,
        string title,
        string evidence,
        string recommendedAction,
        string clientExplanation,
        string sourceReference)
    {
        result.Findings.Add(new Finding
        {
            Severity = severity,
            Title = title,
            Evidence = evidence,
            RecommendedAction = recommendedAction,
            ClientExplanation = clientExplanation,
            SourceReference = sourceReference
        });
    }
}

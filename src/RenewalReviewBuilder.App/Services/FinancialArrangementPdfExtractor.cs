using System.Globalization;
using System.Text.RegularExpressions;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

/// <summary>Extracts only normalized financial terms; it does not persist or calculate them.</summary>
public sealed class FinancialArrangementPdfExtractor
{
    private readonly PdfTextExtractor _pdfTextExtractor;

    public FinancialArrangementPdfExtractor(PdfTextExtractor? pdfTextExtractor = null) => _pdfTextExtractor = pdfTextExtractor ?? new PdfTextExtractor();

    public FinancialArrangementExtraction Extract(string path, DateOnly effectiveDate)
    {
        var result = new FinancialArrangementExtraction { EffectiveDate = effectiveDate };
        foreach (var page in _pdfTextExtractor.Extract(path))
        {
            for (var lineIndex = 0; lineIndex < page.Lines.Count; lineIndex++)
            {
                var line = page.Lines[lineIndex];
                var evidenceLine = string.Join(" ", page.Lines.Skip(lineIndex).Take(3));
                if (IsTargetLossRatioLine(line) && string.IsNullOrWhiteSpace(result.TargetLossRatioSourceText))
                {
                    result.TargetLossRatio = FindPercent(evidenceLine);
                    result.TlrBasis = FindTlrBasis(evidenceLine);
                    result.TargetLossRatioSourceText = evidenceLine;
                    result.TargetLossRatioSourcePage = page.PageNumber;
                    result.TargetLossRatioExtractionConfidence = result.TargetLossRatio.HasValue && !string.IsNullOrWhiteSpace(result.TlrBasis) ? 0.85m : 0.55m;
                    if (result.HasIncompleteTargetLossRatioEvidence)
                    {
                        result.Warnings.Add("Target loss ratio evidence requires review before it can become current client terms.");
                    }
                }
                var chargeType = ClassifyCharge(line);
                var percent = FindPercent(line);
                var sourceLine = line;
                if (chargeType.HasValue && !percent.HasValue)
                {
                    for (var lookAhead = lineIndex + 1; lookAhead < Math.Min(page.Lines.Count, lineIndex + 4); lookAhead++)
                    {
                        percent = FindPercent(page.Lines[lookAhead]);
                        if (percent.HasValue)
                        {
                            sourceLine = $"{line} {page.Lines[lookAhead]}";
                            break;
                        }
                    }
                }
                if (chargeType.HasValue && percent.HasValue)
                {
                    if (result.Charges.Any(charge => charge.ChargeType == chargeType.Value && charge.Benefit.Equals("All", StringComparison.OrdinalIgnoreCase)))
                    {
                        continue;
                    }
                    result.Charges.Add(new FinancialArrangementCharge
                    {
                        ChargeType = chargeType.Value,
                        Rate = percent.Value,
                        Basis = FindBasis(sourceLine),
                        SourceText = sourceLine,
                        SourcePage = page.PageNumber,
                        ExtractionConfidence = 0.85m
                    });
                }
                else if (chargeType.HasValue && TryFindFormula(line, out var formula) && !result.Charges.Any(charge => charge.ChargeType == chargeType.Value && charge.Formula.Equals(formula, StringComparison.OrdinalIgnoreCase)))
                {
                    result.Charges.Add(new FinancialArrangementCharge { ChargeType = chargeType.Value, Formula = formula, Basis = FindBasis(line), SourceText = line, SourcePage = page.PageNumber, ExtractionConfidence = 0.8m });
                }

                if (line.Contains("pool", StringComparison.OrdinalIgnoreCase))
                {
                    var pooling = new PoolingArrangement { EffectiveDate = effectiveDate, SourceDocumentId = null, Benefit = evidenceLine.Contains("EHC", StringComparison.OrdinalIgnoreCase) ? "EHC" : "All", Basis = FindPoolingBasis(evidenceLine), PoolingType = evidenceLine.Contains("large amount", StringComparison.OrdinalIgnoreCase) ? "Large Amount" : "", SourceText = evidenceLine, SourcePage = page.PageNumber };
                    var threshold = FindDollar(evidenceLine);
                    var poolingPercent = FindPercent(evidenceLine);
                    if (poolingPercent.HasValue) pooling.Rate = poolingPercent;
                    if (threshold.HasValue) pooling.Threshold = threshold;
                    if (pooling.Rate.HasValue || pooling.Threshold.HasValue)
                    {
                        pooling.ExtractionConfidence = pooling.Rate.HasValue && pooling.Threshold.HasValue ? 0.85m : 0.65m;
                        // Adjacent PDF text lines can repeat the same pooling wording in the
                        // three-line evidence window. Keep one canonical extracted term.
                        if (!result.Pooling.Any(existing =>
                            existing.Benefit.Equals(pooling.Benefit, StringComparison.OrdinalIgnoreCase) &&
                            existing.Rate == pooling.Rate &&
                            existing.Threshold == pooling.Threshold &&
                            existing.Basis.Equals(pooling.Basis, StringComparison.OrdinalIgnoreCase)))
                        {
                            result.Pooling.Add(pooling);
                        }
                    }
                }
            }
        }
        if (result.Charges.Count == 0) result.Warnings.Add("No carrier administrative charges were extracted.");
        return result;
    }

    private static FinancialChargeType? ClassifyCharge(string text)
    {
        if (text.Contains("general administration", StringComparison.OrdinalIgnoreCase)) return FinancialChargeType.GeneralAdministration;
        if (text.Contains("benefit payment", StringComparison.OrdinalIgnoreCase) && text.Contains("administration", StringComparison.OrdinalIgnoreCase)) return FinancialChargeType.BenefitPaymentsAdministration;
        if (text.Contains("profit charge", StringComparison.OrdinalIgnoreCase)) return FinancialChargeType.ProfitCharge;
        if (text.Contains("printing", StringComparison.OrdinalIgnoreCase)) return FinancialChargeType.PrintingExpense;
        if (text.Contains("report preparation", StringComparison.OrdinalIgnoreCase)) return FinancialChargeType.ReportPreparationExpense;
        if (text.Contains("trip cancellation", StringComparison.OrdinalIgnoreCase)) return FinancialChargeType.TripCancellation;
        if (text.Contains("other administrative charge", StringComparison.OrdinalIgnoreCase) || text.Contains("referral expense", StringComparison.OrdinalIgnoreCase)) return FinancialChargeType.OtherCarrierCharge;
        if (text.Contains("premium tax", StringComparison.OrdinalIgnoreCase) || text.Contains("tax is payable", StringComparison.OrdinalIgnoreCase)) return FinancialChargeType.PremiumTax;
        if (text.Contains("interest rate", StringComparison.OrdinalIgnoreCase) || text.Contains("negative balance", StringComparison.OrdinalIgnoreCase) || text.Contains("positive balance", StringComparison.OrdinalIgnoreCase)) return FinancialChargeType.InterestTerm;
        return null;
    }

    private static decimal? FindPercent(string text)
    {
        var match = Regex.Match(text, @"(?<value>\d{1,3}(?:\.\d+)?)\s*%");
        return match.Success && decimal.TryParse(match.Groups["value"].Value, CultureInfo.InvariantCulture, out var value) ? value / 100m : null;
    }

    private static decimal? FindDollar(string text)
    {
        var match = Regex.Match(text, @"\$\s*(?<value>[\d,]+(?:\.\d{2})?)");
        return match.Success && decimal.TryParse(match.Groups["value"].Value.Replace(",", ""), CultureInfo.InvariantCulture, out var value) ? value : null;
    }

    private static string FindBasis(string text) => text.Contains("billed premium", StringComparison.OrdinalIgnoreCase) ? "Billed premium" : text.Contains("non-pooled", StringComparison.OrdinalIgnoreCase) ? "Non-pooled benefit payments" : "Total benefit payments";

    private static string FindPoolingBasis(string text) => text.Contains("non-pooled", StringComparison.OrdinalIgnoreCase) && text.Contains("EHC", StringComparison.OrdinalIgnoreCase) ? "Non-pooled EHC benefit payments" : FindBasis(text);

    private static bool IsTargetLossRatioLine(string text) =>
        text.Contains("target loss ratio", StringComparison.OrdinalIgnoreCase) ||
        Regex.IsMatch(text, @"\bTLR\b", RegexOptions.IgnoreCase);

    private static string FindTlrBasis(string text)
    {
        if (text.Contains("net of tax and commission", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("net of premium tax and commission", StringComparison.OrdinalIgnoreCase)) return "Net of tax and commission";
        if (text.Contains("gross billed premium", StringComparison.OrdinalIgnoreCase)) return "Gross billed premium";
        if (text.Contains("includes admin margin", StringComparison.OrdinalIgnoreCase)) return "Includes admin margin";
        if (text.Contains("excludes admin margin", StringComparison.OrdinalIgnoreCase)) return "Excludes admin margin";
        if (text.Contains("net of premium tax", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("net tlr", StringComparison.OrdinalIgnoreCase)) return "Net of premium tax only";
        return "";
    }

    private static bool TryFindFormula(string text, out string formula)
    {
        formula = text.Contains("actual cost", StringComparison.OrdinalIgnoreCase) ? text.Trim() : "";
        return formula.Length > 0;
    }
}

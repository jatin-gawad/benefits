using System.Globalization;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed class ProposalComparisonDataService
{
    private readonly ProposalIntelligenceAnalysisService _analysisService;
    private readonly CurrentVsQuotePlanDesignService _currentVsQuoteService;

    public ProposalComparisonDataService(string? databasePath = null)
        : this(
            new ProposalIntelligenceAnalysisService(databasePath),
            new CurrentVsQuotePlanDesignService(databasePath))
    {
    }

    public ProposalComparisonDataService(
        ProposalIntelligenceAnalysisService analysisService,
        CurrentVsQuotePlanDesignService currentVsQuoteService)
    {
        _analysisService = analysisService ?? throw new ArgumentNullException(nameof(analysisService));
        _currentVsQuoteService = currentVsQuoteService ?? throw new ArgumentNullException(nameof(currentVsQuoteService));
    }

    public async Task<ProposalComparisonViewModel> LoadAsync(
        long clientId,
        string? fallbackClientName = null,
        string? fallbackCarrier = null)
    {
        var analysis = await _analysisService.BuildAsync(clientId);
        var currentVsQuote = await _currentVsQuoteService.BuildAsync(clientId);
        return Build(analysis, currentVsQuote, fallbackClientName, fallbackCarrier);
    }

    public ProposalComparisonViewModel Build(
        ProposalIntelligenceAnalysis analysis,
        CurrentVsQuotePlanDesignComparison currentVsQuote,
        string? fallbackClientName = null,
        string? fallbackCarrier = null)
    {
        ArgumentNullException.ThrowIfNull(analysis);
        ArgumentNullException.ThrowIfNull(currentVsQuote);

        var quoteCount = analysis.Quotes.Count;
        var currentRowCount = currentVsQuote.Rows.Count;
        var financialRowCount = analysis.PremiumComparisons.Count + analysis.RateComparisons.Count;
        var model = new ProposalComparisonViewModel
        {
            ClientDatabaseId = analysis.ClientId > 0 ? analysis.ClientId : currentVsQuote.ClientId,
            ClientName = FirstNonBlank(
                analysis.ClientName,
                currentVsQuote.ClientName,
                fallbackClientName,
                $"Client #{(analysis.ClientId > 0 ? analysis.ClientId : currentVsQuote.ClientId).ToString(CultureInfo.InvariantCulture)}"),
            CurrentCarrier = FirstNonBlank(currentVsQuote.CurrentCarrier, fallbackCarrier, "Not available"),
            CurrentPlanStatus = currentVsQuote.HasCurrentPlan ? "Available" : "Missing",
            QuoteCountText = $"{quoteCount.ToString(CultureInfo.InvariantCulture)} saved {(quoteCount == 1 ? "proposal" : "proposals")}",
            ComparisonCountText = $"{currentRowCount.ToString(CultureInfo.InvariantCulture)} current-plan {(currentRowCount == 1 ? "provision" : "provisions")}",
            FinancialComparisonCountText = $"{financialRowCount.ToString(CultureInfo.InvariantCulture)} financial {(financialRowCount == 1 ? "row" : "rows")}",
            CarrierPlanDesignCountText = $"{analysis.PlanDesignComparisons.Count.ToString(CultureInfo.InvariantCulture)} carrier plan-design {(analysis.PlanDesignComparisons.Count == 1 ? "row" : "rows")}",
            UnderwritingComparisonCountText = $"{analysis.UnderwritingComparisons.Count.ToString(CultureInfo.InvariantCulture)} underwriting {(analysis.UnderwritingComparisons.Count == 1 ? "row" : "rows")}",
            ReviewItemCountText = $"{analysis.Findings.Count.ToString(CultureInfo.InvariantCulture)} {(analysis.Findings.Count == 1 ? "finding" : "findings")}",
            LowestAnnualPremiumText = FormatMoney(analysis.LowestAnnualPremium),
            AnnualPremiumSpreadText = FormatPremiumSpread(analysis),
            ReadinessStatus = analysis.ReadinessStatus,
            ReadinessMessage = analysis.ReadinessMessage,
            ProfileVersion = analysis.ProfileVersion,
            AnalysisVersion = analysis.AnalysisVersion,
            BuiltAtText = analysis.BuiltAtUtc.ToLocalTime().ToString("g", CultureInfo.CurrentCulture),
            IsEmpty = quoteCount == 0 && currentRowCount == 0,
            StatusText = BuildStatusText(analysis, currentVsQuote)
        };

        MapProposalSummaries(analysis, model);
        MapPremiumComparisons(analysis, model);
        MapRateComparisons(analysis, model);
        MapCarrierPlanDesignComparisons(analysis, model);
        MapUnderwritingComparisons(analysis, model);
        MapCurrentVsQuote(currentVsQuote, model);
        MapFindings(analysis, currentVsQuote, model);
        model.ReviewItemCountText = $"{model.Issues.Count.ToString(CultureInfo.InvariantCulture)} {(model.Issues.Count == 1 ? "finding" : "findings")}";
        return model;
    }

    private static void MapProposalSummaries(
        ProposalIntelligenceAnalysis analysis,
        ProposalComparisonViewModel model)
    {
        foreach (var quote in analysis.Quotes)
        {
            model.ProposalSummaries.Add(new ProposalQuoteSummaryDto
            {
                ProposalId = quote.ProposalId,
                CarrierName = quote.CarrierName,
                DisplayName = quote.DisplayName,
                AnnualPremiumText = FormatMoney(quote.AnnualPremium),
                MonthlyPremiumText = FormatMoney(quote.MonthlyPremium),
                DifferenceFromLowestText = FormatDifference(quote),
                ReadinessStatus = quote.ReadinessStatus,
                FinancialStatusText = quote.HasFinancialFacts ? "Available" : "Missing",
                PlanDesignStatusText = quote.HasPlanDesignFacts ? "Available" : "Missing",
                ValidationStatusText = quote.HasBookletValidation
                    ? quote.ValidationMismatchCount > 0 ? "Needs Review" : "Validated"
                    : "Missing",
                ReviewStatusText = quote.ReviewIssueCount == 0
                    ? "No review items"
                    : $"{quote.ReviewIssueCount.ToString(CultureInfo.InvariantCulture)} review {(quote.ReviewIssueCount == 1 ? "item" : "items")}",
                SourceText = BuildEvidenceText(quote.Evidence),
                IsLowestAnnualPremium = quote.IsLowestAnnualPremium
            });
        }
    }

    private static void MapPremiumComparisons(
        ProposalIntelligenceAnalysis analysis,
        ProposalComparisonViewModel model)
    {
        foreach (var row in analysis.PremiumComparisons)
        {
            var viewRow = new ProposalCanonicalComparisonRowDto
            {
                Category = "Premium",
                Subject = row.BenefitLine,
                Detail = JoinParts(row.ClassName, row.Tier),
                ComparisonStatus = row.ComparisonStatus,
                MissingText = FormatMissing(row.MissingProposalCount)
            };
            foreach (var value in row.Values)
            {
                viewRow.Values.Add(new ProposalCanonicalComparisonValueDto
                {
                    ProposalId = value.ProposalId,
                    CarrierName = value.QuoteDisplayName,
                    Value = JoinParts(
                        FormatLabelledMoney("Monthly", value.MonthlyPremium),
                        FormatLabelledMoney("Annual", value.AnnualPremium),
                        FormatLabelledMoney("Total", value.TotalCost)),
                    Detail = JoinParts(
                        FormatLabelledMoney("Employer", value.EmployerCost),
                        FormatLabelledMoney("Employee", value.EmployeeCost)),
                    SourceText = BuildEvidenceText(value.Evidence)
                });
            }

            model.PremiumComparisons.Add(viewRow);
        }
    }

    private static void MapRateComparisons(
        ProposalIntelligenceAnalysis analysis,
        ProposalComparisonViewModel model)
    {
        foreach (var row in analysis.RateComparisons)
        {
            var dimensions = JoinParts(row.ClassName, row.Tier, row.AgeBand, row.Gender, row.Province, row.RateBasis);
            var viewRow = new ProposalCanonicalComparisonRowDto
            {
                Category = "Rate",
                Subject = row.BenefitLine,
                Detail = dimensions,
                ComparisonStatus = row.ComparisonStatus,
                MissingText = FormatMissing(row.MissingProposalCount)
            };
            foreach (var value in row.Values)
            {
                viewRow.Values.Add(new ProposalCanonicalComparisonValueDto
                {
                    ProposalId = value.ProposalId,
                    CarrierName = value.QuoteDisplayName,
                    Value = value.Rate?.ToString("0.####", CultureInfo.InvariantCulture) ?? "Not available",
                    Detail = JoinParts(
                        FormatLabelledNumber("Lives", value.Lives),
                        FormatLabelledNumber("Volume", value.Volume),
                        FormatLabelledMoney("Annual", value.AnnualPremium)),
                    SourceText = BuildEvidenceText(value.Evidence)
                });
            }

            model.RateComparisons.Add(viewRow);
        }
    }

    private static void MapCarrierPlanDesignComparisons(
        ProposalIntelligenceAnalysis analysis,
        ProposalComparisonViewModel model)
    {
        foreach (var row in analysis.PlanDesignComparisons)
        {
            var viewRow = new ProposalCanonicalComparisonRowDto
            {
                Category = "Plan Design",
                Subject = row.BenefitLine,
                Detail = row.ProvisionName,
                ComparisonStatus = row.ComparisonStatus,
                MissingText = FormatMissing(row.MissingProposalCount)
            };
            foreach (var value in row.Values)
            {
                viewRow.Values.Add(new ProposalCanonicalComparisonValueDto
                {
                    ProposalId = value.ProposalId,
                    CarrierName = value.QuoteDisplayName,
                    Value = FirstNonBlank(value.ProposalValue, "Not available"),
                    Detail = JoinParts(
                        string.IsNullOrWhiteSpace(value.BookletValue) ? "" : $"Booklet: {value.BookletValue}",
                        value.ValidationStatus,
                        value.Variance),
                    SourceText = JoinParts(
                        BuildEvidenceText(value.Evidence),
                        string.IsNullOrWhiteSpace(value.BookletEvidence.SourceFileName)
                            ? ""
                            : $"Booklet: {BuildEvidenceText(value.BookletEvidence)}")
                });
            }

            model.CarrierPlanDesignComparisons.Add(viewRow);
        }
    }

    private static void MapUnderwritingComparisons(
        ProposalIntelligenceAnalysis analysis,
        ProposalComparisonViewModel model)
    {
        foreach (var row in analysis.UnderwritingComparisons)
        {
            var viewRow = new ProposalCanonicalComparisonRowDto
            {
                Category = row.Category,
                Subject = row.TermName,
                Detail = row.BenefitLine,
                ComparisonStatus = row.ComparisonStatus,
                MissingText = FormatMissing(row.MissingProposalCount)
            };
            foreach (var value in row.Values)
            {
                viewRow.Values.Add(new ProposalCanonicalComparisonValueDto
                {
                    ProposalId = value.ProposalId,
                    CarrierName = value.QuoteDisplayName,
                    Value = value.Value,
                    SourceText = BuildEvidenceText(value.Evidence)
                });
            }

            model.UnderwritingComparisons.Add(viewRow);
        }
    }

    private static void MapCurrentVsQuote(
        CurrentVsQuotePlanDesignComparison comparison,
        ProposalComparisonViewModel model)
    {
        foreach (var row in comparison.Rows)
        {
            var viewRow = new BenefitComparisonDto
            {
                BenefitCategory = row.BenefitCategory,
                BenefitName = row.BenefitName,
                ProvisionName = row.ProvisionName,
                ClassName = row.ClassName,
                CurrentValue = row.CurrentValue,
                CurrentSourceText = BuildEvidenceText(row.CurrentEvidence),
                ComparisonStatus = row.Status
            };

            foreach (var quoteValue in row.QuoteValues)
            {
                var quote = comparison.Quotes.First(item => item.QuoteKey == quoteValue.QuoteKey);
                viewRow.QuoteValues.Add(new QuoteComparisonValueDto
                {
                    CarrierName = quoteValue.QuoteDisplayName,
                    Reference = quote.Reference,
                    Value = quoteValue.Value,
                    Status = quoteValue.Status,
                    SourceText = BuildEvidenceText(quoteValue.Evidence)
                });
            }

            model.BenefitComparisons.Add(viewRow);
        }
    }

    private static void MapFindings(
        ProposalIntelligenceAnalysis analysis,
        CurrentVsQuotePlanDesignComparison currentVsQuote,
        ProposalComparisonViewModel model)
    {
        foreach (var finding in analysis.Findings)
        {
            AddIssue(model, new PlanDesignComparisonIssueDto
            {
                Code = finding.Code,
                Severity = finding.Severity,
                Category = finding.Category,
                Carrier = finding.Carrier,
                Message = finding.Message,
                Context = JoinParts(finding.Carrier, finding.BenefitLine, finding.FieldName)
            });
        }

        foreach (var issue in currentVsQuote.Issues)
        {
            AddIssue(model, new PlanDesignComparisonIssueDto
            {
                Code = issue.Code,
                Severity = issue.Severity,
                Category = ProposalIntelligenceFindingCategories.PlanDesign,
                Message = issue.Message,
                Context = JoinParts(issue.Benefit, issue.Provision)
            });
        }
    }

    private static void AddIssue(ProposalComparisonViewModel model, PlanDesignComparisonIssueDto issue)
    {
        if (!model.Issues.Any(existing =>
                existing.Code.Equals(issue.Code, StringComparison.OrdinalIgnoreCase) &&
                existing.Message.Equals(issue.Message, StringComparison.OrdinalIgnoreCase) &&
                existing.Context.Equals(issue.Context, StringComparison.OrdinalIgnoreCase)))
        {
            model.Issues.Add(issue);
        }
    }

    private static string BuildStatusText(
        ProposalIntelligenceAnalysis analysis,
        CurrentVsQuotePlanDesignComparison currentVsQuote)
    {
        if (analysis.Quotes.Count == 0 && currentVsQuote.Rows.Count == 0)
        {
            return "No saved proposal or current-plan comparison facts were found for the selected client.";
        }

        if (analysis.Quotes.Count == 0)
        {
            return "The current plan is available, but no saved carrier proposals are ready for Proposal Intelligence analysis.";
        }

        return currentVsQuote.HasCurrentPlan
            ? "Loaded canonical proposal analysis and the saved Current-vs-Quote plan-design comparison."
            : "Loaded canonical proposal analysis. Current-plan facts are not yet available for Current-vs-Quote comparison.";
    }

    private static string FormatPremiumSpread(ProposalIntelligenceAnalysis analysis)
    {
        if (!analysis.AnnualPremiumSpread.HasValue)
        {
            return "N/A";
        }

        var percent = analysis.AnnualPremiumSpreadPercent.HasValue
            ? $" ({analysis.AnnualPremiumSpreadPercent.Value.ToString("0.##", CultureInfo.InvariantCulture)}%)"
            : "";
        return $"{FormatMoney(analysis.AnnualPremiumSpread)}{percent}";
    }

    private static string FormatDifference(ProposalIntelligenceQuoteAnalysis quote)
    {
        if (!quote.AnnualPremiumDifferenceFromLowest.HasValue)
        {
            return "N/A";
        }

        var percent = quote.AnnualPremiumDifferencePercentFromLowest.HasValue
            ? $" ({quote.AnnualPremiumDifferencePercentFromLowest.Value.ToString("0.##", CultureInfo.InvariantCulture)}%)"
            : "";
        return $"{FormatMoney(quote.AnnualPremiumDifferenceFromLowest)}{percent}";
    }

    private static string FormatMoney(decimal? value)
    {
        return value?.ToString("$#,0.##", CultureInfo.InvariantCulture) ?? "N/A";
    }

    private static string FormatLabelledMoney(string label, decimal? value)
    {
        return value.HasValue ? $"{label}: {FormatMoney(value)}" : "";
    }

    private static string FormatLabelledNumber(string label, decimal? value)
    {
        return value.HasValue
            ? $"{label}: {value.Value.ToString("#,0.##", CultureInfo.InvariantCulture)}"
            : "";
    }

    private static string FormatMissing(int missingProposalCount)
    {
        return missingProposalCount == 0
            ? "Complete"
            : $"{missingProposalCount.ToString(CultureInfo.InvariantCulture)} missing";
    }

    private static string BuildEvidenceText(PlanDesignCanonicalEvidence? evidence)
    {
        if (evidence is null)
        {
            return "No source evidence";
        }

        var location = evidence.SourcePage.HasValue
            ? $"page {evidence.SourcePage.Value.ToString(CultureInfo.InvariantCulture)}"
            : evidence.SourceSheet;
        var confidence = evidence.ConfidenceScore.HasValue
            ? $"{evidence.ConfidenceScore.Value:P0} confidence"
            : "";
        return EvidenceText(
            FirstNonBlank(evidence.SourceFileName, evidence.SourceDocument, evidence.SourceTable),
            location,
            confidence);
    }

    private static string BuildEvidenceText(ProposalCanonicalEvidence evidence)
    {
        var location = evidence.SourcePage.HasValue
            ? $"page {evidence.SourcePage.Value.ToString(CultureInfo.InvariantCulture)}"
            : evidence.SourceSheet;
        var confidence = evidence.ConfidenceScore.HasValue
            ? $"{evidence.ConfidenceScore.Value:P0} confidence"
            : "";
        return EvidenceText(
            FirstNonBlank(evidence.SourceFileName, evidence.SourceTable),
            location,
            confidence);
    }

    private static string EvidenceText(params string?[] values)
    {
        var text = JoinParts(values);
        return string.IsNullOrWhiteSpace(text) ? "No source evidence" : text;
    }

    private static string JoinParts(params string?[] values)
    {
        return string.Join(" | ", values.Where(value => !string.IsNullOrWhiteSpace(value)).Select(value => value!.Trim()));
    }

    private static string FirstNonBlank(params string?[] values)
    {
        return values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value))?.Trim() ?? "";
    }
}

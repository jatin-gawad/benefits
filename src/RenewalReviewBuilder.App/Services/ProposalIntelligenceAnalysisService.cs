using System.Globalization;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed class ProposalIntelligenceAnalysisService
{
    private readonly ProposalCanonicalProfileService _profileService;

    public ProposalIntelligenceAnalysisService(string? databasePath = null)
        : this(new ProposalCanonicalProfileService(databasePath))
    {
    }

    public ProposalIntelligenceAnalysisService(ProposalCanonicalProfileService profileService)
    {
        _profileService = profileService ?? throw new ArgumentNullException(nameof(profileService));
    }

    public async Task<ProposalIntelligenceAnalysis> BuildAsync(long clientId)
    {
        var profile = await _profileService.BuildAsync(clientId);
        return Build(profile);
    }

    public ProposalIntelligenceAnalysis Build(ProposalCanonicalProfile profile)
    {
        ArgumentNullException.ThrowIfNull(profile);

        var result = new ProposalIntelligenceAnalysis
        {
            AnalysisVersion = ProposalIntelligenceAnalysisVersions.Current,
            ProfileVersion = profile.ProfileVersion,
            ClientId = profile.ClientId,
            ClientName = profile.ClientName,
            ClientNumber = profile.ClientNumber,
            BuiltAtUtc = DateTime.UtcNow,
            CanCompareCarriers = profile.CanCompareCarriers,
            CanCompareFinancials = profile.CanCompareFinancials,
            CanComparePlanDesign = profile.CanComparePlanDesign
        };
        var quoteContexts = profile.Proposals
            .Select((proposal, index) => BuildQuoteContext(proposal, index))
            .ToList();
        foreach (var context in quoteContexts)
        {
            context.ExpectedProposalCount = quoteContexts.Count;
        }

        result.Quotes.AddRange(quoteContexts.Select(context => context.Analysis));

        ApplyAnnualPremiumBaseline(result.Quotes);
        BuildRateComparisons(quoteContexts, result);
        BuildPremiumComparisons(quoteContexts, result);
        BuildPlanDesignComparisons(quoteContexts, result);
        BuildUnderwritingComparisons(quoteContexts, result);
        AddProfileFindings(profile, result);
        AddDerivedReadinessFindings(profile, result);
        AddComparisonFindings(result);
        ResolveReadiness(profile, result);
        SortOutput(result);
        return result;
    }

    private static QuoteContext BuildQuoteContext(ProposalCanonicalQuote proposal, int index)
    {
        var carrier = FirstNonBlank(proposal.CarrierName, proposal.DetectedCarrier, "Unknown Carrier");
        var quoteKey = proposal.ProposalId > 0
            ? $"proposal:{proposal.ProposalId.ToString(CultureInfo.InvariantCulture)}"
            : !string.IsNullOrWhiteSpace(proposal.SourceFileHash)
                ? $"source:{proposal.SourceFileHash.Trim().ToLowerInvariant()}"
                : $"quote:{index.ToString(CultureInfo.InvariantCulture)}";
        var displayParts = new List<string> { carrier };
        if (proposal.ProposalDate.HasValue)
        {
            displayParts.Add(proposal.ProposalDate.Value.ToString("MMM d, yyyy", CultureInfo.InvariantCulture));
        }

        if (proposal.ProposalId > 0)
        {
            displayParts.Add($"#{proposal.ProposalId.ToString(CultureInfo.InvariantCulture)}");
        }

        var analysis = new ProposalIntelligenceQuoteAnalysis
        {
            QuoteKey = quoteKey,
            ProposalId = proposal.ProposalId,
            CarrierId = proposal.CarrierId,
            CarrierKey = proposal.CarrierKey,
            CarrierName = carrier,
            DisplayName = string.Join(" | ", displayParts),
            ProposalDate = proposal.ProposalDate,
            EffectiveDate = proposal.EffectiveDate,
            Status = proposal.Status,
            ReadinessStatus = proposal.ReadinessStatus,
            MonthlyPremium = proposal.TotalMonthlyPremium,
            AnnualPremium = proposal.TotalAnnualPremium ?? proposal.TotalCost,
            EmployerCost = proposal.TotalEmployerCost,
            EmployeeCost = proposal.TotalEmployeeCost,
            TotalCost = proposal.TotalCost,
            HasCoreFacts = proposal.HasCoreFacts,
            HasFinancialFacts = proposal.HasFinancialFacts,
            HasPlanDesignFacts = proposal.HasPlanDesignFacts,
            HasBookletValidation = proposal.HasBookletValidation,
            HasSourceEvidence = proposal.HasSourceEvidence,
            RateCount = proposal.Rates.Count,
            PremiumCount = proposal.Premiums.Count,
            PlanDesignFactCount = proposal.PlanDesignFacts.Count,
            ValidationMismatchCount = proposal.ValidationMismatchCount,
            UnderwritingTermCount = proposal.UnderwritingAssumptions.Count + proposal.PoolingRetention.Count,
            ReviewIssueCount = proposal.ReviewIssueCount,
            SourceFileName = proposal.SourceFileName,
            SourceFileHash = proposal.SourceFileHash,
            ExtractionConfidence = proposal.ExtractionConfidence,
            ExtractionMethod = proposal.ExtractionMethod,
            Evidence = proposal.Evidence
        };
        return new QuoteContext(proposal, analysis);
    }

    private static void ApplyAnnualPremiumBaseline(IReadOnlyList<ProposalIntelligenceQuoteAnalysis> quotes)
    {
        var premiums = quotes
            .Where(quote => quote.AnnualPremium.HasValue)
            .Select(quote => quote.AnnualPremium!.Value)
            .ToList();
        if (premiums.Count == 0)
        {
            return;
        }

        var lowest = premiums.Min();
        foreach (var quote in quotes.Where(quote => quote.AnnualPremium.HasValue))
        {
            var difference = quote.AnnualPremium!.Value - lowest;
            quote.AnnualPremiumDifferenceFromLowest = difference;
            quote.AnnualPremiumDifferencePercentFromLowest = lowest > 0m
                ? difference / lowest * 100m
                : null;
            quote.IsLowestAnnualPremium = quote.AnnualPremium.Value == lowest;
        }
    }

    private static void BuildRateComparisons(
        IReadOnlyList<QuoteContext> contexts,
        ProposalIntelligenceAnalysis result)
    {
        var rows = new Dictionary<string, ProposalIntelligenceRateComparisonRow>(StringComparer.OrdinalIgnoreCase);
        foreach (var context in contexts)
        {
            if (context.Proposal.Rates.Count > 0)
            {
                foreach (var rate in context.Proposal.Rates)
                {
                    AddRateValue(rows, context, rate);
                }
            }
            else
            {
                foreach (var benefit in context.Proposal.BenefitLines.Where(row => row.Rate.HasValue))
                {
                    AddRateValue(rows, context, new ProposalCanonicalRate
                    {
                        BenefitLine = benefit.BenefitLine,
                        ClassName = benefit.ClassName,
                        Tier = benefit.Tier,
                        Lives = benefit.Lives,
                        Volume = benefit.Volume,
                        Rate = benefit.Rate,
                        MonthlyPremium = benefit.MonthlyPremium,
                        AnnualPremium = benefit.AnnualPremium,
                        Evidence = benefit.Evidence
                    });
                }
            }
        }

        foreach (var row in rows.Values)
        {
            var availableProposalCount = row.Values
                .Where(value => value.Rate.HasValue)
                .Select(value => value.ProposalId)
                .Distinct()
                .Count();
            row.MissingProposalCount = Math.Max(0, row.ExpectedProposalCount - availableProposalCount);
            row.HasRateDifference = DistinctDecimalCount(row.Values.Select(value => value.Rate)) > 1;
            row.HasPremiumDifference = DistinctSignatureCount(row.Values.Select(value =>
                FinancialSignature(value.MonthlyPremium, value.AnnualPremium))) > 1;
            row.ComparisonStatus = ResolveComparisonStatus(
                row.ExpectedProposalCount,
                availableProposalCount,
                row.HasRateDifference || row.HasPremiumDifference);
            result.RateComparisons.Add(row);
        }
    }

    private static void AddRateValue(
        IDictionary<string, ProposalIntelligenceRateComparisonRow> rows,
        QuoteContext context,
        ProposalCanonicalRate rate)
    {
        var benefitLine = FirstNonBlank(rate.BenefitLine, "Other");
        var className = FirstNonBlank(rate.ClassName, "All Classes");
        var tier = rate.Tier.Trim();
        var comparisonKey = BuildComparisonKey(
            "rate",
            benefitLine,
            className,
            tier,
            rate.AgeBand,
            rate.Gender,
            rate.Province,
            rate.RateBasis);
        if (!rows.TryGetValue(comparisonKey, out var row))
        {
            row = new ProposalIntelligenceRateComparisonRow
            {
                ComparisonKey = comparisonKey,
                BenefitLine = benefitLine,
                ClassName = className,
                Tier = tier,
                AgeBand = rate.AgeBand.Trim(),
                Gender = rate.Gender.Trim(),
                Province = rate.Province.Trim(),
                RateBasis = rate.RateBasis.Trim(),
                ExpectedProposalCount = context.ExpectedProposalCount
            };
            rows[comparisonKey] = row;
        }

        row.Values.Add(new ProposalIntelligenceRateValue
        {
            QuoteKey = context.Analysis.QuoteKey,
            ProposalId = context.Analysis.ProposalId,
            CarrierName = context.Analysis.CarrierName,
            QuoteDisplayName = context.Analysis.DisplayName,
            Lives = rate.Lives,
            Volume = rate.Volume,
            Rate = rate.Rate,
            MonthlyPremium = rate.MonthlyPremium,
            AnnualPremium = rate.AnnualPremium,
            Evidence = rate.Evidence
        });
    }

    private static void BuildPremiumComparisons(
        IReadOnlyList<QuoteContext> contexts,
        ProposalIntelligenceAnalysis result)
    {
        var rows = new Dictionary<string, ProposalIntelligencePremiumComparisonRow>(StringComparer.OrdinalIgnoreCase);
        foreach (var context in contexts)
        {
            if (context.Proposal.Premiums.Count > 0)
            {
                foreach (var premium in context.Proposal.Premiums)
                {
                    AddPremiumValue(rows, context, premium);
                }
            }
            else
            {
                foreach (var benefit in context.Proposal.BenefitLines.Where(row =>
                             row.MonthlyPremium.HasValue || row.AnnualPremium.HasValue))
                {
                    AddPremiumValue(rows, context, new ProposalCanonicalPremium
                    {
                        BenefitLine = benefit.BenefitLine,
                        ClassName = benefit.ClassName,
                        Tier = benefit.Tier,
                        MonthlyPremium = benefit.MonthlyPremium,
                        AnnualPremium = benefit.AnnualPremium,
                        Evidence = benefit.Evidence
                    });
                }
            }
        }

        foreach (var row in rows.Values)
        {
            var availableProposalCount = row.Values
                .Where(value => value.MonthlyPremium.HasValue || value.AnnualPremium.HasValue || value.TotalCost.HasValue)
                .Select(value => value.ProposalId)
                .Distinct()
                .Count();
            row.MissingProposalCount = Math.Max(0, row.ExpectedProposalCount - availableProposalCount);
            row.HasPremiumDifference = DistinctSignatureCount(row.Values.Select(value =>
                FinancialSignature(value.MonthlyPremium, value.AnnualPremium, value.TotalCost))) > 1;
            row.HasCostShareDifference = DistinctSignatureCount(row.Values
                .Where(value => value.EmployerCost.HasValue || value.EmployeeCost.HasValue)
                .Select(value => FinancialSignature(value.EmployerCost, value.EmployeeCost))) > 1;
            row.ComparisonStatus = ResolveComparisonStatus(
                row.ExpectedProposalCount,
                availableProposalCount,
                row.HasPremiumDifference || row.HasCostShareDifference);
            result.PremiumComparisons.Add(row);
        }
    }

    private static void AddPremiumValue(
        IDictionary<string, ProposalIntelligencePremiumComparisonRow> rows,
        QuoteContext context,
        ProposalCanonicalPremium premium)
    {
        var benefitLine = FirstNonBlank(premium.BenefitLine, "Other");
        var className = FirstNonBlank(premium.ClassName, "All Classes");
        var tier = premium.Tier.Trim();
        var comparisonKey = BuildComparisonKey("premium", benefitLine, className, tier);
        if (!rows.TryGetValue(comparisonKey, out var row))
        {
            row = new ProposalIntelligencePremiumComparisonRow
            {
                ComparisonKey = comparisonKey,
                BenefitLine = benefitLine,
                ClassName = className,
                Tier = tier,
                ExpectedProposalCount = context.ExpectedProposalCount
            };
            rows[comparisonKey] = row;
        }

        row.Values.Add(new ProposalIntelligencePremiumValue
        {
            QuoteKey = context.Analysis.QuoteKey,
            ProposalId = context.Analysis.ProposalId,
            CarrierName = context.Analysis.CarrierName,
            QuoteDisplayName = context.Analysis.DisplayName,
            MonthlyPremium = premium.MonthlyPremium,
            AnnualPremium = premium.AnnualPremium,
            EmployerCost = premium.EmployerCost,
            EmployeeCost = premium.EmployeeCost,
            TotalCost = premium.TotalCost,
            Evidence = premium.Evidence
        });
    }

    private static void BuildPlanDesignComparisons(
        IReadOnlyList<QuoteContext> contexts,
        ProposalIntelligenceAnalysis result)
    {
        var rows = new Dictionary<string, ProposalIntelligencePlanDesignComparisonRow>(StringComparer.OrdinalIgnoreCase);
        foreach (var context in contexts)
        {
            var planGroups = context.Proposal.PlanDesignFacts
                .GroupBy(BuildPlanDesignKey, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.ToList(), StringComparer.OrdinalIgnoreCase);
            var validationGroups = context.Proposal.ValidationFacts
                .GroupBy(BuildPlanDesignKey, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.ToList(), StringComparer.OrdinalIgnoreCase);
            var comparisonKeys = planGroups.Keys
                .Concat(validationGroups.Keys)
                .Distinct(StringComparer.OrdinalIgnoreCase);

            foreach (var comparisonKey in comparisonKeys)
            {
                var planFacts = planGroups.GetValueOrDefault(comparisonKey) ?? [];
                var validationFacts = validationGroups.GetValueOrDefault(comparisonKey) ?? [];
                var representative = planFacts.Concat(validationFacts).First();
                if (!rows.TryGetValue(comparisonKey, out var row))
                {
                    row = new ProposalIntelligencePlanDesignComparisonRow
                    {
                        ComparisonKey = comparisonKey,
                        BenefitLine = FirstNonBlank(representative.BenefitLine, "Other"),
                        ProvisionName = FirstNonBlank(representative.ProvisionName, "Plan Design"),
                        ExpectedProposalCount = context.ExpectedProposalCount
                    };
                    rows[comparisonKey] = row;
                }

                var proposalValue = JoinDistinct(planFacts.Select(fact => fact.ProposalValue)
                    .Concat(validationFacts.Select(fact => fact.ProposalValue)));
                var bookletValue = JoinDistinct(validationFacts.Select(fact => fact.BookletValue));
                var validationStatus = JoinDistinct(validationFacts.Select(fact => fact.ValidationStatus));
                var variance = JoinDistinct(validationFacts.Select(fact => fact.Variance));
                var evidence = SelectEvidence(planFacts.Select(fact => fact.Evidence)
                    .Concat(validationFacts.Select(fact => fact.Evidence)));
                var bookletEvidence = SelectEvidence(validationFacts.Select(fact => fact.BookletEvidence));
                row.Values.Add(new ProposalIntelligencePlanDesignValue
                {
                    QuoteKey = context.Analysis.QuoteKey,
                    ProposalId = context.Analysis.ProposalId,
                    CarrierName = context.Analysis.CarrierName,
                    QuoteDisplayName = context.Analysis.DisplayName,
                    ProposalValue = proposalValue,
                    BookletValue = bookletValue,
                    ValidationStatus = validationStatus,
                    Variance = variance,
                    IsMismatch = validationFacts.Any(fact => fact.IsMismatch),
                    Evidence = evidence,
                    BookletEvidence = bookletEvidence
                });
            }
        }

        foreach (var row in rows.Values)
        {
            var availableProposalCount = row.Values
                .Where(value => !string.IsNullOrWhiteSpace(value.ProposalValue))
                .Select(value => value.ProposalId)
                .Distinct()
                .Count();
            row.MissingProposalCount = Math.Max(0, row.ExpectedProposalCount - availableProposalCount);
            row.HasProposalDifference = DistinctTextCount(row.Values.Select(value => value.ProposalValue)) > 1;
            row.HasBookletMismatch = row.Values.Any(value => value.IsMismatch);
            row.ComparisonStatus = ResolveComparisonStatus(
                row.ExpectedProposalCount,
                availableProposalCount,
                row.HasProposalDifference,
                row.HasBookletMismatch);
            result.PlanDesignComparisons.Add(row);
        }
    }

    private static void BuildUnderwritingComparisons(
        IReadOnlyList<QuoteContext> contexts,
        ProposalIntelligenceAnalysis result)
    {
        var rows = new Dictionary<string, ProposalIntelligenceTermComparisonRow>(StringComparer.OrdinalIgnoreCase);
        foreach (var context in contexts)
        {
            foreach (var assumption in context.Proposal.UnderwritingAssumptions)
            {
                AddTermValue(
                    rows,
                    context,
                    "Underwriting Assumption",
                    "",
                    FirstNonBlank(assumption.Name, "Assumption"),
                    assumption.Value,
                    assumption.Evidence);
            }

            foreach (var pooling in context.Proposal.PoolingRetention)
            {
                AddTermValue(rows, context, "Pooling / Retention", pooling.BenefitLine, "Pooling Level", pooling.PoolingLevel, pooling.Evidence);
                AddTermValue(rows, context, "Pooling / Retention", pooling.BenefitLine, "Pooling Charge", FormatDecimal(pooling.PoolingCharge), pooling.Evidence);
                AddTermValue(rows, context, "Pooling / Retention", pooling.BenefitLine, "Commission", FormatDecimal(pooling.Commission), pooling.Evidence);
                AddTermValue(rows, context, "Pooling / Retention", pooling.BenefitLine, "Retention", FormatDecimal(pooling.Retention), pooling.Evidence);
                AddTermValue(rows, context, "Pooling / Retention", pooling.BenefitLine, "Target Loss Ratio", FormatDecimal(pooling.TargetLossRatio), pooling.Evidence);
                AddTermValue(rows, context, "Pooling / Retention", pooling.BenefitLine, "Rate Guarantee", pooling.RateGuarantee, pooling.Evidence);
                AddTermValue(rows, context, "Pooling / Retention", pooling.BenefitLine, "Pooled Benefits", pooling.PooledBenefits, pooling.Evidence);
                AddTermValue(rows, context, "Pooling / Retention", pooling.BenefitLine, "Experience Rated Benefits", pooling.ExperienceRatedBenefits, pooling.Evidence);
                AddTermValue(rows, context, "Pooling / Retention", pooling.BenefitLine, "IBNR", pooling.Ibnr, pooling.Evidence);
                AddTermValue(rows, context, "Pooling / Retention", pooling.BenefitLine, "Rate Caps", pooling.RateCaps, pooling.Evidence);
                AddTermValue(rows, context, "Pooling / Retention", pooling.BenefitLine, "First Renewal Cap", pooling.FirstRenewalCap, pooling.Evidence);
                AddTermValue(rows, context, "Pooling / Retention", pooling.BenefitLine, "Second Renewal Cap", pooling.SecondRenewalCap, pooling.Evidence);
            }
        }

        foreach (var row in rows.Values)
        {
            var availableProposalCount = row.Values
                .Where(value => !string.IsNullOrWhiteSpace(value.Value))
                .Select(value => value.ProposalId)
                .Distinct()
                .Count();
            row.MissingProposalCount = Math.Max(0, row.ExpectedProposalCount - availableProposalCount);
            row.HasDifference = DistinctTextCount(row.Values.Select(value => value.Value)) > 1;
            row.ComparisonStatus = ResolveComparisonStatus(
                row.ExpectedProposalCount,
                availableProposalCount,
                row.HasDifference);
            result.UnderwritingComparisons.Add(row);
        }
    }

    private static void AddTermValue(
        IDictionary<string, ProposalIntelligenceTermComparisonRow> rows,
        QuoteContext context,
        string category,
        string benefitLine,
        string termName,
        string value,
        ProposalCanonicalEvidence evidence)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return;
        }

        var normalizedBenefitLine = FirstNonBlank(benefitLine, "All Benefits");
        var comparisonKey = BuildComparisonKey("term", category, normalizedBenefitLine, termName);
        if (!rows.TryGetValue(comparisonKey, out var row))
        {
            row = new ProposalIntelligenceTermComparisonRow
            {
                ComparisonKey = comparisonKey,
                Category = category,
                BenefitLine = normalizedBenefitLine,
                TermName = termName,
                ExpectedProposalCount = context.ExpectedProposalCount
            };
            rows[comparisonKey] = row;
        }

        row.Values.Add(new ProposalIntelligenceTermValue
        {
            QuoteKey = context.Analysis.QuoteKey,
            ProposalId = context.Analysis.ProposalId,
            CarrierName = context.Analysis.CarrierName,
            QuoteDisplayName = context.Analysis.DisplayName,
            Value = value.Trim(),
            Evidence = evidence
        });
    }

    private static void AddProfileFindings(
        ProposalCanonicalProfile profile,
        ProposalIntelligenceAnalysis result)
    {
        foreach (var issue in profile.Issues)
        {
            AddFinding(result, new ProposalIntelligenceFinding
            {
                Code = issue.Code,
                Severity = issue.Severity,
                Category = ResolveFindingCategory(issue.Code, issue.SourceTable),
                Message = issue.Message,
                ProposalId = issue.ProposalId,
                Carrier = issue.Carrier,
                SourceTable = issue.SourceTable,
                SourceId = issue.SourceId
            });
        }
    }

    private static void AddDerivedReadinessFindings(
        ProposalCanonicalProfile profile,
        ProposalIntelligenceAnalysis result)
    {
        if (profile.Proposals.Count == 0)
        {
            AddFinding(result, new ProposalIntelligenceFinding
            {
                Code = "proposals_missing",
                Severity = ProposalCanonicalIssueSeverities.Error,
                Category = ProposalIntelligenceFindingCategories.Readiness,
                Message = "No saved proposals were found for the selected client."
            });
            return;
        }

        foreach (var proposal in profile.Proposals)
        {
            var carrier = FirstNonBlank(proposal.CarrierName, proposal.DetectedCarrier, "Unknown Carrier");
            if (!proposal.HasCoreFacts)
            {
                AddFinding(result, new ProposalIntelligenceFinding
                {
                    Code = "proposal_facts_missing",
                    Severity = ProposalCanonicalIssueSeverities.Error,
                    Category = ProposalIntelligenceFindingCategories.Readiness,
                    Message = "The saved proposal has no usable normalized facts.",
                    ProposalId = proposal.ProposalId,
                    Carrier = carrier,
                    SourceTable = "proposal_files",
                    Evidence = proposal.Evidence
                });
                continue;
            }

            if (!proposal.HasFinancialFacts)
            {
                AddFinding(result, MissingFinding(
                    proposal,
                    carrier,
                    "proposal_financial_facts_missing",
                    ProposalIntelligenceFindingCategories.Financial,
                    "No saved proposal rates or premiums were found.",
                    "proposal_rates / proposal_premiums"));
            }

            if (!proposal.HasPlanDesignFacts)
            {
                AddFinding(result, MissingFinding(
                    proposal,
                    carrier,
                    "proposal_plan_design_missing",
                    ProposalIntelligenceFindingCategories.PlanDesign,
                    "No saved proposal plan design facts were found.",
                    "proposal_plan_design"));
            }

            if (!proposal.HasBookletValidation)
            {
                AddFinding(result, MissingFinding(
                    proposal,
                    carrier,
                    "proposal_validation_missing",
                    ProposalIntelligenceFindingCategories.Validation,
                    "No proposal-to-booklet validation facts were found.",
                    "proposal_booklet_validation"));
            }

            if (!proposal.HasSourceEvidence)
            {
                AddFinding(result, MissingFinding(
                    proposal,
                    carrier,
                    "proposal_source_evidence_missing",
                    ProposalIntelligenceFindingCategories.Evidence,
                    "The proposal does not contain usable source-file evidence.",
                    "proposal_files"));
            }

            if (proposal.ValidationFacts.Any(fact => fact.IsMismatch) &&
                !result.Findings.Any(finding => finding.Code == "proposal_booklet_mismatch" && finding.ProposalId == proposal.ProposalId))
            {
                var mismatch = proposal.ValidationFacts.First(fact => fact.IsMismatch);
                AddFinding(result, new ProposalIntelligenceFinding
                {
                    Code = "proposal_booklet_mismatch",
                    Severity = ProposalCanonicalIssueSeverities.Warning,
                    Category = ProposalIntelligenceFindingCategories.Validation,
                    Message = "One or more saved proposal values do not match the booklet evidence.",
                    ProposalId = proposal.ProposalId,
                    Carrier = carrier,
                    BenefitLine = mismatch.BenefitLine,
                    FieldName = mismatch.ProvisionName,
                    SourceTable = "proposal_booklet_validation",
                    SourceId = mismatch.Evidence.SourceId,
                    Evidence = mismatch.Evidence
                });
            }

            if (proposal.Exceptions.Any(exception => exception.NeedsReview) &&
                !result.Findings.Any(finding => finding.Code == "proposal_exceptions_need_review" && finding.ProposalId == proposal.ProposalId))
            {
                var exception = proposal.Exceptions.First(row => row.NeedsReview);
                AddFinding(result, new ProposalIntelligenceFinding
                {
                    Code = "proposal_exceptions_need_review",
                    Severity = ProposalCanonicalIssueSeverities.Warning,
                    Category = ProposalIntelligenceFindingCategories.Validation,
                    Message = $"{proposal.Exceptions.Count(row => row.NeedsReview).ToString(CultureInfo.InvariantCulture)} saved proposal extraction or validation issue(s) need review.",
                    ProposalId = proposal.ProposalId,
                    Carrier = carrier,
                    BenefitLine = exception.BenefitLine,
                    FieldName = exception.FieldName,
                    SourceTable = "proposal_exceptions",
                    SourceId = exception.Evidence.SourceId,
                    Evidence = exception.Evidence
                });
            }
        }
    }

    private static ProposalIntelligenceFinding MissingFinding(
        ProposalCanonicalQuote proposal,
        string carrier,
        string code,
        string category,
        string message,
        string sourceTable)
    {
        return new ProposalIntelligenceFinding
        {
            Code = code,
            Severity = ProposalCanonicalIssueSeverities.Warning,
            Category = category,
            Message = message,
            ProposalId = proposal.ProposalId,
            Carrier = carrier,
            SourceTable = sourceTable,
            Evidence = proposal.Evidence
        };
    }

    private static void AddComparisonFindings(ProposalIntelligenceAnalysis result)
    {
        if (!result.CanCompareCarriers && result.HasProposals)
        {
            AddFinding(result, new ProposalIntelligenceFinding
            {
                Code = "carrier_comparison_unavailable",
                Severity = ProposalCanonicalIssueSeverities.Warning,
                Category = ProposalIntelligenceFindingCategories.Readiness,
                Message = "At least two proposals with usable normalized facts are required for carrier comparison."
            });
        }

        if (result.CanCompareCarriers && !result.CanCompareFinancials)
        {
            AddFinding(result, new ProposalIntelligenceFinding
            {
                Code = "financial_comparison_incomplete",
                Severity = ProposalCanonicalIssueSeverities.Warning,
                Category = ProposalIntelligenceFindingCategories.Financial,
                Message = "At least two saved proposals need usable rates or premiums for financial comparison."
            });
        }

        if (result.CanCompareCarriers && !result.CanComparePlanDesign)
        {
            AddFinding(result, new ProposalIntelligenceFinding
            {
                Code = "plan_design_comparison_incomplete",
                Severity = ProposalCanonicalIssueSeverities.Warning,
                Category = ProposalIntelligenceFindingCategories.PlanDesign,
                Message = "At least two saved proposals need plan design facts for carrier comparison."
            });
        }

        if (result.AnnualPremiumSpread is > 0m)
        {
            var percent = result.AnnualPremiumSpreadPercent.HasValue
                ? $" ({result.AnnualPremiumSpreadPercent.Value.ToString("0.##", CultureInfo.InvariantCulture)}%)"
                : "";
            AddFinding(result, new ProposalIntelligenceFinding
            {
                Code = "annual_premium_spread",
                Severity = ProposalCanonicalIssueSeverities.Information,
                Category = ProposalIntelligenceFindingCategories.Financial,
                Message = $"The saved annual proposal premium spread is {result.AnnualPremiumSpread.Value.ToString("$#,0.##", CultureInfo.InvariantCulture)}{percent}."
            });
        }

        AddDifferenceSummary(
            result,
            "rate_differences_found",
            ProposalIntelligenceFindingCategories.Financial,
            result.RateComparisons.Count(row => row.HasRateDifference),
            "rate comparison row(s) contain different saved rates.");
        AddDifferenceSummary(
            result,
            "premium_differences_found",
            ProposalIntelligenceFindingCategories.Financial,
            result.PremiumComparisons.Count(row => row.HasPremiumDifference || row.HasCostShareDifference),
            "premium comparison row(s) contain different saved premium or cost-share values.");
        AddDifferenceSummary(
            result,
            "plan_design_differences_found",
            ProposalIntelligenceFindingCategories.PlanDesign,
            result.PlanDesignComparisons.Count(row => row.HasProposalDifference),
            "plan design comparison row(s) contain different saved proposal values.");
        AddDifferenceSummary(
            result,
            "underwriting_term_differences_found",
            ProposalIntelligenceFindingCategories.Underwriting,
            result.UnderwritingComparisons.Count(row => row.HasDifference),
            "underwriting or pooling term row(s) contain different saved values.");
    }

    private static void AddDifferenceSummary(
        ProposalIntelligenceAnalysis result,
        string code,
        string category,
        int count,
        string message)
    {
        if (count == 0)
        {
            return;
        }

        AddFinding(result, new ProposalIntelligenceFinding
        {
            Code = code,
            Severity = ProposalCanonicalIssueSeverities.Information,
            Category = category,
            Message = $"{count.ToString(CultureInfo.InvariantCulture)} {message}"
        });
    }

    private static void ResolveReadiness(
        ProposalCanonicalProfile profile,
        ProposalIntelligenceAnalysis result)
    {
        if (!profile.CanSupportProposalIntelligence)
        {
            result.ReadinessStatus = ProposalCanonicalReadinessStatuses.NotReady;
            result.ReadinessMessage = "No usable saved proposal facts were found for this client.";
            return;
        }

        if (profile.ReadinessStatus == ProposalCanonicalReadinessStatuses.NeedsReview ||
            result.Quotes.Any(quote => quote.ReadinessStatus == ProposalCanonicalReadinessStatuses.NeedsReview))
        {
            result.ReadinessStatus = ProposalCanonicalReadinessStatuses.NeedsReview;
            result.ReadinessMessage = "Saved proposal data is available, but validation or extraction evidence requires review before carrier decisions.";
            return;
        }

        if (result.CanCompareCarriers &&
            result.CanCompareFinancials &&
            result.CanComparePlanDesign &&
            profile.ReadinessStatus == ProposalCanonicalReadinessStatuses.Ready)
        {
            result.ReadinessStatus = ProposalCanonicalReadinessStatuses.Ready;
            result.ReadinessMessage = "Saved carrier proposals are ready for factual financial, plan design, and underwriting comparison.";
            return;
        }

        result.ReadinessStatus = ProposalCanonicalReadinessStatuses.Partial;
        result.ReadinessMessage = "Saved proposal facts are usable, with missing comparison dimensions or evidence identified in the findings.";
    }

    private static void AddFinding(
        ProposalIntelligenceAnalysis result,
        ProposalIntelligenceFinding finding)
    {
        var duplicate = result.Findings.Any(existing =>
            existing.Code.Equals(finding.Code, StringComparison.OrdinalIgnoreCase) &&
            existing.ProposalId == finding.ProposalId &&
            existing.Message.Equals(finding.Message, StringComparison.OrdinalIgnoreCase));
        if (!duplicate)
        {
            result.Findings.Add(finding);
        }
    }

    private static string ResolveFindingCategory(string code, string sourceTable)
    {
        var text = $"{code} {sourceTable}";
        if (text.Contains("rate", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("premium", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("financial", StringComparison.OrdinalIgnoreCase))
        {
            return ProposalIntelligenceFindingCategories.Financial;
        }

        if (text.Contains("plan_design", StringComparison.OrdinalIgnoreCase))
        {
            return ProposalIntelligenceFindingCategories.PlanDesign;
        }

        if (text.Contains("booklet", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("validation", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("exception", StringComparison.OrdinalIgnoreCase))
        {
            return ProposalIntelligenceFindingCategories.Validation;
        }

        if (text.Contains("source", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("evidence", StringComparison.OrdinalIgnoreCase))
        {
            return ProposalIntelligenceFindingCategories.Evidence;
        }

        if (text.Contains("pool", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("retention", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("assumption", StringComparison.OrdinalIgnoreCase))
        {
            return ProposalIntelligenceFindingCategories.Underwriting;
        }

        return ProposalIntelligenceFindingCategories.Readiness;
    }

    private static string ResolveComparisonStatus(
        int expectedProposalCount,
        int availableProposalCount,
        bool hasDifference,
        bool needsReview = false)
    {
        if (needsReview)
        {
            return ProposalIntelligenceComparisonStatuses.NeedsReview;
        }

        if (availableProposalCount < 2)
        {
            return expectedProposalCount > 1 && availableProposalCount < expectedProposalCount
                ? ProposalIntelligenceComparisonStatuses.Incomplete
                : ProposalIntelligenceComparisonStatuses.NotComparable;
        }

        if (availableProposalCount < expectedProposalCount)
        {
            return ProposalIntelligenceComparisonStatuses.Incomplete;
        }

        return hasDifference
            ? ProposalIntelligenceComparisonStatuses.Different
            : ProposalIntelligenceComparisonStatuses.Same;
    }

    private static string BuildPlanDesignKey(ProposalCanonicalPlanDesignFact fact)
    {
        return BuildComparisonKey(
            "plan-design",
            FirstNonBlank(fact.BenefitLine, "Other"),
            FirstNonBlank(fact.ProvisionName, "Plan Design"));
    }

    private static string BuildComparisonKey(string prefix, params string?[] values)
    {
        return $"{prefix}:{string.Join("|", values.Select(NormalizeText))}";
    }

    private static int DistinctDecimalCount(IEnumerable<decimal?> values)
    {
        return values.Where(value => value.HasValue).Select(value => value!.Value).Distinct().Count();
    }

    private static int DistinctSignatureCount(IEnumerable<string> values)
    {
        return values.Where(value => !string.IsNullOrWhiteSpace(value)).Distinct(StringComparer.OrdinalIgnoreCase).Count();
    }

    private static int DistinctTextCount(IEnumerable<string> values)
    {
        return values
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Select(NormalizeText)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Count();
    }

    private static string FinancialSignature(params decimal?[] values)
    {
        return values.Any(value => value.HasValue)
            ? string.Join("|", values.Select(value => value?.ToString("0.############################", CultureInfo.InvariantCulture) ?? "missing"))
            : "";
    }

    private static string JoinDistinct(IEnumerable<string> values)
    {
        return string.Join(
            " | ",
            values
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Select(value => value.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase));
    }

    private static ProposalCanonicalEvidence SelectEvidence(IEnumerable<ProposalCanonicalEvidence> evidence)
    {
        return evidence
            .Where(item => item is not null)
            .OrderByDescending(item => item.ConfidenceScore ?? 0m)
            .ThenByDescending(item => item.UpdatedAt ?? item.CreatedAt ?? DateTime.MinValue)
            .FirstOrDefault() ?? new ProposalCanonicalEvidence();
    }

    private static string FormatDecimal(decimal? value)
    {
        return value?.ToString("0.####", CultureInfo.InvariantCulture) ?? "";
    }

    private static string NormalizeText(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return "";
        }

        return string.Join(" ", value.Trim().ToLowerInvariant().Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
    }

    private static string FirstNonBlank(params string?[] values)
    {
        return values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value))?.Trim() ?? "";
    }

    private static void SortOutput(ProposalIntelligenceAnalysis result)
    {
        result.Quotes.Sort((left, right) =>
        {
            var carrier = string.Compare(left.CarrierName, right.CarrierName, StringComparison.OrdinalIgnoreCase);
            return carrier != 0 ? carrier : left.ProposalId.CompareTo(right.ProposalId);
        });
        result.RateComparisons.Sort((left, right) => CompareKeys(
            [left.BenefitLine, left.ClassName, left.Tier, left.AgeBand, left.Gender, left.Province],
            [right.BenefitLine, right.ClassName, right.Tier, right.AgeBand, right.Gender, right.Province]));
        result.PremiumComparisons.Sort((left, right) => CompareKeys(
            [left.BenefitLine, left.ClassName, left.Tier],
            [right.BenefitLine, right.ClassName, right.Tier]));
        result.PlanDesignComparisons.Sort((left, right) => CompareKeys(
            [left.BenefitLine, left.ProvisionName],
            [right.BenefitLine, right.ProvisionName]));
        result.UnderwritingComparisons.Sort((left, right) => CompareKeys(
            [left.Category, left.BenefitLine, left.TermName],
            [right.Category, right.BenefitLine, right.TermName]));
        result.Findings.Sort((left, right) =>
        {
            var severity = SeverityRank(right.Severity).CompareTo(SeverityRank(left.Severity));
            if (severity != 0)
            {
                return severity;
            }

            return CompareKeys(
                [left.Category, left.Carrier, left.Message],
                [right.Category, right.Carrier, right.Message]);
        });
    }

    private static int CompareKeys(IReadOnlyList<string> left, IReadOnlyList<string> right)
    {
        for (var index = 0; index < Math.Min(left.Count, right.Count); index++)
        {
            var comparison = string.Compare(left[index], right[index], StringComparison.OrdinalIgnoreCase);
            if (comparison != 0)
            {
                return comparison;
            }
        }

        return left.Count.CompareTo(right.Count);
    }

    private static int SeverityRank(string severity)
    {
        return severity switch
        {
            ProposalCanonicalIssueSeverities.Error => 3,
            ProposalCanonicalIssueSeverities.Warning => 2,
            _ => 1
        };
    }

    private sealed record QuoteContext(
        ProposalCanonicalQuote Proposal,
        ProposalIntelligenceQuoteAnalysis Analysis)
    {
        public int ExpectedProposalCount { get; set; }
    }
}

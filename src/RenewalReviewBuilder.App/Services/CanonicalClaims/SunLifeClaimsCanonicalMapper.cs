using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services.CanonicalClaims;

/// <summary>Maps Sun Life Claims Paid by Type of Service PDFs into dated canonical utilization reports.</summary>
public sealed class SunLifeClaimsCanonicalMapper : ICarrierClaimsCanonicalMultiPeriodMapper<ClaimsExperienceDataModelResult>
{
    public const string CurrentMappingVersion = "SUN_LIFE_SERVICE_CLAIMS_CANONICAL_V1";
    public string CarrierName => "Sun Life";
    public string SourceParserName => nameof(SunLifeClaimsPaidByServiceParser);

    public CanonicalClaimsImportEnvelope Map(ClaimsExperienceDataModelResult parsedResult, long clientId, long? sourceDocumentId, long? importBatchId)
    {
        return MapMany(parsedResult, clientId, sourceDocumentId, importBatchId).FirstOrDefault()
            ?? MapLegacySummary(parsedResult, clientId, sourceDocumentId, importBatchId);
    }

    public IReadOnlyList<CanonicalClaimsImportEnvelope> MapMany(ClaimsExperienceDataModelResult parsedResult, long clientId, long? sourceDocumentId, long? importBatchId)
    {
        ArgumentNullException.ThrowIfNull(parsedResult);
        if (parsedResult.ServiceAdjudicationFacts.Count == 0)
        {
            return [MapLegacySummary(parsedResult, clientId, sourceDocumentId, importBatchId)];
        }

        var periods = parsedResult.ServiceAdjudicationFacts
            .Select(fact => (fact.PeriodStart, fact.PeriodEnd, MonthCount: ClaimsExperienceDataModelBuilder.InclusiveMonthCount(fact.PeriodStart, fact.PeriodEnd)))
            .Distinct()
            .OrderByDescending(period => period.PeriodEnd)
            .ThenByDescending(period => period.PeriodStart)
            .Take(4)
            .ToList();

        var envelopes = new List<CanonicalClaimsImportEnvelope>();
        foreach (var period in periods)
        {
            var envelope = CreateEnvelope(parsedResult, clientId, sourceDocumentId, importBatchId, period.PeriodStart, period.PeriodEnd);
            foreach (var sourceFact in DeduplicateFacts(parsedResult.ServiceAdjudicationFacts.Where(fact => fact.PeriodStart == period.PeriodStart && fact.PeriodEnd == period.PeriodEnd), envelope))
            {
                envelope.UtilizationFacts.Add(MapUtilizationFact(sourceFact, parsedResult, clientId, sourceDocumentId, importBatchId, envelope.MappingVersion));
            }

            foreach (var total in parsedResult.ServiceControlTotals.Where(total => total.PeriodStart == period.PeriodStart && total.PeriodEnd == period.PeriodEnd))
            {
                envelope.BenchmarkFacts.Add(MapControlTotal(total, parsedResult, clientId, sourceDocumentId, importBatchId, envelope.MappingVersion));
            }

            if (parsedResult.ExperiencePeriodStart == period.PeriodStart && parsedResult.ExperiencePeriodEnd == period.PeriodEnd)
            {
                AddLegacySummaryFacts(envelope, parsedResult, clientId, sourceDocumentId, importBatchId);
            }

            AddReconciliationWarnings(envelope);
            envelopes.Add(envelope);
        }

        AddDurationWarnings(envelopes);
        return envelopes;
    }

    private static CanonicalClaimsImportEnvelope CreateEnvelope(ClaimsExperienceDataModelResult parsed, long clientId, long? sourceDocumentId, long? importBatchId, DateOnly start, DateOnly end)
    {
        return new CanonicalClaimsImportEnvelope
        {
            MappingVersion = CurrentMappingVersion,
            Report = new CanonicalClaimsReport
            {
                ClientId = clientId,
                Carrier = "Sun Life",
                PolicyNumber = parsed.ClientNumber.Trim(),
                ImportBatchId = importBatchId,
                SourceDocumentId = sourceDocumentId,
                Period = new CanonicalClaimsPeriod
                {
                    Start = start,
                    End = end,
                    PeriodType = "Experience"
                }
            }
        };
    }

    private static CanonicalClaimsImportEnvelope MapLegacySummary(ClaimsExperienceDataModelResult parsed, long clientId, long? sourceDocumentId, long? importBatchId)
    {
        var envelope = CreateEnvelope(
            parsed,
            clientId,
            sourceDocumentId,
            importBatchId,
            parsed.ExperiencePeriodStart ?? DateOnly.FromDateTime(DateTime.Today),
            parsed.ExperiencePeriodEnd ?? DateOnly.FromDateTime(DateTime.Today));
        AddLegacySummaryFacts(envelope, parsed, clientId, sourceDocumentId, importBatchId);
        return envelope;
    }

    private static void AddLegacySummaryFacts(CanonicalClaimsImportEnvelope envelope, ClaimsExperienceDataModelResult parsed, long clientId, long? sourceDocumentId, long? importBatchId)
    {
        foreach (var summary in parsed.Summaries)
        {
            envelope.FinancialFacts.Add(new CanonicalClaimFinancialFact
            {
                FactId = StableId("financial", clientId, "Sun Life", parsed.ClientNumber, NormalizeBenefit(summary.BenefitLine), envelope.Report.Period.Start, envelope.Report.Period.End),
                Benefit = NormalizeBenefit(summary.BenefitLine),
                MeasurementType = "ClaimsExperienceSummary",
                Premium = summary.Premium,
                PaidClaims = summary.PaidClaims,
                IncurredClaims = summary.IncurredClaims,
                LossRatio = summary.LossRatio,
                Lives = summary.Lives,
                Availability = Availability(summary.PaidClaims, summary.IncurredClaims, summary.Premium),
                Evidence = Evidence(sourceDocumentId, importBatchId, summary.BenefitLine, summary.BenefitLine, "Existing Sun Life summary compatibility mapping", envelope.MappingVersion, null, null)
            });
        }

        foreach (var monthly in parsed.MonthlyClaims)
        {
            envelope.MonthlyFacts.Add(new CanonicalClaimMonthlyFact
            {
                FactId = StableId("monthly", clientId, "Sun Life", parsed.ClientNumber, NormalizeBenefit(monthly.BenefitLine), monthly.Month, envelope.Report.Period.End),
                Benefit = NormalizeBenefit(monthly.BenefitLine),
                MeasurementType = "MonthlyClaims",
                DimensionType = "Month",
                DimensionValue = monthly.Month?.ToString("yyyy-MM", CultureInfo.InvariantCulture) ?? "",
                Month = monthly.Month,
                Premium = monthly.Premium,
                PaidClaims = monthly.PaidClaims,
                IncurredClaims = monthly.IncurredClaims,
                ClaimsPlusFees = monthly.ClaimsPlusFees,
                LossRatio = monthly.LossRatio,
                Lives = monthly.Lives,
                Availability = Availability(monthly.PaidClaims, monthly.IncurredClaims, monthly.Premium),
                Evidence = Evidence(sourceDocumentId, importBatchId, monthly.BenefitLine, monthly.BenefitLine, "Existing Sun Life monthly compatibility mapping", envelope.MappingVersion, null, null)
            });
        }
    }

    private static IEnumerable<ClaimsServiceAdjudicationFact> DeduplicateFacts(IEnumerable<ClaimsServiceAdjudicationFact> facts, CanonicalClaimsImportEnvelope envelope)
    {
        foreach (var group in facts.GroupBy(FactDimensionKey, StringComparer.OrdinalIgnoreCase))
        {
            var distinct = group
                .GroupBy(fact => string.Join("|", fact.AmountSubmitted, fact.AmountEligible, fact.AmountPaid, fact.ReportedPaidPercentage), StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (distinct.Count > 1)
            {
                var chosen = group.OrderByDescending(fact => fact.SourceFile, StringComparer.OrdinalIgnoreCase).ThenByDescending(fact => fact.SourcePage).First();
                foreach (var previous in distinct
                             .Select(values => values.OrderByDescending(fact => fact.SourceFile, StringComparer.OrdinalIgnoreCase).ThenByDescending(fact => fact.SourcePage).First())
                             .Where(fact => !FactValueKey(fact).Equals(FactValueKey(chosen), StringComparison.OrdinalIgnoreCase)))
                {
                    var factId = StableFactId(chosen, envelope.Report.ClientId, envelope.Report.Carrier, envelope.Report.PolicyNumber);
                    envelope.FactRevisions.Add(new CanonicalClaimsFactRevision
                    {
                        RevisionId = StableId("revision", factId, previous.SourceFile, previous.SourcePage, chosen.SourceFile, chosen.SourcePage, FactValueKey(previous), FactValueKey(chosen)),
                        FactId = factId,
                        NaturalKey = FactNaturalKey(chosen, envelope.Report.ClientId, envelope.Report.Carrier, envelope.Report.PolicyNumber),
                        PreviousSourceReference = $"{previous.SourceFile} p{previous.SourcePage}",
                        RevisedSourceReference = $"{chosen.SourceFile} p{chosen.SourcePage}",
                        PreviousSubmittedAmount = previous.AmountSubmitted,
                        RevisedSubmittedAmount = chosen.AmountSubmitted,
                        PreviousEligibleAmount = previous.AmountEligible,
                        RevisedEligibleAmount = chosen.AmountEligible,
                        PreviousPaidAmount = previous.AmountPaid,
                        RevisedPaidAmount = chosen.AmountPaid
                    });
                }

                envelope.MappingWarnings.Add(new CanonicalClaimsValidationResult
                {
                    Code = "sun_life_changed_duplicate_fact_review_required",
                    Severity = "Warning",
                    FieldName = "service_fact_revision",
                    FactId = StableFactId(chosen, envelope.Report.ClientId, envelope.Report.Carrier, envelope.Report.PolicyNumber),
                    Message = "Multiple source documents reported different values for the same Sun Life service fact. The latest source row was retained and requires review.",
                    SourceReference = string.Join("; ", group.Select(fact => $"{fact.SourceFile} p{fact.SourcePage}").Distinct(StringComparer.OrdinalIgnoreCase))
                });
                yield return chosen;
                continue;
            }

            yield return group.First();
        }
    }

    private static CanonicalClaimsUtilizationFact MapUtilizationFact(ClaimsServiceAdjudicationFact source, ClaimsExperienceDataModelResult parsed, long clientId, long? sourceDocumentId, long? importBatchId, string mappingVersion)
    {
        var relationship = NormalizeRelationship(source.Relationship);
        return new CanonicalClaimsUtilizationFact
        {
            CanonicalClaimsUtilizationFactId = StableFactId(source, clientId, "Sun Life", parsed.ClientNumber),
            BenefitCode = Code(NormalizeBenefit(source.Benefit)),
            BenefitName = NormalizeBenefit(source.Benefit),
            CategoryCode = Code(source.Category),
            CategoryName = source.Category,
            ServiceCode = Code(source.Service),
            ServiceName = source.Service,
            ProvinceCode = "",
            RelationshipType = relationship,
            SubmittedAmount = source.AmountSubmitted,
            EligibleAmount = source.AmountEligible,
            PaidAmount = source.AmountPaid,
            ReportedPaidPercentage = source.ReportedPaidPercentage,
            Currency = FirstNonBlank(source.Currency, "CAD"),
            ValidationStatus = source.ValidationState.Equals("Valid", StringComparison.OrdinalIgnoreCase) ? CanonicalClaimsValidationStatus.Valid : CanonicalClaimsValidationStatus.NeedsReview,
            MeasurementStartDate = source.PeriodStart,
            MeasurementEndDate = source.PeriodEnd,
            SourceAggregationLevel = CanonicalClaimsSourceAggregationLevel.Service,
            MappingVersion = mappingVersion,
            Evidence = Evidence(sourceDocumentId, importBatchId, $"{source.Benefit}|{source.Relationship}|{source.Category}|{source.Service}", source.SourceEvidence, "Sun Life Claims Paid by Type of Service PDF", mappingVersion, source.SourceFile, source.SourcePage, source.Confidence)
        };
    }

    private static CanonicalBenchmarkFact MapControlTotal(ClaimsServiceControlTotal source, ClaimsExperienceDataModelResult parsed, long clientId, long? sourceDocumentId, long? importBatchId, string mappingVersion)
    {
        return new CanonicalBenchmarkFact
        {
            FactId = StableId("control", clientId, "Sun Life", parsed.ClientNumber, NormalizeBenefit(source.Benefit), source.PeriodStart, source.PeriodEnd, source.Relationship, source.Category, source.Service, source.TotalType),
            Benefit = NormalizeBenefit(source.Benefit),
            Category = source.Category,
            MeasurementType = "ReportedControlTotal",
            DimensionType = source.TotalType,
            DimensionValue = string.Join("|", new[] { source.Relationship, source.Category, source.Service }.Where(value => !string.IsNullOrWhiteSpace(value))),
            BenchmarkDefinition = "Sun Life reported non-additive claims control total",
            MonetaryValue = source.AmountPaid,
            PercentageValue = source.ReportedPaidPercentage,
            Currency = FirstNonBlank(source.Currency, "CAD"),
            UnitOfMeasure = "Paid claims control total",
            Availability = Availability(source.AmountPaid),
            Evidence = Evidence(sourceDocumentId, importBatchId, $"{source.Benefit}|{source.TotalType}|{source.Relationship}|{source.Category}|{source.Service}", source.SourceEvidence, "Sun Life Claims Paid by Type of Service PDF", mappingVersion, source.SourceFile, source.SourcePage)
        };
    }

    private static void AddReconciliationWarnings(CanonicalClaimsImportEnvelope envelope)
    {
        foreach (var fact in envelope.UtilizationFacts)
        {
            if (fact.SubmittedAmount < fact.EligibleAmount || fact.EligibleAmount < fact.PaidAmount)
            {
                envelope.MappingWarnings.Add(new CanonicalClaimsValidationResult
                {
                    Code = "sun_life_amount_order_review",
                    Severity = "Warning",
                    FieldName = "amounts",
                    FactId = fact.CanonicalClaimsUtilizationFactId,
                    Message = "Sun Life service amounts do not follow submitted >= eligible >= paid.",
                    SourceReference = SourceReference(fact.Evidence)
                });
            }
        }

        foreach (var relationshipTotal in envelope.BenchmarkFacts.Where(fact => fact.MeasurementType == "ReportedControlTotal" && fact.DimensionType == "RelationshipTotal"))
        {
            var relationship = relationshipTotal.DimensionValue.Split('|', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() ?? "";
            var relationshipType = NormalizeRelationship(relationship);
            CompareControlTotal(envelope, relationshipTotal, envelope.UtilizationFacts.Where(fact => fact.RelationshipType == relationshipType), "sun_life_relationship_reconciliation");
        }

        foreach (var categoryTotal in envelope.BenchmarkFacts.Where(fact => fact.MeasurementType == "ReportedControlTotal" && fact.DimensionType == "CategorySubtotal"))
        {
            var parts = categoryTotal.DimensionValue.Split('|', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
            var relationship = parts.ElementAtOrDefault(0) ?? "";
            var category = parts.ElementAtOrDefault(1) ?? "";
            var relationshipType = NormalizeRelationship(relationship);
            CompareControlTotal(envelope, categoryTotal, envelope.UtilizationFacts.Where(fact => fact.RelationshipType == relationshipType && fact.CategoryName.Equals(category, StringComparison.OrdinalIgnoreCase)), "sun_life_category_reconciliation");
        }

        foreach (var allService in envelope.BenchmarkFacts.Where(fact => fact.MeasurementType == "ReportedControlTotal" && fact.DimensionType == "AllService"))
        {
            var parts = allService.DimensionValue.Split('|', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
            var category = parts.ElementAtOrDefault(1) ?? "";
            var service = parts.ElementAtOrDefault(2) ?? "";
            CompareControlTotal(envelope, allService, envelope.UtilizationFacts.Where(fact => fact.CategoryName.Equals(category, StringComparison.OrdinalIgnoreCase) && fact.ServiceName.Equals(service, StringComparison.OrdinalIgnoreCase)), "sun_life_all_summary_reconciliation");
        }

        foreach (var grand in envelope.BenchmarkFacts.Where(fact => fact.MeasurementType == "ReportedControlTotal" && fact.DimensionType == "GrandTotal"))
        {
            CompareControlTotal(envelope, grand, envelope.UtilizationFacts, "sun_life_period_reconciliation");
            if (grand.PercentageValue.HasValue && Math.Abs(grand.PercentageValue.Value - 100m) > 0.1m)
            {
                envelope.MappingWarnings.Add(new CanonicalClaimsValidationResult
                {
                    Code = "sun_life_paid_percentage_not_100",
                    Severity = "Warning",
                    FieldName = "reported_paid_percentage",
                    FactId = grand.FactId,
                    Message = "Sun Life grand-total paid percentage was expected to be approximately 100%.",
                    SourceReference = SourceReference(grand.Evidence)
                });
            }
        }
    }

    private static void CompareControlTotal(CanonicalClaimsImportEnvelope envelope, CanonicalBenchmarkFact control, IEnumerable<CanonicalClaimsUtilizationFact> facts, string code)
    {
        var paid = facts.Sum(fact => fact.PaidAmount ?? 0m);
        if (Math.Abs(paid - (control.MonetaryValue ?? 0m)) <= 0.02m)
        {
            return;
        }

        envelope.MappingWarnings.Add(new CanonicalClaimsValidationResult
        {
            Code = code,
            Severity = "Warning",
            FieldName = control.DimensionType,
            FactId = control.FactId,
            Message = $"Sun Life control total paid claims {control.MonetaryValue?.ToString("C2", CultureInfo.CurrentCulture) ?? "blank"} does not reconcile to service facts {paid.ToString("C2", CultureInfo.CurrentCulture)}.",
            SourceReference = SourceReference(control.Evidence)
        });
    }

    private static void AddDurationWarnings(IReadOnlyList<CanonicalClaimsImportEnvelope> envelopes)
    {
        var durations = envelopes
            .Select(envelope => envelope.Report.Period.Start.HasValue && envelope.Report.Period.End.HasValue ? ClaimsExperienceDataModelBuilder.InclusiveMonthCount(envelope.Report.Period.Start.Value, envelope.Report.Period.End.Value) : 0)
            .Where(count => count > 0)
            .Distinct()
            .ToList();
        if (durations.Count <= 1)
        {
            return;
        }

        foreach (var envelope in envelopes)
        {
            envelope.MappingWarnings.Add(new CanonicalClaimsValidationResult
            {
                Code = "sun_life_comparison_duration_mismatch",
                Severity = "Warning",
                FieldName = "period_duration",
                Message = "Comparable Sun Life periods have different month counts and require review before presentation."
            });
        }
    }

    private static string FactNaturalKey(ClaimsServiceAdjudicationFact fact, long? clientId, string carrier, string policy) =>
        string.Join("|", clientId, carrier.Trim(), policy.Trim(), FactDimensionKey(fact));

    private static string FactDimensionKey(ClaimsServiceAdjudicationFact fact) => string.Join("|", NormalizeBenefit(fact.Benefit), fact.PeriodStart, fact.PeriodEnd, NormalizeRelationship(fact.Relationship), fact.Category, fact.Service);

    private static string FactValueKey(ClaimsServiceAdjudicationFact fact) => string.Join("|", fact.AmountSubmitted, fact.AmountEligible, fact.AmountPaid, fact.ReportedPaidPercentage, fact.Currency);

    private static string StableFactId(ClaimsServiceAdjudicationFact fact, long clientId, string carrier, string policy) => StableId("utilization", clientId, carrier, policy, NormalizeBenefit(fact.Benefit), fact.PeriodStart, fact.PeriodEnd, NormalizeRelationship(fact.Relationship), fact.Category, fact.Service);

    private static string StableId(params object?[] parts)
    {
        var raw = string.Join("|", parts.Select(part => part switch
        {
            DateOnly date => date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
            CanonicalClaimsRelationshipType relationship => relationship.ToString(),
            null => "",
            _ => Convert.ToString(part, CultureInfo.InvariantCulture) ?? ""
        }).Select(part => part.Trim().ToUpperInvariant()));
        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(raw))).ToLowerInvariant();
    }

    private static CanonicalClaimsSourceEvidence Evidence(long? documentId, long? batchId, string rawLabel, string rawValue, string method, string mappingVersion, string? sourceFile, int? sourcePage, decimal? confidence = null)
    {
        return new CanonicalClaimsSourceEvidence
        {
            SourceDocumentId = documentId,
            ImportBatchId = batchId,
            SourcePageNumber = sourcePage,
            WorksheetName = sourceFile ?? "Sun Life PDF",
            SourceTableOrSection = "Claims Paid by Type of Service",
            RawCarrierLabel = rawLabel,
            NormalizedLabel = rawLabel,
            RawExtractedValue = rawValue,
            ExtractionMethod = method,
            ParserName = nameof(SunLifeClaimsPaidByServiceParser),
            ParserVersion = "1.0",
            MappingVersion = mappingVersion,
            ConfidenceScore = confidence,
            Precision = sourcePage.HasValue ? CanonicalClaimsEvidencePrecision.Page : CanonicalClaimsEvidencePrecision.Document,
            MappingClassification = CanonicalClaimsMappingClassification.Direct
        };
    }

    private static CanonicalDataAvailability Availability(params decimal?[] values)
    {
        if (values.All(value => !value.HasValue)) return CanonicalDataAvailability.Missing;
        if (values.Where(value => value.HasValue).All(value => value == 0m)) return CanonicalDataAvailability.ReportedZero;
        return CanonicalDataAvailability.Present;
    }

    private static CanonicalClaimsRelationshipType NormalizeRelationship(string value)
    {
        var normalized = new string((value ?? "").Where(char.IsLetterOrDigit).Select(char.ToUpperInvariant).ToArray());
        return normalized switch
        {
            "EMPLOYEE" or "MEMBER" or "CERTIFICATEHOLDER" => CanonicalClaimsRelationshipType.Employee,
            "SPOUSE" or "PARTNER" => CanonicalClaimsRelationshipType.Spouse,
            "CHILD" or "CHILDREN" => CanonicalClaimsRelationshipType.Child,
            "OTHERDEPENDANT" or "OTHERDEPENDENT" or "DEPENDANT" or "DEPENDENT" => CanonicalClaimsRelationshipType.OtherDependant,
            "UNALLOCATED" => CanonicalClaimsRelationshipType.Unallocated,
            "ALL" or "TOTAL" => CanonicalClaimsRelationshipType.Total,
            _ => CanonicalClaimsRelationshipType.Unknown
        };
    }

    private static string NormalizeBenefit(string value)
    {
        var normalized = value.Trim();
        return normalized.Equals("Extended Health Care", StringComparison.OrdinalIgnoreCase) || normalized.Equals("Health", StringComparison.OrdinalIgnoreCase)
            ? "EHC"
            : normalized.Equals("Dental Care", StringComparison.OrdinalIgnoreCase)
                ? "Dental"
                : normalized;
    }

    private static string Code(string value) => string.Concat((value ?? "").Where(char.IsLetterOrDigit)).ToUpperInvariant();
    private static string FirstNonBlank(string? value, string fallback) => string.IsNullOrWhiteSpace(value) ? fallback : value.Trim();
    private static string SourceReference(CanonicalClaimsSourceEvidence evidence) => string.IsNullOrWhiteSpace(evidence.WorksheetName) ? "Sun Life PDF" : $"{evidence.WorksheetName} p{evidence.SourcePageNumber}";
}

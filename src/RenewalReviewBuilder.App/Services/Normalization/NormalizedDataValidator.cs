using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services.Normalization;

public sealed class NormalizedDataValidator
{
    public List<ValidationIssueRecord> Validate(NormalizedImportPayload payload)
    {
        var issues = new List<ValidationIssueRecord>();
        var source = payload.Source;

        if (string.IsNullOrWhiteSpace(payload.Client?.ClientName))
        {
            Add(issues, source, ValidationIssueCodes.MissingClientName, "Client", "client_name", "Client name is missing.");
        }

        if (string.IsNullOrWhiteSpace(payload.Carrier?.CarrierName) &&
            string.IsNullOrWhiteSpace(payload.Policy?.CarrierName) &&
            string.IsNullOrWhiteSpace(source.SourceCarrier))
        {
            Add(issues, source, ValidationIssueCodes.MissingCarrier, "Carrier", "carrier_name", "Carrier is missing.");
        }

        if (payload.DocumentType.Equals(DocumentTypes.Renewal, StringComparison.OrdinalIgnoreCase) &&
            payload.Renewals.All(item => !item.EffectiveDate.HasValue) &&
            payload.Policy?.RenewalDate is null)
        {
            Add(issues, source, ValidationIssueCodes.MissingRenewalEffectiveDate, "Renewal", "effective_date", "Renewal effective date is missing.");
        }

        if (payload.DocumentType.Equals(DocumentTypes.Renewal, StringComparison.OrdinalIgnoreCase) &&
            payload.Premiums.All(item => !item.Amount.HasValue) &&
            payload.Renewals.All(item => !item.CurrentPremium.HasValue && !item.RenewalPremium.HasValue))
        {
            Add(issues, source, ValidationIssueCodes.MissingPremium, "Premium", "amount", "Premium is missing.");
        }

        if ((payload.DocumentType.Equals(DocumentTypes.Renewal, StringComparison.OrdinalIgnoreCase) ||
             payload.DocumentType.Equals(DocumentTypes.ClaimsExperience, StringComparison.OrdinalIgnoreCase)) &&
            payload.Claims.Count == 0 &&
            payload.ClaimsExperience.Count == 0 &&
            payload.Renewals.All(item => !item.Claims.HasValue || item.Claims == 0m))
        {
            Add(issues, source, ValidationIssueCodes.MissingClaims, "Claim", "claims", "Claims are missing.");
        }

        if (payload.DocumentType.Equals(DocumentTypes.Renewal, StringComparison.OrdinalIgnoreCase) &&
            payload.Rates.Count == 0)
        {
            Add(issues, source, ValidationIssueCodes.MissingRates, "Rate", "rate", "Rates are missing.");
        }

        foreach (var rate in payload.Rates.Where(item => string.IsNullOrWhiteSpace(item.ClassName)))
        {
            Add(issues, rate.Source, ValidationIssueCodes.MissingClass, "Rate", "class", $"Class is missing for {FirstNonBlank(rate.BenefitLine, "a rate line")}.");
        }

        foreach (var claim in payload.ClaimsExperience.Where(item => item.PeriodStart.HasValue && item.PeriodEnd.HasValue && item.PeriodStart > item.PeriodEnd))
        {
            Add(issues, claim.Source, ValidationIssueCodes.ClaimsPeriodMismatch, "ClaimsExperience", "experience_period", $"Claims period is invalid for {FirstNonBlank(claim.BenefitLine, "a claims row")}.");
        }

        var renewalDates = payload.Renewals
            .Select(item => item.EffectiveDate)
            .Where(item => item.HasValue)
            .Distinct()
            .ToList();
        if (renewalDates.Count > 1)
        {
            Add(issues, source, ValidationIssueCodes.RenewalPeriodMismatch, "Renewal", "effective_date", "Renewal rows have more than one effective date.");
        }

        foreach (var item in payload.Rates.Where(item => string.IsNullOrWhiteSpace(item.BenefitLine)))
        {
            Add(issues, item.Source, ValidationIssueCodes.BenefitNotMapped, "Rate", "benefit_line", "A rate row is not mapped to a benefit.");
        }

        foreach (var item in payload.BillingStatements.Where(item => string.IsNullOrWhiteSpace(item.Province)))
        {
            Add(issues, item.Source, ValidationIssueCodes.ProvinceMissing, "BillingStatement", "province", "Billing statement province is missing.");
        }

        foreach (var item in payload.BillingStatements.Where(item => string.IsNullOrWhiteSpace(item.OptionLabel) || string.IsNullOrWhiteSpace(item.ClassCode)))
        {
            Add(issues, item.Source, ValidationIssueCodes.OptionClassMissing, "BillingStatement", "option_class", "Billing statement option or class is missing.");
        }

        foreach (var item in payload.PlanDesign)
        {
            if (string.IsNullOrWhiteSpace(item.ProvisionCode) ||
                string.IsNullOrWhiteSpace(item.ProvisionLabel))
            {
                Add(
                    issues,
                    item.Source,
                    ValidationIssueCodes.PlanDesignProvisionIdentityMissing,
                    "PlanDesign",
                    "provision_code",
                    "A Plan Design row could not be assigned a stable provision identity.");
            }

            if (CanonicalSourceLineageValueGuardrails.IsPercentageOutOfRange(item.ReimbursementPercent) ||
                CanonicalSourceLineageValueGuardrails.HasOutOfRangePercentageText(item.ReimbursementText))
            {
                Add(
                    issues,
                    item.Source,
                    ValidationIssueCodes.PlanDesignPercentageOutOfRange,
                    "PlanDesign",
                    "reimbursement_percent",
                    "Reimbursement percentage must be between 0 and 100. The numeric value was rejected and the source text was retained.");
            }
            else if (CanonicalSourceLineageValueGuardrails.HasMultiplePercentageValues(item.ReimbursementText))
            {
                Add(
                    issues,
                    item.Source,
                    ValidationIssueCodes.PlanDesignMultiplePercentagesPreserved,
                    "PlanDesign",
                    "reimbursement_percent",
                    "Multiple reimbursement percentages were preserved as source text and require structured review.");
            }
        }

        foreach (var duplicate in payload.PlanDesign
                     .Where(item => !string.IsNullOrWhiteSpace(item.ProvisionCode))
                     .GroupBy(
                         item => string.Join(
                             "|",
                             PlanDesignCanonicalMapper.ResolveProvisionBenefitCode(new PlanDesignRecord
                             {
                                 BenefitCategory = item.BenefitCategory,
                                 BenefitName = item.BenefitName
                             }),
                             PlanDesignCanonicalMapper.NormalizeComparableText(FirstNonBlank(item.ClassName, "All Classes")),
                             item.ProvisionCode),
                         StringComparer.OrdinalIgnoreCase)
                     .Where(group => group.Count() > 1))
        {
            Add(
                issues,
                duplicate.First().Source,
                ValidationIssueCodes.PlanDesignDuplicateProvisionIdentity,
                "PlanDesign",
                "provision_code",
                $"Provision '{duplicate.First().ProvisionLabel}' occurs {duplicate.Count()} times for the same benefit and class.");
        }

        return issues;
    }

    public List<CanonicalClaimsValidationResult> Validate(CanonicalClaimsImportEnvelope envelope)
    {
        ArgumentNullException.ThrowIfNull(envelope);
        var results = new List<CanonicalClaimsValidationResult>();
        var report = envelope.Report;

        AddIf(report.ClientId <= 0, "canonical_client_missing", "Error", "client_id", "A valid client assignment is required.");
        AddIf(string.IsNullOrWhiteSpace(report.Carrier), "canonical_carrier_missing", "Error", "carrier", "A valid carrier assignment is required.");
        AddIf(!report.Period.Start.HasValue || !report.Period.End.HasValue, "canonical_period_missing", "Error", "reporting_period", "A complete reporting period is required.");
        AddIf(report.Period.Start.HasValue && report.Period.End.HasValue && report.Period.Start > report.Period.End, "canonical_period_invalid", "Error", "reporting_period", "The reporting period start must not be after the end.");
        AddIf(string.IsNullOrWhiteSpace(report.PolicyNumber), "canonical_policy_missing", "Warning", "policy_number", "Policy number was not available from the existing claims result.");
        AddIf(!report.SourceDocumentId.HasValue || !report.ImportBatchId.HasValue, "canonical_evidence_missing", "Warning", "source_evidence", "Document or import-batch evidence is incomplete.");
        AddIf(envelope.FinancialFacts.Count == 0 && envelope.MonthlyFacts.Count == 0 && envelope.CategoryFacts.Count == 0 && envelope.DistributionFacts.Count == 0 && envelope.UtilizationFacts.Count == 0, "canonical_facts_missing", "Error", "facts", "At least one canonical claims fact is required.");

        var allFacts = envelope.FinancialFacts.Cast<CanonicalClaimFact>()
            .Concat(envelope.MonthlyFacts)
            .Concat(envelope.CategoryFacts)
            .Concat(envelope.DistributionFacts)
            .Concat(envelope.BenchmarkFacts)
            .ToList();

        foreach (var duplicate in allFacts.GroupBy(fact => fact.FactId).Where(group => group.Count() > 1))
        {
            results.Add(new CanonicalClaimsValidationResult { Code = "canonical_duplicate_fact", Severity = "Error", FieldName = "fact_id", FactId = duplicate.Key, Message = "Canonical fact ID is duplicated inside the import." });
        }

        foreach (var fact in allFacts)
        {
            if (string.IsNullOrWhiteSpace(fact.Benefit))
                results.Add(new CanonicalClaimsValidationResult { Code = "canonical_benefit_unmapped", Severity = "Warning", FieldName = "benefit", FactId = fact.FactId, Message = "A claims fact has no normalized benefit." });
            if (fact.CountValue < 0)
                results.Add(new CanonicalClaimsValidationResult { Code = "canonical_count_invalid", Severity = "Error", FieldName = "count_value", FactId = fact.FactId, Message = "Counts cannot be negative." });
              if (!fact.Evidence.SourceDocumentId.HasValue || !fact.Evidence.ImportBatchId.HasValue)
                  results.Add(new CanonicalClaimsValidationResult { Code = "canonical_fact_evidence_missing", Severity = "Warning", FieldName = "evidence", FactId = fact.FactId, Message = "Fact evidence is less precise than an import batch and source document." });
              if (!string.IsNullOrWhiteSpace(fact.Evidence.WarningCode))
                  results.Add(new CanonicalClaimsValidationResult { Code = fact.Evidence.WarningCode, Severity = "Warning", FieldName = "evidence", FactId = fact.FactId, Message = fact.Evidence.WarningMessage, MappingClassification = fact.Evidence.MappingClassification, NormalizationRule = fact.Evidence.NormalizationRule });
        }

        foreach (var monthly in envelope.MonthlyFacts.Where(fact => !fact.Month.HasValue))
            results.Add(new CanonicalClaimsValidationResult { Code = "canonical_month_missing", Severity = "Error", FieldName = "month", FactId = monthly.FactId, Message = "A monthly fact requires a normalized month." });

        foreach (var financial in envelope.FinancialFacts.Where(fact => fact.Lives < 0))
            results.Add(new CanonicalClaimsValidationResult { Code = "canonical_lives_invalid", Severity = "Error", FieldName = "lives", FactId = financial.FactId, Message = "Lives cannot be negative." });

        foreach (var monthly in envelope.MonthlyFacts.Where(fact => fact.Lives < 0))
            results.Add(new CanonicalClaimsValidationResult { Code = "canonical_lives_invalid", Severity = "Error", FieldName = "lives", FactId = monthly.FactId, Message = "Lives cannot be negative." });

        foreach (var duplicate in envelope.UtilizationFacts.GroupBy(fact => string.Join("|", fact.BenefitCode, fact.CategoryCode, fact.ServiceCode, fact.ProvinceCode, fact.RelationshipType, fact.DivisionCode, fact.ClassCode, fact.MeasurementStartDate, fact.MeasurementEndDate, fact.SourceAggregationLevel)).Where(group => group.Count() > 1))
            results.Add(new CanonicalClaimsValidationResult { Code = "canonical_utilization_duplicate_grain", Severity = "Warning", FieldName = "utilization_grain", FactId = duplicate.First().CanonicalClaimsUtilizationFactId, Message = "Multiple utilization facts have the same source grain and require source-version review." });

        foreach (var fact in envelope.UtilizationFacts)
        {
            void Warn(bool condition, string code, string field, string message) { if (condition) results.Add(new CanonicalClaimsValidationResult { Code=code, Severity="Warning", FieldName=field, FactId=fact.CanonicalClaimsUtilizationFactId, Message=message, SourceReference=$"{fact.Evidence.WorksheetName}!{fact.Evidence.SourceColumnOrCell}" }); }
            Warn(string.IsNullOrWhiteSpace(fact.BenefitCode), "canonical_utilization_benefit_missing", "benefit", "A utilization fact requires an identified benefit.");
            Warn(!fact.MeasurementStartDate.HasValue || !fact.MeasurementEndDate.HasValue || fact.MeasurementStartDate > fact.MeasurementEndDate, "canonical_utilization_period_invalid", "measurement_period", "A utilization fact requires a valid measurement period.");
            Warn(!Enum.IsDefined(fact.RelationshipType), "canonical_utilization_relationship_invalid", "relationship_type", "Relationship type is not canonical.");
            Warn(!fact.UniqueClaimants.HasValue && !fact.Occurrences.HasValue && !fact.SubmittedAmount.HasValue && !fact.EligibleAmount.HasValue && !fact.PaidAmount.HasValue, "canonical_utilization_measures_missing", "measures", "At least one utilization measure is required.");
            Warn(fact.Occurrences < 0 || fact.UniqueClaimants < 0, "canonical_utilization_negative_count", "counts", "Unique claimants and occurrences cannot be negative.");
            Warn(fact.SubmittedAmount < 0 || fact.EligibleAmount < 0 || fact.PaidAmount < 0, "canonical_utilization_negative_amount", "amounts", "Negative monetary values were preserved as possible adjustments and require review.");
            Warn(fact.Occurrences > 0 && fact.UniqueClaimants > fact.Occurrences, "canonical_utilization_claimants_exceed_occurrences", "unique_claimants", "Unique claimants normally should not exceed occurrences.");
            Warn(fact.PaidAmount > fact.EligibleAmount, "canonical_utilization_paid_exceeds_eligible", "paid_amount", "Paid amount exceeds eligible amount and was preserved for review.");
            Warn(fact.EligibleAmount > fact.SubmittedAmount, "canonical_utilization_eligible_exceeds_submitted", "eligible_amount", "Eligible amount exceeds submitted amount and was preserved for review.");
            Warn(fact.Occurrences == 0 && fact.SubmittedAmount > 0, "canonical_utilization_submitted_without_occurrences", "submitted_amount", "Positive submitted amount with zero occurrences was preserved.");
            Warn(fact.EligibleAmount == 0 && fact.PaidAmount > 0, "canonical_utilization_paid_without_eligible", "paid_amount", "Paid amount with zero eligible amount was preserved.");
            Warn(!fact.Evidence.SourceDocumentId.HasValue || string.IsNullOrWhiteSpace(fact.Evidence.WorksheetName) || !fact.Evidence.SourceRow.HasValue, "canonical_utilization_evidence_missing", "evidence", "Utilization evidence must identify source document, worksheet, and row.");
        }

        return results;

        void AddIf(bool condition, string code, string severity, string field, string message)
        {
            if (condition) results.Add(new CanonicalClaimsValidationResult { Code = code, Severity = severity, FieldName = field, Message = message });
        }
    }

    private static void Add(
        ICollection<ValidationIssueRecord> issues,
        NormalizedSourceMetadata source,
        string code,
        string tableName,
        string fieldName,
        string message)
    {
        issues.Add(new ValidationIssueRecord
        {
            ImportBatchId = source.ImportBatchId,
            IssueCode = code,
            Severity = "Warning",
            TableName = tableName,
            FieldName = fieldName,
            Message = message,
            RawSourceReference = source.RawSourceReference,
            CreatedAt = DateTime.Now
        });
    }

    private static string FirstNonBlank(params string?[] values)
    {
        return values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)) ?? "";
    }
}

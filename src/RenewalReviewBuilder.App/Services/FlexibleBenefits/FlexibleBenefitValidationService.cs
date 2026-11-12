using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services.FlexibleBenefits;

public sealed class FlexibleBenefitValidationService
{
    private readonly FlexibleBenefitCategoryCatalogue _categoryCatalogue;

    public FlexibleBenefitValidationService(FlexibleBenefitCategoryCatalogue? categoryCatalogue = null)
    {
        _categoryCatalogue = categoryCatalogue ?? new FlexibleBenefitCategoryCatalogue();
    }

    public FlexibleBenefitValidationResult Validate(
        FlexibleBenefitExtractionResult extraction,
        long? confirmedClientId)
    {
        ArgumentNullException.ThrowIfNull(extraction);
        var result = new FlexibleBenefitValidationResult();
        foreach (var issue in extraction.ValidationIssues)
        {
            Add(result, issue.Code, issue.Severity, issue.FieldName, issue.Message, issue.SourcePageNumber);
        }

        ValidateIdentity(extraction, confirmedClientId, result);
        ValidatePlan(extraction, result);
        ValidateClasses(extraction, result);
        ValidateCategories(extraction, result);
        ValidateFundingAndFees(extraction, result);
        ValidateAgreements(extraction, result);
        ValidateLineage(extraction, result);

        result.Status = result.Issues.Any(issue => issue.IsFailure)
            ? FlexibleBenefitExtractionStatuses.Failed
            : result.HasWarnings
                ? FlexibleBenefitExtractionStatuses.NeedsReview
                : FlexibleBenefitExtractionStatuses.Completed;
        return result;
    }

    private static void ValidateIdentity(
        FlexibleBenefitExtractionResult extraction,
        long? confirmedClientId,
        FlexibleBenefitValidationResult result)
    {
        if (!confirmedClientId.HasValue || confirmedClientId <= 0)
        {
            Add(result, "client_not_confirmed", "Error", "ClientId", "Confirm the client before importing the flexible-benefits plan.");
        }

        if (string.IsNullOrWhiteSpace(extraction.Client.LegalName))
        {
            Add(result, "client_name_missing", "Warning", "Client.LegalName", "The legal company name was not confirmed from the source document.");
        }

        if (string.IsNullOrWhiteSpace(extraction.Client.PostalCode))
        {
            Add(result, "client_postal_code_missing", "Warning", "Client.PostalCode", "The postal code is missing or invalid and should be reviewed.", 1);
        }

        if (string.IsNullOrWhiteSpace(extraction.Classification.Carrier) ||
            string.IsNullOrWhiteSpace(extraction.Classification.Administrator))
        {
            Add(result, "administrator_not_confirmed", "Error", "Classification.Administrator", "The carrier or administrator could not be confirmed.");
        }

        if (!extraction.Classification.DocumentType.Equals(
                FlexibleBenefitDocumentTypes.FlexibleBenefitsPlanDocument,
                StringComparison.OrdinalIgnoreCase))
        {
            Add(result, "document_type_not_confirmed", "Error", "Classification.DocumentType", "The flexible-benefits document type could not be confirmed.");
        }

        var planAdministrator = extraction.Contacts.FirstOrDefault(contact =>
            contact.ContactRole.Equals("PlanAdministrator", StringComparison.OrdinalIgnoreCase));
        if (planAdministrator is not null &&
            string.IsNullOrWhiteSpace(planAdministrator.Email) &&
            !result.Issues.Any(issue => issue.FieldName.Equals("PlanAdministrator.Email", StringComparison.OrdinalIgnoreCase)))
        {
            Add(result, "plan_administrator_email_needs_review", "Warning", "PlanAdministrator.Email", "The plan administrator email is missing or invalid.", planAdministrator.SourcePageNumber);
        }
    }

    private static void ValidatePlan(
        FlexibleBenefitExtractionResult extraction,
        FlexibleBenefitValidationResult result)
    {
        var plan = extraction.Plan;
        var meaningfulFieldCount = extraction.Classes.Count +
                                   extraction.Categories.Count(category => category.Enabled.HasValue) +
                                   extraction.Fees.Count +
                                   extraction.FundingConfigurations.Count +
                                   (plan.EffectiveDate.HasValue ? 1 : 0) +
                                   (plan.DefaultHcsaPercentage.HasValue || plan.DefaultLsaPercentage.HasValue ? 1 : 0);
        if (meaningfulFieldCount == 0)
        {
            Add(result, "zero_meaningful_plan_fields", "Error", "FlexibleBenefitPlan", "No meaningful flexible-benefits plan-design fields were extracted.");
        }

        if (!plan.EffectiveDate.HasValue)
        {
            Add(result, "effective_date_missing", "Warning", "FlexibleBenefitPlan.EffectiveDate", "The plan effective date was not confirmed.");
        }

        ValidatePercentage(plan.DefaultHcsaPercentage, "FlexibleBenefitPlan.DefaultHcsaPercentage", result);
        ValidatePercentage(plan.DefaultLsaPercentage, "FlexibleBenefitPlan.DefaultLsaPercentage", result);
        if (plan.DefaultAllocationType is FlexibleBenefitAllocationType.AllHcsa or
            FlexibleBenefitAllocationType.AllLsa or
            FlexibleBenefitAllocationType.Split)
        {
            if (!plan.DefaultHcsaPercentage.HasValue || !plan.DefaultLsaPercentage.HasValue)
            {
                Add(result, "fixed_allocation_incomplete", "Warning", "FlexibleBenefitPlan.DefaultAllocation", "A fixed HCSA/LSA allocation is missing one percentage.");
            }
            else if (plan.DefaultHcsaPercentage.Value + plan.DefaultLsaPercentage.Value != 100m)
            {
                Add(result, "fixed_allocation_not_100", "Error", "FlexibleBenefitPlan.DefaultAllocation", "The fixed HCSA and LSA percentages must total 100.");
            }
        }

        if (plan.DefaultAllocationType == FlexibleBenefitAllocationType.AllLsa && plan.HcsaEnabled == false)
        {
            Add(result, "hcsa_disabled_from_lsa_default", "Warning", "FlexibleBenefitPlan.HcsaEnabled", "A 100% LSA default does not by itself confirm that HCSA is unavailable.");
        }

        ValidateNonNegative(plan.RunoffDays, "FlexibleBenefitPlan.RunoffDays", result);
        ValidateNonNegative(plan.DefaultWaitingPeriodDays, "FlexibleBenefitPlan.DefaultWaitingPeriodDays", result);
        if (plan.ChildDependentCutoffAge.HasValue && plan.StudentDependentCutoffAge.HasValue &&
            plan.ChildDependentCutoffAge > plan.StudentDependentCutoffAge)
        {
            Add(result, "dependent_age_order", "Warning", "FlexibleBenefitPlan.DependentAges", "The child cutoff age exceeds the student cutoff age.");
        }
    }

    private static void ValidateClasses(
        FlexibleBenefitExtractionResult extraction,
        FlexibleBenefitValidationResult result)
    {
        if (extraction.Classes.Count == 0)
        {
            Add(result, "class_design_missing", "Warning", "FlexibleBenefitClass", "No class design was confirmed.");
            return;
        }

        foreach (var planClass in extraction.Classes)
        {
            ValidateNonNegative(planClass.AnnualSingleAmount, $"FlexibleBenefitClass[{planClass.ClassName}].AnnualSingleAmount", result);
            ValidateNonNegative(planClass.AnnualCoupleAmount, $"FlexibleBenefitClass[{planClass.ClassName}].AnnualCoupleAmount", result);
            ValidateNonNegative(planClass.AnnualFamilyAmount, $"FlexibleBenefitClass[{planClass.ClassName}].AnnualFamilyAmount", result);
            ValidateNonNegative(planClass.AnnualEmployeeAmount, $"FlexibleBenefitClass[{planClass.ClassName}].AnnualEmployeeAmount", result);
            ValidateNonNegative(planClass.AnnualDependentAmount, $"FlexibleBenefitClass[{planClass.ClassName}].AnnualDependentAmount", result);
            ValidateNonNegative(planClass.HcsaAmount, $"FlexibleBenefitClass[{planClass.ClassName}].HcsaAmount", result);
            ValidateNonNegative(planClass.LsaAmount, $"FlexibleBenefitClass[{planClass.ClassName}].LsaAmount", result);
            ValidateNonNegative(planClass.WaitingPeriodDays, $"FlexibleBenefitClass[{planClass.ClassName}].WaitingPeriodDays", result);
            ValidatePercentage(planClass.HcsaPercentage, $"FlexibleBenefitClass[{planClass.ClassName}].HcsaPercentage", result);
            ValidatePercentage(planClass.LsaPercentage, $"FlexibleBenefitClass[{planClass.ClassName}].LsaPercentage", result);

            var tierValues = new[] { planClass.AnnualSingleAmount, planClass.AnnualCoupleAmount, planClass.AnnualFamilyAmount };
            if (tierValues.Any(value => value.HasValue) && tierValues.Any(value => !value.HasValue))
            {
                Add(result, "class_tier_limits_incomplete", "Warning", $"FlexibleBenefitClass[{planClass.ClassName}]", "Retain single, couple, and family values separately when the form provides all three.", planClass.SourcePageNumber);
            }
        }
    }

    private void ValidateCategories(
        FlexibleBenefitExtractionResult extraction,
        FlexibleBenefitValidationResult result)
    {
        var catalogueCodes = _categoryCatalogue.Load()
            .Select(category => category.CategoryCode)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        var duplicateCodes = extraction.Categories
            .GroupBy(category => category.CategoryCode, StringComparer.OrdinalIgnoreCase)
            .Where(group => group.Count() > 1)
            .Select(group => group.Key)
            .ToList();
        if (duplicateCodes.Count > 0)
        {
            Add(result, "duplicate_category_codes", "Error", "FlexibleBenefitCategory.CategoryCode", "Duplicate normalized category codes were extracted.");
        }

        foreach (var category in extraction.Categories)
        {
            if (!catalogueCodes.Contains(category.CategoryCode))
            {
                Add(result, "uncontrolled_category", "Warning", $"FlexibleBenefitCategory[{category.CategoryCode}]", "The category is not in the controlled flexible-benefits catalogue.", category.SourcePageNumber);
            }

            if (!category.Enabled.HasValue)
            {
                Add(result, "category_checkbox_unclear", "Warning", $"FlexibleBenefitCategory[{category.CategoryCode}].Enabled", "The category checkbox state is unclear.", category.SourcePageNumber);
            }
        }

        var detectedMarks = extraction.OcrDocument?.Pages
            .Where(page => page.PageNumber <= 4)
            .SelectMany(page => page.Marks
                .Where(mark => mark.Shape == "Selection mark")
                .Where(mark => mark.Bounds.Width >= page.PixelWidth * 0.016 && mark.Bounds.Height >= page.PixelWidth * 0.016)
                .Where(mark => mark.Bounds.Width <= page.PixelWidth * 0.035 && mark.Bounds.Height <= page.PixelWidth * 0.035)
                .Where(mark => mark.Bounds.X < page.PixelWidth * 0.25))
            .ToList() ?? [];
        var categorySelectedCount = extraction.Categories.Count(category => category.Enabled == true);
        var detectedSelectedCount = detectedMarks.Count(mark => mark.IsSelected == true);
        if (categorySelectedCount != detectedSelectedCount)
        {
            Add(result, "category_checkbox_reconciliation", "Warning", "FlexibleBenefitCategory.Enabled", $"The normalized selected-category count ({categorySelectedCount}) does not match the detected selected checkbox count ({detectedSelectedCount}).");
        }

        var payCategory = extraction.Categories.FirstOrDefault(category =>
            category.CategoryCode.Equals("PAY_ALL_RECEIPTS_AS_SUBMITTED", StringComparison.OrdinalIgnoreCase));
        if (extraction.Plan.PayAsSubmitted == true && payCategory?.Enabled != true)
        {
            Add(result, "pay_as_submitted_not_selected", "Error", "FlexibleBenefitPlan.PayAsSubmitted", "Pay as submitted cannot be enabled unless its checkbox is selected.", payCategory?.SourcePageNumber);
        }
    }

    private static void ValidateFundingAndFees(
        FlexibleBenefitExtractionResult extraction,
        FlexibleBenefitValidationResult result)
    {
        foreach (var fee in extraction.Fees)
        {
            ValidatePercentage(fee.Percentage, $"FlexibleBenefitFee[{fee.FeeType}].Percentage", result);
            ValidateNonNegative(fee.FlatAmount, $"FlexibleBenefitFee[{fee.FeeType}].FlatAmount", result);
            ValidateNonNegative(fee.MinimumAmount, $"FlexibleBenefitFee[{fee.FeeType}].MinimumAmount", result);
            ValidateNonNegative(fee.MaximumAmount, $"FlexibleBenefitFee[{fee.FeeType}].MaximumAmount", result);
            ValidateNonNegative(fee.NoticePeriodDays, $"FlexibleBenefitFee[{fee.FeeType}].NoticePeriodDays", result);
            if (fee.FeeCalculationType.Equals("Percentage", StringComparison.OrdinalIgnoreCase) && !fee.Percentage.HasValue)
            {
                Add(result, "fee_percentage_missing", "Warning", $"FlexibleBenefitFee[{fee.FeeType}].Percentage", "The fee is percentage-based but no percentage was confirmed.", fee.SourcePageNumber);
            }
        }
    }

    private static void ValidateAgreements(
        FlexibleBenefitExtractionResult extraction,
        FlexibleBenefitValidationResult result)
    {
        foreach (var agreement in extraction.Agreements)
        {
            ValidateSignatureEvidence(extraction, agreement, "SubscriberSignaturePresent", agreement.SubscriberSignaturePresent, result);
            ValidateSignatureEvidence(extraction, agreement, "AdministratorSignaturePresent", agreement.AdministratorSignaturePresent, result);
            if (agreement.SubscriberSignaturePresent != true || agreement.AdministratorSignaturePresent != true)
            {
                Add(result, "agreement_signature_needs_review", "Warning", "FlexibleBenefitAgreement.Signatures", "One or more agreement signatures are missing or unclear.", agreement.SourcePageNumber);
            }
        }
    }

    private static void ValidateSignatureEvidence(
        FlexibleBenefitExtractionResult extraction,
        FlexibleBenefitAgreement agreement,
        string fieldName,
        bool? signaturePresent,
        FlexibleBenefitValidationResult result)
    {
        if (signaturePresent != true)
        {
            return;
        }

        var evidence = extraction.SourceFields.FirstOrDefault(field =>
            field.EntityKey == agreement.SourceKey && field.FieldName.Equals(fieldName, StringComparison.OrdinalIgnoreCase));
        if (evidence is null || evidence.ExtractionConfidence < 0.70m || string.IsNullOrWhiteSpace(evidence.BoundingBoxJson))
        {
            Add(result, "signature_evidence_missing", "Error", $"FlexibleBenefitAgreement.{fieldName}", "Signature presence cannot be confirmed from a typed name alone.", agreement.SourcePageNumber);
        }
    }

    private static void ValidateLineage(
        FlexibleBenefitExtractionResult extraction,
        FlexibleBenefitValidationResult result)
    {
        if (extraction.SourceFields.Count == 0)
        {
            Add(result, "source_lineage_missing", "Error", "FlexibleBenefitPlanSourceField", "Field-level source lineage is required before import.");
            return;
        }

        foreach (var field in extraction.SourceFields.Where(field =>
                     field.ReviewStatus.Equals(FlexibleBenefitReviewStatuses.NeedsReview, StringComparison.OrdinalIgnoreCase)))
        {
            if (result.Issues.Any(issue =>
                    issue.FieldName.Equals(field.FieldName, StringComparison.OrdinalIgnoreCase) ||
                    issue.FieldName.EndsWith($".{field.FieldName}", StringComparison.OrdinalIgnoreCase)))
            {
                continue;
            }

            Add(result, "source_field_needs_review", "Warning", $"{field.EntityType}.{field.FieldName}", "An extracted source field requires review.", field.SourcePageNumber);
        }
    }

    private static void ValidatePercentage(decimal? value, string fieldName, FlexibleBenefitValidationResult result)
    {
        if (value.HasValue && (value < 0m || value > 100m))
        {
            Add(result, "percentage_out_of_range", "Error", fieldName, "Percentages must be between 0 and 100.");
        }
    }

    private static void ValidateNonNegative(decimal? value, string fieldName, FlexibleBenefitValidationResult result)
    {
        if (value < 0m)
        {
            Add(result, "negative_value", "Error", fieldName, "Values cannot be negative.");
        }
    }

    private static void ValidateNonNegative(int? value, string fieldName, FlexibleBenefitValidationResult result)
    {
        if (value < 0)
        {
            Add(result, "negative_value", "Error", fieldName, "Values cannot be negative.");
        }
    }

    private static void Add(
        FlexibleBenefitValidationResult result,
        string code,
        string severity,
        string fieldName,
        string message,
        int? sourcePageNumber = null)
    {
        if (result.Issues.Any(issue =>
                issue.Code.Equals(code, StringComparison.OrdinalIgnoreCase) &&
                issue.FieldName.Equals(fieldName, StringComparison.OrdinalIgnoreCase)))
        {
            return;
        }

        result.Issues.Add(new FlexibleBenefitValidationIssue
        {
            Code = code,
            Severity = severity,
            FieldName = fieldName,
            Message = message,
            SourcePageNumber = sourcePageNumber
        });
    }
}

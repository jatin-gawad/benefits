using Microsoft.Data.Sqlite;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed partial class DatabaseService
{
    private static async Task InsertFlexibleBenefitChildrenAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long planId,
        long sourceDocumentId,
        FlexibleBenefitExtractionResult extraction,
        Dictionary<string, long> entityIds)
    {
        foreach (var item in extraction.Classes)
        {
            var id = await InsertFlexibleBenefitClassAsync(connection, transaction, planId, item);
            RegisterFlexibleBenefitEntity(entityIds, item.SourceKey, id);
        }

        foreach (var item in extraction.Categories)
        {
            var id = await InsertFlexibleBenefitCategoryAsync(connection, transaction, planId, item);
            RegisterFlexibleBenefitEntity(entityIds, item.SourceKey, id);
        }

        foreach (var item in extraction.EligibilityRules)
        {
            var classId = ResolveFlexibleBenefitEntityId(entityIds, item.ClassSourceKey);
            var id = await InsertFlexibleBenefitEligibilityRuleAsync(connection, transaction, planId, classId, item);
            RegisterFlexibleBenefitEntity(entityIds, item.SourceKey, id);
        }

        foreach (var item in extraction.FundingConfigurations)
        {
            var id = await InsertFlexibleBenefitFundingAsync(connection, transaction, planId, item);
            RegisterFlexibleBenefitEntity(entityIds, item.SourceKey, id);
        }

        foreach (var item in extraction.Fees)
        {
            var id = await InsertFlexibleBenefitFeeAsync(connection, transaction, planId, item);
            RegisterFlexibleBenefitEntity(entityIds, item.SourceKey, id);
        }

        foreach (var item in extraction.Agreements)
        {
            var id = await InsertFlexibleBenefitAgreementAsync(connection, transaction, planId, sourceDocumentId, item);
            RegisterFlexibleBenefitEntity(entityIds, item.SourceKey, id);
        }

        foreach (var item in extraction.Contacts)
        {
            var id = await InsertFlexibleBenefitContactAsync(connection, transaction, planId, item);
            RegisterFlexibleBenefitEntity(entityIds, item.SourceKey, id);
        }
    }

    private static async Task<long> InsertFlexibleBenefitClassAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long planId,
        FlexibleBenefitClass item)
    {
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO FlexibleBenefitClass (
    flexible_benefit_plan_id, external_class_code, class_name, job_description,
    eligibility_description, annual_single_amount, annual_couple_amount, annual_family_amount,
    annual_employee_amount, annual_dependent_amount, waiting_period_days,
    new_hire_proration_enabled, hcsa_amount, lsa_amount, hcsa_percentage, lsa_percentage,
    currency_code, effective_date, termination_date, is_current, source_page_number,
    extraction_confidence, review_status
) VALUES (
    $plan_id, $external_class_code, $class_name, $job_description,
    $eligibility_description, $annual_single_amount, $annual_couple_amount, $annual_family_amount,
    $annual_employee_amount, $annual_dependent_amount, $waiting_period_days,
    $new_hire_proration_enabled, $hcsa_amount, $lsa_amount, $hcsa_percentage, $lsa_percentage,
    $currency_code, $effective_date, $termination_date, $is_current, $source_page_number,
    $extraction_confidence, $review_status
);");
        AddParameter(command, "$plan_id", planId);
        AddParameter(command, "$external_class_code", Clean(item.ExternalClassCode));
        AddParameter(command, "$class_name", Clean(item.ClassName));
        AddParameter(command, "$job_description", Clean(item.JobDescription));
        AddParameter(command, "$eligibility_description", Clean(item.EligibilityDescription));
        AddParameter(command, "$annual_single_amount", item.AnnualSingleAmount);
        AddParameter(command, "$annual_couple_amount", item.AnnualCoupleAmount);
        AddParameter(command, "$annual_family_amount", item.AnnualFamilyAmount);
        AddParameter(command, "$annual_employee_amount", item.AnnualEmployeeAmount);
        AddParameter(command, "$annual_dependent_amount", item.AnnualDependentAmount);
        AddParameter(command, "$waiting_period_days", item.WaitingPeriodDays);
        AddParameter(command, "$new_hire_proration_enabled", DbBool(item.NewHireProrationEnabled));
        AddParameter(command, "$hcsa_amount", item.HcsaAmount);
        AddParameter(command, "$lsa_amount", item.LsaAmount);
        AddParameter(command, "$hcsa_percentage", item.HcsaPercentage);
        AddParameter(command, "$lsa_percentage", item.LsaPercentage);
        AddParameter(command, "$currency_code", FirstNonBlank(item.CurrencyCode, "CAD"));
        AddParameter(command, "$effective_date", ToDateText(item.EffectiveDate));
        AddParameter(command, "$termination_date", ToDateText(item.TerminationDate));
        AddParameter(command, "$is_current", item.IsCurrent ? 1 : 0);
        AddParameter(command, "$source_page_number", item.SourcePageNumber);
        AddParameter(command, "$extraction_confidence", item.ExtractionConfidence);
        AddParameter(command, "$review_status", Clean(item.ReviewStatus));
        await command.ExecuteNonQueryAsync();
        return await LastInsertRowIdAsync(connection, transaction);
    }

    private static async Task<long> InsertFlexibleBenefitCategoryAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long planId,
        FlexibleBenefitCategory item)
    {
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO FlexibleBenefitCategory (
    flexible_benefit_plan_id, category_code, category_group, category_name, description,
    enabled, taxable, account_type, requires_receipt, requires_provider_details, exclusions,
    source_page_number, extraction_confidence, review_status
) VALUES (
    $plan_id, $category_code, $category_group, $category_name, $description,
    $enabled, $taxable, $account_type, $requires_receipt, $requires_provider_details, $exclusions,
    $source_page_number, $extraction_confidence, $review_status
);");
        AddParameter(command, "$plan_id", planId);
        AddParameter(command, "$category_code", Clean(item.CategoryCode));
        AddParameter(command, "$category_group", Clean(item.CategoryGroup));
        AddParameter(command, "$category_name", Clean(item.CategoryName));
        AddParameter(command, "$description", Clean(item.Description));
        AddParameter(command, "$enabled", DbBool(item.Enabled));
        AddParameter(command, "$taxable", DbBool(item.Taxable));
        AddParameter(command, "$account_type", item.AccountType.ToString());
        AddParameter(command, "$requires_receipt", DbBool(item.RequiresReceipt));
        AddParameter(command, "$requires_provider_details", DbBool(item.RequiresProviderDetails));
        AddParameter(command, "$exclusions", Clean(item.Exclusions));
        AddParameter(command, "$source_page_number", item.SourcePageNumber);
        AddParameter(command, "$extraction_confidence", item.ExtractionConfidence);
        AddParameter(command, "$review_status", Clean(item.ReviewStatus));
        await command.ExecuteNonQueryAsync();
        return await LastInsertRowIdAsync(connection, transaction);
    }

    private static async Task<long> InsertFlexibleBenefitEligibilityRuleAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long planId,
        long? classId,
        FlexibleBenefitEligibilityRule item)
    {
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO FlexibleBenefitEligibilityRule (
    flexible_benefit_plan_id, flexible_benefit_class_id, rule_type, operator, rule_value,
    rule_unit, description, effective_date, source_page_number, extraction_confidence, review_status
) VALUES (
    $plan_id, $class_id, $rule_type, $operator, $rule_value,
    $rule_unit, $description, $effective_date, $source_page_number, $extraction_confidence, $review_status
);");
        AddParameter(command, "$plan_id", planId);
        AddParameter(command, "$class_id", classId);
        AddParameter(command, "$rule_type", Clean(item.RuleType));
        AddParameter(command, "$operator", Clean(item.Operator));
        AddParameter(command, "$rule_value", Clean(item.RuleValue));
        AddParameter(command, "$rule_unit", Clean(item.RuleUnit));
        AddParameter(command, "$description", Clean(item.Description));
        AddParameter(command, "$effective_date", ToDateText(item.EffectiveDate));
        AddParameter(command, "$source_page_number", item.SourcePageNumber);
        AddParameter(command, "$extraction_confidence", item.ExtractionConfidence);
        AddParameter(command, "$review_status", Clean(item.ReviewStatus));
        await command.ExecuteNonQueryAsync();
        return await LastInsertRowIdAsync(connection, transaction);
    }

    private static async Task<long> InsertFlexibleBenefitFundingAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long planId,
        FlexibleBenefitFundingConfiguration item)
    {
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO FlexibleBenefitFundingConfiguration (
    flexible_benefit_plan_id, funding_method, pre_funding_required, pad_selected,
    continuous_float_required, online_bank_payment_allowed, cheque_allowed,
    payment_timing_description, claims_paid_only_after_funds_clear, interest_paid_on_funds,
    currency_code, effective_date, source_page_number, extraction_confidence, review_status
) VALUES (
    $plan_id, $funding_method, $pre_funding_required, $pad_selected,
    $continuous_float_required, $online_bank_payment_allowed, $cheque_allowed,
    $payment_timing_description, $claims_paid_only_after_funds_clear, $interest_paid_on_funds,
    $currency_code, $effective_date, $source_page_number, $extraction_confidence, $review_status
);");
        AddParameter(command, "$plan_id", planId);
        AddParameter(command, "$funding_method", item.FundingMethod.ToString());
        AddParameter(command, "$pre_funding_required", DbBool(item.PreFundingRequired));
        AddParameter(command, "$pad_selected", DbBool(item.PadSelected));
        AddParameter(command, "$continuous_float_required", DbBool(item.ContinuousFloatRequired));
        AddParameter(command, "$online_bank_payment_allowed", DbBool(item.OnlineBankPaymentAllowed));
        AddParameter(command, "$cheque_allowed", DbBool(item.ChequeAllowed));
        AddParameter(command, "$payment_timing_description", Clean(item.PaymentTimingDescription));
        AddParameter(command, "$claims_paid_only_after_funds_clear", DbBool(item.ClaimsPaidOnlyAfterFundsClear));
        AddParameter(command, "$interest_paid_on_funds", DbBool(item.InterestPaidOnFunds));
        AddParameter(command, "$currency_code", FirstNonBlank(item.CurrencyCode, "CAD"));
        AddParameter(command, "$effective_date", ToDateText(item.EffectiveDate));
        AddParameter(command, "$source_page_number", item.SourcePageNumber);
        AddParameter(command, "$extraction_confidence", item.ExtractionConfidence);
        AddParameter(command, "$review_status", Clean(item.ReviewStatus));
        await command.ExecuteNonQueryAsync();
        return await LastInsertRowIdAsync(connection, transaction);
    }

    private static async Task<long> InsertFlexibleBenefitFeeAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long planId,
        FlexibleBenefitFee item)
    {
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO FlexibleBenefitFee (
    flexible_benefit_plan_id, fee_type, fee_calculation_type, percentage, flat_amount,
    minimum_amount, maximum_amount, taxable, tax_type, description, effective_date,
    notice_period_days, source_page_number, extraction_confidence, review_status
) VALUES (
    $plan_id, $fee_type, $fee_calculation_type, $percentage, $flat_amount,
    $minimum_amount, $maximum_amount, $taxable, $tax_type, $description, $effective_date,
    $notice_period_days, $source_page_number, $extraction_confidence, $review_status
);");
        AddParameter(command, "$plan_id", planId);
        AddParameter(command, "$fee_type", Clean(item.FeeType));
        AddParameter(command, "$fee_calculation_type", Clean(item.FeeCalculationType));
        AddParameter(command, "$percentage", item.Percentage);
        AddParameter(command, "$flat_amount", item.FlatAmount);
        AddParameter(command, "$minimum_amount", item.MinimumAmount);
        AddParameter(command, "$maximum_amount", item.MaximumAmount);
        AddParameter(command, "$taxable", DbBool(item.Taxable));
        AddParameter(command, "$tax_type", Clean(item.TaxType));
        AddParameter(command, "$description", Clean(item.Description));
        AddParameter(command, "$effective_date", ToDateText(item.EffectiveDate));
        AddParameter(command, "$notice_period_days", item.NoticePeriodDays);
        AddParameter(command, "$source_page_number", item.SourcePageNumber);
        AddParameter(command, "$extraction_confidence", item.ExtractionConfidence);
        AddParameter(command, "$review_status", Clean(item.ReviewStatus));
        await command.ExecuteNonQueryAsync();
        return await LastInsertRowIdAsync(connection, transaction);
    }

    private static async Task<long> InsertFlexibleBenefitAgreementAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long planId,
        long sourceDocumentId,
        FlexibleBenefitAgreement item)
    {
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO FlexibleBenefitAgreement (
    flexible_benefit_plan_id, agreement_type, agreement_version, agreement_date, execution_date,
    governing_province, subscriber_legal_name, administrator_legal_name,
    subscriber_signer_name, subscriber_signer_title, administrator_signer_name,
    administrator_signer_title, subscriber_signature_present, administrator_signature_present,
    electronic_signature, signed_city, signed_province, indemnity_agreement_signed,
    funding_agreement_signed, agreement_status, source_document_id, source_page_number, review_status
) VALUES (
    $plan_id, $agreement_type, $agreement_version, $agreement_date, $execution_date,
    $governing_province, $subscriber_legal_name, $administrator_legal_name,
    $subscriber_signer_name, $subscriber_signer_title, $administrator_signer_name,
    $administrator_signer_title, $subscriber_signature_present, $administrator_signature_present,
    $electronic_signature, $signed_city, $signed_province, $indemnity_agreement_signed,
    $funding_agreement_signed, $agreement_status, $source_document_id, $source_page_number, $review_status
);");
        AddParameter(command, "$plan_id", planId);
        AddParameter(command, "$agreement_type", Clean(item.AgreementType));
        AddParameter(command, "$agreement_version", Clean(item.AgreementVersion));
        AddParameter(command, "$agreement_date", ToDateText(item.AgreementDate));
        AddParameter(command, "$execution_date", item.ExecutionDate.HasValue ? ToDateTimeText(item.ExecutionDate.Value) : null);
        AddParameter(command, "$governing_province", Clean(item.GoverningProvince));
        AddParameter(command, "$subscriber_legal_name", Clean(item.SubscriberLegalName));
        AddParameter(command, "$administrator_legal_name", Clean(item.AdministratorLegalName));
        AddParameter(command, "$subscriber_signer_name", Clean(item.SubscriberSignerName));
        AddParameter(command, "$subscriber_signer_title", Clean(item.SubscriberSignerTitle));
        AddParameter(command, "$administrator_signer_name", Clean(item.AdministratorSignerName));
        AddParameter(command, "$administrator_signer_title", Clean(item.AdministratorSignerTitle));
        AddParameter(command, "$subscriber_signature_present", DbBool(item.SubscriberSignaturePresent));
        AddParameter(command, "$administrator_signature_present", DbBool(item.AdministratorSignaturePresent));
        AddParameter(command, "$electronic_signature", DbBool(item.ElectronicSignature));
        AddParameter(command, "$signed_city", Clean(item.SignedCity));
        AddParameter(command, "$signed_province", Clean(item.SignedProvince));
        AddParameter(command, "$indemnity_agreement_signed", DbBool(item.IndemnityAgreementSigned));
        AddParameter(command, "$funding_agreement_signed", DbBool(item.FundingAgreementSigned));
        AddParameter(command, "$agreement_status", Clean(item.AgreementStatus));
        AddParameter(command, "$source_document_id", sourceDocumentId);
        AddParameter(command, "$source_page_number", item.SourcePageNumber);
        AddParameter(command, "$review_status", Clean(item.ReviewStatus));
        await command.ExecuteNonQueryAsync();
        return await LastInsertRowIdAsync(connection, transaction);
    }

    private static async Task<long> InsertFlexibleBenefitContactAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long planId,
        FlexibleBenefitContact item)
    {
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO FlexibleBenefitContact (
    flexible_benefit_plan_id, client_contact_id, contact_role, contact_name, email, phone,
    is_primary, source_page_number, extraction_confidence, review_status
) VALUES (
    $plan_id, $client_contact_id, $contact_role, $contact_name, $email, $phone,
    $is_primary, $source_page_number, $extraction_confidence, $review_status
);");
        AddParameter(command, "$plan_id", planId);
        AddParameter(command, "$client_contact_id", item.ClientContactId);
        AddParameter(command, "$contact_role", Clean(item.ContactRole));
        AddParameter(command, "$contact_name", Clean(item.ContactName));
        AddParameter(command, "$email", Clean(item.Email));
        AddParameter(command, "$phone", Clean(item.Phone));
        AddParameter(command, "$is_primary", item.IsPrimary ? 1 : 0);
        AddParameter(command, "$source_page_number", item.SourcePageNumber);
        AddParameter(command, "$extraction_confidence", item.ExtractionConfidence);
        AddParameter(command, "$review_status", Clean(item.ReviewStatus));
        await command.ExecuteNonQueryAsync();
        return await LastInsertRowIdAsync(connection, transaction);
    }

    private static void RegisterFlexibleBenefitEntity(IDictionary<string, long> entityIds, string sourceKey, long entityId)
    {
        if (!string.IsNullOrWhiteSpace(sourceKey))
        {
            entityIds[sourceKey] = entityId;
        }
    }

    private static long? ResolveFlexibleBenefitEntityId(IDictionary<string, long> entityIds, string sourceKey)
    {
        return !string.IsNullOrWhiteSpace(sourceKey) && entityIds.TryGetValue(sourceKey, out var entityId)
            ? entityId
            : null;
    }
}

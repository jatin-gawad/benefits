using Microsoft.Data.Sqlite;

namespace RenewalReviewBuilder.App.Services;

public sealed partial class DatabaseService
{
    private static async Task EnsureFlexibleBenefitSchemaAsync(SqliteConnection connection)
    {
        await EnsureColumnExistsAsync(connection, "Clients", "legal_name", "TEXT");
        await EnsureColumnExistsAsync(connection, "Clients", "dba_name", "TEXT");
        await EnsureColumnExistsAsync(connection, "Clients", "address", "TEXT");
        await EnsureColumnExistsAsync(connection, "Clients", "city", "TEXT");
        await EnsureColumnExistsAsync(connection, "Clients", "postal_code", "TEXT");
        await EnsureColumnExistsAsync(connection, "Clients", "phone", "TEXT");

        await ExecuteNonQueryAsync(connection, null, @"
CREATE TABLE IF NOT EXISTS FlexibleBenefitDocumentVersion (
    flexible_benefit_document_version_id INTEGER PRIMARY KEY AUTOINCREMENT,
    document_id INTEGER NOT NULL,
    client_id INTEGER NOT NULL,
    document_hash TEXT NOT NULL,
    document_subtype TEXT NOT NULL,
    effective_date TEXT,
    execution_date TEXT,
    version_number INTEGER NOT NULL,
    version_action TEXT NOT NULL,
    is_current INTEGER NOT NULL DEFAULT 1 CHECK (is_current IN (0, 1)),
    supersedes_document_version_id INTEGER,
    review_status TEXT NOT NULL DEFAULT 'Pending',
    created_at TEXT NOT NULL,
    created_by TEXT NOT NULL,
    FOREIGN KEY (document_id) REFERENCES Documents(document_id),
    FOREIGN KEY (client_id) REFERENCES Clients(client_id),
    FOREIGN KEY (supersedes_document_version_id) REFERENCES FlexibleBenefitDocumentVersion(flexible_benefit_document_version_id),
    UNIQUE (client_id, version_number, document_id)
);

CREATE INDEX IF NOT EXISTS IX_FlexibleBenefitDocumentVersion_ClientCurrent
ON FlexibleBenefitDocumentVersion(client_id, is_current, effective_date, version_number);

CREATE INDEX IF NOT EXISTS IX_FlexibleBenefitDocumentVersion_Hash
ON FlexibleBenefitDocumentVersion(document_hash, client_id);

CREATE TABLE IF NOT EXISTS FlexibleBenefitPlan (
    flexible_benefit_plan_id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL,
    carrier_id INTEGER,
    administrator_id INTEGER,
    source_document_id INTEGER NOT NULL,
    source_document_version_id INTEGER NOT NULL,
    policy_number TEXT,
    account_number TEXT,
    plan_name TEXT,
    plan_type TEXT NOT NULL,
    hcsa_enabled INTEGER CHECK (hcsa_enabled IN (0, 1) OR hcsa_enabled IS NULL),
    lsa_enabled INTEGER CHECK (lsa_enabled IN (0, 1) OR lsa_enabled IS NULL),
    wellness_account_enabled INTEGER CHECK (wellness_account_enabled IN (0, 1) OR wellness_account_enabled IS NULL),
    combined_account_enabled INTEGER CHECK (combined_account_enabled IN (0, 1) OR combined_account_enabled IS NULL),
    effective_date TEXT,
    termination_date TEXT,
    benefit_year_type TEXT NOT NULL,
    benefit_year_start_month INTEGER,
    benefit_year_start_day INTEGER,
    default_allocation_type TEXT NOT NULL,
    default_hcsa_percentage NUMERIC,
    default_lsa_percentage NUMERIC,
    pay_as_submitted INTEGER CHECK (pay_as_submitted IN (0, 1) OR pay_as_submitted IS NULL),
    runoff_days INTEGER,
    carry_forward_type TEXT NOT NULL,
    credit_carry_forward_enabled INTEGER CHECK (credit_carry_forward_enabled IN (0, 1) OR credit_carry_forward_enabled IS NULL),
    expense_carry_forward_enabled INTEGER CHECK (expense_carry_forward_enabled IN (0, 1) OR expense_carry_forward_enabled IS NULL),
    carry_forward_maximum NUMERIC,
    carry_forward_expiry_months INTEGER,
    default_waiting_period_days INTEGER,
    new_hire_proration_default INTEGER CHECK (new_hire_proration_default IN (0, 1) OR new_hire_proration_default IS NULL),
    child_dependent_cutoff_age INTEGER,
    student_dependent_cutoff_age INTEGER,
    funding_method TEXT NOT NULL,
    pre_funded INTEGER CHECK (pre_funded IN (0, 1) OR pre_funded IS NULL),
    pay_as_you_go INTEGER CHECK (pay_as_you_go IN (0, 1) OR pay_as_you_go IS NULL),
    pad_requested INTEGER CHECK (pad_requested IN (0, 1) OR pad_requested IS NULL),
    setup_fee NUMERIC,
    administration_fee_percentage NUMERIC,
    administration_fee_flat_amount NUMERIC,
    administration_fee_basis TEXT,
    tax_treatment TEXT,
    currency_code TEXT NOT NULL DEFAULT 'CAD',
    is_current INTEGER NOT NULL DEFAULT 1 CHECK (is_current IN (0, 1)),
    version_number INTEGER NOT NULL DEFAULT 1,
    supersedes_plan_id INTEGER,
    review_status TEXT NOT NULL DEFAULT 'Pending',
    extraction_status TEXT NOT NULL DEFAULT 'Extracted',
    created_at TEXT NOT NULL,
    created_by TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    updated_by TEXT NOT NULL,
    FOREIGN KEY (client_id) REFERENCES Clients(client_id),
    FOREIGN KEY (carrier_id) REFERENCES Carrier(carrier_id),
    FOREIGN KEY (administrator_id) REFERENCES Carrier(carrier_id),
    FOREIGN KEY (source_document_id) REFERENCES Documents(document_id),
    FOREIGN KEY (source_document_version_id) REFERENCES FlexibleBenefitDocumentVersion(flexible_benefit_document_version_id),
    FOREIGN KEY (supersedes_plan_id) REFERENCES FlexibleBenefitPlan(flexible_benefit_plan_id),
    UNIQUE (source_document_version_id)
);

CREATE INDEX IF NOT EXISTS IX_FlexibleBenefitPlan_ClientCurrent
ON FlexibleBenefitPlan(client_id, is_current, effective_date, version_number);

CREATE TABLE IF NOT EXISTS FlexibleBenefitClass (
    flexible_benefit_class_id INTEGER PRIMARY KEY AUTOINCREMENT,
    flexible_benefit_plan_id INTEGER NOT NULL,
    external_class_code TEXT,
    class_name TEXT NOT NULL,
    job_description TEXT,
    eligibility_description TEXT,
    annual_single_amount NUMERIC,
    annual_couple_amount NUMERIC,
    annual_family_amount NUMERIC,
    annual_employee_amount NUMERIC,
    annual_dependent_amount NUMERIC,
    waiting_period_days INTEGER,
    new_hire_proration_enabled INTEGER CHECK (new_hire_proration_enabled IN (0, 1) OR new_hire_proration_enabled IS NULL),
    hcsa_amount NUMERIC,
    lsa_amount NUMERIC,
    hcsa_percentage NUMERIC,
    lsa_percentage NUMERIC,
    currency_code TEXT NOT NULL DEFAULT 'CAD',
    effective_date TEXT,
    termination_date TEXT,
    is_current INTEGER NOT NULL DEFAULT 1 CHECK (is_current IN (0, 1)),
    source_page_number INTEGER,
    extraction_confidence NUMERIC,
    review_status TEXT NOT NULL DEFAULT 'Pending',
    FOREIGN KEY (flexible_benefit_plan_id) REFERENCES FlexibleBenefitPlan(flexible_benefit_plan_id)
);

CREATE INDEX IF NOT EXISTS IX_FlexibleBenefitClass_Plan
ON FlexibleBenefitClass(flexible_benefit_plan_id, is_current, class_name);

CREATE TABLE IF NOT EXISTS FlexibleBenefitCategory (
    flexible_benefit_category_id INTEGER PRIMARY KEY AUTOINCREMENT,
    flexible_benefit_plan_id INTEGER NOT NULL,
    category_code TEXT NOT NULL,
    category_group TEXT NOT NULL,
    category_name TEXT NOT NULL,
    description TEXT,
    enabled INTEGER CHECK (enabled IN (0, 1) OR enabled IS NULL),
    taxable INTEGER CHECK (taxable IN (0, 1) OR taxable IS NULL),
    account_type TEXT NOT NULL,
    requires_receipt INTEGER CHECK (requires_receipt IN (0, 1) OR requires_receipt IS NULL),
    requires_provider_details INTEGER CHECK (requires_provider_details IN (0, 1) OR requires_provider_details IS NULL),
    exclusions TEXT,
    source_page_number INTEGER,
    extraction_confidence NUMERIC,
    review_status TEXT NOT NULL DEFAULT 'Pending',
    FOREIGN KEY (flexible_benefit_plan_id) REFERENCES FlexibleBenefitPlan(flexible_benefit_plan_id),
    UNIQUE (flexible_benefit_plan_id, category_code)
);

CREATE INDEX IF NOT EXISTS IX_FlexibleBenefitCategory_PlanGroup
ON FlexibleBenefitCategory(flexible_benefit_plan_id, category_group, enabled);

CREATE TABLE IF NOT EXISTS FlexibleBenefitEligibilityRule (
    flexible_benefit_eligibility_rule_id INTEGER PRIMARY KEY AUTOINCREMENT,
    flexible_benefit_plan_id INTEGER NOT NULL,
    flexible_benefit_class_id INTEGER,
    rule_type TEXT NOT NULL,
    operator TEXT,
    rule_value TEXT,
    rule_unit TEXT,
    description TEXT,
    effective_date TEXT,
    source_page_number INTEGER,
    extraction_confidence NUMERIC,
    review_status TEXT NOT NULL DEFAULT 'Pending',
    FOREIGN KEY (flexible_benefit_plan_id) REFERENCES FlexibleBenefitPlan(flexible_benefit_plan_id),
    FOREIGN KEY (flexible_benefit_class_id) REFERENCES FlexibleBenefitClass(flexible_benefit_class_id)
);

CREATE INDEX IF NOT EXISTS IX_FlexibleBenefitEligibilityRule_Plan
ON FlexibleBenefitEligibilityRule(flexible_benefit_plan_id, flexible_benefit_class_id, rule_type);

CREATE TABLE IF NOT EXISTS FlexibleBenefitFundingConfiguration (
    flexible_benefit_funding_configuration_id INTEGER PRIMARY KEY AUTOINCREMENT,
    flexible_benefit_plan_id INTEGER NOT NULL,
    funding_method TEXT NOT NULL,
    pre_funding_required INTEGER CHECK (pre_funding_required IN (0, 1) OR pre_funding_required IS NULL),
    pad_selected INTEGER CHECK (pad_selected IN (0, 1) OR pad_selected IS NULL),
    continuous_float_required INTEGER CHECK (continuous_float_required IN (0, 1) OR continuous_float_required IS NULL),
    online_bank_payment_allowed INTEGER CHECK (online_bank_payment_allowed IN (0, 1) OR online_bank_payment_allowed IS NULL),
    cheque_allowed INTEGER CHECK (cheque_allowed IN (0, 1) OR cheque_allowed IS NULL),
    payment_timing_description TEXT,
    claims_paid_only_after_funds_clear INTEGER CHECK (claims_paid_only_after_funds_clear IN (0, 1) OR claims_paid_only_after_funds_clear IS NULL),
    interest_paid_on_funds INTEGER CHECK (interest_paid_on_funds IN (0, 1) OR interest_paid_on_funds IS NULL),
    currency_code TEXT NOT NULL DEFAULT 'CAD',
    effective_date TEXT,
    source_page_number INTEGER,
    extraction_confidence NUMERIC,
    review_status TEXT NOT NULL DEFAULT 'Pending',
    FOREIGN KEY (flexible_benefit_plan_id) REFERENCES FlexibleBenefitPlan(flexible_benefit_plan_id)
);

CREATE TABLE IF NOT EXISTS FlexibleBenefitFee (
    flexible_benefit_fee_id INTEGER PRIMARY KEY AUTOINCREMENT,
    flexible_benefit_plan_id INTEGER NOT NULL,
    fee_type TEXT NOT NULL,
    fee_calculation_type TEXT NOT NULL,
    percentage NUMERIC,
    flat_amount NUMERIC,
    minimum_amount NUMERIC,
    maximum_amount NUMERIC,
    taxable INTEGER CHECK (taxable IN (0, 1) OR taxable IS NULL),
    tax_type TEXT,
    description TEXT,
    effective_date TEXT,
    notice_period_days INTEGER,
    source_page_number INTEGER,
    extraction_confidence NUMERIC,
    review_status TEXT NOT NULL DEFAULT 'Pending',
    FOREIGN KEY (flexible_benefit_plan_id) REFERENCES FlexibleBenefitPlan(flexible_benefit_plan_id)
);

CREATE INDEX IF NOT EXISTS IX_FlexibleBenefitFee_Plan
ON FlexibleBenefitFee(flexible_benefit_plan_id, fee_type);

CREATE TABLE IF NOT EXISTS FlexibleBenefitAgreement (
    flexible_benefit_agreement_id INTEGER PRIMARY KEY AUTOINCREMENT,
    flexible_benefit_plan_id INTEGER NOT NULL,
    agreement_type TEXT NOT NULL,
    agreement_version TEXT,
    agreement_date TEXT,
    execution_date TEXT,
    governing_province TEXT,
    subscriber_legal_name TEXT,
    administrator_legal_name TEXT,
    subscriber_signer_name TEXT,
    subscriber_signer_title TEXT,
    administrator_signer_name TEXT,
    administrator_signer_title TEXT,
    subscriber_signature_present INTEGER CHECK (subscriber_signature_present IN (0, 1) OR subscriber_signature_present IS NULL),
    administrator_signature_present INTEGER CHECK (administrator_signature_present IN (0, 1) OR administrator_signature_present IS NULL),
    electronic_signature INTEGER CHECK (electronic_signature IN (0, 1) OR electronic_signature IS NULL),
    signed_city TEXT,
    signed_province TEXT,
    indemnity_agreement_signed INTEGER CHECK (indemnity_agreement_signed IN (0, 1) OR indemnity_agreement_signed IS NULL),
    funding_agreement_signed INTEGER CHECK (funding_agreement_signed IN (0, 1) OR funding_agreement_signed IS NULL),
    agreement_status TEXT,
    source_document_id INTEGER NOT NULL,
    source_page_number INTEGER,
    review_status TEXT NOT NULL DEFAULT 'Pending',
    FOREIGN KEY (flexible_benefit_plan_id) REFERENCES FlexibleBenefitPlan(flexible_benefit_plan_id),
    FOREIGN KEY (source_document_id) REFERENCES Documents(document_id)
);

CREATE TABLE IF NOT EXISTS FlexibleBenefitContact (
    flexible_benefit_contact_id INTEGER PRIMARY KEY AUTOINCREMENT,
    flexible_benefit_plan_id INTEGER NOT NULL,
    client_contact_id INTEGER,
    contact_role TEXT NOT NULL,
    contact_name TEXT,
    email TEXT,
    phone TEXT,
    is_primary INTEGER NOT NULL DEFAULT 0 CHECK (is_primary IN (0, 1)),
    source_page_number INTEGER,
    extraction_confidence NUMERIC,
    review_status TEXT NOT NULL DEFAULT 'Pending',
    FOREIGN KEY (flexible_benefit_plan_id) REFERENCES FlexibleBenefitPlan(flexible_benefit_plan_id)
);

CREATE INDEX IF NOT EXISTS IX_FlexibleBenefitContact_PlanRole
ON FlexibleBenefitContact(flexible_benefit_plan_id, contact_role);

CREATE TABLE IF NOT EXISTS FlexibleBenefitPlanSourceField (
    flexible_benefit_plan_source_field_id INTEGER PRIMARY KEY AUTOINCREMENT,
    source_document_id INTEGER NOT NULL,
    source_document_version_id INTEGER NOT NULL,
    entity_type TEXT NOT NULL,
    entity_id INTEGER,
    field_name TEXT NOT NULL,
    raw_value TEXT,
    normalized_value TEXT,
    source_page_number INTEGER,
    bounding_box_json TEXT,
    extraction_method TEXT NOT NULL,
    extraction_confidence NUMERIC,
    was_manually_corrected INTEGER NOT NULL DEFAULT 0 CHECK (was_manually_corrected IN (0, 1)),
    corrected_by TEXT,
    corrected_at TEXT,
    review_status TEXT NOT NULL DEFAULT 'Pending',
    FOREIGN KEY (source_document_id) REFERENCES Documents(document_id),
    FOREIGN KEY (source_document_version_id) REFERENCES FlexibleBenefitDocumentVersion(flexible_benefit_document_version_id)
);

CREATE INDEX IF NOT EXISTS IX_FlexibleBenefitPlanSourceField_Entity
ON FlexibleBenefitPlanSourceField(entity_type, entity_id, field_name);

CREATE INDEX IF NOT EXISTS IX_FlexibleBenefitPlanSourceField_DocumentPage
ON FlexibleBenefitPlanSourceField(source_document_id, source_document_version_id, source_page_number);

CREATE TABLE IF NOT EXISTS FlexibleBenefitCorrectionAudit (
    flexible_benefit_correction_audit_id INTEGER PRIMARY KEY AUTOINCREMENT,
    flexible_benefit_plan_source_field_id INTEGER NOT NULL,
    source_document_id INTEGER NOT NULL,
    source_document_version_id INTEGER NOT NULL,
    entity_type TEXT NOT NULL,
    entity_id INTEGER,
    field_name TEXT NOT NULL,
    previous_normalized_value TEXT,
    new_normalized_value TEXT,
    reason TEXT,
    corrected_at TEXT NOT NULL,
    corrected_by TEXT NOT NULL,
    FOREIGN KEY (flexible_benefit_plan_source_field_id) REFERENCES FlexibleBenefitPlanSourceField(flexible_benefit_plan_source_field_id),
    FOREIGN KEY (source_document_id) REFERENCES Documents(document_id),
    FOREIGN KEY (source_document_version_id) REFERENCES FlexibleBenefitDocumentVersion(flexible_benefit_document_version_id)
);

CREATE INDEX IF NOT EXISTS IX_FlexibleBenefitCorrectionAudit_SourceField
ON FlexibleBenefitCorrectionAudit(flexible_benefit_plan_source_field_id, corrected_at);

CREATE TABLE IF NOT EXISTS FlexibleBenefitExtractionStaging (
    flexible_benefit_extraction_staging_id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL,
    source_document_id INTEGER NOT NULL,
    source_document_version_id INTEGER NOT NULL,
    document_hash TEXT NOT NULL,
    carrier TEXT NOT NULL,
    document_type TEXT NOT NULL,
    raw_extraction_json TEXT NOT NULL,
    extraction_status TEXT NOT NULL,
    review_status TEXT NOT NULL,
    created_at TEXT NOT NULL,
    created_by TEXT NOT NULL,
    FOREIGN KEY (client_id) REFERENCES Clients(client_id),
    FOREIGN KEY (source_document_id) REFERENCES Documents(document_id),
    FOREIGN KEY (source_document_version_id) REFERENCES FlexibleBenefitDocumentVersion(flexible_benefit_document_version_id)
);

CREATE INDEX IF NOT EXISTS IX_FlexibleBenefitExtractionStaging_Document
ON FlexibleBenefitExtractionStaging(source_document_id, source_document_version_id);
");
    }
}

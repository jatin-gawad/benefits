using System.Globalization;
using Microsoft.Data.Sqlite;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed partial class DatabaseService
{
    private static async Task<bool> FlexibleBenefitClientExistsAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long clientId)
    {
        using var command = CreateCommand(connection, transaction, "SELECT COUNT(*) FROM Clients WHERE client_id = $client_id;");
        AddParameter(command, "$client_id", clientId);
        return Convert.ToInt32(await command.ExecuteScalarAsync()) > 0;
    }

    private static async Task<bool> SameFlexibleBenefitHashExistsForAnotherClientAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long clientId,
        string fileHash)
    {
        using var command = CreateCommand(connection, transaction, @"
SELECT COUNT(*)
FROM Documents
WHERE document_type = $document_type
  AND file_hash = $file_hash
  AND client_id <> $client_id;");
        AddParameter(command, "$document_type", DocumentTypes.FlexibleBenefitsPlan);
        AddParameter(command, "$file_hash", fileHash);
        AddParameter(command, "$client_id", clientId);
        return Convert.ToInt32(await command.ExecuteScalarAsync()) > 0;
    }

    private static async Task<FlexibleBenefitCurrentPlan?> FindCurrentFlexibleBenefitPlanAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long clientId,
        DateOnly? effectiveDate)
    {
        using var command = CreateCommand(connection, transaction, @"
SELECT flexible_benefit_plan_id,
       source_document_version_id,
       version_number,
       effective_date
FROM FlexibleBenefitPlan
WHERE client_id = $client_id
  AND is_current = 1
  AND ($filter_period = 0 OR COALESCE(effective_date, '') = COALESCE($effective_date, ''))
ORDER BY version_number DESC
LIMIT 1;");
        AddParameter(command, "$client_id", clientId);
        AddParameter(command, "$filter_period", effectiveDate.HasValue ? 1 : 0);
        AddParameter(command, "$effective_date", ToDateText(effectiveDate));
        using var reader = await command.ExecuteReaderAsync();
        if (!await reader.ReadAsync())
        {
            return null;
        }

        return new FlexibleBenefitCurrentPlan(
            reader.GetInt64(0),
            reader.GetInt64(1),
            reader.GetInt32(2),
            reader.IsDBNull(3) ? null : DateOnly.Parse(reader.GetString(3), CultureInfo.InvariantCulture));
    }

    private static async Task<int> NextFlexibleBenefitVersionNumberAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long clientId)
    {
        using var command = CreateCommand(connection, transaction, @"
SELECT COALESCE(MAX(version_number), 0) + 1
FROM FlexibleBenefitPlan
WHERE client_id = $client_id;");
        AddParameter(command, "$client_id", clientId);
        return Convert.ToInt32(await command.ExecuteScalarAsync(), CultureInfo.InvariantCulture);
    }

    private static async Task UpdateFlexibleBenefitClientIdentityAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long clientId,
        FlexibleBenefitClientIdentity identity)
    {
        using var command = CreateCommand(connection, transaction, @"
UPDATE Clients
SET legal_name = COALESCE(NULLIF(TRIM(legal_name), ''), NULLIF($legal_name, '')),
    dba_name = COALESCE(NULLIF(TRIM(dba_name), ''), NULLIF($dba_name, '')),
    address = COALESCE(NULLIF(TRIM(address), ''), NULLIF($address, '')),
    city = COALESCE(NULLIF(TRIM(city), ''), NULLIF($city, '')),
    province = COALESCE(NULLIF(TRIM(province), ''), NULLIF($province, '')),
    postal_code = COALESCE(NULLIF(TRIM(postal_code), ''), NULLIF($postal_code, '')),
    phone = COALESCE(NULLIF(TRIM(phone), ''), NULLIF($phone, '')),
    date_updated = $date_updated
WHERE client_id = $client_id;");
        AddParameter(command, "$legal_name", Clean(identity.LegalName));
        AddParameter(command, "$dba_name", Clean(identity.DbaName));
        AddParameter(command, "$address", Clean(identity.Address));
        AddParameter(command, "$city", Clean(identity.City));
        AddParameter(command, "$province", Clean(identity.Province));
        AddParameter(command, "$postal_code", Clean(identity.PostalCode));
        AddParameter(command, "$phone", Clean(identity.Phone));
        AddParameter(command, "$date_updated", ToDateTimeText(DateTime.Now));
        AddParameter(command, "$client_id", clientId);
        await command.ExecuteNonQueryAsync();
    }

    private static async Task<long?> FindCarrierIdByNameAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        string carrierName)
    {
        if (string.IsNullOrWhiteSpace(carrierName))
        {
            return null;
        }

        using var command = CreateCommand(connection, transaction, @"
SELECT carrier_id
FROM Carrier
WHERE carrier_name = $carrier_name COLLATE NOCASE
LIMIT 1;");
        AddParameter(command, "$carrier_name", Clean(carrierName));
        return ToNullableLong(await command.ExecuteScalarAsync());
    }

    private static async Task MarkFlexibleBenefitVersionHistoricalAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        FlexibleBenefitCurrentPlan prior,
        bool replaced)
    {
        await ExecuteNonQueryAsync(connection, transaction, @"
UPDATE FlexibleBenefitPlan
SET is_current = 0,
    review_status = CASE WHEN $replaced = 1 THEN $replaced_status ELSE review_status END,
    updated_at = $updated_at
WHERE flexible_benefit_plan_id = $plan_id;", command =>
        {
            AddParameter(command, "$replaced", replaced ? 1 : 0);
            AddParameter(command, "$replaced_status", FlexibleBenefitReviewStatuses.Replaced);
            AddParameter(command, "$updated_at", ToDateTimeText(DateTime.Now));
            AddParameter(command, "$plan_id", prior.PlanId);
        });
        await ExecuteNonQueryAsync(connection, transaction, @"
UPDATE FlexibleBenefitDocumentVersion
SET is_current = 0,
    review_status = CASE WHEN $replaced = 1 THEN $replaced_status ELSE review_status END
WHERE flexible_benefit_document_version_id = $document_version_id;", command =>
        {
            AddParameter(command, "$replaced", replaced ? 1 : 0);
            AddParameter(command, "$replaced_status", FlexibleBenefitReviewStatuses.Replaced);
            AddParameter(command, "$document_version_id", prior.DocumentVersionId);
        });
    }

    private static async Task<long> InsertFlexibleBenefitDocumentVersionAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long documentId,
        long clientId,
        FlexibleBenefitExtractionResult extraction,
        int versionNumber,
        long? supersedesDocumentVersionId,
        FlexibleBenefitImportRequest request,
        string finalStatus)
    {
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO FlexibleBenefitDocumentVersion (
    document_id, client_id, document_hash, document_subtype, effective_date, execution_date,
    version_number, version_action, is_current, supersedes_document_version_id,
    review_status, created_at, created_by
)
VALUES (
    $document_id, $client_id, $document_hash, $document_subtype, $effective_date, $execution_date,
    $version_number, $version_action, 1, $supersedes_document_version_id,
    $review_status, $created_at, $created_by
);");
        AddParameter(command, "$document_id", documentId);
        AddParameter(command, "$client_id", clientId);
        AddParameter(command, "$document_hash", extraction.SourceHash.ToUpperInvariant());
        AddParameter(command, "$document_subtype", extraction.Classification.DocumentSubtype.ToString());
        AddParameter(command, "$effective_date", ToDateText(extraction.Plan.EffectiveDate));
        var executionDate = extraction.Agreements.FirstOrDefault()?.ExecutionDate;
        AddParameter(command, "$execution_date", executionDate.HasValue ? ToDateTimeText(executionDate.Value) : null);
        AddParameter(command, "$version_number", versionNumber);
        AddParameter(command, "$version_action", supersedesDocumentVersionId.HasValue ? request.ImportMode.ToString() : "InitialImport");
        AddParameter(command, "$supersedes_document_version_id", supersedesDocumentVersionId);
        AddParameter(command, "$review_status", ToFlexibleReviewStatus(finalStatus));
        AddParameter(command, "$created_at", ToDateTimeText(DateTime.Now));
        AddParameter(command, "$created_by", Clean(request.ImportedBy));
        await command.ExecuteNonQueryAsync();
        return await LastInsertRowIdAsync(connection, transaction);
    }

    private static async Task<long> InsertFlexibleBenefitPlanAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long clientId,
        long? carrierId,
        long? administratorId,
        long documentId,
        long documentVersionId,
        long? supersedesPlanId,
        int versionNumber,
        FlexibleBenefitImportRequest request,
        FlexibleBenefitExtractionResult extraction,
        string finalStatus)
    {
        var plan = extraction.Plan;
        using var command = CreateCommand(connection, transaction, @"
INSERT INTO FlexibleBenefitPlan (
    client_id, carrier_id, administrator_id, source_document_id, source_document_version_id,
    policy_number, account_number, plan_name, plan_type, hcsa_enabled, lsa_enabled,
    wellness_account_enabled, combined_account_enabled, effective_date, termination_date,
    benefit_year_type, benefit_year_start_month, benefit_year_start_day, default_allocation_type,
    default_hcsa_percentage, default_lsa_percentage, pay_as_submitted, runoff_days,
    carry_forward_type, credit_carry_forward_enabled, expense_carry_forward_enabled,
    carry_forward_maximum, carry_forward_expiry_months, default_waiting_period_days,
    new_hire_proration_default, child_dependent_cutoff_age, student_dependent_cutoff_age,
    funding_method, pre_funded, pay_as_you_go, pad_requested, setup_fee,
    administration_fee_percentage, administration_fee_flat_amount, administration_fee_basis,
    tax_treatment, currency_code, is_current, version_number, supersedes_plan_id,
    review_status, extraction_status, created_at, created_by, updated_at, updated_by
)
VALUES (
    $client_id, $carrier_id, $administrator_id, $source_document_id, $source_document_version_id,
    NULLIF($policy_number, ''), NULLIF($account_number, ''), NULLIF($plan_name, ''), $plan_type,
    $hcsa_enabled, $lsa_enabled, $wellness_account_enabled, $combined_account_enabled,
    NULLIF($effective_date, ''), NULLIF($termination_date, ''), $benefit_year_type,
    $benefit_year_start_month, $benefit_year_start_day, $default_allocation_type,
    $default_hcsa_percentage, $default_lsa_percentage, $pay_as_submitted, $runoff_days,
    $carry_forward_type, $credit_carry_forward_enabled, $expense_carry_forward_enabled,
    $carry_forward_maximum, $carry_forward_expiry_months, $default_waiting_period_days,
    $new_hire_proration_default, $child_dependent_cutoff_age, $student_dependent_cutoff_age,
    $funding_method, $pre_funded, $pay_as_you_go, $pad_requested, $setup_fee,
    $administration_fee_percentage, $administration_fee_flat_amount, NULLIF($administration_fee_basis, ''),
    NULLIF($tax_treatment, ''), $currency_code, 1, $version_number, $supersedes_plan_id,
    $review_status, $extraction_status, $created_at, $created_by, $updated_at, $updated_by
);");
        AddParameter(command, "$client_id", clientId);
        AddParameter(command, "$carrier_id", carrierId);
        AddParameter(command, "$administrator_id", administratorId);
        AddParameter(command, "$source_document_id", documentId);
        AddParameter(command, "$source_document_version_id", documentVersionId);
        AddParameter(command, "$policy_number", Clean(plan.PolicyNumber));
        AddParameter(command, "$account_number", Clean(plan.AccountNumber));
        AddParameter(command, "$plan_name", Clean(plan.PlanName));
        AddParameter(command, "$plan_type", plan.PlanType.ToString());
        AddParameter(command, "$hcsa_enabled", DbBool(plan.HcsaEnabled));
        AddParameter(command, "$lsa_enabled", DbBool(plan.LsaEnabled));
        AddParameter(command, "$wellness_account_enabled", DbBool(plan.WellnessAccountEnabled));
        AddParameter(command, "$combined_account_enabled", DbBool(plan.CombinedAccountEnabled));
        AddParameter(command, "$effective_date", ToDateText(plan.EffectiveDate));
        AddParameter(command, "$termination_date", ToDateText(plan.TerminationDate));
        AddParameter(command, "$benefit_year_type", plan.BenefitYearType.ToString());
        AddParameter(command, "$benefit_year_start_month", plan.BenefitYearStartMonth);
        AddParameter(command, "$benefit_year_start_day", plan.BenefitYearStartDay);
        AddParameter(command, "$default_allocation_type", plan.DefaultAllocationType.ToString());
        AddParameter(command, "$default_hcsa_percentage", plan.DefaultHcsaPercentage);
        AddParameter(command, "$default_lsa_percentage", plan.DefaultLsaPercentage);
        AddParameter(command, "$pay_as_submitted", DbBool(plan.PayAsSubmitted));
        AddParameter(command, "$runoff_days", plan.RunoffDays);
        AddParameter(command, "$carry_forward_type", plan.CarryForwardType.ToString());
        AddParameter(command, "$credit_carry_forward_enabled", DbBool(plan.CreditCarryForwardEnabled));
        AddParameter(command, "$expense_carry_forward_enabled", DbBool(plan.ExpenseCarryForwardEnabled));
        AddParameter(command, "$carry_forward_maximum", plan.CarryForwardMaximum);
        AddParameter(command, "$carry_forward_expiry_months", plan.CarryForwardExpiryMonths);
        AddParameter(command, "$default_waiting_period_days", plan.DefaultWaitingPeriodDays);
        AddParameter(command, "$new_hire_proration_default", DbBool(plan.NewHireProrationDefault));
        AddParameter(command, "$child_dependent_cutoff_age", plan.ChildDependentCutoffAge);
        AddParameter(command, "$student_dependent_cutoff_age", plan.StudentDependentCutoffAge);
        AddParameter(command, "$funding_method", plan.FundingMethod.ToString());
        AddParameter(command, "$pre_funded", DbBool(plan.PreFunded));
        AddParameter(command, "$pay_as_you_go", DbBool(plan.PayAsYouGo));
        AddParameter(command, "$pad_requested", DbBool(plan.PadRequested));
        AddParameter(command, "$setup_fee", plan.SetupFee);
        AddParameter(command, "$administration_fee_percentage", plan.AdministrationFeePercentage);
        AddParameter(command, "$administration_fee_flat_amount", plan.AdministrationFeeFlatAmount);
        AddParameter(command, "$administration_fee_basis", Clean(plan.AdministrationFeeBasis));
        AddParameter(command, "$tax_treatment", Clean(plan.TaxTreatment));
        AddParameter(command, "$currency_code", FirstNonBlank(plan.CurrencyCode, "CAD"));
        AddParameter(command, "$version_number", versionNumber);
        AddParameter(command, "$supersedes_plan_id", supersedesPlanId);
        AddParameter(command, "$review_status", ToFlexibleReviewStatus(finalStatus));
        AddParameter(command, "$extraction_status", finalStatus);
        AddParameter(command, "$created_at", ToDateTimeText(DateTime.Now));
        AddParameter(command, "$created_by", Clean(request.ImportedBy));
        AddParameter(command, "$updated_at", ToDateTimeText(DateTime.Now));
        AddParameter(command, "$updated_by", Clean(request.ImportedBy));
        await command.ExecuteNonQueryAsync();
        return await LastInsertRowIdAsync(connection, transaction);
    }

    private static async Task<long> LastInsertRowIdAsync(SqliteConnection connection, SqliteTransaction transaction)
    {
        using var command = CreateCommand(connection, transaction, "SELECT last_insert_rowid();");
        return Convert.ToInt64(await command.ExecuteScalarAsync(), CultureInfo.InvariantCulture);
    }

    private static object? DbBool(bool? value)
    {
        return value.HasValue ? value.Value ? 1 : 0 : null;
    }

    private static string ToFlexibleReviewStatus(string importStatus)
    {
        return importStatus == ImportStatuses.Completed
            ? FlexibleBenefitReviewStatuses.Confirmed
            : FlexibleBenefitReviewStatuses.NeedsReview;
    }
}

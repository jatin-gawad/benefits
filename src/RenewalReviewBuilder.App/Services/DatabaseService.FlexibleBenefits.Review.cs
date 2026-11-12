using System.Globalization;
using Microsoft.Data.Sqlite;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed partial class DatabaseService
{
    public async Task<DatabaseOperationResult> MarkFlexibleBenefitPlanReviewedAsync(long planId, string reviewedBy)
    {
        if (planId <= 0)
        {
            return DatabaseOperationResult.Failure("Select a saved flexible-benefits plan before marking it reviewed.", DatabasePath);
        }

        try
        {
            await InitializeCoreAsync();
            using var connection = await OpenConnectionAsync();
            using var transaction = connection.BeginTransaction();
            var identity = await GetFlexibleBenefitReviewIdentityAsync(connection, transaction, planId);
            if (identity is null)
            {
                return DatabaseOperationResult.Failure("The flexible-benefits plan could not be found.", DatabasePath);
            }

            var reviewer = FirstNonBlank(reviewedBy, Environment.UserName, "Desktop user");
            await ExecuteNonQueryAsync(connection, transaction, @"
UPDATE FlexibleBenefitPlan
SET review_status = $review_status,
    updated_at = $updated_at,
    updated_by = $updated_by
WHERE flexible_benefit_plan_id = $plan_id;", command =>
            {
                AddParameter(command, "$review_status", FlexibleBenefitReviewStatuses.Reviewed);
                AddParameter(command, "$updated_at", ToDateTimeText(DateTime.Now));
                AddParameter(command, "$updated_by", reviewer);
                AddParameter(command, "$plan_id", planId);
            });

            await ExecuteNonQueryAsync(connection, transaction, @"
UPDATE FlexibleBenefitDocumentVersion
SET review_status = $review_status
WHERE flexible_benefit_document_version_id = $document_version_id;", command =>
            {
                AddParameter(command, "$review_status", FlexibleBenefitReviewStatuses.Reviewed);
                AddParameter(command, "$document_version_id", identity.DocumentVersionId);
            });

            foreach (var table in FlexibleBenefitReviewChildTables)
            {
                await ExecuteNonQueryAsync(
                    connection,
                    transaction,
                    $"UPDATE {table} SET review_status = $review_status WHERE flexible_benefit_plan_id = $plan_id;",
                    command =>
                    {
                        AddParameter(command, "$review_status", FlexibleBenefitReviewStatuses.Reviewed);
                        AddParameter(command, "$plan_id", planId);
                    });
            }

            await ExecuteNonQueryAsync(connection, transaction, @"
UPDATE FlexibleBenefitPlanSourceField
SET review_status = $review_status
WHERE source_document_version_id = $document_version_id;", command =>
            {
                AddParameter(command, "$review_status", FlexibleBenefitReviewStatuses.Reviewed);
                AddParameter(command, "$document_version_id", identity.DocumentVersionId);
            });

            transaction.Commit();
            return DatabaseOperationResult.Success("Flexible-benefits plan and its extracted source fields were marked reviewed.", DatabasePath);
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            return DatabaseOperationResult.Failure("The flexible-benefits review status could not be saved.", DatabasePath);
        }
    }

    public async Task<DatabaseOperationResult> CorrectFlexibleBenefitSourceFieldAsync(
        long sourceFieldId,
        string newNormalizedValue,
        string correctedBy,
        string reason)
    {
        if (sourceFieldId <= 0)
        {
            return DatabaseOperationResult.Failure("Select a source field before editing it.", DatabasePath);
        }

        try
        {
            await InitializeCoreAsync();
            using var connection = await OpenConnectionAsync();
            using var transaction = connection.BeginTransaction();
            var sourceField = await ReadFlexibleBenefitCorrectionSourceAsync(connection, transaction, sourceFieldId);
            if (sourceField is null || !sourceField.EntityId.HasValue)
            {
                return DatabaseOperationResult.Failure("The selected source field is not linked to a normalized record.", DatabasePath);
            }

            var target = ResolveFlexibleBenefitCorrectionTarget(sourceField.EntityType, sourceField.FieldName);
            if (target is null)
            {
                return DatabaseOperationResult.Failure("This field is audit-only and cannot be edited from the review screen.", DatabasePath);
            }

            if (!TryNormalizeFlexibleBenefitCorrection(
                    target.ValueKind,
                    sourceField.FieldName,
                    newNormalizedValue,
                    out var databaseValue,
                    out var normalizedText,
                    out var validationMessage))
            {
                return DatabaseOperationResult.Failure(validationMessage, DatabasePath);
            }

            var user = FirstNonBlank(correctedBy, Environment.UserName, "Desktop user");
            var correctionReason = FirstNonBlank(reason, "Confirmed in flexible-benefits extraction review.");
            var correctedAt = DateTime.Now;
            await ExecuteNonQueryAsync(
                connection,
                transaction,
                $"UPDATE {target.TableName} SET {target.ColumnName} = $value WHERE {target.IdColumnName} = $entity_id;",
                command =>
                {
                    AddParameter(command, "$value", databaseValue);
                    AddParameter(command, "$entity_id", sourceField.EntityId.Value);
                });

            await ExecuteNonQueryAsync(connection, transaction, @"
INSERT INTO FlexibleBenefitCorrectionAudit (
    flexible_benefit_plan_source_field_id, source_document_id, source_document_version_id,
    entity_type, entity_id, field_name, previous_normalized_value, new_normalized_value,
    reason, corrected_at, corrected_by
) VALUES (
    $source_field_id, $source_document_id, $source_document_version_id,
    $entity_type, $entity_id, $field_name, $previous_value, $new_value,
    $reason, $corrected_at, $corrected_by
);", command =>
            {
                AddParameter(command, "$source_field_id", sourceField.SourceFieldId);
                AddParameter(command, "$source_document_id", sourceField.SourceDocumentId);
                AddParameter(command, "$source_document_version_id", sourceField.SourceDocumentVersionId);
                AddParameter(command, "$entity_type", sourceField.EntityType);
                AddParameter(command, "$entity_id", sourceField.EntityId);
                AddParameter(command, "$field_name", sourceField.FieldName);
                AddParameter(command, "$previous_value", sourceField.NormalizedValue);
                AddParameter(command, "$new_value", normalizedText);
                AddParameter(command, "$reason", correctionReason);
                AddParameter(command, "$corrected_at", ToDateTimeText(correctedAt));
                AddParameter(command, "$corrected_by", user);
            });

            await ExecuteNonQueryAsync(connection, transaction, @"
UPDATE FlexibleBenefitPlanSourceField
SET normalized_value = $normalized_value,
    was_manually_corrected = 1,
    corrected_by = $corrected_by,
    corrected_at = $corrected_at,
    review_status = $review_status
WHERE flexible_benefit_plan_source_field_id = $source_field_id;", command =>
            {
                AddParameter(command, "$normalized_value", normalizedText);
                AddParameter(command, "$corrected_by", user);
                AddParameter(command, "$corrected_at", ToDateTimeText(correctedAt));
                AddParameter(command, "$review_status", FlexibleBenefitReviewStatuses.Reviewed);
                AddParameter(command, "$source_field_id", sourceField.SourceFieldId);
            });

            transaction.Commit();
            return DatabaseOperationResult.Success("The normalized value was corrected and the original PDF source value was preserved.", DatabasePath);
        }
        catch (Exception ex)
        {
            AppLogger.Error(ex);
            return DatabaseOperationResult.Failure("The source-field correction could not be saved.", DatabasePath);
        }
    }

    private static async Task<FlexibleBenefitReviewIdentity?> GetFlexibleBenefitReviewIdentityAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long planId)
    {
        using var command = CreateCommand(connection, transaction, @"
SELECT source_document_id, source_document_version_id
FROM FlexibleBenefitPlan
WHERE flexible_benefit_plan_id = $plan_id
LIMIT 1;");
        AddParameter(command, "$plan_id", planId);
        using var reader = await command.ExecuteReaderAsync();
        return await reader.ReadAsync()
            ? new FlexibleBenefitReviewIdentity(reader.GetInt64(0), reader.GetInt64(1))
            : null;
    }

    private static async Task<FlexibleBenefitCorrectionSource?> ReadFlexibleBenefitCorrectionSourceAsync(
        SqliteConnection connection,
        SqliteTransaction transaction,
        long sourceFieldId)
    {
        using var command = CreateCommand(connection, transaction, @"
SELECT flexible_benefit_plan_source_field_id, source_document_id, source_document_version_id,
       entity_type, entity_id, field_name, normalized_value
FROM FlexibleBenefitPlanSourceField
WHERE flexible_benefit_plan_source_field_id = $source_field_id
LIMIT 1;");
        AddParameter(command, "$source_field_id", sourceFieldId);
        using var reader = await command.ExecuteReaderAsync();
        if (!await reader.ReadAsync())
        {
            return null;
        }

        return new FlexibleBenefitCorrectionSource(
            reader.GetInt64(0),
            reader.GetInt64(1),
            reader.GetInt64(2),
            reader.GetString(3),
            reader.IsDBNull(4) ? null : reader.GetInt64(4),
            reader.GetString(5),
            reader.IsDBNull(6) ? "" : reader.GetString(6));
    }

    private static FlexibleBenefitCorrectionTarget? ResolveFlexibleBenefitCorrectionTarget(string entityType, string fieldName)
    {
        return FlexibleBenefitCorrectionTargets.TryGetValue($"{entityType}.{fieldName}", out var target)
            ? target
            : null;
    }

    private static bool TryNormalizeFlexibleBenefitCorrection(
        string valueKind,
        string fieldName,
        string input,
        out object? databaseValue,
        out string normalizedText,
        out string message)
    {
        var value = input?.Trim() ?? "";
        databaseValue = null;
        normalizedText = value;
        message = "";
        switch (valueKind)
        {
            case "Text":
                databaseValue = value;
                return true;
            case "Integer":
                if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var integer) && integer >= 0)
                {
                    databaseValue = integer;
                    normalizedText = integer.ToString(CultureInfo.InvariantCulture);
                    return true;
                }

                message = "Enter a valid non-negative whole number.";
                return false;
            case "Decimal":
                var cleanedDecimal = value.Replace("$", "", StringComparison.Ordinal).Replace("%", "", StringComparison.Ordinal).Replace(",", "", StringComparison.Ordinal);
                if (decimal.TryParse(cleanedDecimal, NumberStyles.Number, CultureInfo.InvariantCulture, out var number) && number >= 0)
                {
                    databaseValue = number;
                    normalizedText = number.ToString("0.####", CultureInfo.InvariantCulture);
                    return true;
                }

                message = "Enter a valid non-negative amount or percentage.";
                return false;
            case "Boolean":
                if (TryParseFlexibleBenefitBoolean(value, out var boolean))
                {
                    databaseValue = boolean ? 1 : 0;
                    normalizedText = boolean ? "true" : "false";
                    return true;
                }

                message = "Enter Yes or No (or true or false).";
                return false;
            case "Date":
                if (DateOnly.TryParse(value, CultureInfo.GetCultureInfo("en-CA"), DateTimeStyles.AllowWhiteSpaces, out var date))
                {
                    databaseValue = date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                    normalizedText = (string)databaseValue;
                    return true;
                }

                message = "Enter a valid date, for example 2023-04-01.";
                return false;
            case "DateTime":
                if (DateTime.TryParse(value, CultureInfo.GetCultureInfo("en-CA"), DateTimeStyles.AllowWhiteSpaces, out var dateTime))
                {
                    databaseValue = ToDateTimeText(dateTime);
                    normalizedText = (string)databaseValue;
                    return true;
                }

                message = "Enter a valid date and time.";
                return false;
            case "Enum":
                if (IsValidFlexibleBenefitEnumValue(fieldName, value))
                {
                    databaseValue = value;
                    return true;
                }

                message = $"'{value}' is not a supported value for {fieldName}.";
                return false;
            default:
                message = "This value type is not supported by the review screen.";
                return false;
        }
    }

    private static bool TryParseFlexibleBenefitBoolean(string value, out bool result)
    {
        if (bool.TryParse(value, out result))
        {
            return true;
        }

        if (value.Equals("yes", StringComparison.OrdinalIgnoreCase) || value.Equals("selected", StringComparison.OrdinalIgnoreCase))
        {
            result = true;
            return true;
        }

        if (value.Equals("no", StringComparison.OrdinalIgnoreCase) || value.Equals("unselected", StringComparison.OrdinalIgnoreCase))
        {
            result = false;
            return true;
        }

        result = false;
        return false;
    }

    private static bool IsValidFlexibleBenefitEnumValue(string fieldName, string value)
    {
        return fieldName switch
        {
            "PlanType" => Enum.TryParse<FlexibleBenefitPlanType>(value, true, out _),
            "BenefitYearType" => Enum.TryParse<FlexibleBenefitYearType>(value, true, out _),
            "DefaultAllocationType" => Enum.TryParse<FlexibleBenefitAllocationType>(value, true, out _),
            "CarryForwardType" => Enum.TryParse<FlexibleBenefitCarryForwardType>(value, true, out _),
            "FundingMethod" => Enum.TryParse<FlexibleBenefitFundingMethod>(value, true, out _),
            _ => false
        };
    }

    private static readonly string[] FlexibleBenefitReviewChildTables =
    [
        "FlexibleBenefitClass",
        "FlexibleBenefitCategory",
        "FlexibleBenefitEligibilityRule",
        "FlexibleBenefitFundingConfiguration",
        "FlexibleBenefitFee",
        "FlexibleBenefitAgreement",
        "FlexibleBenefitContact"
    ];

    private static readonly IReadOnlyDictionary<string, FlexibleBenefitCorrectionTarget> FlexibleBenefitCorrectionTargets =
        new Dictionary<string, FlexibleBenefitCorrectionTarget>(StringComparer.OrdinalIgnoreCase)
        {
            ["Client.LegalName"] = new("Clients", "client_id", "legal_name", "Text"),
            ["Client.DbaName"] = new("Clients", "client_id", "dba_name", "Text"),
            ["Client.Address"] = new("Clients", "client_id", "address", "Text"),
            ["Client.City"] = new("Clients", "client_id", "city", "Text"),
            ["Client.Province"] = new("Clients", "client_id", "province", "Text"),
            ["Client.PostalCode"] = new("Clients", "client_id", "postal_code", "Text"),
            ["Client.Phone"] = new("Clients", "client_id", "phone", "Text"),
            ["FlexibleBenefitPlan.PlanName"] = new("FlexibleBenefitPlan", "flexible_benefit_plan_id", "plan_name", "Text"),
            ["FlexibleBenefitPlan.PlanType"] = new("FlexibleBenefitPlan", "flexible_benefit_plan_id", "plan_type", "Enum"),
            ["FlexibleBenefitPlan.HcsaEnabled"] = new("FlexibleBenefitPlan", "flexible_benefit_plan_id", "hcsa_enabled", "Boolean"),
            ["FlexibleBenefitPlan.LsaEnabled"] = new("FlexibleBenefitPlan", "flexible_benefit_plan_id", "lsa_enabled", "Boolean"),
            ["FlexibleBenefitPlan.DefaultAllocationType"] = new("FlexibleBenefitPlan", "flexible_benefit_plan_id", "default_allocation_type", "Enum"),
            ["FlexibleBenefitPlan.DefaultHcsaPercentage"] = new("FlexibleBenefitPlan", "flexible_benefit_plan_id", "default_hcsa_percentage", "Decimal"),
            ["FlexibleBenefitPlan.DefaultLsaPercentage"] = new("FlexibleBenefitPlan", "flexible_benefit_plan_id", "default_lsa_percentage", "Decimal"),
            ["FlexibleBenefitPlan.EffectiveDate"] = new("FlexibleBenefitPlan", "flexible_benefit_plan_id", "effective_date", "Date"),
            ["FlexibleBenefitPlan.BenefitYearType"] = new("FlexibleBenefitPlan", "flexible_benefit_plan_id", "benefit_year_type", "Enum"),
            ["FlexibleBenefitPlan.RunoffDays"] = new("FlexibleBenefitPlan", "flexible_benefit_plan_id", "runoff_days", "Integer"),
            ["FlexibleBenefitPlan.CarryForwardType"] = new("FlexibleBenefitPlan", "flexible_benefit_plan_id", "carry_forward_type", "Enum"),
            ["FlexibleBenefitPlan.ChildDependentCutoffAge"] = new("FlexibleBenefitPlan", "flexible_benefit_plan_id", "child_dependent_cutoff_age", "Integer"),
            ["FlexibleBenefitPlan.StudentDependentCutoffAge"] = new("FlexibleBenefitPlan", "flexible_benefit_plan_id", "student_dependent_cutoff_age", "Integer"),
            ["FlexibleBenefitPlan.PayAsSubmitted"] = new("FlexibleBenefitPlan", "flexible_benefit_plan_id", "pay_as_submitted", "Boolean"),
            ["FlexibleBenefitClass.ClassName"] = new("FlexibleBenefitClass", "flexible_benefit_class_id", "class_name", "Text"),
            ["FlexibleBenefitClass.AnnualSingleAmount"] = new("FlexibleBenefitClass", "flexible_benefit_class_id", "annual_single_amount", "Decimal"),
            ["FlexibleBenefitClass.AnnualCoupleAmount"] = new("FlexibleBenefitClass", "flexible_benefit_class_id", "annual_couple_amount", "Decimal"),
            ["FlexibleBenefitClass.AnnualFamilyAmount"] = new("FlexibleBenefitClass", "flexible_benefit_class_id", "annual_family_amount", "Decimal"),
            ["FlexibleBenefitClass.WaitingPeriodDays"] = new("FlexibleBenefitClass", "flexible_benefit_class_id", "waiting_period_days", "Integer"),
            ["FlexibleBenefitClass.NewHireProrationEnabled"] = new("FlexibleBenefitClass", "flexible_benefit_class_id", "new_hire_proration_enabled", "Boolean"),
            ["FlexibleBenefitCategory.Enabled"] = new("FlexibleBenefitCategory", "flexible_benefit_category_id", "enabled", "Boolean"),
            ["FlexibleBenefitEligibilityRule.RuleValue"] = new("FlexibleBenefitEligibilityRule", "flexible_benefit_eligibility_rule_id", "rule_value", "Text"),
            ["FlexibleBenefitFundingConfiguration.PadSelected"] = new("FlexibleBenefitFundingConfiguration", "flexible_benefit_funding_configuration_id", "pad_selected", "Boolean"),
            ["FlexibleBenefitFundingConfiguration.OnlineBankPaymentAllowed"] = new("FlexibleBenefitFundingConfiguration", "flexible_benefit_funding_configuration_id", "online_bank_payment_allowed", "Boolean"),
            ["FlexibleBenefitFundingConfiguration.ChequeAllowed"] = new("FlexibleBenefitFundingConfiguration", "flexible_benefit_funding_configuration_id", "cheque_allowed", "Boolean"),
            ["FlexibleBenefitFundingConfiguration.ClaimsPaidOnlyAfterFundsClear"] = new("FlexibleBenefitFundingConfiguration", "flexible_benefit_funding_configuration_id", "claims_paid_only_after_funds_clear", "Boolean"),
            ["FlexibleBenefitFundingConfiguration.InterestPaidOnFunds"] = new("FlexibleBenefitFundingConfiguration", "flexible_benefit_funding_configuration_id", "interest_paid_on_funds", "Boolean"),
            ["FlexibleBenefitFee.FlatAmount"] = new("FlexibleBenefitFee", "flexible_benefit_fee_id", "flat_amount", "Decimal"),
            ["FlexibleBenefitFee.Percentage"] = new("FlexibleBenefitFee", "flexible_benefit_fee_id", "percentage", "Decimal"),
            ["FlexibleBenefitFee.TaxType"] = new("FlexibleBenefitFee", "flexible_benefit_fee_id", "tax_type", "Text"),
            ["FlexibleBenefitFee.NoticePeriodDays"] = new("FlexibleBenefitFee", "flexible_benefit_fee_id", "notice_period_days", "Integer"),
            ["FlexibleBenefitAgreement.AgreementDate"] = new("FlexibleBenefitAgreement", "flexible_benefit_agreement_id", "agreement_date", "Date"),
            ["FlexibleBenefitAgreement.ExecutionDate"] = new("FlexibleBenefitAgreement", "flexible_benefit_agreement_id", "execution_date", "DateTime"),
            ["FlexibleBenefitAgreement.GoverningProvince"] = new("FlexibleBenefitAgreement", "flexible_benefit_agreement_id", "governing_province", "Text"),
            ["FlexibleBenefitAgreement.SubscriberSignerName"] = new("FlexibleBenefitAgreement", "flexible_benefit_agreement_id", "subscriber_signer_name", "Text"),
            ["FlexibleBenefitAgreement.AdministratorSignerName"] = new("FlexibleBenefitAgreement", "flexible_benefit_agreement_id", "administrator_signer_name", "Text"),
            ["FlexibleBenefitAgreement.SubscriberSignaturePresent"] = new("FlexibleBenefitAgreement", "flexible_benefit_agreement_id", "subscriber_signature_present", "Boolean"),
            ["FlexibleBenefitAgreement.AdministratorSignaturePresent"] = new("FlexibleBenefitAgreement", "flexible_benefit_agreement_id", "administrator_signature_present", "Boolean"),
            ["FlexibleBenefitContact.ContactName"] = new("FlexibleBenefitContact", "flexible_benefit_contact_id", "contact_name", "Text"),
            ["FlexibleBenefitContact.Email"] = new("FlexibleBenefitContact", "flexible_benefit_contact_id", "email", "Text")
        };

    private sealed record FlexibleBenefitReviewIdentity(long DocumentId, long DocumentVersionId);
    private sealed record FlexibleBenefitCorrectionSource(
        long SourceFieldId,
        long SourceDocumentId,
        long SourceDocumentVersionId,
        string EntityType,
        long? EntityId,
        string FieldName,
        string NormalizedValue);
    private sealed record FlexibleBenefitCorrectionTarget(
        string TableName,
        string IdColumnName,
        string ColumnName,
        string ValueKind);
}

using System.Globalization;
using Microsoft.Data.Sqlite;
using RenewalReviewBuilder.App.Models;
using RenewalReviewBuilder.App.Services.FlexibleBenefits;

namespace RenewalReviewBuilder.App.Services;

public sealed partial class DatabaseService
{
    public async Task<FlexibleBenefitWorkspaceViewModel> GetFlexibleBenefitWorkspaceAsync(long clientId)
    {
        if (clientId <= 0)
        {
            return new FlexibleBenefitWorkspaceViewModel();
        }

        await InitializeCoreAsync();
        using var connection = await OpenConnectionAsync();
        var planRecord = await ReadCurrentFlexibleBenefitPlanAsync(connection, clientId);
        if (planRecord is null)
        {
            return new FlexibleBenefitWorkspaceViewModel();
        }

        var classes = await ReadFlexibleBenefitClassesAsync(connection, planRecord.Plan.FlexibleBenefitPlanId!.Value);
        var summaryService = new FlexibleBenefitSummaryService();
        var model = new FlexibleBenefitWorkspaceViewModel
        {
            HasPlan = true,
            PlanId = planRecord.Plan.FlexibleBenefitPlanId,
            SourceDocumentId = planRecord.Plan.SourceDocumentId,
            SourceDocumentVersionId = planRecord.Plan.SourceDocumentVersionId,
            Administrator = FirstNonBlank(planRecord.Administrator, "Not confirmed"),
            PlanType = HumanizeFlexibleBenefitValue(planRecord.Plan.PlanType.ToString()),
            EffectiveDate = FormatFlexibleBenefitDate(planRecord.Plan.EffectiveDate),
            HcsaSummary = summaryService.BuildHcsaSummary(planRecord.Plan, classes),
            LsaSummary = summaryService.BuildLsaSummary(planRecord.Plan, classes),
            DefaultAllocation = FormatFlexibleBenefitAllocation(planRecord.Plan),
            BenefitYear = HumanizeFlexibleBenefitValue(planRecord.Plan.BenefitYearType.ToString()),
            Runoff = planRecord.Plan.RunoffDays.HasValue ? $"{planRecord.Plan.RunoffDays.Value:#,0} days" : "Not confirmed",
            CarryForward = HumanizeFlexibleBenefitValue(planRecord.Plan.CarryForwardType.ToString()),
            DependentAges = FormatFlexibleBenefitDependentAges(planRecord.Plan),
            FundingSummary = FormatFlexibleBenefitFunding(planRecord.Plan),
            FeeSummary = FormatFlexibleBenefitFees(planRecord.Plan),
            AgreementSummary = "Not confirmed",
            ReviewStatus = FirstNonBlank(planRecord.Plan.ReviewStatus, "Not confirmed"),
            SourceDocumentPath = planRecord.SourceDocumentPath,
            VersionNumber = planRecord.Plan.VersionNumber
        };

        foreach (var item in classes)
        {
            model.Classes.Add(new FlexibleBenefitClassDisplayRow
            {
                ClassName = FirstNonBlank(item.ClassName, "Unnamed class"),
                SingleAmount = FormatFlexibleBenefitMoney(item.AnnualSingleAmount),
                CoupleAmount = FormatFlexibleBenefitMoney(item.AnnualCoupleAmount),
                FamilyAmount = FormatFlexibleBenefitMoney(item.AnnualFamilyAmount),
                WaitingPeriod = item.WaitingPeriodDays.HasValue ? $"{item.WaitingPeriodDays.Value:#,0} days" : "Not confirmed",
                Proration = FormatFlexibleBenefitBoolean(item.NewHireProrationEnabled),
                SourcePage = FormatFlexibleBenefitSourcePage(item.SourcePageNumber)
            });
        }

        await ReadFlexibleBenefitCategoriesAsync(connection, planRecord.Plan.FlexibleBenefitPlanId.Value, model);
        await ReadFlexibleBenefitFundingAsync(connection, planRecord.Plan.FlexibleBenefitPlanId.Value, model);
        await ReadFlexibleBenefitFeesAsync(connection, planRecord.Plan.FlexibleBenefitPlanId.Value, model);
        model.AgreementSummary = await ReadFlexibleBenefitAgreementsAsync(connection, planRecord.Plan.FlexibleBenefitPlanId.Value, model);
        await ReadFlexibleBenefitContactsAsync(connection, planRecord.Plan.FlexibleBenefitPlanId.Value, model);
        await ReadFlexibleBenefitSourceFieldsAsync(connection, planRecord.Plan.SourceDocumentVersionId!.Value, model);
        await ReadFlexibleBenefitVersionsAsync(connection, clientId, model);
        await ReadFlexibleBenefitReviewItemsAsync(connection, planRecord.Plan.SourceDocumentId!.Value, model);
        return model;
    }

    private static async Task<FlexibleBenefitPlanRecord?> ReadCurrentFlexibleBenefitPlanAsync(
        SqliteConnection connection,
        long clientId)
    {
        using var command = CreateCommand(connection, null, @"
SELECT fp.*,
       COALESCE(NULLIF(a.carrier_name, ''), NULLIF(d.carrier, ''), '') AS administrator_name,
       d.file_path AS source_document_path
FROM FlexibleBenefitPlan fp
INNER JOIN Documents d ON d.document_id = fp.source_document_id
LEFT JOIN Carrier a ON a.carrier_id = fp.administrator_id
WHERE fp.client_id = $client_id
  AND fp.is_current = 1
ORDER BY fp.version_number DESC
LIMIT 1;");
        AddParameter(command, "$client_id", clientId);
        using var reader = await command.ExecuteReaderAsync();
        if (!await reader.ReadAsync())
        {
            return null;
        }

        return new FlexibleBenefitPlanRecord(
            new FlexibleBenefitPlan
            {
                FlexibleBenefitPlanId = ReadLong(reader, "flexible_benefit_plan_id"),
                ClientId = ReadLong(reader, "client_id"),
                CarrierId = ReadNullableLong(reader, "carrier_id"),
                AdministratorId = ReadNullableLong(reader, "administrator_id"),
                SourceDocumentId = ReadNullableLong(reader, "source_document_id"),
                SourceDocumentVersionId = ReadNullableLong(reader, "source_document_version_id"),
                PolicyNumber = ReadString(reader, "policy_number"),
                AccountNumber = ReadString(reader, "account_number"),
                PlanName = ReadString(reader, "plan_name"),
                PlanType = ParseFlexibleBenefitEnum(ReadString(reader, "plan_type"), FlexibleBenefitPlanType.Other),
                HcsaEnabled = ReadNullableFlexibleBenefitBoolean(reader, "hcsa_enabled"),
                LsaEnabled = ReadNullableFlexibleBenefitBoolean(reader, "lsa_enabled"),
                WellnessAccountEnabled = ReadNullableFlexibleBenefitBoolean(reader, "wellness_account_enabled"),
                CombinedAccountEnabled = ReadNullableFlexibleBenefitBoolean(reader, "combined_account_enabled"),
                EffectiveDate = ReadDateOnly(reader, "effective_date"),
                TerminationDate = ReadDateOnly(reader, "termination_date"),
                BenefitYearType = ParseFlexibleBenefitEnum(ReadString(reader, "benefit_year_type"), FlexibleBenefitYearType.Unknown),
                BenefitYearStartMonth = ReadInt(reader, "benefit_year_start_month"),
                BenefitYearStartDay = ReadInt(reader, "benefit_year_start_day"),
                DefaultAllocationType = ParseFlexibleBenefitEnum(ReadString(reader, "default_allocation_type"), FlexibleBenefitAllocationType.Unknown),
                DefaultHcsaPercentage = ReadDecimal(reader, "default_hcsa_percentage"),
                DefaultLsaPercentage = ReadDecimal(reader, "default_lsa_percentage"),
                PayAsSubmitted = ReadNullableFlexibleBenefitBoolean(reader, "pay_as_submitted"),
                RunoffDays = ReadInt(reader, "runoff_days"),
                CarryForwardType = ParseFlexibleBenefitEnum(ReadString(reader, "carry_forward_type"), FlexibleBenefitCarryForwardType.Unknown),
                CreditCarryForwardEnabled = ReadNullableFlexibleBenefitBoolean(reader, "credit_carry_forward_enabled"),
                ExpenseCarryForwardEnabled = ReadNullableFlexibleBenefitBoolean(reader, "expense_carry_forward_enabled"),
                CarryForwardMaximum = ReadDecimal(reader, "carry_forward_maximum"),
                CarryForwardExpiryMonths = ReadInt(reader, "carry_forward_expiry_months"),
                DefaultWaitingPeriodDays = ReadInt(reader, "default_waiting_period_days"),
                NewHireProrationDefault = ReadNullableFlexibleBenefitBoolean(reader, "new_hire_proration_default"),
                ChildDependentCutoffAge = ReadInt(reader, "child_dependent_cutoff_age"),
                StudentDependentCutoffAge = ReadInt(reader, "student_dependent_cutoff_age"),
                FundingMethod = ParseFlexibleBenefitEnum(ReadString(reader, "funding_method"), FlexibleBenefitFundingMethod.Unknown),
                PreFunded = ReadNullableFlexibleBenefitBoolean(reader, "pre_funded"),
                PayAsYouGo = ReadNullableFlexibleBenefitBoolean(reader, "pay_as_you_go"),
                PadRequested = ReadNullableFlexibleBenefitBoolean(reader, "pad_requested"),
                SetupFee = ReadDecimal(reader, "setup_fee"),
                AdministrationFeePercentage = ReadDecimal(reader, "administration_fee_percentage"),
                AdministrationFeeFlatAmount = ReadDecimal(reader, "administration_fee_flat_amount"),
                AdministrationFeeBasis = ReadString(reader, "administration_fee_basis"),
                TaxTreatment = ReadString(reader, "tax_treatment"),
                CurrencyCode = ReadString(reader, "currency_code"),
                IsCurrent = ReadNullableFlexibleBenefitBoolean(reader, "is_current") == true,
                VersionNumber = ReadInt(reader, "version_number") ?? 1,
                SupersedesPlanId = ReadNullableLong(reader, "supersedes_plan_id"),
                ReviewStatus = ReadString(reader, "review_status"),
                ExtractionStatus = ReadString(reader, "extraction_status")
            },
            ReadString(reader, "administrator_name"),
            ReadString(reader, "source_document_path"));
    }

    private static async Task<List<FlexibleBenefitClass>> ReadFlexibleBenefitClassesAsync(
        SqliteConnection connection,
        long planId)
    {
        using var command = CreateCommand(connection, null, @"
SELECT *
FROM FlexibleBenefitClass
WHERE flexible_benefit_plan_id = $plan_id
ORDER BY class_name COLLATE NOCASE, flexible_benefit_class_id;");
        AddParameter(command, "$plan_id", planId);
        var rows = new List<FlexibleBenefitClass>();
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            rows.Add(new FlexibleBenefitClass
            {
                FlexibleBenefitClassId = ReadLong(reader, "flexible_benefit_class_id"),
                FlexibleBenefitPlanId = planId,
                ExternalClassCode = ReadString(reader, "external_class_code"),
                ClassName = ReadString(reader, "class_name"),
                JobDescription = ReadString(reader, "job_description"),
                EligibilityDescription = ReadString(reader, "eligibility_description"),
                AnnualSingleAmount = ReadDecimal(reader, "annual_single_amount"),
                AnnualCoupleAmount = ReadDecimal(reader, "annual_couple_amount"),
                AnnualFamilyAmount = ReadDecimal(reader, "annual_family_amount"),
                AnnualEmployeeAmount = ReadDecimal(reader, "annual_employee_amount"),
                AnnualDependentAmount = ReadDecimal(reader, "annual_dependent_amount"),
                WaitingPeriodDays = ReadInt(reader, "waiting_period_days"),
                NewHireProrationEnabled = ReadNullableFlexibleBenefitBoolean(reader, "new_hire_proration_enabled"),
                HcsaAmount = ReadDecimal(reader, "hcsa_amount"),
                LsaAmount = ReadDecimal(reader, "lsa_amount"),
                HcsaPercentage = ReadDecimal(reader, "hcsa_percentage"),
                LsaPercentage = ReadDecimal(reader, "lsa_percentage"),
                CurrencyCode = ReadString(reader, "currency_code"),
                EffectiveDate = ReadDateOnly(reader, "effective_date"),
                TerminationDate = ReadDateOnly(reader, "termination_date"),
                IsCurrent = ReadNullableFlexibleBenefitBoolean(reader, "is_current") == true,
                SourcePageNumber = ReadInt(reader, "source_page_number"),
                ExtractionConfidence = ReadDecimal(reader, "extraction_confidence"),
                ReviewStatus = ReadString(reader, "review_status")
            });
        }

        return rows;
    }

    private static async Task ReadFlexibleBenefitCategoriesAsync(
        SqliteConnection connection,
        long planId,
        FlexibleBenefitWorkspaceViewModel model)
    {
        using var command = CreateCommand(connection, null, @"
SELECT category_group, category_name, enabled, source_page_number
FROM FlexibleBenefitCategory
WHERE flexible_benefit_plan_id = $plan_id
ORDER BY category_group COLLATE NOCASE, category_name COLLATE NOCASE;");
        AddParameter(command, "$plan_id", planId);
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            model.Categories.Add(new FlexibleBenefitCategoryDisplayRow
            {
                CategoryGroup = ReadString(reader, "category_group"),
                CategoryName = ReadString(reader, "category_name"),
                Enabled = FormatFlexibleBenefitBoolean(ReadNullableFlexibleBenefitBoolean(reader, "enabled")),
                SourcePage = FormatFlexibleBenefitSourcePage(ReadInt(reader, "source_page_number"))
            });
        }
    }

    private static async Task ReadFlexibleBenefitFundingAsync(
        SqliteConnection connection,
        long planId,
        FlexibleBenefitWorkspaceViewModel model)
    {
        using var command = CreateCommand(connection, null, @"
SELECT *
FROM FlexibleBenefitFundingConfiguration
WHERE flexible_benefit_plan_id = $plan_id
ORDER BY flexible_benefit_funding_configuration_id;");
        AddParameter(command, "$plan_id", planId);
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            var sourcePage = FormatFlexibleBenefitSourcePage(ReadInt(reader, "source_page_number"));
            AddFlexibleBenefitDetail(model.Funding, "Funding method", HumanizeFlexibleBenefitValue(ReadString(reader, "funding_method")), sourcePage);
            AddFlexibleBenefitDetail(model.Funding, "PAD selected", FormatFlexibleBenefitBoolean(ReadNullableFlexibleBenefitBoolean(reader, "pad_selected")), sourcePage);
            AddFlexibleBenefitDetail(model.Funding, "Pre-funding required", FormatFlexibleBenefitBoolean(ReadNullableFlexibleBenefitBoolean(reader, "pre_funding_required")), sourcePage);
            AddFlexibleBenefitDetail(model.Funding, "Online bank payment", FormatFlexibleBenefitBoolean(ReadNullableFlexibleBenefitBoolean(reader, "online_bank_payment_allowed")), sourcePage);
            AddFlexibleBenefitDetail(model.Funding, "Cheque allowed", FormatFlexibleBenefitBoolean(ReadNullableFlexibleBenefitBoolean(reader, "cheque_allowed")), sourcePage);
            AddFlexibleBenefitDetail(model.Funding, "Claims paid after funds clear", FormatFlexibleBenefitBoolean(ReadNullableFlexibleBenefitBoolean(reader, "claims_paid_only_after_funds_clear")), sourcePage);
            AddFlexibleBenefitDetail(model.Funding, "Interest paid on funds", FormatFlexibleBenefitBoolean(ReadNullableFlexibleBenefitBoolean(reader, "interest_paid_on_funds")), sourcePage);
        }
    }

    private static async Task ReadFlexibleBenefitFeesAsync(
        SqliteConnection connection,
        long planId,
        FlexibleBenefitWorkspaceViewModel model)
    {
        using var command = CreateCommand(connection, null, @"
SELECT *
FROM FlexibleBenefitFee
WHERE flexible_benefit_plan_id = $plan_id
ORDER BY flexible_benefit_fee_id;");
        AddParameter(command, "$plan_id", planId);
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            var percentage = ReadDecimal(reader, "percentage");
            var amount = ReadDecimal(reader, "flat_amount");
            var value = percentage.HasValue
                ? $"{percentage.Value:0.##}%"
                : amount.HasValue
                    ? FormatFlexibleBenefitMoney(amount)
                    : "Not confirmed";
            var details = new List<string>();
            var taxType = ReadString(reader, "tax_type");
            var noticeDays = ReadInt(reader, "notice_period_days");
            if (!string.IsNullOrWhiteSpace(taxType))
            {
                details.Add(taxType);
            }

            if (noticeDays.HasValue)
            {
                details.Add($"{noticeDays.Value:#,0}-day change notice");
            }

            if (details.Count > 0)
            {
                value = $"{value} ({string.Join(", ", details)})";
            }

            AddFlexibleBenefitDetail(
                model.Fees,
                HumanizeFlexibleBenefitValue(ReadString(reader, "fee_type")),
                value,
                FormatFlexibleBenefitSourcePage(ReadInt(reader, "source_page_number")));
        }
    }

    private static async Task<string> ReadFlexibleBenefitAgreementsAsync(
        SqliteConnection connection,
        long planId,
        FlexibleBenefitWorkspaceViewModel model)
    {
        using var command = CreateCommand(connection, null, @"
SELECT *
FROM FlexibleBenefitAgreement
WHERE flexible_benefit_plan_id = $plan_id
ORDER BY flexible_benefit_agreement_id;");
        AddParameter(command, "$plan_id", planId);
        using var reader = await command.ExecuteReaderAsync();
        string summary = "Not confirmed";
        while (await reader.ReadAsync())
        {
            var page = FormatFlexibleBenefitSourcePage(ReadInt(reader, "source_page_number"));
            var subscriberSigned = ReadNullableFlexibleBenefitBoolean(reader, "subscriber_signature_present");
            var administratorSigned = ReadNullableFlexibleBenefitBoolean(reader, "administrator_signature_present");
            var status = subscriberSigned == true && administratorSigned == true ? "Signed" : "Needs review";
            var execution = ReadDateTime(reader, "execution_date");
            summary = execution.HasValue ? $"{status} {execution.Value:MMM d, yyyy}" : status;
            AddFlexibleBenefitDetail(model.Agreements, "Agreement", HumanizeFlexibleBenefitValue(ReadString(reader, "agreement_type")), page);
            AddFlexibleBenefitDetail(model.Agreements, "Status", status, page);
            AddFlexibleBenefitDetail(model.Agreements, "Execution date", execution.HasValue ? execution.Value.ToString("MMM d, yyyy", CultureInfo.InvariantCulture) : "Not confirmed", page);
            AddFlexibleBenefitDetail(model.Agreements, "Subscriber signer", FirstNonBlank(ReadString(reader, "subscriber_signer_name"), "Not confirmed"), page);
            AddFlexibleBenefitDetail(model.Agreements, "Administrator signer", FirstNonBlank(ReadString(reader, "administrator_signer_name"), "Not confirmed"), page);
            AddFlexibleBenefitDetail(model.Agreements, "Governing province", FirstNonBlank(ReadString(reader, "governing_province"), "Not confirmed"), page);
        }

        return summary;
    }

    private static async Task ReadFlexibleBenefitContactsAsync(
        SqliteConnection connection,
        long planId,
        FlexibleBenefitWorkspaceViewModel model)
    {
        using var command = CreateCommand(connection, null, @"
SELECT contact_role, contact_name, email, phone, source_page_number
FROM FlexibleBenefitContact
WHERE flexible_benefit_plan_id = $plan_id
ORDER BY is_primary DESC, contact_role COLLATE NOCASE, contact_name COLLATE NOCASE;");
        AddParameter(command, "$plan_id", planId);
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            var detail = string.Join(" | ", new[]
            {
                ReadString(reader, "contact_name"),
                ReadString(reader, "email"),
                ReadString(reader, "phone")
            }.Where(value => !string.IsNullOrWhiteSpace(value)));
            AddFlexibleBenefitDetail(
                model.Contacts,
                HumanizeFlexibleBenefitValue(ReadString(reader, "contact_role")),
                FirstNonBlank(detail, "Not confirmed"),
                FormatFlexibleBenefitSourcePage(ReadInt(reader, "source_page_number")));
        }
    }

    private static async Task ReadFlexibleBenefitSourceFieldsAsync(
        SqliteConnection connection,
        long sourceDocumentVersionId,
        FlexibleBenefitWorkspaceViewModel model)
    {
        using var command = CreateCommand(connection, null, @"
SELECT flexible_benefit_plan_source_field_id, entity_type, field_name, raw_value,
       normalized_value, source_page_number, extraction_confidence,
       review_status, was_manually_corrected
FROM FlexibleBenefitPlanSourceField
WHERE source_document_version_id = $source_document_version_id
ORDER BY source_page_number, entity_type COLLATE NOCASE, field_name COLLATE NOCASE;");
        AddParameter(command, "$source_document_version_id", sourceDocumentVersionId);
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            var confidence = ReadDecimal(reader, "extraction_confidence");
            model.SourceFields.Add(new FlexibleBenefitSourceFieldDisplayRow
            {
                SourceFieldId = ReadLong(reader, "flexible_benefit_plan_source_field_id"),
                EntityType = HumanizeFlexibleBenefitValue(ReadString(reader, "entity_type").Replace("FlexibleBenefit", "", StringComparison.Ordinal)),
                FieldName = HumanizeFlexibleBenefitValue(ReadString(reader, "field_name")),
                RawValue = ReadString(reader, "raw_value"),
                NormalizedValue = ReadString(reader, "normalized_value"),
                SourcePage = FormatFlexibleBenefitSourcePage(ReadInt(reader, "source_page_number")),
                Confidence = confidence.HasValue ? $"{confidence.Value:P0}" : "Not confirmed",
                ReviewStatus = ReadString(reader, "review_status"),
                WasManuallyCorrected = ReadNullableFlexibleBenefitBoolean(reader, "was_manually_corrected") == true
            });
        }
    }

    private static async Task ReadFlexibleBenefitVersionsAsync(
        SqliteConnection connection,
        long clientId,
        FlexibleBenefitWorkspaceViewModel model)
    {
        using var command = CreateCommand(connection, null, @"
SELECT fp.flexible_benefit_plan_id, fp.version_number, fp.effective_date,
       fp.review_status, fp.is_current, d.file_name
FROM FlexibleBenefitPlan fp
INNER JOIN Documents d ON d.document_id = fp.source_document_id
WHERE fp.client_id = $client_id
ORDER BY fp.version_number DESC;");
        AddParameter(command, "$client_id", clientId);
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            model.Versions.Add(new FlexibleBenefitVersionDisplayRow
            {
                PlanId = ReadLong(reader, "flexible_benefit_plan_id"),
                VersionNumber = ReadInt(reader, "version_number") ?? 1,
                EffectiveDate = FormatFlexibleBenefitDate(ReadDateOnly(reader, "effective_date")),
                Status = ReadString(reader, "review_status"),
                SourceFile = ReadString(reader, "file_name"),
                IsCurrent = ReadNullableFlexibleBenefitBoolean(reader, "is_current") == true
            });
        }
    }

    private static async Task ReadFlexibleBenefitReviewItemsAsync(
        SqliteConnection connection,
        long sourceDocumentId,
        FlexibleBenefitWorkspaceViewModel model)
    {
        using var command = CreateCommand(connection, null, @"
SELECT severity, field_name, message
FROM ValidationIssue
WHERE source_document_id = $source_document_id
ORDER BY CASE severity WHEN 'Error' THEN 0 WHEN 'Warning' THEN 1 ELSE 2 END,
         validation_issue_id;");
        AddParameter(command, "$source_document_id", sourceDocumentId);
        using var reader = await command.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            AddFlexibleBenefitDetail(
                model.ReviewItems,
                FirstNonBlank(ReadString(reader, "field_name"), ReadString(reader, "severity")),
                ReadString(reader, "message"),
                "");
        }
    }

    private static void AddFlexibleBenefitDetail(
        System.Collections.ObjectModel.ObservableCollection<FlexibleBenefitDetailDisplayRow> rows,
        string label,
        string value,
        string sourcePage)
    {
        rows.Add(new FlexibleBenefitDetailDisplayRow
        {
            Label = label,
            Value = FirstNonBlank(value, "Not confirmed"),
            SourcePage = sourcePage
        });
    }

    private static bool? ReadNullableFlexibleBenefitBoolean(SqliteDataReader reader, string name)
    {
        var ordinal = reader.GetOrdinal(name);
        return reader.IsDBNull(ordinal) ? null : Convert.ToInt32(reader.GetValue(ordinal), CultureInfo.InvariantCulture) == 1;
    }

    private static TEnum ParseFlexibleBenefitEnum<TEnum>(string value, TEnum fallback)
        where TEnum : struct, Enum
    {
        return Enum.TryParse<TEnum>(value, true, out var parsed) ? parsed : fallback;
    }

    private static string FormatFlexibleBenefitMoney(decimal? value)
    {
        return value.HasValue ? value.Value.ToString("C0", CultureInfo.GetCultureInfo("en-CA")) : "Not confirmed";
    }

    private static string FormatFlexibleBenefitDate(DateOnly? value)
    {
        return value.HasValue ? value.Value.ToString("MMM d, yyyy", CultureInfo.InvariantCulture) : "Not confirmed";
    }

    private static string FormatFlexibleBenefitBoolean(bool? value)
    {
        return value.HasValue ? value.Value ? "Yes" : "No" : "Not confirmed";
    }

    private static string FormatFlexibleBenefitSourcePage(int? page)
    {
        return page.HasValue ? $"Page {page.Value:#,0}" : "";
    }

    private static string FormatFlexibleBenefitAllocation(FlexibleBenefitPlan plan)
    {
        var allocation = HumanizeFlexibleBenefitValue(plan.DefaultAllocationType.ToString());
        if (plan.DefaultHcsaPercentage.HasValue || plan.DefaultLsaPercentage.HasValue)
        {
            return $"{allocation} ({plan.DefaultHcsaPercentage?.ToString("0.##", CultureInfo.InvariantCulture) ?? "?"}% HCSA / {plan.DefaultLsaPercentage?.ToString("0.##", CultureInfo.InvariantCulture) ?? "?"}% LSA)";
        }

        return allocation;
    }

    private static string FormatFlexibleBenefitDependentAges(FlexibleBenefitPlan plan)
    {
        if (!plan.ChildDependentCutoffAge.HasValue && !plan.StudentDependentCutoffAge.HasValue)
        {
            return "Not confirmed";
        }

        return $"Child {plan.ChildDependentCutoffAge?.ToString(CultureInfo.InvariantCulture) ?? "?"}; student {plan.StudentDependentCutoffAge?.ToString(CultureInfo.InvariantCulture) ?? "?"}";
    }

    private static string FormatFlexibleBenefitFunding(FlexibleBenefitPlan plan)
    {
        var parts = new List<string>();
        if (plan.FundingMethod != FlexibleBenefitFundingMethod.Unknown)
        {
            parts.Add(HumanizeFlexibleBenefitValue(plan.FundingMethod.ToString()));
        }

        if (plan.PadRequested.HasValue)
        {
            parts.Add(plan.PadRequested.Value ? "PAD requested" : "PAD not selected");
        }

        if (plan.PayAsYouGo == true)
        {
            parts.Add("pay as you go available");
        }

        return parts.Count > 0 ? string.Join("; ", parts) : "Not confirmed";
    }

    private static string FormatFlexibleBenefitFees(FlexibleBenefitPlan plan)
    {
        var parts = new List<string>();
        if (plan.SetupFee.HasValue)
        {
            parts.Add($"{FormatFlexibleBenefitMoney(plan.SetupFee)} setup");
        }

        if (plan.AdministrationFeePercentage.HasValue)
        {
            parts.Add($"{plan.AdministrationFeePercentage.Value:0.##}% administration");
        }
        else if (plan.AdministrationFeeFlatAmount.HasValue)
        {
            parts.Add($"{FormatFlexibleBenefitMoney(plan.AdministrationFeeFlatAmount)} administration");
        }

        if (!string.IsNullOrWhiteSpace(plan.TaxTreatment))
        {
            parts.Add(plan.TaxTreatment);
        }

        return parts.Count > 0 ? string.Join("; ", parts) : "Not confirmed";
    }

    private static string HumanizeFlexibleBenefitValue(string value)
    {
        if (string.IsNullOrWhiteSpace(value) || value.Equals("Unknown", StringComparison.OrdinalIgnoreCase))
        {
            return "Not confirmed";
        }

        var text = System.Text.RegularExpressions.Regex.Replace(value.Trim(), "(?<=[a-z0-9])(?=[A-Z])", " ");
        return text
            .Replace("Hcsa", "HCSA", StringComparison.OrdinalIgnoreCase)
            .Replace("Lsa", "LSA", StringComparison.OrdinalIgnoreCase)
            .Replace("Pad", "PAD", StringComparison.OrdinalIgnoreCase);
    }

    private sealed record FlexibleBenefitPlanRecord(
        FlexibleBenefitPlan Plan,
        string Administrator,
        string SourceDocumentPath);
}

using System.Globalization;
using Microsoft.Data.Sqlite;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public sealed partial class DatabaseService
{
    public async Task<IReadOnlyList<FlexibleBenefitVersionDifferenceRow>> CompareFlexibleBenefitVersionsAsync(
        long currentPlanId,
        long priorPlanId)
    {
        if (currentPlanId <= 0 || priorPlanId <= 0 || currentPlanId == priorPlanId)
        {
            return [];
        }

        await InitializeCoreAsync();
        using var connection = await OpenConnectionAsync();
        var current = await ReadFlexibleBenefitVersionValuesAsync(connection, currentPlanId);
        var prior = await ReadFlexibleBenefitVersionValuesAsync(connection, priorPlanId);
        return current.Keys
            .Union(prior.Keys, StringComparer.OrdinalIgnoreCase)
            .Select(key =>
            {
                current.TryGetValue(key, out var currentValue);
                prior.TryGetValue(key, out var priorValue);
                var identity = currentValue ?? priorValue!;
                return new FlexibleBenefitVersionDifferenceRow
                {
                    Area = identity.Area,
                    Field = identity.Field,
                    PriorValue = FirstNonBlank(priorValue?.Value, "Not available"),
                    CurrentValue = FirstNonBlank(currentValue?.Value, "Not available")
                };
            })
            .Where(row => !row.PriorValue.Equals(row.CurrentValue, StringComparison.OrdinalIgnoreCase))
            .OrderBy(row => row.Area, StringComparer.OrdinalIgnoreCase)
            .ThenBy(row => row.Field, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static async Task<Dictionary<string, FlexibleBenefitVersionValue>> ReadFlexibleBenefitVersionValuesAsync(
        SqliteConnection connection,
        long planId)
    {
        var values = new Dictionary<string, FlexibleBenefitVersionValue>(StringComparer.OrdinalIgnoreCase);
        using (var command = CreateCommand(connection, null, @"
SELECT plan_type, hcsa_enabled, lsa_enabled, effective_date, benefit_year_type,
       default_allocation_type, default_hcsa_percentage, default_lsa_percentage,
       runoff_days, carry_forward_type, child_dependent_cutoff_age,
       student_dependent_cutoff_age, pay_as_submitted, funding_method,
       pad_requested, setup_fee, administration_fee_percentage, tax_treatment
FROM FlexibleBenefitPlan
WHERE flexible_benefit_plan_id = $plan_id
LIMIT 1;"))
        {
            AddParameter(command, "$plan_id", planId);
            using var reader = await command.ExecuteReaderAsync();
            if (await reader.ReadAsync())
            {
                AddFlexibleBenefitVersionValue(values, "Plan", "Plan type", ReadString(reader, "plan_type"));
                AddFlexibleBenefitVersionValue(values, "Plan", "HCSA enabled", FormatFlexibleBenefitBoolean(ReadNullableFlexibleBenefitBoolean(reader, "hcsa_enabled")));
                AddFlexibleBenefitVersionValue(values, "Plan", "LSA enabled", FormatFlexibleBenefitBoolean(ReadNullableFlexibleBenefitBoolean(reader, "lsa_enabled")));
                AddFlexibleBenefitVersionValue(values, "Plan", "Effective date", FormatFlexibleBenefitDate(ReadDateOnly(reader, "effective_date")));
                AddFlexibleBenefitVersionValue(values, "Plan", "Benefit year", ReadString(reader, "benefit_year_type"));
                AddFlexibleBenefitVersionValue(values, "Plan", "Default allocation", ReadString(reader, "default_allocation_type"));
                AddFlexibleBenefitVersionValue(values, "Plan", "HCSA percentage", FormatFlexibleBenefitVersionPercent(ReadDecimal(reader, "default_hcsa_percentage")));
                AddFlexibleBenefitVersionValue(values, "Plan", "LSA percentage", FormatFlexibleBenefitVersionPercent(ReadDecimal(reader, "default_lsa_percentage")));
                AddFlexibleBenefitVersionValue(values, "Plan", "Runoff", FormatFlexibleBenefitVersionInteger(ReadInt(reader, "runoff_days"), "days"));
                AddFlexibleBenefitVersionValue(values, "Plan", "Carry forward", ReadString(reader, "carry_forward_type"));
                AddFlexibleBenefitVersionValue(values, "Plan", "Child dependent age", FormatFlexibleBenefitVersionInteger(ReadInt(reader, "child_dependent_cutoff_age"), ""));
                AddFlexibleBenefitVersionValue(values, "Plan", "Student dependent age", FormatFlexibleBenefitVersionInteger(ReadInt(reader, "student_dependent_cutoff_age"), ""));
                AddFlexibleBenefitVersionValue(values, "Plan", "Pay as submitted", FormatFlexibleBenefitBoolean(ReadNullableFlexibleBenefitBoolean(reader, "pay_as_submitted")));
                AddFlexibleBenefitVersionValue(values, "Funding", "Funding method", ReadString(reader, "funding_method"));
                AddFlexibleBenefitVersionValue(values, "Funding", "PAD requested", FormatFlexibleBenefitBoolean(ReadNullableFlexibleBenefitBoolean(reader, "pad_requested")));
                AddFlexibleBenefitVersionValue(values, "Fees", "Setup fee", FormatFlexibleBenefitMoney(ReadDecimal(reader, "setup_fee")));
                AddFlexibleBenefitVersionValue(values, "Fees", "Administration fee", FormatFlexibleBenefitVersionPercent(ReadDecimal(reader, "administration_fee_percentage")));
                AddFlexibleBenefitVersionValue(values, "Fees", "Tax treatment", ReadString(reader, "tax_treatment"));
            }
        }

        using (var command = CreateCommand(connection, null, @"
SELECT class_name, annual_single_amount, annual_couple_amount, annual_family_amount,
       waiting_period_days, new_hire_proration_enabled
FROM FlexibleBenefitClass
WHERE flexible_benefit_plan_id = $plan_id
ORDER BY class_name COLLATE NOCASE;"))
        {
            AddParameter(command, "$plan_id", planId);
            using var reader = await command.ExecuteReaderAsync();
            while (await reader.ReadAsync())
            {
                var className = FirstNonBlank(ReadString(reader, "class_name"), "Unnamed class");
                AddFlexibleBenefitVersionValue(values, "Class Limits", $"{className} - Single", FormatFlexibleBenefitMoney(ReadDecimal(reader, "annual_single_amount")));
                AddFlexibleBenefitVersionValue(values, "Class Limits", $"{className} - Couple", FormatFlexibleBenefitMoney(ReadDecimal(reader, "annual_couple_amount")));
                AddFlexibleBenefitVersionValue(values, "Class Limits", $"{className} - Family", FormatFlexibleBenefitMoney(ReadDecimal(reader, "annual_family_amount")));
                AddFlexibleBenefitVersionValue(values, "Class Limits", $"{className} - Waiting period", FormatFlexibleBenefitVersionInteger(ReadInt(reader, "waiting_period_days"), "days"));
                AddFlexibleBenefitVersionValue(values, "Class Limits", $"{className} - New-hire proration", FormatFlexibleBenefitBoolean(ReadNullableFlexibleBenefitBoolean(reader, "new_hire_proration_enabled")));
            }
        }

        using (var command = CreateCommand(connection, null, @"
SELECT category_code, category_name, enabled
FROM FlexibleBenefitCategory
WHERE flexible_benefit_plan_id = $plan_id
ORDER BY category_code;"))
        {
            AddParameter(command, "$plan_id", planId);
            using var reader = await command.ExecuteReaderAsync();
            while (await reader.ReadAsync())
            {
                AddFlexibleBenefitVersionValue(
                    values,
                    "LSA Categories",
                    FirstNonBlank(ReadString(reader, "category_name"), ReadString(reader, "category_code")),
                    FormatFlexibleBenefitBoolean(ReadNullableFlexibleBenefitBoolean(reader, "enabled")),
                    ReadString(reader, "category_code"));
            }
        }

        using (var command = CreateCommand(connection, null, @"
SELECT fee_type, percentage, flat_amount, tax_type, notice_period_days
FROM FlexibleBenefitFee
WHERE flexible_benefit_plan_id = $plan_id
ORDER BY fee_type;"))
        {
            AddParameter(command, "$plan_id", planId);
            using var reader = await command.ExecuteReaderAsync();
            while (await reader.ReadAsync())
            {
                var feeType = FirstNonBlank(ReadString(reader, "fee_type"), "Other fee");
                var percentage = ReadDecimal(reader, "percentage");
                var amount = ReadDecimal(reader, "flat_amount");
                var value = percentage.HasValue
                    ? FormatFlexibleBenefitVersionPercent(percentage)
                    : FormatFlexibleBenefitMoney(amount);
                var tax = ReadString(reader, "tax_type");
                var notice = ReadInt(reader, "notice_period_days");
                if (!string.IsNullOrWhiteSpace(tax))
                {
                    value = $"{value}; {tax}";
                }

                if (notice.HasValue)
                {
                    value = $"{value}; {notice.Value:#,0}-day notice";
                }

                AddFlexibleBenefitVersionValue(values, "Fees", feeType, value, feeType);
            }
        }

        return values;
    }

    private static void AddFlexibleBenefitVersionValue(
        IDictionary<string, FlexibleBenefitVersionValue> values,
        string area,
        string field,
        string value,
        string? keyOverride = null)
    {
        var key = $"{area}|{FirstNonBlank(keyOverride, field)}";
        values[key] = new FlexibleBenefitVersionValue(area, field, FirstNonBlank(value, "Not confirmed"));
    }

    private static string FormatFlexibleBenefitVersionPercent(decimal? value)
    {
        return value.HasValue ? $"{value.Value:0.##}%" : "Not confirmed";
    }

    private static string FormatFlexibleBenefitVersionInteger(int? value, string unit)
    {
        return value.HasValue
            ? string.IsNullOrWhiteSpace(unit) ? value.Value.ToString(CultureInfo.InvariantCulture) : $"{value.Value:#,0} {unit}"
            : "Not confirmed";
    }

    private sealed record FlexibleBenefitVersionValue(string Area, string Field, string Value);
}

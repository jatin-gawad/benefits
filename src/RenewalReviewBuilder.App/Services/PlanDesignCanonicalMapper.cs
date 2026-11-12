using System.Globalization;
using System.Text.RegularExpressions;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public static class PlanDesignCanonicalMapper
{
    public const string ProvisionIdentityVersion = "CANONICAL_PROVISION_IDENTITY_V1";

    public static void CompleteFact(PlanDesignCanonicalFact fact)
    {
        var lineage = CanonicalSourceLineageValueGuardrails.ResolveSourceFileName(
            fact.SourceDocumentId,
            FirstNonBlank(fact.TrackedSourceFileName, fact.SourceFileName),
            fact.SourceDocumentId.HasValue ? fact.SourceDocument : "");
        fact.SourceFileName = lineage.ResolvedSourceFileName;
        fact.SourceLineageStatus = lineage.Status;
        fact.SourceLineageMessage = lineage.Message;

        var percentage = CanonicalSourceLineageValueGuardrails.EvaluatePercentage(
            fact.ReimbursementPercent,
            fact.Coinsurance,
            fact.FeatureValue,
            fact.SourceText);
        fact.ReimbursementPercent = percentage.NormalizedValue;
        fact.ValueGuardrailStatus = percentage.Status;
        fact.ValueGuardrailMessage = percentage.Message;

        var benefitText = FirstNonBlank(fact.BenefitName, fact.BenefitCategory, fact.FeatureName);
        fact.NormalizedBenefitKey = NormalizeBenefitKey(benefitText);
        fact.CanonicalBenefitCode = ResolveCanonicalBenefitCode(benefitText);
        fact.BenefitCategory = FirstNonBlank(fact.BenefitCategory, NormalizeBenefitCategory(benefitText));
        if (fact.CanonicalBenefitCode == CanonicalBenefitTypeCodes.Dental)
        {
            fact.FeatureName = NormalizeDentalProvisionName(
                FirstNonBlank(fact.FeatureName, fact.BenefitName, fact.BenefitCategory));
        }
        fact.ClassName = FirstNonBlank(fact.ClassName, "All Classes");
        fact.ComparisonKey = BuildComparisonKey(fact);
        fact.Availability = ResolveAvailability(fact);
        fact.SourceCarrier = FirstNonBlank(fact.SourceCarrier, fact.Carrier);
        fact.RawSourceReference = FirstNonBlank(fact.RawSourceReference, BuildSourceReference(fact));
    }

    public static string ResolveCanonicalBenefitCode(string? value)
    {
        var direct = BenefitMatcher.ResolveCanonicalBenefitTypeCode(value ?? "");
        if (!string.IsNullOrWhiteSpace(direct))
        {
            return direct;
        }

        return NormalizeBenefitKey(value) switch
        {
            "ehc" => CanonicalBenefitTypeCodes.Ehc,
            "dental" => CanonicalBenefitTypeCodes.Dental,
            "vision" => "VISION",
            "basic life" => CanonicalBenefitTypeCodes.Life,
            "add" => CanonicalBenefitTypeCodes.AdAndD,
            "std" => "STD",
            "ltd" => CanonicalBenefitTypeCodes.Ltd,
            "critical illness" => "CRITICAL_ILLNESS",
            _ => ""
        };
    }

    public static string NormalizeBenefitCategory(string? value)
    {
        var key = NormalizeBenefitKey(value);
        return key switch
        {
            "ehc" => "EHC",
            "dental" => "Dental",
            "vision" => "Vision",
            "basic life" => "Basic Life",
            "add" => "AD&D",
            "std" => "STD",
            "ltd" => "LTD",
            "critical illness" => "Critical Illness",
            _ => FirstNonBlank(value, "Other")
        };
    }

    public static string NormalizeBenefitKey(string? value)
    {
        var text = NormalizeComparableText(value);
        if (string.IsNullOrWhiteSpace(text))
        {
            return "";
        }

        if (text.Contains("basic restorative", StringComparison.Ordinal) &&
            text.Contains("endodont", StringComparison.Ordinal) &&
            text.Contains("periodont", StringComparison.Ordinal))
        {
            return "Basic Services";
        }

        if (text.Contains("accidental death", StringComparison.Ordinal) ||
            text.Contains("dismemberment", StringComparison.Ordinal) ||
            text.Equals("ad d", StringComparison.Ordinal) ||
            text.Equals("add", StringComparison.Ordinal))
        {
            return "add";
        }

        if (text.Contains("long term disability", StringComparison.Ordinal) || text.Equals("ltd", StringComparison.Ordinal))
        {
            return "ltd";
        }

        if (text.Contains("short term disability", StringComparison.Ordinal) || text.Equals("std", StringComparison.Ordinal))
        {
            return "std";
        }

        if (text.Contains("critical illness", StringComparison.Ordinal))
        {
            return "critical illness";
        }

        if (text.Contains("life", StringComparison.Ordinal) && !text.Contains("disability", StringComparison.Ordinal))
        {
            return "basic life";
        }

        if (text.Contains("dental", StringComparison.Ordinal))
        {
            return "dental";
        }

        if (text.Contains("vision", StringComparison.Ordinal))
        {
            return "vision";
        }

        if (text.Contains("drug", StringComparison.Ordinal) ||
            text.Contains("prescription", StringComparison.Ordinal) ||
            text.Contains("paramedical", StringComparison.Ordinal) ||
            text.Contains("health", StringComparison.Ordinal) ||
            text.Contains("medical", StringComparison.Ordinal) ||
            text.Equals("ehc", StringComparison.Ordinal))
        {
            return "ehc";
        }

        return text;
    }

    public static string NormalizeComparableText(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return "";
        }

        var cleaned = Regex.Replace(value.Trim().ToLowerInvariant(), "[^a-z0-9%$]+", " ");
        return Regex.Replace(cleaned, "\\s+", " ").Trim();
    }

    public static string NormalizeDentalProvisionName(string? value)
    {
        var original = value?.Trim() ?? "";
        var text = NormalizeComparableText(value);
        if (string.IsNullOrWhiteSpace(text))
        {
            return "";
        }

        var matches = new List<string>();
        if (text.Contains("accidental dental injury", StringComparison.Ordinal) ||
            text.Contains("dental accidental injury", StringComparison.Ordinal))
        {
            matches.Add("Accidental Dental Injury");
        }
        if (text.Contains("all other treatment", StringComparison.Ordinal) ||
            text.Contains("other dental treatment", StringComparison.Ordinal))
        {
            matches.Add("Other Dental Treatment");
        }
        if (Regex.IsMatch(text, @"\blevel\s+(?:v|5)\b", RegexOptions.IgnoreCase) ||
            text.Contains("orthodont", StringComparison.Ordinal) ||
            Regex.IsMatch(text, @"\bortho\b", RegexOptions.IgnoreCase))
        {
            matches.Add("Orthodontic Services");
        }
        if (Regex.IsMatch(text, @"\blevel\s+(?:iv|4)\b", RegexOptions.IgnoreCase) ||
            text.Contains("major", StringComparison.Ordinal) ||
            text.Contains("dental restorative", StringComparison.Ordinal))
        {
            matches.Add("Major Restorative Services");
        }
        if (Regex.IsMatch(text, @"\blevel\s+(?:iii|3)\b", RegexOptions.IgnoreCase) ||
            text.Contains("denture", StringComparison.Ordinal))
        {
            matches.Add("Dentures");
        }
        if (text.Contains("endodont", StringComparison.Ordinal) &&
            text.Contains("periodont", StringComparison.Ordinal))
        {
            matches.Add("Endodontic / Periodontic Services");
        }
        if (text.Contains("routine dental", StringComparison.Ordinal) ||
            text.Contains("dental routine", StringComparison.Ordinal))
        {
            matches.Add("Routine Dental Care");
        }
        if (text.Contains("preventive", StringComparison.Ordinal) ||
            text.Contains("preventative", StringComparison.Ordinal))
        {
            matches.Add("Preventive Services");
        }
        if (Regex.IsMatch(text, @"\blevel\s+(?:i|1)\b", RegexOptions.IgnoreCase) ||
            Regex.IsMatch(text, @"\bbasic\b", RegexOptions.IgnoreCase))
        {
            matches.Add("Basic Services");
        }

        return matches.Distinct(StringComparer.OrdinalIgnoreCase).ToList() switch
        {
            [var single] => single,
            _ => original
        };
    }

    public static List<PlanDesignRecord> NormalizeProvisionRows(IEnumerable<PlanDesignRecord> sourceRows)
    {
        var rows = sourceRows.Select(ClonePlanDesignRow).ToList();
        rows = rows.SelectMany(ExpandLabeledMultiRateRow).ToList();
        rows = RemoveDerivedCostSharingRows(rows);
        rows = RemoveSyntheticSummaryRows(rows);
        rows = RemoveRedundantDentalSummaryRows(rows);

        foreach (var group in rows.GroupBy(
                     row => $"{NormalizeComparableText(FirstNonBlank(row.ClassName, "All Classes"))}|{ResolveProvisionBenefitCode(row)}",
                     StringComparer.OrdinalIgnoreCase))
        {
            var hasSiblingRows = group.Count() > 1;
            foreach (var row in group)
            {
                if (!string.IsNullOrWhiteSpace(row.ProvisionCode) &&
                    !string.IsNullOrWhiteSpace(row.ProvisionLabel))
                {
                    continue;
                }

                var label = ResolveProvisionLabel(row, hasSiblingRows);
                row.ProvisionLabel = label;
                row.ProvisionCode = BuildProvisionCode(label);
            }
        }

        return rows;
    }

    private static List<PlanDesignRecord> RemoveSyntheticSummaryRows(IEnumerable<PlanDesignRecord> rows)
    {
        return rows
            .Where(row =>
            {
                var category = NormalizeComparableText(row.BenefitCategory);
                var benefit = NormalizeComparableText(row.BenefitName);
                var provision = NormalizeComparableText(FirstNonBlank(row.ProvisionCode, row.ProvisionLabel));
                return category != "class summary" &&
                       benefit != "class benchmark summary" &&
                       provision != "class benchmark summary";
            })
            .ToList();
    }

    public static string ResolveProvisionBenefitCode(PlanDesignRecord row)
    {
        var category = NormalizeComparableText(row.BenefitCategory);
        var categoryCombinesLifeAndAdAndD =
            category.Contains("life", StringComparison.Ordinal) &&
            category.Contains("ad d", StringComparison.Ordinal);
        var value = categoryCombinesLifeAndAdAndD && !string.IsNullOrWhiteSpace(row.BenefitName)
            ? NormalizeComparableText(row.BenefitName)
            : NormalizeComparableText(FirstNonBlank(
                row.BenefitCategory,
                row.SourceSection,
                row.BenefitName));
        if (string.IsNullOrWhiteSpace(value))
        {
            return "";
        }

        if (value.Contains("dental", StringComparison.Ordinal))
        {
            if (value.Contains("orthodont", StringComparison.Ordinal)) return "DENTAL_ORTHODONTIC";
            if (value.Contains("major", StringComparison.Ordinal)) return "DENTAL_MAJOR";
            if (value.Contains("basic", StringComparison.Ordinal)) return "DENTAL_BASIC";
            if (value.Contains("recall", StringComparison.Ordinal)) return "DENTAL_RECALL";
            return "DENTAL";
        }

        if (value.Contains("prescription", StringComparison.Ordinal) || value == "drugs" || value == "drug")
            return "PRESCRIPTION_DRUGS";
        if (value.Contains("medical services", StringComparison.Ordinal) || value.Contains("medical supplies", StringComparison.Ordinal))
            return "MEDICAL_SERVICES_SUPPLIES";
        if (value.Contains("paramed", StringComparison.Ordinal)) return "PARAMEDICAL";
        if (value.Contains("vision", StringComparison.Ordinal)) return "VISION";
        if (value.Contains("hospital", StringComparison.Ordinal)) return "HOSPITAL";
        if (value.Contains("travel", StringComparison.Ordinal)) return "TRAVEL";
        if (value.Contains("survivor", StringComparison.Ordinal)) return "SURVIVOR_BENEFIT";
        if (value.Contains("virtual", StringComparison.Ordinal)) return "VIRTUAL_CARE";
        if (value == "eap" || value.Contains("employee assistance", StringComparison.Ordinal)) return "EAP";
        if (value.Contains("medical second opinion", StringComparison.Ordinal)) return "MEDICAL_SECOND_OPINION";
        if (value.Contains("dependent life", StringComparison.Ordinal)) return "DEPENDENT_LIFE";
        if (value.Contains("life", StringComparison.Ordinal) && value.Contains("ad d", StringComparison.Ordinal)) return "LIFE_AD_AND_D";
        if (value == "ad d" || value.Contains("accidental death", StringComparison.Ordinal)) return "AD_AND_D";
        if (value == "life" || value.Contains("group life", StringComparison.Ordinal)) return "LIFE";
        if (value == "std" || value.Contains("short term disability", StringComparison.Ordinal)) return "STD";
        if (value == "ltd" || value.Contains("long term disability", StringComparison.Ordinal)) return "LTD";
        if (value.Contains("critical illness", StringComparison.Ordinal)) return "CRITICAL_ILLNESS";
        if (value == "ehc" || value.Contains("extended health", StringComparison.Ordinal)) return "EHC";

        return BuildProvisionCode(value);
    }

    public static string BuildProvisionComparisonKey(PlanDesignRecord row)
    {
        var benefitCode = ResolveProvisionBenefitCode(row);
        var className = NormalizeComparableText(FirstNonBlank(row.ClassName, "All Classes"));
        if (className is "" or "all classes" or "all")
        {
            className = "all";
        }

        var provisionCode = FirstNonBlank(
            row.ProvisionCode,
            BuildProvisionCode(row.ProvisionLabel),
            BuildProvisionCode(ResolveProvisionLabel(row, hasSiblingRows: false)));
        return $"{FirstNonBlank(benefitCode, "UNMAPPED")}|{className}|{FirstNonBlank(provisionCode, "UNRESOLVED")}";
    }

    private static IEnumerable<PlanDesignRecord> ExpandLabeledMultiRateRow(PlanDesignRecord row)
    {
        var firstLine = FirstDescriptionLine(row.CoverageDescription);
        var separator = firstLine.IndexOf(':');
        if (separator <= 0)
        {
            return [row];
        }

        var fieldLabel = NormalizeComparableText(firstLine[..separator]);
        if (fieldLabel is not ("coinsurance" or "reimbursement" or "reimbursement level" or "reimbursement levels"))
        {
            return [row];
        }

        var segments = firstLine[(separator + 1)..]
            .Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (segments.Length < 2)
        {
            return [row];
        }

        var expanded = new List<PlanDesignRecord>();
        foreach (var segment in segments)
        {
            if (!TryParseLabeledRate(segment, out var context, out var percentage))
            {
                return [row];
            }

            var clone = ClonePlanDesignRow(row);
            var contextLabel = Capitalize(context);
            clone.ProvisionLabel = $"{contextLabel} Coinsurance";
            clone.ProvisionCode = BuildProvisionCode(clone.ProvisionLabel);
            clone.CoverageDescription = $"{contextLabel} Coinsurance: {percentage:0.####}%";
            clone.ReimbursementPercent = percentage;
            clone.ReimbursementText = $"{percentage:0.####}%";
            clone.Notes = AppendCanonicalNote(
                clone.Notes,
                $"{ProvisionIdentityVersion}: split labeled multi-rate provision.");
            expanded.Add(clone);
        }

        return expanded;
    }

    private static bool TryParseLabeledRate(string segment, out string context, out decimal percentage)
    {
        context = "";
        percentage = 0m;
        var percentFirst = Regex.Match(
            segment,
            @"^(?<percent>100(?:\.0+)?|\d{1,2}(?:\.\d+)?)\s*%\s*(?:(?:for|at|on)\s+)?(?<context>.+?)\s*$",
            RegexOptions.IgnoreCase);
        var labelFirst = percentFirst.Success
            ? Match.Empty
            : Regex.Match(
                segment,
                @"^(?<context>.+?)\s*[:\-]?\s*(?<percent>100(?:\.0+)?|\d{1,2}(?:\.\d+)?)\s*%\s*$",
                RegexOptions.IgnoreCase);
        var match = percentFirst.Success ? percentFirst : labelFirst;
        if (!match.Success ||
            !decimal.TryParse(match.Groups["percent"].Value, NumberStyles.Number, CultureInfo.InvariantCulture, out percentage) ||
            percentage is < 0m or > 100m)
        {
            return false;
        }

        context = Regex.Replace(match.Groups["context"].Value.Trim(" .:-".ToCharArray()), @"^(?:for|at|on)\s+", "", RegexOptions.IgnoreCase);
        return !string.IsNullOrWhiteSpace(context);
    }

    private static List<PlanDesignRecord> RemoveDerivedCostSharingRows(List<PlanDesignRecord> rows)
    {
        var groupsWithCoinsurance = rows
            .Where(IsCoinsuranceProvision)
            .Select(BuildBenefitClassGroupKey)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        return rows
            .Where(row => !IsCostSharingProvision(row) || !groupsWithCoinsurance.Contains(BuildBenefitClassGroupKey(row)))
            .ToList();
    }

    private static List<PlanDesignRecord> RemoveRedundantDentalSummaryRows(List<PlanDesignRecord> rows)
    {
        var result = new List<PlanDesignRecord>();
        foreach (var classGroup in rows.GroupBy(row => NormalizeComparableText(FirstNonBlank(row.ClassName, "All Classes"))))
        {
            var classRows = classGroup.ToList();
            var benefitCodes = classRows.Select(ResolveProvisionBenefitCode).ToHashSet(StringComparer.OrdinalIgnoreCase);
            foreach (var row in classRows)
            {
                if (!ResolveProvisionBenefitCode(row).Equals("DENTAL", StringComparison.OrdinalIgnoreCase))
                {
                    result.Add(row);
                    continue;
                }

                var label = NormalizeComparableText(ExtractLeadingLabel(row.CoverageDescription));
                var redundant =
                    (label == "basic services" && benefitCodes.Contains("DENTAL_BASIC")) ||
                    (label == "major restorative services" && benefitCodes.Contains("DENTAL_MAJOR")) ||
                    (label == "orthodontics" && benefitCodes.Contains("DENTAL_ORTHODONTIC")) ||
                    (label == "basic and major maximum" &&
                     (benefitCodes.Contains("DENTAL_BASIC") || benefitCodes.Contains("DENTAL_MAJOR"))) ||
                    (label == "orthodontic maximum" && benefitCodes.Contains("DENTAL_ORTHODONTIC"));
                if (!redundant)
                {
                    result.Add(row);
                }
            }
        }

        return result;
    }

    private static string ResolveProvisionLabel(PlanDesignRecord row, bool hasSiblingRows)
    {
        if (!string.IsNullOrWhiteSpace(row.ProvisionLabel))
        {
            return row.ProvisionLabel.Trim();
        }

        var benefitCode = ResolveProvisionBenefitCode(row);
        var specific = benefitCode switch
        {
            "DENTAL_BASIC" => "Basic Services",
            "DENTAL_MAJOR" => "Major Restorative Services",
            "DENTAL_ORTHODONTIC" => "Orthodontic Services",
            "DENTAL_RECALL" => "Dental Recall",
            _ => ""
        };
        if (!string.IsNullOrWhiteSpace(specific))
        {
            return specific;
        }

        var benefitName = row.BenefitName.Trim();
        if (!string.IsNullOrWhiteSpace(benefitName) &&
            !NormalizeComparableText(benefitName).Equals(
                NormalizeComparableText(row.BenefitCategory),
                StringComparison.OrdinalIgnoreCase))
        {
            return NormalizeDentalProvisionName(benefitName);
        }

        var hasMultipleDescriptionLines = HasMultipleDescriptionLines(row.CoverageDescription);
        var descriptionLabel = ExtractLeadingLabel(row.CoverageDescription);
        if (!hasSiblingRows &&
            NormalizeComparableText(descriptionLabel) is "coverage" or "benefit coverage")
        {
            return "Plan Summary";
        }

        if (!hasMultipleDescriptionLines && !string.IsNullOrWhiteSpace(descriptionLabel))
        {
            return descriptionLabel;
        }

        if (!hasSiblingRows || hasMultipleDescriptionLines)
        {
            return "Plan Summary";
        }

        return descriptionLabel;
    }

    private static string BuildBenefitClassGroupKey(PlanDesignRecord row) =>
        $"{NormalizeComparableText(FirstNonBlank(row.ClassName, "All Classes"))}|{ResolveProvisionBenefitCode(row)}";

    private static bool IsCoinsuranceProvision(PlanDesignRecord row)
    {
        var label = NormalizeComparableText(FirstNonBlank(row.ProvisionLabel, ExtractLeadingLabel(row.CoverageDescription)));
        return label.Contains("coinsurance", StringComparison.Ordinal) || label == "reimbursement";
    }

    private static bool IsCostSharingProvision(PlanDesignRecord row) =>
        NormalizeComparableText(ExtractLeadingLabel(row.CoverageDescription)) == "cost sharing";

    private static bool HasMultipleDescriptionLines(string value) =>
        value.Replace("\r\n", "\n", StringComparison.Ordinal)
            .Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Length > 1;

    private static string FirstDescriptionLine(string value) =>
        value.Replace("\r\n", "\n", StringComparison.Ordinal)
            .Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .FirstOrDefault() ?? "";

    private static string ExtractLeadingLabel(string value)
    {
        var line = FirstDescriptionLine(value);
        var match = Regex.Match(line, @"^(?<label>[A-Za-z][A-Za-z0-9 &/()'.+\-]{1,100}):\s*.+$");
        return match.Success ? match.Groups["label"].Value.Trim() : "";
    }

    private static string BuildProvisionCode(string? value)
    {
        var code = Regex.Replace(value?.Trim().ToUpperInvariant() ?? "", "[^A-Z0-9]+", "_").Trim('_');
        return code;
    }

    private static string Capitalize(string value) =>
        string.IsNullOrWhiteSpace(value)
            ? ""
            : char.ToUpperInvariant(value[0]) + value[1..];

    private static string AppendCanonicalNote(string notes, string note) =>
        string.IsNullOrWhiteSpace(notes)
            ? note
            : notes.Contains(note, StringComparison.OrdinalIgnoreCase) ? notes : $"{notes} {note}";

    private static PlanDesignRecord ClonePlanDesignRow(PlanDesignRecord row) => new()
    {
        PlanDesignId = row.PlanDesignId,
        ClientId = row.ClientId,
        ClientName = row.ClientName,
        ClientNumber = row.ClientNumber,
        Carrier = row.Carrier,
        ClassName = row.ClassName,
        BenefitCategory = row.BenefitCategory,
        BenefitName = row.BenefitName,
        ProvisionCode = row.ProvisionCode,
        ProvisionLabel = row.ProvisionLabel,
        CoverageDescription = row.CoverageDescription,
        ReimbursementPercent = row.ReimbursementPercent,
        ReimbursementText = row.ReimbursementText,
        MaximumAmount = row.MaximumAmount,
        Frequency = row.Frequency,
        WaitingPeriod = row.WaitingPeriod,
        TerminationAge = row.TerminationAge,
        SourceDocumentId = row.SourceDocumentId,
        SourceFileName = row.SourceFileName,
        SourceSection = row.SourceSection,
        Notes = row.Notes,
        SourcePage = row.SourcePage,
        DateLoaded = row.DateLoaded
    };

    public static decimal? ParseConfidence(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        var text = value.Trim();
        if (decimal.TryParse(text.TrimEnd('%'), NumberStyles.Any, CultureInfo.InvariantCulture, out var numeric))
        {
            return text.EndsWith('%') || numeric > 1m ? numeric / 100m : numeric;
        }

        return text.ToLowerInvariant() switch
        {
            "high" => 0.90m,
            "medium" => 0.70m,
            "low" => 0.40m,
            _ => null
        };
    }

    public static string ResolveSourceType(string? fileName)
    {
        return Path.GetExtension(fileName ?? "").ToLowerInvariant() switch
        {
            ".pdf" => SourceTypes.Pdf,
            ".xls" or ".xlsx" or ".xlsm" => SourceTypes.Excel,
            _ => ""
        };
    }

    public static bool IsFailedValidation(PlanDesignCanonicalFact fact)
    {
        var status = fact.ValidationStatus;
        if (string.IsNullOrWhiteSpace(status))
        {
            return false;
        }

        return !status.Equals("valid", StringComparison.OrdinalIgnoreCase) &&
               !status.Equals("match", StringComparison.OrdinalIgnoreCase) &&
               !status.Equals("matched", StringComparison.OrdinalIgnoreCase) &&
               !status.Equals("pass", StringComparison.OrdinalIgnoreCase) &&
               !status.Equals("passes", StringComparison.OrdinalIgnoreCase);
    }

    private static string BuildComparisonKey(PlanDesignCanonicalFact fact)
    {
        var benefit = FirstNonBlank(fact.CanonicalBenefitCode, fact.NormalizedBenefitKey, "unmapped").ToLowerInvariant();
        var className = NormalizeComparableText(fact.ClassName);
        if (className is "" or "all classes" or "all")
        {
            className = "all";
        }

        var provision = NormalizeComparableText(FirstNonBlank(fact.FeatureName, fact.BenefitName, fact.BenefitCategory));
        return $"{benefit}|{className}|{provision}";
    }

    private static string ResolveAvailability(PlanDesignCanonicalFact fact)
    {
        if (IsFailedValidation(fact) ||
            fact.SourceLineageStatus == CanonicalSourceLineageStatuses.NeedsReview ||
            (fact.Scope == PlanDesignCanonicalScopes.CurrentPlan &&
             !string.IsNullOrWhiteSpace(fact.SourceId) &&
             fact.SourceLineageStatus is CanonicalSourceLineageStatuses.Unlinked or CanonicalSourceLineageStatuses.Missing) ||
            fact.ValueGuardrailStatus is CanonicalValueGuardrailStatuses.PreservedForReview or CanonicalValueGuardrailStatuses.Rejected)
        {
            return PlanDesignCanonicalAvailability.NeedsReview;
        }

        var hasValue = fact.ReimbursementPercent.HasValue ||
                       new[]
                       {
                           fact.FeatureValue,
                           fact.BookletValue,
                           fact.Coinsurance,
                           fact.Deductible,
                           fact.Maximum,
                           fact.Frequency,
                           fact.WaitingPeriod,
                           fact.TerminationAge
                       }.Any(value => !string.IsNullOrWhiteSpace(value));
        if (!hasValue)
        {
            return PlanDesignCanonicalAvailability.Missing;
        }

        return string.IsNullOrWhiteSpace(fact.CanonicalBenefitCode) ||
               (fact.ConfidenceScore.HasValue && fact.ConfidenceScore.Value < 0.70m)
            ? PlanDesignCanonicalAvailability.Partial
            : PlanDesignCanonicalAvailability.Available;
    }

    private static string BuildSourceReference(PlanDesignCanonicalFact fact)
    {
        var location = fact.SourcePage.HasValue
            ? $"page {fact.SourcePage.Value.ToString(CultureInfo.InvariantCulture)}"
            : fact.SourceSheet;
        return string.Join(" | ", new[] { fact.SourceDocument, location }.Where(value => !string.IsNullOrWhiteSpace(value)));
    }

    private static string FirstNonBlank(params string?[] values)
    {
        return values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value))?.Trim() ?? "";
    }
}

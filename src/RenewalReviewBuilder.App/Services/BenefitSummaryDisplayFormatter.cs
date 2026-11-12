using System.Globalization;
using System.Text.RegularExpressions;
using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services;

public static class BenefitSummaryDisplayFormatter
{
    public const string DefaultDisclaimer =
        "E&OE: Should there be a discrepancy between this summary and the group contract, the group contract will be considered correct and legally binding.";

    private static readonly string[] SectionOrder =
    [
        "Carrier & Policy",
        "Eligibility Requirements",
        "Life / AD&D",
        "Dependent Life",
        "Optional Life",
        "Short-Term Disability",
        "Long-Term Disability",
        "Critical Illness",
        "Prescription Drugs",
        "Medical Services & Supplies",
        "Paramedicals",
        "Mental Health Practitioners",
        "Hospital",
        "Vision Care",
        "Eye Exams",
        "Emergency Travel",
        "Dental Care",
        "Health Spending Account",
        "Wellness Spending Account",
        "Health & Dental Benefit Period",
        "Employee Assistance Program",
        "Virtual Health Care",
        "Medical Second Opinion",
        "Digital Wellness Program"
    ];

    private static readonly string[] KnownProviders =
    [
        "Manulife Financial",
        "Pacific Blue Cross",
        "Equitable Life",
        "Empire Life",
        "Canada Life",
        "Green Shield",
        "GreenShield",
        "Sun Life",
        "Blue Cross",
        "Allstate",
        "Desjardins",
        "Homewood Health",
        "TELUS Health",
        "LifeWorks",
        "National HealthClaim",
        "Manulife"
    ];

    private static readonly string[] MentalHealthPractitionerTerms =
    [
        "psychologist",
        "social worker",
        "counsellor",
        "counselor",
        "marriage and family therapist",
        "psychoanalyst",
        "psychotherapist"
    ];

    private static readonly string[] NonSummaryLabels =
    [
        "notes",
        "description",
        "policy wording",
        "source text",
        "raw value",
        "clause",
        "definitions"
    ];

    private static readonly string[] PolicyNarrativeTerms =
    [
        "insured person",
        "terms and conditions",
        "subject to the terms",
        "notice of claim",
        "proof of claim",
        "the insurer will pay",
        "we must receive",
        "right to deny",
        "as used in this benefit",
        "as defined in this section",
        "under the policy",
        "under this benefit"
    ];

    public static IReadOnlyList<BenefitSummaryViewModel> Build(BenefitSummaryBuildRequest request)
    {
        ArgumentNullException.ThrowIfNull(request);

        var rows = request.Rows
            .Where(HasMeaningfulValue)
            .ToList();
        var concreteClasses = rows
            .Select(row => NormalizeClassKey(row.ClassName))
            .Where(value => !IsGenericClass(value))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (concreteClasses.Count == 0)
        {
            concreteClasses.Add("All Classes");
        }

        return concreteClasses
            .Select(className => BuildSummary(
                request,
                className,
                rows.Where(row => IsGenericClass(NormalizeClassKey(row.ClassName)) ||
                                  NormalizeClassKey(row.ClassName).Equals(className, StringComparison.OrdinalIgnoreCase))
                    .ToList()))
            .ToList();
    }

    public static string NormalizeBenefitName(string value)
    {
        var text = CleanWhitespace(value);
        if (string.IsNullOrWhiteSpace(text))
        {
            return "";
        }

        if (Regex.IsMatch(text, @"\bLTD\b|long[ -]?term disability", RegexOptions.IgnoreCase)) return "Long-Term Disability";
        if (Regex.IsMatch(text, @"\bSTD\b|short[ -]?term disability", RegexOptions.IgnoreCase)) return "Short-Term Disability";
        if (Regex.IsMatch(text, @"\bAD\s*&\s*D\b|accidental death", RegexOptions.IgnoreCase)) return "AD&D";
        if (Regex.IsMatch(text, @"\b(?:dep(?:endent|endant)?|dependant)\s+life\b", RegexOptions.IgnoreCase)) return "Dependent Life";
        if (text.Contains("optional life", StringComparison.OrdinalIgnoreCase)) return "Optional Life";
        if (text.Contains("critical illness", StringComparison.OrdinalIgnoreCase)) return "Critical Illness";
        if (Regex.IsMatch(text, @"medical second opinion|Teladoc Medical Experts", RegexOptions.IgnoreCase)) return "Medical Second Opinion";
        if (Regex.IsMatch(text, @"virtual (?:health )?care|Lumino Health Virtual Care", RegexOptions.IgnoreCase)) return "Virtual Health Care";
        if (Regex.IsMatch(text, @"\b(?:prescription )?drugs?\b", RegexOptions.IgnoreCase)) return "Prescription Drugs";
        if (Regex.IsMatch(text, @"\b(?:other medical(?: expenses?)?|medical (?:services|supplies)(?:\s*&\s*| and )?(?:services|supplies)?)\b", RegexOptions.IgnoreCase)) return "Medical Services & Supplies";
        if (Regex.IsMatch(text, @"\bpara\s*medical", RegexOptions.IgnoreCase)) return "Paramedicals";
        if (Regex.IsMatch(text, @"mental health|psychological services", RegexOptions.IgnoreCase)) return "Mental Health Practitioners";
        if (text.Contains("dental", StringComparison.OrdinalIgnoreCase)) return "Dental Care";
        if (text.Contains("vision", StringComparison.OrdinalIgnoreCase)) return "Vision Care";
        if (Regex.IsMatch(text, @"\b(?:HCSA|HSCA|HSA)\b|health(?: care)? spending account", RegexOptions.IgnoreCase)) return "Health Spending Account";
        if (Regex.IsMatch(text, @"\b(?:LSA|WSA)\b|lifestyle spending account|wellness(?: spending)? account", RegexOptions.IgnoreCase)) return "Wellness Spending Account";
        return text;
    }

    public static string NormalizeLabel(string value)
    {
        var text = CleanWhitespace(value).Trim(' ', ':', '-', '_');
        if (string.IsNullOrWhiteSpace(text))
        {
            return "";
        }

        var key = Regex.Replace(text, @"[^a-z0-9]+", " ", RegexOptions.IgnoreCase).Trim().ToLowerInvariant();
        if (key is "benefit amt" or "benefit amount" or "life benefit amount") return "Benefit Amount";
        if (key is "ad d waiver of premium") return "Waiver of Premium";
        if (key is "max") return "Maximum";
        if (key is "nel" or "non evidence limit" or "non evidence maximum") return "Non-Evidence Maximum";
        if (key is "elim period" or "elimination period" or "qualification period") return "Qualification Period";
        if (key is "own occ" or "own occupation" or "definition of disability") return "Definition of Disability";
        if (key is "term age" or "termination age") return "Termination";
        if (key is "co ins" or "co insurance" or "coinsurance" or "reimbursement percentage") return "Coinsurance";
        if (key is "ded" or "deductible" or "dental deductible") return "Deductible";
        if (key is "annual max" or "annual maximum") return "Annual Maximum";
        if (key is "lifetime max" or "lifetime maximum") return "Lifetime Maximum";
        if (key is "freq" or "frequency") return "Frequency";
        if (key is "fee guide yr" or "fee guide year" or "fee guide" or "dental fee guide") return "Fee Guide";
        if (key is "dep spouse" or "dependent spouse") return "Spouse";
        if (key is "dep child" or "dependent child") return "Child";
        var isDentalProvisionLabel =
            key is "basic" or "basic services" or "preventive services" or "preventative services" or
                "routine dental care" or "dentures" or "major" or "major restorative" or "ortho" or
                "orthodontics" or "orthodontic services" or "all other treatment" ||
            key.StartsWith("level ", StringComparison.Ordinal) ||
            key.Contains("dental", StringComparison.Ordinal) ||
            key.Contains("restorative", StringComparison.Ordinal) ||
            key.Contains("orthodont", StringComparison.Ordinal) ||
            key.Contains("denture", StringComparison.Ordinal) ||
            (key.Contains("endodont", StringComparison.Ordinal) && key.Contains("periodont", StringComparison.Ordinal));
        var dentalProvision = isDentalProvisionLabel
            ? PlanDesignCanonicalMapper.NormalizeDentalProvisionName(text)
            : "";
        if (dentalProvision is "Preventive Services" or "Basic Services" or "Endodontic / Periodontic Services" or
            "Routine Dental Care" or "Dentures" or "Major Restorative Services" or "Orthodontic Services" or
            "Accidental Dental Injury" or "Other Dental Treatment")
        {
            return dentalProvision;
        }
        if (key is "diagnostic preventative") return "Basic Services";
        if (key is "ltd benefit formula" or "benefit formula" or "benefit schedule") return "Benefit Schedule";
        if (key is "ltd taxability" or "taxability" or "tax status") return "Tax Status";
        if (key.Contains("waiting period", StringComparison.Ordinal)) return "Waiting Period";
        if (key.Contains("cost sharing", StringComparison.Ordinal)) return "Cost Sharing";
        if (key.Contains("maximum", StringComparison.Ordinal) && key.Contains("per trip", StringComparison.Ordinal)) return "Per-Trip Maximum";
        return text;
    }

    public static string NormalizePractitionerName(string value)
    {
        var text = CleanWhitespace(value).Trim(' ', '-', '\u2022', '.', ':');
        if (string.IsNullOrWhiteSpace(text))
        {
            return "";
        }

        if (Regex.IsMatch(text, @"^chiro(?:practic|practor)?$", RegexOptions.IgnoreCase)) return "Chiropractor";
        if (Regex.IsMatch(text, @"^physio(?:therapy|therapist)?$", RegexOptions.IgnoreCase)) return "Physiotherapist";
        if (Regex.IsMatch(text, @"^massage(?: therapy| therapist)?$", RegexOptions.IgnoreCase)) return "Massage Therapist";
        if (Regex.IsMatch(text, @"^(?:registered )?psychologist$", RegexOptions.IgnoreCase)) return "Psychologist";
        if (Regex.IsMatch(text, @"^clinical counsel(?:l)?or$", RegexOptions.IgnoreCase)) return "Clinical Counsellor";
        if (Regex.IsMatch(text, @"^podiatrist\s*/\s*chiropodist$", RegexOptions.IgnoreCase)) return "Podiatrist / Chiropodist";
        if (text.Equals("social worker", StringComparison.OrdinalIgnoreCase)) return "Social Worker";
        if (text.Equals("marriage and family therapist", StringComparison.OrdinalIgnoreCase)) return "Marriage and Family Therapist";
        return CultureInfo.InvariantCulture.TextInfo.ToTitleCase(text.ToLowerInvariant());
    }

    public static string FormatCurrency(decimal value, string? unit = null)
    {
        var amount = value == decimal.Truncate(value)
            ? value.ToString("$#,##0", CultureInfo.InvariantCulture)
            : value.ToString("$#,##0.00", CultureInfo.InvariantCulture);
        return string.IsNullOrWhiteSpace(unit) ? amount : $"{amount} {NormalizeUnit(unit)}";
    }

    public static string FormatPercentage(decimal value, BenefitSummaryPercentageScale scale)
    {
        var percentage = scale == BenefitSummaryPercentageScale.Decimal ? value * 100m : value;
        return $"{percentage.ToString(percentage == decimal.Truncate(percentage) ? "0" : "0.##", CultureInfo.InvariantCulture)}%";
    }

    public static string FormatWaitingPeriod(string value)
    {
        var text = CleanWhitespace(value);
        if (Regex.IsMatch(text, @"^(?:none|nil|n/?a|0(?:\.0+)?(?:\s*(?:day|days|month|months))?)$", RegexOptions.IgnoreCase))
        {
            return "No waiting period";
        }

        return text;
    }

    public static string FormatFrequency(string value)
    {
        var text = CleanWhitespace(value);
        var monthMatch = Regex.Match(text, @"^(?:once\s+)?(?:every\s+)?(?<months>\d+)\s+months?$", RegexOptions.IgnoreCase);
        if (monthMatch.Success)
        {
            return $"Once every {monthMatch.Groups["months"].Value} months";
        }

        var ageQualifiedMonthMatch = Regex.Match(
            text,
            @"^(?:once\s+)?(?:every\s+)?(?<months>\d+)\s+months?\s+(?:for\s+)?(?<qualifier>under\s+age\s+\d+|age\s+\d+\s+and\s+over)$",
            RegexOptions.IgnoreCase);
        if (ageQualifiedMonthMatch.Success)
        {
            var qualifier = ageQualifiedMonthMatch.Groups["qualifier"].Value.ToLowerInvariant();
            return $"Once every {ageQualifiedMonthMatch.Groups["months"].Value} months for {qualifier}";
        }

        if (Regex.IsMatch(text, @"^(?:annual|annually|per year|(?:per )?calendar year)$", RegexOptions.IgnoreCase)) return "Per calendar year";
        if (Regex.IsMatch(text, @"^(?:monthly|per month)$", RegexOptions.IgnoreCase)) return "Per month";
        if (Regex.IsMatch(text, @"^lifetime$", RegexOptions.IgnoreCase)) return "Lifetime maximum";
        return text;
    }

    public static string FormatDisplayValue(string value, string label = "")
    {
        var text = CleanClientFacingLine(value);
        if (string.IsNullOrWhiteSpace(text))
        {
            return "";
        }

        if (label.Contains("Waiting Period", StringComparison.OrdinalIgnoreCase))
        {
            text = FormatWaitingPeriod(text);
        }
        else if (label.Contains("Frequency", StringComparison.OrdinalIgnoreCase) ||
                 label.Contains("Recall", StringComparison.OrdinalIgnoreCase))
        {
            text = FormatFrequency(text);
        }
        else if (label.Contains("Deductible", StringComparison.OrdinalIgnoreCase) &&
                 Regex.IsMatch(text, @"^(?:none|nil)$", RegexOptions.IgnoreCase))
        {
            text = "None";
        }

        if ((label.Contains("Age", StringComparison.OrdinalIgnoreCase) ||
             label.Contains("Termination", StringComparison.OrdinalIgnoreCase)) &&
            Regex.IsMatch(text, @"^\d{1,3}$"))
        {
            text = $"Age {text}";
        }

        text = Regex.Replace(
            text,
            @"^(?:termination|terminate)\s+at\s+(?:age\s+)?(?<age>\d{1,3})$",
            "Terminates at age ${age}",
            RegexOptions.IgnoreCase);
        if (label.Contains("Reduction", StringComparison.OrdinalIgnoreCase))
        {
            text = Regex.Replace(text, @"(?<percent>\d+(?:\.\d+)?\s*%)\s+at\s+(?:age\s+)?(?<age>\d{1,3})\b", "${percent} at age ${age}", RegexOptions.IgnoreCase);
        }

        var unmarkedAmount = Regex.Match(text, @"^(?<amount>\d[\d,]*(?:\.\d{1,2})?)\s+(?<unit>monthly|annual|annually|lifetime)$", RegexOptions.IgnoreCase);
        if (unmarkedAmount.Success && IsMoneyLabel(label) &&
            decimal.TryParse(unmarkedAmount.Groups["amount"].Value.Replace(",", "", StringComparison.Ordinal), NumberStyles.Number, CultureInfo.InvariantCulture, out var numericAmount))
        {
            text = FormatCurrency(numericAmount, unmarkedAmount.Groups["unit"].Value);
        }

        text = Regex.Replace(text, @"\$(?<amount>\d[\d,]*(?:\.\d{1,2})?)", match =>
        {
            if (!decimal.TryParse(match.Groups["amount"].Value.Replace(",", "", StringComparison.Ordinal), NumberStyles.Number, CultureInfo.InvariantCulture, out var amount))
            {
                return match.Value;
            }

            return amount == decimal.Truncate(amount)
                ? amount.ToString("$#,##0", CultureInfo.InvariantCulture)
                : amount.ToString("$#,##0.00", CultureInfo.InvariantCulture);
        });
        text = Regex.Replace(text, @"(?<amount>\$[\d,]+(?:\.\d{1,2})?)\s+monthly\b", "${amount} per month", RegexOptions.IgnoreCase);
        text = Regex.Replace(text, @"(?<amount>\$[\d,]+(?:\.\d{1,2})?)\s+(?:annual|annually)\b", "${amount} per calendar year", RegexOptions.IgnoreCase);
        return CleanWhitespace(text);
    }

    /// <summary>
    /// Applies the same client-facing value gate used by Benefit Summary. Narrative policy wording is rejected.
    /// </summary>
    public static bool TryFormatClientFacingValue(
        string sectionTitle,
        string label,
        string value,
        out string normalizedLabel,
        out string displayValue)
    {
        return TryPrepareSummaryLine(sectionTitle, label, value, out normalizedLabel, out displayValue);
    }

    public static bool TryFormatLongTermDisabilityBenefitFormula(
        string sectionTitle,
        string value,
        out string displayValue)
    {
        displayValue = "";
        if (!sectionTitle.Equals("Long-Term Disability", StringComparison.OrdinalIgnoreCase) &&
            !sectionTitle.Equals("Disability", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        var text = CleanWhitespace(value);
        var formula = Regex.Match(
            text,
            @"(?<tier1>\d{1,3}(?:\.\d+)?)\s*%\s+of\s+(?:the\s+)?first\s+\$?\s*(?<amount1>\d[\d,]*(?:\.\d{1,2})?)(?:\s+of\s+(?:your\s+)?(?:basic\s+)?monthly(?:\s+basic)?\s+earnings?)?[\s,;]*(?:plus|then)\s+(?<tier2>\d{1,3}(?:\.\d+)?)\s*%\s+of\s+(?:the\s+)?next\s+\$?\s*(?<amount2>\d[\d,]*(?:\.\d{1,2})?)(?:\s+of\s+(?:your\s+)?(?:basic\s+)?monthly(?:\s+basic)?\s+earnings?)?[\s,;]*(?:plus|then)\s+(?<tier3>\d{1,3}(?:\.\d+)?)\s*%\s+of\s+(?:the\s+)?(?:any\s+)?(?:excess(?:\s+amount)?|remainder|balance)",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        if (!formula.Success ||
            !decimal.TryParse(formula.Groups["tier1"].Value, NumberStyles.Number, CultureInfo.InvariantCulture, out var tier1Percent) ||
            !decimal.TryParse(formula.Groups["tier2"].Value, NumberStyles.Number, CultureInfo.InvariantCulture, out var tier2Percent) ||
            !decimal.TryParse(formula.Groups["tier3"].Value, NumberStyles.Number, CultureInfo.InvariantCulture, out var tier3Percent) ||
            !decimal.TryParse(formula.Groups["amount1"].Value, NumberStyles.Number, CultureInfo.InvariantCulture, out var tier1Amount) ||
            !decimal.TryParse(formula.Groups["amount2"].Value, NumberStyles.Number, CultureInfo.InvariantCulture, out var tier2Amount) ||
            tier1Percent is < 0m or > 100m ||
            tier2Percent is < 0m or > 100m ||
            tier3Percent is < 0m or > 100m ||
            tier1Amount <= 0m ||
            tier2Amount <= 0m)
        {
            return false;
        }

        var thresholdBasis = Regex.IsMatch(formula.Value, @"\bmonthly(?:\s+basic)?\s+earnings?\b", RegexOptions.IgnoreCase)
            ? " of monthly earnings"
            : "";
        displayValue = string.Join(
            "; ",
            $"{FormatPercentage(tier1Percent, BenefitSummaryPercentageScale.Whole)} of first {FormatCurrency(tier1Amount)}{thresholdBasis}",
            $"{FormatPercentage(tier2Percent, BenefitSummaryPercentageScale.Whole)} of next {FormatCurrency(tier2Amount)}{thresholdBasis}",
            $"{FormatPercentage(tier3Percent, BenefitSummaryPercentageScale.Whole)} of excess");

        var maximum = Regex.Match(
            text,
            @"\b(?:up\s+to\s+(?:a\s+)?)?maximum(?:\s+benefit)?\s+(?:of\s+)?\$?\s*(?<amount>\d[\d,]*(?:\.\d{1,2})?)",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        if (maximum.Success &&
            decimal.TryParse(maximum.Groups["amount"].Value, NumberStyles.Number, CultureInfo.InvariantCulture, out var maximumAmount) &&
            maximumAmount > 0m)
        {
            displayValue += $" (maximum {FormatCurrency(maximumAmount)})";
        }

        return true;
    }

    private static BenefitSummaryViewModel BuildSummary(
        BenefitSummaryBuildRequest request,
        string classKey,
        IReadOnlyList<BenefitSummarySourceRow> rows)
    {
        var sections = new Dictionary<string, MutableSection>(StringComparer.OrdinalIgnoreCase);
        var hasPrescriptionDrugEvidence = rows.Any(IsPrescriptionDrugRow);
        foreach (var row in rows)
        {
            var sectionTitle = ResolveSection(row, hasPrescriptionDrugEvidence);
            if (string.IsNullOrWhiteSpace(sectionTitle))
            {
                continue;
            }

            if (sectionTitle.Equals("Paramedicals", StringComparison.OrdinalIgnoreCase) &&
                AddParamedicalRows(sections, row))
            {
                continue;
            }

            AddStandardRow(GetSection(sections, sectionTitle), row, sectionTitle);
        }

        NormalizePrescriptionDrugSection(sections, request, rows, hasPrescriptionDrugEvidence);
        NormalizeMedicalServicesSection(sections, rows, hasPrescriptionDrugEvidence);
        ApplyTieredReimbursementRules(sections, rows, hasPrescriptionDrugEvidence);

        var carrierSection = BuildCarrierSection(request, rows, hasPrescriptionDrugEvidence);
        if (carrierSection.Lines.Count > 0)
        {
            sections[carrierSection.Title] = carrierSection;
        }

        var finalizedSections = sections.Values
            .Select(FinalizeSection)
            .Where(section => section.Lines.Count > 0)
            .OrderBy(section => section.DisplayOrder)
            .ThenBy(section => section.SectionTitle, StringComparer.OrdinalIgnoreCase)
            .ToList();

        var (className, optionName) = SplitClassAndOption(classKey);
        var subtitleSelection = string.Join(" ", new[] { className, optionName }.Where(value => !string.IsNullOrWhiteSpace(value)));
        return new BenefitSummaryViewModel
        {
            ClientName = request.ClientName,
            SummaryTitle = string.IsNullOrWhiteSpace(subtitleSelection)
                ? "Employee Benefits Summary"
                : $"Employee Benefits Summary - {subtitleSelection}",
            ClassName = className,
            OptionName = optionName,
            EffectiveDate = NormalizeMissingValue(request.EffectiveDate),
            UpdatedDate = request.UpdatedDate,
            UpdatedDateSource = request.UpdatedDateSource,
            DisclaimerText = DefaultDisclaimer,
            Sections = finalizedSections
        };
    }

    private static void AddStandardRow(MutableSection section, BenefitSummarySourceRow row, string sectionTitle)
    {
        var rawValue = FirstNonBlank(row.RawValue, row.Deductible, row.Maximum, row.WaitingPeriod, row.TerminationAge);
        if (string.IsNullOrWhiteSpace(rawValue))
        {
            return;
        }

        if (ShouldSuppressRedundantCombinedLifeAmount(sectionTitle, row, rawValue))
        {
            return;
        }

        var evidence = $"{rawValue}\n{row.SourceText}";
        var reconciledSameAsLife = sectionTitle.Equals("Life / AD&D", StringComparison.OrdinalIgnoreCase) &&
            IsAdAndD(row) && evidence.Contains("same as life", StringComparison.OrdinalIgnoreCase) &&
            evidence.Contains("contract/booklet", StringComparison.OrdinalIgnoreCase);
        if (reconciledSameAsLife)
        {
            rawValue = ReconcileSameAsLife(rawValue);
        }

        var defaultLabel = InferLabel(sectionTitle, row, rawValue);
        var lines = Regex.Replace(
                rawValue,
                @";\s*(?=[A-Za-z][A-Za-z0-9 &/()'\-]{1,80}:\s*)",
                "\n")
            .Replace("\r\n", "\n", StringComparison.Ordinal)
            .Replace('\r', '\n')
            .Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .ToList();
        var provider = SupportsProvider(sectionTitle) ? FindProvider(evidence) : "";
        if (!string.IsNullOrWhiteSpace(provider))
        {
            section.Add("Provider", provider, false, 0, 0);
        }

        var hasLongTermDisabilityFormula = false;
        for (var index = 0; index < lines.Count; index++)
        {
            var cleaned = CleanClientFacingLine(lines[index]);
            if (string.IsNullOrWhiteSpace(cleaned) || cleaned.StartsWith("Practitioners:", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var (label, value) = SplitLabelledLine(cleaned, index == 0 ? defaultLabel : "Notes");
            if (TryPrepareSummaryLine(sectionTitle, label, value, out var normalizedLabel, out var displayValue))
            {
                var lineOrder = GetLineOrder(sectionTitle, normalizedLabel);
                if (ShouldHideCombinedLifeAmountLabel(sectionTitle, row, normalizedLabel))
                {
                    normalizedLabel = "";
                }

                section.Add(normalizedLabel, displayValue, false, 0, lineOrder);
                hasLongTermDisabilityFormula |= normalizedLabel.Equals("Benefit Schedule", StringComparison.OrdinalIgnoreCase) ||
                                                normalizedLabel.Equals("Benefit Formula", StringComparison.OrdinalIgnoreCase);
            }
        }

        if (!hasLongTermDisabilityFormula &&
            !sectionTitle.Equals("Life / AD&D", StringComparison.OrdinalIgnoreCase))
        {
            var supplementalCoinsuranceLabel = ResolveSupplementalCoinsuranceLabel(sectionTitle, row);
            var supplementalCoinsurance = supplementalCoinsuranceLabel.Equals("Coinsurance", StringComparison.OrdinalIgnoreCase) ||
                                           HasValidPercentages(row.Coinsurance)
                ? row.Coinsurance
                : "";
            AddSupplementalLine(
                section,
                sectionTitle,
                supplementalCoinsuranceLabel,
                supplementalCoinsurance);
        }
        AddSupplementalLine(section, sectionTitle, "Deductible", row.Deductible);
        if (!reconciledSameAsLife &&
            !ShouldSuppressEmbeddedCombinedLifeMaximum(sectionTitle, row, rawValue))
        {
            AddSupplementalLine(
                section,
                sectionTitle,
                "Maximum",
                NormalizeDentalMaximumForDisplay(sectionTitle, row.Maximum));
        }
        if (!IsDentalRecallRow(sectionTitle, row))
        {
            AddSupplementalLine(section, sectionTitle, "Frequency", row.Frequency);
        }
        AddSupplementalLine(
            section,
            sectionTitle,
            sectionTitle.Equals("Long-Term Disability", StringComparison.OrdinalIgnoreCase) ||
            sectionTitle.Equals("Short-Term Disability", StringComparison.OrdinalIgnoreCase)
                ? "Qualification Period"
                : "Waiting Period",
            row.WaitingPeriod);
        AddSupplementalLine(section, sectionTitle, "Termination", row.TerminationAge);
    }

    private static bool AddParamedicalRows(
        IDictionary<string, MutableSection> sections,
        BenefitSummarySourceRow row)
    {
        var evidence = FirstNonBlank(row.RawValue, row.SourceText);
        var practitioners = ExtractPractitioners(evidence);
        if (practitioners.Count == 0)
        {
            return false;
        }

        var physical = practitioners.Where(name => !IsMentalHealthPractitioner(name)).ToList();
        var mental = practitioners.Where(IsMentalHealthPractitioner).ToList();
        var coinsurance = FirstNonBlank(row.Coinsurance, Regex.Match(evidence, @"(?<!\d)(?:100|[1-9]?\d(?:\.\d+)?)\s*%(?!\d)").Value);
        var provider = FindProvider($"{evidence}\n{row.SourceText}");
        var standardMaximum = ExtractStandardParamedicalMaximum(row, evidence);
        var mentalMaximum = ExtractMentalHealthMaximum($"{row.Maximum}\n{evidence}");

        if (physical.Count > 0)
        {
            var section = GetSection(sections, "Paramedicals");
            if (!string.IsNullOrWhiteSpace(provider)) section.Add("Provider", provider, false, 0, 0);
            AddSupplementalLine(section, "Paramedicals", "Coinsurance", coinsurance);
            AddSupplementalLine(section, "Paramedicals", "Benefit Amount", standardMaximum);
            if (evidence.Contains("no practitioner limit", StringComparison.OrdinalIgnoreCase)) section.Add("Practitioner Limit", "No practitioner limit", false, 0, 30);
            section.Add("Coverage", FormatPractitionerCount(physical.Count), false, 0, 100);
            foreach (var practitioner in physical)
            {
                section.Add("", practitioner, true, 1, 110);
            }
        }

        if (mental.Count > 0 || !string.IsNullOrWhiteSpace(mentalMaximum))
        {
            var section = GetSection(sections, "Mental Health Practitioners");
            if (!string.IsNullOrWhiteSpace(provider)) section.Add("Provider", provider, false, 0, 0);
            AddSupplementalLine(section, "Mental Health Practitioners", "Coinsurance", coinsurance);
            AddSupplementalLine(section, "Mental Health Practitioners", "Benefit Amount", mentalMaximum);
            if (mental.Count > 0)
            {
                section.Add("Coverage", FormatPractitionerCount(mental.Count), false, 0, 100);
                foreach (var practitioner in mental)
                {
                    section.Add("", practitioner, true, 1, 110);
                }
            }
        }

        return true;
    }

    private static MutableSection BuildCarrierSection(
        BenefitSummaryBuildRequest request,
        IReadOnlyList<BenefitSummarySourceRow> rows,
        bool hasPrescriptionDrugEvidence)
    {
        var section = new MutableSection("Carrier & Policy");
        var values = new List<BenefitSummaryCarrierPolicy>();
        if (!string.IsNullOrWhiteSpace(request.Carrier))
        {
            values.Add(new BenefitSummaryCarrierPolicy { Carrier = request.Carrier, PolicyNumber = request.PolicyNumber });
        }

        values.AddRange(request.CarrierPolicies);
        values.AddRange(rows
            .Where(row => !string.IsNullOrWhiteSpace(row.Carrier) &&
                          (!row.Carrier.Equals(request.Carrier, StringComparison.OrdinalIgnoreCase) ||
                           (!string.IsNullOrWhiteSpace(row.PolicyNumber) &&
                            !NormalizePolicy(row.PolicyNumber).Equals(NormalizePolicy(request.PolicyNumber), StringComparison.OrdinalIgnoreCase))))
            .Select(row => new BenefitSummaryCarrierPolicy
            {
                Carrier = row.Carrier,
                PolicyNumber = row.PolicyNumber,
                Coverage = ResolveSection(row, hasPrescriptionDrugEvidence)
            }));

        foreach (var row in rows)
        {
            var coverage = ResolveSection(row, hasPrescriptionDrugEvidence);
            if (!SupportsProvider(coverage))
            {
                continue;
            }

            var provider = FindProvider($"{row.RawValue}\n{row.SourceText}");
            if (!string.IsNullOrWhiteSpace(provider) &&
                (string.IsNullOrWhiteSpace(request.Carrier) ||
                 !NormalizeProviderName(request.Carrier).Equals(provider, StringComparison.OrdinalIgnoreCase)))
            {
                values.Add(new BenefitSummaryCarrierPolicy { Carrier = provider, Coverage = coverage });
            }
        }

        var carrierGroups = values
            .Where(value => !IsMissing(value.Carrier))
            .GroupBy(value => NormalizeProviderName(value.Carrier), StringComparer.OrdinalIgnoreCase)
            .OrderBy(group => group.Key, StringComparer.OrdinalIgnoreCase)
            .ToList();
        foreach (var group in carrierGroups)
        {
            var policies = group
                .Select(value => NormalizePolicy(value.PolicyNumber))
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            var coverages = group.Select(value => NormalizeMissingValue(value.Coverage))
                .Where(value => !string.IsNullOrWhiteSpace(value) && !value.Equals("Carrier & Policy", StringComparison.OrdinalIgnoreCase))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(SectionSortOrder)
                .ThenBy(value => value, StringComparer.OrdinalIgnoreCase)
                .ToList();
            var valueText = JoinPolicyNumbers(policies);
            if (coverages.Count > 0)
            {
                valueText = string.IsNullOrWhiteSpace(valueText)
                    ? string.Join(", ", coverages)
                    : $"{valueText} - {string.Join(", ", coverages)}";
            }

            if (string.IsNullOrWhiteSpace(valueText))
            {
                section.Add("Carrier", group.Key, false, 0, 10);
            }
            else
            {
                var label = carrierGroups.Count == 1 && policies.Count > 1 ? "" : group.Key;
                section.Add(label, valueText, false, 0, 10);
            }
        }

        return section;
    }

    private static void NormalizePrescriptionDrugSection(
        IDictionary<string, MutableSection> sections,
        BenefitSummaryBuildRequest request,
        IReadOnlyList<BenefitSummarySourceRow> rows,
        bool hasPrescriptionDrugEvidence)
    {
        if (!hasPrescriptionDrugEvidence)
        {
            return;
        }

        var drugRows = rows.Where(IsPrescriptionDrugRow).ToList();
        if (drugRows.Count == 0)
        {
            return;
        }

        var evidence = BuildEvidence(drugRows);
        var section = new MutableSection("Prescription Drugs");
        if (Regex.IsMatch(evidence, @"\b(?:coverage\s*:\s*)?not covered\b", RegexOptions.IgnoreCase))
        {
            section.Add("Coverage", "Not covered", false, 0, 10);
            sections[section.Title] = section;
            return;
        }

        var coinsurance = ExtractPreferredPercentage(drugRows, "Prescription Drugs");
        var hasConfirmedCoverage = !string.IsNullOrWhiteSpace(coinsurance) ||
            Regex.IsMatch(
                evidence,
                @"\bPrescription Drugs?\s*[-:]\s*(?:Included|Covered)\b|\b(?:Pay Direct|Direct) Drug Card\b|\bmandatory generic\b|\bgeneric substitution\b|\bFormuCARE\b",
                RegexOptions.IgnoreCase);
        if (!string.IsNullOrWhiteSpace(coinsurance))
        {
            var hasOptionPercentages = coinsurance.Contains("Option 1", StringComparison.OrdinalIgnoreCase);
            if (!hasOptionPercentages && evidence.Contains("mandatory generic", StringComparison.OrdinalIgnoreCase))
            {
                coinsurance = $"{coinsurance} for Mandatory Generic Drugs";
            }
            else if (!hasOptionPercentages && evidence.Contains("generic substitution", StringComparison.OrdinalIgnoreCase))
            {
                coinsurance = $"{coinsurance} with Generic Substitution";
            }
            else if (!hasOptionPercentages && evidence.Contains("FormuCARE", StringComparison.OrdinalIgnoreCase))
            {
                coinsurance = $"{coinsurance} for FormuCARE Drugs";
            }

            section.Add("Coinsurance", coinsurance, false, 0, GetLineOrder(section.Title, "Coinsurance"));
        }

        var relatedRows = rows
            .Where(row => ResolveSection(row, true).Equals("Prescription Drugs", StringComparison.OrdinalIgnoreCase))
            .Concat(drugRows)
            .Concat(rows.Where(IsExtendedHealthDeductibleRow))
            .Concat(rows.Where(IsMedicalServicesRow))
            .Distinct()
            .ToList();
        var deductible = ExtractHealthDeductible(relatedRows);
        if (!string.IsNullOrWhiteSpace(deductible))
        {
            section.Add("Deductible", deductible, false, 0, GetLineOrder(section.Title, "Deductible"));
        }

        var maximum = ExtractPrescriptionDrugMaximum(drugRows);
        if (!string.IsNullOrWhiteSpace(maximum) || hasConfirmedCoverage)
        {
            section.Add(
                "Maximum",
                string.IsNullOrWhiteSpace(maximum) ? "Unlimited" : maximum,
                false,
                0,
                GetLineOrder(section.Title, "Maximum"));
        }

        if (Regex.IsMatch(
                evidence,
                @"\b(?:covers|covered at) the lowest[- ]cost generic equivalent product\b",
                RegexOptions.IgnoreCase))
        {
            section.Add(
                "Generic Substitution",
                "Covered at the lowest-cost generic equivalent product",
                false,
                0,
                GetLineOrder(section.Title, "Generic Substitution"));
        }

        var electiveMedications = ExtractElectiveMedicationSummary(drugRows);
        if (!string.IsNullOrWhiteSpace(electiveMedications))
        {
            section.Add("Elective Medications", electiveMedications, false, 0, GetLineOrder(section.Title, "Elective Medications"));
        }

        section.Add(
            "Personalized Medicine Program",
            HasConfirmedManulifeDrugCoverage(request, drugRows) ? "Pharmacogenetic testing" : "Unknown",
            false,
            0,
            GetLineOrder(section.Title, "Personalized Medicine Program"));
        sections[section.Title] = section;
    }

    private static void ApplyTieredReimbursementRules(
        IDictionary<string, MutableSection> sections,
        IReadOnlyList<BenefitSummarySourceRow> rows,
        bool hasPrescriptionDrugEvidence)
    {
        var rulesBySection = rows
            .SelectMany(row => row.Rules
                .Where(rule => rule.RuleType == BenefitRuleType.TieredReimbursement)
                .Select(rule => new
                {
                    SectionTitle = ResolveSection(row, hasPrescriptionDrugEvidence),
                    DisplayText = BenefitRuleDisplayFormatter.FormatFull(rule),
                    RequiresReview = !rule.IsDisplayable
                }))
            .Where(item => !string.IsNullOrWhiteSpace(item.SectionTitle))
            .GroupBy(item => item.SectionTitle, StringComparer.OrdinalIgnoreCase);

        foreach (var sectionGroup in rulesBySection)
        {
            if (!sections.TryGetValue(sectionGroup.Key, out var section))
            {
                section = GetSection(sections, sectionGroup.Key);
            }

            section.Lines.RemoveAll(line =>
                line.Label.Equals("Coinsurance", StringComparison.OrdinalIgnoreCase) ||
                line.Label.Equals("Reimbursement", StringComparison.OrdinalIgnoreCase) ||
                line.Label.Equals("Cost Sharing", StringComparison.OrdinalIgnoreCase) ||
                Regex.IsMatch(
                    line.Value,
                    @"^\s*(?:100|[1-9]?\d(?:\.\d+)?)\s*%\s*\*?\s*$",
                    RegexOptions.IgnoreCase));

            var displayTexts = sectionGroup
                .Select(item => item.DisplayText)
                .Where(displayText => !string.IsNullOrWhiteSpace(displayText))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            var hasParsedDisabilitySchedule =
                (sectionGroup.Key.Equals("Long-Term Disability", StringComparison.OrdinalIgnoreCase) ||
                 sectionGroup.Key.Equals("Short-Term Disability", StringComparison.OrdinalIgnoreCase)) &&
                section.Lines.Any(line =>
                    (line.Label.Equals("Benefit Schedule", StringComparison.OrdinalIgnoreCase) ||
                     line.Label.Equals("Benefit Formula", StringComparison.OrdinalIgnoreCase)) &&
                    Regex.Matches(line.Value, @"(?<!\d)(?:100|[1-9]?\d(?:\.\d+)?)\s*%(?!\d)").Count >= 2);
            if (displayTexts.Count == 0 &&
                sectionGroup.Any(item => item.RequiresReview) &&
                !hasParsedDisabilitySchedule)
            {
                displayTexts.Add(BenefitRuleDisplayFormatter.ReviewRequiredDisplayText);
            }

            foreach (var displayText in displayTexts)
            {
                section.Add(
                    "Coinsurance",
                    displayText,
                    false,
                    0,
                    GetLineOrder(section.Title, "Coinsurance"));
            }
        }
    }

    private static void NormalizeMedicalServicesSection(
        IDictionary<string, MutableSection> sections,
        IReadOnlyList<BenefitSummarySourceRow> rows,
        bool hasPrescriptionDrugEvidence)
    {
        sections.TryGetValue("Medical Services & Supplies", out var existingSection);
        var medicalRows = rows.Where(IsMedicalServicesRow).ToList();
        var sourceEvidence = BuildEvidence(rows);
        var otherMedicalCoinsurance = ExtractOptionPercentages(sourceEvidence, "Other Medical Expenses");
        if (string.IsNullOrWhiteSpace(otherMedicalCoinsurance))
        {
            otherMedicalCoinsurance = ExtractLabelledPercentage(sourceEvidence, "Other Medical Expenses");
        }
        if (existingSection is null && medicalRows.Count == 0 && string.IsNullOrWhiteSpace(otherMedicalCoinsurance))
        {
            return;
        }

        var section = new MutableSection("Medical Services & Supplies");
        var coinsurance = ExtractPreferredPercentage(medicalRows, "Other Medical Expenses");
        if (string.IsNullOrWhiteSpace(coinsurance))
        {
            coinsurance = otherMedicalCoinsurance;
        }

        if (!string.IsNullOrWhiteSpace(coinsurance))
        {
            section.Add("Coinsurance", coinsurance, false, 0, GetLineOrder(section.Title, "Coinsurance"));
        }

        var deductibleRows = medicalRows
            .Concat(rows.Where(IsExtendedHealthDeductibleRow))
            .Concat(hasPrescriptionDrugEvidence ? rows.Where(IsPrescriptionDrugRow) : [])
            .Distinct()
            .ToList();
        var deductible = ExtractHealthDeductible(deductibleRows);
        if (!string.IsNullOrWhiteSpace(deductible))
        {
            section.Add("Deductible", deductible, false, 0, GetLineOrder(section.Title, "Deductible"));
        }

        if (existingSection is not null)
        {
            foreach (var line in existingSection.Lines.Where(line =>
                         !line.Label.Equals("Coinsurance", StringComparison.OrdinalIgnoreCase) &&
                         !line.Label.Equals("Cost Sharing", StringComparison.OrdinalIgnoreCase) &&
                         !line.Label.Equals("Deductible", StringComparison.OrdinalIgnoreCase) &&
                         !line.Label.Contains("Other Medical", StringComparison.OrdinalIgnoreCase)))
            {
                section.Add(line.Label, line.Value, line.IsBullet, line.IndentLevel, line.Order);
            }
        }

        sections[section.Title] = section;
    }

    private static string ExtractPreferredPercentage(
        IReadOnlyList<BenefitSummarySourceRow> rows,
        string scheduleLabel)
    {
        var evidence = BuildEvidence(rows);
        var optionPercentages = ExtractOptionPercentages(evidence, scheduleLabel);
        if (!string.IsNullOrWhiteSpace(optionPercentages))
        {
            return optionPercentages;
        }

        foreach (var candidate in rows.Select(row => row.Coinsurance))
        {
            if (TryNormalizePercentage(candidate, out var percentage))
            {
                return percentage;
            }
        }

        var labelled = ExtractLabelledPercentage(evidence, scheduleLabel);
        if (!string.IsNullOrWhiteSpace(labelled))
        {
            return labelled;
        }

        foreach (var row in rows)
        {
            if (TryNormalizePercentage(row.RawValue, out var percentage))
            {
                return percentage;
            }
        }

        return "";
    }

    private static string ExtractLabelledPercentage(string evidence, string label)
    {
        var labelPattern = label.Equals("Prescription Drugs", StringComparison.OrdinalIgnoreCase)
            ? @"Prescription Drugs?"
            : label.Equals("Other Medical Expenses", StringComparison.OrdinalIgnoreCase)
                ? @"Other Medical(?: Expenses?)?"
                : Regex.Escape(label);
        var match = Regex.Match(
            evidence,
            $@"\*?\s*{labelPattern}(?:\s*\([^)]{{1,80}}\))?\s*\**\s*[.\s]*(?<value>\d{{1,3}}(?:\.\d+)?)\s*%",
            RegexOptions.IgnoreCase);
        return match.Success && TryNormalizePercentage(match.Value, out var percentage) ? percentage : "";
    }

    private static string ExtractOptionPercentages(string evidence, string label)
    {
        var explicitOptions = Regex.Match(
            evidence,
            @"Option\s+1\s*:\s*(?<one>\d{1,3}%|N/?A)\s*[;/|]\s*Option\s+2\s*:\s*(?<two>\d{1,3}%|N/?A)\s*[;/|]\s*Option\s+3\s*:\s*(?<three>\d{1,3}%|N/?A)",
            RegexOptions.IgnoreCase);
        if (!explicitOptions.Success)
        {
            var labelPattern = label.Equals("Prescription Drugs", StringComparison.OrdinalIgnoreCase)
                ? @"Prescription Drugs?"
                : label.Equals("Other Medical Expenses", StringComparison.OrdinalIgnoreCase)
                    ? @"Other Medical(?: Expenses?)?"
                    : Regex.Escape(label);
            explicitOptions = Regex.Match(
                evidence,
                $@"{labelPattern}(?:\s*\([^)]{{1,80}}\))?\s+(?<one>\d{{1,3}}%|N/?A)\*?\s+(?<two>\d{{1,3}}%|N/?A)\*?\s+(?<three>\d{{1,3}}%|N/?A)\*?",
                RegexOptions.IgnoreCase);
        }

        return explicitOptions.Success
            ? $"Option 1: {explicitOptions.Groups["one"].Value} / Option 2: {explicitOptions.Groups["two"].Value} / Option 3: {explicitOptions.Groups["three"].Value}"
            : "";
    }

    private static bool TryNormalizePercentage(string value, out string percentage)
    {
        percentage = "";
        if (string.IsNullOrWhiteSpace(value) || !HasValidPercentages(value))
        {
            return false;
        }

        var match = Regex.Match(value, @"(?<!\d)(?<value>100|[1-9]?\d(?:\.\d+)?)\s*%(?!\d)");
        if (!match.Success ||
            !decimal.TryParse(match.Groups["value"].Value, NumberStyles.Number, CultureInfo.InvariantCulture, out var numeric) ||
            numeric is < 0m or > 100m)
        {
            return false;
        }

        percentage = $"{numeric:0.##}%";
        return true;
    }

    private static string ExtractHealthDeductible(IEnumerable<BenefitSummarySourceRow> sourceRows)
    {
        var rows = sourceRows.ToList();
        foreach (var value in rows.Select(row => row.Deductible).Where(value => !string.IsNullOrWhiteSpace(value)))
        {
            return NormalizeDeductible(value, false);
        }

        foreach (var row in rows.Where(row =>
                     Regex.IsMatch($"{row.Benefit} {row.Category} {row.ProvisionName}", @"\bdeductible\b", RegexOptions.IgnoreCase)))
        {
            if (!string.IsNullOrWhiteSpace(row.RawValue))
            {
                return NormalizeDeductible(row.RawValue, false);
            }
        }

        var evidence = BuildEvidence(rows);
        var calendarYearMatch = Regex.Match(
            evidence,
            @"\bCalendar Year Deductible\s+(?<value>.{1,100}?)(?=\s+(?:Deductible Does Not Apply To\b|%\s*Reimbursement\b|Allowable Expenses\b|Prescription Drugs?\b)|\r|\n|$)",
            RegexOptions.IgnoreCase);
        if (calendarYearMatch.Success)
        {
            return NormalizeDeductible(calendarYearMatch.Groups["value"].Value, true);
        }

        var labelledMatch = Regex.Match(
            evidence,
            @"\bDeductible\s*:\s*(?<value>[^\r\n]{1,100})",
            RegexOptions.IgnoreCase);
        return labelledMatch.Success ? NormalizeDeductible(labelledMatch.Groups["value"].Value, false) : "";
    }

    private static string NormalizeDeductible(string value, bool calendarYear)
    {
        var cleaned = CleanClientFacingLine(value);
        if (Regex.IsMatch(cleaned, @"^(?:no deductible|none|nil)$", RegexOptions.IgnoreCase))
        {
            return "None";
        }

        var formatted = FormatDisplayValue(cleaned, "Deductible");
        if (calendarYear &&
            !string.IsNullOrWhiteSpace(formatted) &&
            !formatted.Equals("None", StringComparison.OrdinalIgnoreCase) &&
            !Regex.IsMatch(formatted, @"\b(?:per|each)\s+(?:calendar\s+)?year\b|\bannual(?:ly)?\b", RegexOptions.IgnoreCase))
        {
            formatted = $"{formatted} per calendar year";
        }

        return formatted;
    }

    private static string ExtractPrescriptionDrugMaximum(IReadOnlyList<BenefitSummarySourceRow> rows)
    {
        const string moneyPattern = @"\$\s*\d[\d,]*(?:\.\d{1,2})?";
        var evidence = BuildEvidence(rows);
        var amountMatch = Regex.Match(
            evidence,
            $@"(?<amount>{moneyPattern})\s+(?:annual\s+)?(?:prescription\s+)?drugs?\s+(?:max(?:imum)?)\b|(?:annual\s+)?(?:prescription\s+)?drugs?\s+(?:max(?:imum)?)\s*[:\-]?\s*(?<amount>{moneyPattern})",
            RegexOptions.IgnoreCase);

        if (!amountMatch.Success)
        {
            foreach (var row in rows)
            {
                var maximum = FirstNonBlank(row.Maximum, row.RawValue);
                var rowEvidence = $"{row.Maximum}\n{row.RawValue}\n{row.SourceText}";
                var amounts = Regex.Matches(maximum, moneyPattern);
                if (amounts.Count == 1 &&
                    !Regex.IsMatch(rowEvidence, @"\bfertility\b|anti[- ]?smoking|smoking cessation", RegexOptions.IgnoreCase))
                {
                    amountMatch = amounts[0];
                    break;
                }
            }
        }

        if (!amountMatch.Success)
        {
            return Regex.IsMatch(evidence, @"\b(?:prescription\s+)?drugs?\s+(?:max(?:imum)?)\s*[:\-]?\s*Unlimited\b|\bMaximum\s*:\s*Unlimited\b", RegexOptions.IgnoreCase)
                ? "Unlimited"
                : "";
        }

        var amount = amountMatch.Groups["amount"].Success
            ? amountMatch.Groups["amount"].Value
            : amountMatch.Value;
        var formatted = FormatDisplayValue(amount, "Maximum");
        var isAnnual = Regex.IsMatch(
            evidence,
            @"\b(?:per\s+(?:calendar\s+)?year|annual(?:ly)?|calendar year)\b",
            RegexOptions.IgnoreCase);
        return isAnnual && !Regex.IsMatch(formatted, @"\bper\s+year\b", RegexOptions.IgnoreCase)
            ? $"{formatted} per year"
            : formatted;
    }

    private static string ExtractElectiveMedicationSummary(IReadOnlyList<BenefitSummarySourceRow> rows)
    {
        var evidence = BuildEvidence(rows);
        var values = new List<string>();
        var fertility = ExtractMedicationLimit(evidence, @"fertility(?:\s+drugs?)?", "fertility drugs");
        if (!string.IsNullOrWhiteSpace(fertility))
        {
            values.Add(fertility);
        }

        var antiSmoking = ExtractMedicationLimit(
            evidence,
            @"anti[- ]?smoking(?:\s+prescription\s+drugs?)?|smoking cessation(?:\s+(?:aids?|drugs?))?",
            "anti-smoking prescription drugs");
        if (!string.IsNullOrWhiteSpace(antiSmoking))
        {
            values.Add(antiSmoking);
        }

        return string.Join(Environment.NewLine, values);
    }

    private static string ExtractMedicationLimit(
        string evidence,
        string medicationPattern,
        string description)
    {
        const string moneyPattern = @"\$\s*\d[\d,]*(?:\.\d{1,2})?";
        var match = Regex.Match(
            evidence,
            $@"(?<amount>{moneyPattern})[^$;\r\n]{{0,60}}\b(?:{medicationPattern})\b(?<tail>[^$;\r\n]{{0,35}})|\b(?:{medicationPattern})\b[^$;\r\n]{{0,60}}(?<amount>{moneyPattern})(?<tail>[^$;\r\n]{{0,35}})",
            RegexOptions.IgnoreCase);
        if (!match.Success)
        {
            return "";
        }

        var amount = FormatDisplayValue(match.Groups["amount"].Value, "Maximum");
        if (Regex.IsMatch(match.Value, @"\blifetime\b", RegexOptions.IgnoreCase))
        {
            return $"{amount} lifetime maximum on {description}";
        }

        if (Regex.IsMatch(match.Value, @"\bper\s+(?:calendar\s+)?year\b|\bannual(?:ly)?\b", RegexOptions.IgnoreCase))
        {
            return $"{amount} per calendar year for {description}";
        }

        return $"{amount} for {description}";
    }

    private static bool HasConfirmedManulifeDrugCoverage(
        BenefitSummaryBuildRequest request,
        IReadOnlyList<BenefitSummarySourceRow> drugRows)
    {
        var specificCarriers = drugRows
            .Select(row => row.Carrier)
            .Where(carrier => !IsMissing(carrier))
            .Concat(request.CarrierPolicies
                .Where(policy => Regex.IsMatch(policy.Coverage ?? "", @"\b(?:prescription )?drugs?\b|extended health", RegexOptions.IgnoreCase))
                .Select(policy => policy.Carrier))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        var evidenceProvider = FindProvider(BuildEvidence(drugRows));
        if (!string.IsNullOrWhiteSpace(evidenceProvider))
        {
            specificCarriers.Add(evidenceProvider);
        }

        if (specificCarriers.Count > 0)
        {
            return specificCarriers.Any(IsManulifeName);
        }

        return IsManulifeName(request.Carrier);
    }

    private static bool IsManulifeName(string value)
    {
        return Regex.IsMatch(value ?? "", @"\bManulife(?: Financial)?\b", RegexOptions.IgnoreCase);
    }

    private static bool IsMedicalServicesRow(BenefitSummarySourceRow row)
    {
        var text = $"{row.Benefit} {row.Category} {row.ProvisionName}";
        return Regex.IsMatch(
            text,
            @"\b(?:other medical(?: expenses?)?|medical (?:services|supplies)(?:\s*&\s*| and )?(?:services|supplies)?|private duty nursing)\b",
            RegexOptions.IgnoreCase);
    }

    private static bool IsExtendedHealthDeductibleRow(BenefitSummarySourceRow row)
    {
        var text = $"{row.Benefit} {row.Category} {row.ProvisionName}";
        return Regex.IsMatch(text, @"\bdeductible\b", RegexOptions.IgnoreCase) &&
               Regex.IsMatch(text, @"extended health|\bEHC\b", RegexOptions.IgnoreCase);
    }

    private static string BuildEvidence(IEnumerable<BenefitSummarySourceRow> rows)
    {
        return string.Join(
            Environment.NewLine,
            rows.SelectMany(row => new[]
                {
                    row.RawValue,
                    row.Coinsurance,
                    row.Deductible,
                    row.Maximum,
                    row.Frequency,
                    row.Carrier,
                    row.SourceText
                })
                .Where(value => !string.IsNullOrWhiteSpace(value)));
    }

    private static void AddDerivedCostSharing(MutableSection section)
    {
        if (!SupportsCostSharing(section.Title) ||
            section.Lines.Any(line => line.Label.Equals("Cost Sharing", StringComparison.OrdinalIgnoreCase)))
        {
            return;
        }

        var coinsurance = section.Lines.FirstOrDefault(line => line.Label.Equals("Coinsurance", StringComparison.OrdinalIgnoreCase));
        var percentageMatches = Regex.Matches(coinsurance?.Value ?? "", @"(?<value>100|[1-9]?\d(?:\.\d+)?)\s*%");
        if (percentageMatches.Count != 1)
        {
            return;
        }

        var match = percentageMatches[0];
        if (!match.Success ||
            !decimal.TryParse(match.Groups["value"].Value, NumberStyles.Number, CultureInfo.InvariantCulture, out var planPercent) ||
            planPercent is < 0m or > 100m)
        {
            return;
        }

        section.Add("Cost Sharing", $"Plan {planPercent:0.##}% / Member {100m - planPercent:0.##}%", false, 0, 15);
    }

    private static BenefitSummarySectionViewModel FinalizeSection(MutableSection section)
    {
        var distinct = section.Lines
            .Where(line => !string.IsNullOrWhiteSpace(line.Value) || !string.IsNullOrWhiteSpace(line.Label))
            .GroupBy(line => $"{line.Label}\u001f{line.Value}\u001f{line.IsBullet}", StringComparer.OrdinalIgnoreCase)
            .Select(group => group.First())
            .OrderBy(line => line.Order)
            .ThenBy(line => line.Sequence)
            .Select(line => new BenefitSummaryLineViewModel
            {
                Label = line.Label,
                DisplayValue = line.Value,
                IndentLevel = line.IndentLevel,
                IsBullet = line.IsBullet,
                IsBold = !line.IsBullet && !string.IsNullOrWhiteSpace(line.Label),
                DisplayOrder = line.Order
            })
            .ToList();
        return new BenefitSummarySectionViewModel
        {
            SectionCode = Regex.Replace(section.Title.ToUpperInvariant(), @"[^A-Z0-9]+", "_").Trim('_'),
            SectionTitle = section.Title,
            DisplayOrder = SectionSortOrder(section.Title),
            Lines = distinct
        };
    }

    private static string ResolveSection(BenefitSummarySourceRow row, bool hasPrescriptionDrugEvidence = false)
    {
        var text = $"{row.Benefit} {row.Category} {row.ProvisionName}";
        if (Regex.IsMatch(text, @"\b(?:dep(?:endent|endant)?|dependant)\s+life\b", RegexOptions.IgnoreCase)) return "Dependent Life";
        if (text.Contains("optional life", StringComparison.OrdinalIgnoreCase)) return "Optional Life";
        if (Regex.IsMatch(text, @"\bSTD\b|short[ -]?term disability", RegexOptions.IgnoreCase)) return "Short-Term Disability";
        if (Regex.IsMatch(text, @"\bLTD\b|long[ -]?term disability", RegexOptions.IgnoreCase)) return "Long-Term Disability";
        if (text.Contains("critical illness", StringComparison.OrdinalIgnoreCase)) return "Critical Illness";
        if (Regex.IsMatch(text, @"\b(?:prescription )?drugs?\b|drug plan|drug max", RegexOptions.IgnoreCase)) return "Prescription Drugs";
        if (Regex.IsMatch(text, @"\b(?:other medical(?: expenses?)?|medical (?:services|supplies)(?:\s*&\s*| and )?(?:services|supplies)?)\b", RegexOptions.IgnoreCase)) return "Medical Services & Supplies";
        if (Regex.IsMatch(text, @"mental health|psychological services", RegexOptions.IgnoreCase)) return "Mental Health Practitioners";
        if (Regex.IsMatch(text, @"\bpara\s*medical", RegexOptions.IgnoreCase) || text.Contains("practitioner", StringComparison.OrdinalIgnoreCase)) return "Paramedicals";
        if (text.Contains("hospital", StringComparison.OrdinalIgnoreCase)) return "Hospital";
        if (text.Contains("eye exam", StringComparison.OrdinalIgnoreCase)) return "Eye Exams";
        if (text.Contains("vision", StringComparison.OrdinalIgnoreCase)) return "Vision Care";
        if (text.Contains("travel", StringComparison.OrdinalIgnoreCase) || text.Contains("out of province", StringComparison.OrdinalIgnoreCase)) return "Emergency Travel";
        if (text.Contains("dental", StringComparison.OrdinalIgnoreCase) || text.Contains("orthodont", StringComparison.OrdinalIgnoreCase)) return "Dental Care";
        if (Regex.IsMatch(text, @"\b(?:HCSA|HSCA|HSA)\b|health(?: care)? spending account", RegexOptions.IgnoreCase)) return "Health Spending Account";
        if (Regex.IsMatch(text, @"\b(?:LSA|WSA)\b|lifestyle spending account|wellness(?: spending)? account", RegexOptions.IgnoreCase)) return "Wellness Spending Account";
        if (text.Contains("National HealthClaim", StringComparison.OrdinalIgnoreCase) && Regex.IsMatch(row.RawValue, @"\$\s*\d")) return "Health Spending Account";
        if (text.Contains("survivor benefit", StringComparison.OrdinalIgnoreCase)) return "Health & Dental Benefit Period";
        if (Regex.IsMatch(text, @"\bEAP\b|\bEFAP\b|employee assistance", RegexOptions.IgnoreCase)) return "Employee Assistance Program";
        if (Regex.IsMatch(text, @"medical second opinion|Teladoc Medical Experts", RegexOptions.IgnoreCase)) return "Medical Second Opinion";
        if (text.Contains("virtual", StringComparison.OrdinalIgnoreCase)) return "Virtual Health Care";
        if (text.Contains("digital wellness", StringComparison.OrdinalIgnoreCase)) return "Digital Wellness Program";
        if (text.Contains("eligibility", StringComparison.OrdinalIgnoreCase) || text.Contains("hours worked", StringComparison.OrdinalIgnoreCase)) return "Eligibility Requirements";
        if (text.Contains("medical supplies", StringComparison.OrdinalIgnoreCase) || text.Contains("private duty nursing", StringComparison.OrdinalIgnoreCase)) return "Medical Services & Supplies";
        if (Regex.IsMatch(text, @"\bAD\s*&\s*D\b|accidental death|\blife\b", RegexOptions.IgnoreCase)) return "Life / AD&D";
        if (text.Contains("extended health", StringComparison.OrdinalIgnoreCase) || Regex.IsMatch(text, @"\bEHC\b", RegexOptions.IgnoreCase))
        {
            return hasPrescriptionDrugEvidence && row.ProvisionName.Contains("Deductible", StringComparison.OrdinalIgnoreCase)
                ? "Prescription Drugs"
                : "Medical Services & Supplies";
        }

        return NormalizeBenefitName(FirstNonBlank(row.Benefit, row.Category));
    }

    private static string InferLabel(string sectionTitle, BenefitSummarySourceRow row, string rawValue)
    {
        var label = NormalizeLabel(FirstNonBlank(row.ProvisionName, row.Benefit));
        var benefitText = $"{row.Benefit} {row.ProvisionName}";
        if (sectionTitle.Equals("Life / AD&D", StringComparison.OrdinalIgnoreCase))
        {
            if (IsCombinedLifeAndAdAndD(row)) return label;
            if (IsAdAndD(row)) return "AD&D";
            if (Regex.IsMatch(benefitText, @"\b(?:employee )?(?:basic )?life\b", RegexOptions.IgnoreCase)) return "Employee Basic Life";
        }

        if (sectionTitle.Equals("Dependent Life", StringComparison.OrdinalIgnoreCase) && label.Contains("Dependent Life", StringComparison.OrdinalIgnoreCase)) return "Benefit Amount";
        if (sectionTitle.Equals("Critical Illness", StringComparison.OrdinalIgnoreCase) && label.Equals("Critical Illness", StringComparison.OrdinalIgnoreCase) && rawValue.Contains('$')) return "Benefit Amount";
        if (sectionTitle.Equals("Dental Care", StringComparison.OrdinalIgnoreCase) &&
            Regex.IsMatch(benefitText, @"\b(?:dental\s+)?recall\b", RegexOptions.IgnoreCase)) return "Recall";
        if (sectionTitle.Equals("Prescription Drugs", StringComparison.OrdinalIgnoreCase) && label.Equals("Prescription Drugs", StringComparison.OrdinalIgnoreCase) && rawValue.Contains('%')) return "Coinsurance";
        if (sectionTitle.Equals("Medical Services & Supplies", StringComparison.OrdinalIgnoreCase) &&
            Regex.IsMatch(label, @"\bother medical(?: expenses?)?\b", RegexOptions.IgnoreCase) &&
            rawValue.Contains('%')) return "Coinsurance";
        if ((sectionTitle.Contains("Spending Account", StringComparison.OrdinalIgnoreCase) || sectionTitle.Contains("Wellness Spending", StringComparison.OrdinalIgnoreCase)) &&
            (label.Equals(sectionTitle, StringComparison.OrdinalIgnoreCase) || label.Equals("Amount", StringComparison.OrdinalIgnoreCase))) return "Benefit Amount";
        if (sectionTitle.Equals("Employee Assistance Program", StringComparison.OrdinalIgnoreCase) ||
            sectionTitle.Equals("Virtual Health Care", StringComparison.OrdinalIgnoreCase) ||
            sectionTitle.Equals("Medical Second Opinion", StringComparison.OrdinalIgnoreCase) ||
            sectionTitle.Equals("Digital Wellness Program", StringComparison.OrdinalIgnoreCase)) return label.Equals(sectionTitle, StringComparison.OrdinalIgnoreCase) ? "Coverage" : label;
        return label;
    }

    private static string ResolveSupplementalCoinsuranceLabel(
        string sectionTitle,
        BenefitSummarySourceRow row)
    {
        if (!sectionTitle.Equals("Dental Care", StringComparison.OrdinalIgnoreCase))
        {
            return "Coinsurance";
        }

        var label = NormalizeLabel(FirstNonBlank(row.ProvisionName, row.Benefit, row.Category));
        return label is "Basic Services" or "Preventive Services" or "Endodontic / Periodontic Services" or
            "Routine Dental Care" or "Dentures" or "Major Restorative Services" or "Orthodontic Services" or
            "Accidental Dental Injury" or "Other Dental Treatment"
            ? label
            : "Coinsurance";
    }

    private static string NormalizeDentalMaximumForDisplay(string sectionTitle, string value)
    {
        return sectionTitle.Equals("Dental Care", StringComparison.OrdinalIgnoreCase)
            ? Regex.Replace(value, @"\bInsured Person\b", "person", RegexOptions.IgnoreCase)
            : value;
    }

    private static bool IsDentalRecallRow(string sectionTitle, BenefitSummarySourceRow row)
    {
        return sectionTitle.Equals("Dental Care", StringComparison.OrdinalIgnoreCase) &&
               Regex.IsMatch(
                   $"{row.Benefit} {row.Category} {row.ProvisionName}",
                   @"\b(?:dental\s+)?recall\b",
                   RegexOptions.IgnoreCase);
    }

    private static (string Label, string Value) SplitLabelledLine(string line, string fallbackLabel)
    {
        var match = Regex.Match(line, @"^(?<label>[A-Za-z][A-Za-z0-9 &/()'\-]{1,80}):\s*(?<value>.+)$");
        return match.Success
            ? (match.Groups["label"].Value, match.Groups["value"].Value)
            : (fallbackLabel, line);
    }

    private static void AddSupplementalLine(
        MutableSection section,
        string sectionTitle,
        string label,
        string value)
    {
        if (!TryPrepareSummaryLine(sectionTitle, label, value, out var normalizedLabel, out var displayValue) ||
            section.Lines.Any(line => line.Label.Equals(normalizedLabel, StringComparison.OrdinalIgnoreCase) &&
                                      line.Value.Equals(displayValue, StringComparison.OrdinalIgnoreCase)))
        {
            return;
        }

        section.Add(normalizedLabel, displayValue, false, 0, GetLineOrder(sectionTitle, normalizedLabel));
    }

    private static bool TryPrepareSummaryLine(
        string sectionTitle,
        string label,
        string value,
        out string normalizedLabel,
        out string displayValue)
    {
        normalizedLabel = NormalizeLabel(label);
        displayValue = "";
        var cleaned = CleanClientFacingLine(value);
        var labelForChecks = normalizedLabel;
        if (TryFormatLongTermDisabilityBenefitFormula(sectionTitle, cleaned, out var longTermDisabilityFormula))
        {
            normalizedLabel = "Benefit Schedule";
            displayValue = longTermDisabilityFormula;
            return true;
        }

        if (string.IsNullOrWhiteSpace(normalizedLabel) ||
            string.IsNullOrWhiteSpace(cleaned) ||
            normalizedLabel.Equals("Cost Sharing", StringComparison.OrdinalIgnoreCase) ||
            NonSummaryLabels.Any(item => labelForChecks.Equals(item, StringComparison.OrdinalIgnoreCase)) ||
            IsPolicyNarrative(cleaned))
        {
            return false;
        }

        if (normalizedLabel.Contains("Coinsurance", StringComparison.OrdinalIgnoreCase) && !HasValidPercentages(cleaned))
        {
            return false;
        }

        if (IsMaximumLikeLabel(normalizedLabel) && HasAmbiguousAmountList(cleaned))
        {
            return false;
        }

        displayValue = FormatDisplayValue(cleaned, normalizedLabel);
        return !string.IsNullOrWhiteSpace(displayValue) && !IsPolicyNarrative(displayValue);
    }

    private static bool IsPolicyNarrative(string value)
    {
        var text = CleanWhitespace(value);
        if (text.Length > 180 || Regex.Matches(text, @"\b\w+\b").Count > 30)
        {
            return true;
        }

        return PolicyNarrativeTerms.Any(term => text.Contains(term, StringComparison.OrdinalIgnoreCase)) ||
               Regex.IsMatch(text, @"\.{4,}");
    }

    private static bool HasValidPercentages(string value)
    {
        if (Regex.IsMatch(value, @"\d{4,}\s*%"))
        {
            return false;
        }

        var matches = Regex.Matches(value, @"(?<!\d)(?<value>\d{1,3}(?:\.\d+)?)\s*%(?!\d)");
        return matches.Count > 0 && matches
            .Select(match => decimal.TryParse(match.Groups["value"].Value, NumberStyles.Number, CultureInfo.InvariantCulture, out var percentage)
                ? percentage
                : -1m)
            .All(percentage => percentage is >= 0m and <= 100m);
    }

    private static bool IsMaximumLikeLabel(string label)
    {
        return label.Contains("Maximum", StringComparison.OrdinalIgnoreCase) ||
               label.Contains("Benefit Amount", StringComparison.OrdinalIgnoreCase);
    }

    private static bool HasAmbiguousAmountList(string value)
    {
        var amounts = Regex.Matches(value, @"\$\s*\d[\d,]*(?:\.\d{1,2})?");
        if (amounts.Count < 3)
        {
            return false;
        }

        return !Regex.IsMatch(
            value,
            @"\b(?:employee|spouse|child|basic|major|orthodont|drug|fertility|smoking|per trip|per practitioner|mental health)\b",
            RegexOptions.IgnoreCase);
    }

    private static string FormatPractitionerCount(int count)
    {
        return count == 1 ? "1 practitioner type" : $"{count} practitioner types";
    }

    private static string ReconcileSameAsLife(string value)
    {
        var output = new List<string> { "Same as Employee Basic Life" };
        foreach (var line in value.Replace("\r\n", "\n", StringComparison.Ordinal).Replace('\r', '\n')
                     .Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            if (line.Contains("same as life", StringComparison.OrdinalIgnoreCase) ||
                line.StartsWith("Maximum:", StringComparison.OrdinalIgnoreCase) ||
                line.Contains("contract/booklet", StringComparison.OrdinalIgnoreCase) ||
                line.Contains("review against sold", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            output.Add(line);
        }

        return string.Join(Environment.NewLine, output);
    }

    private static string CleanClientFacingLine(string line)
    {
        if (Regex.IsMatch(line, @"^\s*(?:sold design|sold plan)\s+only\b", RegexOptions.IgnoreCase))
        {
            return "";
        }

        var cleaned = Regex.Replace(line, @"\s*(?:/|;)?\s*per\s+sold\s+(?:plan\s+)?design\b", "", RegexOptions.IgnoreCase);
        cleaned = Regex.Replace(cleaned, @";?\s*review\s+(?:the\s+)?booklet.*$", "", RegexOptions.IgnoreCase);
        cleaned = Regex.Replace(cleaned, @";?\s*review\s+against\s+.*$", "", RegexOptions.IgnoreCase);
        return CleanWhitespace(cleaned).Trim().TrimEnd(';', '/', '-', ',');
    }

    private static IReadOnlyList<string> ExtractPractitioners(string value)
    {
        var match = Regex.Match(value, @"Practitioners:\s*(?<list>.+?)(?:\n\s*(?:Maximum|Mental Health)\s*:|$)", RegexOptions.IgnoreCase | RegexOptions.Singleline);
        if (!match.Success)
        {
            return [];
        }

        return match.Groups["list"].Value
            .Split([';', '\n'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(NormalizePractitionerName)
            .Where(name => !string.IsNullOrWhiteSpace(name))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static string ExtractStandardParamedicalMaximum(BenefitSummarySourceRow row, string evidence)
    {
        var explicitMaximum = row.Maximum.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .FirstOrDefault(value => value.Contains('$') && !value.Contains("mental health", StringComparison.OrdinalIgnoreCase));
        if (!string.IsNullOrWhiteSpace(explicitMaximum))
        {
            return explicitMaximum;
        }

        var beforePractitioners = evidence.Split("Practitioners:", 2, StringSplitOptions.TrimEntries)[0];
        return Regex.Match(beforePractitioners, @"\$[\d,]+(?:\.\d{1,2})?(?:\s+(?:combined|per)\s+[^.;\n]+)?", RegexOptions.IgnoreCase).Value;
    }

    private static string ExtractMentalHealthMaximum(string evidence)
    {
        var match = Regex.Match(evidence, @"(?:mental health|psycholog(?:y|ist)?)[^$\n]{0,40}(?<amount>\$[\d,]+(?:\.\d{1,2})?)", RegexOptions.IgnoreCase);
        return match.Success ? match.Groups["amount"].Value : "";
    }

    private static bool IsMentalHealthPractitioner(string value)
    {
        return MentalHealthPractitionerTerms.Any(term => value.Contains(term, StringComparison.OrdinalIgnoreCase));
    }

    private static string FindProvider(string evidence)
    {
        foreach (var provider in KnownProviders)
        {
            if ((provider.Equals("National HealthClaim", StringComparison.OrdinalIgnoreCase) &&
                 Regex.IsMatch(evidence, @"\bNHC\b|National Health\s*Claim", RegexOptions.IgnoreCase)) ||
                Regex.IsMatch(evidence, $@"\b{Regex.Escape(provider)}\b", RegexOptions.IgnoreCase))
            {
                return NormalizeProviderName(provider);
            }
        }

        return "";
    }

    private static string NormalizeProviderName(string value)
    {
        if (value.Equals("Manulife Financial", StringComparison.OrdinalIgnoreCase)) return "Manulife";
        if (value.Equals("GreenShield", StringComparison.OrdinalIgnoreCase)) return "Green Shield";
        return CleanWhitespace(value);
    }

    private static string NormalizePolicy(string value)
    {
        var text = NormalizeMissingValue(value);
        return Regex.Replace(text, @"^policy\s*(?:number|no\.?|#)?\s*[:\-]?\s*", "", RegexOptions.IgnoreCase).Trim();
    }

    private static string JoinPolicyNumbers(IReadOnlyList<string> policies)
    {
        return policies.Count switch
        {
            0 => "",
            1 => policies[0],
            2 => $"{policies[0]} & {policies[1]}",
            _ => $"{string.Join(", ", policies.Take(policies.Count - 1))} & {policies[^1]}"
        };
    }

    private static string NormalizeUnit(string value)
    {
        var unit = CleanWhitespace(value);
        if (Regex.IsMatch(unit, @"^(?:monthly|per month)$", RegexOptions.IgnoreCase)) return "per month";
        if (Regex.IsMatch(unit, @"^(?:annual|annually|per year|calendar year)$", RegexOptions.IgnoreCase)) return "per calendar year";
        if (unit.Equals("lifetime", StringComparison.OrdinalIgnoreCase)) return "lifetime";
        return unit;
    }

    private static (string ClassName, string OptionName) SplitClassAndOption(string value)
    {
        var cleaned = Regex.Replace(CleanWhitespace(value), @"^Class\s+\d+\s*[-:]\s*", "", RegexOptions.IgnoreCase);
        var optionMatch = Regex.Match(cleaned, @"^(?<class>.+?)\s*[-|]\s*(?<option>(?:Flex|Option|Plan Option)\s*.+)$", RegexOptions.IgnoreCase);
        var className = optionMatch.Success ? optionMatch.Groups["class"].Value.Trim() : cleaned;
        var optionName = optionMatch.Success ? optionMatch.Groups["option"].Value.Trim() : "";
        if (Regex.IsMatch(className, @"^[A-Za-z0-9]{1,3}$") && !className.Equals("All", StringComparison.OrdinalIgnoreCase))
        {
            className = $"Class {className}";
        }

        return (FirstNonBlank(className, "All Classes"), optionName);
    }

    private static int GetLineOrder(string section, string label)
    {
        var text = label.ToLowerInvariant();
        if (text.Contains("provider")) return 0;
        if (section.Equals("Dental Care", StringComparison.OrdinalIgnoreCase))
        {
            if (text.Contains("basic")) return 10;
            if (text.Contains("major")) return 20;
            if (text.Contains("ortho")) return 30;
            if (text.Contains("deductible")) return 40;
            if (text.Equals("maximum", StringComparison.Ordinal) || text.Contains("annual maximum")) return 50;
            if (text.Contains("frequency") || text.Contains("recall")) return 60;
            if (text.Contains("fee guide")) return 70;
            return 200;
        }
        if (section.Equals("Life / AD&D", StringComparison.OrdinalIgnoreCase))
        {
            if (text.Contains("employee basic life") || text.Contains("benefit amount")) return 10;
            if (text.Contains("ad&d")) return 20;
            if (text.Contains("no-evidence") || text.Contains("non-evidence")) return 30;
            if (text.Contains("reduction")) return 40;
            if (text.Contains("waiver")) return 50;
        }
        if (section.Equals("Dependent Life", StringComparison.OrdinalIgnoreCase))
        {
            if (text.Contains("spouse")) return 10;
            if (text.Contains("dependent") || text.Contains("child")) return 20;
            if (text.Contains("termination")) return 30;
        }
        if (text.Contains("benefit amount") || text.Contains("benefit schedule")) return 10;
        if (text.Equals("coinsurance", StringComparison.Ordinal)) return 10;
        if (text.Contains("cost sharing")) return 15;
        if (text.Contains("deductible")) return 20;
        if (text.Equals("maximum", StringComparison.Ordinal) || text.Contains("annual maximum")) return 30;
        if (text.Contains("generic substitution")) return 35;
        if (text.Contains("elective medications")) return 40;
        if (text.Contains("personalized medicine")) return 50;
        if (text.Contains("non-evidence")) return 40;
        if (text.Contains("tax status")) return 50;
        if (text.Contains("qualification")) return 60;
        if (text.Contains("definition of disability")) return 70;
        if (text.Contains("benefit period")) return 80;
        if (text.Contains("termination")) return 90;
        if (text.Contains("frequency") || text.Contains("recall")) return 100;
        return 200;
    }

    private static int SectionSortOrder(string value)
    {
        var index = Array.FindIndex(SectionOrder, item => item.Equals(value, StringComparison.OrdinalIgnoreCase));
        return index < 0 ? 999 : index + 1;
    }

    private static bool SupportsProvider(string section)
    {
        return section.Equals("Critical Illness", StringComparison.OrdinalIgnoreCase) ||
               section.Equals("Health Spending Account", StringComparison.OrdinalIgnoreCase) ||
               section.Equals("Wellness Spending Account", StringComparison.OrdinalIgnoreCase) ||
               section.Equals("Employee Assistance Program", StringComparison.OrdinalIgnoreCase);
    }

    private static bool SupportsCostSharing(string section)
    {
        return section.Equals("Medical Services & Supplies", StringComparison.OrdinalIgnoreCase) ||
               section.Equals("Paramedicals", StringComparison.OrdinalIgnoreCase) ||
               section.Equals("Mental Health Practitioners", StringComparison.OrdinalIgnoreCase) ||
               section.Equals("Hospital", StringComparison.OrdinalIgnoreCase) ||
               section.Equals("Vision Care", StringComparison.OrdinalIgnoreCase) ||
               section.Equals("Eye Exams", StringComparison.OrdinalIgnoreCase) ||
               section.Equals("Emergency Travel", StringComparison.OrdinalIgnoreCase) ||
               section.Equals("Dental Care", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsCombinedLifeAndAdAndD(BenefitSummarySourceRow row)
    {
        return Regex.IsMatch(row.Benefit, @"^\s*Life\s*/\s*AD\s*&\s*D\s*$", RegexOptions.IgnoreCase);
    }

    private static bool IsAdAndD(BenefitSummarySourceRow row)
    {
        if (Regex.IsMatch(row.ProvisionName, @"\bAD\s*&\s*D\b|accidental death", RegexOptions.IgnoreCase))
        {
            return true;
        }

        return !IsCombinedLifeAndAdAndD(row) &&
               Regex.IsMatch(row.Benefit, @"\bAD\s*&\s*D\b|accidental death", RegexOptions.IgnoreCase);
    }

    private static bool ShouldSuppressRedundantCombinedLifeAmount(
        string sectionTitle,
        BenefitSummarySourceRow row,
        string rawValue)
    {
        return sectionTitle.Equals("Life / AD&D", StringComparison.OrdinalIgnoreCase) &&
               IsCombinedLifeAndAdAndD(row) &&
               Regex.IsMatch(row.ProvisionName, @"\bAD\s*&\s*D\b.*\bBenefit Amount\b", RegexOptions.IgnoreCase) &&
               Regex.IsMatch(rawValue, @"\bsame as (?:employee )?life (?:schedule|coverage|benefit)\b", RegexOptions.IgnoreCase);
    }

    private static bool ShouldHideCombinedLifeAmountLabel(
        string sectionTitle,
        BenefitSummarySourceRow row,
        string normalizedLabel)
    {
        return sectionTitle.Equals("Life / AD&D", StringComparison.OrdinalIgnoreCase) &&
               IsCombinedLifeAndAdAndD(row) &&
               normalizedLabel.Equals("Benefit Amount", StringComparison.OrdinalIgnoreCase);
    }

    private static bool ShouldSuppressEmbeddedCombinedLifeMaximum(
        string sectionTitle,
        BenefitSummarySourceRow row,
        string rawValue)
    {
        var maximum = CleanWhitespace(row.Maximum);
        return sectionTitle.Equals("Life / AD&D", StringComparison.OrdinalIgnoreCase) &&
               IsCombinedLifeAndAdAndD(row) &&
               NormalizeLabel(row.ProvisionName).Equals("Benefit Amount", StringComparison.OrdinalIgnoreCase) &&
               !string.IsNullOrWhiteSpace(maximum) &&
               CleanWhitespace(rawValue).Contains(maximum, StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsPrescriptionDrugRow(BenefitSummarySourceRow row)
    {
        return Regex.IsMatch($"{row.Benefit} {row.Category} {row.ProvisionName}", @"\b(?:prescription )?drugs?\b|drug plan|drug max", RegexOptions.IgnoreCase);
    }

    private static bool IsMoneyLabel(string label)
    {
        return label.Contains("Amount", StringComparison.OrdinalIgnoreCase) ||
               label.Contains("Maximum", StringComparison.OrdinalIgnoreCase) ||
               label.Contains("Deductible", StringComparison.OrdinalIgnoreCase);
    }

    private static bool HasMeaningfulValue(BenefitSummarySourceRow row)
    {
        return !string.IsNullOrWhiteSpace(FirstNonBlank(
            row.RawValue,
            row.Coinsurance,
            row.Deductible,
            row.Maximum,
            row.Frequency,
            row.WaitingPeriod,
            row.TerminationAge));
    }

    private static bool IsGenericClass(string value)
    {
        return string.IsNullOrWhiteSpace(value) ||
               value.Equals("All Classes", StringComparison.OrdinalIgnoreCase) ||
               value.Equals("All Class", StringComparison.OrdinalIgnoreCase) ||
               value.Equals("All Eligible", StringComparison.OrdinalIgnoreCase) ||
               value.Equals("Class 01", StringComparison.OrdinalIgnoreCase);
    }

    private static string NormalizeClassKey(string value)
    {
        return FirstNonBlank(CleanWhitespace(value), "All Classes");
    }

    private static string NormalizeMissingValue(string value)
    {
        return IsMissing(value) ? "" : CleanWhitespace(value);
    }

    private static bool IsMissing(string value)
    {
        return string.IsNullOrWhiteSpace(value) ||
               Regex.IsMatch(value.Trim(), @"^(?:n/?a|not (?:saved|available|provided|confirmed)|unknown|null|-)$", RegexOptions.IgnoreCase);
    }

    private static string CleanWhitespace(string value)
    {
        return Regex.Replace(value ?? "", @"[ \t]+", " ").Trim();
    }

    private static string FirstNonBlank(params string[] values)
    {
        return values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value))?.Trim() ?? "";
    }

    private static MutableSection GetSection(IDictionary<string, MutableSection> sections, string title)
    {
        if (!sections.TryGetValue(title, out var section))
        {
            section = new MutableSection(title);
            sections[title] = section;
        }

        return section;
    }

    private sealed class MutableSection(string title)
    {
        private int _sequence;

        public string Title { get; } = title;
        public List<MutableLine> Lines { get; } = [];

        public void Add(string label, string value, bool isBullet, int indentLevel, int order)
        {
            if (string.IsNullOrWhiteSpace(label) && string.IsNullOrWhiteSpace(value))
            {
                return;
            }

            Lines.Add(new MutableLine(label.Trim(), value.Trim(), isBullet, indentLevel, order, _sequence++));
        }
    }

    private sealed record MutableLine(
        string Label,
        string Value,
        bool IsBullet,
        int IndentLevel,
        int Order,
        int Sequence);
}

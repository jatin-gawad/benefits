using RenewalReviewBuilder.App.Models;

namespace RenewalReviewBuilder.App.Services.CanonicalClaims;

public sealed record CanadaLifeTermMapping(
    string RawValue,
    string CanonicalValue,
    CanonicalClaimsMappingClassification Classification,
    string RuleName)
{
    public bool IsMapped => Classification is CanonicalClaimsMappingClassification.Direct or CanonicalClaimsMappingClassification.Normalized;
}

/// <summary>Explicit Canada Life aliases. Unknown values are never guessed.</summary>
public static class CanadaLifeClaimsTerminology
{
    private static readonly IReadOnlyDictionary<string, string> Benefits = Dictionary(
        ("Extended Health Care", "EHC"), ("Healthcare", "EHC"), ("Health", "EHC"), ("Medical", "EHC"), ("EHC", "EHC"),
        ("Dental", "Dental"), ("Dental Care", "Dental"), ("DTL", "Dental"),
        ("Long Term Disability", "LTD"), ("Long-Term Disability", "LTD"), ("LTD", "LTD"),
        ("Short Term Disability", "STD"), ("Short-Term Disability", "STD"), ("STD", "STD"),
        ("Employee Term Life", "Basic Life"), ("Basic Life", "Basic Life"), ("Life", "Basic Life"),
        ("AD&D", "AD&D"), ("Accidental Death & Dismemberment", "AD&D"), ("Accidental Death and Dismemberment", "AD&D"),
        ("Dependent Life", "Dependent Life"), ("Critical Illness", "Critical Illness"), ("Employee Critical Illness", "Critical Illness"),
        ("Employee Assistance Program", "EAP"), ("EAP", "EAP"), ("Contact", "EAP"),
        ("Vision", "Vision"), ("Visioncare", "Vision"), ("Drugs", "Drugs"), ("Drug", "Drugs"),
        ("Virtual Health", "Virtual Health"), ("Virtual Healthcare", "Virtual Health"));

    private static readonly IReadOnlyDictionary<string, string> Categories = Dictionary(
        ("Total", "Total"), ("Drug", "Drugs"), ("Drugs", "Drugs"), ("Prescription Drugs", "Drugs"),
        ("Paramedical", "Paramedical"), ("Paramedical Services", "Paramedical"),
        ("Medical Supplies", "Medical Supplies"), ("Services and Supplies", "Medical Supplies"),
        ("Out of Country", "Out of Country"), ("Emergency Travel", "Out of Country"),
        ("Vision", "Vision"), ("Vision Care", "Vision"), ("Dental", "Dental"),
        ("Hospital", "Hospital"), ("Ambulance", "Ambulance"), ("Out of Province", "Out of Province"),
        ("Adjustment", "Adjustment"), ("Claims Adjustment", "Adjustment"), ("Reversal", "Reversal"),
        ("Other", "Other"), ("Miscellaneous", "Other"),
        ("Massage", "Massage"), ("Massage Therapy", "Massage"), ("Masseur", "Massage"), ("Registered Massage Therapist", "Massage"),
        ("Physiotherapy", "Physiotherapy"), ("Physiotherapist", "Physiotherapy"),
        ("Chiropractic", "Chiropractic"), ("Chiropractor", "Chiropractic"), ("Chiropractic Treatment", "Chiropractic"),
        ("Naturopath", "Naturopath"), ("Naturopath Treatment", "Naturopath"),
        ("Psychologist", "Psychologist"), ("Psychology", "Psychologist"), ("Psychologist's Office Visit", "Psychologist"),
        ("Psychotherapist", "Psychotherapist"), ("Clinical Counsellor", "Clinical Counsellor"),
        ("Social Worker", "Social Worker"), ("Acupuncture", "Acupuncture"), ("Osteopath", "Osteopath"),
        ("Podiatrist", "Podiatrist"), ("Chiropodist", "Chiropodist"), ("Dietitian", "Dietitian"),
        ("Dietician", "Dietitian"), ("Speech Therapy", "Speech Therapy"));

    private static readonly IReadOnlyDictionary<string, string> Funding = Dictionary(
        ("ASO", "Administrative Services Only"), ("Administrative Services Only", "Administrative Services Only"),
        ("Insured", "Insured"), ("Fully Insured", "Insured"), ("Refund Accounting", "Refund Accounting"),
        ("Retention Accounting", "Retention Accounting"), ("Non-Refund", "Non-Refund"),
        ("Pooled", "Pooled"), ("Non-Pooled", "Non-Pooled"), ("Large Amount Pooling", "Large Amount Pooling"));

    private static readonly IReadOnlyDictionary<string, string> Provinces = Dictionary(
        ("BC", "BC"), ("British Columbia", "BC"), ("AB", "AB"), ("Alberta", "AB"),
        ("SK", "SK"), ("Saskatchewan", "SK"), ("MB", "MB"), ("Manitoba", "MB"),
        ("ON", "ON"), ("Ontario", "ON"), ("QC", "QC"), ("Quebec", "QC"), ("Qu\u00e9bec", "QC"),
        ("NB", "NB"), ("New Brunswick", "NB"), ("NS", "NS"), ("Nova Scotia", "NS"),
        ("PE", "PE"), ("Prince Edward Island", "PE"), ("NL", "NL"), ("Newfoundland and Labrador", "NL"),
        ("YT", "YT"), ("Yukon", "YT"), ("NT", "NT"), ("Northwest Territories", "NT"),
        ("NU", "NU"), ("Nunavut", "NU"));

    private static readonly IReadOnlyDictionary<string, string> MembershipUnits = Dictionary(
        ("Employee", "Employee"), ("Employees", "Employee"), ("Certificate", "Certificate"), ("Certificates", "Certificate"),
        ("Insured Life", "Insured Life"), ("Insured Lives", "Insured Life"), ("Member", "Member"), ("Members", "Member"),
        ("Dependant", "Dependant"), ("Dependants", "Dependant"), ("Dependent", "Dependant"), ("Dependents", "Dependant"),
        ("Single", "Single"), ("Family", "Family"), ("Couple", "Couple"),
        ("Parent and Child", "Parent and Child"), ("Employee Only", "Single"),
        ("Employee + Spouse", "Couple"), ("Employee + Child", "Parent and Child"), ("Employee + Children", "Parent and Child"));

    public static CanadaLifeTermMapping NormalizeBenefit(string value) => Map(value, Benefits, "CanadaLife.BenefitAliases");
    public static CanadaLifeTermMapping NormalizeCategory(string value) => Map(value, Categories, "CanadaLife.CategoryAliases");
    public static CanadaLifeTermMapping NormalizeFunding(string value) => Map(value, Funding, "CanadaLife.FundingAliases");
    public static CanadaLifeTermMapping NormalizeProvince(string value) => Map(value, Provinces, "CanadaLife.ProvinceAliases");
    public static CanadaLifeTermMapping NormalizeMembershipUnit(string value) => Map(value, MembershipUnits, "CanadaLife.MembershipAliases");

    public static CanadaLifeTermMapping NormalizeDimensionIdentifier(string value, string dimension)
    {
        var raw = value.Trim();
        if (raw.Length == 0) return new(raw, "", CanonicalClaimsMappingClassification.Unavailable, $"CanadaLife.{dimension}.Unavailable");
        if (raw.Equals("ALL", StringComparison.OrdinalIgnoreCase)) return new(raw, "All", CanonicalClaimsMappingClassification.Normalized, $"CanadaLife.{dimension}.All");
        return raw.All(character => char.IsLetterOrDigit(character) || character is '-' or '_' or ' ')
            ? new(raw, raw, CanonicalClaimsMappingClassification.Direct, $"CanadaLife.{dimension}.Direct")
            : new(raw, "", CanonicalClaimsMappingClassification.Ambiguous, $"CanadaLife.{dimension}.Ambiguous");
    }

    private static CanadaLifeTermMapping Map(string value, IReadOnlyDictionary<string, string> mappings, string rule)
    {
        var raw = value.Trim();
        if (raw.Length == 0) return new(raw, "", CanonicalClaimsMappingClassification.Unavailable, $"{rule}.Unavailable");
        if (!mappings.TryGetValue(raw, out var canonical)) return new(raw, "", CanonicalClaimsMappingClassification.Ambiguous, $"{rule}.Unmapped");
        return new(raw, canonical, raw.Equals(canonical, StringComparison.OrdinalIgnoreCase) ? CanonicalClaimsMappingClassification.Direct : CanonicalClaimsMappingClassification.Normalized, rule);
    }

    private static IReadOnlyDictionary<string, string> Dictionary(params (string Raw, string Canonical)[] values) =>
        values.ToDictionary(value => value.Raw, value => value.Canonical, StringComparer.OrdinalIgnoreCase);
}

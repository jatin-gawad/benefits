using System.Globalization;

namespace RenewalReviewBuilder.App.Models;

public sealed class BenefitSummaryBuildRequest
{
    public string ClientName { get; init; } = "";
    public string Carrier { get; init; } = "";
    public string PolicyNumber { get; init; } = "";
    public string EffectiveDate { get; init; } = "";
    public DateTime? UpdatedDate { get; init; }
    public string UpdatedDateSource { get; init; } = "";
    public IReadOnlyList<BenefitSummarySourceRow> Rows { get; init; } = [];
    public IReadOnlyList<BenefitSummaryCarrierPolicy> CarrierPolicies { get; init; } = [];
}

public sealed class BenefitSummarySourceRow
{
    public long? PlanDesignFeatureId { get; init; }
    public string ClassName { get; init; } = "";
    public string Benefit { get; init; } = "";
    public string Category { get; init; } = "";
    public string ProvisionName { get; init; } = "";
    public string RawValue { get; init; } = "";
    public string Coinsurance { get; init; } = "";
    public string Deductible { get; init; } = "";
    public string Maximum { get; init; } = "";
    public string Frequency { get; init; } = "";
    public string WaitingPeriod { get; init; } = "";
    public string TerminationAge { get; init; } = "";
    public string Carrier { get; init; } = "";
    public string PolicyNumber { get; init; } = "";
    public string SourceText { get; init; } = "";
    public string SourceDocument { get; init; } = "";
    public int? SourcePage { get; init; }
    public IReadOnlyList<PlanBenefitRule> Rules { get; set; } = [];
}

public sealed class BenefitSummaryCarrierPolicy
{
    public string Carrier { get; init; } = "";
    public string PolicyNumber { get; init; } = "";
    public string Coverage { get; init; } = "";
}

public sealed class BenefitSummaryViewModel
{
    public string ClientName { get; init; } = "";
    public string SummaryTitle { get; init; } = "Employee Benefits Summary";
    public string ClassName { get; init; } = "";
    public string OptionName { get; init; } = "";
    public string EffectiveDate { get; init; } = "";
    public DateTime? UpdatedDate { get; init; }
    public string UpdatedDateSource { get; init; } = "";
    public string DisclaimerText { get; init; } = "";
    public IReadOnlyList<BenefitSummarySectionViewModel> Sections { get; init; } = [];

    public string SelectionLabel => string.Join(" - ", new[] { ClassName, OptionName }
        .Where(value => !string.IsNullOrWhiteSpace(value)));

    public string EffectiveDateText => string.IsNullOrWhiteSpace(EffectiveDate)
        ? ""
        : $"Effective Date: {EffectiveDate}";

    public string UpdatedDateText => UpdatedDate.HasValue
        ? $"UPDATED: {UpdatedDate.Value.ToString("MMMM d, yyyy", CultureInfo.InvariantCulture).ToUpperInvariant()}"
        : "";
}

public sealed class BenefitSummarySectionViewModel
{
    public string SectionCode { get; init; } = "";
    public string SectionTitle { get; init; } = "";
    public int DisplayOrder { get; init; }
    public IReadOnlyList<BenefitSummaryLineViewModel> Lines { get; init; } = [];

    public string PlainText => string.Join(Environment.NewLine, Lines.Select(line => line.PlainText));
}

public sealed class BenefitSummaryLineViewModel
{
    public string Label { get; init; } = "";
    public string DisplayValue { get; init; } = "";
    public int IndentLevel { get; init; }
    public bool IsBullet { get; init; }
    public bool IsBold { get; init; }
    public int DisplayOrder { get; init; }

    public string LabelSeparator => string.IsNullOrWhiteSpace(Label) ? "" : ": ";
    public string BulletMarker => IsBullet ? "\u2022" : "";
    public string PlainText => $"{(IsBullet ? "- " : "")}{Label}{LabelSeparator}{DisplayValue}".TrimEnd();
}

public enum BenefitSummaryPercentageScale
{
    Decimal,
    Whole
}

using RenewalReviewBuilder.App.Services;

namespace RenewalReviewBuilder.App.Models;

public sealed class AnalysisResult
{
    private RenewalCase _renewalCase = new();

    public RenewalCase RenewalCase
    {
        get => _renewalCase;
        set
        {
            _renewalCase = value ?? new RenewalCase();
            ExportInputFingerprint = RenewalCaseInputFingerprint.Create(_renewalCase);
        }
    }

    public string ExportInputFingerprint { get; private set; } =
        RenewalCaseInputFingerprint.Create(new RenewalCase());
    public RenewalDocumentDiscoveryResult? DocumentDiscovery { get; set; }
    public PdfExtractionResult Pdf { get; set; } = new();
    public WorkbookExtractionResult Workbook { get; set; } = new();
    public List<PreviousRenewalSnapshot> PreviousRenewals { get; } = [];
    public List<Finding> Findings { get; } = [];
    public List<SourceCheck> SourceChecks { get; } = [];
    public List<RenewalVersion> RenewalVersions { get; } = [];
    public List<NemMaximumRow> NemMaximums { get; } = [];
    public List<AsoBillingActivityRecord> AsoBillingActivity { get; } = [];
    public List<AsoFinancialActivityStatement> AsoFinancialActivityStatements { get; } = [];
    public AssurePlusExperienceSummary? AssurePlusExperience { get; set; }
    public ClaimsRangeComparison? ClaimsRange { get; set; }
    public string SelectedRenewalVersionName { get; set; } = RenewalVersionNames.CarrierProposed;
    public DateTime CreatedAt { get; set; } = DateTime.Now;
    public long? DatabaseClientId { get; set; }
    public int? YearsAsClient { get; set; }

    public bool MatchesExportInputs(RenewalCase renewalCase) =>
        ExportInputFingerprint.Equals(
            RenewalCaseInputFingerprint.Create(renewalCase),
            StringComparison.Ordinal);

    public RenewalVersion? SelectedRenewalVersion => RenewalVersions.FirstOrDefault(version =>
        version.VersionName.Equals(SelectedRenewalVersionName, StringComparison.OrdinalIgnoreCase));
    public PreviousRenewalSnapshot? MostRelevantPreviousRenewal => PreviousRenewals
        .OrderByDescending(snapshot => snapshot.RenewalDate ?? DateOnly.MinValue)
        .FirstOrDefault();

    public decimal? CurrentMonthlyPremium => SelectedRenewalVersion?.CurrentMonthlyPremium ??
                                             (HasRateLines ? Pdf.RateLines.Sum(line => line.CurrentPremium ?? 0m) : Pdf.CurrentMonthlyPremium);
    public decimal? RenewalMonthlyPremium => SelectedRenewalVersion?.RenewalMonthlyPremium ??
                                             (HasRateLines ? Pdf.RateLines.Sum(line => line.RenewalPremium ?? 0m) : Pdf.RenewalMonthlyPremium);
    public decimal? MonthlyPremiumChange => CurrentMonthlyPremium.HasValue && RenewalMonthlyPremium.HasValue
        ? RenewalMonthlyPremium.Value - CurrentMonthlyPremium.Value
        : Pdf.MonthlyPremiumChange;
    public decimal? MonthlyPremiumChangePercent => CurrentMonthlyPremium is > 0m && RenewalMonthlyPremium.HasValue
        ? UnderwritingEngine.CalculatePercentChange(CurrentMonthlyPremium, RenewalMonthlyPremium)
        : Pdf.MonthlyPremiumChangePercent;

    private bool HasRateLines => Pdf.RateLines.Count > 0;
}

public sealed class PreviousRenewalSnapshot
{
    public string SourcePath { get; set; } = "";
    public string SourceFileName { get; set; } = "";
    public string SourceType { get; set; } = "";
    public string Carrier { get; set; } = "";
    public string ClientName { get; set; } = "";
    public string PolicyNumber { get; set; } = "";
    public DateOnly? RenewalDate { get; set; }
    public decimal? CurrentMonthlyPremium { get; set; }
    public decimal? RenewalMonthlyPremium { get; set; }
    public decimal? MonthlyPremiumChange { get; set; }
    public decimal? MonthlyPremiumChangePercent { get; set; }
    public decimal? CurrentAnnualPremium => CurrentMonthlyPremium.HasValue ? CurrentMonthlyPremium.Value * 12m : null;
    public decimal? RenewalAnnualPremium => RenewalMonthlyPremium.HasValue ? RenewalMonthlyPremium.Value * 12m : null;
    public string SourceReference { get; set; } = "";

    public bool HasUsefulValues =>
        RenewalDate.HasValue ||
        CurrentMonthlyPremium.HasValue ||
        RenewalMonthlyPremium.HasValue ||
        MonthlyPremiumChange.HasValue ||
        MonthlyPremiumChangePercent.HasValue ||
        !string.IsNullOrWhiteSpace(ClientName) ||
        !string.IsNullOrWhiteSpace(PolicyNumber);
}

public sealed class ExportResult
{
    public string OutputFolder { get; set; } = "";
    public string? WorkbookPath { get; set; }
    public string? PowerPointPath { get; set; }
    public string? DeckMessage { get; set; }
    public string? SupportingPdfFolder { get; set; }
    public int SupportingPdfCount { get; set; }
    public string? SupportingWorkbookFolder { get; set; }
    public int SupportingWorkbookCount { get; set; }
}

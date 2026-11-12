namespace RenewalReviewBuilder.App.Models;

public sealed class ClaimsExperienceDataModelRequest
{
    public string SourceFolder { get; set; } = "";
    public List<string> SourceFiles { get; set; } = [];
    public string OutputFolder { get; set; } = "";
    // Zero means detect the inclusive month count from the experience report.
    public int ReportMonthCount { get; set; }
    public string BookletMasterWorkbookPath { get; set; } = "";
    public List<string> BenefitBookletPaths { get; set; } = [];
    public ClaimsExperienceRenewalTerms RenewalTerms { get; set; } = new();
}

public sealed class ClaimsExperienceRenewalTerms
{
    public const decimal DefaultEhcPoolingCharge = 0.09m;
    public const decimal DefaultTrendFactor = 0.15m;
    public const decimal DefaultTargetLossRatio = 0.763m;

    public decimal? EhcPoolingCharge { get; init; }
    public decimal? EhcTrendFactor { get; init; }
    public decimal? DentalTrendFactor { get; init; }
    public decimal? EhcTargetLossRatio { get; init; }
    public decimal? DentalTargetLossRatio { get; init; }

    public decimal EffectiveEhcPoolingCharge => EhcPoolingCharge ?? DefaultEhcPoolingCharge;
    public decimal EffectiveEhcTrendFactor => EhcTrendFactor ?? DefaultTrendFactor;
    public decimal EffectiveDentalTrendFactor => DentalTrendFactor ?? DefaultTrendFactor;
    public decimal EffectiveEhcTargetLossRatio => EhcTargetLossRatio ?? DefaultTargetLossRatio;
    public decimal EffectiveDentalTargetLossRatio => DentalTargetLossRatio ?? DefaultTargetLossRatio;

    public static ClaimsExperienceRenewalTerms FromRenewalAssumptions(
        IEnumerable<RenewalAssumptionRecord> assumptions,
        string? carrier)
    {
        var source = assumptions.ToList();
        if (!string.IsNullOrWhiteSpace(carrier))
        {
            var carrierRows = source
                .Where(row => row.Carrier.Equals(carrier, StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (carrierRows.Count > 0)
            {
                source = carrierRows;
            }
        }

        return new ClaimsExperienceRenewalTerms
        {
            EhcPoolingCharge = Resolve(source, "EHC", row => row.PoolingCharge),
            EhcTrendFactor = Resolve(source, "EHC", row => row.Trend),
            DentalTrendFactor = Resolve(source, "Dental", row => row.Trend),
            EhcTargetLossRatio = Resolve(source, "EHC", row => row.TargetLossRatio),
            DentalTargetLossRatio = Resolve(source, "Dental", row => row.TargetLossRatio)
        };
    }

    private static decimal? Resolve(
        IEnumerable<RenewalAssumptionRecord> assumptions,
        string coverage,
        Func<RenewalAssumptionRecord, decimal?> selector)
    {
        return assumptions
            .Select(row => new
            {
                Row = row,
                Priority = BenefitPriority(row.BenefitLine, coverage),
                Value = selector(row)
            })
            .Where(item => item.Priority >= 0 && item.Value.HasValue)
            .OrderByDescending(item => item.Row.EffectiveDate ?? DateOnly.MinValue)
            .ThenBy(item => item.Priority)
            .Select(item => NormalizeRate(item.Value))
            .FirstOrDefault();
    }

    private static int BenefitPriority(string benefitLine, string coverage)
    {
        if (benefitLine.Equals("All", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }

        var isDental = benefitLine.Contains("dental", StringComparison.OrdinalIgnoreCase);
        if (coverage.Equals("Dental", StringComparison.OrdinalIgnoreCase))
        {
            return isDental ? 0 : -1;
        }

        var isEhc = !isDental &&
                    (benefitLine.Contains("ehc", StringComparison.OrdinalIgnoreCase) ||
                     benefitLine.Contains("health", StringComparison.OrdinalIgnoreCase) ||
                     benefitLine.Contains("extended", StringComparison.OrdinalIgnoreCase));
        return isEhc ? 0 : -1;
    }

    private static decimal? NormalizeRate(decimal? value)
    {
        return value.HasValue && Math.Abs(value.Value) > 1m
            ? value.Value / 100m
            : value;
    }
}

public sealed class ClaimsExperienceDataModelResult
{
    public string OutputFolder { get; set; } = "";
    public string WorkbookPath { get; set; } = "";
    public string ClientName { get; set; } = "";
    public string ClientNumber { get; set; } = "";
    public string Carrier { get; set; } = "";
    public int ReportMonthCount { get; set; }
    public DateOnly? ExperiencePeriodStart { get; set; }
    public DateOnly? ExperiencePeriodEnd { get; set; }
    public int TableCount { get; set; }
    public int SourceRowCount { get; set; }
    public string ScheduleOfBenefitsSourcePath { get; set; } = "";
    public List<string> SourceDocumentPaths { get; } = [];
    public List<ClaimsExperienceSummary> Summaries { get; } = [];
    public List<MonthlyClaimsRecord> MonthlyClaims { get; } = [];
    public List<MonthlyClaimsRecord> AllPeriodMonthlyClaims { get; } = [];
    public List<ClaimsSourcePeriod> SourcePeriods { get; } = [];
    public List<ClaimsServiceAdjudicationFact> ServiceAdjudicationFacts { get; } = [];
    public List<ClaimsServiceControlTotal> ServiceControlTotals { get; } = [];
    public List<string> Warnings { get; } = [];
}

public sealed class ClaimsExperienceSummary
{
    public string BenefitLine { get; set; } = "";
    public decimal? Premium { get; set; }
    public decimal? PaidClaims { get; set; }
    public decimal? IncurredClaims { get; set; }
    public decimal? LossRatio { get; set; }
    public int? Lives { get; set; }
}

public sealed class ClaimsSourcePeriod
{
    public string SourceRole { get; set; } = "";
    public DateOnly Start { get; set; }
    public DateOnly End { get; set; }
    public int MonthCount { get; set; }
    public string SourceFile { get; set; } = "";
}

public sealed class ClaimsServiceAdjudicationFact
{
    public string Benefit { get; set; } = "";
    public string Relationship { get; set; } = "";
    public string Category { get; set; } = "";
    public string Service { get; set; } = "";
    public DateOnly PeriodStart { get; set; }
    public DateOnly PeriodEnd { get; set; }
    public decimal AmountSubmitted { get; set; }
    public decimal AmountEligible { get; set; }
    public decimal AmountPaid { get; set; }
    public decimal ReportedPaidPercentage { get; set; }
    public string Currency { get; set; } = "CAD";
    public string SourceFile { get; set; } = "";
    public int SourcePage { get; set; }
    public string SourceEvidence { get; set; } = "";
    public decimal Confidence { get; set; } = 0.95m;
    public string ValidationState { get; set; } = "Valid";
}

public sealed class ClaimsServiceControlTotal
{
    public string Benefit { get; set; } = "";
    public string Relationship { get; set; } = "";
    public string Category { get; set; } = "";
    public string Service { get; set; } = "";
    public string TotalType { get; set; } = "";
    public DateOnly PeriodStart { get; set; }
    public DateOnly PeriodEnd { get; set; }
    public decimal AmountSubmitted { get; set; }
    public decimal AmountEligible { get; set; }
    public decimal AmountPaid { get; set; }
    public decimal ReportedPaidPercentage { get; set; }
    public string Currency { get; set; } = "CAD";
    public string SourceFile { get; set; } = "";
    public int SourcePage { get; set; }
    public string SourceEvidence { get; set; } = "";
}

public sealed class ClaimsExperiencePowerPointRequest
{
    public string WorkbookPath { get; set; } = "";
    public string TemplatePath { get; set; } = "";
    public string OutputFolder { get; set; } = "";
}

public sealed class ClaimsExperiencePowerPointResult
{
    public string OutputFolder { get; set; } = "";
    public string PowerPointPath { get; set; } = "";
    public string SourceWorkbookPath { get; set; } = "";
    public int SlideCountUpdated { get; set; }
    public List<string> Warnings { get; } = [];
}


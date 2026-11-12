using System.Globalization;

namespace RenewalReviewBuilder.App.Services;

public sealed class ClientWorkspaceRenewalHistoryPresentationService
{
    internal const int MaximumCycles = 5;
    internal const double LossChartWidth = 435d;
    internal const double LossChartHeight = 104d;
    private const double LossChartPadding = 8d;
    private const double MaximumPremiumBarHeight = 92d;

    public ClientWorkspaceRenewalHistoryPresentation Build(
        IEnumerable<ClientWorkspaceRenewalHistoryEntry> source,
        string clientStatus)
    {
        ArgumentNullException.ThrowIfNull(source);

        var cycles = source
            .Where(HasRenewalIdentity)
            .GroupBy(CycleKey, StringComparer.OrdinalIgnoreCase)
            .Select(group => group
                .OrderByDescending(item => item.DateLoaded)
                .First())
            .OrderBy(CycleDate)
            .ThenBy(item => item.DateLoaded)
            .TakeLast(MaximumCycles)
            .ToList();

        var current = cycles.LastOrDefault();
        var prior = cycles.Count > 1 ? cycles[^2] : null;
        var premiumMaximum = cycles
            .SelectMany(item => new[] { item.CurrentPremium, item.RenewalPremium })
            .Where(value => value is > 0m)
            .Select(value => value!.Value)
            .DefaultIfEmpty(0m)
            .Max();
        var lossMaximum = CalculateLossMaximum(cycles);

        var premiumPoints = cycles
            .Select(item => new ClientWorkspacePremiumTrendPoint
            {
                Label = CycleLabel(item),
                CurrentPremium = item.CurrentPremium,
                RenewalPremium = item.RenewalPremium,
                CurrentBarHeight = ScalePremium(item.CurrentPremium, premiumMaximum),
                RenewalBarHeight = ScalePremium(item.RenewalPremium, premiumMaximum)
            })
            .ToList();

        var lossPoints = cycles
            .Select((item, index) => new ClientWorkspaceLossRatioTrendPoint
            {
                Label = CycleLabel(item),
                X = ScaleX(index, cycles.Count),
                EhcLossRatio = item.EhcLossRatio,
                DentalLossRatio = item.DentalLossRatio,
                EhcY = ScaleLoss(item.EhcLossRatio, lossMaximum),
                DentalY = ScaleLoss(item.DentalLossRatio, lossMaximum)
            })
            .ToList();

        var currentScore = CalculateHealthScore(current, clientStatus);
        var priorScore = CalculateHealthScore(prior, clientStatus);

        return new ClientWorkspaceRenewalHistoryPresentation
        {
            Current = current,
            Prior = prior,
            CurrentHealthScore = currentScore,
            HealthScoreComparison = CompareScore(currentScore, priorScore),
            EhcComparison = CompareLossRatio(current?.EhcLossRatio, prior?.EhcLossRatio),
            DentalComparison = CompareLossRatio(current?.DentalLossRatio, prior?.DentalLossRatio),
            PremiumPoints = premiumPoints,
            LossRatioPoints = lossPoints,
            PeriodsText = BuildPeriodsText(cycles)
        };
    }

    public static int? CalculateHealthScore(
        ClientWorkspaceRenewalHistoryEntry? renewal,
        string clientStatus)
    {
        if (clientStatus.Contains("Prospect", StringComparison.OrdinalIgnoreCase) || renewal is null)
        {
            return null;
        }

        var score = 100;
        if (renewal.PremiumChangePercent is >= 10m)
        {
            score -= 20;
        }
        else if (renewal.PremiumChangePercent is >= 5m)
        {
            score -= 10;
        }

        if (renewal.EhcLossRatio is >= 100m)
        {
            score -= 18;
        }
        else if (renewal.EhcLossRatio is >= 85m)
        {
            score -= 10;
        }

        if (renewal.DentalLossRatio is >= 100m)
        {
            score -= 18;
        }
        else if (renewal.DentalLossRatio is >= 85m)
        {
            score -= 10;
        }

        if (renewal.IsHighRisk)
        {
            score -= 12;
        }

        return Math.Clamp(score, 0, 100);
    }

    private static bool HasRenewalIdentity(ClientWorkspaceRenewalHistoryEntry item)
    {
        return item.RenewalDate.HasValue || item.RenewalYear > 0;
    }

    private static string CycleKey(ClientWorkspaceRenewalHistoryEntry item)
    {
        return item.RenewalDate.HasValue
            ? $"date:{item.RenewalDate.Value:yyyy-MM-dd}"
            : $"year:{item.RenewalYear}";
    }

    private static DateOnly CycleDate(ClientWorkspaceRenewalHistoryEntry item)
    {
        return item.RenewalDate ??
               (item.RenewalYear > 0 ? new DateOnly(item.RenewalYear, 1, 1) : DateOnly.MinValue);
    }

    private static string CycleLabel(ClientWorkspaceRenewalHistoryEntry item)
    {
        return (item.RenewalDate?.Year ?? item.RenewalYear).ToString(CultureInfo.InvariantCulture);
    }

    private static double ScalePremium(decimal? value, decimal maximum)
    {
        if (value is not > 0m || maximum <= 0m)
        {
            return 0d;
        }

        return Math.Max(4d, Math.Round((double)(value.Value / maximum) * MaximumPremiumBarHeight, 1));
    }

    private static decimal CalculateLossMaximum(
        IReadOnlyCollection<ClientWorkspaceRenewalHistoryEntry> cycles)
    {
        var maximum = cycles
            .SelectMany(item => new[] { item.EhcLossRatio, item.DentalLossRatio })
            .Where(value => value is >= 0m)
            .Select(value => value!.Value)
            .DefaultIfEmpty(0m)
            .Max();

        return Math.Max(100m, Math.Ceiling(maximum / 10m) * 10m);
    }

    private static double ScaleX(int index, int count)
    {
        var plotWidth = LossChartWidth - (LossChartPadding * 2d);
        if (count <= 1)
        {
            return LossChartPadding + (plotWidth / 2d);
        }

        return LossChartPadding + (plotWidth * index / (count - 1d));
    }

    private static double? ScaleLoss(decimal? value, decimal maximum)
    {
        if (!value.HasValue || value.Value < 0m || maximum <= 0m)
        {
            return null;
        }

        var normalized = Math.Clamp(value.Value / maximum, 0m, 1m);
        var plotHeight = LossChartHeight - (LossChartPadding * 2d);
        return LossChartPadding + ((1d - (double)normalized) * plotHeight);
    }

    private static ClientWorkspaceMetricComparison CompareScore(int? current, int? prior)
    {
        if (!current.HasValue)
        {
            return ClientWorkspaceMetricComparison.Unavailable("Score unavailable");
        }

        if (!prior.HasValue)
        {
            return ClientWorkspaceMetricComparison.Unavailable("No prior renewal");
        }

        var difference = current.Value - prior.Value;
        return new ClientWorkspaceMetricComparison
        {
            Text = difference == 0
                ? "No change vs last renewal"
                : $"{difference:+0;-0} pts vs last renewal",
            Tone = difference switch
            {
                > 0 => ClientWorkspaceComparisonTone.Positive,
                < 0 => ClientWorkspaceComparisonTone.Negative,
                _ => ClientWorkspaceComparisonTone.Neutral
            }
        };
    }

    private static ClientWorkspaceMetricComparison CompareLossRatio(decimal? current, decimal? prior)
    {
        if (!current.HasValue)
        {
            return ClientWorkspaceMetricComparison.Unavailable("Current result unavailable");
        }

        if (!prior.HasValue)
        {
            return ClientWorkspaceMetricComparison.Unavailable("No prior renewal");
        }

        var difference = current.Value - prior.Value;
        return new ClientWorkspaceMetricComparison
        {
            Text = $"vs prior: {prior.Value:0.0}%",
            Tone = difference switch
            {
                < 0m => ClientWorkspaceComparisonTone.Positive,
                > 0m => ClientWorkspaceComparisonTone.Negative,
                _ => ClientWorkspaceComparisonTone.Neutral
            }
        };
    }

    private static string BuildPeriodsText(
        IReadOnlyCollection<ClientWorkspaceRenewalHistoryEntry> cycles)
    {
        if (cycles.Count == 0)
        {
            return "No saved renewal history";
        }

        var labels = cycles.Select(CycleLabel).ToList();
        return cycles.Count == 1
            ? $"{labels[0]} saved renewal"
            : $"{cycles.Count} saved renewals: {labels[0]} to {labels[^1]}";
    }
}

public sealed class ClientWorkspaceRenewalHistoryEntry
{
    public DateOnly? RenewalDate { get; init; }
    public int RenewalYear { get; init; }
    public decimal? CurrentPremium { get; init; }
    public decimal? RenewalPremium { get; init; }
    public decimal? PremiumChangePercent { get; init; }
    public decimal? EhcLossRatio { get; init; }
    public decimal? DentalLossRatio { get; init; }
    public bool IsHighRisk { get; init; }
    public DateTime DateLoaded { get; init; }
}

public sealed class ClientWorkspaceRenewalHistoryPresentation
{
    public ClientWorkspaceRenewalHistoryEntry? Current { get; init; }
    public ClientWorkspaceRenewalHistoryEntry? Prior { get; init; }
    public int? CurrentHealthScore { get; init; }
    public ClientWorkspaceMetricComparison HealthScoreComparison { get; init; } =
        ClientWorkspaceMetricComparison.Unavailable("No prior renewal");
    public ClientWorkspaceMetricComparison EhcComparison { get; init; } =
        ClientWorkspaceMetricComparison.Unavailable("No prior renewal");
    public ClientWorkspaceMetricComparison DentalComparison { get; init; } =
        ClientWorkspaceMetricComparison.Unavailable("No prior renewal");
    public IReadOnlyList<ClientWorkspacePremiumTrendPoint> PremiumPoints { get; init; } = [];
    public IReadOnlyList<ClientWorkspaceLossRatioTrendPoint> LossRatioPoints { get; init; } = [];
    public string PeriodsText { get; init; } = "No saved renewal history";
}

public sealed class ClientWorkspacePremiumTrendPoint
{
    public string Label { get; init; } = "";
    public decimal? CurrentPremium { get; init; }
    public decimal? RenewalPremium { get; init; }
    public double CurrentBarHeight { get; init; }
    public double RenewalBarHeight { get; init; }
    public string CurrentPremiumText => CurrentPremium.HasValue
        ? $"Current: {CurrentPremium.Value.ToString("$#,0", CultureInfo.InvariantCulture)}"
        : "Current: N/A";
    public string RenewalPremiumText => RenewalPremium.HasValue
        ? $"Renewal: {RenewalPremium.Value.ToString("$#,0", CultureInfo.InvariantCulture)}"
        : "Renewal: N/A";
}

public sealed class ClientWorkspaceLossRatioTrendPoint
{
    public string Label { get; init; } = "";
    public double X { get; init; }
    public decimal? EhcLossRatio { get; init; }
    public decimal? DentalLossRatio { get; init; }
    public double? EhcY { get; init; }
    public double? DentalY { get; init; }
    public string EhcToolTip => EhcLossRatio.HasValue
        ? $"{Label} EHC: {EhcLossRatio.Value:0.0}%"
        : $"{Label} EHC: N/A";
    public string DentalToolTip => DentalLossRatio.HasValue
        ? $"{Label} Dental: {DentalLossRatio.Value:0.0}%"
        : $"{Label} Dental: N/A";
}

public sealed class ClientWorkspaceMetricComparison
{
    public string Text { get; init; } = "";
    public ClientWorkspaceComparisonTone Tone { get; init; }

    public static ClientWorkspaceMetricComparison Unavailable(string text)
    {
        return new ClientWorkspaceMetricComparison
        {
            Text = text,
            Tone = ClientWorkspaceComparisonTone.Unavailable
        };
    }
}

public enum ClientWorkspaceComparisonTone
{
    Unavailable,
    Neutral,
    Positive,
    Negative
}

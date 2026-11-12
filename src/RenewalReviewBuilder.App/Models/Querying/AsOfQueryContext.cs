using System.Threading;

namespace RenewalReviewBuilder.App.Models.Querying;

/// <summary>
/// Describes how completely a source can answer a historical query.
/// </summary>
public enum HistoricalQuerySupport
{
    Supported,
    Limited,
    Unavailable
}
/// <summary>
/// Reusable, read-only point-in-time query context for canonical projections.
/// </summary>
public sealed class AsOfQueryContext
{
    public long ClientId { get; init; }
    public DateOnly? AsOfDate { get; init; }
    public bool IncludeManualEvents { get; init; } = true;
    public bool IncludeUnvalidatedEvents { get; init; }
    public bool IncludeFutureEvents { get; init; }
    public CancellationToken CancellationToken { get; init; }
}

/// <summary>
/// A contributor's explicit historical-query capability for one query.
/// </summary>
public sealed class HistoricalQuerySupportStatus
{
    public string SourceModule { get; init; } = "";
    public HistoricalQuerySupport Support { get; init; }
    public string Detail { get; init; } = "";
}
using System.Globalization;

namespace RenewalReviewBuilder.App.Models.History;

public enum HistoryDatePrecision
{
    Unknown,
    ExactDate,
    MonthAndYear,
    YearOnly
}

/// <summary>
/// Represents the precision supplied by a canonical history source without manufacturing a day.
/// </summary>
public sealed class HistoryDateReference
{
    public DateOnly? ExactDate { get; init; }
    public int? Year { get; init; }
    public int? Month { get; init; }
    public HistoryDatePrecision Precision { get; init; } = HistoryDatePrecision.Unknown;

    public int SortYear => ExactDate?.Year ?? Year ?? 0;
    public int SortMonth => ExactDate?.Month ?? Month ?? 0;
    public int SortDay => ExactDate?.Day ?? 0;

    public string DisplayText => Precision switch
    {
        HistoryDatePrecision.ExactDate when ExactDate.HasValue => ExactDate.Value.ToString("MMM d, yyyy", CultureInfo.InvariantCulture),
        HistoryDatePrecision.MonthAndYear when Year.HasValue && Month is >= 1 and <= 12 =>
            CultureInfo.InvariantCulture.DateTimeFormat.GetAbbreviatedMonthName(Month.Value) + " " + Year.Value.ToString(CultureInfo.InvariantCulture),
        HistoryDatePrecision.YearOnly when Year.HasValue => Year.Value.ToString(CultureInfo.InvariantCulture),
        _ => "Date unavailable"
    };

    public bool IsOnOrAfter(DateOnly boundary) => Precision switch
    {
        HistoryDatePrecision.ExactDate when ExactDate.HasValue => ExactDate.Value >= boundary,
        HistoryDatePrecision.MonthAndYear when Year.HasValue && Month.HasValue =>
            Year.Value > boundary.Year || (Year.Value == boundary.Year && Month.Value >= boundary.Month),
        HistoryDatePrecision.YearOnly when Year.HasValue => Year.Value >= boundary.Year,
        _ => true
    };

    public bool IsOnOrBefore(DateOnly boundary) => Precision switch
    {
        HistoryDatePrecision.ExactDate when ExactDate.HasValue => ExactDate.Value <= boundary,
        HistoryDatePrecision.MonthAndYear when Year.HasValue && Month.HasValue =>
            Year.Value < boundary.Year || (Year.Value == boundary.Year && Month.Value <= boundary.Month),
        HistoryDatePrecision.YearOnly when Year.HasValue => Year.Value <= boundary.Year,
        _ => true
    };

    public static HistoryDateReference FromExact(DateOnly date) => new()
    {
        ExactDate = date,
        Year = date.Year,
        Month = date.Month,
        Precision = HistoryDatePrecision.ExactDate
    };

    public static HistoryDateReference FromMonth(int year, int month)
    {
        if (month is < 1 or > 12)
        {
            throw new ArgumentOutOfRangeException(nameof(month));
        }

        return new HistoryDateReference
        {
            Year = year,
            Month = month,
            Precision = HistoryDatePrecision.MonthAndYear
        };
    }

    public static HistoryDateReference FromYear(int year) => new()
    {
        Year = year,
        Precision = HistoryDatePrecision.YearOnly
    };
}

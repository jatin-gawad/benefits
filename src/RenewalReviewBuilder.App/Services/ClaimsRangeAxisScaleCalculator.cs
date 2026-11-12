namespace RenewalReviewBuilder.App.Services;

public sealed record ClaimsRangeAxisScale(double Minimum, double Maximum, double MajorUnit);

/// <summary>Builds a readable, unbounded whole-certificate scale with space for data labels.</summary>
public sealed class ClaimsRangeAxisScaleCalculator
{
    private const int TargetIntervals = 5;

    public ClaimsRangeAxisScale Calculate(IEnumerable<int> counts)
    {
        ArgumentNullException.ThrowIfNull(counts);
        var values = counts.ToList();
        if (values.Any(value => value < 0))
        {
            throw new ArgumentOutOfRangeException(nameof(counts), "Claims Range counts cannot be negative.");
        }

        var maximumCount = values.Count == 0 ? 0 : values.Max();
        if (maximumCount == 0)
        {
            return new ClaimsRangeAxisScale(0d, 1d, 1d);
        }

        var roughMajorUnit = maximumCount / (double)TargetIntervals;
        var majorUnit = Math.Max(1d, NiceCeiling(roughMajorUnit));
        var maximum = Math.Ceiling(maximumCount / majorUnit) * majorUnit;

        // Keep a label from sitting on the chart boundary when the maximum lands on a gridline.
        if (maximum - maximumCount < majorUnit * 0.25d)
        {
            maximum += majorUnit;
        }

        return new ClaimsRangeAxisScale(0d, maximum, majorUnit);
    }

    private static double NiceCeiling(double value)
    {
        var exponent = Math.Floor(Math.Log10(value));
        var magnitude = Math.Pow(10d, exponent);
        var fraction = value / magnitude;
        var niceFraction = fraction switch
        {
            <= 1d => 1d,
            <= 2d => 2d,
            <= 2.5d => 2.5d,
            <= 5d => 5d,
            _ => 10d
        };
        return niceFraction * magnitude;
    }
}

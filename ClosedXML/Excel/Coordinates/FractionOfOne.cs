using System;

namespace ClosedXML.Excel;

/// <summary>
/// Represents a percentage position from 0 to 1 represented as a double (=inherently imprecise).
/// In most cases, it represents a position relative to a different size. Some other percentage
/// that are also limited to values 0%-100% often have a fixed precision (e.g.,
/// <c>ST_PositiveFixedPercentage</c> has precision of 1/1000).
/// </summary>
internal readonly record struct FractionOfOne : IEquatable<double>
{
    private FractionOfOne(double value)
    {
        if (double.IsInfinity(value) || double.IsNaN(value) || value is < 0 or > 1)
            throw new ArgumentOutOfRangeException(nameof(value), "Value must be between 0 and 1.");

        Value = value;
    }

    public double Value { get; }

    public static implicit operator FractionOfOne(double value) => new(value);

    public bool Equals(double other)
    {
        return Value.Equals(other);
    }
}

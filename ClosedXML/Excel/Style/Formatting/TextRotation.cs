using System;

namespace ClosedXML.Excel.Formatting;

/// <summary>
/// An angle of text rotation of a text in a cell.
/// </summary>
internal readonly record struct TextRotation
{
    public static readonly TextRotation None = new(0);

    public static readonly TextRotation VerticalText = new(255);

    public TextRotation(int value)
    {
        if (value is not (>= -90 and <= 90 or 255))
            throw new ArgumentOutOfRangeException(nameof(value), value, "Rotation must be between [-90,90] or 255 for vertical text.");

        Value = value;
    }

    public int Value { get; }

    /// <summary>
    /// Get value that is stored in ISO-29500 (unsigned int)
    /// </summary>
    internal uint GetIso()
    {
        return Value switch
        {
            >= 0 and <= 90 => (uint)Value,
            >= -90 and <= 0 => (uint)(90 - Value),
            _ => (uint)Value
        };
    }
}

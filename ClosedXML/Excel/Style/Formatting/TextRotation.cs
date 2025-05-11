using System;

namespace ClosedXML.Excel.Formatting;

/// <summary>
/// An angle of text rotation of a text in a cell.
/// </summary>
internal readonly record struct TextRotation
{
    public static readonly TextRotation VerticalText = new(255);

    public TextRotation(int value)
    {
        if (value is not (>= -90 and <= 90 or 255))
            throw new ArgumentOutOfRangeException();

        Value = value;
    }

    public int Value { get; }
}

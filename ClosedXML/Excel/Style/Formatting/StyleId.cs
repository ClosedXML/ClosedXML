using System;

namespace ClosedXML.Excel.Formatting;

/// <summary>
/// A key to a named style that is used to connect immutable <see cref="XLCellFormatValue"/> to
/// a changeable <see cref="XLCellStyleValue"/>. API methods that needs access to a resolved format
/// value should pass <see cref="XLWorkbookStyles.DefaultFormat"/> as a parameter.
/// </summary>
internal readonly record struct StyleId(int Value) : IEquatable<int>
{
    public static implicit operator StyleId(int v) => new(v);

    bool IEquatable<int>.Equals(int other)
    {
        return Value == other;
    }
}

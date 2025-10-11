using System;
using System.Text;
using ClosedXML.Parser;

namespace ClosedXML.Excel;

/// <summary>
/// An immutable row address within a workbook.
/// </summary>
internal readonly record struct XLRowArea
{
    public XLRowArea(string name, int rowNumber)
    {
        if (string.IsNullOrEmpty(name))
            throw new ArgumentException(nameof(name));

        if (rowNumber is < XLHelper.MinRowNumber or > XLHelper.MaxRowNumber)
            throw new ArgumentOutOfRangeException(nameof(rowNumber));

        Name = name;
        RowNumber = rowNumber;
    }

    /// <summary>
    /// Name of the sheet. Sheet may exist or not (e.g. deleted). Never null.
    /// </summary>
    public string Name { get; }

    /// <summary>
    /// Row number, ranges from 1 to <see cref="XLHelper.MaxRowNumber"/>.
    /// </summary>
    public int RowNumber { get; }

    /// <summary>
    /// Get the area of the row.
    /// </summary>
    public XLBookArea Area => new(Name, new XLSheetRange(RowNumber, XLHelper.MinColumnNumber, RowNumber, XLHelper.MaxColumnNumber));

    public bool Equals(XLRowArea other)
    {
        return RowNumber == other.RowNumber && XLHelper.SheetComparer.Equals(Name, other.Name);
    }

    public override int GetHashCode()
    {
        unchecked
        {
            return (XLHelper.SheetComparer.GetHashCode(Name) * 397) ^ RowNumber.GetHashCode();
        }
    }

    public override string ToString()
    {
        var name = NameUtils.ShouldQuote(Name.AsSpan()) ? Name.AlwaysEscapeSheetName() : Name;
        return new StringBuilder(name.Length + 1 + 7 + 1 + 7)
            .Append(name).Append('!').Append(RowNumber).Append(':').Append(RowNumber)
            .ToString();
    }
}

using System;

namespace ClosedXML.Excel;

/// <summary>
/// An immutable column address within a workbook.
/// </summary>
internal readonly record struct XLColumnArea
{
    public XLColumnArea(string name, int columnNumber)
    {
        if (string.IsNullOrEmpty(name))
            throw new ArgumentException(nameof(name));

        if (columnNumber is < XLHelper.MinColumnNumber or > XLHelper.MaxColumnNumber)
            throw new ArgumentOutOfRangeException(nameof(columnNumber));

        Name = name;
        ColumNumber = columnNumber;
    }

    /// <summary>
    /// Name of the sheet. Sheet may exist or not (e.g. deleted). Never null.
    /// </summary>
    public string Name { get; }

    /// <summary>
    /// Column number, ranges from 1 to <see cref="XLHelper.MaxColumnNumber"/>.
    /// </summary>
    public int ColumNumber { get; }

    public XLBookArea Area => new(Name, new XLSheetRange(XLHelper.MinRowNumber, ColumNumber, XLHelper.MaxRowNumber, ColumNumber));

    public bool Equals(XLColumnArea other)
    {
        return ColumNumber == other.ColumNumber && XLHelper.SheetComparer.Equals(Name, other.Name);
    }

    public override int GetHashCode()
    {
        unchecked
        {
            return (XLHelper.SheetComparer.GetHashCode(Name) * 397) ^ ColumNumber.GetHashCode();
        }
    }
}

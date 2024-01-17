using ClosedXML.Extensions;
using ClosedXML.Parser;

namespace ClosedXML.Excel;

/// <summary>
/// A reference without a sheet. Can represent single cell (<c>A1</c>), area
/// (<c>B$4:$D$10</c>), row span (<c>4:10</c>) and col span (<c>G:H</c>).
/// </summary>
/// <remarks>
/// This is an actual representation of a reference, while the <see cref="XLSheetRange"/> is for
/// an absolute are of a sheet and <see cref="XLAddress"/> is only for a cell reference and
/// <see cref="XLRangeAddress"/> only for area reference.
/// </remarks>
internal readonly record struct XLReference
{
    private readonly ReferenceArea _reference;

    internal XLReference(ReferenceArea reference)
    {
        _reference = reference;
    }

    internal string GetA1()
    {
        return _reference.GetDisplayStringA1();
    }

    internal XLRangeAddress ToRangeAddress(XLWorksheet? sheet, XLSheetPoint anchor)
    {
        var area = _reference.ToSheetRange(anchor);
        var firstColAbs = _reference.First.ColumnType == ReferenceAxisType.Absolute;
        var firstRowAbs = _reference.First.RowType == ReferenceAxisType.Absolute;
        var secondColAbs = _reference.Second.ColumnType == ReferenceAxisType.Absolute;
        var secondRowAbs = _reference.Second.RowType == ReferenceAxisType.Absolute;
        if (_reference.First.IsColumn)
        {
            // Column span
            firstRowAbs = true;
            secondRowAbs = true;
        }

        if (_reference.First.IsRow)
        {
            // Row span
            firstColAbs = true;
            secondColAbs = true;
        }

        return new XLRangeAddress(
            new XLAddress(sheet, area.TopRow, area.LeftColumn, firstRowAbs, firstColAbs),
            new XLAddress(sheet, area.BottomRow, area.RightColumn, secondRowAbs, secondColAbs));
    }
}

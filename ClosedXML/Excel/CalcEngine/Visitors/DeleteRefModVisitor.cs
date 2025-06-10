using ClosedXML.Extensions;
using ClosedXML.Parser;
using System;
using System.Diagnostics;

namespace ClosedXML.Excel.CalcEngine.Visitors;

/// <summary>
/// A RefModVisitor that adjusts a reference in a formula when an area is deleted and shifted up/left.
/// </summary>
internal class ReferenceShiftOnDeleteRefModVisitor : CopyVisitor
{
    private const string RefError = "#REF!";
    private readonly XLBookArea _deletedBookArea;
    private readonly XLShiftDeletedCells _shift;

    public ReferenceShiftOnDeleteRefModVisitor(XLBookArea deletedBookArea, XLShiftDeletedCells shift)
    {
        _deletedBookArea = deletedBookArea;
        _shift = shift;
    }

    public override TransformedSymbol Reference(ModContext ctx, SymbolRange range, ReferenceArea referenceToShift)
    {
        // If reference is for a different sheet than deleted one, return the original reference.
        if (!XLHelper.SheetComparer.Equals(_deletedBookArea.Name, ctx.Sheet))
            return TransformedSymbol.CopyOriginal(ctx.Formula, range);

        // The two methods could be transposed into a single case, but it is hard to debug and I
        // will rather take some duplication. Plus there are some problems like transposing row
        // beyond last column and so on. Make sure to fix/refactor BOTH variants.
        return _shift switch
        {
            XLShiftDeletedCells.ShiftCellsUp => DeleteAndShiftUp(ctx, range, referenceToShift),
            XLShiftDeletedCells.ShiftCellsLeft => DeleteAndShiftLeft(ctx, range, referenceToShift),
            _ => throw new UnreachableException()
        };
    }

    private TransformedSymbol DeleteAndShiftUp(ModContext ctx, SymbolRange range, ReferenceArea referenceToShift)
    {
        // Columns are never changed by shift up deletion.
        if (referenceToShift.IsColumnSpan())
            return TransformedSymbol.CopyOriginal(ctx.Formula, range);

        var referenceToShiftArea = referenceToShift.ToSheetRangeA1();
        var deletedArea = _deletedBookArea.Area;

        // Deleted area is fully to the left of the reference - reference will never be shifted up
        if (deletedArea.RightColumn < referenceToShiftArea.LeftColumn)
            return TransformedSymbol.CopyOriginal(ctx.Formula, range);

        // Deleted area is fully to the right of the reference - reference will never be shifted up
        if (deletedArea.LeftColumn > referenceToShiftArea.RightColumn)
            return TransformedSymbol.CopyOriginal(ctx.Formula, range);

        // Deleted area doesn't fully cover reference columns. It will either cause splits or not move -> keep original.
        if (deletedArea.LeftColumn > referenceToShiftArea.LeftColumn &&
            deletedArea.RightColumn < referenceToShiftArea.RightColumn)
            return TransformedSymbol.CopyOriginal(ctx.Formula, range);

        // Deleted area fully cover reference columns.
        if (deletedArea.LeftColumn <= referenceToShiftArea.LeftColumn &&
            deletedArea.RightColumn >= referenceToShiftArea.RightColumn)
        {
            // Deleted area is below the reference
            if (deletedArea.TopRow > referenceToShiftArea.BottomRow)
                return TransformedSymbol.CopyOriginal(ctx.Formula, range);

            // The deleted area either partially covers the reference or is above the reference -> shift up.

            // How many rows to shift up the reference.
            var shiftUp = Math.Min(deletedArea.BottomRow + 1, referenceToShiftArea.TopRow) - Math.Min(deletedArea.TopRow, referenceToShiftArea.TopRow);

            // By how many rows to shrink the reference
            var shrinkBy = deletedArea.BottomRow >= referenceToShiftArea.TopRow
                ? Math.Min(referenceToShiftArea.BottomRow, deletedArea.BottomRow) - Math.Max(referenceToShiftArea.TopRow, deletedArea.TopRow) + 1
                : 0;

            // The reference was completely removed by deleted area
            if (shrinkBy >= referenceToShiftArea.Height)
                return TransformedSymbol.ToText(ctx.Formula, range, RefError);

            var first = Shift(referenceToShift.First, shiftUp, 0);
            var second = Shift(referenceToShift.Second, shiftUp + shrinkBy, 0);
            var shiftedReference = new ReferenceArea(first, second);
            return TransformedSymbol.ToText(ctx.Formula, range, shiftedReference.GetDisplayStringA1());
        }

        // Deleted area slices off a left part of the reference
        if (deletedArea.LeftColumn <= referenceToShiftArea.LeftColumn &&
            deletedArea.RightColumn < referenceToShiftArea.RightColumn &&
            deletedArea.RightColumn >= referenceToShiftArea.LeftColumn)
        {
            if (deletedArea.TopRow <= referenceToShiftArea.TopRow &&
                deletedArea.BottomRow >= referenceToShiftArea.BottomRow)
            {
                // Slice off the right side of the reference
                var sliceOffLeftColumns = referenceToShiftArea.LeftColumn - deletedArea.RightColumn - 1;
                var first = Shift(referenceToShift.First, 0, sliceOffLeftColumns);
                var shiftedReference = new ReferenceArea(first, referenceToShift.Second);
                return TransformedSymbol.ToText(ctx.Formula, range, shiftedReference.GetDisplayStringA1());
            }

            // Slicing would not change anything or would cause split -> keep original.
            return TransformedSymbol.CopyOriginal(ctx.Formula, range);
        }

        // Deleted area slices off a right part of the reference
        if (deletedArea.RightColumn >= referenceToShiftArea.RightColumn &&
            deletedArea.LeftColumn > referenceToShiftArea.LeftColumn &&
            deletedArea.LeftColumn <= referenceToShiftArea.RightColumn)
        {
            if (deletedArea.TopRow <= referenceToShiftArea.TopRow &&
                deletedArea.BottomRow >= referenceToShiftArea.BottomRow)
            {
                // Slice off the right side of the reference
                var sliceOffRightColumns = referenceToShiftArea.RightColumn - deletedArea.LeftColumn + 1;
                var second = Shift(referenceToShift.Second, 0, sliceOffRightColumns);
                var shiftedReference = new ReferenceArea(referenceToShift.First, second);
                return TransformedSymbol.ToText(ctx.Formula, range, shiftedReference.GetDisplayStringA1());
            }

            // Slicing would not change anything or would cause split -> keep original.
            return TransformedSymbol.CopyOriginal(ctx.Formula, range);
        }

        throw new UnreachableException($"Unhandled case between a delete area {deletedArea} and a reference area {referenceToShiftArea}.");
    }

    private TransformedSymbol DeleteAndShiftLeft(ModContext ctx, SymbolRange range, ReferenceArea referenceToShift)
    {
        // Rows are never changed by shift left deletion.
        if (referenceToShift.IsRowSpan())
            return TransformedSymbol.CopyOriginal(ctx.Formula, range);

        var referenceToShiftArea = referenceToShift.ToSheetRangeA1();
        var deletedArea = _deletedBookArea.Area;

        // Subtraction would cause split -> return original
        if (!referenceToShiftArea.TrySubtract(deletedArea, out var subtracted))
            return TransformedSymbol.CopyOriginal(ctx.Formula, range);

        // Whole area was subtracted -> #REF!
        if (subtracted is null)
            return TransformedSymbol.ToText(ctx.Formula, range, RefError);

        // If delete area is to the left and covers full height of the subtracted area, then shift
        var shouldShiftToLeft = deletedArea.RightColumn < subtracted.Value.LeftColumn &&
                                deletedArea.TopRow <= subtracted.Value.TopRow &&
                                deletedArea.BottomRow >= subtracted.Value.BottomRow;
        var result = shouldShiftToLeft
            ? subtracted.Value.ShiftColumns(-deletedArea.Width)
            : subtracted.Value;
        var first = Set(referenceToShift.First, result.TopRow, result.LeftColumn);
        var second = Set(referenceToShift.Second, result.BottomRow, result.RightColumn);
        var shiftedReference = new ReferenceArea(first, second);
        return TransformedSymbol.ToText(ctx.Formula, range, shiftedReference.GetDisplayStringA1());
    }

    private static RowCol Shift(RowCol rowCol, int rowUpShift, int columnLeftShift)
    {
        var r = rowCol.RowType != ReferenceAxisType.None ? rowCol.RowValue - rowUpShift : rowCol.RowValue;
        var c = rowCol.ColumnType != ReferenceAxisType.None ? rowCol.ColumnValue - columnLeftShift : rowCol.ColumnValue;
        return new RowCol(rowCol.RowType, r, rowCol.ColumnType, c, rowCol.Style);
    }
    private static RowCol Set(RowCol rowCol, int row, int column)
    {
        var r = rowCol.RowType != ReferenceAxisType.None ? row : rowCol.RowValue;
        var c = rowCol.ColumnType != ReferenceAxisType.None ? column : rowCol.ColumnValue;
        return new RowCol(rowCol.RowType, r, rowCol.ColumnType, c, rowCol.Style);
    }
}

using System.Diagnostics;
using ClosedXML.Extensions;
using ClosedXML.Parser;

namespace ClosedXML.Excel.CalcEngine.Visitors;

/// <summary>
/// A RefModVisitor that adjusts a reference in a formula when an area is deleted and shifted up/left.
/// </summary>
internal class ReferenceShiftOnDeleteRefModVisitor : CopyVisitor
{
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
        // will rather take some duplication.
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

        // Subtraction would cause split -> return original
        if (!referenceToShiftArea.TrySubtract(deletedArea, out var subtracted))
            return TransformedSymbol.CopyOriginal(ctx.Formula, range);

        // Whole area was subtracted -> #REF!
        if (subtracted is null)
            return TransformedSymbol.ToText(ctx.Formula, range, XLHelper.RefError);

        // If delete area is upwards and covers full width of the subtracted area, then shift
        var shouldShiftUpwards = deletedArea.BottomRow < subtracted.Value.TopRow &&
                                 deletedArea.LeftColumn <= subtracted.Value.LeftColumn &&
                                 deletedArea.RightColumn >= subtracted.Value.RightColumn;
        var result = shouldShiftUpwards
            ? subtracted.Value.ShiftRows(-deletedArea.Height)
            : subtracted.Value;

        // If there was no change, just use original text. Don't allocate new symbol needlessly.
        if (result == referenceToShiftArea)
            return TransformedSymbol.CopyOriginal(ctx.Formula, range);

        var first = Set(referenceToShift.First, result.TopRow, result.LeftColumn);
        var second = Set(referenceToShift.Second, result.BottomRow, result.RightColumn);
        var shiftedReference = new ReferenceArea(first, second);
        return TransformedSymbol.ToText(ctx.Formula, range, shiftedReference.GetDisplayStringA1());
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
            return TransformedSymbol.ToText(ctx.Formula, range, XLHelper.RefError);

        // If delete area is to the left and covers full height of the subtracted area, then shift
        var shouldShiftToLeft = deletedArea.RightColumn < subtracted.Value.LeftColumn &&
                                deletedArea.TopRow <= subtracted.Value.TopRow &&
                                deletedArea.BottomRow >= subtracted.Value.BottomRow;
        var result = shouldShiftToLeft
            ? subtracted.Value.ShiftColumns(-deletedArea.Width)
            : subtracted.Value;

        // If there was no change, just use original text. Don't allocate new symbol needlessly.
        if (result == referenceToShiftArea)
            return TransformedSymbol.CopyOriginal(ctx.Formula, range);

        var first = Set(referenceToShift.First, result.TopRow, result.LeftColumn);
        var second = Set(referenceToShift.Second, result.BottomRow, result.RightColumn);
        var shiftedReference = new ReferenceArea(first, second);
        return TransformedSymbol.ToText(ctx.Formula, range, shiftedReference.GetDisplayStringA1());
    }

    private static RowCol Set(RowCol rowCol, int row, int column)
    {
        var r = rowCol.RowType != ReferenceAxisType.None ? row : rowCol.RowValue;
        var c = rowCol.ColumnType != ReferenceAxisType.None ? column : rowCol.ColumnValue;
        return new RowCol(rowCol.RowType, r, rowCol.ColumnType, c, rowCol.Style);
    }
}

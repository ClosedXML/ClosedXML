using System;
using ClosedXML.Extensions;
using ClosedXML.Parser;

namespace ClosedXML.Excel.CalcEngine.Visitors;

/// <summary>
/// A RefModVisitor that adjusts a reference in a formula when an area is inserted and cells are shifted down/right.
/// </summary>
internal class ReferenceShiftOnInsertRefModVisitor : CopyVisitor
{
    private readonly XLBookArea _insertedBookArea;
    private readonly bool _shiftDown;

    internal ReferenceShiftOnInsertRefModVisitor(XLBookArea insertedBookArea, bool shiftDown)
    {
        _insertedBookArea = insertedBookArea;
        _shiftDown = shiftDown;
    }

    public override TransformedSymbol Reference(ModContext ctx, SymbolRange range, ReferenceArea referenceToShift)
    {
        if (!XLHelper.SheetComparer.Equals(_insertedBookArea.Name, ctx.Sheet))
            return TransformedSymbol.CopyOriginal(ctx.Formula, range);

        return _shiftDown ? InsertAndShiftDown(ctx, range, referenceToShift) : throw new NotImplementedException();
    }

    private TransformedSymbol InsertAndShiftDown(ModContext ctx, SymbolRange range, ReferenceArea referenceToShift)
    {
        // Return original reference if the shift would cause a split
        if (!referenceToShift.TryInsertAndShiftDown(_insertedBookArea.Area, out var shiftedReference))
            return TransformedSymbol.CopyOriginal(ctx.Formula, range);

        // If reference was shifted out of sheet, return #REF!
        if (shiftedReference is null)
            return TransformedSymbol.ToText(ctx.Formula, range, XLHelper.RefError);

        // Do not allocate a new string unless necessary
        if (referenceToShift == shiftedReference.Value)
            return TransformedSymbol.CopyOriginal(ctx.Formula, range);

        return TransformedSymbol.ToText(ctx.Formula, range, shiftedReference.Value.GetDisplayStringA1());
    }
}

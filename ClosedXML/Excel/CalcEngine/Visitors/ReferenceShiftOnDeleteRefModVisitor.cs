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

    public override TransformedSymbol SheetReference(ModContext ctx, SymbolRange range, string sheet, ReferenceArea reference)
    {
        return ShiftFormulaReferences(ctx, range, sheet, reference);
    }

    public override TransformedSymbol Reference(ModContext ctx, SymbolRange range, ReferenceArea reference)
    {
        return ShiftFormulaReferences(ctx, range, null, reference);
    }

    private TransformedSymbol ShiftFormulaReferences(ModContext ctx, SymbolRange range, string? referenceSheetName, ReferenceArea referenceToShift)
    {
        if (!XLHelper.SheetComparer.Equals(_deletedBookArea.Name, referenceSheetName ?? ctx.Sheet))
            return TransformedSymbol.CopyOriginal(ctx.Formula, range);

        var wouldSplitArea = _shift == XLShiftDeletedCells.ShiftCellsUp
            ? !referenceToShift.TryDeleteAndShiftUp(_deletedBookArea.Area, out var shiftedReference)
            : !referenceToShift.TryDeleteAndShiftLeft(_deletedBookArea.Area, out shiftedReference);

        // Return original reference if the shift would cause a split
        if (wouldSplitArea)
            return TransformedSymbol.CopyOriginal(ctx.Formula, range);

        // If reference was shifted out of sheet, return #REF!
        if (shiftedReference is null)
            return TransformedSymbol.ToText(ctx.Formula, range, XLHelper.RefError);

        // Do not allocate a new string unless necessary
        if (referenceToShift == shiftedReference.Value)
            return TransformedSymbol.CopyOriginal(ctx.Formula, range);

        var shiftedReferenceA1 = shiftedReference.Value.GetDisplayStringA1(referenceSheetName);
        return TransformedSymbol.ToText(ctx.Formula, range, shiftedReferenceA1);
    }
}

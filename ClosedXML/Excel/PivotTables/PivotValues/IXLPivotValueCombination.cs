#nullable disable

namespace ClosedXML.Excel
{
    public interface IXLPivotValueCombination
    {
        IXLPivotValue And(XLCellValue item);

        IXLPivotValue AndNext();

        IXLPivotValue AndPrevious();
    }
}

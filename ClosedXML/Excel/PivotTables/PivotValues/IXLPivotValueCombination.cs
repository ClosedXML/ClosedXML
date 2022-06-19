namespace ClosedXML.Excel
{
    public interface IXLPivotValueCombination
    {
        IXLPivotValue And(string item);
        IXLPivotValue AndPrevious();
        IXLPivotValue AndNext();
    }
}

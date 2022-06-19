namespace ClosedXML.Excel
{
    public enum XLCFIconSetOperator {GreaterThan, EqualOrGreaterThan}
    public interface IXLCFIconSet
    {
        IXLCFIconSet AddValue(XLCFIconSetOperator setOperator, string value, XLCFContentType type);
        IXLCFIconSet AddValue(XLCFIconSetOperator setOperator, double value, XLCFContentType type);
    }
}

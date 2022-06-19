namespace ClosedXML.Excel
{
    public enum XLCFContentType { Number, Percent, Formula, Percentile, Minimum, Maximum }
    public interface IXLCFColorScaleMin
    {
        IXLCFColorScaleMid Minimum(XLCFContentType type, string value, XLColor color);
        IXLCFColorScaleMid Minimum(XLCFContentType type, double value, XLColor color);
        IXLCFColorScaleMid LowestValue(XLColor color);
    }
}

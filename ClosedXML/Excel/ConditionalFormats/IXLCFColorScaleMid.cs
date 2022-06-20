namespace ClosedXML.Excel
{
    public interface IXLCFColorScaleMid
    {
        IXLCFColorScaleMax Midpoint(XLCFContentType type, string value, XLColor color);
        IXLCFColorScaleMax Midpoint(XLCFContentType type, double value, XLColor color);
        void Maximum(XLCFContentType type, string value, XLColor color);
        void Maximum(XLCFContentType type, double value, XLColor color);
        void HighestValue(XLColor color);
    }
}

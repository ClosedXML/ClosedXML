namespace ClosedXML.Excel
{
    public interface IXLCFDataBarMin
    {
        IXLCFDataBarMax Minimum(XLCFContentType type, string value);
        IXLCFDataBarMax Minimum(XLCFContentType type, double value);
        IXLCFDataBarMax LowestValue();
    }
}

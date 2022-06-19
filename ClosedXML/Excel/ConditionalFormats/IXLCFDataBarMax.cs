namespace ClosedXML.Excel
{
    public interface IXLCFDataBarMax
    {
        void Maximum(XLCFContentType type, string value);
        void Maximum(XLCFContentType type, double value);
        void HighestValue();
    }
}

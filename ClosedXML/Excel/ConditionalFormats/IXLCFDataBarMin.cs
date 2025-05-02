#nullable disable

using System;

namespace ClosedXML.Excel
{
    public interface IXLCFDataBarMin
    {
        IXLCFDataBarMax Minimum(XLCFContentType type, String value);
        IXLCFDataBarMax Minimum(XLCFContentType type, Double value);
        IXLCFDataBarMax LowestValue();
    }
}

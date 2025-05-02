#nullable disable

using System;

namespace ClosedXML.Excel
{
    public interface IXLCFDataBarMax
    {
        void Maximum(XLCFContentType type, String value);
        void Maximum(XLCFContentType type, Double value);
        void HighestValue();
    }
}

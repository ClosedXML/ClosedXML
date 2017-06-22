using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLCFColorScaleMax
    {
        void Maximum(XLCFContentType type, String value, XLColor color);
        void Maximum(XLCFContentType type, Double value, XLColor color);
        void HighestValue(XLColor color);
    }
}

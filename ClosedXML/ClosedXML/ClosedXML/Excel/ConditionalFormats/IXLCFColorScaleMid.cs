using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLCFColorScaleMid
    {
        IXLCFColorScaleMax Midpoint(XLCFContentType type, String value, IXLColor color);
        void Maximum(XLCFContentType type, String value, IXLColor color);
        void HighestValue(IXLColor color);
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLCFColorScaleMax
    {
        void Maximum(XLCFContentType type, String value, IXLColor color);
        void HighestValue(IXLColor color);
    }
}

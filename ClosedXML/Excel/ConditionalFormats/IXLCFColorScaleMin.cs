using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public enum XLCFContentType { Number, Percent, Formula, Percentile, Minimum, Maximum }
    public interface IXLCFColorScaleMin
    {
        IXLCFColorScaleMid Minimum(XLCFContentType type, String value, XLColor color);
        IXLCFColorScaleMid Minimum(XLCFContentType type, Double value, XLColor color);
        IXLCFColorScaleMid LowestValue(XLColor color);
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLCFDataBarMin
    {
        IXLCFDataBarMax Minimum(XLCFContentType type, String value);
        IXLCFDataBarMax Minimum(XLCFContentType type, Double value);
        IXLCFDataBarMax LowestValue();
    }
}

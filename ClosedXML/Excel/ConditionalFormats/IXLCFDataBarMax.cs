#nullable disable

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLCFDataBarMax
    {
        IXLCFDataBar Maximum(XLCFContentType type, String value);

        IXLCFDataBar Maximum(XLCFContentType type, Double value);

        IXLCFDataBar HighestValue();
    }
}

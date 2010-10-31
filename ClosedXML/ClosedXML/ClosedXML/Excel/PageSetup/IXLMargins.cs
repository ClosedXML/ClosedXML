using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLMargins
    {
        Double Left { get; set; }
        Double Right { get; set; }
        Double Top { get; set; }
        Double Bottom { get; set; }
        Double Header { get; set; }
        Double Footer { get; set; }
    }
}

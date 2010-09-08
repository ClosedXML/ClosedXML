using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLStyle
    {
        IXLAlignment Alignment { get; set; }
        IXLBorder Border { get; set; }
        IXLFill Fill { get; set; }
        IXLFont Font { get; set; }
        IXLNumberFormat NumberFormat { get; set; }

    }
}

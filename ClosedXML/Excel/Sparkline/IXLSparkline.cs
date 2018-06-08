using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLSparkline
    {        
        XLFormula Formula { get; set; }
        IXLCell Cell { get; set; }

        IXLSparklineGroup SparklineGroup { get; }

        IXLSparkline SetFormula(XLFormula value);
        IXLSparkline SetCell(IXLCell value);
    }
}


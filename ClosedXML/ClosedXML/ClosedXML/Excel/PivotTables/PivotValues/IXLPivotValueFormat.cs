using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLPivotValueFormat: IXLNumberFormatBase
    {
        IXLPivotValue SetNumberFormatId(Int32 value);
        IXLPivotValue SetFormat(String value);
    }
}

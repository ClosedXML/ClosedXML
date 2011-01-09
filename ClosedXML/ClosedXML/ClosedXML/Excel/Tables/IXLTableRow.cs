using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLTableRow: IXLRangeRow
    {
        IXLCell Field(Int32 index);
        IXLCell Field(String name);
    }
}

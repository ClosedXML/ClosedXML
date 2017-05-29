using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel.Drawings
{
    internal interface IXLMarker
    {
        Int32 ColumnId { get; set; }
        Int32 RowId { get; set; }
        Double ColumnOffset { get; set; }
        Double RowOffset { get; set; }
    }
}

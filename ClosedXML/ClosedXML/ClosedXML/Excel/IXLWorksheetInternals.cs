using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal interface IXLWorksheetInternals
    {
        Dictionary<IXLAddress, IXLCell> CellsCollection { get; }
        XLColumnsCollection ColumnsCollection { get; }
        XLRowsCollection RowsCollection { get; }
        List<String> MergedCells { get; }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal interface IXLWorksheetInternals
    {
        XLCellCollection CellsCollection { get; }
        XLColumnsCollection ColumnsCollection { get; }
        XLRowsCollection RowsCollection { get; }
        List<String> MergedCells { get; }
        XLWorkbook Workbook { get; }
    }
}

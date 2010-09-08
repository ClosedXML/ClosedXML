using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLWorksheetInternals
    {
        IXLAddress FirstCellAddress { get; }
        IXLAddress LastCellAddress { get; }
        Dictionary<IXLAddress, IXLCell> CellsCollection { get; }
        Dictionary<Int32, IXLColumn> ColumnsCollection { get; }
        Dictionary<Int32, IXLRow> RowsCollection { get; }
        List<String> MergedCells { get; }
    }
}

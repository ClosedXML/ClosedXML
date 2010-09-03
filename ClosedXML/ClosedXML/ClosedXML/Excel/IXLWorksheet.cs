using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLWorksheet: IXLRange
    {
        Dictionary<Int32, IXLColumn> ColumnsCollection { get; }
        Dictionary<Int32, IXLRow> RowsCollection { get; }
        new IXLRow Row(Int32 column);
        new IXLColumn Column(Int32 column);
        new IXLColumn Column(String column);
        String Name { get; set; }
        List<IXLColumn> Columns();

        void SetPrintArea(IXLRange range);
        void SetPrintArea(String rangeAddress);
        void SetPrintArea(IXLCell firstCell, IXLCell lastCell);
        void SetPrintArea(String firstCellAddress, String lastCellAddress);
        void SetPrintArea(IXLAddress firstCellAddress, IXLAddress lastCellAddress);
        void SetPrintArea(Int32 firstCellRow, Int32 firstCellColumn, Int32 lastCellRow, Int32 lastCellColumn);
        
        
    }
}

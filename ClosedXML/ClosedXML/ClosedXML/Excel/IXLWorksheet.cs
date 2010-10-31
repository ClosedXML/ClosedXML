using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLWorksheet: IXLRangeBase
    {
        Double ColumnWidth { get; set; }
        Double RowHeight { get; set; }

        String Name { get; set; }
        IXLPageSetup PageSetup { get; }

        IXLRow FirstRowUsed();
        IXLRow LastRowUsed();
        IXLColumn FirstColumnUsed();
        IXLColumn LastColumnUsed();
        IXLColumns Columns();
        IXLColumns Columns(String columns);
        IXLColumns Columns(String firstColumn, String lastColumn);
        IXLColumns Columns(Int32 firstColumn, Int32 lastColumn);
        IXLRows Rows();
        IXLRows Rows(String rows);
        IXLRows Rows(Int32 firstRow, Int32 lastRow);
        IXLRow Row(Int32 row);
        IXLColumn Column(Int32 column);
        IXLColumn Column(String column);
        IXLRange Range(int firstCellRow, int firstCellColumn, int lastCellRow, int lastCellColumn);

        IXLCell Cell(int row, int column);
        IXLCell Cell(string cellAddressInRange);
        IXLCell Cell(int row, string column);
        IXLCell Cell(IXLAddress cellAddressInRange);

        int RowCount();
        int ColumnCount();
    }
}

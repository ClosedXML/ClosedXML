using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLWorksheet: IXLRangeBase
    {
        Double DefaultColumnWidth { get; set; }
        Double DefaultRowHeight { get; set; }

        String Name { get; set; }
        IXLPageOptions PageSetup { get; }

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
    }
}

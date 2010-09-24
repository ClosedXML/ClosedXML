using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLWorksheet: IXLRange
    {
        new IXLRow Row(Int32 row);
        new IXLColumn Column(Int32 column);
        new IXLColumn Column(String column);
        String Name { get; set; }
        IXLColumns Columns();
        IXLColumns Columns(String columns);
        IXLColumns Columns(String firstColumn, String lastColumn);
        IXLColumns Columns(Int32 firstColumn, Int32 lastColumn);
        IXLRows Rows();
        IXLRows Rows(String rows);
        IXLRows Rows(Int32 firstRow, Int32 lastRow); 

        IXLPageSetup PageSetup { get; }
        new IXLWorksheetInternals Internals { get; }
    }
}

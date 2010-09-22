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
        List<IXLColumn> Columns(String columns);
        List<IXLColumn> Columns(String firstColumn, String lastColumn);
        List<IXLColumn> Columns(Int32 firstColumn, Int32 lastColumn);
        List<IXLRow> Rows();
        List<IXLRow> Rows(String rows);
        List<IXLRow> Rows(Int32 firstRow, Int32 lastRow); 

        IXLPageSetup PageSetup { get; }
        new IXLWorksheetInternals Internals { get; }
    }
}

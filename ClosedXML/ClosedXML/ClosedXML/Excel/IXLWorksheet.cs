using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLWorksheet: IXLRange
    {
        new IXLRow Row(Int32 column);
        new IXLColumn Column(Int32 column);
        new IXLColumn Column(String column);
        String Name { get; set; }
        List<IXLColumn> Columns();

        IXLPrintOptions PrintOptions { get; }
        new IXLWorksheetInternals Internals { get; }
    }
}

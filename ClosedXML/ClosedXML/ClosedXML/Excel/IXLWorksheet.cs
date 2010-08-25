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
    }
}

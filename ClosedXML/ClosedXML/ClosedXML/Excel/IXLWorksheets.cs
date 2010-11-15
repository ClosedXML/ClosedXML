using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLWorksheets: IEnumerable<IXLWorksheet>
    {
        IXLWorksheet Worksheet(String sheetName);
        IXLWorksheet Worksheet(Int32 sheetIndex);
        IXLWorksheet Add(String sheetName);
        void Delete(String sheetName);
        void Delete(Int32 sheetIndex);
    }
}

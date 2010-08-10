using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLWorksheets: IEnumerable<IXLWorksheet>
    {
        IXLWorksheet GetWorksheet(String sheetName);
        IXLWorksheet GetWorksheet(Int32 sheetIndex);
        IXLWorksheet Add(String sheetName);
        void Delete(String sheetName);
    }
}

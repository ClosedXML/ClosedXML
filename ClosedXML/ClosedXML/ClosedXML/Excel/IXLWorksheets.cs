using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace ClosedXML.Excel
{
    public interface IXLWorksheets: IEnumerable<IXLWorksheet>
    {
        IXLWorksheet Worksheet(String sheetName);
        IXLWorksheet Worksheet(Int32 position);
        IXLWorksheet Add(String sheetName);
        IXLWorksheet Add(String sheetName, Int32 position);
        IXLWorksheet Add(DataTable dataTable);
        void Add(DataSet dataSet);
        void Delete(String sheetName);
        void Delete(Int32 position);

        
    }
}

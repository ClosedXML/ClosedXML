using System;
using System.Collections.Generic;
using System.Data;

namespace ClosedXML.Excel
{
    public interface IXLWorksheets : IEnumerable<IXLWorksheet>
    {
        int Count { get; }

        IXLWorksheet Add();

        IXLWorksheet Add(Int32 position);

        IXLWorksheet Add(String sheetName);

        IXLWorksheet Add(String sheetName, Int32 position);

        IXLWorksheet Add(DataTable dataTable);

        IXLWorksheet Add(DataTable dataTable, String sheetName);

        void Add(DataSet dataSet);

        Boolean Contains(String sheetName);

        void Delete(String sheetName);

        void Delete(Int32 position);

        bool TryGetWorksheet(string sheetName, out IXLWorksheet worksheet);

        IXLWorksheet Worksheet(String sheetName);

        IXLWorksheet Worksheet(Int32 position);
    }
}

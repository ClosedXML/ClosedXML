using System.Collections.Generic;
using System.Data;

namespace ClosedXML.Excel
{
    public interface IXLWorksheets : IEnumerable<IXLWorksheet>
    {
        int Count { get; }

        IXLWorksheet Add();

        IXLWorksheet Add(int position);

        IXLWorksheet Add(string sheetName);

        IXLWorksheet Add(string sheetName, int position);

        IXLWorksheet Add(DataTable dataTable);

        IXLWorksheet Add(DataTable dataTable, string sheetName);

        void Add(DataSet dataSet);

        bool Contains(string sheetName);

        void Delete(string sheetName);

        void Delete(int position);

        bool TryGetWorksheet(string sheetName, out IXLWorksheet worksheet);

        IXLWorksheet Worksheet(string sheetName);

        IXLWorksheet Worksheet(int position);
    }
}

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

        /// <summary>
        /// Gets an <see cref="IXLWorksheet"/> by it's ID.
        /// </summary>
        /// <param name="id">The ID of the worksheet to look for</param>
        /// <returns>The requested <see cref="IXLWorksheet"/></returns>
        /// <throws>If no <see cref="IXLWorksheet"/> with the given ID was not found, an <see cref="InvalidOperationException"/> is thrown</throws>
        IXLWorksheet GetWorksheetById(int id);

        /// <summary>
        /// Trys to gets an <see cref="IXLWorksheet"/> by it's ID.
        /// </summary>
        /// <param name="id">The ID of the worksheet to look for</param>
        /// <param name="result">The <see cref="IXLWorksheet"/> if found, otherwise <see langword="null"/></param>
        /// <returns><see langword="true"/> if found, otherwise <see langword="false"/></returns>
        bool TryGetWorksheetById(int id, out IXLWorksheet result);
    }
}

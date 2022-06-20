using System.Collections;
using System.Collections.Generic;
using System.Data;

namespace ClosedXML.Excel
{
    public interface IXLTable : IXLRange
    {
        IXLAutoFilter AutoFilter { get; }
        IXLTableRange DataRange { get; }
        bool EmphasizeFirstColumn { get; set; }
        bool EmphasizeLastColumn { get; set; }
        IEnumerable<IXLTableField> Fields { get; }
        string Name { get; set; }
        bool ShowAutoFilter { get; set; }
        bool ShowColumnStripes { get; set; }
        bool ShowHeaderRow { get; set; }
        bool ShowRowStripes { get; set; }
        bool ShowTotalsRow { get; set; }
        XLTableTheme Theme { get; set; }

        /// <summary>
        /// Clears the contents of this table.
        /// </summary>
        /// <param name="clearOptions">Specify what you want to clear.</param>
        new IXLTable Clear(XLClearOptions clearOptions = XLClearOptions.All);

        IXLTableField Field(string fieldName);

        IXLTableField Field(int fieldIndex);

        IXLRangeRow HeadersRow();

        /// <summary>
        /// Appends the IEnumerable data elements and returns the range of the new rows.
        /// </summary>
        /// <param name="data">The IEnumerable data.</param>
        /// <param name="propagateExtraColumns">if set to <c>true</c> propagate extra columns' values and formulas.</param>
        /// <returns>
        /// The range of the new rows.
        /// </returns>
        IXLRange AppendData(IEnumerable data, bool propagateExtraColumns = false);

        /// <summary>
        /// Appends the IEnumerable data elements and returns the range of the new rows.
        /// </summary>
        /// <param name="data">The IEnumerable data.</param>
        /// <param name="transpose">if set to <c>true</c> the data will be transposed before inserting.</param>
        /// <param name="propagateExtraColumns">if set to <c>true</c> propagate extra columns' values and formulas.</param>
        /// <returns>
        /// The range of the new rows.
        /// </returns>
        IXLRange AppendData(IEnumerable data, bool transpose, bool propagateExtraColumns = false);

        /// <summary>
        /// Appends the data of a data table and returns the range of the new rows.
        /// </summary>
        /// <param name="dataTable">The data table.</param>
        /// <param name="propagateExtraColumns">if set to <c>true</c> propagate extra columns' values and formulas.</param>
        /// <returns>
        /// The range of the new rows.
        /// </returns>
        IXLRange AppendData(DataTable dataTable, bool propagateExtraColumns = false);

        /// <summary>
        /// Appends the IEnumerable data elements and returns the range of the new rows.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data">The table data.</param>
        /// <param name="propagateExtraColumns">if set to <c>true</c> propagate extra columns' values and formulas.</param>
        /// <returns>
        /// The range of the new rows.
        /// </returns>
        IXLRange AppendData<T>(IEnumerable<T> data, bool propagateExtraColumns = false);

        /// <summary>
        /// Replaces the IEnumerable data elements and returns the table's data range.
        /// </summary>
        /// <param name="data">The IEnumerable data.</param>
        /// <param name="propagateExtraColumns">if set to <c>true</c> propagate extra columns' values and formulas.</param>
        /// <returns>
        /// The table's data range.
        /// </returns>
        IXLRange ReplaceData(IEnumerable data, bool propagateExtraColumns = false);

        /// <summary>
        /// Replaces the IEnumerable data elements and returns the table's data range.
        /// </summary>
        /// <param name="data">The IEnumerable data.</param>
        /// <param name="transpose">if set to <c>true</c> the data will be transposed before inserting.</param>
        /// <param name="propagateExtraColumns">if set to <c>true</c> propagate extra columns' values and formulas.</param>
        /// <returns>
        /// The table's data range.
        /// </returns>
        IXLRange ReplaceData(IEnumerable data, bool transpose, bool propagateExtraColumns = false);

        /// <summary>
        /// Replaces the data from the records of a data table and returns the table's data range.
        /// </summary>
        /// <param name="dataTable">The data table.</param>
        /// <param name="propagateExtraColumns">if set to <c>true</c> propagate extra columns' values and formulas.</param>
        /// <returns>
        /// The table's data range.
        /// </returns>
        IXLRange ReplaceData(DataTable dataTable, bool propagateExtraColumns = false);

        /// <summary>
        /// Replaces the IEnumerable data elements as a table and the table's data range.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data">The table data.</param>
        /// <param name="propagateExtraColumns">if set to <c>true</c> propagate extra columns' values and formulas.</param>
        /// <returns>
        /// The table's data range.
        /// </returns>
        IXLRange ReplaceData<T>(IEnumerable<T> data, bool propagateExtraColumns = false);

        /// <summary>
        /// Resizes the table to the specified range address.
        /// </summary>
        /// <param name="range">The new table range.</param>
        /// <returns></returns>
        IXLTable Resize(IXLRange range);

        /// <summary>
        /// Resizes the table to the specified range address.
        /// </summary>
        /// <param name="rangeAddress">The range boundaries.</param>
        /// <returns></returns>
        IXLTable Resize(IXLRangeAddress rangeAddress);

        /// <summary>
        /// Resizes the table to the specified range address.
        /// </summary>
        /// <param name="rangeAddress">The range boundaries.</param>
        /// <returns></returns>
        IXLTable Resize(string rangeAddress);

        /// <summary>
        /// Resizes the table to the specified range.
        /// </summary>
        /// <param name="firstCell">The first cell in the range.</param>
        /// <param name="lastCell">The last cell in the range.</param>
        /// <returns></returns>
        IXLTable Resize(IXLCell firstCell, IXLCell lastCell);

        /// <summary>
        /// Resizes the table to the specified range.
        /// </summary>
        /// <param name="firstCellAddress">The first cell address in the worksheet.</param>
        /// <param name="lastCellAddress">The last cell address in the worksheet.</param>
        /// <returns></returns>
        IXLTable Resize(string firstCellAddress, string lastCellAddress);

        /// <summary>
        /// Resizes the table to the specified range.
        /// </summary>
        /// <param name="firstCellAddress">The first cell address in the worksheet.</param>
        /// <param name="lastCellAddress">The last cell address in the worksheet.</param>
        /// <returns></returns>
        IXLTable Resize(IXLAddress firstCellAddress, IXLAddress lastCellAddress);

        /// <summary>
        /// Resizes the table to the specified range.
        /// </summary>
        /// <param name="firstCellRow">The first cell's row of the range to return.</param>
        /// <param name="firstCellColumn">The first cell's column of the range to return.</param>
        /// <param name="lastCellRow">The last cell's row of the range to return.</param>
        /// <param name="lastCellColumn">The last cell's column of the range to return.</param>
        /// <returns></returns>
        IXLTable Resize(int firstCellRow, int firstCellColumn, int lastCellRow, int lastCellColumn);

        new IXLAutoFilter SetAutoFilter();

        IXLTable SetEmphasizeFirstColumn();

        IXLTable SetEmphasizeFirstColumn(bool value);

        IXLTable SetEmphasizeLastColumn();

        IXLTable SetEmphasizeLastColumn(bool value);

        IXLTable SetShowAutoFilter();

        IXLTable SetShowAutoFilter(bool value);

        IXLTable SetShowColumnStripes();

        IXLTable SetShowColumnStripes(bool value);

        IXLTable SetShowHeaderRow();

        IXLTable SetShowHeaderRow(bool value);

        IXLTable SetShowRowStripes();

        IXLTable SetShowRowStripes(bool value);

        IXLTable SetShowTotalsRow();

        IXLTable SetShowTotalsRow(bool value);

        IXLRangeRow TotalsRow();

        /// <summary>
        /// Converts the table to an enumerable of dynamic objects
        /// </summary>
        /// <returns></returns>
        IEnumerable<dynamic> AsDynamicEnumerable();

        /// <summary>
        /// Converts the table to a standard .NET System.Data.DataTable
        /// </summary>
        /// <returns></returns>
        DataTable AsNativeDataTable();

        IXLTable CopyTo(IXLWorksheet targetSheet);
    }
}

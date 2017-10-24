using System;
using System.Collections.Generic;
using System.Data;

namespace ClosedXML.Excel
{
    public interface IXLTable : IXLRange
    {
        IXLBaseAutoFilter AutoFilter { get; }
        IXLTableRange DataRange { get; }
        Boolean EmphasizeFirstColumn { get; set; }
        Boolean EmphasizeLastColumn { get; set; }
        IEnumerable<IXLTableField> Fields { get; }
        string Name { get; set; }
        Boolean ShowAutoFilter { get; set; }
        Boolean ShowColumnStripes { get; set; }
        Boolean ShowHeaderRow { get; set; }
        Boolean ShowRowStripes { get; set; }
        Boolean ShowTotalsRow { get; set; }
        XLTableTheme Theme { get; set; }

        /// <summary>
        /// Clears the contents of this table.
        /// </summary>
        /// <param name="clearOptions">Specify what you want to clear.</param>
        new IXLTable Clear(XLClearOptions clearOptions = XLClearOptions.ContentsAndFormats);

        IXLTableField Field(string fieldName);

        IXLTableField Field(int fieldIndex);

        IXLRangeRow HeadersRow();

        /// <summary>
        /// Resizes the table to the specified range.
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

        new IXLBaseAutoFilter SetAutoFilter();

        IXLTable SetEmphasizeFirstColumn();

        IXLTable SetEmphasizeFirstColumn(Boolean value);

        IXLTable SetEmphasizeLastColumn();

        IXLTable SetEmphasizeLastColumn(Boolean value);

        IXLTable SetShowAutoFilter();

        IXLTable SetShowAutoFilter(Boolean value);

        IXLTable SetShowColumnStripes();

        IXLTable SetShowColumnStripes(Boolean value);

        IXLTable SetShowHeaderRow();

        IXLTable SetShowHeaderRow(Boolean value);

        IXLTable SetShowRowStripes();

        IXLTable SetShowRowStripes(Boolean value);

        IXLTable SetShowTotalsRow();

        IXLTable SetShowTotalsRow(Boolean value);

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
    }
}
#nullable disable

// Keep this file CodeMaid organised and cleaned
using System;

namespace ClosedXML.Excel
{
    public interface IXLTableRange : IXLRange
    {
        IXLTable Table { get; }

        IXLTableRow FirstRow(Func<IXLTableRow, Boolean> predicate = null);

        IXLTableRow FirstRowUsed(XLCellsUsedOptions options, Func<IXLTableRow, Boolean> predicate = null);

        IXLTableRow FirstRowUsed(Func<IXLTableRow, Boolean> predicate = null);

        new IXLTableRows InsertRowsAbove(int numberOfRows);

        new IXLTableRows InsertRowsBelow(int numberOfRows);

        IXLTableRow LastRow(Func<IXLTableRow, Boolean> predicate = null);

        IXLTableRow LastRowUsed(XLCellsUsedOptions options, Func<IXLTableRow, Boolean> predicate = null);

        IXLTableRow LastRowUsed(Func<IXLTableRow, Boolean> predicate = null);

        /// <summary>
        /// Rows the specified row.
        /// </summary>
        /// <param name="row">1-based row number relative to the first row of this range.</param>
        new IXLTableRow Row(int row);

        IXLTableRows Rows(Func<IXLTableRow, Boolean> predicate = null);

        /// <summary>
        /// Returns a subset of the rows
        /// </summary>
        /// <param name="firstRow">The first row to return. 1-based row number relative to the first row of this range.</param>
        /// <param name="lastRow">The last row to return. 1-based row number relative to the first row of this range.</param>
        new IXLTableRows Rows(int firstRow, int lastRow);

        new IXLTableRows Rows(string rows);

        IXLTableRows RowsUsed(XLCellsUsedOptions options, Func<IXLTableRow, Boolean> predicate = null);

        IXLTableRows RowsUsed(Func<IXLTableRow, Boolean> predicate = null);
    }
}

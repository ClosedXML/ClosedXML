using System;

namespace ClosedXML.Excel
{
    public enum XLShiftDeletedCells { ShiftCellsUp, ShiftCellsLeft }

    public enum XLTransposeOptions { MoveCells, ReplaceCells }

    public enum XLSearchContents { Values, Formulas, ValuesAndFormulas }

    public interface IXLRange : IXLRangeBase
    {
        /// <summary>
        /// Gets the cell at the specified row and column.
        /// <para>The cell address is relative to the parent range.</para>
        /// </summary>
        /// <param name="row">The cell's row.</param>
        /// <param name="column">The cell's column.</param>
        IXLCell Cell(int row, int column);

        /// <summary>Gets the cell at the specified address.</summary>
        /// <para>The cell address is relative to the parent range.</para>
        /// <param name="cellAddressInRange">The cell address in the parent range.</param>
        IXLCell Cell(string cellAddressInRange);

        /// <summary>
        /// Gets the cell at the specified row and column.
        /// <para>The cell address is relative to the parent range.</para>
        /// </summary>
        /// <param name="row">The cell's row.</param>
        /// <param name="column">The cell's column.</param>
        IXLCell Cell(int row, string column);

        /// <summary>Gets the cell at the specified address.</summary>
        /// <para>The cell address is relative to the parent range.</para>
        /// <param name="cellAddressInRange">The cell address in the parent range.</param>
        IXLCell Cell(IXLAddress cellAddressInRange);

        /// <summary>
        /// Gets the specified column of the range.
        /// </summary>
        /// <param name="columnNumber">1-based column number relative to the first column of this range.</param>
        /// <returns>The relevant column</returns>
        IXLRangeColumn Column(int columnNumber);

        /// <summary>
        /// Gets the specified column of the range.
        /// </summary>
        /// <param name="columnLetter">Column letter.</param>
        IXLRangeColumn Column(string columnLetter);

        /// <summary>
        /// Gets the first column of the range.
        /// </summary>
        IXLRangeColumn FirstColumn(Func<IXLRangeColumn, Boolean> predicate = null);

        /// <summary>
        /// Gets the first column of the range that contains a cell with a value.
        /// </summary>
        [Obsolete("Use the overload with XLCellsUsedOptions")]
        IXLRangeColumn FirstColumnUsed(Boolean includeFormats, Func<IXLRangeColumn, Boolean> predicate = null);

        IXLRangeColumn FirstColumnUsed(XLCellsUsedOptions options, Func<IXLRangeColumn, Boolean> predicate = null);

        IXLRangeColumn FirstColumnUsed(Func<IXLRangeColumn, Boolean> predicate = null);

        /// <summary>
        /// Gets the last column of the range.
        /// </summary>
        IXLRangeColumn LastColumn(Func<IXLRangeColumn, Boolean> predicate = null);

        /// <summary>
        /// Gets the last column of the range that contains a cell with a value.
        /// </summary>
        [Obsolete("Use the overload with XLCellsUsedOptions")]
        IXLRangeColumn LastColumnUsed(Boolean includeFormats, Func<IXLRangeColumn, Boolean> predicate = null);

        IXLRangeColumn LastColumnUsed(XLCellsUsedOptions options, Func<IXLRangeColumn, Boolean> predicate = null);

        IXLRangeColumn LastColumnUsed(Func<IXLRangeColumn, Boolean> predicate = null);

        /// <summary>
        /// Gets a collection of all columns in this range.
        /// </summary>
        IXLRangeColumns Columns(Func<IXLRangeColumn, Boolean> predicate = null);

        /// <summary>
        /// Gets a collection of the specified columns in this range.
        /// </summary>
        /// <param name="firstColumn">The first column to return. 1-based column number relative to the first column of this range.</param>
        /// <param name="lastColumn">The last column to return. 1-based column number relative to the first column of this range.</param>
        IXLRangeColumns Columns(int firstColumn, int lastColumn);

        /// <summary>
        /// Gets a collection of the specified columns in this range.
        /// </summary>
        /// <param name="firstColumn">The first column to return.</param>
        /// <param name="lastColumn">The last column to return.</param>
        /// <returns>The relevant columns</returns>
        IXLRangeColumns Columns(string firstColumn, string lastColumn);

        /// <summary>
        /// Gets a collection of the specified columns in this range, separated by commas.
        /// <para>e.g. Columns("G:H"), Columns("10:11,13:14"), Columns("P:Q,S:T"), Columns("V")</para>
        /// </summary>
        /// <param name="columns">The columns to return.</param>
        IXLRangeColumns Columns(string columns);

        /// <summary>
        /// Returns the first row that matches the given predicate
        /// </summary>
        IXLRangeColumn FindColumn(Func<IXLRangeColumn, Boolean> predicate);

        /// <summary>
        /// Returns the first row that matches the given predicate
        /// </summary>
        IXLRangeRow FindRow(Func<IXLRangeRow, Boolean> predicate);

        /// <summary>
        /// Gets the first row of the range.
        /// </summary>
        IXLRangeRow FirstRow(Func<IXLRangeRow, Boolean> predicate = null);

        /// <summary>
        /// Gets the first row of the range that contains a cell with a value.
        /// </summary>
        [Obsolete("Use the overload with XLCellsUsedOptions")]
        IXLRangeRow FirstRowUsed(Boolean includeFormats, Func<IXLRangeRow, Boolean> predicate = null);

        IXLRangeRow FirstRowUsed(XLCellsUsedOptions options, Func<IXLRangeRow, Boolean> predicate = null);

        IXLRangeRow FirstRowUsed(Func<IXLRangeRow, Boolean> predicate = null);

        /// <summary>
        /// Gets the last row of the range.
        /// </summary>
        IXLRangeRow LastRow(Func<IXLRangeRow, Boolean> predicate = null);

        /// <summary>
        /// Gets the last row of the range that contains a cell with a value.
        /// </summary>
        [Obsolete("Use the overload with XLCellsUsedOptions")]
        IXLRangeRow LastRowUsed(Boolean includeFormats, Func<IXLRangeRow, Boolean> predicate = null);

        IXLRangeRow LastRowUsed(XLCellsUsedOptions options, Func<IXLRangeRow, Boolean> predicate = null);

        IXLRangeRow LastRowUsed(Func<IXLRangeRow, Boolean> predicate = null);

        /// <summary>
        /// Gets the specified row of the range.
        /// </summary>
        /// <param name="row">1-based row number relative to the first row of this range.</param>
        /// <returns>The relevant row</returns>
        IXLRangeRow Row(int row);

        IXLRangeRows Rows(Func<IXLRangeRow, Boolean> predicate = null);

        /// <summary>
        /// Gets a collection of the specified rows in this range.
        /// </summary>
        /// <param name="firstRow">The first row to return. 1-based row number relative to the first row of this range.</param>
        /// <param name="lastRow">The last row to return. 1-based row number relative to the first row of this range.</param>
        /// <returns></returns>
        IXLRangeRows Rows(int firstRow, int lastRow);

        /// <summary>
        /// Gets a collection of the specified rows in this range, separated by commas.
        /// <para>e.g. Rows("4:5"), Rows("7:8,10:11"), Rows("13")</para>
        /// </summary>
        /// <param name="rows">The rows to return.</param>
        IXLRangeRows Rows(string rows);

        /// <summary>
        /// Returns the specified range.
        /// </summary>
        /// <param name="rangeAddress">The range boundaries.</param>
        IXLRange Range(IXLRangeAddress rangeAddress);

        /// <summary>Returns the specified range.</summary>
        /// <para>e.g. Range("A1"), Range("A1:C2")</para>
        /// <param name="rangeAddress">The range boundaries.</param>
        IXLRange Range(string rangeAddress);

        /// <summary>Returns the specified range.</summary>
        /// <param name="firstCell">The first cell in the range.</param>
        /// <param name="lastCell"> The last cell in the range.</param>
        IXLRange Range(IXLCell firstCell, IXLCell lastCell);

        /// <summary>Returns the specified range.</summary>
        /// <param name="firstCellAddress">The first cell address in the range.</param>
        /// <param name="lastCellAddress"> The last cell address in the range.</param>
        IXLRange Range(string firstCellAddress, string lastCellAddress);

        /// <summary>Returns the specified range.</summary>
        /// <param name="firstCellAddress">The first cell address in the range.</param>
        /// <param name="lastCellAddress"> The last cell address in the range.</param>
        IXLRange Range(IXLAddress firstCellAddress, IXLAddress lastCellAddress);

        /// <summary>Returns a collection of ranges, separated by commas.</summary>
        /// <para>e.g. Ranges("A1"), Ranges("A1:C2"), Ranges("A1:B2,D1:D4")</para>
        /// <param name="ranges">The ranges to return.</param>
        IXLRanges Ranges(string ranges);

        /// <summary>Returns the specified range.</summary>
        /// <param name="firstCellRow">   The first cell's row of the range to return.</param>
        /// <param name="firstCellColumn">The first cell's column of the range to return.</param>
        /// <param name="lastCellRow">    The last cell's row of the range to return.</param>
        /// <param name="lastCellColumn"> The last cell's column of the range to return.</param>
        /// <returns>.</returns>
        IXLRange Range(int firstCellRow, int firstCellColumn, int lastCellRow, int lastCellColumn);

        /// <summary>Gets the number of rows in this range.</summary>
        int RowCount();

        /// <summary>Gets the number of columns in this range.</summary>
        int ColumnCount();

        /// <summary>
        /// Inserts X number of columns to the right of this range.
        /// <para>All cells to the right of this range will be shifted X number of columns.</para>
        /// </summary>
        /// <param name="numberOfColumns">Number of columns to insert.</param>
        IXLRangeColumns InsertColumnsAfter(int numberOfColumns);

        IXLRangeColumns InsertColumnsAfter(int numberOfColumns, Boolean expandRange);

        /// <summary>
        /// Inserts X number of columns to the left of this range.
        /// <para>This range and all cells to the right of this range will be shifted X number of columns.</para>
        /// </summary>
        /// <param name="numberOfColumns">Number of columns to insert.</param>
        IXLRangeColumns InsertColumnsBefore(int numberOfColumns);

        IXLRangeColumns InsertColumnsBefore(int numberOfColumns, Boolean expandRange);

        /// <summary>
        /// Inserts X number of rows on top of this range.
        /// <para>This range and all cells below this range will be shifted X number of rows.</para>
        /// </summary>
        /// <param name="numberOfRows">Number of rows to insert.</param>
        IXLRangeRows InsertRowsAbove(int numberOfRows);

        IXLRangeRows InsertRowsAbove(int numberOfRows, Boolean expandRange);

        /// <summary>
        /// Inserts X number of rows below this range.
        /// <para>All cells below this range will be shifted X number of rows.</para>
        /// </summary>
        /// <param name="numberOfRows">Number of rows to insert.</param>
        IXLRangeRows InsertRowsBelow(int numberOfRows);

        IXLRangeRows InsertRowsBelow(int numberOfRows, Boolean expandRange);

        /// <summary>
        /// Deletes this range and shifts the surrounding cells accordingly.
        /// </summary>
        /// <param name="shiftDeleteCells">How to shift the surrounding cells.</param>
        void Delete(XLShiftDeletedCells shiftDeleteCells);

        /// <summary>
        /// Transposes the contents and styles of all cells in this range.
        /// </summary>
        /// <param name="transposeOption">How to handle the surrounding cells when transposing the range.</param>
        void Transpose(XLTransposeOptions transposeOption);

        IXLTable AsTable();

        IXLTable AsTable(String name);

        IXLTable CreateTable();

        IXLTable CreateTable(String name);

        IXLRange RangeUsed();

        IXLRange CopyTo(IXLCell target);

        IXLRange CopyTo(IXLRangeBase target);

        IXLSortElements SortRows { get; }
        IXLSortElements SortColumns { get; }

        IXLRange Sort();

        IXLRange Sort(String columnsToSortBy, XLSortOrder sortOrder = XLSortOrder.Ascending, Boolean matchCase = false, Boolean ignoreBlanks = true);

        IXLRange Sort(Int32 columnToSortBy, XLSortOrder sortOrder = XLSortOrder.Ascending, Boolean matchCase = false, Boolean ignoreBlanks = true);

        IXLRange SortLeftToRight(XLSortOrder sortOrder = XLSortOrder.Ascending, Boolean matchCase = false, Boolean ignoreBlanks = true);

        IXLRange SetDataType(XLDataType dataType);

        /// <summary>
        /// Clears the contents of this range.
        /// </summary>
        /// <param name="clearOptions">Specify what you want to clear.</param>
        new IXLRange Clear(XLClearOptions clearOptions = XLClearOptions.All);

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        IXLRangeRows RowsUsed(Boolean includeFormats, Func<IXLRangeRow, Boolean> predicate = null);

        IXLRangeRows RowsUsed(XLCellsUsedOptions options, Func<IXLRangeRow, Boolean> predicate = null);

        IXLRangeRows RowsUsed(Func<IXLRangeRow, Boolean> predicate = null);

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        IXLRangeColumns ColumnsUsed(Boolean includeFormats, Func<IXLRangeColumn, Boolean> predicate = null);

        IXLRangeColumns ColumnsUsed(XLCellsUsedOptions options, Func<IXLRangeColumn, Boolean> predicate = null);

        IXLRangeColumns ColumnsUsed(Func<IXLRangeColumn, Boolean> predicate = null);
    }
}

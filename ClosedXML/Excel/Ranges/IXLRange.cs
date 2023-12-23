#nullable disable

using System;

namespace ClosedXML.Excel
{
    public enum XLShiftDeletedCells { ShiftCellsUp, ShiftCellsLeft }

    /// <summary>
    /// A behavior of extra outside cells for transpose operation. The option
    /// is meaningful only for transposition of non-squared ranges, because
    /// squared ranges can always be transposed without effecting outside cells. 
    /// </summary>
    public enum XLTransposeOptions
    {
        /// <summary>
        /// Shift cells of the smaller side to its direction so
        /// there is a space to transpose other side (e.g. if A1:C5
        /// range is transposed, move D1:XFD5 are moved 2 columns
        /// to the right).
        /// </summary>
        MoveCells,

        /// <summary>
        /// Data of the cells are replaced by the transposed cells.
        /// </summary>
        ReplaceCells
    }

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
        /// Gets the first non-empty column of the range that contains a cell with a value.
        /// </summary>
        /// <param name="options">The options to determine whether a cell is used.</param>
        /// <param name="predicate">The predicate to choose cells.</param>
        IXLRangeColumn FirstColumnUsed(XLCellsUsedOptions options, Func<IXLRangeColumn, Boolean> predicate = null);

        IXLRangeColumn FirstColumnUsed(Func<IXLRangeColumn, Boolean> predicate = null);

        /// <summary>
        /// Gets the last column of the range.
        /// </summary>
        IXLRangeColumn LastColumn(Func<IXLRangeColumn, Boolean> predicate = null);

        /// <summary>
        /// Gets the last non-empty column of the range that contains a cell with a value.
        /// </summary>
        /// <param name="options">The options to determine whether a cell is used.</param>
        /// <param name="predicate">The predicate to choose cells.</param>
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
        /// Gets the first non-empty row of the range that contains a cell with a value.
        /// </summary>
        /// <param name="options">The options to determine whether a cell is used.</param>
        /// <param name="predicate">The predicate to choose cells.</param>
        IXLRangeRow FirstRowUsed(XLCellsUsedOptions options, Func<IXLRangeRow, Boolean> predicate = null);

        IXLRangeRow FirstRowUsed(Func<IXLRangeRow, Boolean> predicate = null);

        /// <summary>
        /// Gets the last row of the range.
        /// </summary>
        IXLRangeRow LastRow(Func<IXLRangeRow, Boolean> predicate = null);

        /// <summary>
        /// Gets the last non-empty row of the range that contains a cell with a value.
        /// </summary>
        /// <param name="options">The options to determine whether a cell is used.</param>
        /// <param name="predicate">The predicate to choose cells.</param>
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

        /// <summary>
        /// Use this range as a table, but do not add it to the Tables list
        /// </summary>
        /// <remarks>
        /// NOTES:<br/>
        ///     The AsTable method will use the first row of the range as a header row.<br/>
        ///     If this range contains only one row, then an empty data row will be inserted into the returned table.
        /// </remarks>
        IXLTable AsTable();

        /// <summary>
        /// Use this range as a table with the passed name, but do not add it to the Tables list
        /// </summary>
        /// <param name="name">Table name to be used.</param>
        /// <remarks>
        /// NOTES:<br/>
        ///     The AsTable method will use the first row of the range as a header row.<br/>
        ///     If this range contains only one row, then an empty data row will be inserted into the returned table.
        /// </remarks>
        IXLTable AsTable(String name);

        IXLTable CreateTable();

        IXLTable CreateTable(String name);

        IXLRange RangeUsed();

        IXLRange CopyTo(IXLCell target);

        IXLRange CopyTo(IXLRangeBase target);

        /// <summary>
        /// Rows used for sorting columns. Automatically updated each time a <see cref="SortLeftToRight(XLSortOrder, bool, bool)"/>
        /// is called.
        /// </summary>
        IXLSortElements SortRows { get; }

        /// <summary>
        /// Columns used for sorting rows. Automatically updated each time a <see cref="Sort(String, XLSortOrder, bool, bool)"/>
        /// or <see cref="Sort(Int32, XLSortOrder, bool, bool)"/>.
        /// </summary>
        /// <remarks>
        /// User can set desired sorting order here and then call <see cref="Sort()"/> method.
        /// </remarks>
        IXLSortElements SortColumns { get; }

        /// <summary>
        /// Sort rows of the range using the <see cref="SortColumns"/> (if non-empty) or by using
        /// all columns of the range in ascending order.
        /// </summary>
        /// <remarks>
        /// This method can be used fort sorting, after user specified desired sorting order
        /// in <see cref="SortColumns"/>.
        /// </remarks>
        /// <returns>This range.</returns>
        IXLRange Sort();

        /// <summary>
        /// Sort rows of the range according to values in columns specified by <paramref name="columnsToSortBy"/>.
        /// </summary>
        /// <param name="columnsToSortBy">
        /// Columns which should be used to sort the range and their order. Columns are separated
        /// by a comma (<strong>,</strong>). The column can be specified either by column number or
        /// by column letter. Sort order is parsed case insensitive and can be <c>ASC</c> or
        /// <c>DESC</c>. The specified column is relative to the origin of the range.
        /// <para>
        /// <example><c>2 DESC, 1, C asc</c> means sort by second column of a range in descending
        /// order, then by first column of a range in <paramref name="sortOrder"/> and then by
        /// column <c>C</c> in ascending order.</example>.
        /// </para>
        /// </param>
        /// <param name="sortOrder">
        /// What should be the default sorting order or columns in <paramref name="columnsToSortBy"/>
        /// without specified sorting order.
        /// </param>
        /// <param name="matchCase">
        /// When cell value is a <see cref="XLDataType.Text"/>, should sorting be case insensitive
        /// (<c>false</c>, Excel default behavior) or case sensitive (<c>true</c>). Doesn't affect
        /// other cell value types.
        /// </param>
        /// <param name="ignoreBlanks">
        /// When <c>true</c> (recommended, matches Excel behavior), blank cell values are always
        /// sorted at the end regardless of sorting order. When <c>false</c>, blank values are
        /// considered empty strings and are sorted among other cell values with a type
        /// <see cref="XLDataType.Text"/>.
        /// </param>
        /// <returns>This range.</returns>
        IXLRange Sort(String columnsToSortBy, XLSortOrder sortOrder = XLSortOrder.Ascending, Boolean matchCase = false, Boolean ignoreBlanks = true);

        /// <summary>
        /// Sort rows of the range according to values in <paramref name="columnToSortBy"/> column.
        /// </summary>
        /// <param name="columnToSortBy">Column number that will be used to sort the range rows.</param>
        /// <param name="sortOrder">Sorting order used by <paramref name="columnToSortBy"/>.</param>
        /// <param name="matchCase"><inheritdoc cref="Sort(String, XLSortOrder, bool, bool)"/></param>
        /// <param name="ignoreBlanks"><inheritdoc cref="Sort(String, XLSortOrder, bool, bool)"/></param>
        /// <returns>This range.</returns>
        IXLRange Sort(Int32 columnToSortBy, XLSortOrder sortOrder = XLSortOrder.Ascending, Boolean matchCase = false, Boolean ignoreBlanks = true);

        /// <summary>
        /// Sort columns in a range. The sorting is done using the values in each column of the range.
        /// </summary>
        /// <param name="sortOrder">In what order should columns be sorted</param>
        /// <param name="matchCase"><inheritdoc cref="Sort(String, XLSortOrder, bool, bool)"/></param>
        /// <param name="ignoreBlanks"><inheritdoc cref="Sort(String, XLSortOrder, bool, bool)"/></param>
        /// <returns>This range.</returns>
        IXLRange SortLeftToRight(XLSortOrder sortOrder = XLSortOrder.Ascending, Boolean matchCase = false, Boolean ignoreBlanks = true);

        /// <summary>
        /// Clears the contents of this range.
        /// </summary>
        /// <param name="clearOptions">Specify what you want to clear.</param>
        new IXLRange Clear(XLClearOptions clearOptions = XLClearOptions.All);

        IXLRangeRows RowsUsed(XLCellsUsedOptions options, Func<IXLRangeRow, Boolean> predicate = null);

        IXLRangeRows RowsUsed(Func<IXLRangeRow, Boolean> predicate = null);

        IXLRangeColumns ColumnsUsed(XLCellsUsedOptions options, Func<IXLRangeColumn, Boolean> predicate = null);

        IXLRangeColumns ColumnsUsed(Func<IXLRangeColumn, Boolean> predicate = null);
    }
}

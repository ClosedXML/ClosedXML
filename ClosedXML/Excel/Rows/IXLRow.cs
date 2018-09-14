using System;

namespace ClosedXML.Excel
{
    public interface IXLRow : IXLRangeBase
    {
        /// <summary>
        /// Gets or sets the height of this row.
        /// </summary>
        /// <value>
        /// The width of this row.
        /// </value>
        Double Height { get; set; }

        /// <summary>
        /// Clears the height for the row and defaults it to the spreadsheet row height.
        /// </summary>
        void ClearHeight();

        /// <summary>
        /// Deletes this row and shifts the rows below this one accordingly.
        /// </summary>
        void Delete();

        /// <summary>
        /// Gets this row's number
        /// </summary>
        Int32 RowNumber();

        /// <summary>
        /// Inserts X number of rows below this one.
        /// <para>All rows below will be shifted accordingly.</para>
        /// </summary>
        /// <param name="numberOfRows">The number of rows to insert.</param>
        IXLRows InsertRowsBelow(Int32 numberOfRows);

        /// <summary>
        /// Inserts X number of rows above this one.
        /// <para>This row and all below will be shifted accordingly.</para>
        /// </summary>
        /// <param name="numberOfRows">The number of rows to insert.</param>
        IXLRows InsertRowsAbove(Int32 numberOfRows);

        IXLRow AdjustToContents();

        /// <summary>
        /// Adjusts the height of the row based on its contents, starting from the startColumn.
        /// </summary>
        /// <param name="startColumn">The column to start calculating the row height.</param>
        IXLRow AdjustToContents(Int32 startColumn);

        /// <summary>
        /// Adjusts the height of the row based on its contents, starting from the startColumn and ending at endColumn.
        /// </summary>
        /// <param name="startColumn">The column to start calculating the row height.</param>
        /// <param name="endColumn">The column to end calculating the row height.</param>
        IXLRow AdjustToContents(Int32 startColumn, Int32 endColumn);

        IXLRow AdjustToContents(Double minHeight, Double maxHeight);

        IXLRow AdjustToContents(Int32 startColumn, Double minHeight, Double maxHeight);

        IXLRow AdjustToContents(Int32 startColumn, Int32 endColumn, Double minHeight, Double maxHeight);

        /// <summary>Hides this row.</summary>
        IXLRow Hide();

        /// <summary>Unhides this row.</summary>
        IXLRow Unhide();

        /// <summary>
        /// Gets a value indicating whether this row is hidden or not.
        /// </summary>
        /// <value>
        ///   <c>true</c> if this row is hidden; otherwise, <c>false</c>.
        /// </value>
        Boolean IsHidden { get; }

        /// <summary>
        /// Gets or sets the outline level of this row.
        /// </summary>
        /// <value>
        /// The outline level of this row.
        /// </value>
        Int32 OutlineLevel { get; set; }

        /// <summary>
        /// Adds this row to the next outline level (Increments the outline level for this row by 1).
        /// </summary>
        IXLRow Group();

        /// <summary>
        /// Adds this row to the next outline level (Increments the outline level for this row by 1).
        /// </summary>
        /// <param name="collapse">If set to <c>true</c> the row will be shown collapsed.</param>
        IXLRow Group(Boolean collapse);

        /// <summary>
        /// Sets outline level for this row.
        /// </summary>
        /// <param name="outlineLevel">The outline level.</param>
        IXLRow Group(Int32 outlineLevel);

        /// <summary>
        /// Sets outline level for this row.
        /// </summary>
        /// <param name="outlineLevel">The outline level.</param>
        /// <param name="collapse">If set to <c>true</c> the row will be shown collapsed.</param>
        IXLRow Group(Int32 outlineLevel, Boolean collapse);

        /// <summary>
        /// Adds this row to the previous outline level (decrements the outline level for this row by 1).
        /// </summary>
        IXLRow Ungroup();

        /// <summary>
        /// Adds this row to the previous outline level (decrements the outline level for this row by 1).
        /// </summary>
        /// <param name="fromAll">If set to <c>true</c> it will remove this row from all outline levels.</param>
        IXLRow Ungroup(Boolean fromAll);

        /// <summary>
        /// Show this row as collapsed.
        /// </summary>
        IXLRow Collapse();

        /// <summary>
        /// Gets the cell in the specified column.
        /// </summary>
        /// <param name="columnNumber">The cell's column.</param>
        IXLCell Cell(Int32 columnNumber);

        /// <summary>
        /// Gets the cell in the specified column.
        /// </summary>
        /// <param name="columnLetter">The cell's column.</param>
        IXLCell Cell(String columnLetter);

        /// <summary>
        /// Returns the specified group of cells, separated by commas.
        /// <para>e.g. Cells("1"), Cells("1:5"), Cells("1,3:5")</para>
        /// </summary>
        /// <param name="cellsInRow">The row's cells to return.</param>
        new IXLCells Cells(String cellsInRow);

        /// <summary>
        /// Returns the specified group of cells.
        /// </summary>
        /// <param name="firstColumn">The first column in the group of cells to return.</param>
        /// <param name="lastColumn">The last column in the group of cells to return.</param>
        IXLCells Cells(Int32 firstColumn, Int32 lastColumn);

        /// <summary>
        /// Returns the specified group of cells.
        /// </summary>
        /// <param name="firstColumn">The first column in the group of cells to return.</param>
        /// <param name="lastColumn">The last column in the group of cells to return.</param>
        IXLCells Cells(String firstColumn, String lastColumn);

        /// <summary>Expands this row (if it's collapsed).</summary>
        IXLRow Expand();

        Int32 CellCount();

        IXLRangeRow CopyTo(IXLCell cell);

        IXLRangeRow CopyTo(IXLRangeBase range);

        IXLRow CopyTo(IXLRow row);

        IXLRow Sort();

        IXLRow SortLeftToRight(XLSortOrder sortOrder = XLSortOrder.Ascending, Boolean matchCase = false, Boolean ignoreBlanks = true);

        IXLRangeRow Row(Int32 start, Int32 end);

        IXLRangeRow Row(IXLCell start, IXLCell end);

        IXLRangeRows Rows(String columns);

        /// <summary>
        /// Adds a horizontal page break after this row.
        /// </summary>
        IXLRow AddHorizontalPageBreak();

        IXLRow SetDataType(XLDataType dataType);

        IXLRow RowAbove();

        IXLRow RowAbove(Int32 step);

        IXLRow RowBelow();

        IXLRow RowBelow(Int32 step);

        /// <summary>
        /// Clears the contents of this row.
        /// </summary>
        /// <param name="clearOptions">Specify what you want to clear.</param>
        new IXLRow Clear(XLClearOptions clearOptions = XLClearOptions.All);

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        IXLRangeRow RowUsed(Boolean includeFormats);

        IXLRangeRow RowUsed(XLCellsUsedOptions options = XLCellsUsedOptions.AllContents);
    }
}

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
        double Height { get; set; }

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
        int RowNumber();

        /// <summary>
        /// Inserts X number of rows below this one.
        /// <para>All rows below will be shifted accordingly.</para>
        /// </summary>
        /// <param name="numberOfRows">The number of rows to insert.</param>
        IXLRows InsertRowsBelow(int numberOfRows);

        /// <summary>
        /// Inserts X number of rows above this one.
        /// <para>This row and all below will be shifted accordingly.</para>
        /// </summary>
        /// <param name="numberOfRows">The number of rows to insert.</param>
        IXLRows InsertRowsAbove(int numberOfRows);

        IXLRow AdjustToContents();

        /// <summary>
        /// Adjusts the height of the row based on its contents, starting from the startColumn.
        /// </summary>
        /// <param name="startColumn">The column to start calculating the row height.</param>
        IXLRow AdjustToContents(int startColumn);

        /// <summary>
        /// Adjusts the height of the row based on its contents, starting from the startColumn and ending at endColumn.
        /// </summary>
        /// <param name="startColumn">The column to start calculating the row height.</param>
        /// <param name="endColumn">The column to end calculating the row height.</param>
        IXLRow AdjustToContents(int startColumn, int endColumn);

        IXLRow AdjustToContents(double minHeight, double maxHeight);

        IXLRow AdjustToContents(int startColumn, double minHeight, double maxHeight);

        IXLRow AdjustToContents(int startColumn, int endColumn, double minHeight, double maxHeight);

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
        bool IsHidden { get; }

        /// <summary>
        /// Gets or sets the outline level of this row.
        /// </summary>
        /// <value>
        /// The outline level of this row.
        /// </value>
        int OutlineLevel { get; set; }

        /// <summary>
        /// Adds this row to the next outline level (Increments the outline level for this row by 1).
        /// </summary>
        IXLRow Group();

        /// <summary>
        /// Adds this row to the next outline level (Increments the outline level for this row by 1).
        /// </summary>
        /// <param name="collapse">If set to <c>true</c> the row will be shown collapsed.</param>
        IXLRow Group(bool collapse);

        /// <summary>
        /// Sets outline level for this row.
        /// </summary>
        /// <param name="outlineLevel">The outline level.</param>
        IXLRow Group(int outlineLevel);

        /// <summary>
        /// Sets outline level for this row.
        /// </summary>
        /// <param name="outlineLevel">The outline level.</param>
        /// <param name="collapse">If set to <c>true</c> the row will be shown collapsed.</param>
        IXLRow Group(int outlineLevel, bool collapse);

        /// <summary>
        /// Adds this row to the previous outline level (decrements the outline level for this row by 1).
        /// </summary>
        IXLRow Ungroup();

        /// <summary>
        /// Adds this row to the previous outline level (decrements the outline level for this row by 1).
        /// </summary>
        /// <param name="fromAll">If set to <c>true</c> it will remove this row from all outline levels.</param>
        IXLRow Ungroup(bool fromAll);

        /// <summary>
        /// Show this row as collapsed.
        /// </summary>
        IXLRow Collapse();

        /// <summary>
        /// Gets the cell in the specified column.
        /// </summary>
        /// <param name="columnNumber">The cell's column.</param>
        IXLCell Cell(int columnNumber);

        /// <summary>
        /// Gets the cell in the specified column.
        /// </summary>
        /// <param name="columnLetter">The cell's column.</param>
        IXLCell Cell(string columnLetter);

        /// <summary>
        /// Returns the specified group of cells, separated by commas.
        /// <para>e.g. Cells("1"), Cells("1:5"), Cells("1,3:5")</para>
        /// </summary>
        /// <param name="cellsInRow">The row's cells to return.</param>
        new IXLCells Cells(string cellsInRow);

        /// <summary>
        /// Returns the specified group of cells.
        /// </summary>
        /// <param name="firstColumn">The first column in the group of cells to return.</param>
        /// <param name="lastColumn">The last column in the group of cells to return.</param>
        IXLCells Cells(int firstColumn, int lastColumn);

        /// <summary>
        /// Returns the specified group of cells.
        /// </summary>
        /// <param name="firstColumn">The first column in the group of cells to return.</param>
        /// <param name="lastColumn">The last column in the group of cells to return.</param>
        IXLCells Cells(string firstColumn, string lastColumn);

        /// <summary>Expands this row (if it's collapsed).</summary>
        IXLRow Expand();

        int CellCount();

        IXLRangeRow CopyTo(IXLCell cell);

        IXLRangeRow CopyTo(IXLRangeBase range);

        IXLRow CopyTo(IXLRow row);

        IXLRow Sort();

        IXLRow SortLeftToRight(XLSortOrder sortOrder = XLSortOrder.Ascending, bool matchCase = false, bool ignoreBlanks = true);

        IXLRangeRow Row(int start, int end);

        IXLRangeRow Row(IXLCell start, IXLCell end);

        IXLRangeRows Rows(string columns);

        /// <summary>
        /// Adds a horizontal page break after this row.
        /// </summary>
        IXLRow AddHorizontalPageBreak();

        IXLRow SetDataType(XLDataType dataType);

        IXLRow RowAbove();

        IXLRow RowAbove(int step);

        IXLRow RowBelow();

        IXLRow RowBelow(int step);

        /// <summary>
        /// Clears the contents of this row.
        /// </summary>
        /// <param name="clearOptions">Specify what you want to clear.</param>
        new IXLRow Clear(XLClearOptions clearOptions = XLClearOptions.All);

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        IXLRangeRow RowUsed(bool includeFormats);

        IXLRangeRow RowUsed(XLCellsUsedOptions options = XLCellsUsedOptions.AllContents);
    }
}

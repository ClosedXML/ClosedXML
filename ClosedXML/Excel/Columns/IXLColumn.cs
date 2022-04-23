using System;

namespace ClosedXML.Excel
{
    public interface IXLColumn : IXLRangeBase
    {
        /// <summary>
        /// Gets or sets the width of this column.
        /// </summary>
        /// <value>
        /// The width of this column.
        /// </value>
        double Width { get; set; }

        /// <summary>
        /// Deletes this column and shifts the columns at the right of this one accordingly.
        /// </summary>
        void Delete();

        /// <summary>
        /// Gets this column's number
        /// </summary>
        int ColumnNumber();

        /// <summary>
        /// Gets this column's letter
        /// </summary>
        string ColumnLetter();

        /// <summary>
        /// Inserts X number of columns at the right of this one.
        /// <para>All columns at the right will be shifted accordingly.</para>
        /// </summary>
        /// <param name="numberOfColumns">The number of columns to insert.</param>
        IXLColumns InsertColumnsAfter(int numberOfColumns);

        /// <summary>
        /// Inserts X number of columns at the left of this one.
        /// <para>This column and all at the right will be shifted accordingly.</para>
        /// </summary>
        /// <param name="numberOfColumns">The number of columns to insert.</param>
        IXLColumns InsertColumnsBefore(int numberOfColumns);

        /// <summary>
        /// Gets the cell in the specified row.
        /// </summary>
        /// <param name="rowNumber">The cell's row.</param>
        IXLCell Cell(int rowNumber);

        /// <summary>
        /// Returns the specified group of cells, separated by commas.
        /// <para>e.g. Cells("1"), Cells("1:5"), Cells("1,3:5")</para>
        /// </summary>
        /// <param name="cellsInColumn">The column cells to return.</param>
        new IXLCells Cells(string cellsInColumn);

        /// <summary>
        /// Returns the specified group of cells.
        /// </summary>
        /// <param name="firstRow">The first row in the group of cells to return.</param>
        /// <param name="lastRow">The last row in the group of cells to return.</param>
        IXLCells Cells(int firstRow, int lastRow);

        /// <summary>
        /// Adjusts the width of the column based on its contents.
        /// </summary>
        IXLColumn AdjustToContents();

        /// <summary>
        /// Adjusts the width of the column based on its contents, starting from the startRow.
        /// </summary>
        /// <param name="startRow">The row to start calculating the column width.</param>
        IXLColumn AdjustToContents(int startRow);

        /// <summary>
        /// Adjusts the width of the column based on its contents, starting from the startRow and ending at endRow.
        /// </summary>
        /// <param name="startRow">The row to start calculating the column width.</param>
        /// <param name="endRow">The row to end calculating the column width.</param>
        IXLColumn AdjustToContents(int startRow, int endRow);

        IXLColumn AdjustToContents(double minWidth, double maxWidth);

        IXLColumn AdjustToContents(int startRow, double minWidth, double maxWidth);

        IXLColumn AdjustToContents(int startRow, int endRow, double minWidth, double maxWidth);

        /// <summary>
        /// Hides this column.
        /// </summary>
        IXLColumn Hide();

        /// <summary>Unhides this column.</summary>
        IXLColumn Unhide();

        /// <summary>
        /// Gets a value indicating whether this column is hidden or not.
        /// </summary>
        /// <value>
        ///   <c>true</c> if this column is hidden; otherwise, <c>false</c>.
        /// </value>
        bool IsHidden { get; }

        /// <summary>
        /// Gets or sets the outline level of this column.
        /// </summary>
        /// <value>
        /// The outline level of this column.
        /// </value>
        int OutlineLevel { get; set; }

        /// <summary>
        /// Adds this column to the next outline level (Increments the outline level for this column by 1).
        /// </summary>
        IXLColumn Group();

        /// <summary>
        /// Adds this column to the next outline level (Increments the outline level for this column by 1).
        /// </summary>
        /// <param name="collapse">If set to <c>true</c> the column will be shown collapsed.</param>
        IXLColumn Group(bool collapse);

        /// <summary>
        /// Sets outline level for this column.
        /// </summary>
        /// <param name="outlineLevel">The outline level.</param>
        IXLColumn Group(int outlineLevel);

        /// <summary>
        /// Sets outline level for this column.
        /// </summary>
        /// <param name="outlineLevel">The outline level.</param>
        /// <param name="collapse">If set to <c>true</c> the column will be shown collapsed.</param>
        IXLColumn Group(int outlineLevel, bool collapse);

        /// <summary>
        /// Adds this column to the previous outline level (decrements the outline level for this column by 1).
        /// </summary>
        IXLColumn Ungroup();

        /// <summary>
        /// Adds this column to the previous outline level (decrements the outline level for this column by 1).
        /// </summary>
        /// <param name="fromAll">If set to <c>true</c> it will remove this column from all outline levels.</param>
        IXLColumn Ungroup(bool fromAll);

        /// <summary>
        /// Show this column as collapsed.
        /// </summary>
        IXLColumn Collapse();

        /// <summary>Expands this column (if it's collapsed).</summary>
        IXLColumn Expand();

        int CellCount();

        IXLRangeColumn CopyTo(IXLCell cell);

        IXLRangeColumn CopyTo(IXLRangeBase range);

        IXLColumn CopyTo(IXLColumn column);

        IXLColumn Sort(XLSortOrder sortOrder = XLSortOrder.Ascending, bool matchCase = false, bool ignoreBlanks = true);

        IXLRangeColumn Column(int start, int end);

        IXLRangeColumn Column(IXLCell start, IXLCell end);

        IXLRangeColumns Columns(string columns);

        /// <summary>
        /// Adds a vertical page break after this column.
        /// </summary>
        IXLColumn AddVerticalPageBreak();

        IXLColumn SetDataType(XLDataType dataType);

        IXLColumn ColumnLeft();

        IXLColumn ColumnLeft(int step);

        IXLColumn ColumnRight();

        IXLColumn ColumnRight(int step);

        /// <summary>
        /// Clears the contents of this column.
        /// </summary>
        /// <param name="clearOptions">Specify what you want to clear.</param>
        new IXLColumn Clear(XLClearOptions clearOptions = XLClearOptions.All);

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        IXLRangeColumn ColumnUsed(bool includeFormats);

        IXLRangeColumn ColumnUsed(XLCellsUsedOptions options = XLCellsUsedOptions.AllContents);
    }
}
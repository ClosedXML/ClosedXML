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
        Double Width { get; set; }

        /// <summary>
        /// Deletes this column and shifts the columns at the right of this one accordingly.
        /// </summary>
        void Delete();

        /// <summary>
        /// Gets this column's number
        /// </summary>
        Int32 ColumnNumber();

        /// <summary>
        /// Gets this column's letter
        /// </summary>
        String ColumnLetter();

        /// <summary>
        /// Inserts X number of columns at the right of this one.
        /// <para>All columns at the right will be shifted accordingly.</para>
        /// </summary>
        /// <param name="numberOfColumns">The number of columns to insert.</param>
        IXLColumns InsertColumnsAfter(Int32 numberOfColumns);

        /// <summary>
        /// Inserts X number of columns at the left of this one.
        /// <para>This column and all at the right will be shifted accordingly.</para>
        /// </summary>
        /// <param name="numberOfColumns">The number of columns to insert.</param>
        IXLColumns InsertColumnsBefore(Int32 numberOfColumns);

        /// <summary>
        /// Gets the cell in the specified row.
        /// </summary>
        /// <param name="rowNumber">The cell's row.</param>
        IXLCell Cell(Int32 rowNumber);

        /// <summary>
        /// Returns the specified group of cells, separated by commas.
        /// <para>e.g. Cells("1"), Cells("1:5"), Cells("1,3:5")</para>
        /// </summary>
        /// <param name="cellsInColumn">The column cells to return.</param>
        new IXLCells Cells(String cellsInColumn);

        /// <summary>
        /// Returns the specified group of cells.
        /// </summary>
        /// <param name="firstRow">The first row in the group of cells to return.</param>
        /// <param name="lastRow">The last row in the group of cells to return.</param>
        IXLCells Cells(Int32 firstRow, Int32 lastRow);

        /// <summary>
        /// Adjusts the width of the column based on its contents.
        /// </summary>
        IXLColumn AdjustToContents();

        /// <summary>
        /// Adjusts the width of the column based on its contents, starting from the startRow.
        /// </summary>
        /// <param name="startRow">The row to start calculating the column width.</param>
        IXLColumn AdjustToContents(Int32 startRow);

        /// <summary>
        /// Adjusts the width of the column based on its contents, starting from the startRow and ending at endRow.
        /// </summary>
        /// <param name="startRow">The row to start calculating the column width.</param>
        /// <param name="endRow">The row to end calculating the column width.</param>
        IXLColumn AdjustToContents(Int32 startRow, Int32 endRow);

        IXLColumn AdjustToContents(Double minWidth, Double maxWidth);

        IXLColumn AdjustToContents(Int32 startRow, Double minWidth, Double maxWidth);

        IXLColumn AdjustToContents(Int32 startRow, Int32 endRow, Double minWidth, Double maxWidth);

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
        Boolean IsHidden { get; }

        /// <summary>
        /// Gets or sets the outline level of this column.
        /// </summary>
        /// <value>
        /// The outline level of this column.
        /// </value>
        Int32 OutlineLevel { get; set; }

        /// <summary>
        /// Adds this column to the next outline level (Increments the outline level for this column by 1).
        /// </summary>
        IXLColumn Group();

        /// <summary>
        /// Adds this column to the next outline level (Increments the outline level for this column by 1).
        /// </summary>
        /// <param name="collapse">If set to <c>true</c> the column will be shown collapsed.</param>
        IXLColumn Group(Boolean collapse);

        /// <summary>
        /// Sets outline level for this column.
        /// </summary>
        /// <param name="outlineLevel">The outline level.</param>
        IXLColumn Group(Int32 outlineLevel);

        /// <summary>
        /// Sets outline level for this column.
        /// </summary>
        /// <param name="outlineLevel">The outline level.</param>
        /// <param name="collapse">If set to <c>true</c> the column will be shown collapsed.</param>
        IXLColumn Group(Int32 outlineLevel, Boolean collapse);

        /// <summary>
        /// Adds this column to the previous outline level (decrements the outline level for this column by 1).
        /// </summary>
        IXLColumn Ungroup();

        /// <summary>
        /// Adds this column to the previous outline level (decrements the outline level for this column by 1).
        /// </summary>
        /// <param name="fromAll">If set to <c>true</c> it will remove this column from all outline levels.</param>
        IXLColumn Ungroup(Boolean fromAll);

        /// <summary>
        /// Show this column as collapsed.
        /// </summary>
        IXLColumn Collapse();

        /// <summary>Expands this column (if it's collapsed).</summary>
        IXLColumn Expand();

        Int32 CellCount();

        IXLRangeColumn CopyTo(IXLCell cell);

        IXLRangeColumn CopyTo(IXLRangeBase range);

        IXLColumn CopyTo(IXLColumn column);

        IXLColumn Sort(XLSortOrder sortOrder = XLSortOrder.Ascending, Boolean matchCase = false, Boolean ignoreBlanks = true);

        IXLRangeColumn Column(Int32 start, Int32 end);

        IXLRangeColumn Column(IXLCell start, IXLCell end);

        IXLRangeColumns Columns(String columns);

        /// <summary>
        /// Adds a vertical page break after this column.
        /// </summary>
        IXLColumn AddVerticalPageBreak();

        IXLColumn SetDataType(XLDataType dataType);

        IXLColumn ColumnLeft();

        IXLColumn ColumnLeft(Int32 step);

        IXLColumn ColumnRight();

        IXLColumn ColumnRight(Int32 step);

        /// <summary>
        /// Clears the contents of this column.
        /// </summary>
        /// <param name="clearOptions">Specify what you want to clear.</param>
        new IXLColumn Clear(XLClearOptions clearOptions = XLClearOptions.All);

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        IXLRangeColumn ColumnUsed(Boolean includeFormats);

        IXLRangeColumn ColumnUsed(XLCellsUsedOptions options = XLCellsUsedOptions.AllContents);
    }
}

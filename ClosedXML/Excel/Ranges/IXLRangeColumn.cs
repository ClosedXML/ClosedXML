using System;

namespace ClosedXML.Excel
{
    public interface IXLRangeColumn : IXLRangeBase
    {
        /// <summary>
        /// Gets the cell in the specified row.
        /// </summary>
        /// <param name="rowNumber">The cell's row.</param>
        IXLCell Cell(Int32 rowNumber);

        /// <summary>
        /// Returns the specified group of cells, separated by commas.
        /// <para>e.g. Cells("1"), Cells("1:5"), Cells("1:2,4:5")</para>
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
        /// Inserts X number of cells on top of this column.
        /// <para>This column and all cells below it will be shifted X number of rows.</para>
        /// </summary>
        /// <param name="numberOfRows">Number of cells to insert.</param>
        IXLCells InsertCellsAbove(int numberOfRows);

        IXLCells InsertCellsAbove(int numberOfRows, Boolean expandRange);

        /// <summary>
        /// Inserts X number of cells below this range.
        /// <para>All cells below this column will be shifted X number of rows.</para>
        /// </summary>
        /// <param name="numberOfRows">Number of cells to insert.</param>
        IXLCells InsertCellsBelow(int numberOfRows);

        IXLCells InsertCellsBelow(int numberOfRows, Boolean expandRange);

        /// <summary>
        /// Deletes this range and shifts the cells at the right.
        /// </summary>
        void Delete();

        /// <summary>
        /// Deletes this range and shifts the surrounding cells accordingly.
        /// </summary>
        /// <param name="shiftDeleteCells">How to shift the surrounding cells.</param>
        void Delete(XLShiftDeletedCells shiftDeleteCells);

        /// <summary>
        /// Gets this column's number in the range
        /// </summary>
        Int32 ColumnNumber();

        /// <summary>
        /// Gets this column's letter in the range
        /// </summary>
        String ColumnLetter();

        Int32 CellCount();

        IXLRangeColumn CopyTo(IXLCell target);

        IXLRangeColumn CopyTo(IXLRangeBase target);

        IXLRangeColumn Sort(XLSortOrder sortOrder = XLSortOrder.Ascending, Boolean matchCase = false, Boolean ignoreBlanks = true);

        IXLRangeColumn Column(Int32 start, Int32 end);

        IXLRangeColumn Column(IXLCell start, IXLCell end);

        IXLRangeColumns Columns(String columns);

        IXLRangeColumn SetDataType(XLDataType dataType);

        IXLRangeColumn ColumnLeft();

        IXLRangeColumn ColumnLeft(Int32 step);

        IXLRangeColumn ColumnRight();

        IXLRangeColumn ColumnRight(Int32 step);

        IXLColumn WorksheetColumn();

        IXLTable AsTable();

        IXLTable AsTable(String name);

        IXLTable CreateTable();

        IXLTable CreateTable(String name);

        /// <summary>
        /// Clears the contents of this column.
        /// </summary>
        /// <param name="clearOptions">Specify what you want to clear.</param>
        new IXLRangeColumn Clear(XLClearOptions clearOptions = XLClearOptions.All);

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        IXLRangeColumn ColumnUsed(Boolean includeFormats);

        IXLRangeColumn ColumnUsed(XLCellsUsedOptions options = XLCellsUsedOptions.AllContents);
    }
}

using System;

namespace ClosedXML.Excel
{
    public interface IXLRangeRow : IXLRangeBase
    {
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
        /// <para>e.g. Cells("1"), Cells("1:5"), Cells("1:2,4:5")</para>
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

        /// <summary>
        /// Inserts X number of cells to the right of this row.
        /// <para>All cells to the right of this row will be shifted X number of columns.</para>
        /// </summary>
        /// <param name="numberOfColumns">Number of cells to insert.</param>
        IXLCells InsertCellsAfter(int numberOfColumns);

        IXLCells InsertCellsAfter(int numberOfColumns, Boolean expandRange);

        /// <summary>
        /// Inserts X number of cells to the left of this row.
        /// <para>This row and all cells to the right of it will be shifted X number of columns.</para>
        /// </summary>
        /// <param name="numberOfColumns">Number of cells to insert.</param>
        IXLCells InsertCellsBefore(int numberOfColumns);

        IXLCells InsertCellsBefore(int numberOfColumns, Boolean expandRange);

        /// <summary>
        /// Inserts X number of rows on top of this row.
        /// <para>This row and all cells below it will be shifted X number of rows.</para>
        /// </summary>
        /// <param name="numberOfRows">Number of rows to insert.</param>
        IXLRangeRows InsertRowsAbove(int numberOfRows);

        IXLRangeRows InsertRowsAbove(int numberOfRows, Boolean expandRange);

        /// <summary>
        /// Inserts X number of rows below this row.
        /// <para>All cells below this row will be shifted X number of rows.</para>
        /// </summary>
        /// <param name="numberOfRows">Number of rows to insert.</param>
        IXLRangeRows InsertRowsBelow(int numberOfRows);

        IXLRangeRows InsertRowsBelow(int numberOfRows, Boolean expandRange);

        /// <summary>
        /// Deletes this range and shifts the cells below.
        /// </summary>
        void Delete();

        /// <summary>
        /// Deletes this range and shifts the surrounding cells accordingly.
        /// </summary>
        /// <param name="shiftDeleteCells">How to shift the surrounding cells.</param>
        void Delete(XLShiftDeletedCells shiftDeleteCells);

        /// <summary>
        /// Gets this row's number in the range
        /// </summary>
        Int32 RowNumber();

        Int32 CellCount();

        IXLRangeRow CopyTo(IXLCell target);

        IXLRangeRow CopyTo(IXLRangeBase target);

        IXLRangeRow Sort();

        IXLRangeRow SortLeftToRight(XLSortOrder sortOrder = XLSortOrder.Ascending, Boolean matchCase = false, Boolean ignoreBlanks = true);

        IXLRangeRow Row(Int32 start, Int32 end);

        IXLRangeRow Row(IXLCell start, IXLCell end);

        IXLRangeRows Rows(String rows);

        IXLRangeRow SetDataType(XLDataType dataType);

        IXLRangeRow RowAbove();

        IXLRangeRow RowAbove(Int32 step);

        IXLRangeRow RowBelow();

        IXLRangeRow RowBelow(Int32 step);

        IXLRow WorksheetRow();

        /// <summary>
        /// Clears the contents of this row.
        /// </summary>
        /// <param name="clearOptions">Specify what you want to clear.</param>
        new IXLRangeRow Clear(XLClearOptions clearOptions = XLClearOptions.All);

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        IXLRangeRow RowUsed(Boolean includeFormats);

        IXLRangeRow RowUsed(XLCellsUsedOptions options = XLCellsUsedOptions.AllContents);
    }
}

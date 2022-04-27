using System;

namespace ClosedXML.Excel
{
    public interface IXLRangeRow : IXLRangeBase
    {
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
        /// <para>e.g. Cells("1"), Cells("1:5"), Cells("1:2,4:5")</para>
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

        /// <summary>
        /// Inserts X number of cells to the right of this row.
        /// <para>All cells to the right of this row will be shifted X number of columns.</para>
        /// </summary>
        /// <param name="numberOfColumns">Number of cells to insert.</param>
        IXLCells InsertCellsAfter(int numberOfColumns);

        IXLCells InsertCellsAfter(int numberOfColumns, bool expandRange);

        /// <summary>
        /// Inserts X number of cells to the left of this row.
        /// <para>This row and all cells to the right of it will be shifted X number of columns.</para>
        /// </summary>
        /// <param name="numberOfColumns">Number of cells to insert.</param>
        IXLCells InsertCellsBefore(int numberOfColumns);

        IXLCells InsertCellsBefore(int numberOfColumns, bool expandRange);

        /// <summary>
        /// Inserts X number of rows on top of this row.
        /// <para>This row and all cells below it will be shifted X number of rows.</para>
        /// </summary>
        /// <param name="numberOfRows">Number of rows to insert.</param>
        IXLRangeRows InsertRowsAbove(int numberOfRows);

        IXLRangeRows InsertRowsAbove(int numberOfRows, bool expandRange);

        /// <summary>
        /// Inserts X number of rows below this row.
        /// <para>All cells below this row will be shifted X number of rows.</para>
        /// </summary>
        /// <param name="numberOfRows">Number of rows to insert.</param>
        IXLRangeRows InsertRowsBelow(int numberOfRows);

        IXLRangeRows InsertRowsBelow(int numberOfRows, bool expandRange);

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
        int RowNumber();

        int CellCount();

        IXLRangeRow CopyTo(IXLCell target);

        IXLRangeRow CopyTo(IXLRangeBase target);

        IXLRangeRow Sort();

        IXLRangeRow SortLeftToRight(XLSortOrder sortOrder = XLSortOrder.Ascending, bool matchCase = false, bool ignoreBlanks = true);

        IXLRangeRow Row(int start, int end);

        IXLRangeRow Row(IXLCell start, IXLCell end);

        IXLRangeRows Rows(string rows);

        IXLRangeRow SetDataType(XLDataType dataType);

        IXLRangeRow RowAbove();

        IXLRangeRow RowAbove(int step);

        IXLRangeRow RowBelow();

        IXLRangeRow RowBelow(int step);

        IXLRow WorksheetRow();

        /// <summary>
        /// Clears the contents of this row.
        /// </summary>
        /// <param name="clearOptions">Specify what you want to clear.</param>
        new IXLRangeRow Clear(XLClearOptions clearOptions = XLClearOptions.All);

        [Obsolete("Use the overload with XLCellsUsedOptions")]
        IXLRangeRow RowUsed(bool includeFormats);

        IXLRangeRow RowUsed(XLCellsUsedOptions options = XLCellsUsedOptions.AllContents);
    }
}

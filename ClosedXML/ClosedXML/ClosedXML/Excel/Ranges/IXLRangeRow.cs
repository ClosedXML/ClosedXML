using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace ClosedXML.Excel
{
    public interface IXLRangeRow: IXLRangeBase
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
        IXLCells Cells(String cellsInRow);
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
        /// Converts this row to a range object.
        /// </summary>
        IXLRange AsRange();

        /// <summary>
        /// Inserts X number of cells to the right of this row.
        /// <para>All cells to the right of this row will be shifted X number of columns.</para>
        /// </summary>
        /// <param name="numberOfColumns">Number of cells to insert.</param>
        void InsertCellsAfter(int numberOfColumns);
        /// <summary>
        /// Inserts X number of cells to the left of this row.
        /// <para>This row and all cells to the right of it will be shifted X number of columns.</para>
        /// </summary>
        /// <param name="numberOfColumns">Number of cells to insert.</param>
        void InsertCellsBefore(int numberOfColumns);
        /// <summary>
        /// Inserts X number of rows on top of this row.
        /// <para>This row and all cells below it will be shifted X number of rows.</para>
        /// </summary>
        /// <param name="numberOfRows">Number of rows to insert.</param>
        void InsertRowsAbove(int numberOfRows);
        /// <summary>
        /// Inserts X number of rows below this row.
        /// <para>All cells below this row will be shifted X number of rows.</para>
        /// </summary>
        /// <param name="numberOfRows">Number of rows to insert.</param>
        void InsertRowsBelow(int numberOfRows);

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
        /// Clears the contents of the row (including styles).
        /// </summary>
        void Clear();
        /// <summary>
        /// Sets the formula for all cells in the row in A1 notation.
        /// </summary>
        /// <value>
        /// The formula A1.
        /// </value>
        String FormulaA1 { set; }
        /// <summary>
        /// Sets the formula for all cells in the row in R1C1 notation.
        /// </summary>
        /// <value>
        /// The formula R1C1.
        /// </value>
        String FormulaR1C1 { set; }
    }
}


using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace ClosedXML.Excel
{
    public interface IXLRangeColumn: IXLRangeBase
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
        IXLCells Cells(String cellsInColumn);
        /// <summary>
        /// Returns the specified group of cells.
        /// </summary>
        /// <param name="firstRow">The first row in the group of cells to return.</param>
        /// <param name="lastRow">The last row in the group of cells to return.</param>
        IXLCells Cells(Int32 firstRow, Int32 lastRow);

        /// <summary>
        /// Converts this column to a range object.
        /// </summary>
        IXLRange AsRange();

        /// <summary>
        /// Inserts X number of columns to the right of this range.
        /// <para>All cells to the right of this range will be shifted X number of columns.</para>
        /// </summary>
        /// <param name="numberOfColumns">Number of columns to insert.</param>
        void InsertColumnsAfter(int numberOfColumns);
        /// <summary>
        /// Inserts X number of columns to the left of this range.
        /// <para>This range and all cells to the right of this range will be shifted X number of columns.</para>
        /// </summary>
        /// <param name="numberOfColumns">Number of columns to insert.</param>
        void InsertColumnsBefore(int numberOfColumns);
        /// <summary>
        /// Inserts X number of cells on top of this column.
        /// <para>This column and all cells below it will be shifted X number of rows.</para>
        /// </summary>
        /// <param name="numberOfRows">Number of cells to insert.</param>
        void InsertCellsAbove(int numberOfRows);
        /// <summary>
        /// Inserts X number of cells below this range.
        /// <para>All cells below this column will be shifted X number of rows.</para>
        /// </summary>
        /// <param name="numberOfRows">Number of cells to insert.</param>
        void InsertCellsBelow(int numberOfRows);

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
        /// Clears the contents of the column (including styles).
        /// </summary>
        void Clear();
        /// <summary>
        /// Sets the formula for all cells in the column in A1 notation.
        /// </summary>
        /// <value>
        /// The formula A1.
        /// </value>
        String FormulaA1 { set; }
        /// <summary>
        /// Sets the formula for all cells in the column in R1C1 notation.
        /// </summary>
        /// <value>
        /// The formula R1C1.
        /// </value>
        String FormulaR1C1 { set; }
    }
}


using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace ClosedXML.Excel
{
    public interface IXLRangeColumn: IXLRangeBase
    {
        IXLCell Cell(int row);

        IEnumerable<IXLCell> Cells(int firstRow, int lastRow);
        IXLRange Range(int firstColumn, int lastColumn);

        int RowCount();

        void InsertColumnsAfter(int numberOfColumns);
        void InsertColumnsBefore(int numberOfColumns);
        void InsertRowsAbove(int numberOfRows);
        void InsertRowsBelow(int numberOfRows);

        void Delete(XLShiftDeletedCells shiftDeleteCells = XLShiftDeletedCells.ShiftCellsLeft);
        void Clear();
    }
}


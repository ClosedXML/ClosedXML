using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace ClosedXML.Excel
{
    public interface IXLRangeRow: IXLRangeBase
    {
        IXLCell Cell(int column);
        IXLCell Cell(string column);

        IEnumerable<IXLCell> Cells(int firstColumn, int lastColumn);
        IEnumerable<IXLCell> Cells(String firstColumn, String lastColumn);
        IXLRange Range(int firstRow, int lastRow);

        int ColumnCount();

        void InsertColumnsAfter(int numberOfColumns);
        void InsertColumnsBefore(int numberOfColumns);
        void InsertRowsAbove(int numberOfRows);
        void InsertRowsBelow(int numberOfRows);

        void Delete(XLShiftDeletedCells shiftDeleteCells = XLShiftDeletedCells.ShiftCellsUp);
        void Clear();
    }
}


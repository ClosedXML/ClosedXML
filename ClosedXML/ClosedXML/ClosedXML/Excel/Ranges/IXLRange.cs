using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace ClosedXML.Excel
{
    public enum XLShiftDeletedCells { ShiftCellsUp, ShiftCellsLeft }
    public interface IXLRange: IXLRangeBase
    {
        IXLRange Column(int column);
        IXLRange Column(string column);
        IXLRanges Columns();
        IXLRanges Columns(int firstColumn, int lastColumn);
        IXLRanges Columns(string columns);
        IXLRanges Columns(string firstColumn, string lastColumn);
        IXLRange FirstColumn();
        IXLRange FirstColumnUsed();
        IXLRange FirstRow();
        IXLRange FirstRowUsed();
        IXLRange LastColumn();
        IXLRange LastColumnUsed();
        IXLRange LastRow();
        IXLRange LastRowUsed();
        IXLRange Row(int row);
        IXLRanges Rows();
        IXLRanges Rows(int firstRow, int lastRow);
        IXLRanges Rows(string rows);
        void InsertColumnsAfter(int numberOfColumns);
        void InsertColumnsBefore(int numberOfColumns);
        void InsertRowsAbove(int numberOfRows);
        void InsertRowsBelow(int numberOfRows);
        void Delete(XLShiftDeletedCells shiftDeleteCells);
        void Clear();
    }
}


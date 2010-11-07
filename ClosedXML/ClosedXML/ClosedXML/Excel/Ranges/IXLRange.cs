using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace ClosedXML.Excel
{
    public enum XLShiftDeletedCells { ShiftCellsUp, ShiftCellsLeft }
    public enum XLTransposeOptions { MoveCells, ReplaceCells }
    public interface IXLRange: IXLRangeBase
    {
        IXLCell Cell(int row, int column);
        IXLCell Cell(string cellAddressInRange);
        IXLCell Cell(int row, string column);
        IXLCell Cell(IXLAddress cellAddressInRange);

        IXLRangeColumn Column(int column);
        IXLRangeColumn Column(string column);
        IXLRangeColumn FirstColumn();
        IXLRangeColumn FirstColumnUsed();
        IXLRangeColumn LastColumn();
        IXLRangeColumn LastColumnUsed();
        IXLRangeColumns Columns();
        IXLRangeColumns Columns(int firstColumn, int lastColumn);
        IXLRangeColumns Columns(string firstColumn, string lastColumn);
        IXLRangeColumns Columns(string columns);

        IXLRangeRow FirstRow();
        IXLRangeRow FirstRowUsed();
        IXLRangeRow LastRow();
        IXLRangeRow LastRowUsed();
        IXLRangeRow Row(int row);
        IXLRangeRows Rows();
        IXLRangeRows Rows(int firstRow, int lastRow);
        IXLRangeRows Rows(string rows);

        IXLRange Range(int firstCellRow, int firstCellColumn, int lastCellRow, int lastCellColumn);

        int RowCount();
        int ColumnCount();

        void InsertColumnsAfter(int numberOfColumns);
        void InsertColumnsBefore(int numberOfColumns);
        void InsertRowsAbove(int numberOfRows);
        void InsertRowsBelow(int numberOfRows);
        void Delete(XLShiftDeletedCells shiftDeleteCells);
        void Clear();

        void Transpose(XLTransposeOptions transposeOption);
    }
}


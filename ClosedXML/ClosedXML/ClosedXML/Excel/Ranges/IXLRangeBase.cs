using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLRangeBase: IXLStylized
    {
        IXLCell Cell(ClosedXML.Excel.IXLAddress cellAddressInRange);
        IXLCell Cell(int row, int column);
        IXLCell Cell(int row, string column);
        IXLCell Cell(string cellAddressInRange);
        IEnumerable<ClosedXML.Excel.IXLCell> Cells();
        IEnumerable<ClosedXML.Excel.IXLCell> CellsUsed();
        IXLAddress FirstAddressInSheet { get; }
        IXLAddress LastAddressInSheet { get; }
        IXLCell FirstCell();
        IXLCell LastCell();
        IXLRange Range(ClosedXML.Excel.IXLAddress firstCellAddress, ClosedXML.Excel.IXLAddress lastCellAddress);
        IXLRange Range(int firstCellRow, int firstCellColumn, int lastCellRow, int lastCellColumn);
        IXLRange Range(string firstCellAddress, string lastCellAddress);
        IXLRange Range(string rangeAddress);
        IXLRanges Ranges(params string[] ranges);
        IXLRanges Ranges(string ranges);
        int RowCount();
        int ColumnCount();
        void Unmerge();
        void Merge();
        IXLRange AsRange();
    }
}

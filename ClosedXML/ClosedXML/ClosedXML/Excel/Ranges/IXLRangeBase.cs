using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public enum XLScope { Workbook, Worksheet };
    public interface IXLRangeBase: IXLStylized
    {
        IEnumerable<IXLCell> Cells();
        IEnumerable<IXLCell> CellsUsed();
        IXLRangeAddress RangeAddress { get; }
        IXLCell FirstCell();
        IXLCell FirstCellUsed(Boolean ignoreStyle = true);
        IXLCell LastCell();
        IXLCell LastCellUsed(Boolean ignoreStyle = true);
        IXLRange Range(IXLRangeAddress rangeAddress);
        IXLRange Range(string rangeAddress);
        IXLRange Range(string firstCellAddress, string lastCellAddress);
        IXLRange Range(IXLAddress firstCellAddress, IXLAddress lastCellAddress);
        IXLRanges Ranges(params string[] ranges);
        IXLRanges Ranges(string ranges);
        void Unmerge();
        void Merge();
        IXLRange AsRange();
        Boolean ContainsRange(String rangeAddress);
        void CreateNamedRange(String rangeName, XLScope scope = XLScope.Workbook, String comment = null);
    }
}

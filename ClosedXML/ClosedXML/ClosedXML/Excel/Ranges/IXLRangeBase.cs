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
        IXLCell FirstCellUsed();
        IXLCell FirstCellUsed(Boolean ignoreStyle);
        IXLCell LastCell();
        IXLCell LastCellUsed();
        IXLCell LastCellUsed(Boolean ignoreStyle);
        IXLRange Range(IXLRangeAddress rangeAddress);
        IXLRange Range(string rangeAddress);
        IXLRange Range(string firstCellAddress, string lastCellAddress);
        IXLRange Range(IXLAddress firstCellAddress, IXLAddress lastCellAddress);
        IXLRanges Ranges(params string[] ranges);
        IXLRanges Ranges(string ranges);
        IXLRange Unmerge();
        IXLRange Merge();
        IXLRange AsRange();
        Boolean ContainsRange(String rangeAddress);
        IXLRange AddToNamed(String rangeName);
        IXLRange AddToNamed(String rangeName, XLScope scope);
        IXLRange AddToNamed(String rangeName, XLScope scope, String comment);
    }
}

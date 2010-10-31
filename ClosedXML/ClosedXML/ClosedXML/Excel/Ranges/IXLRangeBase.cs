using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLRangeBase: IXLStylized
    {
        IEnumerable<IXLCell> Cells();
        IEnumerable<IXLCell> CellsUsed();
        IXLAddress FirstAddressInSheet { get; }
        IXLAddress LastAddressInSheet { get; }
        IXLCell FirstCell();
        IXLCell FirstCellUsed(Boolean ignoreStyle = true);
        IXLCell LastCell();
        IXLCell LastCellUsed(Boolean ignoreStyle = true);
        IXLRange Range(string rangeAddress);
        IXLRange Range(string firstCellAddress, string lastCellAddress);
        IXLRange Range(IXLAddress firstCellAddress, IXLAddress lastCellAddress);
        IXLRanges Ranges(params string[] ranges);
        IXLRanges Ranges(string ranges);
        void Unmerge();
        void Merge();
        IXLRange AsRange();
    }
}

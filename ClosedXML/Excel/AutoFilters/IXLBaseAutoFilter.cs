// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public enum XLFilterDynamicType { AboveAverage, BelowAverage }

    public enum XLFilterType { Regular, Custom, TopBottom, Dynamic, DateTimeGrouping }

    public enum XLTopBottomPart { Top, Bottom }

    public interface IXLBaseAutoFilter
    {
        Boolean Enabled { get; set; }
        IXLRange Range { get; set; }
        Int32 SortColumn { get; set; }
        Boolean Sorted { get; set; }
        XLSortOrder SortOrder { get; set; }
        IEnumerable<IXLRangeRow> VisibleRows { get; }

        IXLBaseAutoFilter Clear();

        IXLFilterColumn Column(String column);

        IXLFilterColumn Column(Int32 column);

        IXLBaseAutoFilter Set(IXLRangeBase range);

        IXLBaseAutoFilter Sort(Int32 columnToSortBy, XLSortOrder sortOrder = XLSortOrder.Ascending, Boolean matchCase = false, Boolean ignoreBlanks = true);
    }
}

using System;
namespace ClosedXML.Excel
{
    using System.Collections.Generic;
    public enum XLFilterType { Regular, Custom, TopBottom, Dynamic }
    public enum XLFilterDynamicType { AboveAverage, BelowAverage }
    public enum XLTopBottomPart { Top, Bottom}
    public interface IXLBaseAutoFilter
    {
        Boolean Enabled { get; set; }
        IXLRange Range { get; set; }
        IXLBaseAutoFilter Set(IXLRangeBase range);
        IXLBaseAutoFilter Clear();

        IXLFilterColumn Column(String column);
        IXLFilterColumn Column(Int32 column);

        IXLBaseAutoFilter Sort(Int32 columnToSortBy, XLSortOrder sortOrder = XLSortOrder.Ascending, Boolean matchCase = false, Boolean ignoreBlanks = true);
        Boolean Sorted { get; set; }
        XLSortOrder SortOrder { get; set; }
        Int32 SortColumn { get; set; }


    }
}
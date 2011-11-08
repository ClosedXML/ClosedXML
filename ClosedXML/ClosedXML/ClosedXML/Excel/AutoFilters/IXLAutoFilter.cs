using System;
namespace ClosedXML.Excel
{
    using System.Collections.Generic;

    public interface IXLAutoFilter
    {
        //IXLAutoFilter Set();
        //List<IXLFilter> Filters { get; }
        //List<IXLCustomFilter> CustomFilters { get; }
        IXLFilterColumn Column(String column);
        IXLFilterColumn Column(Int32 column);

        IXLAutoFilter Sort(Int32 columnToSortBy, XLSortOrder sortOrder = XLSortOrder.Ascending, Boolean matchCase = false, Boolean ignoreBlanks = true);
        Boolean Sorted { get; set; }
        XLSortOrder SortOrder { get; set; }
        Int32 SortColumn { get; set; }
    }
}
using System;
namespace ClosedXML.Excel
{
    using System.Collections.Generic;

    public interface IXLAutoFilter: IDisposable
    {
        IXLFilterColumn Column(String column);
        IXLFilterColumn Column(Int32 column);

        IXLAutoFilter Sort(Int32 columnToSortBy = 1, XLSortOrder sortOrder = XLSortOrder.Ascending, Boolean matchCase = false, Boolean ignoreBlanks = true);
        Boolean Sorted { get; set; }
        XLSortOrder SortOrder { get; set; }
        Int32 SortColumn { get; set; }
    }
}
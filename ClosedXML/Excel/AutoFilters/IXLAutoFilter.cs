// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public enum XLFilterDynamicType { AboveAverage, BelowAverage }

    public enum XLFilterType { Regular, Custom, TopBottom, Dynamic, DateTimeGrouping }

    public enum XLTopBottomPart { Top, Bottom }

    public interface IXLAutoFilter
    {
        [Obsolete("Use IsEnabled")]
        bool Enabled { get; set; }
        IEnumerable<IXLRangeRow> HiddenRows { get; }
        bool IsEnabled { get; set; }
        IXLRange Range { get; set; }
        int SortColumn { get; set; }
        bool Sorted { get; set; }
        XLSortOrder SortOrder { get; set; }
        IEnumerable<IXLRangeRow> VisibleRows { get; }

        IXLAutoFilter Clear();

        IXLFilterColumn Column(string column);

        IXLFilterColumn Column(int column);

        IXLAutoFilter Reapply();

        IXLAutoFilter Sort(int columnToSortBy = 1, XLSortOrder sortOrder = XLSortOrder.Ascending, bool matchCase = false, bool ignoreBlanks = true);
    }
}

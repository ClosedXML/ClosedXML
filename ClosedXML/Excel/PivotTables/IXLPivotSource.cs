using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLPivotSource
    {
        IDictionary<String, IList<Object>> CachedFields { get; }
        XLItemsToRetain ItemsToRetainPerField { get; set; }
        IXLPivotSourceReference PivotSourceReference { get; }
        Boolean RefreshDataOnOpen { get; set; }

        Boolean SaveSourceData { get; set; }
        IList<String> SourceRangeFields { get; }

        IXLPivotSource Refresh();

        IXLPivotSource SetItemsToRetainPerField(XLItemsToRetain value);

        IXLPivotSource SetRefreshDataOnOpen();

        IXLPivotSource SetRefreshDataOnOpen(Boolean value);

        IXLPivotSource SetSaveSourceData();

        IXLPivotSource SetSaveSourceData(Boolean value);
    }
}

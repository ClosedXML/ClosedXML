using System;

namespace ClosedXML.Excel
{
    public enum XLSortOrder { Ascending, Descending }
    public enum XLSortOrientation { TopToBottom, LeftToRight }
    public interface IXLSortElement
    {
        Int32 ElementNumber { get; }
        XLSortOrder SortOrder { get; }
        Boolean IgnoreBlanks { get; }
        Boolean MatchCase { get; }
    }
}

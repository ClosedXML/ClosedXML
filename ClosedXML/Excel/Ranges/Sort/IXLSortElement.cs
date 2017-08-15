using System;

namespace ClosedXML.Excel
{
    public enum XLSortOrder { Ascending, Descending }
    public enum XLSortOrientation { TopToBottom, LeftToRight }
    public interface IXLSortElement
    {
        Int32 ElementNumber { get; set; }
        XLSortOrder SortOrder { get; set; }
        Boolean IgnoreBlanks { get; set; }
        Boolean MatchCase { get; set; }
    }
}

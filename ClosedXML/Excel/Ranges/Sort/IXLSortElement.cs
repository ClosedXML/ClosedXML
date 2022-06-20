namespace ClosedXML.Excel
{
    public enum XLSortOrder { Ascending, Descending }
    public enum XLSortOrientation { TopToBottom, LeftToRight }
    public interface IXLSortElement
    {
        int ElementNumber { get; set; }
        XLSortOrder SortOrder { get; set; }
        bool IgnoreBlanks { get; set; }
        bool MatchCase { get; set; }
    }
}

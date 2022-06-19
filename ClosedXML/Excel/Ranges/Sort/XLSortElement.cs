namespace ClosedXML.Excel
{
    internal class XLSortElement: IXLSortElement
    {
        public int ElementNumber { get; set; }
        public XLSortOrder SortOrder { get; set; }
        public bool IgnoreBlanks { get; set; }
        public bool MatchCase { get; set; }
    }
}

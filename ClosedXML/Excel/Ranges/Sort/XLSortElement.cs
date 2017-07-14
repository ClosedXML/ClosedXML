using System;

namespace ClosedXML.Excel
{
    internal class XLSortElement: IXLSortElement
    {
        public Int32 ElementNumber { get; set; }
        public XLSortOrder SortOrder { get; set; }
        public Boolean IgnoreBlanks { get; set; }
        public Boolean MatchCase { get; set; }
    }
}

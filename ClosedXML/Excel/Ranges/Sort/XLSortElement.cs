// Keep this file CodeMaid organised and cleaned
using System;

namespace ClosedXML.Excel
{
    internal class XLSortElement : IXLSortElement
    {
        public XLSortElement(Int32 elementNumber, XLSortOrder sortOrder, Boolean matchCase = false, Boolean ignoreBlanks = true)
        {
            this.ElementNumber = elementNumber;
            this.SortOrder = sortOrder;
            this.MatchCase = matchCase;
            this.IgnoreBlanks = ignoreBlanks;
            this.CellComparer = new XLCellComparer(this);
        }

        public Int32 ElementNumber { get; }
        public Boolean IgnoreBlanks { get; }
        public Boolean MatchCase { get; }
        public XLSortOrder SortOrder { get; }
        internal XLCellComparer CellComparer { get; }
    }
}

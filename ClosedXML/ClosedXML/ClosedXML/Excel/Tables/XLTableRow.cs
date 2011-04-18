using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLTableRow: XLRangeRow, IXLTableRow
    {
        private XLTable table;
        public XLTableRow(XLTable table, XLRangeRow rangeRow)
            : base(rangeRow.RangeParameters)
        {
            this.table = table;
        }

        public IXLCell Field(Int32 index)
        {
            return Cell(index + 1);
        }

        public IXLCell Field(String name)
        {
            Int32 fieldIndex = table.GetFieldIndex(name);
            return Cell(fieldIndex + 1);
        }

        public new IXLTableRow Sort()
        {
            this.AsRange().Sort(XLSortOrientation.LeftToRight);
            return this;
        }
        public new IXLTableRow Sort(XLSortOrder sortOrder)
        {
            this.AsRange().Sort(XLSortOrientation.LeftToRight, sortOrder);
            return this;
        }
        public new IXLTableRow Sort(Boolean matchCase)
        {
            this.AsRange().Sort(XLSortOrientation.LeftToRight, matchCase);
            return this;
        }
        public new IXLTableRow Sort(XLSortOrder sortOrder, Boolean matchCase)
        {
            this.AsRange().Sort(XLSortOrientation.LeftToRight, sortOrder, matchCase);
            return this;
        }
    }
}

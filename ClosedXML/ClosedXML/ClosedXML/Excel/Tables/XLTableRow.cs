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
    }
}

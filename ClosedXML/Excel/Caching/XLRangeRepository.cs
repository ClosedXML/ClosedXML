using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace ClosedXML.Excel.Caching
{
    internal class XLRangeRepository : XLWorkbookElementRepositoryBase<XLRangeAddress, XLRange>
    {
        public XLRangeRepository(XLWorkbook workbook, Func<XLRangeAddress, XLRange> createNew) : base(workbook, createNew)
        {
        }

        public XLRangeRepository(XLWorkbook workbook, Func<XLRangeAddress, XLRange> createNew, IEqualityComparer<XLRangeAddress> сomparer) : base(workbook, createNew, сomparer)
        {
        }

        public override XLRange GetOrCreate(XLRangeAddress key)
        {
            var range = base.GetOrCreate(key);
            Debug.Assert(key.Equals(range.RangeAddress), "Range address differs");
            return range;
        }
    }
}

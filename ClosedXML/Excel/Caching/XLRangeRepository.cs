using System;
using System.Collections.Generic;

namespace ClosedXML.Excel.Caching
{
    internal class XLRangeRepository : XLWorkbookElementRepositoryBase<XLRangeKey, XLRangeBase>
    {
        public XLRangeRepository(XLWorkbook workbook, Func<XLRangeKey, XLRangeBase> createNew) : base(workbook, createNew)
        {
        }

        public XLRangeRepository(XLWorkbook workbook, Func<XLRangeKey, XLRangeBase> createNew, IEqualityComparer<XLRangeKey> сomparer) : base(workbook, createNew, сomparer)
        {
        }
    }
}

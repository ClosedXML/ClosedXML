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

        public override XLRangeBase Store(XLRangeKey key, XLRangeBase value)
        {
            //DEBUG
            if (value != null && !key.RangeAddress.Equals(value.RangeAddress))
                throw new ArgumentException("Range address differs when stored");
            return base.Store(key, value);
        }

        public override XLRangeBase GetOrCreate(XLRangeKey key)
        {
            //DEBUG
            var range = base.GetOrCreate(key);
            if (!key.RangeAddress.Equals(range.RangeAddress))
                throw new ArgumentException("Range address differs when obtained");
            return range;
        }
    }
}

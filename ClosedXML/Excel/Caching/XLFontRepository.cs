using System;
using System.Collections.Generic;

namespace ClosedXML.Excel.Caching
{
    internal sealed class XLFontRepository : XLRepositoryBase<XLFontKey, XLFontValue>
    {
        #region constructors
        public XLFontRepository(Func<XLFontKey, XLFontValue> createNew) : base(createNew)
        {
        }

        public XLFontRepository(Func<XLFontKey, XLFontValue> createNew, IEqualityComparer<XLFontKey> comparer) : base(createNew, comparer)
        {
        }
        #endregion
    }
}

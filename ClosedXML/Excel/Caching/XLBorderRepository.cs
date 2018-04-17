using System;
using System.Collections.Generic;

namespace ClosedXML.Excel.Caching
{
    internal sealed class XLBorderRepository : XLRepositoryBase<XLBorderKey, XLBorderValue>
    {
        #region Constructors

        public XLBorderRepository(Func<XLBorderKey, XLBorderValue> createNew)
            : base(createNew)
        {
        }

        public XLBorderRepository(Func<XLBorderKey, XLBorderValue> createNew, IEqualityComparer<XLBorderKey> comparer)
            : base(createNew, comparer)
        {
        }

        #endregion Constructors
    }
}

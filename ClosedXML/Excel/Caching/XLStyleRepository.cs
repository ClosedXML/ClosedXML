using System;
using System.Collections.Generic;

namespace ClosedXML.Excel.Caching
{
    internal sealed class XLStyleRepository : XLRepositoryBase<XLStyleKey, XLStyleValue>
    {
        #region Constructors

        public XLStyleRepository(Func<XLStyleKey, XLStyleValue> createNew)
            : base(createNew)
        {
        }

        public XLStyleRepository(Func<XLStyleKey, XLStyleValue> createNew, IEqualityComparer<XLStyleKey> comparer)
            : base(createNew, comparer)
        {
        }

        #endregion Constructors
    }
}

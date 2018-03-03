using System;
using System.Collections.Generic;

namespace ClosedXML.Excel.Caching
{
    internal sealed class XLColorRepository : XLRepositoryBase<XLColorKey, XLColor>
    {
        #region constructors
        public XLColorRepository(Func<XLColorKey, XLColor> createNew) : base(createNew)
        {
        }

        public XLColorRepository(Func<XLColorKey, XLColor> createNew, IEqualityComparer<XLColorKey> comparer) : base(createNew, comparer)
        {
        }
        #endregion
    }
}

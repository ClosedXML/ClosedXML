using System;
using System.Collections.Generic;

namespace ClosedXML.Excel.Caching
{
    internal sealed class XLProtectionRepository : XLRepositoryBase<XLProtectionKey, XLProtectionValue>
    {
        #region constructors
        public XLProtectionRepository(Func<XLProtectionKey, XLProtectionValue> createNew) : base(createNew)
        {
        }

        public XLProtectionRepository(Func<XLProtectionKey, XLProtectionValue> createNew, IEqualityComparer<XLProtectionKey> comparer) : base(createNew, comparer)
        {
        }
        #endregion
    }
}

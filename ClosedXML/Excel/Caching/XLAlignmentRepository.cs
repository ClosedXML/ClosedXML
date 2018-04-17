using System;
using System.Collections.Generic;

namespace ClosedXML.Excel.Caching
{
    internal sealed class XLAlignmentRepository : XLRepositoryBase<XLAlignmentKey, XLAlignmentValue>
    {
        #region Constructors

        public XLAlignmentRepository(Func<XLAlignmentKey, XLAlignmentValue> createNew)
            : base(createNew)
        {
        }

        public XLAlignmentRepository(Func<XLAlignmentKey, XLAlignmentValue> createNew, IEqualityComparer<XLAlignmentKey> comparer)
            : base(createNew, comparer)
        {
        }

        #endregion Constructors
    }
}

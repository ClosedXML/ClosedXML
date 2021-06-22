using System.Collections.Generic;

namespace ClosedXML.Excel.CalcEngine
{
    internal class XLRangeByAddressComparer : IEqualityComparer<IXLRange>
    {
        private readonly XLRangeAddressComparer _rangeAddressComparer;

        public XLRangeByAddressComparer()
        {
            _rangeAddressComparer = new XLRangeAddressComparer(true);
        }

        public bool Equals(IXLRange x, IXLRange y)
        {
            if (ReferenceEquals(x, y)) return true;

            if (ReferenceEquals(x, null) ||
                ReferenceEquals(y, null))
                return false;

            return _rangeAddressComparer.Equals(x.RangeAddress, y.RangeAddress);
        }

        public int GetHashCode(IXLRange obj)
        {
            return _rangeAddressComparer.GetHashCode(obj.RangeAddress);
        }
    }
}

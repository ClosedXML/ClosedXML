using System.Collections.Generic;

namespace ClosedXML.Excel.CalcEngine
{
    internal class XLRangeAddressComparer : IEqualityComparer<IXLRangeAddress>
    {
        private readonly XLAddressComparer _addressComparer;

        public XLRangeAddressComparer(bool ignoreFixed)
        {
            _addressComparer = new XLAddressComparer(ignoreFixed);
        }

        public bool Equals(IXLRangeAddress x, IXLRangeAddress y)
        {
            return (x == null && y == null) ||
                   (x != null && y != null &&
                    _addressComparer.Equals(x.FirstAddress, y.FirstAddress) &&
                    _addressComparer.Equals(x.LastAddress, y.LastAddress));
        }

        public int GetHashCode(IXLRangeAddress obj)
        {
            return new
            {
                FirstHash = _addressComparer.GetHashCode(obj.FirstAddress),
                LastHash = _addressComparer.GetHashCode(obj.LastAddress),
            }.GetHashCode();
        }
    }
}

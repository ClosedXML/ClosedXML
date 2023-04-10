#nullable disable

using System.Collections.Generic;

namespace ClosedXML.Excel.CalcEngine
{
    internal class XLRangeAddressComparer : IEqualityComparer<IXLRangeAddress>
    {
        private readonly XLAddressComparer _addressComparer;

        /// <summary>
        /// Comparer of ranges that ignores whether row/column is fixes or not.
        /// </summary>
        internal static readonly XLRangeAddressComparer IgnoreFixed = new(true);

        private XLRangeAddressComparer(bool ignoreFixed)
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

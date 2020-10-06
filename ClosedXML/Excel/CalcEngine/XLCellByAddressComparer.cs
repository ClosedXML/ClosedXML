using System.Collections.Generic;

namespace ClosedXML.Excel.CalcEngine
{
    internal class XLCellByAddressComparer : IEqualityComparer<IXLCell>
    {
        private readonly XLAddressComparer _addressComparer;

        public XLCellByAddressComparer()
        {
            _addressComparer = new XLAddressComparer(true);
        }

        public bool Equals(IXLCell x, IXLCell y)
        {
            if (ReferenceEquals(x, y)) return true;

            if (ReferenceEquals(x, null) ||
                ReferenceEquals(y, null))
                return false;

            return _addressComparer.Equals(x.Address, y.Address);
        }

        public int GetHashCode(IXLCell obj)
        {
            return _addressComparer.GetHashCode(obj.Address);
        }
    }
}

using System;
using System.Collections.Generic;

namespace ClosedXML.Excel.CalcEngine
{
    internal class XLAddressComparer : IEqualityComparer<IXLAddress>
    {
        private readonly bool _ignoreFixed;

        public XLAddressComparer(bool ignoreFixed)
        {
            _ignoreFixed = ignoreFixed;
        }

        public bool Equals(IXLAddress x, IXLAddress y)
        {
            return (x == null && y == null) ||
                   (x != null && y != null &&
                    string.Equals(x.Worksheet.Name, y.Worksheet.Name, StringComparison.InvariantCultureIgnoreCase) &&
                    x.ColumnNumber == y.ColumnNumber &&
                    x.RowNumber == y.RowNumber &&
                    (_ignoreFixed || x.FixedColumn == y.FixedColumn &&
                     x.FixedRow == y.FixedRow));
        }

        public int GetHashCode(IXLAddress obj)
        {
            return new
            {
                WorksheetName = obj.Worksheet.Name.ToUpperInvariant(),
                obj.ColumnNumber,
                obj.RowNumber,
                FixedColumn = (_ignoreFixed ? false : obj.FixedColumn),
                FixedRow = (_ignoreFixed ? false : obj.FixedRow)
            }.GetHashCode();
        }
    }
}

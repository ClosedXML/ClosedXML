#nullable disable

using System;
using System.Collections.Generic;

namespace ClosedXML.Excel.Cells
{
    internal class XLCellValueComparer : IEqualityComparer<XLCellValue>
    {
        private readonly StringComparer _textComparer;

        internal static readonly XLCellValueComparer OrdinalIgnoreCase = new(StringComparer.OrdinalIgnoreCase);

        private XLCellValueComparer(StringComparer textComparer)
        {
            _textComparer = textComparer;
        }

        public bool Equals(XLCellValue x, XLCellValue y)
        {
            if (x.Type != y.Type)
                return false;

            return x.Type switch
            {
                XLDataType.Blank => true,
                XLDataType.Boolean => x.GetBoolean() == y.GetBoolean(),
                XLDataType.Number => x.GetNumber().Equals(y.GetNumber()),
                XLDataType.Text => _textComparer.Equals(x.GetText(), y.GetText()),
                XLDataType.Error => x.GetError() == y.GetError(),
                XLDataType.DateTime => x.GetUnifiedNumber().Equals(y.GetUnifiedNumber()),
                XLDataType.TimeSpan => x.GetUnifiedNumber().Equals(y.GetUnifiedNumber()),
                _ => throw new NotSupportedException()
            };
        }

        public int GetHashCode(XLCellValue obj)
        {
            unchecked
            {
                var hashCode = obj.Type.GetHashCode();
                var valueHashCode = obj.Type switch
                {
                    XLDataType.Blank => 0,
                    XLDataType.Boolean => obj.GetBoolean().GetHashCode(),
                    XLDataType.Text => _textComparer.GetHashCode(obj.GetText()),
                    XLDataType.Error => obj.GetError().GetHashCode(),
                    _ => obj.GetUnifiedNumber().GetHashCode()
                };
                hashCode = (hashCode * 397) ^ valueHashCode;
                return hashCode;
            }
        }
    }
}

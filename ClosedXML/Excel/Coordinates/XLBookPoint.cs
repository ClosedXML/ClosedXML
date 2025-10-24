using System;
using ClosedXML.Parser;

namespace ClosedXML.Excel
{
    /// <summary>
    /// A single point in a workbook. The book point might point to a deleted
    /// worksheet, so it might be invalid. Make sure it is checked when
    /// determining the properties of the actual data of the point.
    /// </summary>
    internal readonly struct XLBookPoint : IEquatable<XLBookPoint>
    {
        internal XLBookPoint(string sheetName, int row, int col)
            : this(sheetName, new XLSheetPoint(row, col))
        {
        }

        internal XLBookPoint(string sheetName, XLSheetPoint point)
        {
            if (string.IsNullOrEmpty(sheetName))
                throw new ArgumentException(nameof(sheetName));

            SheetName = sheetName;
            Point = point;
        }

        /// <summary>
        /// Name of the sheet. The sheet may be deleted.
        /// </summary>
        public string SheetName { get; }

        /// <inheritdoc cref="XLSheetPoint.Row"/>
        public int Row => Point.Column;

        /// <inheritdoc cref="XLSheetPoint.Column"/>
        public int Column => Point.Column;

        /// <summary>
        /// A point in the sheet.
        /// </summary>
        public XLSheetPoint Point { get; }

        public static bool operator ==(XLBookPoint lhs, XLBookPoint rhs) => lhs.Equals(rhs);

        public static bool operator !=(XLBookPoint lhs, XLBookPoint rhs) => !(lhs == rhs);

        public bool Equals(XLBookPoint other)
        {
            return Point.Equals(other.Point) && XLHelper.SheetComparer.Equals(SheetName, other.SheetName);
        }

        public override bool Equals(object? obj)
        {
            return obj is XLBookPoint other && Equals(other);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                return (XLHelper.SheetComparer.GetHashCode(SheetName) * 397) ^ Point.GetHashCode();
            }
        }

        public override string ToString()
        {
            var name = NameUtils.ShouldQuote(SheetName.AsSpan())
                ? SheetName.AlwaysEscapeSheetName()
                : SheetName;
            return $"{name}!{Point}";
        }
    }
}

using System;

namespace ClosedXML.Excel
{
    /// <summary>
    /// An point (address) in a worksheet.
    /// </summary>
    /// <remarks>Unlike the XLAddress, sheet can never be invalid.</remarks>
    internal readonly struct XLSheetPoint : IEquatable<XLSheetPoint>
    {
        public XLSheetPoint(Int32 row, Int32 column)
        {
            Row = row;
            Column = column;
        }

        /// <summary>
        /// 1-based row number in a sheet.
        /// </summary>
        public readonly Int32 Row;

        /// <summary>
        /// 1-based column number in a sheet.
        /// </summary>
        public readonly Int32 Column;

        public override bool Equals(object obj)
        {
            return obj is XLSheetPoint point && Equals(point);
        }

        public bool Equals(XLSheetPoint other)
        {
            return Row == other.Row && Column == other.Column;
        }

        public override int GetHashCode()
        {
            return (Row * -1) ^ Column;
        }

        public static bool operator ==(XLSheetPoint a, XLSheetPoint b)
        {
            return a.Row == b.Row && a.Column == b.Column;
        }

        public static bool operator !=(XLSheetPoint a, XLSheetPoint b)
        {
            return a.Row != b.Row || a.Column != b.Column;
        }
    }
}

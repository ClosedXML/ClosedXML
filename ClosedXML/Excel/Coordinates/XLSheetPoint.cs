using System;

namespace ClosedXML.Excel
{
    internal struct XLSheetPoint:IEquatable<XLSheetPoint>
    {
        public XLSheetPoint(Int32  row, Int32 column)
        {
            Row = row;
            Column = column;
        }

        public readonly Int32 Row;
        public readonly Int32 Column;

        public override bool Equals(object obj)
        {
            return Equals((XLSheetPoint)obj);
        }

        public bool Equals(XLSheetPoint other)
        {
            return Row == other.Row && Column == other.Column;
        }

        public override int GetHashCode()
        {
            return (Row * -1) ^ Column;
        }

        public static bool operator==(XLSheetPoint a, XLSheetPoint b)
        {
            return a.Row == b.Row && a.Column == b.Column;
        }

        public static bool operator !=(XLSheetPoint a, XLSheetPoint b)
        {
            return a.Row != b.Row || a.Column != b.Column;
        }
    }
}

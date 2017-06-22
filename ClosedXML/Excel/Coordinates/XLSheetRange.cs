using System;

namespace ClosedXML.Excel
{
    internal struct XLSheetRange:IEquatable<XLSheetRange>
    {
        public XLSheetRange(XLSheetPoint firstPoint, XLSheetPoint lastPoint)
        {
            FirstPoint = firstPoint;
            LastPoint = lastPoint;
        }

        public readonly XLSheetPoint FirstPoint;
        public readonly XLSheetPoint LastPoint;

        public bool Equals(XLSheetRange other)
        {
            return FirstPoint.Equals(other.FirstPoint) && LastPoint.Equals(other.LastPoint);
        }

        public override int GetHashCode()
        {
            return FirstPoint.GetHashCode() ^ LastPoint.GetHashCode();
        }
    }
}

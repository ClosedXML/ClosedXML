using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal struct XLSheetRange:IEquatable<XLSheetRange>
    {
        public XLSheetRange(XLSheetPoint firstPoint, XLSheetPoint lastPoint)
        {
            FirstPoint = firstPoint;
            LastPoint = lastPoint;
        }

        public XLSheetPoint FirstPoint;
        public XLSheetPoint LastPoint;

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

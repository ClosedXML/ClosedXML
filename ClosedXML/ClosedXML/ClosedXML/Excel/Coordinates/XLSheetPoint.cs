using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal struct XLSheetPoint:IEquatable<XLSheetPoint>
    {
        public XLSheetPoint(Int32  row, Int32 column)
        {
            Row = row;
            Column = column;
        }

        public Int32 Row;
        public Int32 Column;

        public bool Equals(XLSheetPoint other)
        {
            return Row == other.Row && Column == other.Column;
        }

        public override int GetHashCode()
        {
            return (Row * -1) ^ Column;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ClosedXML.Excel
{
    public struct XLCellAddress : IEqualityComparer<XLCellAddress>, IEquatable<XLCellAddress>, IComparable
    {

        public XLCellAddress(UInt32 row, UInt32 column)
        {
            this.row = row;
            this.column = column;
        }

        public XLCellAddress(String cellAddressString)
        {
            Match m = Regex.Match(cellAddressString, @"^([a-zA-Z]+)(\d+)$");
            String columnLetter = m.Groups[1].Value;
            this.row = UInt32.Parse(m.Groups[2].Value);
            this.column = XLWorksheet.ColumnLetterToNumber(columnLetter);
        }

        private UInt32 row;
        public UInt32 Row
        {
            get { return row; }
            private set { row = value; }
        }

        private UInt32 column;
        public UInt32 Column
        {
            get { return column; }
            private set { column = value; }
        }

        public static XLCellAddress operator +(XLCellAddress cellAddress1, XLCellAddress cellAddress2)
        {
            return new XLCellAddress() { Row = cellAddress1.Row + cellAddress2.Row, Column = cellAddress1.Column + cellAddress2.Column };
        }

        public static XLCellAddress operator -(XLCellAddress cellAddress1, XLCellAddress cellAddress2)
        {
            return new XLCellAddress() { Row = cellAddress1.Row - cellAddress2.Row, Column = cellAddress1.Column - cellAddress2.Column };
        }

        public static Boolean operator ==(XLCellAddress cellAddress1, XLCellAddress cellAddress2)
        {
            return
                cellAddress1.Row == cellAddress2.Row
                && cellAddress1.Column == cellAddress2.Column;
        }

        public static Boolean operator !=(XLCellAddress cellAddress1, XLCellAddress cellAddress2)
        {
            return !(cellAddress1 == cellAddress2);
        }

        public static Boolean operator >(XLCellAddress cellAddress1, XLCellAddress cellAddress2)
        {
            return !(cellAddress1 == cellAddress2)
                && cellAddress1.Row >= cellAddress2.Row && cellAddress1.Column >= cellAddress2.Column;
        }

        public static Boolean operator <(XLCellAddress cellAddress1, XLCellAddress cellAddress2)
        {
            return !(cellAddress1 == cellAddress2)
                && cellAddress1.Row <= cellAddress2.Row && cellAddress1.Column <= cellAddress2.Column;
        }

        public static Boolean operator >=(XLCellAddress cellAddress1, XLCellAddress cellAddress2)
        {
            return
                cellAddress1.Row >= cellAddress2.Row
                && cellAddress1.Column >= cellAddress2.Column;
        }

        public static Boolean operator <=(XLCellAddress cellAddress1, XLCellAddress cellAddress2)
        {
            return
                cellAddress1.Row <= cellAddress2.Row
                && cellAddress1.Column <= cellAddress2.Column;
        }

        public override String ToString()
        {
            return XLWorksheet.ColumnNumberToLetter(Column) + Row.ToString();
        }

        #region IEqualityComparer<XLCellAddress> Members

        public Boolean Equals(XLCellAddress x, XLCellAddress y)
        {
            return x == y;
        }

        public Int32 GetHashCode(XLCellAddress obj)
        {
            return obj.GetHashCode();
        }

        new public Boolean Equals(Object x, Object y)
        {
            return x == y;
        }

        public Int32 GetHashCode(Object obj)
        {
            return obj.GetHashCode();
        }

        public override Int32 GetHashCode()
        {
            return this.ToString().GetHashCode();
        }

        #endregion

        #region IEquatable<XLCellAddress> Members

        public Boolean Equals(XLCellAddress other)
        {
            return this == other;
        }

        public override Boolean Equals(Object other)
        {
            return this == (XLCellAddress)other;
        }

        #endregion

        #region IComparable Members

        public Int32 CompareTo(object obj)
        {
            var other = (XLCellAddress)obj;
            if (this == other)
                return 0;
            else if (this > other)
                return 1;
            else
                return -1;
        }

        #endregion
    }
}

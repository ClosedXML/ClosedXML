using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ClosedXML.Excel
{
    public struct XLAddress : IEqualityComparer<XLAddress>, IEquatable<XLAddress>, IComparable, IComparable<XLAddress>
    {
        #region Constructors
        public XLAddress(UInt32 row, UInt32 column)
        {
            this.row = row;
            this.column = column;
            this.columnLetter = ColumnNumberToLetter(column);
        }

        public XLAddress(UInt32 row, String columnLetter)
        {
            this.row = row;
            this.column = ColumnLetterToNumber(columnLetter);
            this.columnLetter = columnLetter;
        }

        public XLAddress(String cellAddressString)
        {
            Match m = Regex.Match(cellAddressString, @"^([a-zA-Z]+)(\d+)$");
            columnLetter = m.Groups[1].Value;
            this.row = UInt32.Parse(m.Groups[2].Value);
            this.column = ColumnLetterToNumber(columnLetter);
        }
        #endregion

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

        private String columnLetter;
        public String ColumnLetter
        {
            get { return columnLetter; }
            private set { columnLetter = value; }
        }

        public static UInt32 ColumnLetterToNumber(String column)
        {
            Int32 intColumnLetterLength = column.Length;
            Int32 retVal = 0;
            for (Int32 intCount = 0; intCount < intColumnLetterLength; intCount++)
            {
                retVal = retVal * 26 + (column.Substring(intCount, 1).ToUpper().ToCharArray()[0] - 64);
            }
            return (UInt32)retVal;
        }

        public static String ColumnNumberToLetter(UInt32 column)
        {
            String s = String.Empty;
            for (
                Int32 i = Convert.ToInt32(
                    Math.Log(
                        Convert.ToDouble(
                            25 * (
                                Convert.ToDouble(column)
                                + 1
                            )
                         )
                     ) / Math.Log(26)
                 ) - 1
                ; i >= 0
                ; i--
                )
            {
                Int32 x = Convert.ToInt32(Math.Pow(26, i + 1) - 1) / 25 - 1;
                if (column > x)
                {
                    s += (Char)(((column - x - 1) / Convert.ToInt32(Math.Pow(26, i))) % 26 + 65);
                }
            }
            return s;
        }

        public override String ToString()
        {
            return ColumnLetter + Row.ToString();
        }

        #region Operator Overloads

        public static XLAddress operator +(XLAddress xlCellAddressLeft, XLAddress xlCellAddressRight)
        {
            return new XLAddress(xlCellAddressLeft.Row + xlCellAddressRight.Row, xlCellAddressLeft.Column + xlCellAddressRight.Column);
        }

        public static XLAddress operator -(XLAddress xlCellAddressLeft, XLAddress xlCellAddressRight)
        {
            return new XLAddress(xlCellAddressLeft.Row - xlCellAddressRight.Row, xlCellAddressLeft.Column - xlCellAddressRight.Column);
        }

        public static XLAddress operator +(XLAddress xlCellAddressLeft, UInt32 right)
        {
            return new XLAddress(xlCellAddressLeft.Row + right, xlCellAddressLeft.Column + right);
        }

        public static XLAddress operator -(XLAddress xlCellAddressLeft, UInt32 right)
        {
            return new XLAddress(xlCellAddressLeft.Row - right, xlCellAddressLeft.Column - right);
        }

        public static Boolean operator ==(XLAddress xlCellAddressLeft, XLAddress xlCellAddressRight)
        {
            return
                xlCellAddressLeft.Row == xlCellAddressRight.Row
                && xlCellAddressLeft.Column == xlCellAddressRight.Column;
        }

        public static Boolean operator !=(XLAddress xlCellAddressLeft, XLAddress xlCellAddressRight)
        {
            return !(xlCellAddressLeft == xlCellAddressRight);
        }

        public static Boolean operator >(XLAddress xlCellAddressLeft, XLAddress xlCellAddressRight)
        {
            return !(xlCellAddressLeft == xlCellAddressRight)
                && (xlCellAddressLeft.Row > xlCellAddressRight.Row || xlCellAddressLeft.Column > xlCellAddressRight.Column);
        }

        public static Boolean operator <(XLAddress xlCellAddressLeft, XLAddress xlCellAddressRight)
        {
            return !(xlCellAddressLeft == xlCellAddressRight)
                && (xlCellAddressLeft.Row < xlCellAddressRight.Row || xlCellAddressLeft.Column < xlCellAddressRight.Column);
        }

        public static Boolean operator >=(XLAddress xlCellAddressLeft, XLAddress xlCellAddressRight)
        {
            return xlCellAddressLeft == xlCellAddressRight || xlCellAddressLeft > xlCellAddressRight;
        }

        public static Boolean operator <=(XLAddress xlCellAddressLeft, XLAddress xlCellAddressRight)
        {
            return xlCellAddressLeft == xlCellAddressRight || xlCellAddressLeft < xlCellAddressRight;
        }

        #endregion

        #region Interface Requirements

        #region IEqualityComparer<XLCellAddress> Members

        public Boolean Equals(XLAddress x, XLAddress y)
        {
            return x == y;
        }

        public Int32 GetHashCode(XLAddress obj)
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

        public Boolean Equals(XLAddress other)
        {
            return this == other;
        }

        public override Boolean Equals(Object other)
        {
            return this == (XLAddress)other;
        }

        #endregion

        #region IComparable Members

        public Int32 CompareTo(object obj)
        {
            var other = (XLAddress)obj;
            if (this == other)
                return 0;
            else if (this > other)
                return 1;
            else
                return -1;
        }

        #endregion

        #region IComparable<XLCellAddress> Members

        public int CompareTo(XLAddress other)
        {
            throw new NotImplementedException();
        }

        #endregion

        #endregion
    }
}

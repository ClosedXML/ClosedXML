using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ClosedXML.Excel
{
    internal struct XLAddress: IXLAddress
    {
        #region Constructors
        /// <summary>
        /// Initializes a new <see cref="XLAddress"/> struct using R1C1 notation.
        /// </summary>
        /// <param name="rowNumber">The row number of the cell address.</param>
        /// <param name="columnNumber">The column number of the cell address.</param>
        public XLAddress(Int32 rowNumber, Int32 columnNumber)
        {
            this.rowNumber = rowNumber;
            this.columnNumber = columnNumber;
            this.columnLetter = GetColumnLetterFromNumber(columnNumber);
        }

        /// <summary>
        /// Initializes a new <see cref="XLAddress"/> struct using a mixed notation.
        /// </summary>
        /// <param name="rowNumber">The row number of the cell address.</param>
        /// <param name="columnLetter">The column letter of the cell address.</param>
        public XLAddress(Int32 rowNumber, String columnLetter)
        {
            this.rowNumber = rowNumber;
            this.columnNumber = GetColumnNumberFromLetter(columnLetter);
            this.columnLetter = columnLetter;
        }


        /// <summary>
        /// Initializes a new <see cref="XLAddress"/> struct using A1 notation.
        /// </summary>
        /// <param name="cellAddressString">The cell address.</param>
        public XLAddress(String cellAddressString)
        {
            Match m = Regex.Match(cellAddressString, @"^([a-zA-Z]+)(\d+)$");
            columnLetter = m.Groups[1].Value;
            this.rowNumber = Int32.Parse(m.Groups[2].Value);
            this.columnNumber = GetColumnNumberFromLetter(columnLetter);
        }

        #endregion

        #region Static

        /// <summary>
        /// Gets the column number of a given column letter.
        /// </summary>
        /// <param name="columnLetter">The column letter to translate into a column number.</param>
        public static Int32 GetColumnNumberFromLetter(String columnLetter)
        {
            Int32 iTest;
            if (Int32.TryParse(columnLetter, out iTest)) return iTest;

            columnLetter = columnLetter.ToUpper();

            if (columnLetter.Length == 1)
            {
                return Convert.ToByte(Convert.ToChar(columnLetter)) - 64;
            }
            else if (columnLetter.Length == 2)
            {
                return
                    ((Convert.ToByte(columnLetter[0]) - 64) * 26) +
                    (Convert.ToByte(columnLetter[1]) - 64);

            }
            else if (columnLetter.Length == 3)
            {
                return
                    ((Convert.ToByte(columnLetter[0]) - 64) * 26 * 26) +
                    ((Convert.ToByte(columnLetter[1]) - 64) * 26) +
                    (Convert.ToByte(columnLetter[2]) - 64);
            }
            else
            {
                throw new ApplicationException("Column Length must be between 1 and 3.");
            }
        }

        /// <summary>
        /// Gets the column letter of a given column number.
        /// </summary>
        /// <param name="columnNumber">The column number to translate into a column letter.</param>
        public static String GetColumnLetterFromNumber(Int32 columnNumber)
        {
            String s = String.Empty;
            for (
                Int32 i = Convert.ToInt32(
                    Math.Log(
                        Convert.ToDouble(
                            25 * (
                                Convert.ToDouble(columnNumber)
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
                if (columnNumber > x)
                {
                    s += (Char)(((columnNumber - x - 1) / Convert.ToInt32(Math.Pow(26, i))) % 26 + 65);
                }
            }
            return s;
        }

        #endregion

        #region Properties

        private Int32 rowNumber;
        /// <summary>
        /// Gets the row number of this address.
        /// </summary>
        public Int32 RowNumber
        {
            get { return rowNumber; }
            private set { rowNumber = value; }
        }

        private Int32 columnNumber;
        /// <summary>
        /// Gets the column number of this address.
        /// </summary>
        public Int32 ColumnNumber
        {
            get { return columnNumber; }
            private set { columnNumber = value; }
        }

        private String columnLetter;
        /// <summary>
        /// Gets the column letter(s) of this address.
        /// </summary>
        public String ColumnLetter
        {
            get { return columnLetter; }
            private set { columnLetter = value; }
        }

        #endregion

        #region Overrides
        public override string ToString()
        {
            return this.columnLetter + this.rowNumber.ToString();
        }
        #endregion

        #region Operator Overloads

        public static XLAddress operator +(XLAddress xlCellAddressLeft, XLAddress xlCellAddressRight)
        {
            return new XLAddress(xlCellAddressLeft.RowNumber + xlCellAddressRight.RowNumber, xlCellAddressLeft.ColumnNumber + xlCellAddressRight.ColumnNumber);
        }

        public static XLAddress operator -(XLAddress xlCellAddressLeft, XLAddress xlCellAddressRight)
        {
            return new XLAddress(xlCellAddressLeft.RowNumber - xlCellAddressRight.RowNumber, xlCellAddressLeft.ColumnNumber - xlCellAddressRight.ColumnNumber);
        }

        public static XLAddress operator +(XLAddress xlCellAddressLeft, Int32 right)
        {
            return new XLAddress(xlCellAddressLeft.RowNumber + right, xlCellAddressLeft.ColumnNumber + right);
        }

        public static XLAddress operator -(XLAddress xlCellAddressLeft, Int32 right)
        {
            return new XLAddress(xlCellAddressLeft.RowNumber - right, xlCellAddressLeft.ColumnNumber - right);
        }

        public static Boolean operator ==(XLAddress xlCellAddressLeft, XLAddress xlCellAddressRight)
        {
            return
                xlCellAddressLeft.RowNumber == xlCellAddressRight.RowNumber
                && xlCellAddressLeft.ColumnNumber == xlCellAddressRight.ColumnNumber;
        }

        public static Boolean operator !=(XLAddress xlCellAddressLeft, XLAddress xlCellAddressRight)
        {
            return !(xlCellAddressLeft == xlCellAddressRight);
        }

        public static Boolean operator >(XLAddress xlCellAddressLeft, XLAddress xlCellAddressRight)
        {
            return !(xlCellAddressLeft == xlCellAddressRight)
                && (xlCellAddressLeft.RowNumber >= xlCellAddressRight.RowNumber 
                && xlCellAddressLeft.ColumnNumber >= xlCellAddressRight.ColumnNumber);
        }

        public static Boolean operator <(XLAddress xlCellAddressLeft, XLAddress xlCellAddressRight)
        {
            return !(xlCellAddressLeft == xlCellAddressRight)
                && (xlCellAddressLeft.RowNumber <= xlCellAddressRight.RowNumber 
                && xlCellAddressLeft.ColumnNumber <= xlCellAddressRight.ColumnNumber);
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

        public Boolean Equals(IXLAddress x, IXLAddress y)
        {
            return x == y;
        }

        public Int32 GetHashCode(IXLAddress obj)
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

        public Boolean Equals(IXLAddress other)
        {
            return this == (XLAddress)other;
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

        public Int32 CompareTo(IXLAddress other)
        {
            return CompareTo((Object)other);
        }

        #endregion

        #endregion
    }
}

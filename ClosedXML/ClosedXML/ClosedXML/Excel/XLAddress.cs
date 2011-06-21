using System;
using System.Diagnostics;
using System.Globalization;

namespace ClosedXML.Excel
{
    internal class XLAddress : IXLAddress
    {
        #region Static
        public static XLAddress Create(XLWorksheet worksheet, string cellAddressString)
        {
            var fixedColumn = cellAddressString[0] == '$';
            Int32 startPos;
            if (fixedColumn)
            {
                startPos = 1;
            }
            else
            {
                startPos = 0;
            }

            int rowPos = startPos;
            while (cellAddressString[rowPos] > '9')
            {
                rowPos++;
            }

            var fixedRow = cellAddressString[rowPos] == '$';
            string columnLetter;
            int rowNumber;
            if (fixedRow)
            {
                if (fixedColumn)
                {
                    columnLetter = cellAddressString.Substring(startPos, rowPos - 1);
                }
                else
                {
                    columnLetter = cellAddressString.Substring(startPos, rowPos);
                }

                rowNumber = Int32.Parse(cellAddressString.Substring(rowPos + 1), ms_nfi);
            }
            else
            {
                if (fixedColumn)
                {
                    columnLetter = cellAddressString.Substring(startPos, rowPos - 1);
                }
                else
                {
                    columnLetter = cellAddressString.Substring(startPos, rowPos);
                }

                rowNumber = Int32.Parse(cellAddressString.Substring(rowPos), ms_nfi);
            }
            return new XLAddress(worksheet, rowNumber, columnLetter, fixedRow, fixedColumn);
        }

        private static readonly NumberFormatInfo ms_nfi = CultureInfo.InvariantCulture.NumberFormat;
        private const Int32 TwoT26 = 26*26;
        /// <summary>
        /// 	Gets the column number of a given column letter.
        /// </summary>
        /// <param name = "columnLetter">The column letter to translate into a column number.</param>
        public static Int32 GetColumnNumberFromLetter(String columnLetter)
        {
            if (columnLetter[0] <= '9')
            {
                return Int32.Parse(columnLetter, ms_nfi);
            }

            columnLetter = columnLetter.ToUpper();
            var length = columnLetter.Length;
            if (length == 1)
            {
                return Convert.ToByte(columnLetter[0]) - 64;
            }
            if (length == 2)
            {
                return
                        ((Convert.ToByte(columnLetter[0]) - 64)*26) +
                        (Convert.ToByte(columnLetter[1]) - 64);
            }
            if (length == 3)
            {
                return
                        ((Convert.ToByte(columnLetter[0]) - 64)*TwoT26) +
                        ((Convert.ToByte(columnLetter[1]) - 64)*26) +
                        (Convert.ToByte(columnLetter[2]) - 64);
            }
            throw new ApplicationException("Column Length must be between 1 and 3.");
        }

        public static Boolean IsValidColumn(String column)
        {
            if (StringExtensions.IsNullOrWhiteSpace(column) || column.Length > 3)
            {
                return false;
            }

            Boolean retVal = true;
            String theColumn = column.ToUpper();
            for (Int32 i = 0; i < column.Length; i++)
            {
                if (theColumn[i] < 'A' || theColumn[i] > 'Z' || (i == 2 && theColumn[i] > 'D'))
                {
                    retVal = false;
                    break;
                }
            }
            return retVal;
        }

        public static Boolean IsValidRow(String rowString)
        {
            Int32 row;
            if (Int32.TryParse(rowString, out row))
            {
                return row > 0 && row <= XLWorksheet.MaxNumberOfRows;
            }
            return false;
        }

        public static Boolean IsValidA1Address(String address)
        {
            address = address.Replace("$", "");
            Int32 rowPos = 0;
            Int32 addressLength = address.Length;
            while (rowPos < addressLength && (address[rowPos] > '9' || address[rowPos] < '0'))
            {
                rowPos++;
            }

            return
                    rowPos < addressLength
                    && IsValidRow(address.Substring(rowPos))
                    && IsValidColumn(address.Substring(0, rowPos));
        }

        /// <summary>
        /// 	Gets the column letter of a given column number.
        /// </summary>
        /// <param name = "columnNumber">The column number to translate into a column letter.</param>
        public static String GetColumnLetterFromNumber(Int32 columnNumber)
        {
            String s = String.Empty;
            for (
                    Int32 i = Convert.ToInt32(
                            Math.Log(
                                    Convert.ToDouble(
                                            25*(
                                                       Convert.ToDouble(columnNumber)
                                                       + 1
                                               )
                                            )
                                    )/Math.Log(26)
                                      ) - 1
                    ; i >= 0
                    ; i--
                    )
            {
                Int32 x = Convert.ToInt32(Math.Pow(26, i + 1) - 1)/25 - 1;
                if (columnNumber > x)
                {
                    s += (Char) (((columnNumber - x - 1)/Convert.ToInt32(Math.Pow(26, i)))%26 + 65);
                }
            }
            return s;
        }

        public static Int32 GetRowFromAddress1(String cellAddressString)
        {
            Int32 rowPos = 1;
            while (cellAddressString[rowPos] > '9')
            {
                rowPos++;
            }

            return Int32.Parse(cellAddressString.Substring(rowPos), ms_nfi);
        }

        public static Int32 GetColumnNumberFromAddress1(String cellAddressString)
        {
            Int32 rowPos = 0;
            while (cellAddressString[rowPos] > '9')
            {
                rowPos++;
            }

            return GetColumnNumberFromLetter(cellAddressString.Substring(0, rowPos));
        }

        public static Int32 GetRowFromAddress2(String cellAddressString)
        {
            Int32 rowPos = 1;
            while (cellAddressString[rowPos] > '9')
            {
                rowPos++;
            }

            if (cellAddressString[rowPos] == '$')
            {
                return Int32.Parse(cellAddressString.Substring(rowPos + 1), ms_nfi);
            }
            return Int32.Parse(cellAddressString.Substring(rowPos), ms_nfi);
        }

        public static Int32 GetColumnNumberFromAddress2(String cellAddressString)
        {
            Int32 startPos;
            if (cellAddressString[0] == '$')
            {
                startPos = 1;
            }
            else
            {
                startPos = 0;
            }

            Int32 rowPos = startPos;
            while (cellAddressString[rowPos] > '9')
            {
                rowPos++;
            }

            if (cellAddressString[rowPos] == '$')
            {
                return GetColumnNumberFromLetter(cellAddressString.Substring(startPos, rowPos - 1));
            }
            return GetColumnNumberFromLetter(cellAddressString.Substring(startPos, rowPos));
        }
        #endregion
        #region Private fields
        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        private bool m_fixedRow;
        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        private bool m_fixedColumn;
        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        private string m_columnLetter;

        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        private readonly int m_rowNumber;
        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        private readonly int m_columnNumber;
        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        private readonly int m_hashCode;
        private String m_trimmedAddress;
        #endregion
        #region Constructors
        /// <summary>
        /// 	Initializes a new <see cref = "XLAddress" /> struct using a mixed notation.
        /// </summary>
        /// <param name = "worksheet"></param>
        /// <param name = "rowNumber">The row number of the cell address.</param>
        /// <param name = "columnLetter">The column letter of the cell address.</param>
        /// <param name = "fixedRow"></param>
        /// <param name = "fixedColumn"></param>
        public XLAddress(XLWorksheet worksheet, int rowNumber, string columnLetter, bool fixedRow, bool fixedColumn)
                : this(worksheet, rowNumber, GetColumnNumberFromLetter(columnLetter), fixedRow, fixedColumn)
        {
            m_columnLetter = columnLetter;
        }

        /// <summary>
        /// 	Initializes a new <see cref = "XLAddress" /> struct using R1C1 notation.
        /// </summary>
        /// <param name = "worksheet"></param>
        /// <param name = "rowNumber">The row number of the cell address.</param>
        /// <param name = "columnNumber">The column number of the cell address.</param>
        /// <param name = "fixedRow"></param>
        /// <param name = "fixedColumn"></param>
        public XLAddress(XLWorksheet worksheet, int rowNumber, int columnNumber, bool fixedRow, bool fixedColumn)

        {
            Worksheet = worksheet;

            m_rowNumber = rowNumber;
            m_columnNumber = columnNumber;
            m_columnLetter = null;
            m_fixedColumn = fixedColumn;
            m_fixedRow = fixedRow;

            m_hashCode = m_rowNumber ^ m_columnNumber;
        }
        #endregion
        #region Properties
        public XLWorksheet Worksheet { get; internal set; }
        IXLWorksheet IXLAddress.Worksheet
        {
            [DebuggerStepThrough]
            get { return Worksheet; }
        }

        public Boolean FixedRow
        {
            get { return m_fixedRow; }
            set { m_fixedRow = value; }
        }

        public Boolean FixedColumn
        {
            get { return m_fixedColumn; }
            set { m_fixedColumn = value; }
        }

        /// <summary>
        /// 	Gets the row number of this address.
        /// </summary>
        public Int32 RowNumber
        {
            get { return m_rowNumber; }
        }

        /// <summary>
        /// 	Gets the column number of this address.
        /// </summary>
        public Int32 ColumnNumber
        {
            get { return m_columnNumber; }
        }

        /// <summary>
        /// 	Gets the column letter(s) of this address.
        /// </summary>
        public String ColumnLetter
        {
            get { return m_columnLetter ?? (m_columnLetter = GetColumnLetterFromNumber(m_columnNumber)); }
        }
        #endregion
        #region Overrides
        public override string ToString()
        {
            String retVal = ColumnLetter;
            if (m_fixedColumn)
            {
                retVal = "$" + retVal;
            }
            if (m_fixedRow)
            {
                retVal += "$";
            }
            retVal += m_rowNumber.ToStringLookup();
            return retVal;
        }

        public String ToString(XLReferenceStyle referenceStyle)
        {
            if (referenceStyle == XLReferenceStyle.A1)
            {
                return ColumnLetter + m_rowNumber.ToStringLookup();
            }
            if (referenceStyle == XLReferenceStyle.R1C1)
            {
                return m_rowNumber.ToStringLookup() + "," + ColumnNumber;
            }
            if ((Worksheet).Internals.Workbook.ReferenceStyle == XLReferenceStyle.R1C1)
            {
                return m_rowNumber.ToStringLookup() + "," + ColumnNumber;
            }
            return ColumnLetter + m_rowNumber.ToStringLookup();
        }
        #endregion
        #region Methods
        public string GetTrimmedAddress()
        {
            return m_trimmedAddress ?? (m_trimmedAddress = ColumnLetter + m_rowNumber.ToStringLookup());
        }

        public string ToStringRelative()
        {
            return GetTrimmedAddress();
        }

        public string ToStringFixed()
        {
            return "$" + ColumnLetter + "$" + m_rowNumber.ToStringLookup();
        }
        #endregion
        #region Operator Overloads
        public static XLAddress operator +(XLAddress left, XLAddress right)
        {
            return new XLAddress(left.Worksheet,
                                 left.RowNumber + right.RowNumber,
                                 left.ColumnNumber + right.ColumnNumber,
                                 left.m_fixedRow,
                                 left.m_fixedColumn);
        }

        public static XLAddress operator -(XLAddress left, XLAddress right)
        {
            return new XLAddress(left.Worksheet,
                                 left.RowNumber - right.RowNumber,
                                 left.ColumnNumber - right.ColumnNumber,
                                 left.m_fixedRow,
                                 left.m_fixedColumn);
        }

        public static XLAddress operator +(XLAddress left, Int32 right)
        {
            return new XLAddress(left.Worksheet,
                                 left.RowNumber + right,
                                 left.ColumnNumber + right,
                                 left.m_fixedRow,
                                 left.m_fixedColumn);
        }

        public static XLAddress operator -(XLAddress left, Int32 right)
        {
            return new XLAddress(left.Worksheet,
                                 left.RowNumber - right,
                                 left.ColumnNumber - right,
                                 left.m_fixedRow,
                                 left.m_fixedColumn);
        }

        public static Boolean operator ==(XLAddress left, XLAddress right)
        {
            if (ReferenceEquals(left, right))
            {
                return true;
            }
            if (ReferenceEquals(left, null))
            {
                return false;
            }
            return left.Equals(right);
        }

        public static Boolean operator !=(XLAddress left, XLAddress right)
        {
            return !(left == right);
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

        public new Boolean Equals(Object x, Object y)
        {
            return x == y;
        }

        public Int32 GetHashCode(Object obj)
        {
            return (obj).GetHashCode();
        }

        public override Int32 GetHashCode()
        {
            return m_hashCode;
        }
        #endregion
        #region IEquatable<XLCellAddress> Members
        public bool Equals(IXLAddress other)
        {
            var right = other as XLAddress;
            if (ReferenceEquals(right, null))
            {
                return false;
            }
            if (m_hashCode != right.m_hashCode)
            {
                return false;
            }

            return m_rowNumber == right.m_rowNumber && m_columnNumber == right.m_columnNumber;
        }

        public override Boolean Equals(Object other)
        {
            return Equals((XLAddress) other);
        }
        #endregion
        #endregion
    }
}
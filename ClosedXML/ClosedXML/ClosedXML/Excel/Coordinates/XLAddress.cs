using System;
using System.Diagnostics;

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

                rowNumber = int.Parse(cellAddressString.Substring(rowPos + 1), ExcelHelper.NumberFormatForParse);
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

                rowNumber = Int32.Parse(cellAddressString.Substring(rowPos), ExcelHelper.NumberFormatForParse);
            }
            return new XLAddress(worksheet, rowNumber, columnLetter, fixedRow, fixedColumn);
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
        private string m_trimmedAddress;
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
                : this(worksheet, rowNumber, ExcelHelper.GetColumnNumberFromLetter(columnLetter), fixedRow, fixedColumn)
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

        public bool FixedRow
        {
            get { return m_fixedRow; }
            set { m_fixedRow = value; }
        }

        public bool FixedColumn
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
            get { return m_columnLetter ?? (m_columnLetter = ExcelHelper.GetColumnLetterFromNumber(m_columnNumber)); }
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
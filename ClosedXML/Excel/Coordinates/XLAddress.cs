using ClosedXML.Extensions;
using System;
using System.Diagnostics;

namespace ClosedXML.Excel
{
    internal class XLAddress : IXLAddress
    {
        #region Static
        /// <summary>
        /// Create address without worksheet. For calculation only!
        /// </summary>
        /// <param name="cellAddressString"></param>
        /// <returns></returns>
        public static XLAddress Create(string cellAddressString)
        {
            return Create(null, cellAddressString);
        }

        public static XLAddress Create(XLAddress cellAddress)
        {
            return new XLAddress(cellAddress.Worksheet, cellAddress.RowNumber, cellAddress.ColumnNumber, cellAddress.FixedRow, cellAddress.FixedColumn);
        }

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

                rowNumber = int.Parse(cellAddressString.Substring(rowPos + 1), XLHelper.NumberStyle, XLHelper.ParseCulture);
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

                rowNumber = Int32.Parse(cellAddressString.Substring(rowPos), XLHelper.NumberStyle, XLHelper.ParseCulture);
            }
            return new XLAddress(worksheet, rowNumber, columnLetter, fixedRow, fixedColumn);
        }
        #endregion
        #region Private fields
        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        private bool _fixedRow;
        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        private bool _fixedColumn;
        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        private string _columnLetter;

        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        private readonly int _rowNumber;
        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        private readonly int _columnNumber;
        private string _trimmedAddress;
        #endregion
        #region Constructors
        /// <summary>
        /// 	Initializes a new <see cref = "XLAddress" /> struct using a mixed notation.  Attention: without worksheet for calculation only!
        /// </summary>
        /// <param name = "rowNumber">The row number of the cell address.</param>
        /// <param name = "columnLetter">The column letter of the cell address.</param>
        /// <param name = "fixedRow"></param>
        /// <param name = "fixedColumn"></param>
        public XLAddress(int rowNumber, string columnLetter, bool fixedRow, bool fixedColumn)
                : this(null, rowNumber, columnLetter, fixedRow, fixedColumn)
        {
        }
        /// <summary>
        /// 	Initializes a new <see cref = "XLAddress" /> struct using a mixed notation.
        /// </summary>
        /// <param name = "worksheet"></param>
        /// <param name = "rowNumber">The row number of the cell address.</param>
        /// <param name = "columnLetter">The column letter of the cell address.</param>
        /// <param name = "fixedRow"></param>
        /// <param name = "fixedColumn"></param>
        public XLAddress(XLWorksheet worksheet, int rowNumber, string columnLetter, bool fixedRow, bool fixedColumn)
                : this(worksheet, rowNumber, XLHelper.GetColumnNumberFromLetter(columnLetter), fixedRow, fixedColumn)
        {
            _columnLetter = columnLetter;
        }

        /// <summary>
        /// 	Initializes a new <see cref = "XLAddress" /> struct using R1C1 notation. Attention: without worksheet for calculation only!
        /// </summary>
        /// <param name = "rowNumber">The row number of the cell address.</param>
        /// <param name = "columnNumber">The column number of the cell address.</param>
        /// <param name = "fixedRow"></param>
        /// <param name = "fixedColumn"></param>
        public XLAddress(int rowNumber, int columnNumber, bool fixedRow, bool fixedColumn)
                : this(null, rowNumber, columnNumber, fixedRow, fixedColumn)
        {
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

            _rowNumber = rowNumber;
            _columnNumber = columnNumber;
            _columnLetter = null;
            _fixedColumn = fixedColumn;
            _fixedRow = fixedRow;


        }
        #endregion
        #region Properties
        public XLWorksheet Worksheet { get; internal set; }
        IXLWorksheet IXLAddress.Worksheet
        {
            [DebuggerStepThrough]
            get { return Worksheet; }
        }

        public bool HasWorksheet
        {
            [DebuggerStepThrough]
            get { return Worksheet != null; }
        }

        public bool FixedRow
        {
            get { return _fixedRow; }
            set { _fixedRow = value; }
        }

        public bool FixedColumn
        {
            get { return _fixedColumn; }
            set { _fixedColumn = value; }
        }

        /// <summary>
        /// 	Gets the row number of this address.
        /// </summary>
        public Int32 RowNumber
        {
            get { return _rowNumber; }
        }

        /// <summary>
        /// 	Gets the column number of this address.
        /// </summary>
        public Int32 ColumnNumber
        {
            get { return _columnNumber; }
        }

        /// <summary>
        /// 	Gets the column letter(s) of this address.
        /// </summary>
        public String ColumnLetter
        {
            get { return _columnLetter ?? (_columnLetter = XLHelper.GetColumnLetterFromNumber(_columnNumber)); }
        }
        #endregion
        #region Overrides
        public override string ToString()
        {
            String retVal = ColumnLetter;
            if (_fixedColumn)
            {
                retVal = "$" + retVal;
            }
            if (_fixedRow)
            {
                retVal += "$";
            }
            retVal += _rowNumber.ToInvariantString();
            return retVal;
        }

        public string ToString(XLReferenceStyle referenceStyle)
        {
            if (referenceStyle == XLReferenceStyle.A1)
            {
                return ColumnLetter + _rowNumber.ToInvariantString();
            }
            if (referenceStyle == XLReferenceStyle.R1C1)
            {
                return String.Format("R{0}C{1}", _rowNumber.ToInvariantString(), ColumnNumber);
            }
            if (HasWorksheet && Worksheet.Workbook.ReferenceStyle == XLReferenceStyle.R1C1)
            {
                return String.Format("R{0}C{1}", _rowNumber.ToInvariantString(), ColumnNumber);
            }
            return ColumnLetter + _rowNumber.ToInvariantString();
        }
        #endregion
        #region Methods
        public string GetTrimmedAddress()
        {
            return _trimmedAddress ?? (_trimmedAddress = ColumnLetter + _rowNumber.ToInvariantString());
        }



        #endregion
        #region Operator Overloads
        public static XLAddress operator +(XLAddress left, XLAddress right)
        {
            return new XLAddress(left.Worksheet,
                                 left.RowNumber + right.RowNumber,
                                 left.ColumnNumber + right.ColumnNumber,
                                 left._fixedRow,
                                 left._fixedColumn);
        }

        public static XLAddress operator -(XLAddress left, XLAddress right)
        {
            return new XLAddress(left.Worksheet,
                                 left.RowNumber - right.RowNumber,
                                 left.ColumnNumber - right.ColumnNumber,
                                 left._fixedRow,
                                 left._fixedColumn);
        }

        public static XLAddress operator +(XLAddress left, Int32 right)
        {
            return new XLAddress(left.Worksheet,
                                 left.RowNumber + right,
                                 left.ColumnNumber + right,
                                 left._fixedRow,
                                 left._fixedColumn);
        }

        public static XLAddress operator -(XLAddress left, Int32 right)
        {
            return new XLAddress(left.Worksheet,
                                 left.RowNumber - right,
                                 left.ColumnNumber - right,
                                 left._fixedRow,
                                 left._fixedColumn);
        }

        public static Boolean operator ==(XLAddress left, XLAddress right)
        {
            if (ReferenceEquals(left, right))
            {
                return true;
            }
            return !ReferenceEquals(left, null) && left.Equals(right);
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

        public new Boolean Equals(object x, object y)
        {
            return x == y;
        }

        public Int32 GetHashCode(object obj)
        {
            return (obj).GetHashCode();
        }

        public override int GetHashCode()
        {
            return _rowNumber ^ _columnNumber;
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
            return _rowNumber == right._rowNumber && _columnNumber == right._columnNumber;
        }

        public override Boolean Equals(Object other)
        {
            return Equals((XLAddress) other);
        }
        #endregion
        #endregion

        public String ToStringRelative()
        {
            return ToStringRelative(false);
        }

        public String ToStringFixed()
        {
            return ToStringFixed(XLReferenceStyle.Default);
        }

        public String ToStringRelative(Boolean includeSheet)
        {
            if (includeSheet)
                return String.Format("{0}!{1}",
                    Worksheet.Name.WrapSheetNameInQuotesIfRequired(),
                    GetTrimmedAddress());

            return GetTrimmedAddress();
        }

        public String ToStringFixed(XLReferenceStyle referenceStyle)
        {
            return ToStringFixed(referenceStyle, false);
        }

        public String ToStringFixed(XLReferenceStyle referenceStyle, Boolean includeSheet)
        {
            String address;
            if (referenceStyle == XLReferenceStyle.A1)
                address = String.Format("${0}${1}", ColumnLetter, _rowNumber.ToInvariantString());
            else if (referenceStyle == XLReferenceStyle.R1C1)
                address = String.Format("R{0}C{1}", _rowNumber.ToInvariantString(), ColumnNumber);
            else if (HasWorksheet && Worksheet.Workbook.ReferenceStyle == XLReferenceStyle.R1C1)
                address = String.Format("R{0}C{1}", _rowNumber.ToInvariantString(), ColumnNumber);
            else
                address = String.Format("${0}${1}", ColumnLetter, _rowNumber.ToInvariantString());

            if (includeSheet)
                return String.Format("{0}!{1}",
                    Worksheet.Name.WrapSheetNameInQuotesIfRequired(),
                    address);

            return address;
        }

        public String UniqueId { get { return RowNumber.ToString("0000000") + ColumnNumber.ToString("00000"); } }
    }
}

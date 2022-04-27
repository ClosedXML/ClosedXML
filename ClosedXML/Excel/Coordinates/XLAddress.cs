using System;
using System.Diagnostics;

namespace ClosedXML.Excel
{
    internal struct XLAddress : IXLAddress, IEquatable<XLAddress>
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

        public static XLAddress Create(XLWorksheet worksheet, string cellAddressString)
        {
            var fixedColumn = cellAddressString[0] == '$';
            int startPos;
            if (fixedColumn)
            {
                startPos = 1;
            }
            else
            {
                startPos = 0;
            }

            var rowPos = startPos;
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

                rowNumber = int.Parse(cellAddressString.Substring(rowPos), XLHelper.NumberStyle, XLHelper.ParseCulture);
            }
            return new XLAddress(worksheet, rowNumber, columnLetter, fixedRow, fixedColumn);
        }

        #endregion Static

        #region Private fields

        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        private readonly bool _fixedRow;

        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        private readonly bool _fixedColumn;

        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        private readonly int _rowNumber;

        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        private readonly int _columnNumber;

        private string _trimmedAddress;

        #endregion Private fields

        #region Constructors

        /// <summary>
        /// Initializes a new <see cref = "XLAddress" /> struct using a mixed notation.  Attention: without worksheet for calculation only!
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
        /// Initializes a new <see cref = "XLAddress" /> struct using a mixed notation.
        /// </summary>
        /// <param name = "worksheet"></param>
        /// <param name = "rowNumber">The row number of the cell address.</param>
        /// <param name = "columnLetter">The column letter of the cell address.</param>
        /// <param name = "fixedRow"></param>
        /// <param name = "fixedColumn"></param>
        public XLAddress(XLWorksheet worksheet, int rowNumber, string columnLetter, bool fixedRow, bool fixedColumn)
                : this(worksheet, rowNumber, XLHelper.GetColumnNumberFromLetter(columnLetter), fixedRow, fixedColumn)
        {
        }

        /// <summary>
        /// Initializes a new <see cref = "XLAddress" /> struct using R1C1 notation. Attention: without worksheet for calculation only!
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
        /// Initializes a new <see cref = "XLAddress" /> struct using R1C1 notation.
        /// </summary>
        /// <param name = "worksheet"></param>
        /// <param name = "rowNumber">The row number of the cell address.</param>
        /// <param name = "columnNumber">The column number of the cell address.</param>
        /// <param name = "fixedRow"></param>
        /// <param name = "fixedColumn"></param>
        public XLAddress(XLWorksheet worksheet, int rowNumber, int columnNumber, bool fixedRow, bool fixedColumn) : this()

        {
            Worksheet = worksheet;

            _rowNumber = rowNumber;
            _columnNumber = columnNumber;
            _fixedColumn = fixedColumn;
            _fixedRow = fixedRow;
        }

        #endregion Constructors

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

        public bool FixedRow => _fixedRow;

        public bool FixedColumn => _fixedColumn;

        /// <summary>
        /// Gets the row number of this address.
        /// </summary>
        public int RowNumber => _rowNumber;

        /// <summary>
        /// Gets the column number of this address.
        /// </summary>
        public int ColumnNumber => _columnNumber;

        /// <summary>
        /// Gets the column letter(s) of this address.
        /// </summary>
        public string ColumnLetter => XLHelper.GetColumnLetterFromNumber(_columnNumber);

        #endregion Properties

        #region Overrides

        public override string ToString()
        {
            if (!IsValid)
            {
                return "#REF!";
            }

            var retVal = ColumnLetter;
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
            return ToString(referenceStyle, false);
        }

        public string ToString(XLReferenceStyle referenceStyle, bool includeSheet)
        {
            string address;
            if (!IsValid)
            {
                address = "#REF!";
            }
            else if (referenceStyle == XLReferenceStyle.A1)
            {
                address = GetTrimmedAddress();
            }
            else if (referenceStyle == XLReferenceStyle.R1C1
                     || HasWorksheet && Worksheet.Workbook.ReferenceStyle == XLReferenceStyle.R1C1)
            {
                address = "R" + _rowNumber.ToInvariantString() + "C" + ColumnNumber.ToInvariantString();
            }
            else
            {
                address = GetTrimmedAddress();
            }

            if (includeSheet)
            {
                return string.Concat(
                    WorksheetIsDeleted ? "#REF" : Worksheet.Name.EscapeSheetName(),
                    '!',
                    address);
            }

            return address;
        }

        #endregion Overrides

        #region Methods

        public string GetTrimmedAddress()
        {
            return _trimmedAddress ??= ColumnLetter + _rowNumber.ToInvariantString();
        }

        #endregion Methods

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

        public static XLAddress operator +(XLAddress left, int right)
        {
            return new XLAddress(left.Worksheet,
                                 left.RowNumber + right,
                                 left.ColumnNumber + right,
                                 left._fixedRow,
                                 left._fixedColumn);
        }

        public static XLAddress operator -(XLAddress left, int right)
        {
            return new XLAddress(left.Worksheet,
                                 left.RowNumber - right,
                                 left.ColumnNumber - right,
                                 left._fixedRow,
                                 left._fixedColumn);
        }

        public static bool operator ==(XLAddress left, XLAddress right)
        {
            return object.Equals(left, right);
        }

        public static bool operator !=(XLAddress left, XLAddress right)
        {
            return !(left == right);
        }

        #endregion Operator Overloads

        #region Interface Requirements

        #region IEqualityComparer<XLCellAddress> Members

        public bool Equals(IXLAddress x, IXLAddress y)
        {
            return x == y;
        }

        public new bool Equals(object x, object y)
        {
            return x == y;
        }

        #endregion IEqualityComparer<XLCellAddress> Members

        #region IEquatable<XLCellAddress> Members

        public bool Equals(IXLAddress other)
        {
            if (other == null)
            {
                return false;
            }

            return _rowNumber == other.RowNumber &&
                   _columnNumber == other.ColumnNumber &&
                   _fixedRow == other.FixedRow &&
                   _fixedColumn == other.FixedColumn;
        }

        public bool Equals(XLAddress other)
        {
            return _rowNumber == other._rowNumber &&
                   _columnNumber == other._columnNumber &&
                   _fixedRow == other._fixedRow &&
                   _fixedColumn == other._fixedColumn;
        }

        public override bool Equals(object other)
        {
            return Equals(other as IXLAddress);
        }

        public override int GetHashCode()
        {
            var hashCode = 2122234362;
            hashCode = hashCode * -1521134295 + _fixedRow.GetHashCode();
            hashCode = hashCode * -1521134295 + _fixedColumn.GetHashCode();
            hashCode = hashCode * -1521134295 + _rowNumber.GetHashCode();
            hashCode = hashCode * -1521134295 + _columnNumber.GetHashCode();
            return hashCode;
        }

        public int GetHashCode(IXLAddress obj)
        {
            return ((XLAddress)obj).GetHashCode();
        }

        #endregion IEquatable<XLCellAddress> Members

        #endregion Interface Requirements

        public string ToStringRelative()
        {
            return ToStringRelative(false);
        }

        public string ToStringFixed()
        {
            return ToStringFixed(XLReferenceStyle.Default);
        }

        public string ToStringRelative(bool includeSheet)
        {
            var address = IsValid ? GetTrimmedAddress() : "#REF!";

            if (includeSheet)
            {
                return string.Concat(
                    WorksheetIsDeleted ? "#REF" : Worksheet.Name.EscapeSheetName(),
                    '!',
                    address
                );
            }

            return address;
        }

        internal XLAddress WithoutWorksheet()
        {
            return new XLAddress(RowNumber, ColumnNumber, FixedRow, FixedColumn);
        }

        public string ToStringFixed(XLReferenceStyle referenceStyle)
        {
            return ToStringFixed(referenceStyle, false);
        }

        public string ToStringFixed(XLReferenceStyle referenceStyle, bool includeSheet)
        {
            string address;

            if (referenceStyle == XLReferenceStyle.Default && HasWorksheet)
            {
                referenceStyle = Worksheet.Workbook.ReferenceStyle;
            }

            if (referenceStyle == XLReferenceStyle.Default)
            {
                referenceStyle = XLReferenceStyle.A1;
            }

            Debug.Assert(referenceStyle != XLReferenceStyle.Default);

            if (!IsValid)
            {
                address = "#REF!";
            }
            else
            {
                switch (referenceStyle)
                {
                    case XLReferenceStyle.A1:
                        address = string.Concat('$', ColumnLetter, '$', _rowNumber.ToInvariantString());
                        break;

                    case XLReferenceStyle.R1C1:
                        address = string.Concat('R', _rowNumber.ToInvariantString(), 'C', ColumnNumber);
                        break;

                    default:
                        throw new NotImplementedException();
                }
            }

            if (includeSheet)
            {
                return string.Concat(
                    WorksheetIsDeleted ? "#REF" : Worksheet.Name.EscapeSheetName(),
                    '!',
                    address);
            }

            return address;
        }

        public string UniqueId => RowNumber.ToString("0000000") + ColumnNumber.ToString("00000");

        public bool IsValid => 0 < RowNumber && RowNumber <= XLHelper.MaxRowNumber &&
                       0 < ColumnNumber && ColumnNumber <= XLHelper.MaxColumnNumber;

        private bool WorksheetIsDeleted => Worksheet?.IsDeleted == true;
    }
}

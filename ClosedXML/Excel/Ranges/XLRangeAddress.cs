using ClosedXML.Extensions;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace ClosedXML.Excel
{
    internal struct XLRangeAddress : IXLRangeAddress, IEquatable<XLRangeAddress>
    {
        #region Static members

        public static XLRangeAddress EntireColumn(XLWorksheet worksheet, int column)
        {
            return new XLRangeAddress(
                new XLAddress(worksheet, 1, column, false, false),
                new XLAddress(worksheet, XLHelper.MaxRowNumber, column, false, false));
        }

        public static XLRangeAddress EntireRow(XLWorksheet worksheet, int row)
        {
            return new XLRangeAddress(
                new XLAddress(worksheet, row, 1, false, false),
                new XLAddress(worksheet, row, XLHelper.MaxColumnNumber, false, false));
        }

        #endregion Static members

        #region Private fields

        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        private XLAddress _firstAddress;

        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        private XLAddress _lastAddress;

        #endregion Private fields

        #region Constructor

        public XLRangeAddress(XLAddress firstAddress, XLAddress lastAddress) : this()
        {
            Worksheet = firstAddress.Worksheet;
            _firstAddress = firstAddress;
            _lastAddress = lastAddress;
        }

        public XLRangeAddress(XLWorksheet worksheet, String rangeAddress) : this()
        {
            string addressToUse = rangeAddress.Contains("!")
                                      ? rangeAddress.Substring(rangeAddress.IndexOf("!") + 1)
                                      : rangeAddress;

            string firstPart;
            string secondPart;
            if (addressToUse.Contains(':'))
            {
                var arrRange = addressToUse.Split(':');
                firstPart = arrRange[0];
                secondPart = arrRange[1];
            }
            else
            {
                firstPart = addressToUse;
                secondPart = addressToUse;
            }

            if (XLHelper.IsValidA1Address(firstPart))
            {
                _firstAddress = XLAddress.Create(worksheet, firstPart);
                _lastAddress = XLAddress.Create(worksheet, secondPart);
            }
            else
            {
                firstPart = firstPart.Replace("$", String.Empty);
                secondPart = secondPart.Replace("$", String.Empty);
                if (char.IsDigit(firstPart[0]))
                {
                    _firstAddress = XLAddress.Create(worksheet, "A" + firstPart);
                    _lastAddress = XLAddress.Create(worksheet, XLHelper.MaxColumnLetter + secondPart);
                }
                else
                {
                    _firstAddress = XLAddress.Create(worksheet, firstPart + "1");
                    _lastAddress = XLAddress.Create(worksheet, secondPart + XLHelper.MaxRowNumber.ToInvariantString());
                }
            }

            Worksheet = worksheet;
        }

        #endregion Constructor

        #region Public properties

        public XLWorksheet Worksheet { get; internal set; }

        public XLAddress FirstAddress
        {
            get
            {
                //if (!IsValid)
                //    throw new InvalidOperationException("Range is invalid.");
                return _firstAddress;
            }
        }

        public XLAddress LastAddress
        {
            get
            {
                //if (!IsValid)
                //    throw new InvalidOperationException("Range is invalid.");
                return _lastAddress;
            }
        }

        IXLWorksheet IXLRangeAddress.Worksheet
        {
            get { return Worksheet; }
        }

        IXLAddress IXLRangeAddress.FirstAddress
        {
            [DebuggerStepThrough]
            get { return FirstAddress; }
        }

        IXLAddress IXLRangeAddress.LastAddress
        {
            [DebuggerStepThrough]
            get { return LastAddress; }
        }

        public bool IsValid
        {
            get
            {
                return _firstAddress.IsValid &&
                       _lastAddress.IsValid;
            }
        }

        #endregion Public properties

        #region Public methods

        /// <summary>
        /// Lead a range address to a normal form - when <see cref="FirstAddress"/> points to the top-left address and
        /// <see cref="LastAddress"/> points to the bottom-right address.
        /// </summary>
        /// <returns></returns>
        public XLRangeAddress Normalize()
        {
            if (FirstAddress.RowNumber <= LastAddress.RowNumber &&
                FirstAddress.ColumnNumber <= LastAddress.ColumnNumber)
                return this;

            int firstRow, firstColumn, lastRow, lastColumn;
            bool firstRowFixed, firstColumnFixed, lastRowFixed, lastColumnFixed;

            if (FirstAddress.RowNumber <= LastAddress.RowNumber)
            {
                firstRow = FirstAddress.RowNumber;
                firstRowFixed = FirstAddress.FixedRow;
                lastRow = LastAddress.RowNumber;
                lastRowFixed = LastAddress.FixedRow;
            }
            else
            {
                firstRow = LastAddress.RowNumber;
                firstRowFixed = LastAddress.FixedRow;
                lastRow = FirstAddress.RowNumber;
                lastRowFixed = FirstAddress.FixedRow;
            }

            if (FirstAddress.ColumnNumber <= LastAddress.ColumnNumber)
            {
                firstColumn = FirstAddress.ColumnNumber;
                firstColumnFixed = FirstAddress.FixedColumn;
                lastColumn = LastAddress.ColumnNumber;
                lastColumnFixed = LastAddress.FixedColumn;
            }
            else
            {
                firstColumn = LastAddress.ColumnNumber;
                firstColumnFixed = LastAddress.FixedColumn;
                lastColumn = FirstAddress.ColumnNumber;
                lastColumnFixed = FirstAddress.FixedColumn;
            }

            return new XLRangeAddress(
                new XLAddress(FirstAddress.Worksheet, firstRow, firstColumn, firstRowFixed, firstColumnFixed),
                new XLAddress(LastAddress.Worksheet, lastRow, lastColumn, lastRowFixed, lastColumnFixed));
        }

        public String ToStringRelative()
        {
            return ToStringRelative(false);
        }

        public String ToStringFixed()
        {
            return ToStringFixed(XLReferenceStyle.A1);
        }

        public String ToStringRelative(Boolean includeSheet)
        {
            if (includeSheet)
                return String.Concat(
                    Worksheet.Name.EscapeSheetName(),
                    '!',
                    _firstAddress.ToStringRelative(),
                    ':',
                    _lastAddress.ToStringRelative());
            else
                return string.Concat(
                    _firstAddress.ToStringRelative(),
                    ":",
                    _lastAddress.ToStringRelative());
        }

        public String ToStringFixed(XLReferenceStyle referenceStyle)
        {
            return ToStringFixed(referenceStyle, false);
        }

        public String ToStringFixed(XLReferenceStyle referenceStyle, Boolean includeSheet)
        {
            if (includeSheet)
                return String.Format("{0}!{1}:{2}",
                    Worksheet.Name.EscapeSheetName(),
                    _firstAddress.ToStringFixed(referenceStyle),
                    _lastAddress.ToStringFixed(referenceStyle));

            return _firstAddress.ToStringFixed(referenceStyle) + ":" + _lastAddress.ToStringFixed(referenceStyle);
        }

        public override string ToString()
        {
            return String.Concat(_firstAddress, ':', _lastAddress);
        }

        public string ToString(XLReferenceStyle referenceStyle)
        {
            return ToString(referenceStyle, false);
        }

        public string ToString(XLReferenceStyle referenceStyle, bool includeSheet)
        {
            if (referenceStyle == XLReferenceStyle.R1C1)
                return ToStringFixed(referenceStyle, true);
            else
                return ToStringRelative(includeSheet);
        }

        public override bool Equals(object obj)
        {
            if (!(obj is XLRangeAddress))
            {
                return false;
            }

            var address = (XLRangeAddress)obj;
            return _firstAddress.Equals(address._firstAddress) &&
                   _lastAddress.Equals(address._lastAddress) &&
                   EqualityComparer<XLWorksheet>.Default.Equals(Worksheet, address.Worksheet);
        }

        public override int GetHashCode()
        {
            var hashCode = -778064135;
            hashCode = hashCode * -1521134295 + _firstAddress.GetHashCode();
            hashCode = hashCode * -1521134295 + _lastAddress.GetHashCode();
            hashCode = hashCode * -1521134295 + EqualityComparer<XLWorksheet>.Default.GetHashCode(Worksheet);
            return hashCode;
        }

        public bool Equals(XLRangeAddress other)
        {
            return ReferenceEquals(Worksheet, other.Worksheet) &&
                   _firstAddress == other._firstAddress &&
                   _lastAddress == other._lastAddress;
        }

        #endregion Public methods
    }
}

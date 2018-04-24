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

        #region Constructor

        public XLRangeAddress(XLAddress firstAddress, XLAddress lastAddress) : this()
        {
            Worksheet = firstAddress.Worksheet;
            FirstAddress = firstAddress;
            LastAddress = lastAddress;
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
                FirstAddress = XLAddress.Create(worksheet, firstPart);
                LastAddress = XLAddress.Create(worksheet, secondPart);
            }
            else
            {
                firstPart = firstPart.Replace("$", String.Empty);
                secondPart = secondPart.Replace("$", String.Empty);
                if (char.IsDigit(firstPart[0]))
                {
                    FirstAddress = XLAddress.Create(worksheet, "A" + firstPart);
                    LastAddress = XLAddress.Create(worksheet, XLHelper.MaxColumnLetter + secondPart);
                }
                else
                {
                    FirstAddress = XLAddress.Create(worksheet, firstPart + "1");
                    LastAddress = XLAddress.Create(worksheet, secondPart + XLHelper.MaxRowNumber.ToInvariantString());
                }
            }

            Worksheet = worksheet;
        }

        #endregion Constructor

        #region Public properties

        public XLWorksheet Worksheet { get; }

        public XLAddress FirstAddress { get; }

        public XLAddress LastAddress { get; }

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
                return FirstAddress.IsValid &&
                       LastAddress.IsValid;
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
                    FirstAddress.ToStringRelative(),
                    ':',
                    LastAddress.ToStringRelative());
            else
                return string.Concat(
                    FirstAddress.ToStringRelative(),
                    ":",
                    LastAddress.ToStringRelative());
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
                    FirstAddress.ToStringFixed(referenceStyle),
                    LastAddress.ToStringFixed(referenceStyle));

            return FirstAddress.ToStringFixed(referenceStyle) + ":" + LastAddress.ToStringFixed(referenceStyle);
        }

        public override string ToString()
        {
            return String.Concat(FirstAddress, ':', LastAddress);
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
            return FirstAddress.Equals(address.FirstAddress) &&
                   LastAddress.Equals(address.LastAddress) &&
                   EqualityComparer<XLWorksheet>.Default.Equals(Worksheet, address.Worksheet);
        }

        public override int GetHashCode()
        {
            var hashCode = -778064135;
            hashCode = hashCode * -1521134295 + FirstAddress.GetHashCode();
            hashCode = hashCode * -1521134295 + LastAddress.GetHashCode();
            hashCode = hashCode * -1521134295 + EqualityComparer<XLWorksheet>.Default.GetHashCode(Worksheet);
            return hashCode;
        }

        public bool Equals(XLRangeAddress other)
        {
            return ReferenceEquals(Worksheet, other.Worksheet) &&
                   FirstAddress == other.FirstAddress &&
                   LastAddress == other.LastAddress;
        }

        #endregion Public methods

        #region Operators

        public static bool operator ==(XLRangeAddress left, XLRangeAddress right)
        {
            return left.Equals(right);
        }

        public static bool operator !=(XLRangeAddress left, XLRangeAddress right)
        {
            return !(left == right);
        }
        #endregion
    }
}

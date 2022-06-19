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

        public static readonly XLRangeAddress Invalid = new XLRangeAddress(
            new XLAddress(-1, -1, fixedRow: true, fixedColumn: true),
            new XLAddress(-1, -1, fixedRow: true, fixedColumn: true)
        );

        #endregion Static members

        #region Constructor

        public XLRangeAddress(XLAddress firstAddress, XLAddress lastAddress) : this()
        {
            Worksheet = firstAddress.Worksheet;
            FirstAddress = firstAddress;
            LastAddress = lastAddress;
        }

        public XLRangeAddress(XLWorksheet worksheet, string rangeAddress) : this()
        {
            var addressToUse = rangeAddress.Contains("!")
                ? rangeAddress.Substring(rangeAddress.LastIndexOf("!") + 1)
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
                firstPart = firstPart.Replace("$", string.Empty);
                secondPart = secondPart.Replace("$", string.Empty);
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

        IXLWorksheet IXLRangeAddress.Worksheet => Worksheet;

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

        public bool IsValid => FirstAddress.IsValid && LastAddress.IsValid;

        public int ColumnSpan
        {
            get
            {
                if (!IsValid)
                {
                    throw new InvalidOperationException("Range address is invalid.");
                }

                return Math.Abs(LastAddress.ColumnNumber - FirstAddress.ColumnNumber) + 1;
            }
        }

        public int NumberOfCells => ColumnSpan * RowSpan;

        public int RowSpan
        {
            get
            {
                if (!IsValid)
                {
                    throw new InvalidOperationException("Range address is invalid.");
                }

                return Math.Abs(LastAddress.RowNumber - FirstAddress.RowNumber) + 1;
            }
        }

        private bool WorksheetIsDeleted => Worksheet?.IsDeleted == true;

        #endregion Public properties

        #region Public methods

        public bool IsNormalized => LastAddress.RowNumber >= FirstAddress.RowNumber
                                       && LastAddress.ColumnNumber >= FirstAddress.ColumnNumber;

        /// <summary>
        /// Lead a range address to a normal form - when <see cref="FirstAddress"/> points to the top-left address and
        /// <see cref="LastAddress"/> points to the bottom-right address.
        /// </summary>
        /// <returns></returns>
        public XLRangeAddress Normalize()
        {
            if (FirstAddress.RowNumber <= LastAddress.RowNumber &&
                FirstAddress.ColumnNumber <= LastAddress.ColumnNumber)
            {
                return this;
            }

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

        public bool Intersects(IXLRangeAddress otherAddress)
        {
            var xlOtherAddress = (XLRangeAddress)otherAddress;
            return Intersects(in xlOtherAddress);
        }

        internal bool Intersects(in XLRangeAddress otherAddress)
        {
            return !( // See if the two ranges intersect...
                       otherAddress.FirstAddress.ColumnNumber > LastAddress.ColumnNumber
                    || otherAddress.LastAddress.ColumnNumber < FirstAddress.ColumnNumber
                    || otherAddress.FirstAddress.RowNumber > LastAddress.RowNumber
                    || otherAddress.LastAddress.RowNumber < FirstAddress.RowNumber
                );
        }

        public bool Contains(IXLAddress address)
        {
            var xlAddress = (XLAddress)address;
            return Contains(in xlAddress);
        }

        internal IXLRangeAddress WithoutWorksheet()
        {
            return new XLRangeAddress(
                FirstAddress.WithoutWorksheet(),
                LastAddress.WithoutWorksheet());
        }

        internal bool Contains(in XLAddress address)
        {
            return FirstAddress.RowNumber <= address.RowNumber &&
                   address.RowNumber <= LastAddress.RowNumber &&
                   FirstAddress.ColumnNumber <= address.ColumnNumber &&
                   address.ColumnNumber <= LastAddress.ColumnNumber;
        }

        public string ToStringRelative()
        {
            return ToStringRelative(false);
        }

        public string ToStringFixed()
        {
            return ToStringFixed(XLReferenceStyle.A1);
        }

        public string ToStringRelative(bool includeSheet)
        {
            string address;
            if (!IsValid)
            {
                address = "#REF!";
            }
            else
            {
                if (IsEntireSheet())
                {
                    address = $"1:{XLHelper.MaxRowNumber}";
                }
                else if (IsEntireRow())
                {
                    address = string.Concat(FirstAddress.RowNumber.ToString(), ":", LastAddress.RowNumber.ToString());
                }
                else if (IsEntireColumn())
                {
                    address = string.Concat(FirstAddress.ColumnLetter, ":", LastAddress.ColumnLetter);
                }
                else
                {
                    address = string.Concat(FirstAddress.ToStringRelative(), ":", LastAddress.ToStringRelative());
                }
            }

            if (includeSheet || WorksheetIsDeleted)
            {
                return string.Concat(
                    WorksheetIsDeleted ? "#REF" : Worksheet.Name.EscapeSheetName(),
                    "!", address);
            }

            return address;
        }

        public string ToStringFixed(XLReferenceStyle referenceStyle)
        {
            return ToStringFixed(referenceStyle, false);
        }

        public string ToStringFixed(XLReferenceStyle referenceStyle, bool includeSheet)
        {
            string address;
            if (!IsValid)
            {
                address = "#REF!";
            }
            else
            {
                if (IsEntireSheet())
                {
                    address = $"$1:${XLHelper.MaxRowNumber}";
                }
                else if (IsEntireRow())
                {
                    address = string.Concat("$", FirstAddress.RowNumber.ToString(), ":$", LastAddress.RowNumber.ToString());
                }
                else if (IsEntireColumn())
                {
                    address = string.Concat("$", FirstAddress.ColumnLetter, ":$", LastAddress.ColumnLetter);
                }
                else
                {
                    address = string.Concat(FirstAddress.ToStringFixed(referenceStyle), ":", LastAddress.ToStringFixed(referenceStyle));
                }
            }

            if (includeSheet || WorksheetIsDeleted)
            {
                return string.Concat(
                    WorksheetIsDeleted ? "#REF" : Worksheet.Name.EscapeSheetName(),
                    "!", address);
            }

            return address;
        }

        public override string ToString()
        {
            if (!IsValid || WorksheetIsDeleted)
            {
                var worksheet = WorksheetIsDeleted ? "#REF!" : "";

                var address = (!FirstAddress.IsValid || !LastAddress.IsValid) ? "#REF!" : string.Concat(FirstAddress.ToString(), ":", LastAddress.ToString());
                return string.Concat(worksheet, address);
            }

            if (IsEntireSheet())
            {
                var worksheet = WorksheetIsDeleted ? "#REF!" : "";
                var address = $"1:{XLHelper.MaxRowNumber}";
                return string.Concat(worksheet, address);
            }
            else if (IsEntireRow())
            {
                var worksheet = WorksheetIsDeleted ? "#REF!" : "";
                var firstAddress = FirstAddress.IsValid ? FirstAddress.RowNumber.ToString() : "#REF!";
                var lastAddress = LastAddress.IsValid ? LastAddress.RowNumber.ToString() : "#REF!";

                return string.Concat(worksheet, firstAddress, ':', lastAddress);
            }
            else if (IsEntireColumn())
            {
                var worksheet = WorksheetIsDeleted ? "#REF!" : "";
                var firstAddress = FirstAddress.IsValid ? FirstAddress.ColumnLetter : "#REF!";
                var lastAddress = LastAddress.IsValid ? LastAddress.ColumnLetter : "#REF!";

                return string.Concat(worksheet, firstAddress, ':', lastAddress);
            }
            else
            {
                return string.Concat(FirstAddress.ToString(), ":", LastAddress.ToString());
            }
        }

        public string ToString(XLReferenceStyle referenceStyle)
        {
            return ToString(referenceStyle, false);
        }

        public string ToString(XLReferenceStyle referenceStyle, bool includeSheet)
        {
            if (referenceStyle == XLReferenceStyle.R1C1)
            {
                return ToStringFixed(referenceStyle, true);
            }
            else
            {
                return ToStringRelative(includeSheet);
            }
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

        public bool IsEntireColumn()
        {
            return IsValid
                   && FirstAddress.RowNumber == 1
                   && LastAddress.RowNumber == XLHelper.MaxRowNumber;
        }

        public bool IsEntireRow()
        {
            return IsValid
                   && FirstAddress.ColumnNumber == 1
                   && LastAddress.ColumnNumber == XLHelper.MaxColumnNumber;
        }

        public bool IsEntireSheet()
        {
            return IsValid && IsEntireColumn() && IsEntireRow();
        }

        public IXLRangeAddress Relative(IXLRangeAddress sourceRangeAddress, IXLRangeAddress targetRangeAddress)
        {
            var xlSourceRangeAddress = (XLRangeAddress)sourceRangeAddress;
            var xlTargetRangeAddress = (XLRangeAddress)targetRangeAddress;

            return Relative(in xlSourceRangeAddress, in xlTargetRangeAddress);
        }

        internal XLRangeAddress Relative(in XLRangeAddress sourceRangeAddress, in XLRangeAddress targetRangeAddress)
        {
            var sheet = targetRangeAddress.Worksheet;

            return new XLRangeAddress
            (
                new XLAddress
                (
                    sheet,
                    FirstAddress.RowNumber - sourceRangeAddress.FirstAddress.RowNumber + targetRangeAddress.FirstAddress.RowNumber,
                    FirstAddress.ColumnNumber - sourceRangeAddress.FirstAddress.ColumnNumber + targetRangeAddress.FirstAddress.ColumnNumber,
                    fixedRow: false,
                    fixedColumn: false
                ),
                new XLAddress
                (
                    sheet,
                    LastAddress.RowNumber - sourceRangeAddress.FirstAddress.RowNumber + targetRangeAddress.FirstAddress.RowNumber,
                    LastAddress.ColumnNumber - sourceRangeAddress.FirstAddress.ColumnNumber + targetRangeAddress.FirstAddress.ColumnNumber,
                    fixedRow: false,
                    fixedColumn: false
                )
            );
        }

        public IXLRangeAddress Intersection(IXLRangeAddress otherRangeAddress)
        {
            if (otherRangeAddress == null)
            {
                throw new ArgumentNullException(nameof(otherRangeAddress));
            }

            var xlOtherRangeAddress = (XLRangeAddress)otherRangeAddress;
            return Intersection(in xlOtherRangeAddress);
        }

        internal XLRangeAddress Intersection(in XLRangeAddress otherRangeAddress)
        {
            if (!Worksheet.Equals(otherRangeAddress.Worksheet))
            {
                throw new ArgumentOutOfRangeException(nameof(otherRangeAddress), "The other range address is on a different worksheet");
            }

            var thisRangeAddressNormalized = Normalize();
            var otherRangeAddressNormalized = otherRangeAddress.Normalize();

            var firstRow = Math.Max(thisRangeAddressNormalized.FirstAddress.RowNumber, otherRangeAddressNormalized.FirstAddress.RowNumber);
            var firstColumn = Math.Max(thisRangeAddressNormalized.FirstAddress.ColumnNumber, otherRangeAddressNormalized.FirstAddress.ColumnNumber);
            var lastRow = Math.Min(thisRangeAddressNormalized.LastAddress.RowNumber, otherRangeAddressNormalized.LastAddress.RowNumber);
            var lastColumn = Math.Min(thisRangeAddressNormalized.LastAddress.ColumnNumber, otherRangeAddressNormalized.LastAddress.ColumnNumber);

            if (lastRow < firstRow || lastColumn < firstColumn)
            {
                return Invalid;
            }

            return new XLRangeAddress
            (
                new XLAddress(Worksheet, firstRow, firstColumn, fixedRow: false, fixedColumn: false),
                new XLAddress(Worksheet, lastRow, lastColumn, fixedRow: false, fixedColumn: false)
            );
        }

        public IXLRange AsRange()
        {
            if (Worksheet == null)
            {
                throw new InvalidOperationException("The worksheet of the current range address has not been set.");
            }

            if (!IsValid)
            {
                return null;
            }

            return Worksheet.Range(this);
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

        #endregion Operators
    }
}

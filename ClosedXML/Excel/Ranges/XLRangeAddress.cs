using ClosedXML.Extensions;
using System;
using System.Diagnostics;
using System.Globalization;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLRangeAddress : IXLRangeAddress
    {
        #region Private fields

        [DebuggerBrowsable(DebuggerBrowsableState.Never)] private XLAddress _firstAddress;
        [DebuggerBrowsable(DebuggerBrowsableState.Never)] private XLAddress _lastAddress;

        #endregion

        #region Constructor

        public XLRangeAddress(XLRangeAddress rangeAddress): this(rangeAddress.FirstAddress, rangeAddress.LastAddress)
        {

        }

        public XLRangeAddress(XLAddress firstAddress, XLAddress lastAddress)
        {
            Worksheet = firstAddress.Worksheet;
            FirstAddress = XLAddress.Create(firstAddress);
            LastAddress = XLAddress.Create(lastAddress);
        }

        public XLRangeAddress(XLWorksheet worksheet, String rangeAddress)
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

        #endregion

        #region Public properties

        public XLWorksheet Worksheet { get; internal set; }

        public XLAddress FirstAddress
        {
            get
            {
                if (IsInvalid)
                    throw new Exception("Range is invalid.");

                return _firstAddress;
            }
            set { _firstAddress = value; }
        }

        public XLAddress LastAddress
        {
            get
            {
                if (IsInvalid)
                    throw new Exception("Range is an invalid state.");

                return _lastAddress;
            }
            set { _lastAddress = value; }
        }

        IXLWorksheet IXLRangeAddress.Worksheet
        {
            get { return Worksheet; }
        }

        IXLAddress IXLRangeAddress.FirstAddress
        {
            [DebuggerStepThrough]
            get { return FirstAddress; }
            set { FirstAddress = value as XLAddress; }
        }

        IXLAddress IXLRangeAddress.LastAddress
        {
            [DebuggerStepThrough]
            get { return LastAddress; }
            set { LastAddress = value as XLAddress; }
        }


        public bool IsInvalid { get; set; }

        #endregion

        #region Public methods

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
                return String.Format("{0}!{1}:{2}",
                    Worksheet.Name.WrapSheetNameInQuotesIfRequired(),
                    _firstAddress.ToStringRelative(),
                    _lastAddress.ToStringRelative());

            return _firstAddress.ToStringRelative() + ":" + _lastAddress.ToStringRelative();
        }

        public String ToStringFixed(XLReferenceStyle referenceStyle)
        {
            return ToStringFixed(referenceStyle, false);
        }

        public String ToStringFixed(XLReferenceStyle referenceStyle, Boolean includeSheet)
        {
            if (includeSheet)
                return String.Format("{0}!{1}:{2}",
                    Worksheet.Name.WrapSheetNameInQuotesIfRequired(),
                    _firstAddress.ToStringFixed(referenceStyle),
                    _lastAddress.ToStringFixed(referenceStyle));

            return _firstAddress.ToStringFixed(referenceStyle) + ":" + _lastAddress.ToStringFixed(referenceStyle);
        }

        public override string ToString()
        {
            return _firstAddress + ":" + _lastAddress;
        }

        public override bool Equals(object obj)
        {
            var other = (XLRangeAddress)obj;
            return Worksheet.Equals(other.Worksheet)
                   && FirstAddress.Equals(other.FirstAddress)
                   && LastAddress.Equals(other.LastAddress);
        }

        public override int GetHashCode()
        {
            return
                Worksheet.GetHashCode()
                ^ FirstAddress.GetHashCode()
                ^ LastAddress.GetHashCode();
        }

        #endregion
    }
}

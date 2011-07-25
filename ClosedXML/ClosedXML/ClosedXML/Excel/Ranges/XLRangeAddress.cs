using System;
using System.Diagnostics;
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

        public XLRangeAddress(XLAddress firstAddress, XLAddress lastAddress)
        {
            if (firstAddress.Worksheet != lastAddress.Worksheet)
                throw new ArgumentException("First and last addresses must be in the same worksheet");

            Worksheet = firstAddress.Worksheet;
            FirstAddress = firstAddress;
            LastAddress = lastAddress;
        }

        public XLRangeAddress(XLWorksheet worksheet, String rangeAddress)
        {
            string addressToUse = rangeAddress.Contains("!")
                                      ? rangeAddress.Substring(rangeAddress.IndexOf("!") + 1)
                                      : rangeAddress;

            XLAddress firstAddress;
            XLAddress lastAddress;
            if (addressToUse.Contains(':'))
            {
                var arrRange = addressToUse.Split(':');
                string firstPart = arrRange[0];
                string secondPart = arrRange[1];
                firstAddress = XLAddress.Create(worksheet, firstPart);
                lastAddress = XLAddress.Create(worksheet, secondPart);
            }
            else
            {
                firstAddress = XLAddress.Create(worksheet, addressToUse);
                lastAddress = XLAddress.Create(worksheet, addressToUse);
            }
            FirstAddress = firstAddress;
            LastAddress = lastAddress;
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
                return String.Format("'{0}'!{1}:{2}", 
                    Worksheet.Name, 
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
                return String.Format("'{0}'!{1}:{2}",
                    Worksheet.Name,
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
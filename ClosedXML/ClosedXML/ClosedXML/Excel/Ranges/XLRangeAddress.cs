using System;
using System.Diagnostics;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLRangeAddress : IXLRangeAddress
    {
        #region Private fields
        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        private XLAddress m_firstAddress;
        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        private XLAddress m_lastAddress;
        #endregion
        #region Constructor
        public XLRangeAddress(XLAddress firstAddress, XLAddress lastAddress)
        {
            if (firstAddress.Worksheet != lastAddress.Worksheet)
            {
                throw new ArgumentException("First and last addresses must be in the same worksheet");
            }

            Worksheet = firstAddress.Worksheet;
            FirstAddress = firstAddress;
            LastAddress = lastAddress;
        }

        public XLRangeAddress(XLWorksheet worksheet, String rangeAddress)
        {
            String addressToUse;
            if (rangeAddress.Contains("!"))
            {
                addressToUse = rangeAddress.Substring(rangeAddress.IndexOf("!") + 1);
            }
            else
            {
                addressToUse = rangeAddress;
            }

            XLAddress firstAddress;
            XLAddress lastAddress;
            if (addressToUse.Contains(':'))
            {
                String[] arrRange = addressToUse.Split(':');
                var firstPart = arrRange[0];
                var secondPart = arrRange[1];
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
        IXLWorksheet IXLRangeAddress.Worksheet { get { return Worksheet; } }
        
        public XLAddress FirstAddress
        {
            get
            {
                if (IsInvalid)
                {
                    throw new Exception("Range is invalid.");
                }

                return m_firstAddress;
            }
            set { m_firstAddress = value; }
        }

        IXLAddress IXLRangeAddress.FirstAddress
        {
            [DebuggerStepThrough]
            get { return FirstAddress; }
            set { FirstAddress = value as XLAddress; }
        }

        public XLAddress LastAddress
        {
            get
            {
                if (IsInvalid)
                {
                    throw new Exception("Range is an invalid state.");
                }

                return m_lastAddress;
            }
            set { m_lastAddress = value; }
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

        public override string ToString()
        {
            return m_firstAddress + ":" + m_lastAddress;
        }

        public String ToStringRelative()
        {
            return m_firstAddress.ToStringRelative() + ":" + m_lastAddress.ToStringRelative();
        }
        public String ToStringFixed()
        {
            return m_firstAddress.ToStringFixed() + ":" + m_lastAddress.ToStringFixed();
        }

        public override bool Equals(object obj)
        {
            var other = (XLRangeAddress) obj;
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
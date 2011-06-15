using System;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLRangeAddress : IXLRangeAddress
    {
        #region Private fields
        private IXLAddress m_firstAddress;
        private IXLAddress m_lastAddress;
        #endregion
        #region Constructor
        public XLRangeAddress(IXLAddress firstAddress, IXLAddress lastAddress)
        {
            if (firstAddress.Worksheet != lastAddress.Worksheet)
            {
                throw new ArgumentException("First and last addresses must be in the same worksheet");
            }

            Worksheet = firstAddress.Worksheet;
            FirstAddress = firstAddress;
            LastAddress = lastAddress;
        }

        public XLRangeAddress(IXLWorksheet worksheet, String rangeAddress)
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
        public IXLWorksheet Worksheet { get; internal set; }

        public IXLAddress FirstAddress
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

        public IXLAddress LastAddress
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

        public Boolean IsInvalid { get; set; }
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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLRangeAddress: IXLRangeAddress
    {
        public IXLWorksheet Worksheet { get; internal set; }

        private IXLAddress firstAddress;
        public IXLAddress FirstAddress
        {
            get 
            {
                if (IsInvalid)
                    throw new Exception("Range is invalid.");

                return firstAddress; 
            }
            set { firstAddress = value; }
        }

        private IXLAddress lastAddress;
        public IXLAddress LastAddress
        {
            get 
            {
                if (IsInvalid)
                    throw new Exception("Range is an invalid state.");

                return lastAddress; 
            }
            set { lastAddress = value; }
        }

        public Boolean IsInvalid { get; set; }

        public XLRangeAddress(IXLAddress firstAddress, IXLAddress lastAddress)
        {
            if (firstAddress.Worksheet != lastAddress.Worksheet)
                throw new ArgumentException("First and last addresses must be in the same worksheet");

            Worksheet = firstAddress.Worksheet;
            FirstAddress = firstAddress;
            LastAddress = lastAddress;
        }

        public XLRangeAddress(IXLWorksheet worksheet, String rangeAddress)
        {
            String addressToUse;
            if (rangeAddress.Contains("!"))
                addressToUse = rangeAddress.Substring(rangeAddress.IndexOf("!") + 1);
            else
                addressToUse = rangeAddress;

            XLAddress firstAddress;
            XLAddress lastAddress;
            if (addressToUse.Contains(':'))
            {
                String[] arrRange = addressToUse.Split(':');
                var firstPart = arrRange[0];
                var secondPart = arrRange[1];
                firstAddress = new XLAddress(worksheet, firstPart);
                lastAddress = new XLAddress(worksheet, secondPart);
            }
            else
            {
                firstAddress = new XLAddress(worksheet, addressToUse);
                lastAddress = new XLAddress(worksheet, addressToUse);
            }
            FirstAddress = firstAddress;
            LastAddress = lastAddress;
            Worksheet = worksheet;
        }

        public override string ToString()
        {
            return firstAddress.ToString() + ":" + lastAddress.ToString();
        }

        public String ToStringRelative()
        {
            return firstAddress.ToStringRelative() + ":" + lastAddress.ToStringRelative();
        }
        public String ToStringFixed()
        {
            return firstAddress.ToStringFixed() + ":" + lastAddress.ToStringFixed();
        }

        public override bool Equals(object obj)
        {
            var other = (XLRangeAddress)obj;
            return 
                    this.Worksheet.Equals(other.Worksheet)
                && this.FirstAddress.Equals(other.FirstAddress)
                && this.LastAddress.Equals(other.LastAddress);
        }

        public override int GetHashCode()
        {
            return
                Worksheet.GetHashCode()
                ^ FirstAddress.GetHashCode()
                ^ LastAddress.GetHashCode();
        }
    }
}

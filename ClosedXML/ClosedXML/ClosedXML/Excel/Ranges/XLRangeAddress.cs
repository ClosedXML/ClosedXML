using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLRangeAddress: IXLRangeAddress
    {
        //public IXLWorksheet Worksheet { get; set; }

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

        public XLRangeAddress(String firstCellAddress, String lastCellAddress)
        {
            FirstAddress = new XLAddress(firstCellAddress);
            LastAddress = new XLAddress(lastCellAddress);
        }

        public XLRangeAddress(Int32 firstCellRow, Int32 firstCellColumn, Int32 lastCellRow, Int32 lastCellColumn)
        {
            FirstAddress = new XLAddress(firstCellRow, firstCellColumn);
            LastAddress = new XLAddress(lastCellRow, lastCellColumn);
        }

        public XLRangeAddress(IXLAddress firstAddress, IXLAddress lastAddress)
        {
            FirstAddress = firstAddress;
            LastAddress = lastAddress;
        }

        public XLRangeAddress(String rangeAddress)
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
                firstAddress = new XLAddress(firstPart);
                lastAddress = new XLAddress(secondPart);
            }
            else
            {
                firstAddress = new XLAddress(addressToUse);
                lastAddress = new XLAddress(addressToUse);
            }
            FirstAddress = firstAddress;
            LastAddress = lastAddress;
        }

        public override string ToString()
        {
            return firstAddress.ToString() + ":" + lastAddress.ToString();
        }

        public override bool Equals(object obj)
        {
            var other = (XLRangeAddress)obj;
            return this.FirstAddress.Equals(other.FirstAddress)
                && this.LastAddress.Equals(other.LastAddress);
        }

        public override int GetHashCode()
        {
            return FirstAddress.GetHashCode()
                    ^ LastAddress.GetHashCode();
        }
    }
}

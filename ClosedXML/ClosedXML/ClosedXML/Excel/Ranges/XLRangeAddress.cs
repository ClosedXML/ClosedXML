using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLRangeAddress: IXLRangeAddress
    {
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
                    throw new Exception("Range is invalid.");

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
            XLAddress firstAddress;
            XLAddress lastAddress;
            if (rangeAddress.Contains(':'))
            {
                String[] arrRange = rangeAddress.Split(':');
                var firstPart = arrRange[0];
                var secondPart = arrRange[1];
                firstAddress = new XLAddress(firstPart);
                lastAddress = new XLAddress(secondPart);
            }
            else
            {
                firstAddress = new XLAddress(rangeAddress);
                lastAddress = new XLAddress(rangeAddress);
            }
            FirstAddress = firstAddress;
            LastAddress = lastAddress;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal struct XLAddressLight
    {
        public XLAddressLight(Int32 rowNumber, Int32 columnNumber)
        {
            RowNumber = rowNumber;
            ColumnNumber = columnNumber;
        }
        public Int32 RowNumber;
        public Int32 ColumnNumber;
    }
}

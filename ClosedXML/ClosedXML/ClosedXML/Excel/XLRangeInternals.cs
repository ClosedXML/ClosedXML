using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public class XLRangeInternals: IXLRangeInternals
    {
        public XLRangeInternals(IXLAddress firstCellAddress, IXLAddress lastCellAddress, IXLWorksheet worksheet)
        {
            FirstCellAddress = firstCellAddress;
            LastCellAddress = lastCellAddress;
            Worksheet = worksheet;
        }
        public IXLAddress FirstCellAddress { get; private set; }
        public IXLAddress LastCellAddress { get; private set; }
        public IXLWorksheet Worksheet { get; private set; }
    }
}

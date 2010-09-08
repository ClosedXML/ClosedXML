using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLRangeInternals
    {
        IXLAddress FirstCellAddress { get; }
        IXLAddress LastCellAddress { get; }
        IXLWorksheet Worksheet { get; }
    }
}

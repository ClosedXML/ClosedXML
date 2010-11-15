using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLRangeAddress
    {
        //IXLWorksheet Worksheet { get; set; }
        IXLAddress FirstAddress { get; set; }
        IXLAddress LastAddress { get; set; }
        Boolean IsInvalid { get; set; }
    }
}

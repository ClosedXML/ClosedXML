using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    internal interface IXLStylized
    {
        IXLStyle Style { get; set; }
        IEnumerable<IXLStyle> Styles { get; }
        Boolean UpdatingStyle { get; set; }
        IXLStyle InnerStyle { get; set; }
        IXLRanges RangesUsed { get; }
        Boolean StyleChanged { get; set; }
        //Boolean IsDefault { get; set; }
    }
}

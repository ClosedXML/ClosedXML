using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal interface IXLStylized
    {
        IXLStyle Style { get; set; }
        IEnumerable<IXLStyle> Styles { get; }
        Boolean UpdatingStyle { get; set; }
        IXLStyle InnerStyle { get; set; }
        //Boolean IsDefault { get; set; }
    }
}

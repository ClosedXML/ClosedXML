using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLColumn: IXLRange
    {
        Double Width { get; set; }

    }
}

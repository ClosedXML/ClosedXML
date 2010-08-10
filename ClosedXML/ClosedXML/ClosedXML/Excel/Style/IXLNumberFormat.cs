using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel.Style
{
    public interface IXLNumberFormat
    {
        UInt32? NumberFormatId { get; set; }
        String Format { get; set; }
    }
}

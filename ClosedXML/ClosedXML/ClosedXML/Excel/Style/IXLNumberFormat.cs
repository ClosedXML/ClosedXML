using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLNumberFormat
    {
        Int32 NumberFormatId { get; set; }
        String Format { get; set; }
    }
}

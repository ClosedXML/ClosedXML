using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel.Style
{
    public class XLNumberFormat: IXLNumberFormat
    {
        public uint? NumberFormatId { get; set; }

        public string Format { get; set; }
    }
}

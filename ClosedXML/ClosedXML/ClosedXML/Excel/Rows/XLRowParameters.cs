using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace ClosedXML.Excel
{
    public class XLRowParameters
    {
        public XLRowParameters(IXLWorksheet worksheet, IXLStyle defaultStyle)
        {
            Worksheet = worksheet;
            DefaultStyle = defaultStyle;
        }
        public IXLStyle DefaultStyle { get; set; }
        public IXLWorksheet Worksheet { get; private set; }
    }
}

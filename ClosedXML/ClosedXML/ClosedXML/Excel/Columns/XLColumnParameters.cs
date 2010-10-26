using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace ClosedXML.Excel
{
    internal class XLColumnParameters
    {
        public XLColumnParameters(XLWorksheet worksheet, IXLStyle defaultStyle, Boolean isReference = true)
        {
            Worksheet = worksheet;
            DefaultStyle = defaultStyle;
            IsReference = isReference;
        }
        public IXLStyle DefaultStyle { get; set; }
        public XLWorksheet Worksheet { get; private set; }
        public Boolean IsReference { get; private set; }
    }
}

using System;


namespace ClosedXML.Excel
{
    internal class XLRowParameters
    {
        public XLRowParameters(XLWorksheet worksheet, IXLStyle defaultStyle)
        {
            Worksheet = worksheet;
            DefaultStyle = defaultStyle;
        }

        public IXLStyle DefaultStyle { get; private set; }
        public XLWorksheet Worksheet { get; private set; }
    }
}

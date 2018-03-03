using System;


namespace ClosedXML.Excel
{
    internal class XLRowParameters
    {
        public XLRowParameters(XLWorksheet worksheet, IXLStyle defaultStyle, Boolean isReference = true)
        {
            Worksheet = worksheet;
            DefaultStyle = defaultStyle;
            IsReference = isReference;
        }

        public IXLStyle DefaultStyle { get; private set; }
        public XLWorksheet Worksheet { get; private set; }
        public Boolean IsReference { get; private set; }
    }
}
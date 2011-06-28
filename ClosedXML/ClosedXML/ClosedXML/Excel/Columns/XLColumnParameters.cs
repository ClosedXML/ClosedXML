using System;


namespace ClosedXML.Excel
{
    internal class XLColumnParameters
    {
        public XLColumnParameters(XLWorksheet worksheet, IXLStyle defaultStyle, Boolean isReference)
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

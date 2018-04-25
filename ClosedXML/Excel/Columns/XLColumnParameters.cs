using System;


namespace ClosedXML.Excel
{
    internal class XLColumnParameters
    {
        public XLColumnParameters(XLWorksheet worksheet, IXLStyle defaultStyle)
        {
            Worksheet = worksheet;
            DefaultStyle = defaultStyle;
            //IsReference = isReference;
        }
        public IXLStyle DefaultStyle { get; private set; }
        public XLWorksheet Worksheet { get; private set; }
        //public Boolean IsReference { get; private set; }
    }
}

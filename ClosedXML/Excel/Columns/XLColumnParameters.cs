using System;


namespace ClosedXML.Excel
{
    internal class XLColumnParameters
    {
        public XLColumnParameters(XLWorksheet worksheet, Int32 defaultStyleId, Boolean isReference)
        {
            Worksheet = worksheet;
            DefaultStyleId = defaultStyleId;
            IsReference = isReference;
        }
        public Int32 DefaultStyleId { get; set; }
        public XLWorksheet Worksheet { get; private set; }
        public Boolean IsReference { get; private set; }
    }
}

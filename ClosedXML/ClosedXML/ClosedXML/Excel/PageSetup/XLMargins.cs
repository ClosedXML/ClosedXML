using System;

namespace ClosedXML.Excel
{
    internal class XLMargins: IXLMargins
    {
        public Double Left { get; set; }
        public Double Right { get; set; }
        public Double Top { get; set; }
        public Double Bottom { get; set; }
        public Double Header { get; set; }
        public Double Footer { get; set; }
    }
}

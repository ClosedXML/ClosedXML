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

        public IXLMargins SetLeft(Double value) { Left = value; return this; }
        public IXLMargins SetRight(Double value) { Right = value; return this; }
        public IXLMargins SetTop(Double value) { Top = value; return this; }
        public IXLMargins SetBottom(Double value) { Bottom = value; return this; }
        public IXLMargins SetHeader(Double value) { Header = value; return this; }
        public IXLMargins SetFooter(Double value) { Footer = value; return this; }

    }
}

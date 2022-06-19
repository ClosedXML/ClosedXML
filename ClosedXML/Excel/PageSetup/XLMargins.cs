namespace ClosedXML.Excel
{
    internal class XLMargins: IXLMargins
    {
        public double Left { get; set; }
        public double Right { get; set; }
        public double Top { get; set; }
        public double Bottom { get; set; }
        public double Header { get; set; }
        public double Footer { get; set; }

        public IXLMargins SetLeft(double value) { Left = value; return this; }
        public IXLMargins SetRight(double value) { Right = value; return this; }
        public IXLMargins SetTop(double value) { Top = value; return this; }
        public IXLMargins SetBottom(double value) { Bottom = value; return this; }
        public IXLMargins SetHeader(double value) { Header = value; return this; }
        public IXLMargins SetFooter(double value) { Footer = value; return this; }

    }
}

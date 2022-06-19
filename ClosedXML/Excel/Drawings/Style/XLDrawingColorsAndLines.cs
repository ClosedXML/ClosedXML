namespace ClosedXML.Excel
{
    internal class XLDrawingColorsAndLines: IXLDrawingColorsAndLines
    {
                private readonly IXLDrawingStyle _style;

        public XLDrawingColorsAndLines(IXLDrawingStyle style)
        {
            _style = style;
        }
        public XLColor FillColor { get; set; }		public IXLDrawingStyle SetFillColor(XLColor value) { FillColor = value; return _style; }
        public double FillTransparency { get; set; }		public IXLDrawingStyle SetFillTransparency(double value) { FillTransparency = value; return _style; }
        public XLColor LineColor { get; set; }		public IXLDrawingStyle SetLineColor(XLColor value) { LineColor = value; return _style; }
        public double LineTransparency { get; set; }		public IXLDrawingStyle SetLineTransparency(double value) { LineTransparency = value; return _style; }
        public double LineWeight { get; set; }		public IXLDrawingStyle SetLineWeight(double value) { LineWeight = value; return _style; }
        public XLDashStyle LineDash { get; set; }		public IXLDrawingStyle SetLineDash(XLDashStyle value) { LineDash = value; return _style; }
        public XLLineStyle LineStyle { get; set; }		public IXLDrawingStyle SetLineStyle(XLLineStyle value) { LineStyle = value; return _style; }

    }
}

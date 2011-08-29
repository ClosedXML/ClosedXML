using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLDrawingColorsAndLines: IXLDrawingColorsAndLines
    {
                private readonly IXLDrawingStyle _style;

        public XLDrawingColorsAndLines(IXLDrawingStyle style)
        {
            _style = style;
            FillColor = XLColor.FromArgb(255, 255, 225);
            LineColor = XLColor.Black;
            LineDash = XLDashTypes.Solid;
            LineStyle = XLLineStyles.OneQuarter;
            LineWeight = 0.75;
        }
        public IXLColor FillColor { get; set; }		public IXLDrawingStyle SetFillColor(XLColor value) { FillColor = value; return _style; }
        public Int32 FillTransparency { get; set; }		public IXLDrawingStyle SetFillTransparency(Int32 value) { FillTransparency = value; return _style; }
        public IXLColor LineColor { get; set; }		public IXLDrawingStyle SetLineColor(XLColor value) { LineColor = value; return _style; }
        public Double LineWeight { get; set; }		public IXLDrawingStyle SetLineWeight(Double value) { LineWeight = value; return _style; }
        public XLDashTypes LineDash { get; set; }		public IXLDrawingStyle SetLineDash(XLDashTypes value) { LineDash = value; return _style; }
        public XLLineStyles LineStyle { get; set; }		public IXLDrawingStyle SetLineStyle(XLLineStyles value) { LineStyle = value; return _style; }

    }
}

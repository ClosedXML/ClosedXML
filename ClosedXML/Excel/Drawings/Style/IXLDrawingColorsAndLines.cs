using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public enum XLDashStyle
    {
        Solid,
        RoundDot,
        SquareDot,
        Dash,
        DashDot,
        LongDash,
        LongDashDot,
        LongDashDotDot
    }
    public enum XLLineStyle
    {
        Single, ThinThin, ThinThick, ThickThin, ThickBetweenThin
    }
    public interface IXLDrawingColorsAndLines
    {
        XLColor FillColor { get; set; }
        Double FillTransparency { get; set; }
        XLColor LineColor { get; set; }
        Double LineTransparency { get; set; }
        Double LineWeight { get; set; }
        XLDashStyle LineDash { get; set; }
        XLLineStyle LineStyle { get; set; }

        IXLDrawingStyle SetFillColor(XLColor value);
        IXLDrawingStyle SetFillTransparency(Double value);
        IXLDrawingStyle SetLineColor(XLColor value);
        IXLDrawingStyle SetLineTransparency(Double value);
        IXLDrawingStyle SetLineWeight(Double value);
        IXLDrawingStyle SetLineDash(XLDashStyle value);
        IXLDrawingStyle SetLineStyle(XLLineStyle value);

    }
}

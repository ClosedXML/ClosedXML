using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public enum XLDashTypes
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
    public enum XLLineStyles
    {
        OneQuarter,
        OneHalf,
        ThreeQuarters,
        One,
        OneAndOneHalf,
        TwoAndOneQuarter,
        Three,
        FourAndOneHalf,
        Six,
        ThreeSplit,
        FourAndOneHalfSplit1,
        FourAndOneHalfSplit2,
        SixSplit
    }
    public interface IXLDrawingColorsAndLines
    {
        IXLColor FillColor { get; set; }
        Int32 FillTransparency { get; set; }
        IXLColor LineColor { get; set; }
        Double LineWeight { get; set; }
        XLDashTypes LineDash { get; set; }
        XLLineStyles LineStyle { get; set; }

        IXLDrawingStyle SetFillColor(XLColor value);
        IXLDrawingStyle SetFillTransparency(Int32 value);
        IXLDrawingStyle SetLineColor(XLColor value);
        IXLDrawingStyle SetLineWeight(Double value);
        IXLDrawingStyle SetLineDash(XLDashTypes value);
        IXLDrawingStyle SetLineStyle(XLLineStyles value);

    }
}

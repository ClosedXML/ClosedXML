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
        double FillTransparency { get; set; }
        XLColor LineColor { get; set; }
        double LineTransparency { get; set; }
        double LineWeight { get; set; }
        XLDashStyle LineDash { get; set; }
        XLLineStyle LineStyle { get; set; }

        IXLDrawingStyle SetFillColor(XLColor value);

        IXLDrawingStyle SetFillTransparency(double value);

        IXLDrawingStyle SetLineColor(XLColor value);

        IXLDrawingStyle SetLineTransparency(double value);

        IXLDrawingStyle SetLineWeight(double value);

        IXLDrawingStyle SetLineDash(XLDashStyle value);

        IXLDrawingStyle SetLineStyle(XLLineStyle value);
    }
}
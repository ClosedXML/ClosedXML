using System;

namespace ClosedXML.Excel
{
    public enum XLBorderStyleValues
    {
        DashDot,
        DashDotDot,
        Dashed,
        Dotted,
        Double,
        Hair,
        Medium,
        MediumDashDot,
        MediumDashDotDot,
        MediumDashed,
        None,
        SlantDashDot,
        Thick,
        Thin
    }

    public interface IXLBorder : IEquatable<IXLBorder>
    {
        XLBorderStyleValues OutsideBorder { set; }

        XLColor OutsideBorderColor { set; }

        XLBorderStyleValues InsideBorder { set; }

        XLColor InsideBorderColor { set; }

        XLBorderStyleValues LeftBorder { get; set; }

        XLColor LeftBorderColor { get; set; }

        XLBorderStyleValues RightBorder { get; set; }

        XLColor RightBorderColor { get; set; }

        XLBorderStyleValues TopBorder { get; set; }

        XLColor TopBorderColor { get; set; }

        XLBorderStyleValues BottomBorder { get; set; }

        XLColor BottomBorderColor { get; set; }

        Boolean DiagonalUp { get; set; }

        Boolean DiagonalDown { get; set; }

        XLBorderStyleValues DiagonalBorder { get; set; }

        XLColor DiagonalBorderColor { get; set; }

        IXLStyle SetOutsideBorder(XLBorderStyleValues value);

        IXLStyle SetOutsideBorderColor(XLColor value);

        IXLStyle SetInsideBorder(XLBorderStyleValues value);

        IXLStyle SetInsideBorderColor(XLColor value);

        IXLStyle SetLeftBorder(XLBorderStyleValues value);

        IXLStyle SetLeftBorderColor(XLColor value);

        IXLStyle SetRightBorder(XLBorderStyleValues value);

        IXLStyle SetRightBorderColor(XLColor value);

        IXLStyle SetTopBorder(XLBorderStyleValues value);

        IXLStyle SetTopBorderColor(XLColor value);

        IXLStyle SetBottomBorder(XLBorderStyleValues value);

        IXLStyle SetBottomBorderColor(XLColor value);

        IXLStyle SetDiagonalUp(); IXLStyle SetDiagonalUp(Boolean value);

        IXLStyle SetDiagonalDown(); IXLStyle SetDiagonalDown(Boolean value);

        IXLStyle SetDiagonalBorder(XLBorderStyleValues value);

        IXLStyle SetDiagonalBorderColor(XLColor value);
    }
}

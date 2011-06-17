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
    public interface IXLBorder: IEquatable<IXLBorder>
    {
        XLBorderStyleValues OutsideBorder { set; }
        IXLColor OutsideBorderColor { set; }

        XLBorderStyleValues LeftBorder { get; set; }
        IXLColor LeftBorderColor { get; set; }
        XLBorderStyleValues RightBorder { get; set; }
        IXLColor RightBorderColor { get; set; }
        XLBorderStyleValues TopBorder { get; set; }
        IXLColor TopBorderColor { get; set; }
        XLBorderStyleValues BottomBorder { get; set; }
        IXLColor BottomBorderColor { get; set; }
        Boolean DiagonalUp { get; set; }
        Boolean DiagonalDown { get; set; }
        XLBorderStyleValues DiagonalBorder { get; set; }
        IXLColor DiagonalBorderColor { get; set; }

        IXLStyle SetOutsideBorder(XLBorderStyleValues value);
        IXLStyle SetOutsideBorderColor(IXLColor value);

        IXLStyle SetLeftBorder(XLBorderStyleValues value);
        IXLStyle SetLeftBorderColor(IXLColor value);
        IXLStyle SetRightBorder(XLBorderStyleValues value);
        IXLStyle SetRightBorderColor(IXLColor value);
        IXLStyle SetTopBorder(XLBorderStyleValues value);
        IXLStyle SetTopBorderColor(IXLColor value);
        IXLStyle SetBottomBorder(XLBorderStyleValues value);
        IXLStyle SetBottomBorderColor(IXLColor value);
        IXLStyle SetDiagonalUp(); IXLStyle SetDiagonalUp(Boolean value);
        IXLStyle SetDiagonalDown(); IXLStyle SetDiagonalDown(Boolean value);
        IXLStyle SetDiagonalBorder(XLBorderStyleValues value);
        IXLStyle SetDiagonalBorderColor(IXLColor value);

    }
}

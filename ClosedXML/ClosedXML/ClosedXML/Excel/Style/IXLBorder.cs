using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;

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
    }
}

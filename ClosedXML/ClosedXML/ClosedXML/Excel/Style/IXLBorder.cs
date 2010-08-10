using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;

namespace ClosedXML.Excel.Style
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
    public interface IXLBorder
    {
        XLBorderStyleValues LeftBorder { get; set; }

        Color LeftBorderColor { get; set; }

        XLBorderStyleValues RightBorder { get; set; }

        Color RightBorderColor { get; set; }

        XLBorderStyleValues TopBorder { get; set; }

        Color TopBorderColor { get; set; }

        XLBorderStyleValues BottomBorder { get; set; }

        Color BottomBorderColor { get; set; }

        Boolean DiagonalUp { get; set; }

        Boolean DiagonalDown { get; set; }

        XLBorderStyleValues DiagonalBorder { get; set; }

        Color DiagonalBorderColor { get; set; }
    }
}

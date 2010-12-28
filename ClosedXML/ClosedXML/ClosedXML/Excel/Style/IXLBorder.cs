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
    }
}

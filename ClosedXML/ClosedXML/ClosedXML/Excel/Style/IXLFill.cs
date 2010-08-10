using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;

namespace ClosedXML.Excel.Style
{
    public enum XLFillPatternValues
    {
        DarkDown,
        DarkGray,
        DarkGrid,
        DarkHorizontal,
        DarkTrellis,
        DarkUp,
        DarkVertical,
        Gray0625,
        Gray125,
        LightDown,
        LightGray,
        LightGrid,
        LightHorizontal,
        LightTrellis,
        LightUp,
        LightVertical,
        MediumGray,
        None,
        Solid
    }

    public interface IXLFill
    {
        Color BackgroundColor { get; set; }

        Color PatternColor { get; set; }

        Color PatternBackgroundColor { get; set; }

        XLFillPatternValues PatternType { get; set; }
    }
}

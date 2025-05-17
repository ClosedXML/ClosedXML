using System;

namespace ClosedXML.Excel
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

        /// <summary>
        /// For solid fill, the fill color is taken from the <see cref="IXLFill.PatternColor"/>.
        /// The <see cref="IXLFill.BackgroundColor"/> is not used.
        /// </summary>
        Solid
    }

    public interface IXLFill : IEquatable<IXLFill>
    {
        XLColor BackgroundColor { get; set; }

        XLColor PatternColor { get; set; }

        XLFillPatternValues PatternType { get; set; }

        IXLStyle SetBackgroundColor(XLColor value);

        IXLStyle SetPatternColor(XLColor value);

        IXLStyle SetPatternType(XLFillPatternValues value);
    }
}

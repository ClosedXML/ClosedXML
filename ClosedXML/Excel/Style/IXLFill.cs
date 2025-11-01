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
        /// <summary>
        /// Gets or set the background color of a fill.
        /// </summary>
        /// <remarks>
        /// Setter will automatically change the fill to an empty pattern fill if needed (e.g.,
        /// from a gradient fill). It then sets the background color and updates the pattern type
        /// if necessary:
        /// <list type="bullet">
        ///   <item>If the set color is transparent, and the fill type is <see cref="XLFillPatternValues.Solid"/>
        ///     the fill type becomes <see cref="XLFillPatternValues.None"/>.</item>
        ///   <item>If the set color is not transparent and the fill type is <see cref="XLFillPatternValues.None"/>,
        ///     the fill type is changed to <see cref="XLFillPatternValues.Solid"/>.</item>
        /// </list>
        /// </remarks>
        XLColor BackgroundColor { get; set; }

        /// <summary>
        /// Gets or set the pattern color of a fill.
        /// </summary>
        XLColor PatternColor { get; set; }

        /// <summary>
        /// Change pattern type of a fill.
        /// </summary>
        /// <remarks>
        /// If the pattern type is changed from <see cref="XLFillPatternValues.None"/> to
        /// a different pattern type, the <see cref="BackgroundColor"/> is set to
        /// a <see cref="XLThemeColor.Text1"/> color.
        /// </remarks>
        XLFillPatternValues PatternType { get; set; }

        /// <summary>
        /// Set the background color of a fill.
        /// </summary>
        /// <inheritdoc cref="BackgroundColor"/>
        IXLStyle SetBackgroundColor(XLColor value);

        /// <summary>
        /// Set the pattern color of a fill.
        /// </summary>
        /// <inheritdoc cref="PatternColor"/>
        IXLStyle SetPatternColor(XLColor value);

        /// <summary>
        /// Set the pattern type of a fill.
        /// </summary>
        /// <inheritdoc cref="PatternType"/>
        IXLStyle SetPatternType(XLFillPatternValues value);
    }
}

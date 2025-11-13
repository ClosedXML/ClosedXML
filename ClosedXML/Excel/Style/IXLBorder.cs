#nullable disable

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

    /// <summary>
    /// <para>
    /// The interface is used across many different objects. The value returned by properties is
    /// defined only for <see cref="IXLCell"/>, <see cref="IXLRow"/>, <see cref="IXLColumn"/>,
    /// <see cref="IXLWorksheet"/> and <see cref="XLWorkbook"/>. The returned value is undefined
    /// for other <see cref="IXLRangeBase"/> objects that can contain multiple different property
    /// values (e.g. <see cref="IXLRange"/> can contain multiple cells, <see cref="IXLColumns"/>
    /// can contains multiple columns).
    /// </para>
    /// </summary>
    public interface IXLBorder : IEquatable<IXLBorder>
    {
        XLBorderStyleValues OutsideBorder { set; }

        XLColor OutsideBorderColor { set; }

        XLBorderStyleValues InsideBorder { set; }

        XLColor InsideBorderColor { set; }

        /// <summary>
        /// Get or set style of the left border.
        /// </summary>
        /// <remarks>
        /// When style is set to <see cref="XLBorderStyleValues.None"/>, the <see cref="LeftBorderColor"/>
        /// is set to the default border color.
        /// </remarks>
        XLBorderStyleValues LeftBorder { get; set; }

        /// <summary>
        /// Get or set color of the left border.
        /// </summary>
        /// <remarks>
        /// The color can be set only when the border is visible (=<see cref="LeftBorder"/>
        /// is not <see cref="XLBorderStyleValues.None"/>). Set style first, then color.
        /// </remarks>
        XLColor LeftBorderColor { get; set; }

        /// <summary>
        /// Get or set style of the right border.
        /// </summary>
        /// <remarks>
        /// When style is set to <see cref="XLBorderStyleValues.None"/>, the <see cref="RightBorderColor"/>
        /// is set to the default border color.
        /// </remarks>
        XLBorderStyleValues RightBorder { get; set; }

        /// <summary>
        /// Get or set color of the right border.
        /// </summary>
        /// <remarks>
        /// The color can be set only when the border is visible (=<see cref="RightBorder"/>
        /// is not <see cref="XLBorderStyleValues.None"/>). Set style first, then color.
        /// </remarks>
        XLColor RightBorderColor { get; set; }

        /// <summary>
        /// Get or set style of the top border.
        /// </summary>
        /// <remarks>
        /// When style is set to <see cref="XLBorderStyleValues.None"/>, the <see cref="TopBorderColor"/>
        /// is set to the default border color.
        /// </remarks>
        XLBorderStyleValues TopBorder { get; set; }

        /// <summary>
        /// Get or set color of the top border.
        /// </summary>
        /// <remarks>
        /// The color can be set only when the border is visible (=<see cref="TopBorder"/>
        /// is not <see cref="XLBorderStyleValues.None"/>). Set style first, then color.
        /// </remarks>
        XLColor TopBorderColor { get; set; }

        /// <summary>
        /// Get or set style of the bottom border.
        /// </summary>
        /// <remarks>
        /// When style is set to <see cref="XLBorderStyleValues.None"/>, the <see cref="BottomBorderColor"/>
        /// is set to the default border color.
        /// </remarks>
        XLBorderStyleValues BottomBorder { get; set; }

        /// <summary>
        /// Get or set color of the bottom border.
        /// </summary>
        /// <remarks>
        /// The color can be set only when the border is visible (=<see cref="BottomBorder"/>
        /// is not <see cref="XLBorderStyleValues.None"/>). Set style first, then color.
        /// </remarks>
        XLColor BottomBorderColor { get; set; }

        Boolean DiagonalUp { get; set; }

        Boolean DiagonalDown { get; set; }

        /// <summary>
        /// Get or set style of the diagonal border.
        /// </summary>
        /// <remarks>
        /// When style is set to <see cref="XLBorderStyleValues.None"/>, the <see cref="DiagonalBorderColor"/>
        /// is set to the default border color.
        /// </remarks>
        XLBorderStyleValues DiagonalBorder { get; set; }

        /// <summary>
        /// Get or set color of the diagonal border.
        /// </summary>
        /// <remarks>
        /// The color can be set only when the border line is can be visible
        /// (=<see cref="DiagonalBorder"/> is not <see cref="XLBorderStyleValues.None"/>).
        /// Set style first, then color.
        /// </remarks>
        XLColor DiagonalBorderColor { get; set; }

        IXLStyle SetOutsideBorder(XLBorderStyleValues value);

        IXLStyle SetOutsideBorderColor(XLColor value);

        IXLStyle SetInsideBorder(XLBorderStyleValues value);

        IXLStyle SetInsideBorderColor(XLColor value);

        /// <summary>
        /// Set style of the left border.
        /// </summary>
        /// <inheritdoc cref="LeftBorder"/>
        IXLStyle SetLeftBorder(XLBorderStyleValues value);

        /// <summary>
        /// Set color of the left border.
        /// </summary>
        /// <inheritdoc cref="LeftBorderColor"/>
        IXLStyle SetLeftBorderColor(XLColor value);

        /// <summary>
        /// Set style of the right border.
        /// </summary>
        /// <inheritdoc cref="RightBorder"/>
        IXLStyle SetRightBorder(XLBorderStyleValues value);

        /// <summary>
        /// Set color of the right border.
        /// </summary>
        /// <inheritdoc cref="RightBorderColor"/>
        IXLStyle SetRightBorderColor(XLColor value);

        /// <summary>
        /// Set style of the top border.
        /// </summary>
        /// <inheritdoc cref="TopBorder"/>
        IXLStyle SetTopBorder(XLBorderStyleValues value);

        /// <summary>
        /// Set color of the top border.
        /// </summary>
        /// <inheritdoc cref="TopBorderColor"/>
        IXLStyle SetTopBorderColor(XLColor value);

        /// <summary>
        /// Set style of the bottom border.
        /// </summary>
        /// <inheritdoc cref="BottomBorder"/>
        IXLStyle SetBottomBorder(XLBorderStyleValues value);

        /// <summary>
        /// Set color of the bottom border.
        /// </summary>
        /// <inheritdoc cref="BottomBorderColor"/>
        IXLStyle SetBottomBorderColor(XLColor value);

        IXLStyle SetDiagonalUp(); IXLStyle SetDiagonalUp(Boolean value);

        IXLStyle SetDiagonalDown(); IXLStyle SetDiagonalDown(Boolean value);

        /// <summary>
        /// Set style of the diagonal border.
        /// </summary>
        /// <inheritdoc cref="DiagonalBorder"/>
        IXLStyle SetDiagonalBorder(XLBorderStyleValues value);

        /// <summary>
        /// Set color of the diagonal border.
        /// </summary>
        /// <inheritdoc cref="DiagonalBorderColor"/>
        IXLStyle SetDiagonalBorderColor(XLColor value);
    }
}

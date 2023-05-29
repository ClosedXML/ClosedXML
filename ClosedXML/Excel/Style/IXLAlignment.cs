#nullable disable

using System;

namespace ClosedXML.Excel
{
    public enum XLAlignmentReadingOrderValues : byte
    {
        ContextDependent,
        LeftToRight,
        RightToLeft
    }

    public enum XLAlignmentHorizontalValues : byte
    {
        Center,
        CenterContinuous,
        Distributed,
        Fill,
        General,
        Justify,
        Left,
        Right
    }

    public enum XLAlignmentVerticalValues : byte
    {
        Bottom,
        Center,
        Distributed,
        Justify,
        Top
    }

    public interface IXLAlignment : IEquatable<IXLAlignment>
    {
        /// <summary>
        /// Gets or sets the cell's horizontal alignment.
        /// </summary>
        XLAlignmentHorizontalValues Horizontal { get; set; }

        /// <summary>
        /// Gets or sets the cell's vertical alignment.
        /// </summary>
        XLAlignmentVerticalValues Vertical { get; set; }

        /// <summary>
        /// Gets or sets the cell's text indentation.
        /// </summary>
        Int32 Indent { get; set; }

        /// <summary>
        /// Gets or sets whether the cell's last line is justified or not.
        /// </summary>
        Boolean JustifyLastLine { get; set; }

        /// <summary>
        /// Gets or sets the cell's reading order.
        /// </summary>
        XLAlignmentReadingOrderValues ReadingOrder { get; set; }

        /// <summary>
        /// Gets or sets the cell's relative indent.
        /// </summary>
        Int32 RelativeIndent { get; set; }

        /// <summary>
        /// Gets or sets whether the cell's font size should decrease to fit the contents.
        /// </summary>
        Boolean ShrinkToFit { get; set; }

        /// <summary>
        /// Gets or sets the cell's text rotation in degrees. Allowed values are -90
        /// (text is rotated clockwise) to 90 (text is rotated counterclockwise) and
        /// 255 for vertical layout of a text.
        /// </summary>
        Int32 TextRotation { get; set; }

        /// <summary>
        /// Gets or sets whether the cell's text should wrap if it doesn't fit.
        /// </summary>
        Boolean WrapText { get; set; }

        /// <summary>
        /// Gets or sets whether the cell's text should be displayed from to to bottom
        /// <para>(as opposed to the normal left to right).</para>
        /// </summary>
        Boolean TopToBottom { get; set; }

        IXLStyle SetHorizontal(XLAlignmentHorizontalValues value);

        IXLStyle SetVertical(XLAlignmentVerticalValues value);

        IXLStyle SetIndent(Int32 value);

        IXLStyle SetJustifyLastLine(); IXLStyle SetJustifyLastLine(Boolean value);

        IXLStyle SetReadingOrder(XLAlignmentReadingOrderValues value);

        IXLStyle SetRelativeIndent(Int32 value);

        IXLStyle SetShrinkToFit(); IXLStyle SetShrinkToFit(Boolean value);

        IXLStyle SetTextRotation(Int32 value);

        IXLStyle SetWrapText(); IXLStyle SetWrapText(Boolean value);

        IXLStyle SetTopToBottom(); IXLStyle SetTopToBottom(Boolean value);
    }
}

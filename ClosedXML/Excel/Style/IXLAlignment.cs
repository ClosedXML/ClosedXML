using System;

namespace ClosedXML.Excel
{
    public enum XLAlignmentReadingOrderValues
    {
        ContextDependent,
        LeftToRight,
        RightToLeft
    }

    public enum XLAlignmentHorizontalValues
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

    public enum XLAlignmentVerticalValues
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
        int Indent { get; set; }

        /// <summary>
        /// Gets or sets whether the cell's last line is justified or not.
        /// </summary>
        bool JustifyLastLine { get; set; }

        /// <summary>
        /// Gets or sets the cell's reading order.
        /// </summary>
        XLAlignmentReadingOrderValues ReadingOrder { get; set; }

        /// <summary>
        /// Gets or sets the cell's relative indent.
        /// </summary>
        int RelativeIndent { get; set; }

        /// <summary>
        /// Gets or sets whether the cell's font size should decrease to fit the contents.
        /// </summary>
        bool ShrinkToFit { get; set; }

        /// <summary>
        /// Gets or sets the cell's text rotation.
        /// </summary>
        int TextRotation { get; set; }

        /// <summary>
        /// Gets or sets whether the cell's text should wrap if it doesn't fit.
        /// </summary>
        bool WrapText { get; set; }

        /// <summary>
        /// Gets or sets wheter the cell's text should be displayed from to to bottom
        /// <para>(as opposed to the normal left to right).</para>
        /// </summary>
        bool TopToBottom { get; set; }

        IXLStyle SetHorizontal(XLAlignmentHorizontalValues value);

        IXLStyle SetVertical(XLAlignmentVerticalValues value);

        IXLStyle SetIndent(int value);

        IXLStyle SetJustifyLastLine(); IXLStyle SetJustifyLastLine(bool value);

        IXLStyle SetReadingOrder(XLAlignmentReadingOrderValues value);

        IXLStyle SetRelativeIndent(int value);

        IXLStyle SetShrinkToFit(); IXLStyle SetShrinkToFit(bool value);

        IXLStyle SetTextRotation(int value);

        IXLStyle SetWrapText(); IXLStyle SetWrapText(bool value);

        IXLStyle SetTopToBottom(); IXLStyle SetTopToBottom(bool value);
    }
}

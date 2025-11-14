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

        /// <summary>
        /// <para>
        /// If <see cref="IXLAlignment.JustifyLastLine"/> is <c>false</c>, each line is justified
        /// (flushed left and right), regardless of whether line has explicit line break or not.
        /// This is a difference from <see cref="Justify"/> that only justifies automatically
        /// wrapped lines and doesn't justify last line.
        /// </para>
        /// <para>
        /// If <see cref="IXLAlignment.JustifyLastLine"/> is <c>true</c>, the words of each line
        /// are distributed along the line, so the gap on the left and right side of each word has
        /// same width. The gap is also on between word and left and right side of a cell.
        /// </para>
        /// </summary>
        /// <remarks>
        /// This alignment implicitly turns on <see cref="IXLAlignment.WrapText"/>, regardless of
        /// actual value.
        /// </remarks>
        Distributed,
        Fill,
        General,

        /// <summary>
        /// The text is justified when it is flushed left and right. For each line of text, it
        /// aligns each line of the wrapped text in a cell to the right and left (except the last
        /// line). Only text in wrapped lines (except last one) is justified. Text with explicit
        /// line breaks isn't.
        /// </summary>
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
        /// <value>
        /// The value means a width equal to <c>indent_value * 3 * width_of_space_in_normal_font</c>.
        /// </value>
        /// <exception cref="ArgumentOutOfRangeException">When the value is outside of [0,255].</exception>
        Int32 Indent { get; set; }

        /// <summary>
        /// Gets or sets whether the cell's last line is justified or not.
        /// </summary>
        /// <value>
        /// The value changes the behavior of <see cref="XLAlignmentHorizontalValues.Distributed"/>
        /// alignment. Name of the property doesn't match actual behavior.
        /// </value>
        Boolean JustifyLastLine { get; set; }

        /// <summary>
        /// Gets or sets the cell's reading order.
        /// </summary>
        XLAlignmentReadingOrderValues ReadingOrder { get; set; }

        /// <summary>
        /// Gets or sets the cell's relative indent.
        /// </summary>
        /// <remarks>
        /// This property doesn't seem to work. It only set indent of a cell to 0 when used as part
        /// of differential format (regardless of relative ident value or actual ident value).
        /// </remarks>
        /// <value>
        /// The value is used only in differential formatting. It determines an additional indent
        /// to add to cells exiting <see cref="Indent"/>.
        /// </value>
        Int32 RelativeIndent { get; set; }

        /// <summary>
        /// Gets or sets whether the cell's font size should decrease to fit the contents.
        /// </summary>
        /// <remarks>
        /// Only applicable when <see cref="WrapText"/> is <c>false</c>, either directly or
        /// implicitly through some <see cref="XLAlignmentHorizontalValues"/> alignment (e.g.
        /// <see cref="XLAlignmentHorizontalValues.Distributed"/>).
        /// </remarks>
        Boolean ShrinkToFit { get; set; }

        /// <summary>
        /// Gets or sets the cell's text rotation in degrees. 
        /// </summary>
        /// <value>
        /// Allowed values are -90 (text is rotated clockwise) to 90 (text is rotated
        /// counterclockwise) and 255 for vertical layout of a text.
        /// </value>
        /// <exception cref="ArgumentOutOfRangeException">When setting the rotation to a value outside of [-90,90] or 255.</exception>
        Int32 TextRotation { get; set; }

        /// <summary>
        /// Gets or sets whether the cell's text should wrap if it doesn't fit.
        /// </summary>
        Boolean WrapText { get; set; }

        /// <summary>
        /// Gets or sets whether the cell's text should be displayed from top to bottom
        /// (as opposed to the normal left to right).
        /// </summary>
        /// <remarks>
        /// The setter has same effect as as <c>alignment.TextRotation = topToBottom ? 255 : 0</c>.
        /// </remarks>
        Boolean TopToBottom { get; set; }

        /// <summary>
        /// Sets the cell's horizontal alignment.
        /// </summary>
        /// <inheritdoc cref="Horizontal"/>
        IXLStyle SetHorizontal(XLAlignmentHorizontalValues value);

        /// <summary>
        /// Sets the cell's vertical alignment.
        /// </summary>
        /// <inheritdoc cref="Vertical"/>
        IXLStyle SetVertical(XLAlignmentVerticalValues value);

        /// <summary>
        /// Sets the cell's text indentation.
        /// </summary>
        /// <inheritdoc cref="Indent"/>
        IXLStyle SetIndent(Int32 value);

        /// <summary>
        /// Changes mode of <see cref="XLAlignmentHorizontalValues.Distributed"/> alignment.
        /// </summary>
        /// <inheritdoc cref="JustifyLastLine"/>
        IXLStyle SetJustifyLastLine();

        /// <summary>
        /// Changes mode of <see cref="XLAlignmentHorizontalValues.Distributed"/> alignment.
        /// </summary>
        /// <inheritdoc cref="JustifyLastLine"/>
        IXLStyle SetJustifyLastLine(Boolean value);

        /// <summary>
        /// Sets the cell's reading order.
        /// </summary>
        /// <inheritdoc cref="ReadingOrder"/>
        IXLStyle SetReadingOrder(XLAlignmentReadingOrderValues value);

        IXLStyle SetRelativeIndent(Int32 value);

        /// <summary>
        /// Sets whether the cell's font size should decrease to fit the contents.
        /// </summary>
        /// <inheritdoc cref="ShrinkToFit"/>
        IXLStyle SetShrinkToFit();

        /// <summary>
        /// Sets whether the cell's font size should decrease to fit the contents.
        /// </summary>
        /// <inheritdoc cref="ShrinkToFit"/>
        IXLStyle SetShrinkToFit(Boolean value);

        /// <summary>
        /// Sets the cell's text rotation in degrees. 
        /// </summary>
        /// <inheritdoc cref="TextRotation"/>
        IXLStyle SetTextRotation(Int32 value);

        /// <summary>
        /// Sets whether the cell's text should wrap if it doesn't fit.
        /// </summary>
        /// <inheritdoc cref="WrapText"/>
        IXLStyle SetWrapText();

        /// <summary>
        /// Sets whether the cell's text should wrap if it doesn't fit.
        /// </summary>
        /// <inheritdoc cref="WrapText"/>
        IXLStyle SetWrapText(Boolean value);

        /// <summary>
        /// Sets whether the cell's text should be displayed from top to bottom (as opposed to
        /// the normal left to right).
        /// </summary>
        /// <inheritdoc cref="TopToBottom"/>
        IXLStyle SetTopToBottom();

        /// <summary>
        /// Sets whether the cell's text should be displayed from top to bottom (as opposed to
        /// the normal left to right).
        /// </summary>
        /// <inheritdoc cref="TopToBottom"/>
        IXLStyle SetTopToBottom(Boolean value);
    }
}

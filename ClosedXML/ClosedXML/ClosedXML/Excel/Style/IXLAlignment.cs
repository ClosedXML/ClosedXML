using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

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

    public interface IXLAlignment: IEquatable<IXLAlignment>
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
        /// Gets or sets the cell's text rotation.
        /// </summary>
        Int32 TextRotation { get; set; }
        /// <summary>
        /// Gets or sets whether the cell's text should wrap if it doesn't fit.
        /// </summary>
        Boolean WrapText { get; set; }
        /// <summary>
        /// Gets or sets wheter the cell's text should be displayed from to to bottom
        /// <para>(as opposed to the normal left to right).</para>
        /// </summary>
        Boolean TopToBottom { get; set; }
    }
}

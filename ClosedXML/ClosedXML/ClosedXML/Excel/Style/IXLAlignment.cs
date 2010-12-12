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
        XLAlignmentHorizontalValues Horizontal { get; set; }

        XLAlignmentVerticalValues Vertical { get; set; }

        Int32 Indent { get; set; }

        Boolean JustifyLastLine { get; set; }

        XLAlignmentReadingOrderValues ReadingOrder { get; set; }

        Int32 RelativeIndent { get; set; }

        Boolean ShrinkToFit { get; set; }

        Int32 TextRotation { get; set; }

        Boolean WrapText { get; set; }

        Boolean TopToBottom { get; set; }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public enum XLDrawingTextDirection
    {
        Context,
        LeftToRight,
        RightToLeft
    }
    public enum XLDrawingTextOrientation
    {
        LeftToRight,
        Vertical,
        BottomToTop,
        TopToBottom
    }
    public interface IXLDrawingAlignment
    {
        XLAlignmentHorizontalValues Horizontal { get; set; }
        XLAlignmentVerticalValues Vertical { get; set; }
        Boolean AutomaticSize { get; set; }
        XLDrawingTextDirection Direction { get; set; }
        XLDrawingTextOrientation Orientation { get; set; }

        IXLDrawingStyle SetHorizontal(XLAlignmentHorizontalValues value);
        IXLDrawingStyle SetVertical(XLAlignmentVerticalValues value);
        IXLDrawingStyle SetAutomaticSize(); IXLDrawingStyle SetAutomaticSize(Boolean value);
        IXLDrawingStyle SetDirection(XLDrawingTextDirection value);
        IXLDrawingStyle SetOrientation(XLDrawingTextOrientation value);

    }
}

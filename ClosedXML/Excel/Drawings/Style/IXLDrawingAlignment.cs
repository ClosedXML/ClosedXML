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
    public enum XLDrawingHorizontalAlignment { Left, Justify, Center, Right, Distributed }
    public enum XLDrawingVerticalAlignment { Top, Justify, Center, Bottom, Distributed }
    public interface IXLDrawingAlignment
    {
        XLDrawingHorizontalAlignment Horizontal { get; set; }
        XLDrawingVerticalAlignment Vertical { get; set; }
        Boolean AutomaticSize { get; set; }
        XLDrawingTextDirection Direction { get; set; }
        XLDrawingTextOrientation Orientation { get; set; }

        IXLDrawingStyle SetHorizontal(XLDrawingHorizontalAlignment value);
        IXLDrawingStyle SetVertical(XLDrawingVerticalAlignment value);
        IXLDrawingStyle SetAutomaticSize(); IXLDrawingStyle SetAutomaticSize(Boolean value);
        IXLDrawingStyle SetDirection(XLDrawingTextDirection value);
        IXLDrawingStyle SetOrientation(XLDrawingTextOrientation value);

    }
}

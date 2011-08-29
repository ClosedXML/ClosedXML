using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLDrawingAlignment: IXLDrawingAlignment
    {
        private readonly IXLDrawingStyle _style;

        public XLDrawingAlignment(IXLDrawingStyle style)
        {
            _style = style;
            Horizontal = XLAlignmentHorizontalValues.Left;
            Vertical = XLAlignmentVerticalValues.Top;
            Direction = XLDrawingTextDirection.Context;
            Orientation = XLDrawingTextOrientation.LeftToRight;
        }
        public XLAlignmentHorizontalValues Horizontal { get; set; }		public IXLDrawingStyle SetHorizontal(XLAlignmentHorizontalValues value) { Horizontal = value; return _style; }
        public XLAlignmentVerticalValues Vertical { get; set; }		public IXLDrawingStyle SetVertical(XLAlignmentVerticalValues value) { Vertical = value; return _style; }
        public Boolean AutomaticSize { get; set; }	public IXLDrawingStyle SetAutomaticSize() { AutomaticSize = true; return _style; }	public IXLDrawingStyle SetAutomaticSize(Boolean value) { AutomaticSize = value; return _style; }
        public XLDrawingTextDirection Direction { get; set; }		public IXLDrawingStyle SetDirection(XLDrawingTextDirection value) { Direction = value; return _style; }
        public XLDrawingTextOrientation Orientation { get; set; }		public IXLDrawingStyle SetOrientation(XLDrawingTextOrientation value) { Orientation = value; return _style; }

    }
}

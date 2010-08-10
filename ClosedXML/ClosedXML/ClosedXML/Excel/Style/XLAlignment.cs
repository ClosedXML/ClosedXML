using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel.Style
{
    public class XLAlignment: IXLAlignment
    {
        public XLAlignmentHorizontalValues Horizontal { get; set; }

        public XLAlignmentVerticalValues Vertical { get; set; }

        public uint Indent { get; set; }

        public bool JustifyLastLine { get; set; }

        public XLAlignmentReadingOrderValues ReadingOrder { get; set; }

        public int RelativeIndent { get; set; }

        public bool ShrinkToFit { get; set; }

        public uint TextRotation { get; set; }

        public bool WrapText { get; set; }

        public bool TopToBottom { get; set; }
    }
}

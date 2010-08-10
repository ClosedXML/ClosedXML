using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
namespace ClosedXML.Excel.Style
{
    public class XLBorder: IXLBorder
    {
        public XLBorderStyleValues LeftBorder { get; set; }

        public Color LeftBorderColor { get; set; }

        public XLBorderStyleValues RightBorder { get; set; }

        public Color RightBorderColor { get; set; }

        public XLBorderStyleValues TopBorder { get; set; }

        public Color TopBorderColor { get; set; }

        public XLBorderStyleValues BottomBorder { get; set; }

        public Color BottomBorderColor { get; set; }

        public bool DiagonalUp { get; set; }

        public bool DiagonalDown { get; set; }

        public XLBorderStyleValues DiagonalBorder { get; set; }

        public Color DiagonalBorderColor { get; set; }
    }
}

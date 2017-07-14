using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLDrawingStyle: IXLDrawingStyle
    {
        public XLDrawingStyle()
        {
            //Font = new XLDrawingFont(this);
            Alignment = new XLDrawingAlignment(this);
            ColorsAndLines = new XLDrawingColorsAndLines(this);
            Size = new XLDrawingSize(this);
            Protection = new XLDrawingProtection(this);
            Properties = new XLDrawingProperties(this);
            Margins = new XLDrawingMargins(this);
            Web = new XLDrawingWeb(this);
        }
        //public IXLDrawingFont Font { get; private set; }
        public IXLDrawingAlignment Alignment { get; private set; }
        public IXLDrawingColorsAndLines ColorsAndLines { get; private set; }
        public IXLDrawingSize Size { get; private set; }
        public IXLDrawingProtection Protection { get; private set; }
        public IXLDrawingProperties Properties { get; private set; }
        public IXLDrawingMargins Margins { get; private set; }
        public IXLDrawingWeb Web { get; private set; }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLDrawingWeb : IXLDrawingWeb
    {
                private readonly IXLDrawingStyle _style;

        public XLDrawingWeb(IXLDrawingStyle style)
        {
            _style = style;
        }
        public String AlternativeText { get; set; }		public IXLDrawingStyle SetAlternativeText(String value) { AlternativeText = value; return _style; }

    }
}

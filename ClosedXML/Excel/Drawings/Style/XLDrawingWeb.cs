using System;

namespace ClosedXML.Excel
{
    internal class XLDrawingWeb : IXLDrawingWeb
    {
        private readonly IXLDrawingStyle _style;

        public XLDrawingWeb(IXLDrawingStyle style)
        {
            _style = style;
        }

        public String? AlternateText { get; set; }

        public IXLDrawingStyle SetAlternateText(String? value) { AlternateText = value; return _style; }
    }
}

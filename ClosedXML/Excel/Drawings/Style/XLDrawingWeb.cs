namespace ClosedXML.Excel
{
    internal class XLDrawingWeb : IXLDrawingWeb
    {
        private readonly IXLDrawingStyle _style;

        public XLDrawingWeb(IXLDrawingStyle style)
        {
            _style = style;
        }
        public string AlternateText { get; set; }		public IXLDrawingStyle SetAlternateText(string value) { AlternateText = value; return _style; }

    }
}

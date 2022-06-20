namespace ClosedXML.Excel
{
    internal class XLDrawingProperties : IXLDrawingProperties
    {
                private readonly IXLDrawingStyle _style;

        public XLDrawingProperties(IXLDrawingStyle style)
        {
            _style = style;
        }
        public XLDrawingAnchor Positioning { get; set; }		public IXLDrawingStyle SetPositioning(XLDrawingAnchor value) { Positioning = value; return _style; }

    }
}

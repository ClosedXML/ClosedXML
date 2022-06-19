namespace ClosedXML.Excel
{
    internal class XLDrawingAlignment: IXLDrawingAlignment
    {
        private readonly IXLDrawingStyle _style;

        public XLDrawingAlignment(IXLDrawingStyle style)
        {
            _style = style;
        }
        public XLDrawingHorizontalAlignment Horizontal { get; set; }		public IXLDrawingStyle SetHorizontal(XLDrawingHorizontalAlignment value) { Horizontal = value; return _style; }
        public XLDrawingVerticalAlignment Vertical { get; set; }		public IXLDrawingStyle SetVertical(XLDrawingVerticalAlignment value) { Vertical = value; return _style; }
        public bool AutomaticSize { get; set; }	public IXLDrawingStyle SetAutomaticSize() { AutomaticSize = true; return _style; }	public IXLDrawingStyle SetAutomaticSize(bool value) { AutomaticSize = value; return _style; }
        public XLDrawingTextDirection Direction { get; set; }		public IXLDrawingStyle SetDirection(XLDrawingTextDirection value) { Direction = value; return _style; }
        public XLDrawingTextOrientation Orientation { get; set; }		public IXLDrawingStyle SetOrientation(XLDrawingTextOrientation value) { Orientation = value; return _style; }

    }
}

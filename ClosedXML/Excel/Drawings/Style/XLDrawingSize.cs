namespace ClosedXML.Excel
{
    internal class XLDrawingSize : IXLDrawingSize
    {
                private readonly IXLDrawingStyle _style;

        public XLDrawingSize(IXLDrawingStyle style)
        {
            _style = style;
        }

        public bool AutomaticSize { get { return _style.Alignment.AutomaticSize; } set { _style.Alignment.AutomaticSize = value; } }	
        public IXLDrawingStyle SetAutomaticSize() { AutomaticSize = true; return _style; }	public IXLDrawingStyle SetAutomaticSize(bool value) { AutomaticSize = value; return _style; }
        public double Height { get; set; }		public IXLDrawingStyle SetHeight(double value) { Height = value; return _style; }
        public double Width { get; set; }		public IXLDrawingStyle SetWidth(double value) { Width = value; return _style; }
    }
}

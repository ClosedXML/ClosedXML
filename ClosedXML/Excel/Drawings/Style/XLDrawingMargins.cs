namespace ClosedXML.Excel
{
    internal class XLDrawingMargins: IXLDrawingMargins
    {
        private readonly IXLDrawingStyle _style;
        public XLDrawingMargins(IXLDrawingStyle style)
        {
            _style = style;
        }
        public bool Automatic { get; set; }	public IXLDrawingStyle SetAutomatic() { Automatic = true; return _style; }	public IXLDrawingStyle SetAutomatic(bool value) { Automatic = value; return _style; }
        double _left;
        public double Left { get { return _left; } set { _left = value; Automatic = false; } }		
        public IXLDrawingStyle SetLeft(double value) { Left = value; return _style; }
        double _right;
        public double Right { get { return _right; } set { _right = value; Automatic = false; } }		public IXLDrawingStyle SetRight(double value) { Right = value; return _style; }
        double _top;
        public double Top { get { return _top; } set { _top = value; Automatic = false; } }		public IXLDrawingStyle SetTop(double value) { Top = value; return _style; }
        double _bottom;
        public double Bottom { get { return _bottom; } set { _bottom = value; Automatic = false; } }		public IXLDrawingStyle SetBottom(double value) { Bottom = value; return _style; }
        public double All
        {
            set
            {
                _left = value;
                _right = value;
                _top = value;
                _bottom = value;
                Automatic = false;
            }
        }	
        public IXLDrawingStyle SetAll(double value) { All = value; return _style; }
    }
}

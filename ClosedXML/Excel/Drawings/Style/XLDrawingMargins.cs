using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLDrawingMargins: IXLDrawingMargins
    {
        private readonly IXLDrawingStyle _style;
        public XLDrawingMargins(IXLDrawingStyle style)
        {
            _style = style;
        }
        public Boolean Automatic { get; set; }	public IXLDrawingStyle SetAutomatic() { Automatic = true; return _style; }	public IXLDrawingStyle SetAutomatic(Boolean value) { Automatic = value; return _style; }
        Double _left;
        public Double Left { get { return _left; } set { _left = value; Automatic = false; } }		
        public IXLDrawingStyle SetLeft(Double value) { Left = value; return _style; }
        Double _right;
        public Double Right { get { return _right; } set { _right = value; Automatic = false; } }		public IXLDrawingStyle SetRight(Double value) { Right = value; return _style; }
        Double _top;
        public Double Top { get { return _top; } set { _top = value; Automatic = false; } }		public IXLDrawingStyle SetTop(Double value) { Top = value; return _style; }
        Double _bottom;
        public Double Bottom { get { return _bottom; } set { _bottom = value; Automatic = false; } }		public IXLDrawingStyle SetBottom(Double value) { Bottom = value; return _style; }
        public Double All
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
        public IXLDrawingStyle SetAll(Double value) { All = value; return _style; }
    }
}

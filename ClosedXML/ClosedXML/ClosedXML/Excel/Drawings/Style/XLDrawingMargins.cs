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
            Automatic = true;
        }
        public Boolean Automatic { get; set; }	public IXLDrawingStyle SetAutomatic() { Automatic = true; return _style; }	public IXLDrawingStyle SetAutomatic(Boolean value) { Automatic = value; return _style; }
        public Double Left { get; set; }		public IXLDrawingStyle SetLeft(Double value) { Left = value; return _style; }
        public Double Right { get; set; }		public IXLDrawingStyle SetRight(Double value) { Right = value; return _style; }
        public Double Top { get; set; }		public IXLDrawingStyle SetTop(Double value) { Top = value; return _style; }
        public Double Bottom { get; set; }		public IXLDrawingStyle SetBottom(Double value) { Bottom = value; return _style; }

    }
}

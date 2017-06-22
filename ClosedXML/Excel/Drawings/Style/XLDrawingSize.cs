using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLDrawingSize : IXLDrawingSize
    {
                private readonly IXLDrawingStyle _style;

        public XLDrawingSize(IXLDrawingStyle style)
        {
            _style = style;
        }

        public Boolean AutomaticSize { get { return _style.Alignment.AutomaticSize; } set { _style.Alignment.AutomaticSize = value; } }	
        public IXLDrawingStyle SetAutomaticSize() { AutomaticSize = true; return _style; }	public IXLDrawingStyle SetAutomaticSize(Boolean value) { AutomaticSize = value; return _style; }
        public Double Height { get; set; }		public IXLDrawingStyle SetHeight(Double value) { Height = value; return _style; }
        public Double Width { get; set; }		public IXLDrawingStyle SetWidth(Double value) { Width = value; return _style; }
    }
}

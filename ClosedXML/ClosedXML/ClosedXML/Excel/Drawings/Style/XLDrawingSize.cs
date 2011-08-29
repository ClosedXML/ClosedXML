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
            Height = 0.82;
            Width = 1.5;
            ScaleHeight = 100;
            ScaleWidth = 100;
        }
        public Double Height { get; set; }		public IXLDrawingStyle SetHeight(Double value) { Height = value; return _style; }
        public Double Width { get; set; }		public IXLDrawingStyle SetWidth(Double value) { Width = value; return _style; }
        public Double ScaleHeight { get; set; }		public IXLDrawingStyle SetScaleHeight(Double value) { ScaleHeight = value; return _style; }
        public Double ScaleWidth { get; set; }		public IXLDrawingStyle SetScaleWidth(Double value) { ScaleWidth = value; return _style; }
        public Boolean LockAspectRatio { get; set; }	public IXLDrawingStyle SetLockAspectRatio() { LockAspectRatio = true; return _style; }	public IXLDrawingStyle SetLockAspectRatio(Boolean value) { LockAspectRatio = value; return _style; }

    }
}

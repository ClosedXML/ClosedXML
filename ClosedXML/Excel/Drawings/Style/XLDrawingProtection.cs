using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    internal class XLDrawingProtection : IXLDrawingProtection
    {
                private readonly IXLDrawingStyle _style;

        public XLDrawingProtection(IXLDrawingStyle style)
        {
            _style = style;
        }
        public Boolean Locked { get; set; }	public IXLDrawingStyle SetLocked() { Locked = true; return _style; }	public IXLDrawingStyle SetLocked(Boolean value) { Locked = value; return _style; }
        public Boolean LockText { get; set; }	public IXLDrawingStyle SetLockText() { LockText = true; return _style; }	public IXLDrawingStyle SetLockText(Boolean value) { LockText = value; return _style; }

    }
}

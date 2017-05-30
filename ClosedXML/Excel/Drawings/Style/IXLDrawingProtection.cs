using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLDrawingProtection
    {
        Boolean Locked { get; set; }
        Boolean LockText { get; set; }

        IXLDrawingStyle SetLocked(); IXLDrawingStyle SetLocked(Boolean value);
        IXLDrawingStyle SetLockText(); IXLDrawingStyle SetLockText(Boolean value);

    }
}

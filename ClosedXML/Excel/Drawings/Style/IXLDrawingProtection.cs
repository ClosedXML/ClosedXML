#nullable disable

using System;

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

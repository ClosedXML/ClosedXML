using System;

namespace ClosedXML.Excel
{
    public interface IXLProtection : IEquatable<IXLProtection>
    {
        Boolean Locked { get; set; }

        Boolean Hidden { get; set; }

        IXLStyle SetLocked(); IXLStyle SetLocked(Boolean value);

        IXLStyle SetHidden(); IXLStyle SetHidden(Boolean value);
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

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

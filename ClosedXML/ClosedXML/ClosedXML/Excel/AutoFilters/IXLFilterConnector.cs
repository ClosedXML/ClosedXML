using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLFilterConnector
    {
        IXLCustomFilteredColumn And { get; }
        IXLCustomFilteredColumn Or { get; }
    }
}

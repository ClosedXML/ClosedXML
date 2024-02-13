using System;
using System.Collections.Generic;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLCFDataBar
    {
        bool Gradient { get; }

        IXLCFDataBar SetGradient(bool value = true);
    }
}

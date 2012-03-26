using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLCFDataBarMax
    {
        void Maximum(XLCFContentType type, String value);
        void HighestValue();
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public enum XLHFMode { OddPagesOnly, OddAndEvenPages, Odd }
    public interface IXLHeaderFooter
    {
        IXLHFItem Left { get; }
        IXLHFItem Center { get; }
        IXLHFItem Right { get; }
        String GetText(XLHFOccurrence occurrence);
    }
}

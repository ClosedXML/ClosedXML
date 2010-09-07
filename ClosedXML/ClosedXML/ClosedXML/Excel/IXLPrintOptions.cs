using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public enum XLPageOrientation { Default, Portrait, Landscape }
    public interface IXLPrintOptions
    {
        IXLRange PrintArea { get; set; }
        XLPageOrientation PageOrientation { get; set; }
        Int32 PagesWide { get; set; }
        Int32 PagesTall { get; set; }
    }
}

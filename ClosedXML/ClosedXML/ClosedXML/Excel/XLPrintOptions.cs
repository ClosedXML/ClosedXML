using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public class XLPrintOptions : IXLPrintOptions
    {
        public XLPrintOptions()
        {
            PageOrientation = XLPageOrientation.Default;
        }
        public IXLRange PrintArea { get; set; }
        public XLPageOrientation PageOrientation { get; set; }
        public Int32 PagesWide { get; set; }
        public Int32 PagesTall { get; set; }
    }
}

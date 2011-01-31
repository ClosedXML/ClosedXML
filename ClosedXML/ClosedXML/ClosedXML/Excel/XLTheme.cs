using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;

namespace ClosedXML.Excel
{
    internal class XLTheme: IXLTheme
    {
        public IXLColor Background1 { get; set; }
        public IXLColor Text1 { get; set; }
        public IXLColor Background2 { get; set; }
        public IXLColor Text2 { get; set; }
        public IXLColor Accent1 { get; set; }
        public IXLColor Accent2 { get; set; }
        public IXLColor Accent3 { get; set; }
        public IXLColor Accent4 { get; set; }
        public IXLColor Accent5 { get; set; }
        public IXLColor Accent6 { get; set; }
        public IXLColor Hyperlink { get; set; }
        public IXLColor FollowedHyperlink { get; set; }
    }
}

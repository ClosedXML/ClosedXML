using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;

namespace ClosedXML.Excel
{
    public interface IXLTheme
    {
        IXLColor Background1 { get; set; }
        IXLColor Text1 { get; set; }
        IXLColor Background2 { get; set; }
        IXLColor Text2 { get; set; }
        IXLColor Accent1 { get; set; }
        IXLColor Accent2 { get; set; }
        IXLColor Accent3 { get; set; }
        IXLColor Accent4 { get; set; }
        IXLColor Accent5 { get; set; }
        IXLColor Accent6 { get; set; }
        IXLColor Hyperlink { get; set; }
        IXLColor FollowedHyperlink { get; set; }
    }
}

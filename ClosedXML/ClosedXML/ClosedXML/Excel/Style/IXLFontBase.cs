using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;

namespace ClosedXML.Excel
{
   

    public interface IXLFontBase
    {
        Boolean Bold { get; set; }
        Boolean Italic { get; set; }
        XLFontUnderlineValues Underline { get; set; }
        Boolean Strikethrough { get; set; }
        XLFontVerticalTextAlignmentValues VerticalAlignment { get; set; }
        Boolean Shadow { get; set; }
        Double FontSize { get; set; }
        IXLColor FontColor { get; set; }
        String FontName { get; set; }
        XLFontFamilyNumberingValues FontFamilyNumbering { get; set; }


    }
}

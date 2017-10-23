using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLDrawingFont : IXLFontBase
    {
        IXLDrawingStyle SetBold(); IXLDrawingStyle SetBold(Boolean value);
        IXLDrawingStyle SetItalic(); IXLDrawingStyle SetItalic(Boolean value);
        IXLDrawingStyle SetUnderline(); IXLDrawingStyle SetUnderline(XLFontUnderlineValues value);
        IXLDrawingStyle SetStrikethrough(); IXLDrawingStyle SetStrikethrough(Boolean value);
        IXLDrawingStyle SetVerticalAlignment(XLFontVerticalTextAlignmentValues value);
        IXLDrawingStyle SetShadow(); IXLDrawingStyle SetShadow(Boolean value);
        IXLDrawingStyle SetFontSize(Double value);
        IXLDrawingStyle SetFontColor(XLColor value);
        IXLDrawingStyle SetFontName(String value);
        IXLDrawingStyle SetFontFamilyNumbering(XLFontFamilyNumberingValues value);
        IXLDrawingStyle SetFontCharSet(XLFontCharSet value);
    }
}

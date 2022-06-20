using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLDrawingFont : IXLFontBase
    {
        IXLDrawingStyle SetBold(); IXLDrawingStyle SetBold(bool value);
        IXLDrawingStyle SetItalic(); IXLDrawingStyle SetItalic(bool value);
        IXLDrawingStyle SetUnderline(); IXLDrawingStyle SetUnderline(XLFontUnderlineValues value);
        IXLDrawingStyle SetStrikethrough(); IXLDrawingStyle SetStrikethrough(bool value);
        IXLDrawingStyle SetVerticalAlignment(XLFontVerticalTextAlignmentValues value);
        IXLDrawingStyle SetShadow(); IXLDrawingStyle SetShadow(bool value);
        IXLDrawingStyle SetFontSize(double value);
        IXLDrawingStyle SetFontColor(XLColor value);
        IXLDrawingStyle SetFontName(string value);
        IXLDrawingStyle SetFontFamilyNumbering(XLFontFamilyNumberingValues value);
        IXLDrawingStyle SetFontCharSet(XLFontCharSet value);
    }
}

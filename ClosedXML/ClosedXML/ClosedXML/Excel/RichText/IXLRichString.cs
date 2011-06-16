using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel.RichText
{
    public interface IXLRichString
    {
        IXLRichText AddText(String text);
        IXLRichString Clear();
        IXLRichText Characters(Int32 index, Int32 length);
    }
}

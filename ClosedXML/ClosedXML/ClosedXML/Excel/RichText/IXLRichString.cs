using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLRichString : IEnumerable<IXLRichText>
    {
        IXLRichText AddText(String text);
        IXLRichString Clear();
        IXLRichText Characters(Int32 index, Int32 length);
        Int32 Count { get; }
    }
}

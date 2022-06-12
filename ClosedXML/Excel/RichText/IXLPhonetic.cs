using System;

namespace ClosedXML.Excel
{
    public interface IXLPhonetic: IEquatable<IXLPhonetic>
    {
        string Text { get; set; }
        int Start { get; set; }
        int End { get; set; }
    }
}

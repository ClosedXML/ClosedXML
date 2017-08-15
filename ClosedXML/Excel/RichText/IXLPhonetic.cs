using System;

namespace ClosedXML.Excel
{
    public interface IXLPhonetic: IEquatable<IXLPhonetic>
    {
        String Text { get; set; }
        Int32 Start { get; set; }
        Int32 End { get; set; }
    }
}

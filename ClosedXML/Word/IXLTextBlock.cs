using System;

namespace ClosedXML.Word
{
    public interface IXLTextBlock : IDisposable
    {
        string text { get; set; }
    }
}

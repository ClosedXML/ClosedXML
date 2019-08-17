using System;

namespace ClosedXML.Word
{
    public interface IXLTextBlock : IDisposable, IXLBlock
    {
        string Text { get; }

        //TODO Implement style class for textblocks
        //IXLTextBlockStyle Style;
    }
}

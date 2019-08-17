using DocumentFormat.OpenXml.Drawing;

namespace ClosedXML.Word
{
    public enum XLBlockTypes {TextBlock};

    public interface IXLBlock
    {
        XLBlockTypes BlockType { get; }

        RunProperties RunProperties { get; set; }
    }
}

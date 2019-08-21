using DocumentFormat.OpenXml.Drawing;

namespace ClosedXML.Word
{
    public enum XLBlockTypes {TextBlock};

    public interface IXLBlock
    {
        XLBlockTypes BlockType { get; }

        /// <summary>
        /// If no ID is supplied when creating a new Block, an ID will be generated. An ID will also be generated when opening an existing document.
        /// </summary>
        int BlockId { get; set; }

        RunProperties RunProperties { get; set; }
    }
}

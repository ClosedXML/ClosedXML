using System;
using DocumentFormat.OpenXml.Drawing;

namespace ClosedXML.Word
{
    public class TextBlock : IXLTextBlock
    {
        public XLBlockTypes BlockType => XLBlockTypes.TextBlock;

        public string BlockName { get; set; }
        public int BlockId { get; set; }

        #region Constructor
        public TextBlock( string text )
        {
            Text = text;

            //TODO Access GenerateBlockIds method from IXLBlocks
            //BlockId = GenerateTextBlockId();
        }

        public TextBlock(int blockId, string text)
        {
            BlockId = blockId;
            Text = text;
        }
        #endregion Constructor

        public string Text { get; set; }

        public RunProperties RunProperties { get; set; }

        public void Dispose( )
        {
            throw new NotImplementedException( );
        }
    }
}

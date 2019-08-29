using System;
using DocumentFormat.OpenXml.Drawing;

namespace ClosedXML.Word
{
    public class TextBlock : IXLTextBlock
    {
        public XLBlockTypes BlockType
        {
            get { return XLBlockTypes.TextBlock; }
        }

        public string BlockId
        {
            get { return _blockID; }
            set { _blockID = value; }
        }

        public string Text
        {
            get { return _text; }
            set { _text = value; }
        }

        public IXLDocument Document { get; set; }

        #region Fields

        //private IXLDocument _document;
        private string _blockID;
        private string _text;

        #endregion Fields

        #region Constructors

        internal TextBlock( IXLDocument document, string blockID, string text )
        {
            _blockID = blockID;
            _text = text;
        }

        public TextBlock( string text )
        {
            _text = text;
            //_blockID = GenerateBlockID();
        }

        public TextBlock( string blockId, string text )
        {
            _blockID = blockId;
            _text = text;
        }

        #endregion Constructors

        public RunProperties RunProperties { get; set; }

        public void Dispose( )
        {
            throw new NotImplementedException( );
        }
    }
}

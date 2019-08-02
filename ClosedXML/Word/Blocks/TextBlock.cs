using DocumentFormat.OpenXml.Drawing;

namespace ClosedXML.Word
{
    public class TextBlock : IXLTextBlock
    {
        public TextBlock( string text )
        {
            Text = text;
        }

        public string Text { get; set; }

        public RunProperties RunProperties { get; set; }

        public void Dispose( )
        {
            throw new System.NotImplementedException( );
        }
    }
}

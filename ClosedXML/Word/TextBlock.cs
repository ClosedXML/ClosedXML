namespace ClosedXML.Word
{
    public class TextBlock : IXLTextBlock
    {
        public TextBlock( string text )
        {
            this.text = text;
        }

        public string text { get; set; }

        public void Dispose( )
        {
            throw new System.NotImplementedException( );
        }
    }
}

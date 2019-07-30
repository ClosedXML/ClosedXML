using System.IO;

using ClosedXML.Word;

namespace ClosedXML_Sandbox
{
    internal class WordTest
    {
        public static void CreateDocument( )
        {
            IXLDocument document = new XLDocument( );
            IXLTextBlock p1 = new TextBlock( "This is a test textblock" );
            document.AddBlock( p1 );
            document.SaveAs( Path.GetTempPath( ) + "test.docx" );
        }
    }
}

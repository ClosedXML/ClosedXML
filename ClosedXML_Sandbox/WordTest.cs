using ClosedXML.Word;
using System.IO;

namespace ClosedXML_Sandbox
{
    internal class WordTest
    {
        public static void CreateDocument( )
        {
            IXLDocument document = new XLDocument( );
            IXLTextBlock p1 = new TextBlock( "This is a test textblock" );
            document.AddTextBlock( p1 );
            XLDocStyle.CreateAndAddCharacterStyle( document.Main.StyleDefinitionsPart, "testId", "test" );
            document.SaveAs( Path.GetTempPath( ) + "test.docx" );
        }
    }
}

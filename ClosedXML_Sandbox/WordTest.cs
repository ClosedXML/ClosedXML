using System.Linq;

using ClosedXML.Word;

using DocumentFormat.OpenXml.Wordprocessing;

using Path = System.IO.Path;

namespace ClosedXML_Sandbox
{
    internal class WordTest
    {
        public static void CreateDocument( )
        {
            //Either open a document or create a new one
            IXLDocument document = new XLDocument( );

            //Add blocks: they are used to construct the document
            IXLTextBlock p1 = new TextBlock( "This is a test textblock" );

            //Add the blocks to the document
            //TODO Create method to add blocks from collection to document at once
            document.AddTextBlock( p1 );

            //TODO Create way to access a text block by id or name
            //IXLTextBlock testTextBlock = document.TextBlock( p1 );

            //Save the document
            document.SaveAs( Path.GetTempPath( ) + "test.docx" );
        }
    }
}

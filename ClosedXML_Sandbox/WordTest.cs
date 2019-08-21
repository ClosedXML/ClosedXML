using System;
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
            IXLTextBlock p2 = new TextBlock( "This is a second textblock" );
            IXLTextBlock p3 = p1;
            //TODO Also changes the text of the first textblock
            p3.Text = "Third textblock test";

            //Add the blocks to the document
            //TODO Create method to add blocks from collection to document at once
            document.Blocks( ).Add( p1 );
            document.Blocks( ).Add( p2 );
            document.Blocks( ).Add( p3 );
            document.Blocks( ).AddBlocksToDocument( );

            Console.WriteLine(p1.BlockId);
            Console.WriteLine(p2.BlockId);
            Console.WriteLine(p3.BlockId);

            //TODO Create way to access a text block by id or name
            IXLTextBlock testTextBlock = document.Block( 0 ) as IXLTextBlock;
            //Do something with the textblock

            //Save the document
            document.SaveAs( Path.GetTempPath( ) + "test.docx" );
        }
    }
}

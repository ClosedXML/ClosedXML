using System;
using ClosedXML.Word;
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
            IXLTextBlock p1 = new TextBlock( "Test paragraph one" );
            IXLTextBlock p2 = new TextBlock( "Test paragraph two" );
            IXLTextBlock p3 = new TextBlock( "A third test paragraph" );
            IXLTextBlock p4 = new TextBlock( "A fourth" );
            IXLTextBlock p5 = new TextBlock( "A fifth" );

            //Add the blocks to the document
            //TODO Create method to add blocks from collection to document at once
            document.AddBlock(p1);
            document.AddBlock(p2);
            document.AddBlock(p3);
            document.AddBlock(p4);
            document.AddBlock(p5);

            Console.WriteLine( p1.BlockId );
            Console.WriteLine( p2.BlockId );
            Console.WriteLine( p3.BlockId );
            Console.WriteLine( p4.BlockId );
            Console.WriteLine( p5.BlockId );

            //TODO Create way to access a text block by id
            //IXLTextBlock testTextBlock = document.Block( "TB0" ) as IXLTextBlock;
            //Do something with the textblock

            //Save the document
            document.SaveAs( Path.GetTempPath( ) + "test.docx" );
        }
    }
}

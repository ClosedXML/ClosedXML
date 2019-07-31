using System.Linq;

using ClosedXML.Word;

using DocumentFormat.OpenXml.Wordprocessing;

using Paragraph = DocumentFormat.OpenXml.Drawing.Paragraph;
using Path = System.IO.Path;
using Run = DocumentFormat.OpenXml.Drawing.Run;
using RunProperties = DocumentFormat.OpenXml.Drawing.RunProperties;

namespace ClosedXML_Sandbox
{
    internal class WordTest
    {
        public static void CreateDocument( )
        {
            IXLDocument document = new XLDocument( );
            IXLTextBlock p1 = new TextBlock( "This is a test textblock" );
            document.AddTextBlock( p1 );
            document.SaveAs( Path.GetTempPath( ) + "test.docx" );
        }
    }
}

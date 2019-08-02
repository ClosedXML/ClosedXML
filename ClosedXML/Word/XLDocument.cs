using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.IO;

namespace ClosedXML.Word
{
    public class XLDocument : IXLDocument
    {
        public string FileName
        {
            get
            {
                return string.Empty;
            }
            set
            {
            }
        }

        public WordprocessingDocument Document { get; set; }
        public MainDocumentPart MainDocumentPart { get; set; }
        public Document DocumentPart { get; set; }
        public Body BodyPart { get; set; }
        
        public XLDocument( string file )
        {
            FileName = file;

            if ( File.Exists( FileName ) )
            {
                Load( FileName );
            }
            else
            {
                CreateNewWordDocument( );
            }
        }

        public XLDocument( )
        {
            CreateNewWordDocument( );
        }

        //Temporary method
        private void CreateNewWordDocument( )
        {
            //string path = Path.Combine( Path.GetTempPath( ), "test.docx" );
            //FileName = path;

            using ( MemoryStream ms = new MemoryStream( ) )
            {
                WordprocessingDocument document = WordprocessingDocument.Create( ms, WordprocessingDocumentType.Document, true );
                MainDocumentPart main = document.AddMainDocumentPart( );
                MainDocumentPart = main;
                main.Document = new Document( );
                Body body = new Body( );
                main.Document.Body = body;

                StyleDefinitionsPart part = document.MainDocumentPart.AddNewPart<StyleDefinitionsPart>( );
                Styles root = new Styles( );
                root.Save( part );

                Document = document;
            }
        }

        public void Dispose( )
        {
            throw new NotImplementedException( );
        }

        public void Save( )
        {
            throw new NotImplementedException( );
        }

        public void SaveAs(
            string file )
        {
            Document.SaveAs(
                file != FileName
                    ? file
                    : FileName );
            //this.document.Close( );
        }

        //Temporary method
        private void AddTextToBody( Body body, string text )
        {
            Paragraph para = body.AppendChild( new Paragraph( ) );
            Run run = para.AppendChild( new Run( ) );
            run.AppendChild( new Text( text ) );
        }

        private void Load( string file )
        {
            using ( WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open( file, true ) )
                LoadWordDocument( wordprocessingDocument );
        }

        private void LoadWordDocument( WordprocessingDocument document )
        {
            Body body = document.MainDocumentPart.Document.Body;
            AddTextToBody( body, "ClosedXML Word Test" );
            document.Close( );
        }

        public void AddTextBlock( IXLTextBlock textBlock )
        {
            //TODO Refactor code
            Body body = Document.MainDocumentPart.Document.Body;
            Paragraph para = body.AppendChild( new Paragraph( ) );
            Run run = para.AppendChild( new Run( ) );
            run.AppendChild( new Text( textBlock.Text ) );

            XLDocumentStyle.CreateAndAddCharacterStyle( MainDocumentPart.StyleDefinitionsPart, "testId", "test" );
            run.PrependChild( new RunProperties( ) );
            RunProperties rPr = run.RunProperties;
            RunStyle rStyle = new RunStyle( );
            rStyle.Val = "Test";
            run.RunProperties.AppendChild( rStyle );
        }

        public void AddTextBlock( string text )
        {
            IXLTextBlock textBlock = new TextBlock( text );
            AddTextBlock( textBlock );
        }

        public void AddBlock( )
        {
            throw new NotImplementedException( );
        }

        public IXLBlocks Blocks( )
        {
            throw new NotImplementedException( );
        }
    }
}

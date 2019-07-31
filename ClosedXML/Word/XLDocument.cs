using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.IO;

namespace ClosedXML.Word
{
    public class XLDocument : IXLDocument
    {
        public WordprocessingDocument document;

        public string FileName
        {
            get
            {
                return String.Empty;
            }
            set
            {
            }
        }

        public MainDocumentPart Main { get; set; }

        //TODO Check whether the specified file exists, if it does open it, else create a new document
        public XLDocument( string file )
        {
            FileName = file;
            Load( FileName );
        }

        public XLDocument( )
        {
            CreateNewWordDocument( );
        }

        //Temporary method
        private void CreateNewWordDocument( )
        {
            string path = Path.Combine( Path.GetTempPath( ), "test.docx" );
            FileName = path;

            using ( MemoryStream ms = new MemoryStream( ) )
            {
                WordprocessingDocument document = WordprocessingDocument.Create( ms, WordprocessingDocumentType.Document, true );
                MainDocumentPart main = document.AddMainDocumentPart( );
                Main = main;
                main.Document = new Document( );
                Body body = new Body( );
                main.Document.Body = body;

                StyleDefinitionsPart part = document.MainDocumentPart.AddNewPart<StyleDefinitionsPart>( );
                Styles root = new Styles( );
                root.Save( part );

                this.document = document;
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
            this.document.SaveAs(
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
            Body body = this.document.MainDocumentPart.Document.Body;
            Paragraph para = body.AppendChild( new Paragraph( ) );
            Run run = para.AppendChild( new Run( ) );
            run.AppendChild( new Text( textBlock.text ) );
        }

        public void AddTextBlock( string text )
        {
            IXLTextBlock textBlock = new TextBlock( text );
            AddTextBlock( textBlock );
        }
    }
}

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
            get { return string.Empty; }
            set { }
        }

        public WordprocessingDocument Document { get; set; }
        public MainDocumentPart MainDocumentPart { get; set; }
        public Document DocumentPart { get; set; }
        public Body BodyPart { get; set; }

        private int _counter = 0;

        #region Constructors

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

        #endregion Constructors

        //Temporary method
        private void CreateNewWordDocument( )
        {
            //string path = Path.Combine( Path.GetTempPath( ), "test.docx" );
            //FileName = path;

            using ( MemoryStream ms = new MemoryStream( ) )
            {
                Document = WordprocessingDocument.Create( ms, WordprocessingDocumentType.Document, true );
                MainDocumentPart = Document.AddMainDocumentPart( );
                DocumentPart = new Document( );
                MainDocumentPart.Document = DocumentPart;
                BodyPart = new Body( );
                DocumentPart.Body = BodyPart;

                //TODO Add styling to document elsewhere
                StyleDefinitionsPart part = MainDocumentPart.AddNewPart<StyleDefinitionsPart>( );
                Styles root = new Styles( );
                root.Save( part );
            }
        }

        public void Dispose( )
        {
            throw new NotImplementedException( );
        }

        public void Save( )
        {
            //GenerateBlockIds( );

            Document.Save( );
        }

        public void SaveAs( string file )
        {
            //GenerateBlockIds( );

            Document.SaveAs(
                file != FileName
                    ? file
                    : FileName );
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
            {
                LoadWordDocument( wordprocessingDocument );
            }
        }

        private void LoadWordDocument( WordprocessingDocument document )
        {
            Body body = document.MainDocumentPart.Document.Body;
            AddTextToBody( body, "ClosedXML Word Test" );
            document.Close( );
        }

        public void AddTextBlock( IXLTextBlock textBlock )
        {
            Paragraph para = BodyPart.AppendChild( new Paragraph( ) );
            Run run = para.AppendChild( new Run( ) );
            run.AppendChild( new Text( textBlock.Text ) );

            //TODO Move styling elsewhere
            XLDocumentStyle.CreateAndAddCharacterStyle( MainDocumentPart.StyleDefinitionsPart, "testId", "test" );
            run.PrependChild( new RunProperties( ) );
            RunStyle rStyle = new RunStyle
            {
                Val = "Test"
            };
            run.RunProperties.AppendChild( rStyle );
        }

        public void AddTextBlock( string text )
        {
            IXLTextBlock textBlock = new TextBlock( text );
            AddTextBlock( textBlock );
        }

        public void AddBlock( IXLBlock block )
        {
            if ( null == block )
            {
                throw new NullReferenceException( "A block cannot be null" );
            }

            _counter++;

            //TODO Implement Add method
            //Blocks( ).Add( block );

            switch ( block.BlockType )
            {
                case XLBlockTypes.TextBlock:
                    IXLTextBlock textBlock = block as IXLTextBlock;
                    AddTextBlock( textBlock );
                    textBlock.BlockId = $"{block.BlockType.ToString( )}{_counter}";
                    break;
                default:
                    throw new IndexOutOfRangeException( $"The block type {block.BlockType} is not a valid block type" );
            }
        }

        public string GenerateBlockIds( )
        {
            throw new NotImplementedException( );
        }

        public IXLBlocks Blocks( )
        {
            XLBlocks retVal = new XLBlocks( this );
            return retVal;
        }

        public IXLBlock Block( string blockId )
        {
            try
            {
                foreach ( IXLBlock block in Blocks( ) )
                {
                    if ( block.BlockId == blockId )
                    {
                        return block;
                    }
                }

                throw new NullReferenceException( );
            }
            catch ( NullReferenceException )
            {
                throw new InvalidOperationException( );
            }
        }
    }
}
